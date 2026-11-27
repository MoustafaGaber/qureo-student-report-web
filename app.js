require('dotenv').config();
const express = require('express');
const fs = require('node:fs');
const path = require('node:path');
const https = require('node:https');
const XLSX = require('xlsx');
const { cert, initializeApp } = require('firebase-admin/app');
const { getAuth } = require('firebase-admin/auth');

const HOST = process.env.HOST || (process.env.NODE_ENV === 'production' ? '0.0.0.0' : '127.0.0.1');
const PORT = Number(process.env.PORT || 3456);
const PORTAL = 'https://me-portal.qureo.education';
const APP = 'https://me-tp.qureo.education';
const CHAPTERS_PATH = '/api/study/students/sections/5/chapters';
const TOTAL_CHAPTERS = 42;
const CONCURRENCY = 6;
const MAX_FILE_BYTES = 8 * 1024 * 1024;
const MAX_STUDENTS = 500;
const CLIENT_DIST = path.join(__dirname, 'client', 'dist');

const firebaseWebConfig = {
  apiKey: process.env.VITE_FIREBASE_API_KEY || '',
  authDomain: process.env.VITE_FIREBASE_AUTH_DOMAIN || '',
  projectId: process.env.VITE_FIREBASE_PROJECT_ID || '',
  appId: process.env.VITE_FIREBASE_APP_ID || '',
  messagingSenderId: process.env.VITE_FIREBASE_MESSAGING_SENDER_ID || '',
};
const firebaseWebReady = Object.values(firebaseWebConfig).every(Boolean) && firebaseWebConfig.projectId === process.env.FIREBASE_PROJECT_ID;
let firebaseAuth = null;
if (firebaseWebReady) {
  try {
    let serviceAccount = null;
    if (process.env.FIREBASE_SERVICE_ACCOUNT_JSON) {
      serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT_JSON);
    } else if (process.env.FIREBASE_SERVICE_ACCOUNT_PATH) {
      const serviceAccountPath = path.resolve(__dirname, process.env.FIREBASE_SERVICE_ACCOUNT_PATH);
      serviceAccount = JSON.parse(fs.readFileSync(serviceAccountPath, 'utf8'));
    }
    if (serviceAccount?.project_id === process.env.FIREBASE_PROJECT_ID) {
      initializeApp({ credential: cert(serviceAccount), projectId: process.env.FIREBASE_PROJECT_ID });
      firebaseAuth = getAuth();
    }
  } catch {
    // Fail closed. Never log service-account contents or details from the key file.
    firebaseAuth = null;
  }
}

const allowedEmails = new Set((process.env.ALLOWED_EMAILS || '').split(',').map((value) => value.trim().toLowerCase()).filter(Boolean));
const allowedDomains = new Set((process.env.ALLOWED_EMAIL_DOMAINS || '').split(',').map((value) => value.trim().replace(/^@/, '').toLowerCase()).filter(Boolean));
const hasAllowlist = allowedEmails.size > 0 || allowedDomains.size > 0;

const server = express();
server.disable('x-powered-by');
server.use((req, res, next) => {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('Referrer-Policy', 'strict-origin-when-cross-origin');
  res.setHeader('Cache-Control', 'no-store');
  next();
});

function normalizeStudentRow(rawObj) {
  let name = '', className = '', studentId = '', password = '', inferredClass = '';
  for (const [rawKey, rawVal] of Object.entries(rawObj)) {
    const key = String(rawKey).replace(/[\r\n]+/g, ' ').trim();
    const compactKey = key.replace(/["\s;،:]/g, '').toLowerCase();
    const val = String(rawVal ?? '').trim();
    if (!name && (compactKey === 'name' || compactKey === 'الاسم' || key.includes('اسم الطالب'))) {
      name = val;
      const match = key.match(/(?:المدرسة\s*:\s*)?(.+?)\s*اسم\s*الطالب/);
      if (match && match[1]) inferredClass = match[1].trim();
    } else if (!className && (compactKey === 'class' || compactKey === 'الفصل' || compactKey === 'فصل')) className = val;
    else if (!studentId && (compactKey === 'student_id' || compactKey === 'studentid' || key.includes('المستخدم') || key.includes('كود الطالب'))) studentId = val;
    else if (!password && (compactKey === 'password' || key.includes('كلمة المرور') || key.includes('كلمة السر'))) password = val;
  }
  return { name, class: className || inferredClass, student_id: studentId, password };
}

function parseCsv(text) {
  text = text.replace(/^\uFEFF/, '');
  const firstLine = text.split(/\r?\n/, 1)[0] || '';
  let delimiter = ',', best = -1;
  for (const candidate of [',', ';', '\t']) {
    let count = 0, quoted = false;
    for (let i = 0; i < firstLine.length; i++) {
      if (firstLine[i] === '"' && firstLine[i + 1] === '"' && quoted) i++;
      else if (firstLine[i] === '"') quoted = !quoted;
      else if (!quoted && firstLine[i] === candidate) count++;
    }
    if (count > best) { best = count; delimiter = candidate; }
  }
  const records = []; let row = [], cell = '', quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i], next = text[i + 1];
    if (ch === '"' && quoted && next === '"') { cell += '"'; i++; }
    else if (ch === '"') quoted = !quoted;
    else if (!quoted && ch === delimiter) { row.push(cell); cell = ''; }
    else if (!quoted && (ch === '\n' || ch === '\r')) {
      if (ch === '\r' && next === '\n') i++;
      row.push(cell); records.push(row); row = []; cell = '';
    } else cell += ch;
  }
  if (cell.length || row.length) { row.push(cell); records.push(row); }
  if (quoted) throw new Error('يوجد حقل نصي بعلامة اقتباس غير مكتملة في ملف CSV.');
  const headers = (records.shift() || []).map((h) => h.trim());
  return records.filter((cells) => cells.some((v) => String(v).trim() !== '')).map((cells) => {
    const obj = {}; headers.forEach((h, i) => { obj[h] = cells[i] ?? ''; }); return normalizeStudentRow(obj);
  });
}

function parseStudents(buffer, fileName) {
  const ext = path.extname(fileName).toLowerCase();
  if (!['.xlsx', '.xls', '.csv'].includes(ext)) throw new Error('صيغة الملف غير مدعومة. اختر ملف XLSX أو XLS أو CSV.');
  let rows;
  if (ext === '.csv') rows = parseCsv(buffer.toString('utf8'));
  else {
    let wb;
    try { wb = XLSX.read(buffer, { type: 'buffer' }); } catch { throw new Error('تعذر قراءة ملف Excel. تحقق من سلامة الملف.'); }
    if (!wb.SheetNames.length) throw new Error('ملف Excel فارغ ولا يحتوي على ورقة بيانات.');
    rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' }).map(normalizeStudentRow)
      .filter((r) => r.name || r.class || r.student_id || r.password);
  }
  if (!rows.length) throw new Error('الملف فارغ أو لا يحتوي على صفوف طلاب.');
  if (rows.length > MAX_STUDENTS) throw new Error(`عدد الطلاب يتجاوز الحد المسموح (${MAX_STUDENTS}).`);
  const missing = rows.findIndex((r) => !r.student_id || !r.password);
  if (missing >= 0) throw new Error(`بيانات الطالب في الصف ${missing + 2} ناقصة: يلزم رقم الطالب وكلمة المرور.`);
  return rows;
}

function formatEgyptTime(value) {
  if (!value) return '';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  return new Intl.DateTimeFormat('sv-SE', { timeZone: 'Africa/Cairo', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false }).format(date);
}

class CookieJar {
  constructor() { this.cookies = new Map(); }
  save(res) {
    for (const raw of res.headers['set-cookie'] || []) {
      const pair = raw.split(';')[0], eq = pair.indexOf('=');
      if (eq > 0) this.cookies.set(pair.slice(0, eq).trim(), pair.slice(eq + 1).trim());
    }
  }
  header() { return [...this.cookies].map(([k, v]) => `${k}=${v}`).join('; '); }
}

function httpRequest(url, opts = {}) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    if (u.protocol !== 'https:' || ![new URL(PORTAL).hostname, new URL(APP).hostname].includes(u.hostname)) return reject(new Error('blocked destination'));
    const req = https.request({ hostname: u.hostname, port: 443, path: u.pathname + u.search, method: opts.method || 'GET', headers: opts.headers || {}, timeout: 20000 }, (res) => {
      const chunks = []; res.on('data', (c) => chunks.push(c));
      res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, text: Buffer.concat(chunks).toString('utf8') }));
    });
    req.on('timeout', () => req.destroy(new Error('timeout'))); req.on('error', reject);
    if (opts.body) req.write(opts.body); req.end();
  });
}

async function fetchWithCookies(jar, url, opts = {}, redirects = 0) {
  if (redirects > 8) throw new Error('redirect limit');
  const headers = { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130 Safari/537.36', ...(opts.headers || {}) };
  if (jar.header()) headers.Cookie = jar.header();
  const res = await httpRequest(url, { ...opts, headers }); jar.save(res);
  if ([301, 302, 303, 307, 308].includes(res.status)) {
    const location = res.headers.location; if (!location) return res;
    const nextUrl = new URL(location, url);
    if (![new URL(PORTAL).hostname, new URL(APP).hostname].includes(nextUrl.hostname)) throw new Error('blocked redirect');
    const method = res.status === 303 || ([301, 302].includes(res.status) && (opts.method || 'GET') === 'POST') ? 'GET' : opts.method;
    return fetchWithCookies(jar, nextUrl.toString(), { ...opts, method, body: method === 'GET' ? undefined : opts.body }, redirects + 1);
  }
  return res;
}

async function checkStudent(student) {
  const jar = new CookieJar();
  try {
    const login = await fetchWithCookies(jar, `${PORTAL}/oauth2/password-login`, { method: 'POST', headers: { 'Content-Type': 'application/json', Accept: 'application/json, text/plain, */*', Referer: `${PORTAL}/login?referer=${encodeURIComponent(PORTAL + '/')}`, Origin: PORTAL }, body: JSON.stringify({ student_id: student.student_id.trim(), password: student.password.trim() }) });
    if (login.status < 200 || login.status >= 300) return { status: `تعذر تسجيل الدخول (${login.status})`, chaptersCount: '', diamonds: '', lastActivity: '' };
    await fetchWithCookies(jar, `${PORTAL}/`).catch(() => {}); await fetchWithCookies(jar, `${APP}/`).catch(() => {});
    const res = await fetchWithCookies(jar, `${APP}${CHAPTERS_PATH}`, { headers: { Accept: 'application/json, text/plain, */*', Referer: `${APP}/` } });
    if (res.status < 200 || res.status >= 300) return { status: `تعذر جلب بيانات الفصول (${res.status})`, chaptersCount: '', diamonds: '', lastActivity: '' };
    let chapters; try { chapters = JSON.parse(res.text); } catch { return { status: 'استجابة غير صالحة', chaptersCount: '', diamonds: '', lastActivity: '' }; }
    if (!Array.isArray(chapters)) return { status: 'استجابة غير متوقعة', chaptersCount: '', diamonds: '', lastActivity: '' };
    const dates = chapters.flatMap((c) => [c.started_at, c.completed_at]).filter(Boolean).sort();
    return { status: 'ok', chaptersCount: chapters.length, diamonds: chapters.reduce((sum, c) => sum + (Number(c.diamond) || 0), 0), lastActivity: formatEgyptTime(dates.at(-1)) };
  } catch { return { status: 'تعذر الاتصال بخدمة QUREO', chaptersCount: '', diamonds: '', lastActivity: '' }; }
}

async function runJob(students, onProgress, shouldStop) {
  const started = Date.now();
  let next = 0, done = 0, okCount = 0, failCount = 0;
  const rows = new Array(students.length);
  async function worker() {
    while (next < students.length && !shouldStop()) {
      const i = next++, s = students[i], result = await checkStudent(s);
      rows[i] = { name: s.name, className: s.class, studentId: s.student_id, status: result.status, chaptersCount: result.chaptersCount, totalChapters: TOTAL_CHAPTERS, pct: result.chaptersCount === '' ? '' : Math.round(result.chaptersCount / TOTAL_CHAPTERS * 100), diamonds: result.diamonds, lastActivity: result.lastActivity };
      done++; if (result.status === 'ok') okCount++; else failCount++;
      onProgress({ row: rows[i], total: students.length, done, okCount, failCount, elapsedSec: ((Date.now() - started) / 1000).toFixed(1) });
    }
  }
  await Promise.all(Array.from({ length: Math.min(CONCURRENCY, students.length) }, worker));
  return { rows: rows.filter(Boolean), total: students.length, done, okCount, failCount, stopped: shouldStop(), elapsedSec: ((Date.now() - started) / 1000).toFixed(1) };
}

function isEmailAllowed(email) {
  const normalized = String(email || '').trim().toLowerCase();
  const domain = normalized.split('@')[1] || '';
  return hasAllowlist && (allowedEmails.has(normalized) || allowedDomains.has(domain));
}

async function requireTeacher(req, res, next) {
  if (!firebaseAuth || !hasAllowlist) return res.status(503).json({ error: 'تسجيل الدخول غير مهيأ بعد. يلزم إعداد Firebase واعتماد الخادم وقائمة حسابات المعلمين.' });
  const authorization = req.get('authorization') || '';
  const match = authorization.match(/^Bearer\s+(.+)$/i);
  if (!match) return res.status(401).json({ error: 'سجّل الدخول بحساب Google للمتابعة.' });
  try {
    const decoded = await firebaseAuth.verifyIdToken(match[1]);
    if (!decoded.email_verified || !isEmailAllowed(decoded.email)) return res.status(403).json({ error: 'هذا الحساب غير مصرح له باستخدام التطبيق.' });
    req.teacher = { uid: decoded.uid, email: decoded.email };
    next();
  } catch {
    return res.status(401).json({ error: 'انتهت جلسة الدخول أو تعذر التحقق منها. سجّل الدخول مجددًا.' });
  }
}

server.get('/api/health', (req, res) => res.json({ ok: true }));
server.get('/api/config', (req, res) => res.json({
  firebase: firebaseWebConfig,
  authReady: Boolean(firebaseAuth) && hasAllowlist,
  drivePickerApiKey: process.env.VITE_GOOGLE_PICKER_API_KEY || '',
  drivePickerReady: Boolean(process.env.VITE_GOOGLE_PICKER_API_KEY && firebaseWebConfig.messagingSenderId),
  maxFileBytes: process.env.VERCEL ? 4 * 1024 * 1024 : MAX_FILE_BYTES,
}));
server.use('/api', express.json({ limit: '16kb' }));
server.get('/api/me', requireTeacher, (req, res) => res.json({ email: req.teacher.email }));
server.post('/api/check', requireTeacher, express.raw({ type: 'application/octet-stream', limit: MAX_FILE_BYTES }), async (req, res) => {
  const filename = path.basename(String(req.query.name || ''));
  let students;
  try {
    const buffer = req.body;
    if (!Buffer.isBuffer(buffer) || !buffer.length) return res.status(400).json({ error: 'الملف فارغ.' });
    const maxFileBytes = process.env.VERCEL ? 4 * 1024 * 1024 : MAX_FILE_BYTES;
    if (buffer.length > maxFileBytes) return res.status(413).json({ error: process.env.VERCEL ? 'على نسخة Vercel المجانية، الحد العملي للملف 4 ميجابايت بسبب حد طلبات المنصة.' : 'حجم الملف يتجاوز الحد الأقصى (8 ميجابايت).' });
    students = parseStudents(buffer, filename);
  } catch (error) {
    return res.status(400).json({ error: error.message || 'تعذر التحقق من الملف.' });
  }
  res.status(200);
  res.setHeader('Content-Type', 'application/x-ndjson; charset=utf-8');
  res.setHeader('Cache-Control', 'no-cache, no-transform');
  res.setHeader('X-Accel-Buffering', 'no');
  res.flushHeaders?.();
  let disconnected = false;
  res.on('close', () => { if (!res.writableEnded) disconnected = true; });
  const send = (event) => { if (!disconnected && !res.writableEnded) res.write(`${JSON.stringify(event)}\n`); };
  send({ type: 'start', total: students.length });
  try {
    const result = await runJob(students, (progress) => send({ type: 'progress', ...progress }), () => disconnected);
    send({ type: 'complete', ...result });
    if (!disconnected && !res.writableEnded) res.end();
  } catch {
    send({ type: 'error', error: 'حدث خطأ أثناء الفحص. النتائج المكتملة ظاهرة ويمكن تنزيلها.' });
    if (!disconnected && !res.writableEnded) res.end();
  }
});

if (process.env.VERCEL) {
  // Vercel serves files from /public through its CDN; Express handles API requests only.
} else if (require('node:fs').existsSync(CLIENT_DIST)) {
  server.use(express.static(CLIENT_DIST, { index: false, maxAge: process.env.NODE_ENV === 'production' ? '1h' : 0 }));
  server.get('*', (req, res, next) => req.path.startsWith('/api/') ? next() : res.sendFile(path.join(CLIENT_DIST, 'index.html')));
} else {
  server.get('/', (req, res) => res.status(503).type('text/plain').send('لم يتم بناء واجهة React بعد. نفّذ npm run build ثم npm start.'));
}

server.use((error, req, res, next) => {
  if (res.headersSent) return next(error);
  if (error.type === 'entity.too.large') return res.status(413).json({ error: process.env.VERCEL ? 'حجم الملف يتجاوز الحد المتاح لنسخة Vercel المجانية (4 ميجابايت).' : 'حجم الملف يتجاوز الحد الأقصى (8 ميجابايت).' });
  return res.status(400).json({ error: 'تعذر معالجة الطلب. تحقق من الملف وحاول مرة أخرى.' });
});

if (require.main === module && !process.env.VERCEL) server.listen(PORT, HOST, () => console.log(`QUREO Student Report API listening on ${HOST}:${PORT}`));
module.exports = server;
