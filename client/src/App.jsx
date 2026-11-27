import React, { useEffect, useMemo, useRef, useState } from 'react';
import { getApp, getApps, initializeApp } from 'firebase/app';
import { getAuth, onAuthStateChanged, signInWithPopup, reauthenticateWithPopup, signOut, GoogleAuthProvider } from 'firebase/auth';

const DRIVE_FILE_SCOPE = 'https://www.googleapis.com/auth/drive.file';
let pickerScriptPromise;

function loadGooglePicker() {
  if (window.google?.picker) return Promise.resolve();
  if (!pickerScriptPromise) {
    pickerScriptPromise = new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = 'https://apis.google.com/js/api.js';
      script.async = true;
      script.onload = () => window.gapi.load('picker', { callback: resolve, onerror: reject });
      script.onerror = reject;
      document.head.appendChild(script);
    });
  }
  return pickerScriptPromise;
}

async function request(path, user, options = {}) {
  const token = user ? await user.getIdToken() : '';
  const response = await fetch(path, {
    ...options,
    headers: { ...(options.body instanceof File ? { 'Content-Type': 'application/octet-stream' } : {}), ...(token ? { Authorization: `Bearer ${token}` } : {}), ...(options.headers || {}) },
    cache: 'no-store',
  });
  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    throw new Error(body.error || 'تعذر تنفيذ الطلب.');
  }
  return response;
}

function getSignInErrorMessage(error) {
  const code = String(error?.code || '').replace(/^auth\//, '');
  const known = {
    'unauthorized-domain': 'النطاق الحالي غير مسموح في Firebase Authentication. أضف 127.0.0.1 إلى النطاقات المصرح بها.',
    'operation-not-allowed': 'تسجيل Google غير مفعّل في Firebase Authentication.',
    'invalid-api-key': 'مفتاح Firebase غير صحيح. راجع إعدادات تطبيق الويب.',
    'network-request-failed': 'تعذر الاتصال بخدمة Google. تحقق من الإنترنت ثم أعد المحاولة.',
    'popup-blocked': 'المتصفح منع نافذة الدخول. أعد المحاولة أو اسمح بالنوافذ المنبثقة.',
    'popup-closed-by-user': 'أُغلقت نافذة Google قبل إكمال الدخول. أعد المحاولة واترك النافذة مفتوحة حتى تنتهي.',
  };
  const message = known[code] || 'تعذر إكمال تسجيل الدخول عبر Google.';
  return code ? `${message} (auth/${code})` : message;
}

function App() {
  const [config, setConfig] = useState(null);
  const [user, setUser] = useState(null);
  const [authorizedEmail, setAuthorizedEmail] = useState('');
  const [authError, setAuthError] = useState('');
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState({ running: false, rows: [], total: 0, done: 0, okCount: 0, failCount: 0, elapsedSec: 0, hasReport: false });
  const [message, setMessage] = useState('');
  const [busy, setBusy] = useState(false);
  const [drivePickerBusy, setDrivePickerBusy] = useState(false);
  const fileRef = useRef(null);
  const checkAbortRef = useRef(null);
  const loggedQureoErrorsRef = useRef(new Set());

  useEffect(() => {
    fetch('/api/config', { cache: 'no-store' }).then((res) => res.json()).then(setConfig).catch(() => setAuthError('تعذر الاتصال بالخادم.'));
  }, []);

  const auth = useMemo(() => {
    if (!config?.authReady) return null;
    const firebaseApp = getApps().length ? getApp() : initializeApp(config.firebase);
    return getAuth(firebaseApp);
  }, [config]);

  useEffect(() => {
    if (!auth) return undefined;
    return onAuthStateChanged(auth, (nextUser) => {
      setUser(nextUser);
      setAuthorizedEmail('');
      if (nextUser) setAuthError('');
    });
  }, [auth]);

  useEffect(() => {
    if (!user) return undefined;
    let active = true;
    request('/api/me', user).then(async (res) => {
      const data = await res.json(); if (active) setAuthorizedEmail(data.email);
    }).catch((error) => { if (active) setAuthError(error.message); });
    return () => { active = false; };
  }, [user]);

  async function login() {
    setAuthError('');
    try { await signInWithPopup(auth, new GoogleAuthProvider()); }
    catch (error) { setAuthError(getSignInErrorMessage(error)); }
  }

  async function logout() {
    try { await signOut(auth); setAuthorizedEmail(''); setStatus({ running: false, rows: [], total: 0, done: 0, okCount: 0, failCount: 0, elapsedSec: 0, hasReport: false }); }
    catch { setAuthError('تعذر تسجيل الخروج.'); }
  }

  async function startCheck() {
    if (!file) { setMessage('اختر ملف الطلاب أولًا.'); return; }
    const maxFileBytes = config.maxFileBytes || 8 * 1024 * 1024;
    if (file.size > maxFileBytes) { setMessage(`حجم الملف يتجاوز الحد الأقصى (${Math.round(maxFileBytes / 1024 / 1024)} ميجابايت) لهذه النسخة.`); return; }
    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) { setMessage('صيغة الملف غير مدعومة. اختر XLSX أو XLS أو CSV.'); return; }
    setBusy(true); setMessage('');
    setStatus({ running: true, rows: [], total: 0, done: 0, okCount: 0, failCount: 0, elapsedSec: 0, hasReport: false });
    const controller = new AbortController();
    checkAbortRef.current = controller;
    loggedQureoErrorsRef.current.clear();
    try {
      const token = await user.getIdToken();
      const response = await fetch(`/api/check?name=${encodeURIComponent(file.name)}`, { method: 'POST', body: file, signal: controller.signal, headers: { 'Content-Type': 'application/octet-stream', Authorization: `Bearer ${token}` }, cache: 'no-store' });
      if (!response.ok) { const body = await response.json().catch(() => ({})); throw new Error(body.error || 'تعذر بدء الفحص.'); }
      if (!response.body) throw new Error('المتصفح لا يدعم استقبال تقدم الفحص.');
      const reader = response.body.getReader(), decoder = new TextDecoder();
      let pending = '';
      while (true) {
        const { value, done } = await reader.read();
        pending += decoder.decode(value || new Uint8Array(), { stream: !done });
        const lines = pending.split('\n'); pending = lines.pop() || '';
        for (const line of lines) {
          if (!line.trim()) continue;
          const event = JSON.parse(line);
          if (event.type === 'start') setStatus((prev) => ({ ...prev, total: event.total }));
          if (event.type === 'progress') {
            if (event.row?.diagnostic) {
              const key = `${event.row.status}\n${event.row.diagnostic}`;
              if (!loggedQureoErrorsRef.current.has(key)) {
                loggedQureoErrorsRef.current.add(key);
                console.error('[QUREO] تعذر جلب بيانات الفصول', { status: event.row.status, message: event.row.diagnostic });
              }
            }
            setStatus((prev) => ({ ...prev, running: true, total: event.total, done: event.done, okCount: event.okCount, failCount: event.failCount, elapsedSec: event.elapsedSec, rows: [...prev.rows, event.row] }));
          }
          if (event.type === 'complete') setStatus((prev) => ({ ...prev, running: false, stopRequested: event.stopped, total: event.total, done: event.done, okCount: event.okCount, failCount: event.failCount, elapsedSec: event.elapsedSec, hasReport: prev.rows.length > 0 }));
          if (event.type === 'error') throw new Error(event.error);
        }
        if (done) break;
      }
      setStatus((prev) => ({ ...prev, running: false, hasReport: prev.rows.length > 0 }));
      setMessage('اكتمل الفحص.');
    } catch (error) {
      if (error.name === 'AbortError') setMessage('تم إيقاف الفحص؛ ستبقى النتائج المكتملة ويمكن تنزيلها.');
      else setMessage(error.message || 'انقطع الاتصال أثناء الفحص؛ احتُفظ بالنتائج التي وصلت.');
      setStatus((prev) => ({ ...prev, running: false, stopRequested: true, hasReport: prev.rows.length > 0 }));
    } finally { checkAbortRef.current = null; setBusy(false); }
  }

  function stopCheck() {
    checkAbortRef.current?.abort();
  }

  async function downloadReport() {
    try {
      const XLSX = await import('xlsx');
      const validRows = status.rows.filter(Boolean), wb = XLSX.utils.book_new();
      wb.Workbook = { Views: [{ RTL: true }] };
      const data = [['الاسم', 'الفصل', 'رقم الطالب', 'الحالة', 'الفصول', 'من 42', 'النسبة %', 'الألماس', 'آخر نشاط']];
      const buckets = [0, 0, 0, 0, 0]; let failed = 0;
      for (const row of validRows) {
        data.push([row.name, row.className, row.studentId, row.status, row.chaptersCount === '' ? '' : Number(row.chaptersCount), row.totalChapters, row.pct === '' ? '' : Number(row.pct), row.diamonds === '' ? '' : Number(row.diamonds), row.lastActivity]);
        if (row.status !== 'ok' || row.chaptersCount === '') failed++;
        else { const n = Number(row.chaptersCount); buckets[n === 0 ? 0 : n <= 10 ? 1 : n <= 17 ? 2 : n <= 41 ? 3 : 4]++; }
      }
      const ws = XLSX.utils.aoa_to_sheet(data); ws['!views'] = [{ rightToLeft: true }]; ws['!cols'] = [32, 24, 15, 24, 10, 10, 11, 10, 18].map((wch) => ({ wch })); XLSX.utils.book_append_sheet(wb, ws, 'بيانات الطلاب');
      const total = buckets.reduce((a, b) => a + b, 0), pct = (n) => total ? Math.round(n / total * 1000) / 10 : 0;
      const summary = [['الفئة / مستوى الإنجاز', 'عدد الطلاب', 'النسبة من المفحوصين %'], ['لم ينجزوا أي فصل (0 فصل)', buckets[0], pct(buckets[0])], ['أنجزوا من 1 إلى 10 فصول', buckets[1], pct(buckets[1])], ['أنجزوا من 11 إلى 17 فصلًا', buckets[2], pct(buckets[2])], ['أنجزوا من 18 إلى 41 فصلًا', buckets[3], pct(buckets[3])], ['أتموا المنهج كاملًا (42 فصلًا)', buckets[4], pct(buckets[4])], ['إجمالي الطلاب المفحوصين بنجاح', total, total ? 100 : 0]];
      if (failed) summary.push(['تعذر فحصهم', failed, ''], ['الإجمالي الكلي بالملف', total + failed, '']);
      const ws2 = XLSX.utils.aoa_to_sheet(summary); ws2['!views'] = [{ rightToLeft: true }]; ws2['!cols'] = [36, 16, 24].map((wch) => ({ wch })); XLSX.utils.book_append_sheet(wb, ws2, 'ملخص الإحصائيات');
      const blob = new Blob([XLSX.write(wb, { type: 'array', bookType: 'xlsx' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob), link = document.createElement('a');
      const stamp = new Intl.DateTimeFormat('sv-SE', { timeZone: 'Africa/Cairo', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' }).format(new Date()).replace(' ', '_').replaceAll(':', '-');
      const filename = `qureo-report-${stamp}.xlsx`;
      link.href = url; link.download = filename; link.click(); URL.revokeObjectURL(url);
    } catch (error) { setMessage(error.message); }
  }

  async function pickFromDrive() {
    if (!config.drivePickerReady) {
      setMessage('اختيار Drive يحتاج تفعيل Google Picker API وإضافة مفتاحه المقيّد في إعدادات التطبيق.');
      return;
    }
    setDrivePickerBusy(true); setMessage('جارٍ فتح Google Drive…');
    try {
      const provider = new GoogleAuthProvider();
      provider.addScope(DRIVE_FILE_SCOPE);
      const authResult = await reauthenticateWithPopup(user, provider);
      const accessToken = GoogleAuthProvider.credentialFromResult(authResult)?.accessToken;
      if (!accessToken) throw new Error('لم يمنح Google إذن اختيار ملف Drive.');
      await loadGooglePicker();
      const selectedFile = await new Promise((resolve, reject) => {
        const view = new window.google.picker.DocsView(window.google.picker.ViewId.DOCS)
          .setMimeTypes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv,application/vnd.google-apps.spreadsheet');
        const picker = new window.google.picker.PickerBuilder()
          .addView(view)
          .setOAuthToken(accessToken)
          .setDeveloperKey(config.drivePickerApiKey)
          .setAppId(config.firebase.messagingSenderId)
          .setCallback(async (data) => {
            if (data.action === window.google.picker.Action.CANCEL) return resolve(null);
            if (data.action !== window.google.picker.Action.PICKED) return;
            try {
              const doc = data.docs?.[0];
              if (!doc) return resolve(null);
              const isGoogleSheet = doc.mimeType === 'application/vnd.google-apps.spreadsheet';
              const endpoint = isGoogleSheet
                ? `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(doc.id)}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
                : `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(doc.id)}?alt=media`;
              const response = await fetch(endpoint, { headers: { Authorization: `Bearer ${accessToken}` } });
              if (!response.ok) throw new Error('تعذر تنزيل الملف المختار من Drive. تحقق من إذن الوصول للملف.');
              const filename = isGoogleSheet ? `${doc.name}.xlsx` : doc.name;
              const blob = await response.blob();
              resolve(new File([blob], filename, { type: blob.type || doc.mimeType || 'application/octet-stream' }));
            } catch (error) { reject(error); }
          })
          .build();
        picker.setVisible(true);
      });
      if (selectedFile) { setFile(selectedFile); setMessage(`تم اختيار الملف: ${selectedFile.name}`); }
      else setMessage('لم يتم اختيار ملف من Drive.');
    } catch (error) { setMessage(error.message || 'تعذر فتح Google Drive.'); }
    finally { setDrivePickerBusy(false); }
  }

  if (!config) return <main className="shell"><div className="panel">جارٍ تحميل إعدادات التطبيق…</div></main>;

  return <main className="shell">
    <header className="hero"><div className="brand">Q</div><div><p className="eyebrow">لوحة متابعة تعليمية</p><h1>تقرير تقدم طلاب QUREO</h1><p className="subtitle">فحص الحسابات وعرض تقدم كل طالب في المنهج.</p></div></header>
    <aside className="notice"><strong>تنبيه الخصوصية</strong><p>يرسل هذا التطبيق بيانات الدخول إلى QUREO للتحقق من الحسابات. استخدمه عند وجود تصريح، ووفق سياسة المدرسة. لا يعرض التطبيق كلمات المرور ولا يحفظها في ملفات.</p></aside>

    {!config.authReady ? <section className="panel setup"><h2>إعداد تسجيل Google مطلوب</h2><p>يلزم إعداد Firebase، وملف اعتماد Firebase Admin الخاص بالخادم، وقائمة حسابات المعلمين المصرح بها. لا يقبل الخادم عمليات الفحص قبل اكتمال الإعداد.</p><p className="muted">بعد تسجيل الدخول، يجب أن يكون البريد ضمن قائمة السماح التي يضبطها مسؤول التطبيق.</p></section> : !user ? <section className="panel login"><h2>مرحبًا بك</h2><p>سجّل الدخول بحساب Google المعتمد للمدرسة للمتابعة.</p><button className="button google" onClick={login}>المتابعة باستخدام Google</button>{authError && <p className="error-text">{authError}</p>}</section> : !authorizedEmail ? <section className="panel login"><h2>جارٍ التحقق من صلاحية الحساب</h2>{authError && <><p className="error-text">{authError}</p><button className="button secondary" onClick={logout}>تسجيل الخروج</button></>}</section> : <>
      <section className="panel upload"><div className="section-title"><div><h2>ملف الطلاب</h2><p>XLSX أو XLS أو CSV · حتى {Math.round((config.maxFileBytes || 8 * 1024 * 1024) / 1024 / 1024)} ميجابايت · حتى 500 طالب</p></div><span className="chip">42 فصلًا · 6 عمليات بالتوازي</span></div>
        <div className="controls"><button className="file-button" type="button" onClick={() => fileRef.current?.click()}><span className="upload-icon">↑</span><span><b>{file?.name || 'اختر ملفًا من جهازك'}</b><small>{file ? `${(file.size / 1024 / 1024).toFixed(2)} ميجابايت` : 'يمكنك اختيار الملف من الهاتف'}</small></span></button><input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={(event) => { setFile(event.target.files?.[0] || null); setMessage(''); }} /><button className="button secondary" disabled={drivePickerBusy || !config.drivePickerReady} onClick={pickFromDrive}>{drivePickerBusy ? 'جارٍ فتح Drive…' : 'اختيار من Google Drive'}</button><button className="button primary" disabled={busy || status.running || drivePickerBusy} onClick={startCheck}>ابدأ الفحص</button>{(busy || status.running) && <button className="button danger" onClick={stopCheck}>إيقاف الفحص</button>}{status.hasReport && !status.running && <button className="button success" onClick={downloadReport}>تنزيل تقرير Excel</button>}<button className="button secondary signout" onClick={logout}>تسجيل الخروج</button></div>
        {!config.drivePickerReady && <p className="muted">اختيار Drive يحتاج تفعيل Google Picker API وإضافة مفتاح API مقيّد. الرفع من الهاتف متاح الآن.</p>}
        <p className="message" role="status">{message || authError}</p>
      </section>

      {(status.total > 0 || status.running) && <section className="panel progress"><div className="section-title"><div><h2>حالة الفحص</h2><p>{status.running ? 'جارٍ فحص الحسابات…' : status.stopRequested ? 'تم إيقاف الفحص' : 'اكتمل الفحص'}</p></div><b className="percent">{status.total ? Math.round(status.done / status.total * 100) : 0}٪</b></div><progress max="100" value={status.total ? Math.round(status.done / status.total * 100) : 0} /><div className="stats"><div><span>تم فحصهم</span><b>{status.done} / {status.total}</b></div><div><span>نجح</span><b className="positive">{status.okCount}</b></div><div><span>تعذر فحصهم</span><b className="negative">{status.failCount}</b></div><div><span>المدة</span><b>{status.elapsedSec} ثانية</b></div></div></section>}

      <section className="panel results"><div className="section-title"><div><h2>النتائج</h2><p>تظهر النتائج المكتملة فقط، حتى عند إيقاف الفحص.</p></div><span className="chip">{(status.rows || []).filter(Boolean).length} طالب</span></div><div className="table-scroll"><table><thead><tr><th>#</th><th>الاسم</th><th>الفصل</th><th>رقم الطالب</th><th>الحالة</th><th>الفصول</th><th>النسبة</th><th>الألماس</th><th>آخر نشاط (القاهرة)</th></tr></thead><tbody>{(status.rows || []).filter(Boolean).length ? status.rows.map((row, index) => row && <tr key={`${row.studentId}-${index}`}><td>{index + 1}</td><td>{row.name}</td><td>{row.className}</td><td>{row.studentId}</td><td><span className={`badge ${row.status === 'ok' ? 'ok' : 'fail'}`}>{row.status === 'ok' ? 'ناجح' : row.status}</span></td><td>{row.chaptersCount}</td><td>{row.pct === '' ? '' : `${row.pct}٪`}</td><td>{row.diamonds}</td><td>{row.lastActivity}</td></tr>) : <tr><td className="empty" colSpan="9">لم تكتمل أي نتيجة بعد.</td></tr>}</tbody></table></div></section>
    </>}
    <footer>إعداد: مصطفى جابر حسين · مدرسة الترامسة التجارية · محافظة قنا</footer>
  </main>;
}

export default App;
