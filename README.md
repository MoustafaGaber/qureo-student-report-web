# QUREO Student Report — Web Version

واجهة React وخادم Express لفحص تقدم الطلاب، مع تسجيل دخول Google عبر Firebase Authentication. نسخة Windows الأصلية على فرع `win-x86` منفصلة ولم تُعدّل.

## إعداد مجاني بلا بطاقة

1. أنشئ مشروع Firebase واختر خطة **Spark** المجانية. تتيح تسجيل Google دون إضافة وسيلة دفع.
2. من Firebase Console فعّل **Authentication → Google**، ثم أضف النطاق المحلي أو نطاق الاستضافة ضمن Authorized domains.
3. أنشئ Web App في Firebase وانسخ قيم `apiKey` و`authDomain` و`projectId` و`appId`.
4. انسخ `.env.example` إلى `.env` وضع القيم، ثم أضف بريد المعلم المسموح في `ALLOWED_EMAILS`، أو نطاق المدرسة في `ALLOWED_EMAIL_DOMAINS`. اترك القائمة فارغة إلى أن تحدد الحسابات المصرح بها؛ عندها سيرفض الخادم تشغيل الفحص.
5. من **Project settings → Service accounts** أنشئ مفتاح Firebase Admin الخاص بالخادم، واحفظ ملف JSON محليًا باسم `secrets/firebase-service-account.json`. هذا المفتاح سري وله صلاحيات إدارية؛ لا ترسله في المحادثة ولا ترفعه إلى Git. المجلد `secrets/` مستثنى من Git. لا يلزم إضافة بطاقة دفع لإنشاء مفتاح الخدمة.

هذا الفرع مخصص لتجربة Vercel. خطة Hobby مجانية وفق شروط Vercel ومخصصة للمشاريع الشخصية أو غير التجارية؛ راجع الشروط الحالية قبل استخدامه في نطاق المدرسة.

## التشغيل محليًا

```powershell
npm install
npm run build
npm start
```

افتح `http://127.0.0.1:3456`. يلزم إعداد Firebase في `.env`، وملف اعتماد Admin في المسار المحدد، وتسجيل الحسابات المصرح بها قبل استخدام الفحص. لا ترفع `.env` أو ملف اعتماد Admin إلى Git.

## اختيار ملف من Google Drive

1. فعّل **Google Picker API** في مشروع Google Cloud المرتبط بمشروع Firebase.
2. أنشئ API key منفصلًا لـPicker، وقيّده بواجهة **Google Picker API** و**Google Drive API**، وبـHTTP referrers الخاصة بالموقع و`https://docs.google.com/*`.
3. ضع المفتاح في `VITE_GOOGLE_PICKER_API_KEY` و`967333099086` في `VITE_FIREBASE_MESSAGING_SENDER_ID` داخل `.env`.
4. أعد تشغيل الخادم. عند اختيار الملف، يطلب التطبيق إذن `drive.file` الذي يتيح للمستخدم اختيار الملفات التي يشاركها مع التطبيق عبر Picker.

إنشاء المفتاح غير مطلوب لتسجيل الدخول أو رفع ملف من الهاتف. لا تربط Billing أو خطة مدفوعة؛ إذا طلبت لوحة Google تفعيل الفوترة، اترك Picker غير مفعّل واستخدم رفع الهاتف.

## السلوك والحدود

- الملفات: XLSX وXLS وCSV، حتى 8 ميجابايت محليًا. على Vercel Hobby حد الرفع مضبوط على 4 ميجابايت لتفادي حد جسم الطلب في المنصة، و500 طالب كحد أعلى.
- ست عمليات فحص بالتوازي، و42 فصلًا في المنهج.
- تقرير Excel بورقتي `بيانات الطلاب` و`ملخص الإحصائيات`.
- يرسل الخادم التقدم في اتصال واحد. لا توجد مهمة محفوظة في ذاكرة مشتركة بين الطلبات؛ الإيقاف يغلق الاتصال، وتظل النتائج التي وصلت للمتصفح قابلة للتنزيل كتقرير جزئي.
- لا يعيد الخادم كلمات المرور إلى الواجهة ولا يكتب ملفات الطلاب أو التقرير إلى القرص؛ تبقى بيانات الاعتماد في الذاكرة أثناء الفحص. ملف اعتماد Firebase Admin منفصل وسري، ويحفظ محليًا خارج Git أو كمتغير سري على Vercel.
- يرسل التطبيق بيانات الدخول إلى QUREO. يلزم تصريح المدرسة ومراعاة شروط المنصة وسياساتها.
- دوال Vercel Hobby محدودة بمدة تنفيذ قصوى؛ إذا طال الفحص حتى انتهت المهلة، سيحتفظ المتصفح بالنتائج التي وصلته فقط. لا يمكن ضمان إتمام دفعة 500 طالب في طلب واحد.

## النشر

### النشر على Vercel

1. ارفع فرع `vercel-version` إلى مستودع GitHub ثم اربط المستودع في Vercel. تُبنى واجهة Vite إلى `client/dist`، ويحوّل `api/[...path].js` طلبات `/api/*` إلى تطبيق Express الموجود في `app.js`.
2. أضف Environment Variables من لوحة Vercel: `FIREBASE_PROJECT_ID`, `FIREBASE_SERVICE_ACCOUNT_JSON`, `ALLOWED_EMAILS`, `VITE_FIREBASE_API_KEY`, `VITE_FIREBASE_AUTH_DOMAIN`, `VITE_FIREBASE_PROJECT_ID`, `VITE_FIREBASE_APP_ID`, `VITE_FIREBASE_MESSAGING_SENDER_ID`, و`VITE_GOOGLE_PICKER_API_KEY` إذا كنت تستخدم Drive. ضع JSON كاملًا لحساب Firebase Admin داخل المتغير السري `FIREBASE_SERVICE_ACCOUNT_JSON`.
3. أضف نطاق `*.vercel.app` أو اسم النشر الفعلي إلى Firebase Authentication → Authorized domains. حدّث قيود HTTP referrer لمفتاح Picker عند استخدام Drive.
4. ابدأ الاختبار بملف وهمي صغير. راجع قيود مدة دوال Vercel وقيود حجم الطلب، ولا ترفع ملف `.env` أو JSON حساب الخدمة إلى GitHub.

#### إصلاح تسجيل الدخول بالتحويل على متصفحات الهاتف

إذا عاد Google إلى الصفحة الرئيسية دون جلسة على الهاتف، يستخدم المشروع تحويلات Vercel لمسارات `/__/auth/*` و`/__/firebase/*` إلى `qureo-student-report.firebaseapp.com`. في إعداد Vercel غيّر `VITE_FIREBASE_AUTH_DOMAIN` إلى `qureo-student-report-web.vercel.app` ثم أعد النشر. وفي Google Cloud Console أضف URI التالي إلى OAuth Client ID → Authorized redirect URIs، مع إبقاء URI Firebase الحالي:

```text
https://qureo-student-report-web.vercel.app/__/auth/handler
```

أبقِ نطاق Vercel ضمن Firebase Authentication → Authorized domains. يعتمد هذا الإعداد على توجيه عكسي شفاف، لا على تحويل HTTP 302؛ وهو المسار الذي توصي به [وثائق Firebase](https://firebase.google.com/docs/auth/web/redirect-best-practices) للواجهات المستضافة خارج Firebase Hosting عند استخدام `signInWithRedirect`.

إعداد Vercel مبني على واجهة ثابتة ودالة API واحدة: تقدم الفحص والنتائج تصل مباشرة للمتصفح، والتقرير يُنشأ داخل المتصفح. حد الملف 4 ميجابايت في استضافة Vercel المجانية بسبب حد الطلب المنشور للمنصة، ومدة الدالة تضبط إلى 300 ثانية؛ قد تتوقف دفعة طويلة عند انتهاء المهلة مع بقاء النتائج التي وصلت فقط. [حدود دوال Vercel](https://vercel.com/docs/functions/limitations) · [مدة التنفيذ](https://vercel.com/docs/functions/configuring-functions/duration)
