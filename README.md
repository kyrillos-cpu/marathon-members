# Marathon Members

Web app لمتابعة خدام MARATHON: الأعضاء، التيمز، الحضور بالـ QR، تقييم كل خادم، تسجيل الخدام في الميشنز، وسحب المواعيد المؤكدة من شيت الحجز.

**الواجهة:** GitHub Pages · **البيانات:** Google Sheet + Apps Script

---

## الملفات

| الملف | فين يتحط |
|---|---|
| `index.html` | GitHub |
| `config.js` | GitHub (حط فيه رابط الـ API) |
| `lib/qrcode.min.js` و `lib/html5-qrcode.min.js` | GitHub جوه فولدر `lib` |
| `Code.gs` | Apps Script بس (متحطوش على GitHub) |

---

## خطوات التركيب (حوالي ١٠ دقايق)

### ١) الشيت والـ Apps Script
1. من حساب **kyrillosmoussa** اعمل Google Sheet جديد واسمه `Marathon Members DB`.
2. من **Extensions ← Apps Script**.
3. امسح الكود اللي موجود والصق محتوى `Code.gs` كله.
4. في أول سطر غيّر `ADMIN_CODE = 'CHANGE-ME'` لكود سري ليكم انتوا القادة.
5. من **Project Settings ⚙** خلّي الـ Time zone = `(GMT+02:00) Cairo`.
6. اختار الدالة `setup` من فوق واضغط **Run**، ووافق على الصلاحيات.
   ← هيعمل التابات لوحده: Members · Teams · Sessions · Attendance · Missions · Signups · Bookings · Settings، ومعاهم الـ ٦ تيمز الأساسية.
7. **Deploy ← New deployment ← Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
8. انسخ رابط الـ **/exec**.

### ٢) GitHub
1. اعمل repo جديد اسمه `marathon-members` على حساب `kyrillos-cpu`.
2. ارفع: `index.html` و `config.js` وفولدر `lib` بالملفين اللي جواه.
3. افتح `config.js` والصق رابط الـ /exec مكان `PASTE_YOUR_APPS_SCRIPT_EXEC_URL_HERE`.
4. **Settings ← Pages ← Branch: main / root ← Save**.
5. الرابط هيبقى: `https://kyrillos-cpu.github.io/marathon-members/`

### ٣) ربط شيت الحجز (المواعيد المؤكدة)
1. من التطبيق: **Settings ← Bookings sheet** حط لينك شيت الحجز.
2. لازم الحساب اللي عامل الـ Apps Script يكون عنده صلاحية على شيت الحجز.
3. اضغط **Missions ← ⟳ Sync**. أي صف حالته "مؤكد / تم الاتفاق / Confirmed / Agreement finalized" هيتحول لميشن، والخدام يقدروا يسجلوا فيه.
4. لو عايزه يتزامن لوحده كل ساعة: من Apps Script شغّل الدالة `enableAutoSync` مرة واحدة.
5. لو التطبيق قرأ عمود غلط، اكتب اسم العمود بالظبط في **Settings ← Column names**.

> ⚠️ أي تعديل في `Code.gs` بعد كده: **Deploy ← Manage deployments ← ✏️ ← Version: New version ← Deploy**، عشان الرابط يفضل هو هو.

---

## الاستخدام

| مين | الرابط |
|---|---|
| القادة | `…/marathon-members/` + الكود السري |
| الخدام (التسجيل في الميشنز) | `…/marathon-members/?join=1` |
| كارت الخادم (QR) | `…/marathon-members/?card=MM-001` |

**الحضور:** Home ← *Sunday Meeting — scan now* ← وجّه الكاميرا على QR كل خادم. لو النت قطع، الحضور بيتحفظ على الموبايل ويترفع لوحده لما النت يرجع.

**التقييم (Score من 100):**
- **Attendance**: نسبة الحضور في اجتماع الحد والتدريب (التأخير = ٧٥٪). الأوزان بتتغير من الإعدادات.
- **Mission commitment**: الخادم اللي سجّل في ميشن، حضر فعلاً ولا لأ.
- **Participation**: النجوم (١–٥) اللي بتديها لكل خادم في كل جلسة.
- **Bonus**: +١ على كل ميشن زيادة حضرها (لحد +٥).
- المستويات: Star ≥ 85 · Active ≥ 70 · Average ≥ 50 · Needs follow-up.
