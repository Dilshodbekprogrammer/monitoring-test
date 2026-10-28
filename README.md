# Maktab monitoring test platformasi (MVP)

Choraklik monitoring testlarini avtomatlashtirish uchun veb-platforma.

## Imkoniyatlar
- **Rollar:** administrator, o'qituvchi, o'quvchi
- **O'qituvchi:** o'z fanini (bo'limini) tanlaydi, savollarni qo'lda kiritadi yoki **Excel/Word hujjatdan yuklaydi**
- **O'quvchilar:** ism-familiya kiritiladi → har biriga avtomatik **login va parol** yaratiladi (Excelga yuklab olinadi). O'quvchi faqat test ishlay oladi
- **Test sozlamalari:** savollar soni (bankdan tasodifiy), vaqt (daqiqa), boshlanish/tugash sanasi, urinishlar soni, savol/variantlarni aralashtirish, natijani ko'rsatish, ogohlantirishlar soni
- **Nazorat:** o'quvchi boshqa oyna/ilovaga (Alt+Tab) o'tsa ogohlantirish beriladi; 3-marta (sozlanadi) test avtomatik yakunlanadi. Barcha holatlar bazaga yoziladi
- **Natijalar:** o'qituvchi profilida jadval, statistika, har bir o'quvchi javoblarining tahlili; **Excelga yuklash** (+ ishlamaganlar ro'yxati)
- Taymer server vaqtiga asoslangan, javoblar har bosilganda saqlanadi, vaqt tugasa test avtomatik yopiladi

## Obuna va to'lov (HamyonAPI)
- O'qituvchi o'zi ro'yxatdan o'tadi (`/register`) va **7 kun bepul** sinaydi (`TRIAL_DAYS`).
- Keyin oylik obuna (`PRICE_MONTH`, standart 49 000 so'm). 1/3/6/12 oylik to'lash mumkin.
- To'lov: tizim HamyonAPI orqali to'lov yaratadi → o'qituvchiga karta va **aniq summa** ko'rsatiladi → pul o'tkazadi → HamyonAPI `complete_url` ga xabar yuboradi → obuna avtomatik uzayadi.
- Callback **imzo (md5) va summa** bo'yicha tekshiriladi, bir to'lov ikki marta hisoblanmaydi. Kalit (`shop_key`) faqat serverda.
- Obuna tugasa: yangi savol, o'quvchi, test qo'shish yopiladi; **eski natijalarni ko'rish va Excelga yuklash ochiq**. O'quvchilar yangi test boshlay olmaydi.
- Har bir o'qituvchi faqat o'z o'quvchilari va natijalarini ko'radi.
- Admin: o'qituvchiga qo'lda obuna beradi, barcha to'lovlar va callback jurnalini ko'radi (`/admin/payments`).

**HamyonAPI sozlamalari** (muhit o'zgaruvchilari):
```
HAMYON_SHOP_ID=12
HAMYON_SHOP_KEY=sk_xxxxxxxx
PRICE_MONTH=49000
TRIAL_DAYS=7
SECRET_KEY=uzun-tasodifiy-satr
```
HamyonAPI kabinetida **prepare_url** va **complete_url** ga quyidagilarni kiriting (domeningizga almashtiring, HTTPS bo'lsin):
```
prepare_url:  https://SIZNING-DOMEN/hooks/hamyon/prepare
complete_url: https://SIZNING-DOMEN/hooks/hamyon/complete
```
Callback ishlamay qolsa, o'qituvchi "To'lovni qayta tekshirish" tugmasi bilan to'lovni qo'lda tasdiqlay oladi.

## Ishga tushirish
```bash
pip install -r requirements.txt
python app.py                 # http://localhost:5000
```
Birinchi ishga tushishda administrator yaratiladi: **admin / admin123** (kirgach almashtiring yoki `ADMIN_PASSWORD` muhit o'zgaruvchisini bering).

Demo ma'lumotlar bilan sinab ko'rish (ixtiyoriy):
```bash
python seed_demo.py           # o'qituvchi: oqituvchi / oqituvchi123, 7-A sinf, 10 ta savol
```

Ish jarayoni:
1. `admin` bilan kiring → fan qo'shing → o'qituvchi yarating va fanini biriktiring
2. O'qituvchi bilan kiring → savollar kiriting/yuklang → "O'quvchilar" bo'limida sinf va ism-familiyalarni kiriting → login-parollarni Excelga yuklab tarqating
3. "Testlar" → yangi test (savollar soni, vaqt, sinf) → o'quvchilar kirib ishlaydi → "Natijalar" va Excel

## Hujjatdan savol yuklash formatlari
- **Excel (.xlsx):** `Savol | A | B | C | D | To'g'ri javob` (shablonni tizimdan yuklab oling)
- **Word (.docx) jadval:** shu 6 ustun
- **Word (.docx) matn:**
  ```
  1. Savol matni
  A) variant
  B) variant
  C) variant
  D) variant
  Javob: B
  ```

## Savol va variantlarda rasm
Savol matni, har bir variant — matn, rasm (JPG/PNG) yoki ikkalasi bo'lishi mumkin. Rasmni fayl sifatida tanlash yoki nusxalab **Ctrl+V** bilan qo'yish mumkin (skrinshot ham). Rasmlar tekshiriladi (haqiqiy JPG/PNG bo'lishi shart), 1600 px gacha kichraytiriladi, har biri 6 MB dan oshmasligi kerak. Rasmlar `instance/uploads/` (yoki `UPLOAD_DIR`) papkasida saqlanadi — **bazadan tashqari bu papkani ham nusxalang**.

## Serverga joylash (internetda)
**Tavsiya etilgan buyruq (Linux server):**
```
pip install -r requirements.txt
SECRET_KEY=uzun-tasodifiy-satr ADMIN_PASSWORD=kuchli-parol COOKIE_SECURE=1 \
  gunicorn -w 3 -b 127.0.0.1:8000 app:app
```
oldiga nginx + HTTPS (Let's Encrypt) qo'ying. nginx'da **`client_max_body_size 32m;`** bo'lishi shart (standart 1 MB — rasm yuklanmaydi), va `proxy_set_header X-Forwarded-For $remote_addr; proxy_set_header X-Forwarded-Proto $scheme; proxy_set_header Host $host;`.

**cPanel ("Setup Python App")**: `passenger_wsgi.py` tayyor (kirish nuqtasi `application`). Muhit o'zgaruvchilarini shu yerda kiriting.

**Muhim qoidalar (aks holda foydalanuvchilar "tizimdan chiqib ketaveradi" yoki ma'lumot yo'qoladi):**
1. **`SECRET_KEY`ni albatta bering.** Bir nechta worker bo'lsa hammasi bir xil kalit ishlatishi kerak (kod buni o'zi ham ta'minlaydi, lekin muhit o'zgaruvchisi eng ishonchlisi).
2. **Baza va rasmlar doimiy diskda bo'lsin**: `DB_PATH=/data/monitoring.db`, `UPLOAD_DIR=/data/uploads`. Vaqtinchalik (ephemeral) diskli xostinglarda (Render/Heroku bepul tarif, Vercel va h.k.) har restartda hamma narsa o'chadi — ishlatmang yoki doimiy disk ulang.
3. Admin parolini darhol almashtiring (`admin123` qolsa, admin panelda qizil ogohlantirish chiqadi).
4. Sozlamalarni tekshirish: brauzerda `https://DOMEN/healthz` oching va bir necha marta yangilang — `secret_fp` har safar **bir xil** bo'lishi kerak; `proxy_headers: true`, `db_writable: true`, `uploads_writable: true` bo'lsin; `scheme` — `https`.
5. Bazani muntazam nusxalang: `python backup.py` (cron bilan kuniga 1 marta). `backups/` va `instance/uploads/` ni ham serverdan tashqariga ko'chiring.

Muhit o'zgaruvchilari: `SECRET_KEY`, `ADMIN_PASSWORD`, `DB_PATH`, `UPLOAD_DIR`, `APP_TZ` (standart Asia/Tashkent), `COOKIE_SECURE=1` (HTTPS bo'lsa), `TRUST_PROXY=1` (standart), `PRICE_MONTH`, `TRIAL_DAYS`, `HAMYON_SHOP_ID`, `HAMYON_SHOP_KEY`.

**Quvvat:** SQLite bilan 2–4 worker va bir vaqtda 100–200 o'quvchi odatda muammosiz. Undan ko'p bo'lsa PostgreSQL'ga o'tish tavsiya etiladi.

## Cheklovlar va keyingi qadamlar
- Brauzer orqali Alt+Tab'ni to'liq to'sib bo'lmaydi — tizim faqat oyna almashganini aniqlaydi. Telefon orqali aldashni sinfdagi nazorat bilan cheklash kerak
- O'quvchi parollari (boshlang'ich) o'qituvchi chop etib tarqatishi uchun ochiq saqlanadi; o'qituvchi/admin parollari esa shifrlangan
- Keyingi: bir akkauntdan bir vaqtda kirishni cheklash, savollarga rasm, ochiq (yozma) savollar, ota-onalar ko'rishi, choraklar bo'yicha dinamika grafigi
