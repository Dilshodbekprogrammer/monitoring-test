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

## Real maktabda ishga tushirish
- Server: `waitress-serve --port=8000 app:app` (yoki gunicorn), oldiga nginx + HTTPS qo'ying
- Muhit o'zgaruvchilari: `SECRET_KEY`, `ADMIN_PASSWORD`, `DB_PATH`, `APP_TZ` (standart: Asia/Tashkent)
- Bazani muntazam nusxalang (`instance/monitoring.db`)

## Cheklovlar va keyingi qadamlar
- Brauzer orqali Alt+Tab'ni to'liq to'sib bo'lmaydi — tizim faqat oyna almashganini aniqlaydi. Telefon orqali aldashni sinfdagi nazorat bilan cheklash kerak
- O'quvchi parollari (boshlang'ich) o'qituvchi chop etib tarqatishi uchun ochiq saqlanadi; o'qituvchi/admin parollari esa shifrlangan
- Keyingi: bir akkauntdan bir vaqtda kirishni cheklash, savollarga rasm, ochiq (yozma) savollar, ota-onalar ko'rishi, choraklar bo'yicha dinamika grafigi
