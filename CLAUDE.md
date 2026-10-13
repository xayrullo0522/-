# Mahalla bot — loyiha haqida (Claude uchun)

Foydalanuvchi bilan **o'zbek tilida** gaplash. U dasturchi emas: qisqa, aniq, qadamma-qadam tushuntir.

## Nima qiladi
`mahalla_bot.py` — Excel'dagi fuqarolarni **online-mahalla.uz** saytiga avtomatik kiritadi (Playwright + Chrome).

- Excel: `Yangi_Pasport_Stol_4_5.xlsx`, varaq `Лист1`, ~124 000 qator (butun tuman). Ustunlar:
  `TV_PINPP` (JSHSHIR), `TV_SURNAMECIRILLIC/LATIN`, `TV_NAMECIRILLIC/LATIN`, `TV_PATRONYMCIRILLIC/LATIN`,
  `TV_DATEBIRTH` (dd.mm.yyyy matn), `TV_BIRTHPLACELATIN`, `TV_SERIALNUMBER` (pasport),
  `TV_ALP_CADASTRE` (xonadon kadastri), `TV_ALP_ADDRESS`, `TV_REGISTRATIONDATE`.
  Dastur `JINS` (Erkak/Ayol) ustunini JSHSHIR 1-raqamidan hisoblaydi.
- Sayt jarayoni: xonadonlar ro'yxati
  `https://online-mahalla.uz/tables/survey_homes_street?obl_id=33&area_id=3306&district_id=3306020&street_id=330600447&_mid=3548`
  → kadastr raqami bosiladi → xonadon ochiladi → unga odamlar (7–8 ta) qo'shiladi.
  Xonadon sahifasi ko'rinishi: `/forms/survey_homes/<id>?obl_id=...&street_id=...`

## Tuzilishi
- `config.yaml` — hamma sozlama. `guruhlash` bo'limi: `ustun` (TV_ALP_CADASTRE), `royxat_url`,
  `keyingi_sahifa_selector`, `xonadon_qadamlar` (xonadonni ochish). `qadamlar` — bitta odamni kiritish.
  Qadam turlari: ochish, bosish, yozish, tanlash, royxat, belgilash, fayl, tugma, kutish, pauza, tekshirish, qolda.
  `{Ustun}` — selektor/qiymatga Excel qiymatini qo'yadi. Selektorlar Playwright formatida (CSS, `text=...`, XPath).
- Natija: `Yangi_Pasport_Stol_4_5_natija.xlsx` (qator, F.I.Sh, Holat OK/XATO). Manba Excel o'zgartirilmaydi.
  OK bo'lgan qatorlar qayta ishga tushirganda o'tkaziladi. Xato rasmlari: `xatolar/`.
- Rejimlar: `--yozib_olish` (qo'lda kiritilganni yozib config yaratadi), `--tanlash` (ALT+bosish → selektor),
  `--tekshir` (saqlashdan oldin so'raydi), `--qatorlar 2-10`, `--hammasi`.
- Brauzer profili: `brauzer_profil/` (login saqlanadi). Bir vaqtda ikkita bot oynasi ochilmasin.

## Qoidalar
- Saytga login/ERI/OneID bilan **foydalanuvchi o'zi** kiradi — parol so'rama.
- Saytda "Saqlash"ni bosishdan oldin har doim foydalanuvchidan tasdiq ol; avval `--tekshir` bilan 1 ta odamda sina.
- Bu fuqarolarning shaxsiy ma'lumotlari: Excel/natija fayllarini hech qayerga yuklama, chiqishda ko'p qatorni ko'rsatma.
- Kodni o'zgartirgandan keyin `python -c "import ast; ast.parse(open('mahalla_bot.py', encoding='utf-8').read())"` bilan tekshir.
