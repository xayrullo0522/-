# Excel → Sayt avtomatik kirituvchi (Online Mahalla uchun)

Bu dastur Excel'dagi fuqarolar ma'lumotlarini saytga **avtomatik** kiritadi:
brauzerni ochadi, siz saytga kirasiz, keyin dastur har bir qator uchun formani ochib,
maydonlarni to'ldirib, "Saqlash"ni bosadi. Har bir qator natijasi alohida
`..._natija.xlsx` fayliga yoziladi (`OK` yoki `XATO: sababi`).

Dastur saytning o'ziga bog'lanmagan: qaysi maydonga nima yozilishi `config.yaml`
faylida sozlanadi. Shuning uchun Online Mahalla'da yoki boshqa istalgan saytda ishlaydi.

---

## 1. O'rnatish (bir marta)

1. **Python** o'rnating: https://www.python.org/downloads/
   (o'rnatishda **"Add Python to PATH"** belgisini qo'yishni unutmang).
2. Shu papkadagi **`1_ornatish.bat`** ni ikki marta bosing — kerakli kutubxonalar o'rnatiladi.

Kompyuterda Google Chrome bo'lsa, dastur o'shani ishlatadi.

## 2. Excel faylini qo'yish

- `Yangi_Pasport_Stol_4_5.xlsx` faylini shu papkaga qo'ying (boshqa nom bo'lsa — `config.yaml` da `fayl:` ni o'zgartiring).
- Dastur bu faylni **o'zgartirmaydi**. Natijalar alohida `Yangi_Pasport_Stol_4_5_natija.xlsx` fayliga yoziladi
  (qator raqami, F.I.Sh, JSHSHIR, `OK` yoki `XATO: sababi`).
- Faylda butun tuman (124 000+ fuqaro) bor. Faqat kerakli ko'cha/uylarni olish uchun `config.yaml` da `filtr:` ni yoqing:
  ```yaml
    filtr:
      TV_ALP_ADDRESS: "Г.ГУЛОМ"
  ```
- `JINS` ustuni JSHSHIR'dan avtomatik hisoblanadi (Erkak / Ayol).

## 3. Qanday ishlaydi — xonadonlar bo'yicha

1. Dastur xonadonlar ro'yxati sahifasini (`guruhlash.royxat_url`) ochib, undagi **barcha kadastr raqamlarini** yig'adi.
2. Excel'dagi odamlarni `TV_ALP_CADASTRE` bo'yicha xonadonlarga ajratadi va **faqat ro'yxatda bor xonadonlarni** oladi
   (butun tumandagi 32 000 xonadon emas).
3. Har bir xonadonni ochadi (kadastr raqamini bosadi) va Excel'da shu xonadonga yozilgan odamlarni (7–8 ta) bittadan kiritadi.
4. Ro'yxat bir necha sahifali bo'lsa, `config.yaml` da `keyingi_sahifa_selector` ni yozing
   (masalan `"li.next a"` — `5_selektor_aniqlash.bat` bilan "Keyingi" tugmasini ALT+bosib aniqlang).

## 4. Saytni "o'rgatish" — yozib olish (bir marta)

1. **`2_yozib_olish.bat`** ni ishga tushiring.
2. Brauzerda saytga kiring, so'ng qora oynada **Enter** bosing.
3. Dastur ro'yxatdagi xonadonlarni tekshirib, Excel'da odami bor xonadonlarni ko'rsatadi va namuna odam taklif qiladi — **Enter** bosing.
4. **1-bosqich:** ro'yxatda ko'rsatilgan kadastr raqamini bosing (xonadon ochiladi) → qora oynada **Enter**.
5. **2-bosqich:** shu xonadonga namunadagi odamni **qo'lda** kiriting ("Qo'shish", maydonlar, "Saqlash") → **Enter**.
   Qiymatlarni Excel'dagidek **aynan bir xil** yozing (masalan `IBRAIMOVA`, `15.07.1960`) —
   shunda dastur har bir maydonni qaysi Excel ustunidan olishni o'zi tushunadi.
6. `config_yozilgan.yaml` fayli yaratiladi. Bloknot bilan ochib tekshiring: `ustun:` bog'lanmagan joylarda
   `qiymat:` turadi — kerak bo'lsa tuzating. So'ng uni **`config.yaml`** nomi bilan saqlang (eskisining o'rniga).

> Qaysidir maydon to'g'ri yozilmasa — **`5_selektor_aniqlash.bat`**: ALT bosib turib maydonga bosing,
> qora oynada uning selektori chiqadi.

## 5. Sinab ko'rish

**`3_sinov_tasdiq_bilan.bat`** — dastur maydonlarni to'ldiradi, lekin har safar
saqlashdan oldin sizdan so'raydi. Avval 1–2 ta fuqaroda shunday tekshirib ko'ring.

Faqat ayrim qatorlarni ishlash:
```
python mahalla_bot.py --tekshir --qatorlar 2-20
```

## 6. Asosiy ish

**`4_ishga_tushirish.bat`** — barcha qatorlarni avtomatik kiritadi.

- Brauzer ochilgach saytga kiring (login / ERI / OneID), so'ng qora oynada **Enter** bosing.
- To'xtatish: qora oynada **Ctrl+C**.
- Qayta ishga tushirsangiz, `OK` bo'lgan qatorlar **o'tkazib yuboriladi** — faqat qolganlari va xatolar kiritiladi.
- Xato bo'lgan qatorlar rasmi `xatolar/` papkasiga saqlanadi.
- Login `brauzer_profil/` papkasida saqlanib qoladi, har safar qayta kirish shart bo'lmasligi mumkin.

> ⚠️ Ishlayotganda `..._natija.xlsx` faylini **yopib qo'ying**, aks holda natija yozilmaydi.

---

## `config.yaml` — qadamlar turlari

| amal         | nima qiladi                                     | asosiy parametrlar |
|--------------|--------------------------------------------------|--------------------|
| `ochish`     | sahifaga o'tadi                                  | `url` |
| `bosish`     | tugma/havolani bosadi                            | `selector` |
| `yozish`     | maydonga yozadi                                  | `selector`, `ustun` yoki `qiymat`, `sekin`, `enter`, `tab` |
| `tanlash`    | oddiy `<select>` ro'yxatdan tanlaydi             | `selector`, `ustun` |
| `royxat`     | maxsus ochiluvchi ro'yxat (viloyat, ko'cha...)   | `selector`, `ustun`, `variant_selector`, `yozib_qidirish` |
| `belgilash`  | checkbox / radio                                 | `selector`, (`ustun`: ha/yo'q) |
| `fayl`       | fayl yuklaydi                                    | `selector`, `ustun` (fayl yo'li) |
| `tugma`      | klaviatura tugmasi                               | `tugma: Enter` |
| `kutish`     | element chiqishini kutadi                        | `selector` yoki `url` |
| `pauza`      | shunchaki kutadi                                 | `soniya` |
| `tekshirish` | element ko'rinsa — qatorni XATO deb belgilaydi   | `selector` (xato xabari) |
| `qolda`      | siz qo'lda biror ish qilasiz (captcha, ERI)      | `xabar` |

Qo'shimcha parametrlar:

- `format: "%d.%m.%Y"` — sana ko'rinishi (`%Y-%m-%d` = 1990-05-12).
- `sekin: true` — maskali maydonlar (sana, telefon) uchun harfma-harf yozish.
- `almashtirish: {"E": "Erkak", "A": "Ayol"}` — Excel qiymatini saytdagisiga almashtirish.
- `majburiy: false` — shu qadam bajarilmasa ham qator davom etadi.
- `qiymat: "+998{Telefon}"` — `{Ustun nomi}` orqali bir nechta ustunni birlashtirish.
- Selektor turlari: CSS (`#id`, `input[name="x"]`), XPath (`//button[...]`) yoki matn (`text=Saqlash`).

## Muammolar

| Muammo | Yechim |
|--------|--------|
| `Kerakli kutubxona topilmadi` | `1_ornatish.bat` ni qayta ishga tushiring |
| `... ustuni yo'q` | Excel sarlavhasi `config.yaml`dagi `ustun:` bilan bir xil emas |
| `Timeout ... exceeded` | selektor noto'g'ri yoki sahifa sekin — `kutish_soniya` ni oshiring |
| Ro'yxatdan tanlanmayapti | `variant_selector` ni `--tanlash` rejimida variantni bosib aniqlang |
| Juda tez ishlayapti, sayt ulgurmayapti | `qadamlar_orasi`, `qatorlar_orasi` ni oshiring |

## Muhim

Dastur faqat sizning hisobingiz va ruxsatingiz doirasida, siz qo'lda bajaradigan
ishni avtomatlashtiradi. Fuqarolarning shaxsiy ma'lumotlari bo'lgan Excel fayllarini
boshqalarga bermang va internetga yuklamang (`.gitignore` ularni repoga qo'shmaydi).
