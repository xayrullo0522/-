# Excel → Sayt avtomatik kirituvchi (Online Mahalla uchun)

Bu dastur Excel'dagi fuqarolar ma'lumotlarini saytga **avtomatik** kiritadi:
brauzerni ochadi, siz saytga kirasiz, keyin dastur har bir qator uchun formani ochib,
maydonlarni to'ldirib, "Saqlash"ni bosadi. Har bir qator natijasi Excel'ning
**Holat** ustuniga yoziladi (`OK` yoki `XATO: sababi`).

Dastur saytning o'ziga bog'lanmagan: qaysi maydonga nima yozilishi `config.yaml`
faylida sozlanadi. Shuning uchun Online Mahalla'da yoki boshqa istalgan saytda ishlaydi.

---

## 1. O'rnatish (bir marta)

1. **Python** o'rnating: https://www.python.org/downloads/
   (o'rnatishda **"Add Python to PATH"** belgisini qo'yishni unutmang).
2. Shu papkadagi **`1_ornatish.bat`** ni ikki marta bosing.
   U kerakli kutubxonalarni o'rnatadi va namuna `fuqarolar.xlsx` yaratadi.

Kompyuterda Google Chrome bo'lsa, dastur o'shani ishlatadi.

## 2. Excel faylini tayyorlash

- Faylni shu papkaga `fuqarolar.xlsx` nomi bilan qo'ying (yoki `config.yaml`da nomini o'zgartiring).
- **1-qator — ustun nomlari** (Familiya, Ism, Tug'ilgan sana, JSHSHIR, ...).
- Ustun nomlari `config.yaml` dagi `ustun:` qiymatlari bilan **aynan bir xil** bo'lishi kerak.

## 3. Saytdagi maydonlarni aniqlash (bir marta)

Har bir saytning maydonlari turlicha nomlangan. Ularni aniqlash uchun:

1. **`2_selektor_aniqlash.bat`** ni ishga tushiring.
2. Ochilgan brauzerda saytga kiring va fuqaro qo'shish formasini oching.
3. Klaviaturada **ALT** ni bosib turib, kerakli maydonga **sichqoncha bilan bosing**.
4. Qora oynada tayyor qator chiqadi, masalan:
   ```yaml
     - amal: yozish
       selector: 'input[name="lastName"]'
   ```
5. Shu qatorni `config.yaml` ning `qadamlar:` bo'limiga ko'chirib, `ustun: "Familiya"` qo'shing.

Shu tarzda har bir maydon, ochiluvchi ro'yxat va "Saqlash" tugmasini aniqlab chiqasiz.

## 4. Sinab ko'rish

**`3_sinov_tasdiq_bilan.bat`** — dastur maydonlarni to'ldiradi, lekin har safar
saqlashdan oldin sizdan so'raydi. Avval 1–2 ta fuqaroda shunday tekshirib ko'ring.

Faqat ayrim qatorlarni ishlash:
```
python mahalla_bot.py --tekshir --qatorlar 2-3
```

## 5. Asosiy ish

**`4_ishga_tushirish.bat`** — barcha qatorlarni avtomatik kiritadi.

- Brauzer ochilgach saytga kiring (login / ERI / OneID), so'ng qora oynada **Enter** bosing.
- To'xtatish: qora oynada **Ctrl+C**.
- Qayta ishga tushirsangiz, `OK` bo'lgan qatorlar **o'tkazib yuboriladi** — faqat qolganlari va xatolar kiritiladi.
- Xato bo'lgan qatorlar rasmi `xatolar/` papkasiga saqlanadi.
- Login `brauzer_profil/` papkasida saqlanib qoladi, har safar qayta kirish shart bo'lmasligi mumkin.

> ⚠️ Ishlayotganda Excel faylini **yopib qo'ying**, aks holda natija yozilmaydi.

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
