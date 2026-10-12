"""
Excel -> veb-sayt avtomatik kirituvchi dastur ("Online Mahalla" va shunga o'xshash saytlar uchun).

Ishlash tartibi:
  1. Brauzer ochiladi, siz saytga o'zingiz kirasiz (login / ERI / OneID).
  2. Dastur Excel faylidagi har bir qatorni o'qiydi.
  3. config.yaml dagi "qadamlar" ro'yxati bo'yicha har bir qator uchun
     formani ochadi, maydonlarni to'ldiradi va saqlaydi.
  4. Har bir qator natijasi Excel'dagi "Holat" ustuniga yoziladi
     (OK / XATO). Qayta ishga tushirsangiz, OK bo'lgan qatorlar o'tkazib yuboriladi.

Foydalanish:
  python mahalla_bot.py                 # asosiy ish (config.yaml bo'yicha)
  python mahalla_bot.py --yozib_olish 2 # 2-qatordagi fuqaroni qo'lda kiritasiz, dastur qadamlarni yozadi
  python mahalla_bot.py --tanlash       # sahifadagi elementlar selektorini aniqlash rejimi
  python mahalla_bot.py --tekshir       # saqlashdan oldin har safar tasdiq so'raydi
  python mahalla_bot.py --qatorlar 2-10 # faqat 2..10-qatorlarni ishlash
  python mahalla_bot.py --config boshqa.yaml
"""

from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
import time
from pathlib import Path

try:
    import yaml
    from openpyxl import load_workbook
    from playwright.sync_api import Error as PWError
    from playwright.sync_api import TimeoutError as PWTimeout
    from playwright.sync_api import sync_playwright
except ImportError as e:  # pragma: no cover
    print(f"Kerakli kutubxona topilmadi: {e.name}")
    print("Avval o'rnating:  pip install -r requirements.txt")
    sys.exit(1)


BAZA = Path(__file__).resolve().parent
# Windows konsolida o'zbek harflari/belgilar sabab dastur yiqilmasin
for _oqim in (sys.stdout, sys.stderr):
    try:
        _oqim.reconfigure(errors="replace")
    except AttributeError:
        pass
SHABLON = re.compile(r"\{([^{}]+)\}")


# --------------------------------------------------------------------------- #
#  Yordamchi funksiyalar
# --------------------------------------------------------------------------- #

def chiqar(matn: str = "") -> None:
    print(matn, flush=True)


def config_oqish(yol: Path) -> dict:
    if not yol.exists():
        chiqar(f"Sozlama fayli topilmadi: {yol}")
        sys.exit(1)
    with open(yol, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def qiymat_matnga(qiymat, fmt: str | None = None) -> str:
    """Excel katagidagi qiymatni saytga yoziladigan matnga aylantiradi."""
    if qiymat is None:
        return ""
    if isinstance(qiymat, (dt.datetime, dt.date)):
        return qiymat.strftime(fmt or "%d.%m.%Y")
    if isinstance(qiymat, float) and qiymat.is_integer():
        return str(int(qiymat))
    matn = str(qiymat).strip()
    if fmt and matn:
        # Matn ko'rinishidagi sanani ham formatlashga urinib ko'ramiz
        for kirish in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                return dt.datetime.strptime(matn, kirish).strftime(fmt)
            except ValueError:
                pass
    return matn


def shablon_toldirish(matn: str, qator: dict) -> str:
    """'{Ustun nomi}' ko'rinishidagi joylarni qator qiymatlari bilan almashtiradi."""
    def almashtir(m):
        nom = m.group(1)
        if nom not in qator:
            raise KeyError(f"Excel'da '{nom}' ustuni yo'q")
        return qiymat_matnga(qator[nom])
    return SHABLON.sub(almashtir, matn)


def qatorlar_oraligi(matn: str | None):
    if not matn:
        return None
    toplam = set()
    for qism in matn.split(","):
        qism = qism.strip()
        if "-" in qism:
            a, b = qism.split("-", 1)
            toplam.update(range(int(a), int(b) + 1))
        elif qism:
            toplam.add(int(qism))
    return toplam


# --------------------------------------------------------------------------- #
#  Excel bilan ishlash
# --------------------------------------------------------------------------- #

def ustun_harfi(n: int) -> str:
    harf = ""
    while n:
        n, q = divmod(n - 1, 26)
        harf = chr(65 + q) + harf
    return harf


class ExcelJadval:
    """Manba Excel faqat o'qiladi (u o'zgarmaydi). Natijalar alohida kichik faylga yoziladi:
    <fayl>_natija.xlsx — har qator uchun Holat. Katta fayllarda (100 000+ qator) ham tez ishlaydi."""

    def __init__(self, sozlama: dict):
        self.fayl = (BAZA / sozlama.get("fayl", "fuqarolar.xlsx")).resolve()
        if not self.fayl.exists():
            chiqar(f"Excel fayl topilmadi: {self.fayl}")
            sys.exit(1)
        natija = sozlama.get("natija_fayl") or f"{self.fayl.stem}_natija.xlsx"
        self.natija_fayl = (self.fayl.parent / natija).resolve()
        self.sarlavha_qatori = int(sozlama.get("sarlavha_qatori", 1))
        self.korsatish = list(sozlama.get("korsatish_ustunlari") or [])
        self.filtr = sozlama.get("filtr") or {}
        self.jshshir_ustuni = sozlama.get("jshshir_ustuni")

        chiqar(f"Excel o'qilmoqda: {self.fayl.name} ...")
        wb = load_workbook(self.fayl, read_only=True, data_only=True)
        varaq = sozlama.get("varaq")
        ws = wb[varaq] if varaq else wb.worksheets[0]

        self.ustunlar: list[str] = []
        self._qatorlar: list[tuple[int, dict]] = []
        for r, qiymatlar in enumerate(ws.iter_rows(values_only=True), 1):
            if r < self.sarlavha_qatori:
                continue
            if r == self.sarlavha_qatori:
                self.ustunlar = [str(v).strip() if v not in (None, "") else ustun_harfi(i)
                                 for i, v in enumerate(qiymatlar, 1)]
                continue
            if not self.ustunlar:  # sarlavha_qatori: 0 — sarlavha yo'q, ustunlar A, B, C ...
                self.ustunlar = [ustun_harfi(i) for i in range(1, len(qiymatlar) + 1)]
            if all(v is None or str(v).strip() == "" for v in qiymatlar):
                continue
            qator = dict(zip(self.ustunlar, qiymatlar))
            self.hisoblash(qator)
            if self.filtrdan_otadi(qator):
                self._qatorlar.append((r, qator))
        wb.close()

        for nom in list(self.filtr) + self.korsatish:
            if nom not in self.ustunlar:
                chiqar(f"Diqqat: Excel'da '{nom}' ustuni yo'q. Bor ustunlar: {', '.join(self.ustunlar)}")

        self.holatlar: dict[int, str] = {}
        self._natija_ochish()

    def hisoblash(self, qator: dict) -> None:
        """JSHSHIR'dan qo'shimcha ustunlar: JINS (Erkak/Ayol). JSHSHIR 1-raqami: toq = erkak, juft = ayol."""
        if not self.jshshir_ustuni:
            return
        j = qiymat_matnga(qator.get(self.jshshir_ustuni))
        qator["JINS"] = ("Erkak" if int(j[0]) % 2 else "Ayol") if j[:1].isdigit() and j[0] != "0" else ""

    def filtrdan_otadi(self, qator: dict) -> bool:
        """filtr: {ustun: "matn"} yoki {ustun: ["matn1", "matn2"]} — ustunda shu matn bo'lsa (katta-kichik harf farqsiz)."""
        for nom, kerak in self.filtr.items():
            qiymat = qiymat_matnga(qator.get(nom)).upper()
            variantlar = kerak if isinstance(kerak, list) else [kerak]
            if not any(str(v).upper() in qiymat for v in variantlar):
                return False
        return True

    def _natija_ochish(self) -> None:
        from openpyxl import Workbook
        sarlavha = ["Qator"] + self.korsatish + ["Holat", "Vaqt"]
        if self.natija_fayl.exists():
            self.nwb = load_workbook(self.natija_fayl)
            self.nws = self.nwb.active
            for qiymatlar in self.nws.iter_rows(min_row=2, values_only=True):
                if qiymatlar and isinstance(qiymatlar[0], int):
                    holat = qiymatlar[-2] if len(qiymatlar) >= 2 else ""
                    self.holatlar[qiymatlar[0]] = str(holat or "")
        else:
            self.nwb = Workbook()
            self.nws = self.nwb.active
            self.nws.title = "Natija"
            self.nws.append(sarlavha)
        self._qator_joyi = {self.nws.cell(row=i, column=1).value: i for i in range(2, self.nws.max_row + 1)}

    def qatorlar(self):
        """(Excel qator raqami, {ustun: qiymat}) — filtrdan o'tgan, bo'sh bo'lmagan qatorlar."""
        return list(self._qatorlar)

    def holat(self, r: int) -> str:
        return self.holatlar.get(r, "")

    def holat_yoz(self, r: int, qator: dict, matn: str) -> None:
        self.holatlar[r] = matn
        qiymatlar = [r] + [qiymat_matnga(qator.get(u)) for u in self.korsatish] + [matn, f"{dt.datetime.now():%Y-%m-%d %H:%M:%S}"]
        joy = self._qator_joyi.get(r)
        if joy is None:
            self.nws.append(qiymatlar)
            self._qator_joyi[r] = self.nws.max_row
        else:
            for c, v in enumerate(qiymatlar, 1):
                self.nws.cell(row=joy, column=c, value=v)
        self.saqlash()

    def saqlash(self) -> None:
        try:
            self.nwb.save(self.natija_fayl)
        except PermissionError:
            chiqar(f"  ! {self.natija_fayl.name} saqlanmadi: fayl Excel'da ochiq bo'lsa, yoping.")


# --------------------------------------------------------------------------- #
#  Brauzer bilan ishlash
# --------------------------------------------------------------------------- #

def _ishga_tushir(pw, kanal, umumiy: dict):
    """Brauzerni ochadi va haqiqatan ishlayotganini tekshiradi (darhol yopilib qolmadimi)."""
    context = pw.chromium.launch_persistent_context(channel=kanal, **umumiy) if kanal \
        else pw.chromium.launch_persistent_context(**umumiy)
    try:
        time.sleep(1)
        sahifa = context.pages[0] if context.pages else context.new_page()
        sahifa.evaluate("1")
        return context
    except PWError:
        try:
            context.close()
        except PWError:
            pass
        raise


def brauzer_ochish(pw, sozlama: dict):
    asosiy_profil = (BAZA / sozlama.get("profil_papkasi", "brauzer_profil")).resolve()
    kanal = sozlama.get("kanal", "chrome")
    kanallar = [kanal, None] if kanal else [None]
    oxirgi_xato = None
    # Profil band bo'lsa (avvalgi oyna ochiq qolgan) — keyingi zaxira profil bilan ochamiz
    for n in range(4):
        profil = asosiy_profil if n == 0 else asosiy_profil.with_name(f"{asosiy_profil.name}_{n + 1}")
        profil.mkdir(parents=True, exist_ok=True)
        umumiy = dict(
            user_data_dir=str(profil),
            headless=bool(sozlama.get("yashirin", False)),
            viewport=None,
            args=["--start-maximized"] + list(sozlama.get("qoshimcha_args", [])),
            slow_mo=int(sozlama.get("sekinlashtirish_ms", 0)),
        )
        for k in kanallar:
            try:
                context = _ishga_tushir(pw, k, umumiy)
                if n:
                    chiqar(f"Diqqat: '{asosiy_profil.name}' band edi (avvalgi brauzer oynasi ochiq qolgandir),"
                           f" '{profil.name}' profili ishlatildi — saytga qaytadan kirish kerak bo'lishi mumkin.")
                if k != kanal:
                    chiqar(f"'{kanal}' ochilmadi, Playwright Chromium ishlatildi.")
                return context
            except PWError as e:
                oxirgi_xato = e
    chiqar("Brauzer ochilmadi. Barcha Chrome oynalarini yoping va qaytadan urinib ko'ring.")
    chiqar(f"Xato: {str(oxirgi_xato).splitlines()[0] if oxirgi_xato else ''}")
    chiqar("Agar 'Executable doesn't exist' deyilsa:  python -m playwright install chromium")
    sys.exit(1)


def joriy_sahifa(context):
    """Eng oxirgi ochilgan (faol) varaqni qaytaradi."""
    sahifalar = [p for p in context.pages if not p.is_closed()]
    return sahifalar[-1] if sahifalar else context.new_page()


class Bajaruvchi:
    """config.yaml dagi qadamlarni bitta Excel qatori uchun bajaradi."""

    def __init__(self, context, sozlama: dict, tekshir: bool):
        self.context = context
        self.kutish_ms = int(float(sozlama.get("kutish_soniya", 15)) * 1000)
        self.qadam_pauza = float(sozlama.get("qadamlar_orasi", 0.3))
        self.tekshir = tekshir

    @property
    def page(self):
        return joriy_sahifa(self.context)

    def element(self, selector: str):
        return self.page.locator(selector).first

    def qiymat_olish(self, qadam: dict, qator: dict) -> str:
        if "ustun" in qadam:
            nom = qadam["ustun"]
            if nom not in qator:
                raise KeyError(f"Excel'da '{nom}' ustuni yo'q")
            qiymat = qiymat_matnga(qator[nom], qadam.get("format"))
        else:
            qiymat = shablon_toldirish(str(qadam.get("qiymat", "")), qator)
        almashtirish = qadam.get("almashtirish")
        if almashtirish:
            qiymat = str(almashtirish.get(qiymat, qiymat))
        return qiymat

    def bajar(self, qadamlar: list[dict], qator: dict) -> None:
        for i, qadam in enumerate(qadamlar, 1):
            amal = qadam.get("amal")
            nomi = qadam.get("izoh") or qadam.get("ustun") or qadam.get("selector") or amal
            try:
                self.bitta_qadam(qadam, qator)
            except Exception as e:
                if qadam.get("majburiy", True):
                    raise RuntimeError(f"{i}-qadam ({amal}: {nomi}) — {e}") from e
                chiqar(f"    - {i}-qadam o'tkazib yuborildi ({nomi}): {str(e).splitlines()[0]}")
            if self.qadam_pauza:
                time.sleep(self.qadam_pauza)

    def sahifalab_topish(self, sel: str, keyingi: str, kut: int) -> None:
        """Element joriy sahifada ko'rinmasa, 'keyingi sahifa' tugmasini bosib qidiradi."""
        for _ in range(200):
            try:
                self.element(sel).wait_for(state="visible", timeout=min(kut, 3000))
                return
            except PWTimeout:
                pass
            tugma = self.page.locator(keyingi).first
            if not tugma.is_visible() or not tugma.is_enabled() or "disabled" in (tugma.get_attribute("class") or ""):
                raise RuntimeError("barcha sahifalar ko'rildi, element topilmadi")
            tugma.click(timeout=kut)
            time.sleep(1.5)

    def variant_bosish(self, variant_sel: str | None, qiymat: str, kut: int) -> None:
        """Ochilgan ro'yxatdan matni qiymatga teng variantni bosadi (avval aniq, keyin qisman moslik).
        Matn selektorga qo'yilmaydi, shuning uchun Navro'z kabi apostrofli qiymatlar ham ishlaydi."""
        nomzodlar = self.page.locator(variant_sel or "li, [role='option'], .option, .dropdown-item")
        aniq = re.compile(r"^\s*" + re.escape(qiymat) + r"\s*$", re.IGNORECASE)
        oxiri = time.time() + kut / 1000
        while True:
            for moslik in (aniq, qiymat):
                v = nomzodlar.filter(has_text=moslik).filter(visible=True)
                if v.count():
                    v.first.click(timeout=kut)
                    return
            if time.time() > oxiri:
                raise RuntimeError(f"ro'yxatda '{qiymat}' varianti topilmadi")
            time.sleep(0.25)

    def bitta_qadam(self, q: dict, qator: dict) -> None:
        amal = q.get("amal")
        kut = int(float(q.get("kutish_soniya", self.kutish_ms / 1000)) * 1000)
        sel = shablon_toldirish(q["selector"], qator) if "selector" in q else None

        # Qiymat bo'sh bo'lsa va "bosh_bolsa_otkaz" yoqilgan bo'lsa — qadam bajarilmaydi
        if amal in ("yozish", "tanlash", "royxat", "fayl") and q.get("bosh_bolsa_otkaz", True):
            if not self.qiymat_olish(q, qator):
                return

        if q.get("saqlash") and self.tekshir:
            javob = input("    Saqlash tugmasini bosaymi? [Enter = ha, n = yo'q]: ").strip().lower()
            if javob in ("n", "no", "yoq", "yo'q"):
                raise RuntimeError("foydalanuvchi saqlashni bekor qildi")

        if amal == "ochish":
            url = shablon_toldirish(q["url"], qator)
            self.page.goto(url, timeout=max(kut, 30000), wait_until="domcontentloaded")

        elif amal == "bosish":
            if q.get("keyingi_sahifa"):
                self.sahifalab_topish(sel, q["keyingi_sahifa"], kut)
            el = self.element(sel)
            el.wait_for(state="visible", timeout=kut)
            el.scroll_into_view_if_needed(timeout=kut)
            el.click(timeout=kut, force=bool(q.get("majburan")))

        elif amal == "yozish":
            qiymat = self.qiymat_olish(q, qator)
            el = self.element(sel)
            el.wait_for(state="visible", timeout=kut)
            el.click(timeout=kut)
            if q.get("sekin", False):
                # Maskali (masalan sana, telefon, JSHSHIR) maydonlar uchun harfma-harf yozish
                el.press("Control+A")
                el.press("Delete")
                el.press_sequentially(qiymat, delay=int(q.get("harf_orasi_ms", 60)))
            else:
                el.fill(qiymat, timeout=kut)
            if q.get("enter"):
                el.press("Enter")
            if q.get("tab"):
                el.press("Tab")

        elif amal == "tanlash":
            # Oddiy HTML <select>
            qiymat = self.qiymat_olish(q, qator)
            el = self.element(sel)
            el.wait_for(state="attached", timeout=kut)
            try:
                el.select_option(label=qiymat, timeout=kut)
            except PWError:
                el.select_option(value=qiymat, timeout=kut)

        elif amal == "royxat":
            # Maxsus (div/ul asosidagi) ochiluvchi ro'yxat: bosib ochish, kerak bo'lsa qidirish, variantni bosish
            qiymat = self.qiymat_olish(q, qator)
            el = self.element(sel)
            el.wait_for(state="visible", timeout=kut)
            el.click(timeout=kut)
            if q.get("qidiruv_selector"):
                qidiruv = self.element(q["qidiruv_selector"])
                qidiruv.fill(qiymat, timeout=kut)
            elif q.get("yozib_qidirish"):
                self.page.keyboard.type(qiymat, delay=40)
            time.sleep(float(q.get("variant_kutish", 0.5)))
            self.variant_bosish(q.get("variant_selector"), qiymat, kut)

        elif amal == "belgilash":
            # Checkbox / radio. 'ustun' berilsa, qiymat "ha/1/true/x" bo'lsa belgilanadi
            el = self.element(sel)
            el.wait_for(state="attached", timeout=kut)
            if "ustun" in q or "qiymat" in q:
                kerak = self.qiymat_olish(q, qator).lower() in ("1", "ha", "true", "x", "+", "yes", "да", "bor")
            else:
                kerak = True
            el.set_checked(kerak, timeout=kut, force=True)

        elif amal == "fayl":
            yol = Path(self.qiymat_olish(q, qator))
            if not yol.is_absolute():
                yol = BAZA / yol
            self.element(sel).set_input_files(str(yol), timeout=kut)

        elif amal == "tugma":
            self.page.keyboard.press(q["tugma"])

        elif amal == "kutish":
            if sel:
                holat = q.get("holat", "visible")
                self.element(sel).wait_for(state=holat, timeout=kut)
            elif q.get("url"):
                self.page.wait_for_url(re.compile(q["url"]), timeout=kut)
            else:
                self.page.wait_for_load_state(q.get("holat", "networkidle"), timeout=kut)

        elif amal == "pauza":
            time.sleep(float(q.get("soniya", 1)))

        elif amal == "tekshirish":
            # Sahifada xato xabari chiqqan-chiqmaganini tekshirish
            if sel and self.page.locator(sel).first.is_visible():
                matn = self.page.locator(sel).first.inner_text().strip()
                raise RuntimeError(f"saytdagi xabar: {matn[:200]}")

        elif amal == "qolda":
            # Masalan, captcha yoki ERI tasdig'i uchun
            input(f"    {q.get('xabar', 'Qolda bajaring va Enter bosing')} ")

        else:
            raise ValueError(f"noma'lum amal: {amal!r}")


# --------------------------------------------------------------------------- #
#  Selektor aniqlash rejimi
# --------------------------------------------------------------------------- #

TANLASH_JS = r"""
(() => {
  if (window.__selTanlash) return;
  window.__selTanlash = true;
  const esc = s => (window.CSS && CSS.escape) ? CSS.escape(s) : s.replace(/([^\w-])/g, '\\$1');
  const unik = (sel) => { try { return document.querySelectorAll(sel).length === 1; } catch (e) { return false; } };
  function yol(el) {
    const tag = el.tagName.toLowerCase();
    if (el.id && !/\d{3,}/.test(el.id) && unik('#' + esc(el.id))) return '#' + esc(el.id);
    for (const a of ['name', 'data-testid', 'formcontrolname', 'placeholder', 'aria-label', 'type']) {
      const v = el.getAttribute(a);
      if (v) { const s = `${tag}[${a}="${v.replace(/"/g, '\\"')}"]`; if (unik(s)) return s; }
    }
    const qism = [];
    let e = el;
    while (e && e.nodeType === 1 && e !== document.body) {
      let s = e.tagName.toLowerCase();
      if (e.id && !/\d{3,}/.test(e.id)) { qism.unshift('#' + esc(e.id)); break; }
      const ota = e.parentElement;
      if (ota) {
        const bir = [...ota.children].filter(c => c.tagName === e.tagName);
        if (bir.length > 1) s += `:nth-of-type(${bir.indexOf(e) + 1})`;
      }
      qism.unshift(s);
      e = ota;
    }
    return qism.join(' > ');
  }
  document.addEventListener('click', (ev) => {
    if (!ev.altKey) return;
    ev.preventDefault(); ev.stopPropagation();
    const el = ev.target;
    const matn = (el.innerText || el.value || '').trim().replace(/\s+/g, ' ').slice(0, 60);
    const tag = el.tagName.toLowerCase();
    const old = el.style.outline;
    el.style.outline = '3px solid red';
    setTimeout(() => el.style.outline = old, 1200);
    window.__selTanlandi({
      css: yol(el), tag, matn,
      turi: el.getAttribute('type') || '',
      url: location.href,
    });
  }, true);
})();
"""


def yaml_matn(s: str) -> str:
    """Matnni config.yaml ga to'g'ridan-to'g'ri qo'yish mumkin bo'lgan ko'rinishda qaytaradi."""
    return "'" + s.replace("'", "''") + "'"


def tanlash_rejimi(context, boshlash_url: str) -> None:
    def tanlandi(_source, info):
        tavsiya = {"input": "yozish", "textarea": "yozish", "select": "tanlash"}.get(info["tag"], "bosish")
        if info["turi"] in ("checkbox", "radio"):
            tavsiya = "belgilash"
        chiqar("-" * 60)
        chiqar(f"Sahifa : {info['url']}")
        chiqar(f"Element: <{info['tag']}> {info['matn']!r}")
        chiqar(f"  - amal: {tavsiya}")
        chiqar(f"    selector: {yaml_matn(info['css'])}")
        if info["matn"] and info["tag"] in ("button", "a", "span", "li", "div"):
            chiqar(f"    # yoki matn bo'yicha:  selector: {yaml_matn('text=' + info['matn'])}")

    context.expose_binding("__selTanlandi", tanlandi)
    context.add_init_script(TANLASH_JS)
    page = joriy_sahifa(context)
    if boshlash_url:
        page.goto(boshlash_url, wait_until="domcontentloaded")
    for p in context.pages:
        try:
            p.evaluate(TANLASH_JS)
        except PWError:
            pass

    chiqar("=" * 60)
    chiqar("SELEKTOR ANIQLASH REJIMI")
    chiqar("  1) Saytga kiring, kerakli forma sahifasini oching.")
    chiqar("  2) Klaviaturada ALT tugmasini bosib turib, kerakli maydonga sichqoncha bilan bosing.")
    chiqar("  3) Bu yerda config.yaml ga qo'yish uchun tayyor qator chiqadi.")
    chiqar("  Tugatish uchun brauzerni yoping yoki bu oynada Ctrl+C bosing.")
    chiqar("=" * 60)
    try:
        while context.pages:
            context.pages[0].wait_for_timeout(300)
    except (KeyboardInterrupt, PWError):
        pass


# --------------------------------------------------------------------------- #
#  Xonadonlar bo'yicha ishlash (kadastr raqami -> xonadon -> 7-8 ta odam)
# --------------------------------------------------------------------------- #

KADASTR_RE = re.compile(r"\b\d{2}:\d{2}:\d{2}:\d{2}:\d{2}:\d{3,5}(?::\d{3,5}){0,2}\b")


def kadastr_norm(k) -> str:
    return re.sub(r"\s+", "", qiymat_matnga(k))


def xonadonlarga_ajratish(qatorlar, guruh: dict) -> dict[str, list]:
    """Excel qatorlarini guruhlash ustuni (kadastr) bo'yicha xonadonlarga ajratadi, tartib saqlanadi."""
    ustun = guruh.get("ustun", "TV_ALP_CADASTRE")
    xonadonlar: dict[str, list] = {}
    for r, q in qatorlar:
        k = kadastr_norm(q.get(ustun))
        if k:
            xonadonlar.setdefault(k, []).append((r, q))
    return xonadonlar


def havola_bilan_ochish(havola: str, xonadon_qadamlar: list[dict]) -> list[dict]:
    """Xonadon havolasi ma'lum bo'lsa: ro'yxatda qidirib bosish o'rniga to'g'ridan-to'g'ri ochamiz,
    kadastr bosilgandan keyingi qadamlar (agar bo'lsa) saqlanadi."""
    keyingi = []
    for i, q in enumerate(xonadon_qadamlar):
        if q.get("amal") == "bosish" and "{" in str(q.get("selector", "")):
            keyingi = xonadon_qadamlar[i + 1:]
    return [{"amal": "ochish", "url": havola}] + keyingi


HAVOLA_JS = r"""(re) => {
  const rx = new RegExp(re, 'g'); const natija = {};
  for (const el of document.querySelectorAll('a, td, span, div, button')) {
    if (el.children.length > 2) continue;
    const m = (el.innerText || '').match(rx);
    if (!m || m.length !== 1) continue;
    const a = el.closest('a[href]') || el.querySelector('a[href]');
    const href = a && !a.getAttribute('href').startsWith('#') && !a.href.startsWith('javascript') ? a.href : null;
    if (!(m[0] in natija) || href) natija[m[0]] = href;
  }
  return natija;
}"""


def saytdagi_kadastrlar(bajaruvchi: "Bajaruvchi", guruh: dict) -> dict[str, str | None]:
    """Xonadonlar ro'yxati sahifasidagi barcha kadastr raqamlarini (va havolasini) yig'adi, sahifalab."""
    page = bajaruvchi.page
    page.goto(guruh["royxat_url"], wait_until="domcontentloaded")
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PWTimeout:
        pass
    kutish = guruh.get("royxat_kutish_selector")
    if kutish:
        page.locator(kutish).first.wait_for(timeout=bajaruvchi.kutish_ms)
    topilgan: dict[str, str | None] = {}
    keyingi = guruh.get("keyingi_sahifa_selector")
    for _ in range(int(guruh.get("max_sahifa", 200))):
        time.sleep(float(guruh.get("sahifa_kutish", 1.5)))
        oldin = len(topilgan)
        for k in KADASTR_RE.findall(page.locator("body").inner_text()):
            topilgan.setdefault(k, None)
        try:
            for k, href in page.evaluate(HAVOLA_JS, KADASTR_RE.pattern).items():
                if href or k not in topilgan:
                    topilgan[k] = href
        except PWError:
            pass
        if not keyingi:
            break
        tugma = page.locator(keyingi).first
        try:
            if not tugma.is_visible() or not tugma.is_enabled() or "disabled" in (tugma.get_attribute("class") or ""):
                break
            if len(topilgan) == oldin and oldin:
                break
            tugma.click(timeout=bajaruvchi.kutish_ms)
        except PWError:
            break
    return topilgan


# --------------------------------------------------------------------------- #
#  Yozib olish rejimi: siz bitta fuqaroni qo'lda kiritasiz — dastur qadamlarni yozadi
# --------------------------------------------------------------------------- #

YOZISH_JS = TANLASH_JS.split("document.addEventListener('click'")[0].replace("__selTanlash", "__selYozish") + r"""
  const yubor = (d) => { try { window.__qadamYozildi(Object.assign({url: location.href}, d)); } catch (e) {} };
  const matnOl = el => (el.innerText || el.value || el.getAttribute('aria-label') || '').trim().replace(/\s+/g, ' ').slice(0, 80);
  document.addEventListener('change', (ev) => {
    const el = ev.target; const tag = el.tagName.toLowerCase();
    if (tag === 'select') {
      const opt = el.options[el.selectedIndex];
      yubor({tur: 'tanlash', css: yol(el), qiymat: opt ? opt.text.trim() : el.value});
    } else if (el.type === 'checkbox' || el.type === 'radio') {
      yubor({tur: 'belgilash', css: yol(el), qiymat: el.checked ? 'ha' : 'yoq'});
    } else if (el.type === 'file') {
      yubor({tur: 'fayl', css: yol(el), qiymat: ''});
    } else if (tag === 'input' || tag === 'textarea') {
      yubor({tur: 'yozish', css: yol(el), qiymat: el.value});
    }
  }, true);
  document.addEventListener('click', (ev) => {
    let el = ev.target;
    const tag = el.tagName.toLowerCase();
    if (['input', 'textarea', 'select', 'option'].includes(tag) && !['submit', 'button'].includes(el.type)) return;
    const tugma = el.closest('button, a, [role="button"], [role="option"], li');
    if (tugma) el = tugma;
    yubor({tur: 'bosish', css: yol(el), tag: el.tagName.toLowerCase(), matn: matnOl(el)});
  }, true);
})();
"""


def _norm(v) -> str:
    return re.sub(r"\s+", " ", qiymat_matnga(v)).strip().upper()


def ustun_top(qiymat: str, qator: dict) -> str | None:
    """Kiritilgan qiymat namuna qatordagi qaysi ustunga teng ekanini topadi."""
    q = _norm(qiymat)
    if not q:
        return None
    for nom, v in qator.items():
        if _norm(v) == q:
            return nom
    return None


def hodisalardan_qadamlar(hodisalar: list[dict], qator: dict, boshlash_url: str) -> list[dict]:
    qadamlar: list[dict] = [{"amal": "ochish", "url": boshlash_url}] if boshlash_url else []
    for h in hodisalar:
        tur = h["tur"]
        if tur == "yozish":
            # Bir maydonga bir necha marta yozilgan bo'lsa, oxirgisi qoladi
            if qadamlar and qadamlar[-1].get("amal") == "yozish" and qadamlar[-1]["selector"] == h["css"]:
                qadamlar.pop()
            q = {"amal": "yozish", "selector": h["css"]}
            ustun = ustun_top(h["qiymat"], qator)
            if ustun:
                q["ustun"] = ustun
            else:
                q["qiymat"] = h["qiymat"]
            qadamlar.append(q)
        elif tur in ("tanlash", "belgilash"):
            q = {"amal": tur, "selector": h["css"]}
            ustun = ustun_top(h["qiymat"], qator) if tur == "tanlash" else None
            if ustun:
                q["ustun"] = ustun
            elif tur == "tanlash":
                q["qiymat"] = h["qiymat"]
            qadamlar.append(q)
        elif tur == "fayl":
            qadamlar.append({"amal": "fayl", "selector": h["css"], "qiymat": "FAYL_YOLINI_YOZING"})
        elif tur == "bosish":
            ustun = ustun_top(h.get("matn", ""), qator)
            if ustun:
                # Excel qiymatiga teng variant bosildi — bu ochiluvchi ro'yxat.
                # Oldin: [ro'yxatni ochish bosildi] + [ixtiyoriy: qidiruv maydoniga yozildi]
                qidiruv = qadamlar.pop() if qadamlar and qadamlar[-1]["amal"] == "yozish" else None
                q = {"amal": "royxat", "ustun": ustun, "variant_selector": h.get("tag") or "li"}
                if qadamlar and qadamlar[-1]["amal"] == "bosish" and not qadamlar[-1].get("saqlash"):
                    q["selector"] = qadamlar.pop()["selector"]
                    if qidiruv:
                        q["qidiruv_selector"] = qidiruv["selector"]
                    qadamlar.append(q)
                    continue
                if qidiruv:  # avtoto'ldiriladigan maydon: yozib, variantni tanlash
                    q["selector"] = qidiruv["selector"]
                    q["yozib_qidirish"] = True
                    qadamlar.append(q)
                    continue
            q = {"amal": "bosish", "selector": h["css"]}
            matn = h.get("matn", "")
            if ustun:
                # Masalan kadastr raqami bosildi — har bir qator uchun o'z qiymati bo'yicha bosiladi
                q["selector"] = f'text="{{{ustun}}}"'
                q["izoh"] = ustun
                qadamlar.append(q)
                continue
            if matn and len(matn) <= 40 and h.get("tag") in ("button", "a") and "\n" not in matn:
                q["selector"] = f"text={matn}"
            q["izoh"] = matn[:40] if matn else h.get("tag", "")
            if re.search(r"saqla|сохран|save|yubor|отправ|tasdiq", matn, re.IGNORECASE):
                q["saqlash"] = True
            qadamlar.append(q)
    return qadamlar


def _enter_kutish(context) -> None:
    """Foydalanuvchi Enter bosguncha brauzer hodisalarini qabul qilib turadi."""
    import threading
    tugadi = threading.Event()
    threading.Thread(target=lambda: (input(), tugadi.set()), daemon=True).start()
    while not tugadi.is_set() and context.pages:
        try:
            context.pages[0].wait_for_timeout(200)
        except PWError:
            break
    try:
        context.pages[0].wait_for_timeout(500)  # oxirgi hodisalar yetib kelsin
    except (PWError, IndexError):
        pass


def yozib_olish_rejimi(context, cfg: dict, cfg_yol: Path, namuna_r: int, jadval: "ExcelJadval") -> None:
    qatorlar = dict(jadval.qatorlar())
    guruh = cfg.get("guruhlash") or {}
    royxat_url = guruh.get("royxat_url")

    joriy = {"qator": {}, "yozish": False}
    hodisalar: list[dict] = []

    def yozildi(_source, h):
        if not joriy["yozish"]:
            return
        hodisalar.append(h)
        korinish = h.get("qiymat") if h["tur"] != "bosish" else h.get("matn")
        ustun = ustun_top(korinish or "", joriy["qator"])
        chiqar(f"  + {h['tur']:9} {str(korinish)[:40]!r}" + (f"  -> Excel ustuni: {ustun}" if ustun else ""))

    context.expose_binding("__qadamYozildi", yozildi)
    context.add_init_script(YOZISH_JS)
    page = joriy_sahifa(context)
    bosh_url = royxat_url or cfg.get("sayt", {}).get("kirish_url")
    if bosh_url:
        page.goto(bosh_url, wait_until="domcontentloaded")

    chiqar("=" * 60)
    chiqar("YOZIB OLISH REJIMI")
    input("Brauzerda saytga kiring (login / ERI / OneID), so'ng shu oynada Enter bosing... ")

    if guruh and not namuna_r and royxat_url:
        # Ro'yxat sahifasidagi xonadonlardan Excel'da odami borlarini taklif qilamiz
        bajaruvchi = Bajaruvchi(context, cfg.get("brauzer", {}), False)
        saytdagi = saytdagi_kadastrlar(bajaruvchi, guruh)
        xonadonlar = xonadonlarga_ajratish(qatorlar.items(), guruh)
        mos = [(k, v) for k, v in xonadonlar.items() if kadastr_norm(k) in saytdagi]
        chiqar(f"Ro'yxatda {len(saytdagi)} ta kadastr raqami topildi, Excel'da odami borlari: {len(mos)} ta.")
        for k, v in mos[:10]:
            r0, q0 = v[0]
            ism = " ".join(qiymat_matnga(q0.get(u)) for u in jadval.korsatish)
            chiqar(f"    {k:32} {len(v)} ta odam, masalan {r0}-qator: {ism}")
        if not mos:
            chiqar("Mos xonadon topilmadi. config.yaml dagi guruhlash.royxat_url va Excel faylini tekshiring.")
            return
        javob = input(f"Namuna uchun Excel qator raqamini kiriting [Enter = {mos[0][1][0][0]}]: ").strip()
        namuna_r = int(javob) if javob else mos[0][1][0][0]
        joriy_sahifa(context).goto(royxat_url, wait_until="domcontentloaded")

    if namuna_r not in qatorlar:
        chiqar(f"{namuna_r}-qator topilmadi (filtrni tekshiring). Mavjud qatorlar: {list(qatorlar)[:10]} ...")
        return
    qator = joriy["qator"] = qatorlar[namuna_r]
    joriy["yozish"] = True
    chiqar(f"Namuna — Excel'ning {namuna_r}-qatori:")
    for nom, v in qator.items():
        if v not in (None, ""):
            chiqar(f"    {nom:25} {qiymat_matnga(v)}")
    chiqar("-" * 60)
    for p in context.pages:
        try:
            p.evaluate(YOZISH_JS)
        except PWError:
            pass

    yangi = dict(cfg)
    if guruh:
        kadastr = qiymat_matnga(qator.get(guruh.get("ustun", "TV_ALP_CADASTRE")))
        chiqar(f"1-BOSQICH. Ro'yxatda {kadastr} xonadonini toping va kadastr raqamini bosing")
        chiqar("           (qidiruv maydoni bo'lsa, unga yozib qidirishingiz mumkin).")
        chiqar("           Xonadon ochilgach, shu oynada Enter bosing.")
        bosh = len(hodisalar)
        _enter_kutish(context)
        xonadon_qadamlar = hodisalardan_qadamlar(hodisalar[bosh:], qator, royxat_url)
        for q in xonadon_qadamlar:
            if q.get("amal") == "bosish" and q.get("izoh") in qator:
                guruh = dict(guruh, ustun=q["izoh"])
        bosh = len(hodisalar)
        chiqar("2-BOSQICH. Endi shu xonadonga namunadagi odamni QO'LDA kiriting:")
    else:
        chiqar("Fuqaro qo'shiladigan sahifani oching.")
        bosh = len(hodisalar)
        xonadon_qadamlar = None
    chiqar("   maydonlarni to'ldiring, ro'yxatlardan tanlang, Saqlash'ni bosing.")
    chiqar("   Qiymatlarni yuqoridagi Excel qiymatlari bilan AYNAN bir xil yozing — dastur ularni ustunlarga bog'laydi.")
    chiqar("   Tugatgach, shu oynada Enter bosing.")
    boshlash_url = joriy_sahifa(context).url
    _enter_kutish(context)
    qadamlar = hodisalardan_qadamlar(hodisalar[bosh:], qator, "" if guruh else boshlash_url)

    if not qadamlar:
        chiqar("Hech qanday harakat yozilmadi.")
        return
    if xonadon_qadamlar is not None:
        yangi["guruhlash"] = dict(guruh, xonadon_qadamlar=xonadon_qadamlar)
    else:
        yangi.setdefault("sayt", {})["kirish_url"] = boshlash_url
    yangi["qadamlar"] = qadamlar
    chiqish = cfg_yol.with_name(cfg_yol.stem + "_yozilgan.yaml")
    with open(chiqish, "w", encoding="utf-8") as f:
        f.write("# --yozib_olish rejimida avtomatik yaratildi. Tekshirib chiqing, so'ng config.yaml nomi bilan saqlang\n")
        f.write("# yoki:  python mahalla_bot.py --config " + chiqish.name + "\n")
        yaml.safe_dump(yangi, f, allow_unicode=True, sort_keys=False, width=200)
    chiqar("=" * 60)
    hammasi = (xonadon_qadamlar or []) + qadamlar
    chiqar(f"{len(hammasi)} ta qadam yozildi: {chiqish}")
    bog = sum(1 for q in hammasi if "ustun" in q or "{" in str(q.get("selector", "")))
    chiqar(f"Shundan {bog} tasi Excel ustunlariga bog'landi. Doimiy (qiymat:) qadamlarni tekshirib chiqing.")
    chiqar("Sinash:  python mahalla_bot.py --config " + chiqish.name + " --tekshir")


# --------------------------------------------------------------------------- #
#  Asosiy jarayon
# --------------------------------------------------------------------------- #

def asosiy() -> None:
    ap = argparse.ArgumentParser(description="Excel ma'lumotlarini saytga avtomatik kiritish")
    ap.add_argument("--config", default="config.yaml", help="sozlama fayli (standart: config.yaml)")
    ap.add_argument("--tanlash", action="store_true", help="selektor aniqlash rejimi")
    ap.add_argument("--tekshir", action="store_true", help="saqlashdan oldin tasdiq so'rash")
    ap.add_argument("--qatorlar", help="faqat shu Excel qatorlari, masalan: 2-10,15")
    ap.add_argument("--hammasi", action="store_true", help="OK bo'lgan qatorlarni ham qayta kiritish")
    ap.add_argument("--yozib_olish", type=int, nargs="?", const=0, metavar="QATOR",
                    help="shu Excel qatoridagi fuqaroni qo'lda kiritasiz, dastur qadamlarni yozib oladi")
    args = ap.parse_args()

    cfg_yol = (BAZA / args.config).resolve()
    cfg = config_oqish(cfg_yol)
    sayt = cfg.get("sayt", {})
    brauzer_cfg = cfg.get("brauzer", {})

    # Excel brauzerdan OLDIN o'qiladi (katta fayl uzoq o'qiladi — brauzer bo'sh turib qolmasin)
    jadval = None if args.tanlash else ExcelJadval(cfg.get("excel", {}))
    qadamlar = cfg.get("qadamlar") or []
    if not args.tanlash and args.yozib_olish is None and not qadamlar:
        chiqar("config.yaml da 'qadamlar' bo'limi bo'sh.")
        return
    if jadval is not None:
        chiqar(f"Excel o'qildi. Brauzer ochilmoqda... (brauzer oynasini yopmang)")

    with sync_playwright() as pw:
        context = brauzer_ochish(pw, brauzer_cfg)

        if args.tanlash:
            tanlash_rejimi(context, sayt.get("kirish_url", ""))
            return
        if args.yozib_olish is not None:
            yozib_olish_rejimi(context, cfg, cfg_yol, args.yozib_olish, jadval)
            return

        page = joriy_sahifa(context)
        if sayt.get("kirish_url"):
            page.goto(sayt["kirish_url"], wait_until="domcontentloaded")
        chiqar("=" * 60)
        chiqar("Brauzerda saytga kiring (login / ERI / OneID).")
        input("Kirib bo'lgach, shu oynada Enter bosing... ")

        tanlangan = qatorlar_oraligi(args.qatorlar)
        qator_pauza = float(brauzer_cfg.get("qatorlar_orasi", 1))
        xato_skrin = BAZA / "xatolar"
        hisob = {"ok": 0, "xato": 0, "otkaz": 0}

        barcha = [(r, q) for r, q in jadval.qatorlar() if not tanlangan or r in tanlangan]
        filtr_izoh = " (filtr bo'yicha)" if jadval.filtr else ""
        chiqar(f"Excel'da {len(barcha)} ta qator topildi{filtr_izoh}.")
        bajaruvchi = Bajaruvchi(context, brauzer_cfg, args.tekshir)

        def skrinshot(nom: str) -> None:
            try:
                xato_skrin.mkdir(exist_ok=True)
                bajaruvchi.page.screenshot(path=str(xato_skrin / f"{nom}.png"))
            except Exception:
                pass

        def bitta_odam(r: int, qator: dict, sarlavha: str) -> bool:
            nom = " ".join(qiymat_matnga(qator.get(u)) for u in jadval.korsatish)
            chiqar(f"{sarlavha} {r}-qator {nom}".rstrip())
            try:
                bajaruvchi.bajar(qadamlar, qator)
                jadval.holat_yoz(r, qator, "OK")
                hisob["ok"] += 1
                chiqar("    ✔ kiritildi")
                return True
            except Exception as e:
                hisob["xato"] += 1
                xabar = str(e).splitlines()[0][:300]
                jadval.holat_yoz(r, qator, f"XATO: {xabar}")
                chiqar(f"    ✖ {xabar}")
                skrinshot(f"qator_{r}")
                if brauzer_cfg.get("xatoda_toxtash", False):
                    input("    Xato. Davom etish uchun Enter bosing (to'xtatish: Ctrl+C)... ")
                return False
            finally:
                time.sleep(qator_pauza)

        def bajarilmagan(qatorlar):
            natija = []
            for r, q in qatorlar:
                if jadval.holat(r).startswith("OK") and not args.hammasi:
                    hisob["otkaz"] += 1
                else:
                    natija.append((r, q))
            return natija

        try:
            guruh = cfg.get("guruhlash")
            if not guruh:
                for n, (r, qator) in enumerate(bajarilmagan(barcha), 1):
                    bitta_odam(r, qator, f"[{n}]")
            else:
                xonadonlar = xonadonlarga_ajratish(barcha, guruh)
                if guruh.get("royxatdan_olish", True) and guruh.get("royxat_url"):
                    saytdagi = saytdagi_kadastrlar(bajaruvchi, guruh)
                    chiqar(f"Sayt ro'yxatida {len(saytdagi)} ta xonadon (kadastr raqami) topildi.")
                    if saytdagi:
                        bor = {k: v for k, v in xonadonlar.items() if kadastr_norm(k) in saytdagi}
                        havolali = sum(1 for k in bor if saytdagi.get(k))
                        if havolali and guruh.get("havola_orqali_ochish", True):
                            chiqar(f"{havolali} ta xonadon havolasi topildi — ular to'g'ridan-to'g'ri ochiladi.")
                        chiqar(f"Ulardan Excel'da bor: {len(bor)} ta. Faqat shular kiritiladi.")
                        yoq = sorted(set(saytdagi) - {kadastr_norm(k) for k in xonadonlar})
                        if yoq:
                            chiqar(f"Excel'da odami topilmagan xonadonlar ({len(yoq)}): {', '.join(yoq[:15])}"
                                   + (" ..." if len(yoq) > 15 else ""))
                        xonadonlar = bor
                else:
                    saytdagi = {}
                eng_kop = int(guruh.get("max_odam") or 0)
                xonadon_qadamlar = [dict(q) for q in guruh.get("xonadon_qadamlar") or []]
                for q in xonadon_qadamlar:
                    if guruh.get("keyingi_sahifa_selector") and q.get("amal") == "bosish" and "{" in str(q.get("selector", "")):
                        q.setdefault("keyingi_sahifa", guruh["keyingi_sahifa_selector"])
                for xn, (kadastr, odamlar) in enumerate(xonadonlar.items(), 1):
                    qoldi = bajarilmagan(odamlar)
                    if eng_kop:
                        qoldi = qoldi[:eng_kop]
                    if not qoldi:
                        continue
                    chiqar("=" * 60)
                    chiqar(f"XONADON {xn}/{len(xonadonlar)}: {kadastr} — {len(qoldi)} ta odam kiritiladi")
                    ochiq = False
                    for i, (r, qator) in enumerate(qoldi, 1):
                        if not ochiq:
                            try:
                                havola = saytdagi.get(kadastr) if guruh.get("havola_orqali_ochish", True) else None
                                if havola:
                                    bajaruvchi.bajar(havola_bilan_ochish(havola, xonadon_qadamlar), qator)
                                else:
                                    bajaruvchi.bajar(xonadon_qadamlar, qator)
                                ochiq = True
                            except Exception as e:
                                xabar = f"xonadon ochilmadi: {str(e).splitlines()[0][:250]}"
                                chiqar(f"    ✖ {xabar}")
                                skrinshot(f"xonadon_{kadastr.replace(':', '-')}")
                                for r2, q2 in qoldi[i - 1:]:
                                    jadval.holat_yoz(r2, q2, f"XATO: {xabar}")
                                    hisob["xato"] += 1
                                break
                        # Xatodan keyin sahifa holati noma'lum — xonadonni qaytadan ochamiz
                        ochiq = bitta_odam(r, qator, f"  [{i}/{len(qoldi)}]") and not guruh.get("har_odamda_qayta_ochish", False)
        except KeyboardInterrupt:
            chiqar("\nFoydalanuvchi to'xtatdi.")
        finally:
            jadval.saqlash()
            chiqar("=" * 60)
            chiqar(f"Tayyor. Muvaffaqiyatli: {hisob['ok']}, xato: {hisob['xato']}, avval kiritilgan (o'tkazildi): {hisob['otkaz']}")
            chiqar(f"Natija: {jadval.natija_fayl}")
            if hisob["xato"]:
                chiqar(f"Xato qatorlar rasmlari: {xato_skrin}")
            if not brauzer_cfg.get("yashirin"):
                try:
                    input("Brauzerni yopish uchun Enter bosing... ")
                except (KeyboardInterrupt, EOFError):
                    pass
            context.close()


if __name__ == "__main__":
    try:
        asosiy()
    except KeyboardInterrupt:
        chiqar("\nTo'xtatildi.")
    except PWError as e:
        chiqar("=" * 60)
        if "closed" in str(e).lower():
            chiqar("Brauzer yopilib qoldi. Sabablari:")
            chiqar("  - dastur ishlayotganda brauzer oynasi yopildi;")
            chiqar("  - avvalgi ishga tushirishdan qolgan brauzer oynasi ochiq (uni yoping).")
            chiqar("Barcha shu dastur ochgan Chrome oynalarini yoping va qaytadan ishga tushiring.")
        else:
            chiqar(f"Brauzer xatosi: {str(e).splitlines()[0]}")
        sys.exit(1)
