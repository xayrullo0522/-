"""Namuna fuqarolar.xlsx faylini yaratadi (ustunlar tuzilishini ko'rish uchun)."""
import datetime as dt
from pathlib import Path

from openpyxl import Workbook

ustunlar = ["Familiya", "Ism", "Otasining ismi", "Tug'ilgan sana", "JSHSHIR",
            "Pasport", "Jinsi", "Telefon", "Ko'cha", "Uy raqami"]
qatorlar = [
    ["Aliyev", "Vali", "Karimovich", dt.date(1990, 5, 12), "31205901234567",
     "AB1234567", "E", "901234567", "Navro'z", "12"],
    ["Karimova", "Dilnoza", "Rustamovna", dt.date(1995, 11, 3), "41103951234567",
     "AC7654321", "A", "931112233", "Mustaqillik", "5"],
]

wb = Workbook()
ws = wb.active
ws.title = "Fuqarolar"
ws.append(ustunlar)
for q in qatorlar:
    ws.append(q)
for r in range(2, ws.max_row + 1):
    ws.cell(row=r, column=4).number_format = "DD.MM.YYYY"

yol = Path(__file__).resolve().parent / "fuqarolar.xlsx"
if yol.exists():
    print(f"{yol.name} allaqachon bor — ustiga yozilmadi.")
else:
    wb.save(yol)
    print(f"Yaratildi: {yol}")
