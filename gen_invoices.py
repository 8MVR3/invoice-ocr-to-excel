"""Генерирует тестовые счета на оплату для демонстрации кворка «PDF и сканы → Excel».

Все компании, ИНН и суммы вымышленные. ИНН специально сделаны с неверной
контрольной суммой, чтобы случайно не совпасть с настоящей организацией.
"""
import io
import random
from datetime import date
from pathlib import Path

import numpy as np
import pypdfium2 as pdfium
from num2words import num2words
from PIL import Image, ImageFilter
from reportlab.lib.colors import Color, black
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

random.seed(7)
np.random.seed(7)

FONT_DIR = "/usr/share/fonts/truetype/dejavu/"
pdfmetrics.registerFont(TTFont("DV", FONT_DIR + "DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DVB", FONT_DIR + "DejaVuSans-Bold.ttf"))

OUT = Path("/home/claude/demo/входящие_счета")
OUT.mkdir(parents=True, exist_ok=True)


def bad_inn(length):
    """Случайный ИНН с намеренно неверной контрольной цифрой."""
    digits = [random.randint(0, 9) for _ in range(length)]
    digits[0] = random.randint(1, 9)
    if length == 10:
        w = [2, 4, 10, 3, 5, 9, 4, 6, 8]
        ok = sum(a * b for a, b in zip(w, digits[:9])) % 11 % 10
        digits[9] = (ok + 3) % 10
    else:
        w1 = [7, 2, 4, 10, 3, 5, 9, 4, 6, 8]
        ok = sum(a * b for a, b in zip(w1, digits[:10])) % 11 % 10
        digits[10] = (ok + 5) % 10
    return "".join(map(str, digits))


BUYER = ("ООО «Лайм Логистик»", bad_inn(10), "г. Москва, ул. Складская, д. 12")

SUPPLIERS = [
    ("ООО «Северная бумага»", bad_inn(10), "г. Москва, ул. Полевая, д. 5", [
        ("Бумага офисная А4, 500 листов", "пач", 389.00, (5, 20)),
        ("Ручка шариковая синяя", "шт", 24.50, (20, 100)),
        ("Папка-скоросшиватель пластиковая", "шт", 18.90, (10, 60)),
        ("Стикеры 76×76 мм, 100 листов", "шт", 64.00, (5, 30)),
        ("Степлер №24/6", "шт", 289.00, (1, 4)),
        ("Скобы для степлера №24/6, 1000 шт.", "уп", 39.00, (2, 10)),
    ]),
    ("ИП Ковалёва Марина Игоревна", bad_inn(12), "г. Москва, ул. Литейная, д. 9", [
        ("Печать визиток, 1000 шт.", "усл", 2400.00, (1, 2)),
        ("Печать листовок А5, 500 шт.", "усл", 3150.00, (1, 3)),
        ("Ламинирование А4", "шт", 45.00, (10, 40)),
    ]),
    ("ООО «Техснаб-Плюс»", bad_inn(10), "г. Химки, ул. Заводская, д. 3", [
        ("Картридж для лазерного принтера, чёрный", "шт", 2890.00, (1, 4)),
        ("Кабель USB 2.0, 1,8 м", "шт", 210.00, (2, 8)),
        ("Мышь компьютерная проводная", "шт", 590.00, (1, 6)),
        ("Клавиатура проводная", "шт", 1190.00, (1, 4)),
    ]),
    ("ООО «Чистый офис»", bad_inn(10), "г. Москва, ул. Садовая, д. 18", [
        ("Средство для мытья полов, 5 л", "шт", 640.00, (1, 4)),
        ("Бумажные полотенца, 2 рулона", "уп", 189.00, (4, 20)),
        ("Мешки для мусора 120 л, 10 шт.", "уп", 215.00, (2, 10)),
        ("Перчатки латексные", "пар", 58.00, (10, 50)),
    ]),
    ("ИП Орлов Денис Сергеевич", bad_inn(12), "г. Мытищи, ул. Речная, д. 7", [
        ("Вода питьевая, бутыль 19 л", "бут", 390.00, (4, 16)),
        ("Аренда кулера, месяц", "усл", 500.00, (1, 2)),
    ]),
    ("ООО «Мебельный двор»", bad_inn(10), "г. Москва, ул. Мебельная, д. 1", [
        ("Стул офисный на колёсиках", "шт", 6890.00, (1, 6)),
        ("Тумба приставная, 3 ящика", "шт", 8450.00, (1, 3)),
    ]),
]


def money(x):
    s = f"{x:,.2f}".replace(",", " ").replace(".", ",")
    return s


def rub_words(total):
    rub = int(total)
    kop = round((total - rub) * 100)
    words = num2words(rub, lang="ru")
    return f"{words[0].upper()}{words[1:]} рублей {kop:02d} копеек"


def draw_invoice(path, number, d, supplier, items, wrong_total_delta=0.0, stamp=True):
    name, inn, addr, _ = supplier
    c = canvas.Canvas(str(path), pagesize=A4)
    W, H = A4
    left = 50
    y = H - 60

    c.setFont("DVB", 15)
    c.drawString(left, y, f"Счёт на оплату № {number} от {d.strftime('%d.%m.%Y')}")
    y -= 14
    c.setLineWidth(1.2)
    c.line(left, y, W - left, y)
    y -= 26

    c.setFont("DVB", 10)
    c.drawString(left, y, "Поставщик:")
    c.setFont("DV", 10)
    c.drawString(left + 78, y, f"{name}, ИНН {inn}")
    y -= 14
    c.drawString(left + 78, y, addr)
    y -= 22
    c.setFont("DVB", 10)
    c.drawString(left, y, "Покупатель:")
    c.setFont("DV", 10)
    c.drawString(left + 78, y, f"{BUYER[0]}, ИНН {BUYER[1]}")
    y -= 14
    c.drawString(left + 78, y, BUYER[2])
    y -= 30

    cols = [left, left + 24, left + 260, left + 308, left + 344, left + 416, W - left]
    heads = ["№", "Товары (работы, услуги)", "Кол-во", "Ед.", "Цена", "Сумма"]
    row_h = 20
    c.setLineWidth(0.6)
    c.setFont("DVB", 9)
    c.rect(left, y - row_h, W - 2 * left, row_h)
    for i, h in enumerate(heads):
        c.drawString(cols[i] + 4, y - 14, h)
    for x in cols[1:-1]:
        c.line(x, y, x, y - row_h * (len(items) + 1))
    y -= row_h

    c.setFont("DV", 9)
    total = 0.0
    for i, (title, unit, price, qty) in enumerate(items, start=1):
        s = round(price * qty, 2)
        total += s
        c.rect(left, y - row_h, W - 2 * left, row_h)
        c.drawString(cols[0] + 4, y - 14, str(i))
        c.drawString(cols[1] + 4, y - 14, title)
        c.drawRightString(cols[3] - 6, y - 14, str(qty))
        c.drawString(cols[3] + 4, y - 14, unit)
        c.drawRightString(cols[5] - 6, y - 14, money(price))
        c.drawRightString(cols[6] - 6, y - 14, money(s))
        y -= row_h

    shown_total = round(total + wrong_total_delta, 2)
    y -= 18
    c.setFont("DVB", 10)
    c.drawRightString(W - left - 6, y, f"Итого: {money(shown_total)}")
    y -= 15
    c.setFont("DV", 10)
    c.drawRightString(W - left - 6, y, "Без налога (НДС)")
    y -= 15
    c.setFont("DVB", 10)
    c.drawRightString(W - left - 6, y, f"Всего к оплате: {money(shown_total)}")
    y -= 28
    c.setFont("DV", 9)
    c.drawString(left, y, f"Всего наименований {len(items)}, на сумму {money(shown_total)} руб.")
    y -= 14
    c.setFont("DVB", 9)
    c.drawString(left, y, rub_words(shown_total))
    y -= 50

    c.setFont("DV", 10)
    c.drawString(left, y, "Руководитель ______________")
    c.drawString(left + 280, y, "Бухгалтер ______________")

    if stamp:
        blue = Color(0.16, 0.30, 0.68, alpha=0.75)
        c.setStrokeColor(blue)
        c.setFillColor(blue)
        cx, cy = left + 175, y + 4
        c.setLineWidth(1.4)
        c.circle(cx, cy, 34)
        c.setLineWidth(0.8)
        c.circle(cx, cy, 27)
        short = name.replace("ООО ", "").replace("«", "").replace("»", "")
        if short.startswith("ИП "):
            parts = short.split()
            short = f"ИП {parts[1]} {parts[2][0]}.{parts[3][0]}."
        c.setFont("DVB", 6)
        c.drawCentredString(cx, cy + 2, short)
        c.setFont("DV", 5.5)
        c.drawCentredString(cx, cy - 7, f"ИНН {inn}")
        c.setFillColor(black)
        c.setStrokeColor(black)

    c.showPage()
    c.save()
    return total, shown_total


def make_scan(pdf_path, out_path, angle, as_pdf=False):
    pdf = pdfium.PdfDocument(str(pdf_path))
    img = pdf[0].render(scale=200 / 72).to_pil().convert("L")
    img = img.rotate(angle, resample=Image.BICUBIC, expand=True, fillcolor=255)
    arr = np.asarray(img).astype(np.float32)
    arr = arr * 0.92 + 14  # лёгкий серый фон, как у сканера
    arr += np.random.normal(0, 7, arr.shape)
    img = Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8)).filter(ImageFilter.GaussianBlur(0.5))
    if as_pdf:
        img.convert("RGB").save(out_path, "PDF", resolution=200)
    else:
        img.save(out_path, "JPEG", quality=72)


def pick_items(supplier):
    catalog = supplier[3]
    n = random.randint(2, min(4, len(catalog)))
    chosen = random.sample(catalog, n)
    return [(t, u, p, random.randint(*q)) for t, u, p, q in chosen]


plan = [
    # (поставщик, номер, дата, вид файла, имя файла, ошибка в итоге)
    (0, "147", date(2026, 9, 1), "pdf", "Счёт_147_Северная_бумага.pdf", 0),
    (2, "ТП-0912", date(2026, 9, 2), "pdf", "schet_TP-0912.pdf", 0),
    (1, "58", date(2026, 9, 3), "pdf", "Ковалёва_счёт_58.pdf", 0),
    (3, "1204", date(2026, 9, 4), "scan_jpg", "IMG_2041.jpg", 0),
    (4, "311", date(2026, 9, 7), "pdf", "Вода_сентябрь.pdf", 0),
    (5, "МД-77", date(2026, 9, 8), "scan_pdf", "скан_счёт_МД-77.pdf", 0),
    (0, "152", date(2026, 9, 10), "pdf", "Счёт_152_Северная_бумага.pdf", 0),
    (2, "ТП-0937", date(2026, 9, 14), "scan_jpg", "IMG_2087.jpg", 0),
    (3, "1239", date(2026, 9, 16), "pdf", "Чистый_офис_1239.pdf", 900.0),
    (1, "64", date(2026, 9, 21), "scan_pdf", "scan_0003.pdf", 0),
    (4, "329", date(2026, 9, 24), "pdf", "Вода_сентябрь_2.pdf", 0),
    (5, "МД-81", date(2026, 9, 28), "pdf", "Мебельный_двор_МД-81.pdf", 0),
]

tmp = Path("/home/claude/demo/_tmp")
tmp.mkdir(exist_ok=True)
truth = []
for k, (si, num, d, kind, fname, delta) in enumerate(plan):
    sup = SUPPLIERS[si]
    items = pick_items(sup)
    if kind == "pdf":
        total, shown = draw_invoice(OUT / fname, num, d, sup, items, delta, stamp=(k % 3 != 0))
    else:
        src = tmp / f"src_{k}.pdf"
        total, shown = draw_invoice(src, num, d, sup, items, delta, stamp=True)
        angle = random.choice([-1.4, -0.8, 0.7, 1.2])
        make_scan(src, OUT / fname, angle, as_pdf=(kind == "scan_pdf"))
    truth.append((fname, num, d.isoformat(), sup[0], sup[1], len(items), round(total, 2), shown))

for t in truth:
    print(t)
