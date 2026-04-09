# -*- coding: utf-8 -*-
"""
Агент публікації результатів фестивалю TORONTO
===============================================
Збирає оцінки від членів журі → генерує PDF таблицю результатів → записує Laureate у Bitrix24

Використання:
    python agent_results.py --folder "КВІТЕНЬ 2026/Журі оцінки" --month "Квітень 2026"
    python agent_results.py --folder "..." --month "..." --bitrix-url "https://..." --write-bitrix

Залежності:
    pip install openpyxl reportlab requests
"""

import argparse
import os
import sys
import re
from collections import defaultdict

import openpyxl
import requests
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ---------------------------------------------------------------------------
# Шрифти (DejaVu — підтримує кирилицю, завантажується разом з reportlab)
# ---------------------------------------------------------------------------
def _register_fonts():
    """Реєструє Arial (Windows) або DejaVu — обидва підтримують кирилицю."""
    # Пробуємо Arial (є на всіх Windows)
    win_fonts = r"C:\Windows\Fonts"
    arial_r = os.path.join(win_fonts, "arial.ttf")
    arial_b = os.path.join(win_fonts, "arialbd.ttf")
    if os.path.exists(arial_r) and os.path.exists(arial_b):
        try:
            pdfmetrics.registerFont(TTFont("Arial",      arial_r))
            pdfmetrics.registerFont(TTFont("Arial-Bold", arial_b))
            return "Arial", "Arial-Bold"
        except Exception:
            pass
    # Пробуємо DejaVu — перевіряємо кілька відомих розташувань (Windows, Linux, поруч зі скриптом)
    dejavu_search = [
        os.path.dirname(__file__),
        win_fonts,
        "/usr/share/fonts/truetype/dejavu",          # Debian/Ubuntu/Streamlit Cloud
        "/usr/share/fonts/dejavu",
        "/usr/share/fonts/TTF",
    ]
    for folder in dejavu_search:
        r = os.path.join(folder, "DejaVuSans.ttf")
        b = os.path.join(folder, "DejaVuSans-Bold.ttf")
        if os.path.exists(r) and os.path.exists(b):
            try:
                pdfmetrics.registerFont(TTFont("DejaVu",      r))
                pdfmetrics.registerFont(TTFont("DejaVu-Bold", b))
                return "DejaVu", "DejaVu-Bold"
            except Exception:
                pass
    return "Helvetica", "Helvetica-Bold"


FONT_REGULAR, FONT_BOLD = _register_fonts()


# ---------------------------------------------------------------------------
# Конвертація значення Laureate
# ---------------------------------------------------------------------------
def convert_laureate(value) -> str:
    """
    Конвертує будь-який формат оцінки у фінальний рядок.
    Підтримує: 1/2/3, I/II/III, 1 місце, 1st/2nd/3d degree,
               Gran Pri, Гран Прі, нема, не відкрила, None → 1st degree
    """
    if value is None:
        return "1st degree"
    v = str(value).strip()
    if v == "":
        return "1st degree"
    low = v.lower()

    # Гран Прі — перевіряємо першим (з тире, без тире, різні варіанти)
    low_nodash = low.replace("-", " ").replace("'", " ")
    if any(k in low_nodash for k in ["gran pri", "гран прі", "гран при", "gran prize"]):
        return "Gran Pri"

    # Числові та рядкові варіанти 1/2/3
    # Витягуємо першу цифру або римські
    import re

    # Спочатку — точні відповідності
    exact = {
        "1": "1st degree", "1st": "1st degree", "1st degree": "1st degree",
        "1 місце": "1st degree", "1 место": "1st degree", "перше": "1st degree",
        "перший": "1st degree", "першe": "1st degree", "i": "1st degree",
        "2": "2nd degree", "2nd": "2nd degree", "2nd degree": "2nd degree",
        "2 місце": "2nd degree", "2 место": "2nd degree", "друге": "2nd degree",
        "другий": "2nd degree", "ii": "2nd degree",
        "3": "3d degree",  "3d": "3d degree",  "3d degree": "3d degree",
        "3rd": "3d degree", "3rd degree": "3d degree",
        "3 місце": "3d degree", "3 место": "3d degree", "третє": "3d degree",
        "третій": "3d degree", "iii": "3d degree",
    }
    if low in exact:
        return exact[low]

    # Числа з текстом: "1 ступінь", "2 ступень" тощо
    m = re.match(r'^(\d)', v)
    if m:
        n = int(m.group(1))
        if n == 1: return "1st degree"
        if n == 2: return "2nd degree"
        if n == 3: return "3d degree"

    # Римські цифри — шукаємо ІІІ/ІІ/І в будь-якому місці рядка
    # Лауреат ІІІ ступеня, Лауреат ІІ ступеня, Лауреат І ступеня
    if re.search(r'\biii\b', low) or re.search(r'ііі', low): return "3d degree"
    if re.search(r'\bii\b',  low) or re.search(r'іі',  low): return "2nd degree"
    if re.search(r'\bi\b',   low) or re.search(r'\bі\b', low): return "1st degree"

    # Дипломант → 3d degree
    if "дипломант" in low:
        return "3d degree"

    # Все інше (нема, не відкрила, error, тощо) → 1st degree
    return "1st degree"


# ---------------------------------------------------------------------------
# Читання файлів журі
# ---------------------------------------------------------------------------
SKIP_PIB_PREFIXES = ("сума", "гонорар", "разом", "total", "підсумок")

def _is_data_row(row_id, pib) -> bool:
    """Повертає True якщо рядок є учасником, а не технічним."""
    if pib is None:
        return False
    pib_s = str(pib).strip().lower()
    for prefix in SKIP_PIB_PREFIXES:
        if pib_s.startswith(prefix):
            return False
    return True


def read_jury_file(path: str) -> tuple[list[dict], list[str]]:
    """
    Читає один xlsx файл журі.
    Повертає (список рядків, лог-повідомлення).
    """
    fname = os.path.basename(path)
    log = []
    log.append(f"📂 Файл: {fname}")

    wb = openpyxl.load_workbook(path)
    results = []

    for sheet_name in wb.sheetnames:
        try:
            ws = wb[sheet_name]
        except Exception as e:
            log.append(f"  ⚠️ Аркуш '{sheet_name}': не вдалося відкрити ({e})")
            continue
        if not hasattr(ws, "iter_rows"):
            log.append(f"  ⏭️ Аркуш '{sheet_name}': діаграма, пропущено")
            continue

        headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
        log.append(f"  📋 Аркуш '{sheet_name}' | рядків: {ws.max_row - 1}")
        log.append(f"     Колонки: {headers}")

        def col(name):
            try:
                return headers.index(name)
            except ValueError:
                return None

        idx_id   = col("ID")
        idx_pib  = col("ПІБ Учасника")
        idx_nom  = col("Номінація")
        idx_vik  = col("Вікова категорія")
        idx_nazv = col("Назва або опис роботи") or col("Назва роботи")
        # Laureate або Ступінь (файли ДПМ)
        idx_lau  = col("Laureate") or col("Ступінь") or col("Оцінка") or col("Місце")
        idx_kom  = col("Коментар Журі") or col("Коментар журі")

        # Лог знайдених/відсутніх колонок
        col_status = []
        for cname, cidx in [("ID", idx_id), ("ПІБ Учасника", idx_pib),
                             ("Номінація", idx_nom), ("Назва або опис роботи", idx_nazv),
                             ("Laureate", idx_lau)]:
            if cidx is None:
                col_status.append(f"❌ '{cname}' — НЕ ЗНАЙДЕНО")
            else:
                col_status.append(f"✅ '{cname}' → [{cidx}]")
        for s in col_status:
            log.append(f"     {s}")

        if idx_pib is None or idx_lau is None:
            log.append(f"  🚫 Аркуш пропущено: немає обов'язкових колонок 'ПІБ Учасника' або 'Laureate'")
            continue

        no_id_count = 0
        no_score_count = 0
        no_nazva_count = 0
        sheet_rows = 0

        for row in ws.iter_rows(min_row=2, values_only=True):
            pib = row[idx_pib] if idx_pib is not None else None
            rid = row[idx_id]  if idx_id  is not None else None

            if not _is_data_row(rid, pib):
                continue

            sheet_rows += 1
            lau_raw = row[idx_lau] if idx_lau is not None else None
            nazva   = row[idx_nazv] if idx_nazv is not None else None

            if rid is None:
                no_id_count += 1
            if lau_raw is None:
                no_score_count += 1
            if not nazva:
                no_nazva_count += 1

            results.append({
                "id":       rid,
                "pib":      str(pib).strip() if pib else "",
                "nom":      str(row[idx_nom]).strip()  if idx_nom  is not None and row[idx_nom]  else "",
                "vik":      str(row[idx_vik]).strip()  if idx_vik  is not None and row[idx_vik]  else "",
                "nazva":    str(nazva).strip() if nazva else "",
                "laureate": convert_laureate(lau_raw),
                "raw_laureate": str(lau_raw) if lau_raw is not None else "None",
                "comment":  str(row[idx_kom]).strip() if idx_kom is not None and row[idx_kom] else "",
                "source":   fname,
            })

        log.append(f"  ✅ Зчитано: {sheet_rows} учасників")
        if no_id_count:
            log.append(f"  ⚠️ Без ID: {no_id_count} — запис у Bitrix24 неможливий")
        if no_score_count:
            log.append(f"  ⚠️ Без оцінки (Laureate=None): {no_score_count} → буде '1st degree'")
        if no_nazva_count:
            log.append(f"  ℹ️ Без назви роботи: {no_nazva_count} — колонка порожня у джерелі")

    log.append(f"  📊 Разом з файлу: {len(results)} рядків")
    return results, log


def read_all_jury_with_log(folder: str) -> tuple[list[dict], list[str]]:
    """Читає всі xlsx файли у папці. Повертає (рядки, повний лог)."""
    all_rows = []
    full_log = []
    for fname in sorted(os.listdir(folder)):
        if fname.lower().endswith(".xlsx") and not fname.startswith("~$"):
            path = os.path.join(folder, fname)
            try:
                rows, log = read_jury_file(path)
                all_rows.extend(rows)
                full_log.extend(log)
                full_log.append("")
                print(f"  {fname}: {len(rows)} рядків")
            except Exception as e:
                msg = f"  ❌ ПОМИЛКА {fname}: {e}"
                full_log.append(msg)
                print(msg)
    return all_rows, full_log


def read_all_jury(folder: str) -> list[dict]:
    """Читає всі xlsx файли у папці (без логу — для CLI)."""
    all_rows = []
    for fname in sorted(os.listdir(folder)):
        if fname.lower().endswith(".xlsx") and not fname.startswith("~$"):
            path = os.path.join(folder, fname)
            try:
                rows, log = read_jury_file(path)
                print(f"  {fname}: {len(rows)} рядків")
                for line in log:
                    print(f"    {line}")
                all_rows.extend(rows)
            except Exception as e:
                print(f"  ПОМИЛКА {fname}: {e}")
    return all_rows


# ---------------------------------------------------------------------------
# Генерація PDF — дизайн як у лютому 2026
# ---------------------------------------------------------------------------

YELLOW       = colors.HexColor("#FFD700")   # жовтий фон рядків / заголовок
BLUE_HEADER  = colors.HexColor("#1a237e")   # темно-синій заголовок таблиці
WHITE        = colors.white

# Кольори тільки клітинки Laureate
LAUREATE_CELL_COLORS = {
    "Gran Pri":   colors.HexColor("#FF6F00"),  # насичений помаранчевий — видно на жовтому
    "1st degree": colors.HexColor("#B3E5FC"),  # блакитний
    "2nd degree": YELLOW,                      # той самий жовтий (без виділення)
    "3d degree":  colors.HexColor("#CE93D8"),  # ліловий
}

HEADER_TEXT = (
    "Вітаємо всіх учасників фестивалю з чудовими результатами!\n"
    "Як користуватися таблицею?\n"
    "Знаходимо у стовпчику Artist назву колективу або ПІБ учасника.\n"
    "Сортування за алфавітом. Навпроти\n"
    "ПІБ учасника ви бачите диплом лауреата від 3 до 1-ї або Гран Прі."
)


def build_pdf(rows: list[dict], output_path: str, month: str, publish_date: str = ""):
    """Генерує PDF таблицю результатів (дизайн лютий 2026)."""
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=1.5*cm,  bottomMargin=1.5*cm,
    )

    # --- Стилі тексту ---
    cell_style = ParagraphStyle(
        "cell", fontName=FONT_REGULAR, fontSize=8, leading=10, wordWrap="CJK",
    )
    bold_style = ParagraphStyle(
        "bold", fontName=FONT_BOLD, fontSize=8, leading=10,
    )
    pib_style = ParagraphStyle(                          # ПІБ: жирний
        "pib", fontName=FONT_BOLD, fontSize=8, leading=10, wordWrap="CJK",
    )
    header_instr_style = ParagraphStyle(                 # текст інструкції у жовтому блоці
        "hdr_instr", fontName=FONT_BOLD, fontSize=9, leading=13,
        alignment=1,                                     # по центру
        textColor=colors.black,
    )

    story = []

    # --- Жовтий блок-заголовок (як у лютому) ---
    instr = HEADER_TEXT
    if publish_date:
        instr += f"\nПублікація онлайн дипломів запланована на {publish_date} на сайті\nhttps://toronto.org.ua/"
    else:
        instr += f"\nhttps://toronto.org.ua/"

    instr_para = Paragraph(instr.replace("\n", "<br/>"), header_instr_style)
    instr_table = Table([[instr_para]], colWidths=[18.4*cm])
    instr_table.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), YELLOW),
        ("BOX",           (0, 0), (-1, -1), 1.0, colors.HexColor("#CCCCCC")),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(instr_table)
    story.append(Spacer(1, 0.3*cm))

    # --- Сортуємо за алфавітом ---
    sorted_rows = sorted(rows, key=lambda r: r["pib"].lower())

    # --- Будуємо таблицю ---
    col_widths = [1.3*cm, 5.5*cm, 3.8*cm, 5.0*cm, 2.8*cm]

    # Заголовок таблиці: жовтий фон, жирний текст
    hdr_cell = ParagraphStyle("hdr_cell", fontName=FONT_BOLD, fontSize=8,
                               leading=10, textColor=colors.black, alignment=1)
    table_data = [[
        Paragraph("ID",                       hdr_cell),
        Paragraph("ПІБ Учасника",             hdr_cell),
        Paragraph("Номінація",                hdr_cell),
        Paragraph("Назва або опис\nроботи",   hdr_cell),
        Paragraph("Laureate",                 hdr_cell),
    ]]

    lau_col_idx = 4   # індекс колонки Laureate (для кольору клітинки)
    pib_col_idx = 1   # індекс колонки ПІБ

    lau_cell_styles = []  # накопичуємо (row_idx, color) для клітинки Laureate

    for i, r in enumerate(sorted_rows, start=1):
        lau   = r["laureate"]
        nazva = r.get("nazva", "") or ""
        lau_color = LAUREATE_CELL_COLORS.get(lau, YELLOW)
        lau_cell_styles.append((i, lau_color))

        table_data.append([
            Paragraph(str(r["id"]) if r["id"] else "—", cell_style),
            Paragraph(r["pib"],   pib_style),          # ПІБ — жирний
            Paragraph(r["nom"],   cell_style),
            Paragraph(nazva,      cell_style),
            Paragraph(lau,        bold_style),          # Laureate — жирний
        ])

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)

    ts = [
        # Заголовок таблиці — жовтий фон
        ("BACKGROUND",    (0, 0), (-1, 0),  YELLOW),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  colors.black),
        ("FONTNAME",      (0, 0), (-1, 0),  FONT_BOLD),
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("GRID",          (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("ALIGN",         (0, 0), (-1, 0),  "CENTER"),
        # Всі рядки даних — жовтий фон
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [YELLOW]),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]

    # Кольори тільки клітинки Laureate
    for row_idx, lau_color in lau_cell_styles:
        ts.append(("BACKGROUND", (lau_col_idx, row_idx), (lau_col_idx, row_idx), lau_color))

    tbl.setStyle(TableStyle(ts))
    story.append(tbl)

    doc.build(story)
    print(f"PDF збережено: {output_path}")


# ---------------------------------------------------------------------------
# Запис у Bitrix24
# ---------------------------------------------------------------------------
LAUREATE_BITRIX = {
    "Gran Pri":   "Gran Pri",
    "1st degree": "1st degree",
    "2nd degree": "2nd degree",
    "3d degree":  "3d degree",
}

def write_to_bitrix(rows: list[dict], webhook_url: str):
    """Записує поле Laureate у Bitrix24 для кожного ID."""
    webhook_url = webhook_url.rstrip("/")
    ok = err = skip = 0
    for r in rows:
        rid = r.get("id")
        lau = r.get("laureate")
        if not rid or not lau:
            skip += 1
            continue
        try:
            rid_int = int(rid)
        except (ValueError, TypeError):
            skip += 1
            continue

        url = f"{webhook_url}/crm.deal.update.json"
        payload = {
            "id": rid_int,
            "fields": {"UF_CRM_LAUREATE": LAUREATE_BITRIX.get(lau, lau)},
        }
        try:
            resp = requests.post(url, json=payload, timeout=10)
            if resp.ok and resp.json().get("result"):
                ok += 1
            else:
                print(f"  ПОМИЛКА ID {rid_int}: {resp.text[:120]}")
                err += 1
        except Exception as e:
            print(f"  ПОМИЛКА ID {rid_int}: {e}")
            err += 1

    print(f"Bitrix24: записано {ok}, помилок {err}, пропущено {skip}")


# ---------------------------------------------------------------------------
# Звіт
# ---------------------------------------------------------------------------
def print_report(rows: list[dict]):
    from collections import Counter
    cnt = Counter(r["laureate"] for r in rows)
    print("\n=== ЗВІТ ===")
    print(f"Всього учасників: {len(rows)}")
    for lau in ["Gran Pri", "1st degree", "2nd degree", "3d degree"]:
        print(f"  {lau}: {cnt.get(lau, 0)}")
    no_id = [r for r in rows if not r["id"]]
    if no_id:
        print(f"\nБез ID ({len(no_id)}):")
        for r in no_id:
            print(f"  {r['pib']}")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Агент публікації результатів фестивалю TORONTO"
    )
    parser.add_argument("--folder",      required=True,
                        help="Папка з xlsx файлами оцінок від журі")
    parser.add_argument("--month",       required=True,
                        help='Місяць для заголовку PDF, наприклад "Квітень 2026"')
    parser.add_argument("--output",      default=None,
                        help="Шлях до вихідного PDF (за замовчуванням — поруч з --folder)")
    parser.add_argument("--publish-date", default="",
                        help='Дата публікації онлайн дипломів, наприклад "20 квітня"')
    parser.add_argument("--bitrix-url",  default="",
                        help="Webhook URL Bitrix24")
    parser.add_argument("--write-bitrix", action="store_true",
                        help="Записати Laureate у Bitrix24")
    args = parser.parse_args()

    folder = args.folder
    if not os.path.isdir(folder):
        print(f"ПОМИЛКА: папка не існує: {folder}")
        sys.exit(1)

    # Вихідний файл
    if args.output:
        output_pdf = args.output
    else:
        safe_month = re.sub(r'[^\w\s\-]', '', args.month).strip().upper().replace(" ", "_")
        output_pdf = os.path.join(
            os.path.dirname(folder),
            f"!!!РЕЗУЛЬТАТИ-{safe_month}.pdf"
        )

    print(f"Читаю файли журі з: {folder}")
    rows = read_all_jury(folder)
    if not rows:
        print("ПОМИЛКА: не знайдено жодного рядка з оцінками")
        sys.exit(1)

    print(f"\nЗнайдено учасників: {len(rows)}")
    print(f"Генерую PDF: {output_pdf}")
    build_pdf(rows, output_pdf, args.month, args.publish_date)

    if args.write_bitrix:
        if not args.bitrix_url:
            print("ПОМИЛКА: вкажіть --bitrix-url")
        else:
            print("\nЗаписую у Bitrix24...")
            write_to_bitrix(rows, args.bitrix_url)

    print_report(rows)
    print(f"\nГотово! PDF: {output_pdf}")


if __name__ == "__main__":
    import sys
    if sys.stdout.encoding != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")
    main()
