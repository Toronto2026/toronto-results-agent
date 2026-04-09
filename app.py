# -*- coding: utf-8 -*-
"""
Streamlit web-додаток: Публікація результатів фестивалю TORONTO
"""

import io
import os
import tempfile

import streamlit as st

from agent_results import build_pdf, convert_laureate, read_jury_file, read_all_jury_with_log

# ---------------------------------------------------------------------------
# Конфігурація сторінки
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Результати фестивалю TORONTO",
    page_icon="🏆",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Sidebar — налаштування
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("⚙️ Налаштування")
    month = st.text_input("Місяць", value="Квітень 2026",
                          help="Відображається у заголовку PDF")
    publish_date = st.text_input("Дата публікації дипломів",
                                 value="",
                                 help='Наприклад: "20 квітня". Якщо порожньо — не показується.')
    st.divider()
    st.caption("ТЗ v1.0 · agent_results.py")

# ---------------------------------------------------------------------------
# Головна сторінка
# ---------------------------------------------------------------------------
st.title("🏆 Результати фестивалю TORONTO")
st.markdown("Завантажте файли оцінок від членів журі та отримайте PDF таблицю результатів.")

# ---------------------------------------------------------------------------
# Завантаження файлів
# ---------------------------------------------------------------------------
col1, col2, col3 = st.columns(3)
with col1:
    file1 = st.file_uploader("📋 Оцінки журі №1 (Лариса)", type=["xlsx"],
                              key="jury1")
with col2:
    file2 = st.file_uploader("📋 Оцінки журі №2 (Світлана)", type=["xlsx"],
                              key="jury2")
with col3:
    file3 = st.file_uploader("📋 Оцінки журі №3 (ДПМ)", type=["xlsx"],
                              key="jury3")

uploaded = [f for f in [file1, file2, file3] if f is not None]

if uploaded:
    st.info(f"Завантажено файлів: **{len(uploaded)}** з 3")
else:
    st.warning("Завантажте хоча б один файл оцінок журі")

# ---------------------------------------------------------------------------
# Кнопка запуску
# ---------------------------------------------------------------------------
run_btn = st.button(
    "▶️ Сформувати результати",
    type="primary",
    disabled=len(uploaded) == 0,
)

# ---------------------------------------------------------------------------
# Обробка
# ---------------------------------------------------------------------------
if run_btn and uploaded:
    with st.status("Обробляю файли...", expanded=True) as status:

        all_rows = []
        errors = []
        full_log = []

        # Читаємо кожен файл
        for i, uploaded_file in enumerate(uploaded, 1):
            st.write(f"📂 Читаю файл {i}: {uploaded_file.name}...")
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(uploaded_file.read())
                    tmp_path = tmp.name
                rows, log = read_jury_file(tmp_path)
                os.unlink(tmp_path)
                all_rows.extend(rows)
                full_log.extend(log)
                full_log.append("")
                no_score = sum(1 for r in rows if r.get("raw_laureate") == "None")
                no_id    = sum(1 for r in rows if not r["id"])
                st.write(f"   ✅ {len(rows)} рядків"
                         + (f" | ⚠️ без оцінки: {no_score}" if no_score else "")
                         + (f" | ⚠️ без ID: {no_id}" if no_id else ""))
                if rows and no_score == len(rows):
                    st.warning(f"   ⛔ Файл **{uploaded_file.name}**: всі рядки без оцінки (Laureate порожній). "
                               "Схоже, завантажено вихідний файл до оцінювання. "
                               "Завантажте файл **після** роботи агента журі.")
            except Exception as e:
                errors.append(f"{uploaded_file.name}: {e}")
                full_log.append(f"❌ {uploaded_file.name}: {e}")
                st.write(f"   ❌ Помилка: {e}")

        if not all_rows:
            st.error("Не знайдено жодного рядка з оцінками!")
            status.update(label="Помилка!", state="error")
            st.stop()

        # Дедуплікація по ID (захист від дублів між файлами журі)
        seen_ids = {}
        dedup_rows = []
        dedup_count = 0
        for r in all_rows:
            rid = str(r["id"]).strip() if r["id"] else None
            if rid and rid not in ("None", "—", ""):
                if rid in seen_ids:
                    dedup_count += 1
                    continue
                seen_ids[rid] = True
            dedup_rows.append(r)
        if dedup_count:
            st.warning(f"⚠️ Видалено дублів по ID: **{dedup_count}** (один учасник у кількох файлах журі)")
        all_rows = dedup_rows

        total_no_score = sum(1 for r in all_rows if r.get("raw_laureate") == "None")
        st.write(f"📊 Всього учасників: **{len(all_rows)}**")
        if total_no_score / len(all_rows) > 0.8:
            st.error(f"⛔ {total_no_score} з {len(all_rows)} учасників без оцінки ({total_no_score*100//len(all_rows)}%). "
                     "Схоже, завантажено файли **до** оцінювання агентом журі. PDF буде некоректним.")

        # Генеруємо PDF
        st.write("📄 Генерую PDF...")
        try:
            pdf_buf = io.BytesIO()
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf_path = tmp_pdf.name
            build_pdf(all_rows, tmp_pdf_path, month, publish_date)
            with open(tmp_pdf_path, "rb") as f:
                pdf_bytes = f.read()
            os.unlink(tmp_pdf_path)
            st.write("   ✅ PDF готовий")
        except Exception as e:
            st.error(f"Помилка генерації PDF: {e}")
            status.update(label="Помилка PDF!", state="error")
            st.stop()

        status.update(label="Готово! ✅", state="complete", expanded=False)

    # -------------------------------------------------------------------
    # Результати
    # -------------------------------------------------------------------
    from collections import Counter
    lau_cnt = Counter(r["laureate"] for r in all_rows)
    no_id   = [r for r in all_rows if not r["id"]]

    # Учасники без оцінки (отримали 1st degree автоматично)
    no_score_rows = [r for r in all_rows if r.get("raw_laureate") == "None"]

    # Можливі дублі: різний ID, однакові ПІБ + Назва роботи
    from collections import defaultdict
    _dup_key = {}  # (pib_norm, nazva_norm) -> list of rows
    for r in all_rows:
        pib_n   = r["pib"].strip().lower()
        nazva_n = (r.get("nazva") or "").strip().lower()
        key = (pib_n, nazva_n)
        if key not in _dup_key:
            _dup_key[key] = []
        _dup_key[key].append(r)
    # Тільки групи де >1 рядок
    dup_groups = {k: v for k, v in _dup_key.items() if len(v) > 1}
    # Розділяємо: конфлікт (різні laureate) і просто дублі (однакові laureate)
    conflict_groups = {k: v for k, v in dup_groups.items()
                       if len(set(r["laureate"] for r in v)) > 1}
    same_groups     = {k: v for k, v in dup_groups.items()
                       if len(set(r["laureate"] for r in v)) == 1}

    # Метрики
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Всього", len(all_rows))
    m2.metric("🥇 Gran Pri",    lau_cnt.get("Gran Pri",   0))
    m3.metric("🥈 1st degree",  lau_cnt.get("1st degree", 0))
    m4.metric("🥉 2nd degree",  lau_cnt.get("2nd degree", 0))
    m5.metric("🎖 3d degree",   lau_cnt.get("3d degree",  0))
    m6.metric("❓ Без оцінки→1st", len(no_score_rows),
              delta=f"-{len(no_score_rows)} перевір" if no_score_rows else None,
              delta_color="inverse")

    if conflict_groups:
        st.error(f"🚨 Знайдено **{len(conflict_groups)}** груп з РІЗНИМИ оцінками для однакового учасника+твору! Перевір вкладку «🚨 Конфлікти».")
    elif dup_groups:
        st.warning(f"⚠️ Знайдено **{len(dup_groups)}** дублікатів (різний ID, однакові ім'я+твір). Оцінки співпадають — перевір вкладку «🔁 Дублі».")

    # Кнопка скачати PDF
    safe_month = month.replace(" ", "_")
    st.download_button(
        label="⬇️ Завантажити PDF результатів",
        data=pdf_bytes,
        file_name=f"!!!РЕЗУЛЬТАТИ-{safe_month}.pdf",
        mime="application/pdf",
        type="primary",
    )

    st.divider()

    # Таблиця результатів
    tabs = st.tabs(["📋 Всі результати", "⭐ Gran Pri", "🚨 Конфлікти", "🔁 Дублі", "❓ Без оцінки → 1st", "⚠️ Без ID", "📋 Лог"])

    import pandas as pd
    df = pd.DataFrame([{
        "ID":       r["id"] or "—",
        "ПІБ Учасника": r["pib"],
        "Номінація": r["nom"],
        "Назва роботи": r["nazva"],
        "Laureate":  r["laureate"],
        "Файл журі": r["source"],
    } for r in sorted(all_rows, key=lambda x: x["pib"].lower())])

    # Кольори по laureate
    def color_row(val):
        colors_map = {
            "Gran Pri":   "background-color: #FFD700; color: #000",
            "1st degree": "background-color: #C8E6C9",
            "2nd degree": "background-color: #BBDEFB",
            "3d degree":  "background-color: #F3E5F5",
        }
        return colors_map.get(val, "")

    with tabs[0]:
        st.dataframe(
            df.style.map(color_row, subset=["Laureate"]),
            use_container_width=True,
            height=500,
        )

    with tabs[1]:
        gp = df[df["Laureate"] == "Gran Pri"]
        if len(gp):
            st.dataframe(gp.style.map(color_row, subset=["Laureate"]),
                         use_container_width=True)
        else:
            st.info("Жодного Gran Pri у цих файлах")

    # --- Вкладка: Конфлікти (різні оцінки для однакового учасника+твору) ---
    with tabs[2]:
        if not conflict_groups:
            st.success("Конфліктів не знайдено ✅ — всі однакові учасники+твори мають однакову оцінку.")
        else:
            st.error(f"🚨 {len(conflict_groups)} груп з РІЗНИМИ оцінками для одного учасника/твору. Потрібне ручне рішення!")
            for (pib_n, nazva_n), group in sorted(conflict_groups.items()):
                lau_vals = ", ".join(set(r["laureate"] for r in group))
                st.markdown(f"**{group[0]['pib']}** | *{group[0].get('nazva','') or '—'}*  →  оцінки: `{lau_vals}`")
                g_df = pd.DataFrame([{
                    "ID": r["id"] or "—", "ПІБ": r["pib"],
                    "Номінація": r["nom"], "Назва": r.get("nazva",""),
                    "Laureate": r["laureate"], "Файл журі": r["source"],
                } for r in group])
                st.dataframe(g_df.style.map(color_row, subset=["Laureate"]),
                             use_container_width=True)
                st.divider()
            csv_c = pd.DataFrame([{
                "ID": r["id"] or "—", "ПІБ": r["pib"], "Номінація": r["nom"],
                "Назва": r.get("nazva",""), "Laureate": r["laureate"], "Файл": r["source"],
            } for g in conflict_groups.values() for r in g]).to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ CSV конфліктів", csv_c,
                               f"конфлікти_{month.replace(' ','_')}.csv", "text/csv")

    # --- Вкладка: Дублі (однакові оцінки — можна прибрати) ---
    with tabs[3]:
        if not same_groups:
            st.success("Дублів не знайдено ✅")
        else:
            st.warning(f"🔁 {len(same_groups)} груп — різний ID, однакові ім'я+твір, **оцінки однакові**. "
                       "Можливо, учасник подав заявку двічі. Можна залишити або прибрати один рядок.")
            for (pib_n, nazva_n), group in sorted(same_groups.items()):
                ids = " / ".join(str(r["id"]) for r in group)
                st.markdown(f"**{group[0]['pib']}** | *{group[0].get('nazva','') or '—'}*  →  ID: `{ids}`  →  `{group[0]['laureate']}`")
            st.divider()
            csv_d = pd.DataFrame([{
                "ID": r["id"] or "—", "ПІБ": r["pib"], "Номінація": r["nom"],
                "Назва": r.get("nazva",""), "Laureate": r["laureate"], "Файл": r["source"],
            } for g in same_groups.values() for r in g]).to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ CSV дублів", csv_d,
                               f"дублі_{month.replace(' ','_')}.csv", "text/csv")

    with tabs[4]:
        if no_score_rows:
            st.warning(f"⚠️ {len(no_score_rows)} учасників не мали оцінки від журі → автоматично отримали **1st degree**. Перевір вручну.")
            no_score_df = pd.DataFrame([{
                "ID":           r["id"] or "—",
                "ПІБ Учасника": r["pib"],
                "Номінація":    r["nom"],
                "Назва роботи": r["nazva"],
                "Файл журі":    r["source"],
                "Оригінал":     r.get("raw_laureate", ""),
            } for r in sorted(no_score_rows, key=lambda x: x["pib"].lower())])
            st.dataframe(no_score_df, use_container_width=True)
            # CSV для ручної перевірки
            csv_bytes = no_score_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Завантажити список для перевірки (CSV)",
                data=csv_bytes,
                file_name=f"без_оцінки_{month.replace(' ','_')}.csv",
                mime="text/csv",
            )
        else:
            st.success("Всі учасники мають оцінку від журі ✅")

    with tabs[5]:
        no_id_df = df[df["ID"] == "—"]
        if len(no_id_df):
            st.warning(f"{len(no_id_df)} учасників без Bitrix24 ID — запис у CRM неможливий")
            st.dataframe(no_id_df, use_container_width=True)
        else:
            st.success("Всі учасники мають ID ✅")

    with tabs[6]:
        st.subheader("Лог читання файлів")
        st.caption("Детальна інформація про колонки, знайдені/відсутні дані")
        for line in full_log:
            if line.startswith("📂"):
                st.markdown(f"**{line}**")
            elif "❌" in line or "🚫" in line:
                st.error(line)
            elif "⚠️" in line:
                st.warning(line)
            elif "✅" in line:
                st.success(line)
            elif line == "":
                st.divider()
            else:
                st.text(line)

    # Помилки читання файлів
    if errors:
        with st.expander(f"⚠️ Помилки ({len(errors)})"):
            for e in errors:
                st.error(e)
