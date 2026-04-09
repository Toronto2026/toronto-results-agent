# -*- coding: utf-8 -*-
"""
Streamlit web-додаток: Публікація результатів фестивалю TORONTO
"""

import io, os, json, tempfile
from datetime import datetime
from collections import Counter, defaultdict

import pandas as pd
import streamlit as st

from agent_results import (build_pdf, read_jury_file, write_to_bitrix,
                           import_results_from_excel, import_results_from_pdf)

# ---------------------------------------------------------------------------
# Конфігурація сторінки
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Результати фестивалю TORONTO",
    page_icon="🏆",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("⚙️ Налаштування")
    month = st.text_input("Місяць", value="Квітень 2026",
                          help="Відображається у заголовку PDF")
    publish_date = st.text_input("Дата публікації дипломів", value="",
                                 help='Наприклад: "20 квітня"')
    st.divider()
    st.subheader("🔗 Bitrix24")
    bitrix_url_input = st.text_input(
        "Webhook URL", type="password",
        value=st.session_state.get("bitrix_url", ""),
        placeholder="https://toronto.bitrix24.com/rest/1/.../",
        key="bitrix_url_input",
    )
    if bitrix_url_input:
        st.session_state["bitrix_url"] = bitrix_url_input
    bx_write_lau = st.checkbox("Записати Laureate",      value=True)
    bx_write_com = st.checkbox("Записати Коментар Журі", value=True)
    st.divider()
    st.caption("v2.0 · agent_results.py")

# ---------------------------------------------------------------------------
# Заголовок
# ---------------------------------------------------------------------------
st.title("🏆 Результати фестивалю TORONTO")

# ---------------------------------------------------------------------------
# Хелпери
# ---------------------------------------------------------------------------
def rows_to_json(rows, month_str) -> bytes:
    return json.dumps({
        "month":      month_str,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "total":      len(rows),
        "results":    rows,
    }, ensure_ascii=False, indent=2).encode("utf-8")

def json_to_rows(data: dict):
    return data.get("results", [])

def color_cell(val):
    m = {"Gran Pri":"background-color:#FFE082",
         "1st degree":"background-color:#DCEDC8",
         "2nd degree":"background-color:#BBDEFB",
         "3d degree":"background-color:#E1BEE7"}
    return m.get(val, "")

def build_df(rows):
    return pd.DataFrame([{
        "ID":            r.get("id") or "—",
        "ПІБ Учасника":  r.get("pib",""),
        "Номінація":     r.get("nom",""),
        "Назва роботи":  r.get("nazva",""),
        "Laureate":      r.get("laureate",""),
        "Коментар":      r.get("comment",""),
        "Файл журі":     r.get("source",""),
    } for r in sorted(rows, key=lambda x: x.get("pib","").lower())])

def dedup_by_id(rows):
    seen, out, count = {}, [], 0
    for r in rows:
        rid = str(r.get("id","")).strip()
        if rid and rid not in ("None","—",""):
            if rid in seen:
                count += 1
                continue
            seen[rid] = True
        out.append(r)
    return out, count

def find_duplicates(rows):
    groups = defaultdict(list)
    for r in rows:
        key = (r.get("pib","").strip().lower(),
               (r.get("nazva","") or "").strip().lower())
        groups[key].append(r)
    conflicts = {k:v for k,v in groups.items() if len(v)>1 and len(set(x["laureate"] for x in v))>1}
    same      = {k:v for k,v in groups.items() if len(v)>1 and len(set(x["laureate"] for x in v))==1}
    return conflicts, same

# ---------------------------------------------------------------------------
# РЕЖИМ 1: завантажити результати попереднього конкурсу (JSON / XLSX / PDF)
# ---------------------------------------------------------------------------
st.subheader("📂 Завантажити результати попереднього конкурсу")

prev_file = st.file_uploader(
    "JSON (збережений агентом) · XLSX · PDF",
    type=["json", "xlsx", "pdf"],
    key="prev_upload",
    help="Завантажте результати будь-якого попереднього конкурсу",
)

loaded_rows = None
loaded_meta = {}

if prev_file:
    ext = prev_file.name.rsplit(".", 1)[-1].lower()
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp:
            tmp.write(prev_file.read())
            tmp_path = tmp.name

        if ext == "json":
            data = json.loads(open(tmp_path, encoding="utf-8").read())
            loaded_rows = json_to_rows(data)
            loaded_meta = {k: v for k, v in data.items() if k != "results"}

        elif ext == "xlsx":
            loaded_rows, imp_log = import_results_from_excel(tmp_path)
            loaded_meta = {"month": month}
            with st.expander("Лог імпорту Excel"):
                for line in imp_log:
                    st.text(line)

        elif ext == "pdf":
            loaded_rows, imp_log = import_results_from_pdf(tmp_path)
            loaded_meta = {"month": month}
            with st.expander("Лог імпорту PDF"):
                for line in imp_log:
                    st.text(line)

        os.unlink(tmp_path)

        if loaded_rows:
            upd = loaded_meta.get("updated_at", "—")
            mon = loaded_meta.get("month", "—")
            st.success(
                f"✅ Завантажено **{len(loaded_rows)}** учасників"
                + (f" | Конкурс: **{mon}**" if mon != "—" else "")
                + (f" | Оновлено: **{upd}**" if upd != "—" else "")
            )
        else:
            st.error("Не вдалося знайти рядки результатів у файлі. "
                     "Перевір формат: потрібні колонки ID, ПІБ Учасника, Laureate.")

    except Exception as e:
        st.error(f"Помилка читання файлу: {e}")
        loaded_rows = None

st.divider()

# ---------------------------------------------------------------------------
# РЕЖИМ 2: завантажити файли журі
# ---------------------------------------------------------------------------
st.subheader("📋 Або завантажте нові файли від журі")
col1, col2, col3 = st.columns(3)
with col1:
    file1 = st.file_uploader("Журі №1 (Лариса)", type=["xlsx"], key="jury1")
with col2:
    file2 = st.file_uploader("Журі №2 (Світлана)", type=["xlsx"], key="jury2")
with col3:
    file3 = st.file_uploader("Журі №3 (ДПМ)", type=["xlsx"], key="jury3")

uploaded = [f for f in [file1, file2, file3] if f is not None]

run_btn = st.button(
    "▶️ Сформувати результати з файлів журі",
    type="primary",
    disabled=len(uploaded) == 0,
)

# ---------------------------------------------------------------------------
# Обробка файлів журі
# ---------------------------------------------------------------------------
if run_btn and uploaded:
    with st.status("Обробляю файли...", expanded=True) as status:
        all_rows, errors, full_log = [], [], []

        for i, uf in enumerate(uploaded, 1):
            st.write(f"📂 Читаю файл {i}: {uf.name}...")
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(uf.read())
                    tmp_path = tmp.name
                rows, log = read_jury_file(tmp_path)
                os.unlink(tmp_path)
                all_rows.extend(rows)
                full_log.extend(log + [""])
                no_score = sum(1 for r in rows if r.get("raw_laureate") == "None")
                no_id    = sum(1 for r in rows if not r["id"])
                st.write(f"   ✅ {len(rows)} рядків"
                         + (f" | ⚠️ без оцінки: {no_score}" if no_score else "")
                         + (f" | ⚠️ без ID: {no_id}"        if no_id    else ""))
                if rows and no_score == len(rows):
                    st.warning(f"⛔ **{uf.name}**: всі рядки без оцінки — завантажено файл ДО оцінювання!")
            except Exception as e:
                errors.append(f"{uf.name}: {e}")
                st.write(f"   ❌ Помилка: {e}")

        if not all_rows:
            st.error("Не знайдено жодного рядка!")
            status.update(label="Помилка!", state="error")
            st.stop()

        all_rows, dedup_count = dedup_by_id(all_rows)
        if dedup_count:
            st.warning(f"⚠️ Видалено дублів по ID: **{dedup_count}**")

        total_none = sum(1 for r in all_rows if r.get("raw_laureate") == "None")
        st.write(f"📊 Всього учасників: **{len(all_rows)}**")
        if all_rows and total_none / len(all_rows) > 0.8:
            st.error(f"⛔ {total_none}/{len(all_rows)} без оцінки — файли до оцінювання!")

        st.write("📄 Генерую PDF...")
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf_path = tmp_pdf.name
            build_pdf(all_rows, tmp_pdf_path, month, publish_date)
            with open(tmp_pdf_path, "rb") as f:
                pdf_bytes = f.read()
            os.unlink(tmp_pdf_path)
            st.write("   ✅ PDF готовий")
        except Exception as e:
            st.error(f"Помилка PDF: {e}")
            status.update(label="Помилка PDF!", state="error")
            st.stop()

        status.update(label="Готово! ✅", state="complete", expanded=False)

    # Зберігаємо у session_state
    st.session_state["all_rows"]  = all_rows
    st.session_state["full_log"]  = full_log
    st.session_state["pdf_bytes"] = pdf_bytes
    st.session_state["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    st.session_state["result_month"] = month

# Якщо завантажений JSON — кладемо у session_state
if loaded_rows is not None:
    st.session_state["all_rows"]   = loaded_rows
    st.session_state["full_log"]   = []
    st.session_state["pdf_bytes"]  = None
    st.session_state["updated_at"] = loaded_meta.get("updated_at", "—")
    st.session_state["result_month"] = loaded_meta.get("month", month)

# ---------------------------------------------------------------------------
# Відображення результатів (якщо є у session_state)
# ---------------------------------------------------------------------------
if "all_rows" not in st.session_state:
    st.stop()

all_rows    = st.session_state["all_rows"]
full_log    = st.session_state.get("full_log", [])
pdf_bytes   = st.session_state.get("pdf_bytes")
updated_at  = st.session_state.get("updated_at", "—")
res_month   = st.session_state.get("result_month", month)

st.divider()

# Заголовок результатів
st.subheader(f"📊 Результати — {res_month}")
st.caption(f"🕐 Останнє оновлення: **{updated_at}**")

lau_cnt       = Counter(r["laureate"] for r in all_rows)
no_score_rows = [r for r in all_rows if r.get("raw_laureate") == "None"]
conflict_groups, same_groups = find_duplicates(all_rows)
dup_groups    = {**conflict_groups, **same_groups}
country_cnt   = Counter(r.get("country", "Україна") for r in all_rows)
country_count = len(country_cnt)

# Метрики
m1,m2,m3,m4,m5,m6,m7 = st.columns(7)
m1.metric("Всього",            len(all_rows))
m2.metric("🌍 Країн",          country_count)
m3.metric("🥇 Gran Pri",       lau_cnt.get("Gran Pri",   0))
m4.metric("🥈 1st degree",     lau_cnt.get("1st degree", 0))
m5.metric("🥉 2nd degree",     lau_cnt.get("2nd degree", 0))
m6.metric("🎖 3d degree",      lau_cnt.get("3d degree",  0))
m7.metric("❓ Без оцінки→1st", len(no_score_rows))

# Розбивка по країнах
if country_count > 1:
    with st.expander(f"🌍 Країни ({country_count})"):
        country_df = pd.DataFrame(
            [{"Країна": c, "Учасників": n} for c, n in sorted(country_cnt.items(), key=lambda x: -x[1])]
        )
        st.dataframe(country_df, use_container_width=True, hide_index=True)

if conflict_groups:
    st.error(f"🚨 **{len(conflict_groups)}** конфліктів — однаковий учасник+твір, різні оцінки! → вкладка «🚨 Конфлікти»")
elif dup_groups:
    st.warning(f"🔁 **{len(dup_groups)}** дублів (однакові оцінки) → вкладка «🔁 Дублі»")

# Кнопки завантаження
safe_month = res_month.replace(" ", "_")
btn1, btn2 = st.columns(2)
with btn1:
    if pdf_bytes:
        st.download_button("⬇️ Завантажити PDF",
                           data=pdf_bytes,
                           file_name=f"!!!РЕЗУЛЬТАТИ-{safe_month}.pdf",
                           mime="application/pdf", type="primary")
    else:
        if st.button("🔄 Згенерувати PDF", type="secondary"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf_path = tmp_pdf.name
            build_pdf(all_rows, tmp_pdf_path, res_month, publish_date)
            with open(tmp_pdf_path, "rb") as f:
                st.session_state["pdf_bytes"] = f.read()
            os.unlink(tmp_pdf_path)
            st.rerun()

with btn2:
    json_bytes = rows_to_json(all_rows, res_month)
    st.download_button(
        "💾 Зберегти результати (JSON)",
        data=json_bytes,
        file_name=f"результати-{safe_month}.json",
        mime="application/json",
        help="Завантажте цей файл щоб наступного разу не завантажувати xlsx файли журі",
    )

st.divider()

# ---------------------------------------------------------------------------
# Вкладки
# ---------------------------------------------------------------------------
df = build_df(all_rows)

tabs = st.tabs([
    "📋 Всі результати", "⭐ Gran Pri",
    "🚨 Конфлікти", "🔁 Дублі",
    "❓ Без оцінки", "⚠️ Без ID",
    "📋 Лог",
])

with tabs[0]:
    st.dataframe(df.style.map(color_cell, subset=["Laureate"]),
                 use_container_width=True, height=520)

with tabs[1]:
    gp = df[df["Laureate"] == "Gran Pri"]
    st.dataframe(gp.style.map(color_cell, subset=["Laureate"]),
                 use_container_width=True) if len(gp) else st.info("Жодного Gran Pri")

with tabs[2]:
    if not conflict_groups:
        st.success("Конфліктів не знайдено ✅")
    else:
        st.error(f"🚨 {len(conflict_groups)} конфліктів — потрібне ручне рішення!")
        for (p,n), grp in sorted(conflict_groups.items()):
            lau_vals = " / ".join(sorted(set(r["laureate"] for r in grp)))
            st.markdown(f"**{grp[0]['pib']}** | *{grp[0].get('nazva','') or '—'}* → `{lau_vals}`")
            st.dataframe(build_df(grp).style.map(color_cell, subset=["Laureate"]),
                         use_container_width=True)
            st.divider()
        csv_c = build_df([r for g in conflict_groups.values() for r in g]).to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ CSV конфліктів", csv_c, f"конфлікти-{safe_month}.csv", "text/csv")

with tabs[3]:
    if not same_groups:
        st.success("Дублів не знайдено ✅")
    else:
        st.warning(f"🔁 {len(same_groups)} груп дублів — оцінки однакові, але різні ID")
        for (p,n), grp in sorted(same_groups.items()):
            ids = " / ".join(str(r["id"]) for r in grp)
            st.markdown(f"**{grp[0]['pib']}** | *{grp[0].get('nazva','') or '—'}* → ID: `{ids}` → `{grp[0]['laureate']}`")
        csv_d = build_df([r for g in same_groups.values() for r in g]).to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ CSV дублів", csv_d, f"дублі-{safe_month}.csv", "text/csv")

with tabs[4]:
    if no_score_rows:
        st.warning(f"⚠️ {len(no_score_rows)} учасників без оцінки → автоматично 1st degree")
        ns_df = build_df(no_score_rows)
        st.dataframe(ns_df, use_container_width=True)
        st.download_button("⬇️ CSV для перевірки",
                           ns_df.to_csv(index=False).encode("utf-8-sig"),
                           f"без_оцінки-{safe_month}.csv", "text/csv")
    else:
        st.success("Всі учасники мають оцінку ✅")

with tabs[5]:
    no_id_df = df[df["ID"] == "—"]
    if len(no_id_df):
        st.warning(f"{len(no_id_df)} учасників без Bitrix24 ID")
        st.dataframe(no_id_df, use_container_width=True)
    else:
        st.success("Всі учасники мають ID ✅")

with tabs[6]:
    if full_log:
        for line in full_log:
            if   line.startswith("📂"):  st.markdown(f"**{line}**")
            elif "❌" in line or "🚫" in line: st.error(line)
            elif "⚠️" in line:          st.warning(line)
            elif "✅" in line:           st.success(line)
            elif line == "":             st.divider()
            else:                        st.text(line)
    else:
        st.info("Лог доступний тільки після завантаження файлів журі (не з JSON)")

# ---------------------------------------------------------------------------
# Bitrix24
# ---------------------------------------------------------------------------
st.divider()
st.subheader("🔗 Записати результати у Bitrix24")

bx_url = st.session_state.get("bitrix_url", "").strip()
rows_with_id = [r for r in all_rows if r.get("id") and str(r.get("id","")).strip().isdigit()]
rows_no_id   = len(all_rows) - len(rows_with_id)

c1,c2,c3 = st.columns(3)
c1.metric("Будуть записані", len(rows_with_id))
c2.metric("Без ID (пропуск)", rows_no_id)
c3.metric("Поля", "Laureate + Коментар" if bx_write_lau and bx_write_com
          else "Тільки Laureate" if bx_write_lau else "Тільки Коментар")

if not bx_url:
    st.info("👈 Введіть Webhook URL у бічній панелі")
else:
    st.caption(f"Webhook: `{bx_url[:45]}...`")

    if st.button("▶️ Записати у Bitrix24", type="primary", key="bx_run",
                 disabled=not bx_url):

        progress_bar = st.progress(0.0, text="Починаю...")
        bx_log = []   # список dict для таблиці логу

        def on_progress(done, total, row, status):
            frac = done / total
            pib  = (row.get("pib") or "")[:40]
            rid  = row.get("id","?")
            if status == "ok":
                icon, msg = "✅", ""
                progress_bar.progress(frac, text=f"✅ {done}/{total} — {pib}")
            elif status == "skip":
                icon, msg = "⏭", "немає ID або значення"
                progress_bar.progress(frac, text=f"⏭ {done}/{total} — пропущено")
            else:
                icon, msg = "❌", status[4:80]
                progress_bar.progress(frac, text=f"❌ {done}/{total} — {pib}: помилка")
            bx_log.append({"Статус": icon, "ID": rid, "ПІБ": pib,
                           "Laureate": row.get("laureate",""), "Деталі": msg})

        with st.spinner("Записую у Bitrix24..."):
            result = write_to_bitrix(
                rows_with_id, bx_url,
                write_laureate=bx_write_lau,
                write_comment=bx_write_com,
                progress_cb=on_progress,
            )

        progress_bar.progress(1.0, text="Готово!")

        # Метрики результату
        r1,r2,r3 = st.columns(3)
        r1.metric("✅ Записано",  result["ok"])
        r2.metric("❌ Помилок",   result["err"],
                  delta=f"-{result['err']}" if result["err"] else None,
                  delta_color="inverse")
        r3.metric("⏭ Пропущено", result["skip"])

        if result["err"] == 0:
            st.success(f"✅ Всі {result['ok']} записів збережені у Bitrix24!")
        else:
            st.warning(f"⚠️ {result['err']} помилок з {len(rows_with_id)}")

        # Детальний лог
        with st.expander("📋 Детальний лог Bitrix24", expanded=result["err"] > 0):
            log_df = pd.DataFrame(bx_log)
            def color_status(val):
                return ("color: green" if val == "✅" else
                        "color: red"   if val == "❌" else "color: grey")
            st.dataframe(log_df.style.map(color_status, subset=["Статус"]),
                         use_container_width=True, height=400)
            # CSV логу
            st.download_button(
                "⬇️ Завантажити лог CSV",
                log_df.to_csv(index=False).encode("utf-8-sig"),
                f"bitrix24-лог-{safe_month}.csv", "text/csv",
            )
