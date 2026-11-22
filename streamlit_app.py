"""Browser interface for the Magnetic MRO pricing workflow."""

from __future__ import annotations

import json
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

from scripts.exchange_rates import get_ecb_rate
from scripts.web_processing import load_converter_html, process_workbooks


ROOT = Path(__file__).resolve().parent
CONFIG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))

st.set_page_config(
    page_title="Magnetic MRO hinnakalkulaator",
    layout="wide",
)

st.markdown("""
<style>
  .stApp { background: #f5f7fa; color: #17212b; }
  [data-testid="stMainBlockContainer"] { max-width: 1180px; padding-top: 2.2rem; }
  .hero { background: linear-gradient(135deg, #103b5d 0%, #1f6b83 100%);
          color: white; padding: 2rem 2.2rem; border-radius: 18px; margin-bottom: 1.4rem; }
  .hero h1 { margin: 0 0 .45rem 0; font-size: 2rem; }
  .hero p { margin: 0; font-size: 1.05rem; opacity: .9; }
  .step { background: white; border: 1px solid #dce4ea; border-radius: 14px;
          padding: 1rem 1.1rem; min-height: 112px; }
  .step strong { display: block; color: #103b5d; margin-bottom: .35rem; }
  div.stButton > button, div.stDownloadButton > button {
      min-height: 48px; border-radius: 10px; font-weight: 700; }
  div.stButton > button { background: #d66a2c; color: white; border: 0; }
  .privacy { color: #52606d; font-size: .92rem; }
</style>
<div class="hero">
  <h1>Magnetic MRO hinnakalkulaator</h1>
  <p>Laadi kolm Exceli faili, arvuta teenuste hinnad ja laadi alla üks kontrollitud raport.</p>
</div>
""", unsafe_allow_html=True)

cols = st.columns(3)
labels = {
    "stat": ("1. Work Package STAT", "Aktiivne tööpakett: INITIAL ja võimalusel AVG"),
    "amos": ("2. AMOS hinnad", "Korduvkasutatav uue süsteemi hinnainfo"),
    "uom": ("3. AMOS UOM", "Korduvkasutatavad mõõtühikute teisendusreeglid"),
}
uploads = {}


def clear_result():
    st.session_state.pop("mro_result", None)


for column, (name, (title, help_text)) in zip(cols, labels.items()):
    with column:
        st.markdown(f'<div class="step"><strong>{title}</strong>{help_text}</div>',
                    unsafe_allow_html=True)
        uploads[name] = st.file_uploader(
            title, type=["xlsx"], key=f"upload_{name}", label_visibility="collapsed",
            on_change=clear_result)

st.markdown("<div style='height:.5rem'></div>", unsafe_allow_html=True)
ready = all(uploads.values())
if st.button("ARVUTA JA LOO RAPORT", type="primary", use_container_width=True,
             disabled=not ready):
    try:
        with st.spinner("Kontrollin faile ja arvutan hindu..."):
            exchange_config = CONFIG["exchange_rate"]
            exchange = get_ecb_rate(
                exchange_config["url"], ROOT / exchange_config["cache_file"],
                currency=exchange_config.get("target_currency", "USD"),
                timeout=float(exchange_config.get("timeout_seconds", 10)),
            )
            result = process_workbooks(
                {name: upload.getvalue() for name, upload in uploads.items()},
                CONFIG, exchange,
            )
            st.session_state["mro_result"] = result
    except Exception as error:
        st.session_state.pop("mro_result", None)
        st.error(f"Raporti loomine ebaõnnestus: {error}")

result = st.session_state.get("mro_result")
if result:
    services = [row for row in result["task_summary"] if row["task_card"] != "SUM_AVG"]
    total = next(row for row in result["summary"] if row["task_card"] == "SUM_AVG")
    st.success("Raport on valmis.")
    metric_cols = st.columns(4)
    metric_cols[0].metric("Materjaliread", result["stats"]["stat_rows"])
    metric_cols[1].metric("Lõppsumma EUR", f'{total["final_task_card_price_eur"]:,.2f} €')
    metric_cols[2].metric("Lõppsumma USD", f'${total["final_task_card_price_usd"]:,.2f}')
    metric_cols[3].metric(
        "EUR/USD kurss",
        f'{result["exchange"]["rate"]} ({result["exchange"]["status"]})',
        result["exchange"]["date"],
    )

    if result["stats"]["unresolved_rows"]:
        st.warning(f'Lahendamata materjaliridu: {result["stats"]["unresolved_rows"]}. '
                   'Neid ei esitata kehtiva nullhinnana; vaata erandite lehte raportis.')
    task_table = [{
        "Task Card": row["task_card"],
        "Reference arv": row["reference_count"],
        "Täielikud Reference-id": row["complete_references"],
        "Puudulikud Reference-id": row["incomplete_references"],
        "Uus AVG EUR": row["new_task_card_avg_eur"],
        "Vana AVG EUR": row["old_task_card_avg_eur"],
        "Lõpphind EUR": row["final_task_card_price_eur"],
        "Valitud allikas": row["selected_price_source"],
        "Lõpphind USD": row["final_task_card_price_usd"],
        "Olek": row["status"],
    } for row in services]
    material_table = [{
        "Rida": row["source_row"], "Reference": row["reference"],
        "Task Card": row["task_card"], "Part Number": row["partno"],
        "Kogus": row["original_quantity"], "Ühik": row["original_unit"],
        "Ajalooline hind EUR": row["historical_unit_price_eur"],
        "AMOS algne hind EUR": row["amos_avg_price_eur"],
        "Hind STAT ühikus EUR": row["selected_unit_price_eur"],
        "Teisendustegur": row["price_conversion_factor"],
        "Teisenduse allikas": row["conversion_source"],
        "Teisenduse olek": row["conversion_status"],
        "Rea summa EUR": row["final_line_eur"], "Hinna allikas": row["price_source"],
        "Olek": row["status_et"],
    } for row in result["material_calculations"]]
    task_tab, material_tab = st.tabs(["Task Card kokkuvõte", "Materjaliread"])
    with task_tab:
        st.dataframe(task_table, use_container_width=True, hide_index=True)
    with material_tab:
        st.dataframe(material_table, use_container_width=True, hide_index=True)
    st.download_button(
        "LAADI EXCELI RAPORT ALLA",
        data=result["report"], file_name=result["filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown(
    '<p class="privacy">Failid töödeldakse eraldatud ajutises kaustas ja neid ei '
    'salvestata püsivalt serverisse.</p>', unsafe_allow_html=True)


st.divider()
st.subheader("Mõõtühikute teisendamine")

converter_html = load_converter_html(ROOT)

if converter_html is None:
    st.info("Mõõtühikute konverter ei ole praegu saadaval. Hinnakalkulaator töötab edasi.")
else:
    with st.expander("Ava ühikute konverter"):
        components.html(converter_html, height=690, scrolling=True)
