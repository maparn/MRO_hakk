"""Browser interface for the Magnetic MRO pricing workflow."""

from __future__ import annotations

import json
from pathlib import Path

import streamlit as st

from scripts.exchange_rates import get_ecb_rate
from scripts.web_processing import process_workbooks


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
    "stat": ("1. STAT fail", "Sisaldab INITIAL ja AVG lehti"),
    "amos": ("2. AMOS hinnad", "Uue süsteemi hinnainfo"),
    "uom": ("3. AMOS UOM", "Mõõtühikute teisendusreeglid"),
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
    price_labels = {
        "NEW_PRICE_HIGHER_OR_EQUAL": "Uus hind on suurem",
        "OLD_PRICE_HIGHER": "Vana hind on suurem",
        "NEW_PRICE_ONLY": "Ainult uus hind",
        "OLD_PRICE_PLUS_MARKUP_FALLBACK": "Vana hind + 20%",
        "PARTIAL_NEW_PRICE_HIGHER": "Osaline uus hind on suurem",
        "PARTIAL_NEW_PRICE_ONLY": "Ainult osaline uus hind",
        "REVIEW_REQUIRED": "Vajab hinnakontrolli",
    }
    approval_labels = {
        "WITHIN_LIMIT": "Limiidi piires",
        "REVIEW_OVER_LIMIT": "Vajab kinnitamist",
        "PRICE_REVIEW_REQUIRED": "Hind vajab kontrolli",
    }
    services = [row for row in result["summary"] if row["task_card"] != "SUM_AVG"]
    total = next(row for row in result["summary"] if row["task_card"] == "SUM_AVG")
    st.success("Raport on valmis.")
    metric_cols = st.columns(4)
    metric_cols[0].metric("Teenuseid", len(services))
    metric_cols[1].metric("Lõppsumma EUR", f'{total["final_service_eur"]:,.2f} €')
    metric_cols[2].metric("Lõppsumma USD", f'${total["final_service_usd"]:,.2f}')
    metric_cols[3].metric(
        "EUR/USD kurss",
        f'{result["exchange"]["rate"]} ({result["exchange"]["status"]})',
        result["exchange"]["date"],
    )

    table = [{
        "Reference": row["reference"],
        "Task Card": row["task_card"],
        "Vana AVG EUR": row["historical_avg_eur"],
        "Uus hind EUR": row["calculated_line_sum_eur"],
        "Lõpphind EUR": row["final_service_eur"],
        "Hinna valik": price_labels.get(row["final_price_source"], row["final_price_source"]),
        "Kinnitamine": approval_labels.get(row["approval_status"], row["approval_status"]),
        "Lõpphind USD": row["final_service_usd"],
    } for row in services]
    st.subheader("Teenuste kokkuvõte")
    st.dataframe(table, use_container_width=True, hide_index=True)
    st.download_button(
        "LAADI EXCELI RAPORT ALLA",
        data=result["report"], file_name=result["filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown(
    '<p class="privacy">Failid töödeldakse eraldatud ajutises kaustas ja neid ei '
    'salvestata püsivalt serverisse.</p>', unsafe_allow_html=True)
