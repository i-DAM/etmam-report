# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import datetime, date, time
import base64

import pandas as pd
import streamlit as st

from analysis import read_excel_safe, build, get_pivots_for_ppt
from ppt_fill import fill_ppt

st.set_page_config(page_title="Reports 940", layout="centered")


def get_base64_image(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


amanah_b64 = get_base64_image("./assets/amanah_logo.png")
center_b64 = get_base64_image("./assets/center_logo.png")
vision_b64 = get_base64_image("./assets/vision2030_logo.png")

page_css = f"""
<style>
.stAppViewContainer {{
    position: relative;
    background: linear-gradient(to bottom, #043548 0%, #02141e 55%, #01080c 100%);
    background-attachment: fixed;
}}

main .block-container {{
    max-width: 900px;
    padding-top: 6rem;
}}

[data-testid="stHeader"] {{
    background-color: rgba(0, 0, 0, 0);
}}


.top-logos-main {{
    display: flex;
    justify-content: center;
    margin-bottom: 0.75rem;
}}

.vision-logo {{
    width: 260px;
    height: auto;
}}

.sub-logos {{
    display: flex;
    justify-content: center;
    align-items: center;
    gap: 0.0rem;
    margin-bottom: 0rem;
}}

.center-logo {{
    width: 270px;
    height: auto;
}}

.amanah-logo {{
    width: 270px;
    height: auto;
}}

.title-row {{
    margin-bottom: 0.75rem;
}}

.title-text {{
    font-size: 1.9rem;
    font-weight: 600;
    color: #ffffff;
}}

h1, h2, h3, h4, h5, h6, p, label {{
    color: #ffffff !important;
}}

section[data-testid="stFileUploader"] > div {{
    background-color: #181b24;
    border-radius: 6px;
}}

div[data-baseweb="input"] > div {{
    background-color: #181b24;
}}

button[kind="primary"] {{
    background-color: #22252f;
    border-radius: 6px;
}}
</style>
"""

st.markdown(page_css, unsafe_allow_html=True)

st.markdown(
    f"""
    <div class="top-logos-main">
        <img src="data:image/png;base64,{vision_b64}" class="vision-logo" />
    </div>
    <div class="sub-logos">
        <img src="data:image/png;base64,{center_b64}" class="center-logo" />
        <img src="data:image/png;base64,{amanah_b64}" class="amanah-logo" />
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="title-row">
        <span class="title-text">Reports 940</span>
    </div>
    """,
    unsafe_allow_html=True,
)


uploaded_file = st.file_uploader("Upload file (Excel)", type=["xlsx"])

c1, c2 = st.columns(2)
with c1:
    d_closed = st.date_input("Date", value=date.today())
with c2:
    t_closed = st.time_input("Time", value=time(8, 0))

CLOSED_DT = pd.Timestamp(datetime.combine(d_closed, t_closed))


if uploaded_file is not None:
    df0 = read_excel_safe(uploaded_file)

    out_excel = build(df0, CLOSED_DT)

    p_open_ar, p_sla_ar, p_urbi = get_pivots_for_ppt(df0, CLOSED_DT)

    ppt_template = "templates/balady_template.pptx"
    out_ppt = fill_ppt(ppt_template, p_open_ar, p_sla_ar, p_urbi, CLOSED_DT)

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "Download Excel report",
            data=out_excel.getvalue(),
            file_name="report_balady.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with c2:
        st.download_button(
            "Download PowerPoint report",
            data=out_ppt.getvalue(),
            file_name="report_balady.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
