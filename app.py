# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import datetime, date, time

import pandas as pd
import streamlit as st

from analysis import read_excel_safe, build, get_pivots_for_ppt
from ppt_fill import fill_ppt


st.set_page_config(page_title="dam", layout="centered")
st.title("dam")

f = st.file_uploader("Upload file (.xlsx)", type=["xlsx"])

c1, c2 = st.columns(2)
with c1:
    d_closed = st.date_input("date", value=date.today())
with c2:
    t_closed = st.time_input("time", value=time(8, 0))

CLOSED_DT = pd.Timestamp(datetime.combine(d_closed, t_closed))

if f:
    df0 = read_excel_safe(f)

    out_excel = build(df0, CLOSED_DT)

    p_open_ar, p_sla_ar, p_urbi = get_pivots_for_ppt(df0, CLOSED_DT)

    ppt_template = "templates/balady_template.pptx"
    out_ppt = fill_ppt(ppt_template, p_open_ar, p_sla_ar, p_urbi, CLOSED_DT)

    c1, c2 = st.columns(2)

    with c1:
        st.download_button(
            "تنزيل التقرير (Excel)",
            data=out_excel.getvalue(),
            file_name="report_balady.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with c2:
        st.download_button(
            "تنزيل العرض (PowerPoint)",
            data=out_ppt.getvalue(),
            file_name="report_balady.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    st.success("تم تجهيز التقارير.")
