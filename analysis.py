# -*- coding: utf-8 -*-
from __future__ import annotations

import re
import importlib.util
import tempfile
from io import BytesIO

import numpy as np
import pandas as pd


COL_ID = "رقم البلاغ"
COL_ADMIN = "الإدارة"
COL_STATUS = "حالة البلاغ في النظام"
COL_SOURCE = "مصدر البلاغ"
COL_CREATED = "تاريخ الانشاء"

COL_CLOSED = "تاريخ الاغلاق"
COL_ELAPSED = "الوقت المنقضي"
COL_HOURS = "الساعات"
COL_SLA = "التوصيف"

STATUS_CANON = [
    "انتظار الاستجابة - مقاول",
    "انتظار الاستجابة - مراقب",
    "انتظار الاستجابة - مشرف",
    "قيد التنفيذ - مراقب",
    "قيد التنفيذ - مقاول",
    "معلق - اعادة فتح",
]
SLA_ORDER = ["لم تتجاوز", "قارب على تجاوز SLA", "تجاوز SLA"]

ALLOWED_SOURCES = {"Urbi", "تطبيق بلدي", "توكلنا", "مراكز الاتصال"}


def _norm_ar(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s = (
        s.replace("أ", "ا")
         .replace("إ", "ا")
         .replace("آ", "ا")
         .replace("ى", "ي")
         .replace("ة", "ه")
    )
    s = re.sub(r"\s+", " ", s)
    s = s.replace(" -", "-").replace("- ", "-")
    return s


def canon_status(raw: str) -> str:
    s = _norm_ar(raw)
    if "انتظار" in s and "استجابه" in s:
        if "مقاول" in s:
            return "انتظار الاستجابة - مقاول"
        if "مراقب" in s:
            return "انتظار الاستجابة - مراقب"
        if "مشرف" in s:
            return "انتظار الاستجابة - مشرف"
    if ("قيد" in s and "التنفيذ" in s) or ("التنفيذ" in s and ("قيد" in s or "جاري" in s)):
        if "مراقب" in s:
            return "قيد التنفيذ - مراقب"
        if "مقاول" in s:
            return "قيد التنفيذ - مقاول"
    if "معلق" in s and ("اعاده فتح" in s or "اعادة فتح" in s or "فتح" in s):
        return "معلق - اعادة فتح"
    return raw


def canon_source(x: str) -> str:
    s = _norm_ar(x)
    if "urbi" in s.lower():
        return "Urbi"
    if "بلدي" in s:
        return "تطبيق بلدي"
    if "توكلنا" in s:
        return "توكلنا"
    if ("مراكز" in s and "اتصال" in s) or ("مركز" in s and "اتصال" in s):
        return "مراكز الاتصال"
    return x


def read_excel_safe(uploaded):
    data = uploaded.read()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(data)
        path = tmp.name

    engines = []
    if importlib.util.find_spec("calamine") or importlib.util.find_spec("python_calamine"):
        engines.append("calamine")
    engines += [None, "openpyxl"]

    last = None
    for eng in engines:
        try:
            return pd.read_excel(path, sheet_name=0, engine=eng)
        except Exception as e:
            last = e
            continue

    raise RuntimeError(f"تعذر قراءة الملف: {last}")


def parse_created_col(s):
    a = pd.to_datetime(s, errors="coerce", dayfirst=False)
    b = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if b.notna().sum() > a.notna().sum():
        return b
    return a


def preprocess(df: pd.DataFrame, closed_ts: pd.Timestamp):
    idx_created = df.columns.get_loc(COL_CREATED)

    drop_after = [
        df.columns[i]
        for i in range(idx_created + 1, min(idx_created + 4, len(df.columns)))
    ]
    df = df.drop(columns=drop_after, errors="ignore").copy()

    sidx = df.columns.get_loc(COL_SOURCE)
    right_cols = list(df.columns[sidx + 1:])
    if right_cols:
        df.drop(columns=right_cols, inplace=True)

    NEW_CLASS_COL = "التصنيف الجديد"
    BAD_CLASS = "السيارات التالفة"

    rows_before = len(df)

    if NEW_CLASS_COL in df.columns:
        df = df[
            df[NEW_CLASS_COL]
            .astype(str)
            .map(_norm_ar)
            != _norm_ar(BAD_CLASS)
        ].copy()

    rows_after = len(df)
    deleted = rows_before - rows_after

    df["_deleted_bad_class"] = deleted
    df["_rows_before_filter"] = rows_before
    df["_rows_after_filter"] = rows_after

    if COL_CLOSED not in df.columns:
        df.insert(idx_created + 1, COL_CLOSED, pd.NaT)
    if COL_ELAPSED not in df.columns:
        df.insert(idx_created + 2, COL_ELAPSED, pd.NaT)
    if COL_HOURS not in df.columns:
        df.insert(idx_created + 3, COL_HOURS, pd.NA)

    df[COL_CREATED] = parse_created_col(df[COL_CREATED])
    df[COL_CLOSED] = closed_ts

    df[COL_ELAPSED] = df[COL_CLOSED] - df[COL_CREATED]

    total_seconds = df[COL_ELAPSED].dt.total_seconds()
    hours_floor = np.floor_divide(total_seconds.fillna(0).astype("int64"), 3600)
    df[COL_HOURS] = pd.Series(hours_floor, index=df.index).astype("Int64")

    df["_status_canon"] = df[COL_STATUS].astype(str).map(canon_status)
    df = df[df["_status_canon"].isin(STATUS_CANON)].copy()
    df["_status_canon"] = pd.Categorical(
        df["_status_canon"], categories=STATUS_CANON, ordered=True
    )

    df[COL_SOURCE] = df[COL_SOURCE].astype(str).map(canon_source)
    df = df[df[COL_SOURCE].isin(ALLOWED_SOURCES)].copy()

    h = df[COL_HOURS].astype("float")
    conds = [h < 72, (h >= 72) & (h <= 95)]
    choices = ["لم تتجاوز", "قارب على تجاوز SLA"]
    df[COL_SLA] = np.select(conds, choices, default="تجاوز SLA")
    df[COL_SLA] = pd.Categorical(df[COL_SLA], categories=SLA_ORDER, ordered=True)

    return df



def pivots(df_proc: pd.DataFrame):
    p_open = (
        df_proc.groupby([COL_ADMIN, "_status_canon"])[COL_ID]
        .count()
        .unstack("_status_canon", fill_value=0)
        .reindex(columns=STATUS_CANON, fill_value=0)
    )
    p_open.loc["الإجمالي الكلي"] = p_open.sum(axis=0)
    p_open.columns.name = None

    p_sla = (
        df_proc.groupby([COL_ADMIN, COL_SLA])[COL_ID]
        .count()
        .unstack(COL_SLA, fill_value=0)
        .reindex(columns=SLA_ORDER, fill_value=0)
    )
    p_sla.loc["الإجمالي الكلي"] = p_sla.sum(axis=0)
    p_sla.columns.name = None

    return p_open, p_sla


def build(xls: pd.DataFrame, closed_ts: pd.Timestamp) -> BytesIO:
    df_all = preprocess(xls.copy(), closed_ts)

    df_ar = df_all[df_all[COL_SOURCE] != "Urbi"].copy()
    p_open_ar, p_sla_ar = pivots(df_ar)

    df_urbi = df_all[df_all[COL_SOURCE] == "Urbi"].copy()
    if df_urbi.empty:
        p_urbi = pd.DataFrame({"ملاحظة": ["لا توجد بلاغات Urbi"]})
        urbi_index = False
    else:
        p_urbi = (
            df_urbi.groupby([COL_ADMIN, COL_SOURCE])[COL_ID]
            .count()
            .unstack(COL_SOURCE, fill_value=0)
        )
        p_urbi.loc["الإجمالي الكلي"] = p_urbi.sum(axis=0)
        p_urbi.columns.name = None
        urbi_index = True

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        p_open_ar.to_excel(w, sheet_name="١- المفتوحة والمعاد فتحها")
        p_sla_ar.to_excel(w, sheet_name="٢- التوصيف")
        p_urbi.to_excel(w, sheet_name="٣- مصادر أخرى", index=urbi_index)
    out.seek(0)
    return out


def get_pivots_for_ppt(xls: pd.DataFrame, closed_ts: pd.Timestamp):
    df_all = preprocess(xls.copy(), closed_ts)

    df_ar = df_all[df_all[COL_SOURCE] != "Urbi"].copy()
    p_open_ar, p_sla_ar = pivots(df_ar)

    df_urbi = df_all[df_all[COL_SOURCE] == "Urbi"].copy()
    if df_urbi.empty:
        p_urbi = pd.DataFrame()
    else:
        p_urbi = (
            df_urbi.groupby([COL_ADMIN, COL_SOURCE])[COL_ID]
            .count()
            .unstack(COL_SOURCE, fill_value=0)
        )
        p_urbi.loc["الإجمالي الكلي"] = p_urbi.sum(axis=0)
        p_urbi.columns.name = None

    return p_open_ar, p_sla_ar, p_urbi
