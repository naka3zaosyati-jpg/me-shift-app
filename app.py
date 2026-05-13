import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
import calendar
import io
import jpholiday
import random

# --- 初期設定 ---
st.set_page_config(
    page_title="ME勤務表管理",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- カスタムCSS（ライトモード＆カードUI） ---
st.markdown("""
<style>
    .stApp { background-color: #F8F9FA; color: #212529; }
    h1, h2, h3 { color: #0056B3 !important; }
    .card { background-color: #FFFFFF; padding: 1.5rem; border-radius: 12px; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1); border: 1px solid #E9ECEF; margin-bottom: 1.5rem; }
    .stButton>button { background-color: #007BFF; color: #FFFFFF; font-weight: bold; border-radius: 8px; border: none; width: 100%; padding: 0.6rem; transition: all 0.2s; }
    .stButton>button:hover { background-color: #0056B3; color: #FFFFFF; transform: translateY(-2px); }
    th { background-color: #F1F3F5 !important; color: #495057 !important; }
    input, select, textarea { border-radius: 6px !important; }
</style>
""", unsafe_allow_html=True)

# --- 色設定辞書（UI改善） ---
TASK_COLORS = {
    "ＨＭ": {"bg": "#DC3545", "text": "white"},
    "Ｈサ": {"bg": "#DC3545", "text": "white"},
    "Ａ": {"bg": "#198754", "text": "white"},
    "カ": {"bg": "#9ACD32", "text": "black"},
    "Ｉ": {"bg": "#87CEEB", "text": "black"},
    "Ｏ": {"bg": "#00008B", "text": "white"},
    "Ｍ": {"bg": "#8B4513", "text": "white"},
    "Ｄ": {"bg": "#6C757D", "text": "white"},
    "Ｒ": {"bg": "#FFC0CB", "text": "black"},
    "宿直": {"bg": "transparent", "text": "#800080", "fw": "bold"},
    "日勤": {"bg": "#FD7E14", "text": "white"},
    "フリー": {"bg": "#E2E3E5", "text": "black"}, # 余剰スタッフ用
}

HOLIDAY_COLORS = {
    "土": "#EBF5FB",    # 薄い青
    "日祝": "#FDEDEC",  # 薄い赤
    "今日": "#FFF3CD"   # 薄い黄色
}

def get_styled_task_html(text, is_calendar=False):
    """カレンダー用HTMLバッジ生成"""
    if not text: return ""
    items = text.split('\n') if '\n' in text else text.split('<br>')
    res = []
    for item in items:
        item = item.strip()
        if not item: continue
        matched_color = None
        for k, v in TASK_COLORS.items():
            if is_calendar:
                if f"({k})" in item or f" {k} " in item or item.endswith(k) or k in item:
                    matched_color = (k, v)
                    break
            else:
                if k == item or k in item:
                    matched_color = (k, v)
                    break
        if matched_color:
            k, v = matched_color
            bg = v.get("bg", "transparent")
            color = v.get("text", "inherit")
            fw = v.get("fw", "normal")
            if k == "宿直":
                res.append(f'<span style="color: {color}; font-weight: {fw};">{item}</span>')
            else:
                res.append(f'<span style="background-color: {bg}; color: {color}; font-weight: {fw}; padding: 3px 8px; border-radius: 6px; display: inline-block; margin-bottom: 2px; font-size: 0.85rem; box-shadow: 0 1px 2px rgba(0,0,0,0.1); width: fit-content;">{item}</span>')
        else:
            res.append(item)
    return "<br>".join(res)

def style_shift_matrix(df, today_date):
    """
    Pandas Styler 用の関数（axis=None で DataFrame 全体に適用）。
    列ごとの休日カラー（背景）と、セルごとの業務カラーを組み合わせて適用します。
    """
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    for col in df.columns:
        is_sat = "土" in col
        is_sun_or_hol = "日" in col or "祝" in col
        
        # col には "11日\n月曜日\n" のような文字列が入っているため前方一致等で判定
        is_today = str(col).startswith(f"{today_date.day}日\n")
        
        base_bg = "#FFFFFF
