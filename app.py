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
    "フリー": {"bg": "#E2E3E5", "text": "black"},
}

HOLIDAY_COLORS = {
    "土": "#EBF5FB",
    "日祝": "#FDEDEC",
    "今日": "#FFF3CD"
}

def get_styled_task_html(text, is_calendar=False):
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
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    for col in df.columns:
        is_sat = "土" in col
        is_sun_or_hol = "日" in col or "祝" in col
        is_today = str(col).startswith(f"{today_date.day}日\n")
        
        base_bg = "#FFFFFF"
        if is_sun_or_hol: base_bg = HOLIDAY_COLORS["日祝"]
        elif is_sat: base_bg = HOLIDAY_COLORS["土"]
        if is_today: base_bg = HOLIDAY_COLORS["今日"]
            
        for row in df.index:
            val = df.at[row, col]
            matched_bg = base_bg
            matched_color = "inherit"
            matched_fw = "normal"
            
            if isinstance(val, str) and val:
                for k, v in TASK_COLORS.items():
                    if k in val:
                        if k == "宿直":
                            matched_color = v.get("text", "#800080")
                            matched_fw = v.get("fw", "bold")
                        else:
                            matched_bg = v.get("bg", "transparent")
                            if "text" in v: matched_color = v["text"]
            
            styles.at[row, col] = f"background-color: {matched_bg}; color: {matched_color}; font-weight: {matched_fw}; text-align: center; white-space: pre-wrap;"
    return styles

def safe_int(val):
    if pd.isna(val) or val == "": return 0
    try: return int(float(val))
    except (ValueError, TypeError): return 0

def parse_date(date_val):
    if pd.isna(date_val) or str(date_val).strip() == "": return None
    try:
        dt = pd.to_datetime(str(date_val).strip())
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return None

def get_gspread_client():
    try:
        if "gcp_service_account" in st.secrets:
            scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            info = dict(st.secrets["gcp_service_account"])
            credentials = Credentials.from_service_account_info(info, scopes=scopes)
            client = gspread.authorize(credentials)
            return client, None
        else:
            return None, "st.secrets に 'gcp_service_account' が設定されていません。"
    except Exception as e:
        return None, str(e)

@st.cache_data(ttl=60)
def _fetch_records_cached(sheet_name):
    client, _ = get_gspread_client()
    if not client: return None
    try:
        sheet_id = st.secrets.get("spreadsheet_id", "")
        sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
        return sheet.get_all_records()
    except Exception as e:
        st.error(f"データ取得エラー ({sheet_name}): {e}")
        return None

def fetch_data(sheet_name, expected_columns):
    records = _fetch_records_cached(sheet_name)
    if records is not None:
        if records:
            df = pd.DataFrame(records)
            df.dropna(how='all', inplace=True)
            return df
        else:
            return pd.DataFrame(columns=expected_columns)
    else:
        return pd.DataFrame(columns=expected_columns)

def append_data(sheet_name, row_data):
    client, _ = get_gspread_client()
    if not client: return False
    try:
        sheet_id = st.secrets.get("spreadsheet_id", "")
        sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
        res = sheet.append_row(row_data, value_input_option="USER_ENTERED", insert_data_option="INSERT_ROWS", table_range="A1")
        st.cache_data.clear()
        return res
    except Exception: return False

def update_data(sheet_name, search_col_index, search_value, row_data):
    client, _ = get_gspread_client()
    if not client: return False
    try:
        sheet_id = st.secrets.get("spreadsheet_id", "")
        sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
        try:
            cell = sheet.find(str(search_value), in_column=search_col_index)
        except gspread.exceptions.CellNotFound:
            return False
        row_num = cell.row
        end_col_chr = chr(ord('A') + len(row_data) - 1)
        range_str = f"A{row_num}:{end_col_chr}{row_num}"
        try: res = sheet.update(range_name=range_str, values=[row_data], value_input_option="USER_ENTERED")
        except TypeError: res = sheet.update(range_str, [row_data], value_input_option="USER_ENTERED")
        st.cache_data.clear()
        return res
    except Exception: return False

def overwrite_data(sheet_name, df, columns):
    client, _ = get_gspread_client()
    if not client: return False
    try:
        sheet_id = st.secrets.get("spreadsheet_id", "")
        sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
        sheet.clear()
        df = df.fillna("")
        data = [columns] + df[columns].values.tolist()
        try: res = sheet.update(range_name="A1", values=data, value_input_option="USER_ENTERED")
        except TypeError: res = sheet.update("A1", data, value_input_option="USER_ENTERED")
        st.cache_data.clear()
        return res
    except Exception: return False

COLS_STAFF = ["氏名", "役職", "OPE習熟度", "アンギオ習熟度", "総合コード", "人工心肺メイン回数", "人工心肺サブ回数", "アブレーション回数", "カテ回数", "雇用形態"]
COLS_REQUEST = ["日時", "氏名", "区分", "コメント"]
COLS_OPE_MASTER = ["術式名", "術式レベル"]
COLS_OPE_SCHEDULE = ["日時", "術式", "業務"]
COLS_TASK_MASTER = ["略語", "業務名"]
COLS_SHIFT = ["日時", "氏名", "割り当て業務"]
COLS_PROFICIENCY_MASTER = ["区分", "ランク", "定義"]

def render_leave_request_ui():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 🗓️ 希望休入力（月間カレンダー）")
    st.info("💡 対象の年月を選択し、表のセルをクリックして希望（×, △, 年）を入力してください。修正後は必ず「希望休を一括保存」を押してください。")
    
    df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
    if not df_staff.empty: df_staff = df_staff[df_staff["氏名"].astype(str).str.strip() != ""]
    staff_names = df_staff["氏名"].astype(str).str.strip().tolist() if not df_staff.empty else []
    
    today_date = datetime.date.today()
    years_req = list(range(today_date.year - 1, today_date.year + 2))
    months_req = list(range(1, 13))
    
    col_req1, col_req2, _ = st.columns([1, 1, 4])
    with col_req1: req_year = st.selectbox("対象年", years_req, index=years_req.index(today_date.year), key="req_year")
    with col_req2: req_month = st.selectbox("対象月", months_req, index=months_req.index(today_date.month), key="req_month")
    
    _, num_days_req = calendar.monthrange(req_year, req_month)
    cols_req_strings = []
    weekday_list_mapped = ["月", "火", "水", "木", "金", "土", "日"]
    
    for d in range(1, num_days_req + 1):
        dt = datetime.date(req_year, req_month, d)
        wd = dt.weekday()
        is_hol = jpholiday.is_holiday(dt)
        weekday_str = weekday_list_mapped[wd]
        if is_hol: weekday_str += "(祝)"
        cols_req_strings.append(f"{d}日\n{weekday_str}")
        
    req_matrix = pd.DataFrame(index=staff_names, columns=cols_req_strings)
    req_matrix.fillna("", inplace=True)
    
    df_request_all = fetch_data("希望入力", COLS_REQUEST)
    if not df_request_all.empty:
        df_request_all["日時_date"] = pd.to_datetime(df_request_all["日時"], errors="coerce").dt.date
        df_request_all["氏名"] = df_request_all["氏名"].astype(str).str.strip()
        for _, row in df_request_all.dropna(subset=["日時_date"]).iterrows():
            dt = row["日時_date"]
            if dt.year == req_year and dt.month == req_month:
                s_name = row["氏名"]
                kubun = str(row["区分"]).strip()
                if s_name in req_matrix.index: req_matrix.at[s_name, cols_req_strings[dt.day - 1]] = kubun
                
    column_config_req = {}
    for col in cols_req_strings:
        column_config_req[col] = st.column_config.SelectboxColumn(options=["", "×", "△", "年"], default="")
        
    edited_req_matrix = st.data_editor(req_matrix, use_container_width=True, height=600, key="req_month_editor", column_config=column_config_req)
    
    if st.button("💾 希望休を一括保存"):
        with st.spinner("スプレッドシートに保存中..."):
            new_req_list = []
            for staff in edited_req_matrix.index:
                for d in range(1, num_days_req + 1):
                    dt = datetime.date(req_year, req_month, d)
                    d_str = dt.strftime("%Y-%m-%d")
                    header_str = cols_req_strings[d-1]
                    val = str(edited_req_matrix.at[staff, header_str]).strip()
                    if val and val != "nan" and val != "None":
                        new_req_list.append([d_str, staff, val, ""])
                        
            month_prefix = f"{req_year}-{req_month:02d}"
            if not df_request_all.empty:
                df_request_all["_date_obj"] = pd.to_datetime(df_request_all["日時"], errors="coerce").dt.date
                df_request_all["_date_str"] = df_request_all["_date_obj"].apply(lambda x: x.strftime("%Y-%m") if pd.notnull(x) else "")
                df_remain = df_request_all[df_request_all["_date_str"] != month_prefix].drop(columns=["_date_obj", "_date_str", "日時_date"], errors="ignore")
            else: df_remain = pd.DataFrame(columns=COLS_REQUEST)
                
            df_new_req_month = pd.DataFrame(new_req_list, columns=COLS_REQUEST)
            df_final_req = pd.concat([df_remain, df_new_req_month], ignore_index=True)
            res = overwrite_data("希望入力", df_final_req, COLS_REQUEST)
            if res: st.success(f"{req_year}年{req_month}月の希望休を一括保存しました！")
            else: st.error("保存に失敗しました。")
    st.markdown('</div>', unsafe_allow_html=True)

def page_home(is_admin=False):
    title = "① ホーム（勤務表確認・編集）" if is_admin else "① 勤務表の確認（閲覧のみ）"
    st.markdown(f'<div class="card"><h2>{title}</h2><p>確定した勤務表や予定を確認します。</p></div>', unsafe_allow_html=True)
    
    df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
    if not df_staff.empty:
        df_staff = df_staff[df_staff["氏名"].astype(str).str.strip() != ""]
        df_staff = df_staff[~df_staff["氏名"].astype(str).str.lower().isin(["none", "nan"])]
    staff_list = df_staff["氏名"].astype(str).str.strip().tolist() if not df_staff.empty else []
    
    df_shift = fetch_data("確定勤務表", COLS_SHIFT)
    df_ope = fetch_data("術式予定", COLS_OPE_SCHEDULE)
    
    st.markdown('<div class="card">', unsafe_allow_html=True)
    today_date = datetime.date.today()
    years = list(range(today_date.year - 1, today_date.year + 2))
    months = list(range(1, 13))
    col1, col2, _ = st.columns([1, 1, 4])
    with col1: selected_year = st.selectbox("年を選択", years, index=years.index(today_date.year))
    with col2: selected_month = st.selectbox("月を選択", months, index=months.index(today_date.month))
    
    cal = calendar.Calendar(firstweekday=6)
    month_days = cal.monthdatescalendar(selected_year, selected_month)
    
    shift_dict = {}
    if not df_shift.empty:
        df_shift["日時_date"] = pd.to_datetime(df_shift["日時"], errors="coerce").dt.date
        for _, row in df_shift.dropna(subset=["日時_date"]).iterrows():
            d_key = row["日時_date"].strftime("%Y-%m-%d")
            staff_name = str(row['氏名']).strip()
            staff_info = f"{staff_name} ({row['割り当て業務']})"
            if d_key not in shift_dict: shift_dict[d_key] = []
            shift_dict[d_key].append((staff_name, staff_info))
                
    def get_staff_idx(name):
        try: return staff_list.index(name)
        except ValueError: return 999
        
    for k in shift_dict:
        shift_dict[k].sort(key=lambda x: get_staff_idx(x[0]))
        shift_dict[k] = [x[1] for x in shift_dict[k]]

    ope_dict = {}
    if not df_ope.empty:
        df_ope["日時_date"] = pd.to_datetime(df_ope["日時"], errors="coerce").dt.date
        for _, row in df_ope.dropna(subset=["日時_date"]).iterrows():
            d_key = row["日時_date"].strftime("%Y-%m-%d")
            ope_info = str(row['術式'])
            if d_key not in ope_dict: ope_dict[d_key] = []
            ope_dict[d_key].append(ope_info)
                
    calendar_css = """
    <style>
        .calendar-container { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; margin-top: 15px; }
        .calendar-header { text-align: center; font-weight: bold; background-color: #F8F9FA; padding: 10px; border-radius: 8px; color: #495057; border: 1px solid #E9ECEF; }
        .calendar-day { background-color: #FFFFFF; border: 1px solid #E9ECEF; border-radius: 8px; min-height: 140px; padding: 8px; display: flex; flex-direction: column; box-shadow: 0 2px 4px rgba(0,0,0,0.02); transition: transform 0.1s; }
        .calendar-day:hover { transform: scale(1.02); box-shadow: 0 4px 8px rgba(0,0,0,0.05); z-index: 10; }
        .calendar-day.saturday { background-color: #F0F8FF; }
        .calendar-day.holiday { background-color: #FFF0F5; }
        .calendar-day.today { background-color: #FFF3CD !important; border: 2px solid #FFC107 !important; }
        .calendar-day.other-month { opacity: 0.5; background-color: #F8F9FA; }
        .day-number { font-size: 1.3rem; font-weight: 800; color: #0056B3; text-align: left; margin-bottom: 0px; }
        .calendar-day.saturday .day-number { color: #0D6EFD; }
        .calendar-day.holiday .day-number { color: #DC3545; }
        .calendar-day.today .day-number { color: #856404; }
        .day-weekday { font-size: 0.75rem; color: #6C757D; text-align: left; margin-bottom: 8px; border-bottom: 1px solid #E9ECEF; padding-bottom: 4px; font-weight: bold; }
        .day-staff { font-size: 0.85rem; color: #212529; flex-grow: 1; white-space: pre-wrap; line-height: 1.5; }
        .day-ope { font-size: 0.8rem; color: #D63384; font-weight: bold; margin-top: 8px; background-color: #FFF0F6; padding: 4px 6px; border-radius: 4px; white-space: pre-wrap; border-left: 3px solid #D63384; }
    </style>
    """
    st.markdown(calendar_css, unsafe_allow_html=True)
    
    weekdays = ["日", "月", "火", "水", "木", "金", "土"]
    html = '<div class="calendar-container">'
    for wd in weekdays:
        color_style = ""
        if wd == "日": color_style = "color: #DC3545;"
        elif wd == "土": color_style = "color: #0D6EFD;"
        html += f'<div class="calendar-header" style="{color_style}">{wd}</div>'
    
    for week in month_days:
        for d in week:
            is_other_month = d.month != selected_month
            is_saturday = d.weekday() == 5
            is_sunday = d.weekday() == 6
            is_holiday = jpholiday.is_holiday(d)
            is_today = (d == today_date)
            
            class_name = "calendar-day"
            if is_other_month: class_name += " other-month"
            elif is_holiday or is_sunday: class_name += " holiday"
            elif is_saturday: class_name += " saturday"
            if is_today: class_name += " today"
            
            d_str = d.strftime("%Y-%m-%d")
            day_num = d.day
            holiday_name = jpholiday.is_holiday_name(d)
            weekday_list_mapped = ["月", "火", "水", "木", "金", "土", "日"]
            weekday_str = f"{weekday_list_mapped[d.weekday()]}曜日"
            if holiday_name: weekday_str += f" <span style='color:#DC3545;'>({holiday_name})</span>"
                
            staffs = shift_dict.get(d_str, [])
            opes = ope_dict.get(d_str, [])
            
            staff_html_parts = [get_styled_task_html(s, is_calendar=True) for s in staffs]
            staff_html = "<br>".join(staff_html_parts) if staff_html_parts else ""
            ope_html = "<br>".join(opes) if opes else ""
            
            html += f'<div class="{class_name}">'
            html += f'<div class="day-number">{day_num}</div>'
            html += f'<div class="day-weekday">{weekday_str}</div>'
            html += f'<div class="day-staff">{staff_html}</div>'
            if ope_html: html += f'<div class="day-ope">{ope_html}</div>'
            html += '</div>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    if is_admin:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        tab_shift, tab_summary = st.tabs(["📋 横型シフト表 (クリックで編集・保存)", "📊 業務割り当て集計表"])
        
        with tab_shift:
            st.write(f"### 📋 {selected_year}年{selected_month}月 横型シフト表")
            st.info("💡 表の中のセルをクリックすると、**プルダウン形式**で割り当て業務を直接変更できます。修正後は必ず下の「編集内容を保存」ボタンを押してください。")
            
            _, num_days = calendar.monthrange(selected_year, selected_month)
            cols_strings = []
            weekday_list_mapped = ["月", "火", "水", "木", "金", "土", "日"]
            
            for d in range(1, num_days + 1):
                dt = datetime.date(selected_year, selected_month, d)
                d_str = dt.strftime("%Y-%m-%d")
                wd = dt.weekday()
                is_hol = jpholiday.is_holiday(dt)
                weekday_str = weekday_list_mapped[wd]
                if is_hol: weekday_str += "(祝)"
                opes = ope_dict.get(d_str, [])
                ope_str = "\\n".join(opes) if opes else ""
                header_str = f"{d}日\\n{weekday_str}"
                if ope_str: header_str += f"\\n{ope_str}"
                cols_strings.append(header_str)
                
            shift_matrix = pd.DataFrame(index=staff_list, columns=cols_strings)
            shift_matrix.fillna("", inplace=True)
            
            if not df_shift.empty:
                for _, row in df_shift.dropna(subset=["日時_date"]).iterrows():
                    dt = row["日時_date"]
                    if dt.year == selected_year and dt.month == selected_month:
                        staff_name = str(row["氏名"]).strip()
                        duty = str(row["割り当て業務"]).strip()
                        header_col = cols_strings[dt.day - 1]
                        if staff_name in shift_matrix.index or staff_name.startswith("スタッフ"):
                            if staff_name not in shift_matrix.index:
                                shift_matrix.loc[staff_name] = ""
                            curr_val = shift_matrix.at[staff_name, header_col]
                            if curr_val: shift_matrix.at[staff_name, header_col] = curr_val + "\\n" + duty
                            else: shift_matrix.at[staff_name, header_col] = duty

            base_options = ["", "フリー", "ＨＭ", "Ｈサ", "Ａ", "カ", "Ｉ", "Ｏ", "Ｍ", "Ｄ", "Ｒ", "日勤", "宿直"]
            combos = [o + "\\n宿直" for o in base_options if o and o not in ["宿直", "日勤"]]
            combos.append("宿直\\n日勤")
            all_options = base_options + combos
            
            column_config = {}
            for col in cols_strings:
                column_config[col] = st.column_config.SelectboxColumn(options=all_options, default="")

            styled_matrix = shift_matrix.style.apply(lambda df: style_shift_matrix(df, today_date), axis=None)
                
            edited_matrix = st.data_editor(styled_matrix, use_container_width=True, height=600, key="month_shift_editor", column_config=column_config)
            
            if st.button("💾 編集内容をスプレッドシートに保存"):
                with st.spinner("スプレッドシートに保存中..."):
                    new_draft = []
                    for staff in edited_matrix.index:
                        for d in range(1, num_days + 1):
                            dt = datetime.date(selected_year, selected_month, d)
                            d_str = dt.strftime("%Y-%m-%d")
                            header_str = cols_strings[d-1]
                            val = str(edited_matrix.at[staff, header_str]).strip()
                            if val and val != "nan" and val != "None":
                                tasks = val.split('\\n')
                                for t in tasks:
                                    t = t.strip()
                                    if t: new_draft.append([d_str, staff, t])
                                    
                    if new_draft:
                        df_shift_all = fetch_data("確定勤務表", COLS_SHIFT)
                        month_prefix = f"{selected_year}-{selected_month:02d}"
                        if not df_shift_all.empty:
                            df_shift_all["_date_obj"] = pd.to_datetime(df_shift_all["日時"], errors="coerce").dt.date
                            df_shift_all["_date_str"] = df_shift_all["_date_obj"].apply(lambda x: x.strftime("%Y-%m") if pd.notnull(x) else "")
                            df_remain = df_shift_all[df_shift_all["_date_str"] != month_prefix].drop(columns=["_date_obj", "_date_str"], errors="ignore")
                        else: df_remain = pd.DataFrame(columns=COLS_SHIFT)
                            
                        df_new_month = pd.DataFrame(new_draft, columns=COLS_SHIFT)
                        df_final = pd.concat([df_remain, df_new_month], ignore_index=True)
                        res = overwrite_data("確定勤務表", df_final, COLS_SHIFT)
                        if res: st.success(f"{selected_year}年{selected_month}月の編集内容を一括保存しました！")
                    else: st.warning("保存するデータがありません。")
                        
            try:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer: shift_matrix.to_excel(writer, sheet_name=f"{selected_month}月シフト")
                st.download_button(label="📥 Excel形式でダウンロード (.xlsx)", data=buffer.getvalue(), file_name=f"shift_{selected_year}_{selected_month}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception:
                csv = shift_matrix.to_csv(encoding="utf-8-sig")
                st.download_button(label="📥 CSV形式でダウンロード (.csv)", data=csv, file_name=f"shift_{selected_year}_{selected_month}.csv", mime="text/csv")

        with tab_summary:
            st.write(f"### 📊 {selected_year}年{selected_month}月 業務割り当て集計表")
            if not df_shift.empty:
                df_month = df_shift.copy()
                df_month["_date_obj"] = pd.to_datetime(df_month["日時"], errors="coerce").dt.date
                df_month["_date_str"] = df_month["_date_obj"].apply(lambda x: x.strftime("%Y-%m") if pd.notnull(x) else "")
                month_prefix = f"{selected_year}-{selected_month:02d}"
                df_month = df_month[df_month["_date_str"] == month_prefix]
                if not df_month.empty:
                    summary_df = pd.pivot_table(df_month, index="氏名", columns="割り当て業務", aggfunc="size", fill_value=0)
                    summary_df["合計"] = summary_df.sum(axis=1)
                    st.dataframe(summary_df, use_container_width=True)
                else: st.info("勤務表データがありません。")
            else: st.info("データがありません。")
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write(f"### 📊 {selected_year}年{selected_month}月 業務割り当て集計表")
        if not df_shift.empty:
            df_month = df_shift.copy()
            df_month["_date_obj"] = pd.to_datetime(df_month["日時"], errors="coerce").dt.date
            df_month["_date_str"] = df_month["_date_obj"].apply(lambda x: x.strftime("%Y-%m") if pd.notnull(x) else "")
            month_prefix = f"{selected_year}-{selected_month:02d}"
            df_month = df_month[df_month["_date_str"] == month_prefix]
            if not df_month.empty:
                summary_df = pd.pivot_table(df_month, index="氏名", columns="割り当て業務", aggfunc="size", fill_value=0)
                summary_df["合計"] = summary_df.sum(axis=1)
                st.dataframe(summary_df, use_container_width=True)
            else: st.info("勤務表データがありません。")
        else: st.info("データがありません。")
        st.markdown('</div>', unsafe_allow_html=True)

def page_schedule_task():
    st.markdown('<div class="card"><h2>② 予定・希望休・マスタ管理</h2><p>術式予定、希望休入力、各種マスタの管理を行います。</p></div>', unsafe_allow_html=True)
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["術式予定", "希望休入力", "業務マスタ管理", "術式マスタ管理", "習熟度マスタ確認"])
    
    with tab1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 術式予定の登録")
        with st.form("schedule_form", clear_on_submit=True):
            sched_date = st.date_input("手術日時")
            df_ope_master = fetch_data("術式マスタ", COLS_OPE_MASTER)
            ope_names = df_ope_master["術式名"].tolist() if not df_ope_master.empty else ["CABG", "AVR", "PCI", "アブレーション"]
            sched_ope = st.selectbox("術式", ope_names)
            sched_gyomu = st.selectbox("業務", ["なし", "心外", "アブレーション"])
            if st.form_submit_button("予定追加"):
                res = append_data("術式予定", [str(sched_date), sched_ope, sched_gyomu])
                if res: st.success("スプレッドシートに予定を書き込みました。")
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 登録済み予定一覧")
        st.dataframe(fetch_data("術式予定", COLS_OPE_SCHEDULE), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with tab2:
        render_leave_request_ui()

    with tab3:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        m_tab1, m_tab2 = st.tabs(["新規追加", "情報更新"])
        with m_tab1:
            st.write("### 業務マスタ登録")
            with st.form("task_form", clear_on_submit=True):
                task_abbr = st.text_input("略語 (例: HM, A)")
                task_name = st.text_input("業務名 (例: 血液浄化, 血管造影)")
                if st.form_submit_button("マスタ追加"):
                    res = append_data("業務マスタ", [task_abbr, task_name])
                    if res: st.success("スプレッドシートに業務マスタを書き込みました。")
        with m_tab2:
            st.write("### 業務マスタ更新")
            df_task = fetch_data("業務マスタ", COLS_TASK_MASTER)
            if not df_task.empty:
                abbrs = df_task["略語"].astype(str).tolist()
                selected_abbr = st.selectbox("更新する略語を選択", abbrs)
                curr_task = df_task[df_task["略語"].astype(str) == selected_abbr].iloc[0]
                with st.form("task_update_form"):
                    task_abbr_upd = st.text_input("略語（検索キーのため変更不可）", value=curr_task["略語"], disabled=True)
                    task_name_upd = st.text_input("業務名", value=curr_task["業務名"])
                    if st.form_submit_button("情報を更新"):
                        res = update_data("業務マスタ", 1, task_abbr_upd, [task_abbr_upd, task_name_upd])
                        if res: st.success(f"業務マスタ「{task_abbr_upd}」を更新しました。")
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 業務マスタ一覧")
        st.dataframe(fetch_data("業務マスタ", COLS_TASK_MASTER), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with tab4:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        om_tab1, om_tab2 = st.tabs(["新規登録", "情報更新"])
        with om_tab1:
            st.write("### 術式マスタ登録")
            with st.form("ope_master_form", clear_on_submit=True):
                ope_name = st.text_input("術式名")
                ope_level = st.selectbox("術式レベル", ["A", "B", "C", "D"])
                if st.form_submit_button("マスタ追加"):
                    res = append_data("術式マスタ", [ope_name, ope_level])
                    if res: st.success("スプレッドシートに術式マスタを追加しました。")
        with om_tab2:
            st.write("### 術式マスタ更新")
            df_ope_m = fetch_data("術式マスタ", COLS_OPE_MASTER)
            if not df_ope_m.empty:
                o_names = df_ope_m["術式名"].astype(str).tolist()
                sel_ope = st.selectbox("更新する術式を選択", o_names)
                curr_ope = df_ope_m[df_ope_m["術式名"].astype(str) == sel_ope].iloc[0]
                with st.form("ope_master_update_form"):
                    ope_name_upd = st.text_input("術式名（変更不可）", value=curr_ope["術式名"], disabled=True)
                    ol_idx = ["A", "B", "C", "D"].index(curr_ope["術式レベル"]) if curr_ope["術式レベル"] in ["A", "B", "C", "D"] else 0
                    ope_level_upd = st.selectbox("術式レベル", ["A", "B", "C", "D"], index=ol_idx)
                    if st.form_submit_button("情報を更新"):
                        res = update_data("術式マスタ", 1, ope_name_upd, [ope_name_upd, ope_level_upd])
                        if res: st.success(f"「{ope_name_upd}」の情報を更新しました。")
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 術式マスタ一覧")
        st.dataframe(fetch_data("術式マスタ", COLS_OPE_MASTER), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
    with tab5:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 習熟度マスタ一覧")
        st.info("※ このマスタの追加・編集はスプレッドシートから直接行ってください。")
        st.dataframe(fetch_data("習熟度マスタ", COLS_PROFICIENCY_MASTER), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

def page_staff():
    st.markdown('<div class="card"><h2>③ スタッフマスタ管理</h2><p>スタッフの基本情報や各分野の習熟度を管理します。</p></div>', unsafe_allow_html=True)
    
    df_prof = fetch_data("習熟度マスタ", COLS_PROFICIENCY_MASTER)
    ope_options = ["A", "B", "C", "D"]
    angio_options = ["1", "2", "3", "4"]
    ope_defs, angio_defs = "", ""
    
    if not df_prof.empty:
        df_ope_prof = df_prof[df_prof["区分"] == "OPE"]
        if not df_ope_prof.empty:
            ope_options = df_ope_prof["ランク"].astype(str).tolist()
            ope_defs = "\\n".join([f"- **{r['ランク']}**: {r['定義']}" for _, r in df_ope_prof.iterrows()])
            
        df_angio_prof = df_prof[df_prof["区分"] == "アンギオ"]
        if not df_angio_prof.empty:
            angio_options = df_angio_prof["ランク"].astype(str).tolist()
            angio_defs = "\\n".join([f"- **{r['ランク']}**: {r['定義']}" for _, r in df_angio_prof.iterrows()])

    st.markdown('<div class="card">', unsafe_allow_html=True)
    tab1, tab2 = st.tabs(["新規登録", "情報更新"])
    
    with tab1:
        st.write("### 新規スタッフ登録")
        with st.form("staff_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                staff_name = st.text_input("氏名")
                staff_role = st.selectbox("役職", ["技士長", "副技士長", "主任", "一般", "新人"])
                employment_type = st.selectbox("雇用形態", ["常勤", "非常勤"])
            with col2:
                ope_level = st.selectbox("OPE習熟度", ope_options)
                if ope_defs: st.caption("【OPE習熟度の定義】\\n" + ope_defs)
                angio_level = st.selectbox("アンギオ習熟度", angio_options)
                if angio_defs: st.caption("【アンギオ習熟度の定義】\\n" + angio_defs)
                
            st.markdown("---")
            st.write("#### 経験回数・実績")
            col3, col4 = st.columns(2)
            with col3:
                cpb_main = st.number_input("人工心肺 メイン回数", min_value=0, step=1)
                cpb_sub = st.number_input("人工心肺 サブ回数", min_value=0, step=1)
            with col4:
                ablation_count = st.number_input("アブレーション回数", min_value=0, step=1)
                catha_count = st.number_input("カテ回数", min_value=0, step=1)
                
            if st.form_submit_button("スタッフ登録"):
                total_code = f"{ope_level}-{angio_level}"
                row = [staff_name, staff_role, ope_level, angio_level, total_code, int(cpb_main), int(cpb_sub), int(ablation_count), int(catha_count), employment_type]
                res = append_data("スタッフマスタ", row)
                if res: st.success(f"スプレッドシートに「{staff_name}」さんのデータを新規登録しました。（総合コード: {total_code}）")
                    
    with tab2:
        st.write("### 登録済みスタッフの情報更新")
        df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
        if not df_staff.empty:
            staff_names = df_staff["氏名"].astype(str).tolist()
            selected_staff = st.selectbox("更新するスタッフを選択", staff_names)
            curr = df_staff[df_staff["氏名"].astype(str) == selected_staff].iloc[0]
            
            with st.form("staff_update_form"):
                col1, col2 = st.columns(2)
                with col1:
                    staff_name_upd = st.text_input("氏名（検索キーのため変更不可）", value=curr["氏名"], disabled=True)
                    role_options = ["技士長", "副技士長", "主任", "一般", "新人"]
                    r_idx = role_options.index(curr["役職"]) if curr["役職"] in role_options else 3
                    staff_role_upd = st.selectbox("役職", role_options, index=r_idx)
                    emp_options = ["常勤", "非常勤"]
                    curr_emp = curr.get("雇用形態", "常勤")
                    e_idx = emp_options.index(curr_emp) if curr_emp in emp_options else 0
                    employment_type_upd = st.selectbox("雇用形態", emp_options, index=e_idx)
                with col2:
                    o_idx = ope_options.index(curr["OPE習熟度"]) if curr["OPE習熟度"] in ope_options else 0
                    ope_level_upd = st.selectbox("OPE習熟度", ope_options, index=o_idx)
                    a_idx = angio_options.index(str(curr["アンギオ習熟度"])) if str(curr["アンギオ習熟度"]) in angio_options else 0
                    angio_level_upd = st.selectbox("アンギオ習熟度", angio_options, index=a_idx)
                
                st.markdown("---")
                st.write("#### 経験回数・実績")
                col3, col4 = st.columns(2)
                with col3:
                    cpb_main_upd = st.number_input("人工心肺 メイン回数", min_value=0, step=1, value=safe_int(curr["人工心肺メイン回数"]))
                    cpb_sub_upd = st.number_input("人工心肺 サブ回数", min_value=0, step=1, value=safe_int(curr["人工心肺サブ回数"]))
                with col4:
                    ablation_upd = st.number_input("アブレーション回数", min_value=0, step=1, value=safe_int(curr["アブレーション回数"]))
                    catha_upd = st.number_input("カテ回数", min_value=0, step=1, value=safe_int(curr["カテ回数"]))
                    
                if st.form_submit_button("情報を更新"):
                    total_code_upd = f"{ope_level_upd}-{angio_level_upd}"
                    row_upd = [staff_name_upd, staff_role_upd, ope_level_upd, total_code_upd, int(cpb_main_upd), int(cpb_sub_upd), int(ablation_upd), int(catha_upd), employment_type_upd]
                    res = update_data("スタッフマスタ", 1, staff_name_upd, row_upd)
                    if res: st.success(f"「{staff_name_upd}」さんの情報を上書き更新しました。（新しい総合コード: {total_code_upd}）")
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### スタッフ一覧")
    st.dataframe(fetch_data("スタッフマスタ", COLS_STAFF), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

def page_shift_creation():
    st.markdown('<div class="card"><h2>④ 勤務表作成（自動生成）</h2><p>勤務表の「1ヶ月一括作成」を行います。</p></div>', unsafe_allow_html=True)
    
    with st.expander("現在の自動割り当てアルゴリズム（条件）"):
        st.markdown("""
- **優先順位1 (休み)**: 「×」「年」のスタッフを候補から最優先で完全除外。
- **優先順位2 (宿直・1名体制確定)**: 平日・休日問わず、「△」の人を最優先で1名のみ割り当て（確定したら他は選ばない）。
- **優先順位3 (非常勤)**: 非常勤スタッフを「Ｍ」に配置。
- **優先順位4 (アブレーション)**: 2名配置（カテ兼務。アンギオ3以上）。
- **優先順位5 (心外 HM/Hサ)**: HM設定回数が未達の人を優先しHMを決定、後からHサを補充（OPE C以上。D以上1名、C含むなら3名体制。HM 0回設定の人は除外）。
- **優先順位6 (カテ)**: A枠がない平日のみ1名配置（アンギオ2以上）。
- **優先順位7 (残りの宿直)**: ステップ2で決まらなかった場合のみ、回数が少ない人を1名選出（複数選出禁止）。
- **優先順位8 (基本業務 I,O,M,D,R)**: 毎日処理順序をシャッフルし、前日と同じ業務の連続割り当ては「候補除外（ハード制約）」により完全に防ぐ。
        """)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 曜日別枠設定")
    df_sys = fetch_data("システム設定", ["項目", "月", "火", "水", "木", "金"])
    
    default_settings = {
        "心外": {"月": False, "火": False, "水": False, "木": False, "金": False},
        "アブレーション": {"月": False, "火": False, "水": False, "木": False, "金": False}
    }
    
    if not df_sys.empty:
        for _, row in df_sys.iterrows():
            item = str(row.get("項目", "")).strip()
            if item in default_settings:
                for d in ["月", "火", "水", "木", "金"]:
                    val = str(row.get(d, "")).lower()
                    default_settings[item][d] = val in ['true', '1', 'yes', 'on', 't']

    with st.form("sys_settings_form"):
        st.write("#### 心外枠（ＨＭ・Ｈサ）")
        cols_heart = st.columns(5)
        new_heart = {}
        for i, d in enumerate(["月", "火", "水", "木", "金"]):
            with cols_heart[i]: new_heart[d] = st.checkbox(d, value=default_settings["心外"][d], key=f"heart_{d}")
                
        st.write("#### アブレーション枠（Ａ）")
        cols_ab = st.columns(5)
        new_ab = {}
        for i, d in enumerate(["月", "火", "水", "木", "金"]):
            with cols_ab[i]: new_ab[d] = st.checkbox(d, value=default_settings["アブレーション"][d], key=f"ab_{d}")
                
        if st.form_submit_button("設定を保存"):
            df_new_sys = pd.DataFrame([
                ["心外", new_heart["月"], new_heart["火"], new_heart["水"], new_heart["木"], new_heart["金"]],
                ["アブレーション", new_ab["月"], new_ab["火"], new_ab["水"], new_ab["木"], new_ab["金"]]
            ], columns=["項目", "月", "火", "水", "木", "金"])
            res = overwrite_data("システム設定", df_new_sys, ["項目", "月", "火", "水", "木", "金"])
            if res:
                st.success("システム設定を保存しました。")
                default_settings["心外"] = new_heart
                default_settings["アブレーション"] = new_ab
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 機能A：1ヶ月一括ドラフト作成（ベース作り）")
    st.info("対象月のデータがすべて「洗い替え（上書き）」されます。")
    
    col1, col2 = st.columns(2)
    today = datetime.date.today()
    with col1: target_year = st.number_input("対象年", min_value=2020, max_value=2050, value=today.year, step=1)
    with col2: target_month = st.number_input("対象月", min_value=1, max_value=12, value=today.month, step=1)
    
    st.write("#### ⚕️ 今月の心外メイン（ＨＭ）最低割り当て回数設定")
    st.info("※OPE習熟度「C」以上のスタッフのみ表示されます。設定回数に未到達のスタッフが優先してHMに割り当てられます。**「0」を指定したスタッフはHM候補から完全に除外されます。**")
    
    df_staff_hm = fetch_data("スタッフマスタ", COLS_STAFF)
    hm_cands_init = []
    if not df_staff_hm.empty:
        for _, row in df_staff_hm.iterrows():
            sname = str(row["氏名"]).strip()
            ope_rank = str(row.get("OPE習熟度", "A")).strip()
            if sname and ope_rank >= 'C': hm_cands_init.append(sname)
                
    hm_min_df = pd.DataFrame(index=hm_cands_init, columns=["最低回数"])
    hm_min_df["最低回数"] = 0
    edited_hm_min = st.data_editor(
        hm_min_df, 
        column_config={"最低回数": st.column_config.NumberColumn("最低回数", min_value=0, max_value=31, step=1)},
        use_container_width=True, 
        key="hm_min_editor"
    )
    
    def is_off_day_func(dt):
        return dt.weekday() >= 5 or jpholiday.is_holiday(dt) or (dt.month == 12 and dt.day >= 29) or (dt.month == 1 and dt.day <= 3)

    # --- UI表示用の年間・月間集計表 ---
    df_staff_ui = fetch_data("スタッフマスタ", COLS_STAFF)
    staff_list_summary = []
    if not df_staff_ui.empty:
        df_staff_clean = df_staff_ui[df_staff_ui["氏名"].astype(str).str.strip() != ""]
        staff_list_summary = df_staff_clean[~df_staff_clean["氏名"].astype(str).str.lower().isin(["none", "nan"])]["氏名"].astype(str).str.strip().tolist()

    fiscal_year_ui = target_year if target_month >= 4 else target_year - 1
    fy_start_ui = datetime.date(fiscal_year_ui, 4, 1)
    
    annual_summary = {s: {"心外(HM/Hサ)": 0, "A": 0, "カ": 0, "I": 0, "O": 0, "M": 0, "D": 0, "R": 0, "宿直": 0} for s in staff_list_summary}
    monthly_summary = {s: {"当月宿直(合計)": 0, "当月休日出勤": 0} for s in staff_list_summary}
    
    df_history_ui = fetch_data("確定勤務表", COLS_SHIFT)
    if not df_history_ui.empty:
        df_history_ui["日時_date"] = pd.to_datetime(df_history_ui["日時"], errors="coerce").dt.date
        for _, row in df_history_ui.dropna(subset=["日時_date"]).iterrows():
            h_date = row["日時_date"]
            s_name = str(row["氏名"]).strip()
            t_name = str(row["割り当て業務"]).strip()
            if s_name not in staff_list_summary: continue
            
            if h_date >= fy_start_ui and (h_date.year < target_year or (h_date.year == target_year and h_date.month <= target_month)):
                if t_name in ["ＨＭ", "Ｈサ"]: annual_summary[s_name]["心外(HM/Hサ)"] += 1
                elif t_name == "Ａ": annual_summary[s_name]["A"] += 1
                elif t_name == "カ": annual_summary[s_name]["カ"] += 1
                elif t_name == "Ｉ": annual_summary[s_name]["I"] += 1
                elif t_name == "Ｏ": annual_summary[s_name]["O"] += 1
                elif t_name == "Ｍ": annual_summary[s_name]["M"] += 1
                elif t_name == "Ｄ": annual_summary[s_name]["D"] += 1
                elif t_name == "Ｒ": annual_summary[s_name]["R"] += 1
                elif t_name == "宿直": annual_summary[s_name]["宿直"] += 1
                
            if h_date.year == target_year and h_date.month == target_month:
                if t_name == "宿直": monthly_summary[s_name]["当月宿直(合計)"] += 1
                if t_name == "日勤" and is_off_day_func(h_date):
                    monthly_summary[s_name]["当月休日出勤"] += 1

    if staff_list_summary:
        summary_data = []
        for s in staff_list_summary:
            row = {"氏名": s}
            row.update(annual_summary[s])
            row.update(monthly_summary[s])
            summary_data.append(row)
        st.write(f"#### 📊 {fiscal_year_ui}年度 累計実績 ＆ {target_year}年{target_month}月 実績サマリー")
        st.dataframe(pd.DataFrame(summary_data), use_container_width=True)

    with st.form("bulk_draft_form"):
        if st.form_submit_button("1ヶ月分の一括ドラフトを作成"):
            with st.spinner(f"{target_year}年{target_month}月のドラフトを作成中..."):
                df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
                if not df_staff.empty:
                    df_staff["氏名"] = df_staff["氏名"].astype(str).str.replace(r'[\s　]+', '', regex=True)
                    df_staff = df_staff[df_staff["氏名"] != ""]
                    df_staff = df_staff[~df_staff["氏名"].str.lower().isin(["none", "nan"])]
                
                df_request = fetch_data("希望入力", COLS_REQUEST)
                staff_list = df_staff["氏名"].tolist() if not df_staff.empty else []
                
                part_time_staff = []
                staff_ope_dict = {}
                staff_angio_dict = {}
                if "雇用形態" in df_staff.columns:
                    df_staff["雇用形態"] = df_staff["雇用形態"].astype(str).str.replace(r'[\s　]+', '', regex=True)
                    part_time_staff = df_staff[df_staff["雇用形態"] == "非常勤"]["氏名"].tolist()
                
                if not df_staff.empty:
                    for _, row in df_staff.iterrows():
                        sname = row["氏名"]
                        staff_ope_dict[sname] = str(row["OPE習熟度"]).strip()
                        staff_angio_dict[sname] = str(row["アンギオ習熟度"]).strip()
                
                # 【1. 日付照合の絶対確実化】希望データの取得とデータ型の厳密な照合
                req_unavailable = {}
                req_night_shift = {}
                if not df_request.empty:
                    df_request["日時_date"] = pd.to_datetime(df_request["日時"], errors="coerce").dt.date
                    df_request["氏名"] = df_request["氏名"].astype(str).str.replace(r'[\s　]+', '', regex=True)
                    df_request["区分"] = df_request["区分"].astype(str).str.replace(r'[\s　]+', '', regex=True)
                    for _, row in df_request.dropna(subset=["日時_date"]).iterrows():
                        r_date_str = row["日時_date"].strftime("%Y-%m-%d")
                        kubun = row["区分"]
                        staff_n = row["氏名"]
                        
                        if "×" in kubun or "年" in kubun:
                            if r_date_str not in req_unavailable: req_unavailable[r_date_str] = []
                            req_unavailable[r_date_str].append(staff_n)
                        elif "△" in kubun:
                            if r_date_str not in req_night_shift: req_night_shift[r_date_str] = []
                            req_night_shift[r_date_str].append(staff_n)
                                
                _, num_days = calendar.monthrange(target_year, target_month)
                draft_data = []
                
                # ーーー 月間の「△」を事前確保（ロック） ーーー
                pre_assigned_shukuchoku = {s: set() for s in staff_list}
                for d in range(1, num_days + 1):
                    dt = datetime.date(target_year, target_month, d)
                    d_str = dt.strftime("%Y-%m-%d")
                    abs_day = (dt - datetime.date(2020, 1, 1)).days
                    wishers = [s for s in req_night_shift.get(d_str, []) if s in staff_list and s not in part_time_staff]
                    for w in wishers:
                        pre_assigned_shukuchoku[w].add(abs_day)
                
                # --- 年間を通じた回数均等化の準備 ---
                fiscal_year = target_year if target_month >= 4 else target_year - 1
                fy_start = datetime.date(fiscal_year, 4, 1)
                
                annual_task_counts = {s: {} for s in staff_list}
                monthly_d_count = {s: 0 for s in staff_list}
                staff_night_holiday_abs_days = {s: set() for s in staff_list}
                
                df_history = fetch_data("確定勤務表", COLS_SHIFT)
                if not df_history.empty:
                    df_history["日時_date"] = pd.to_datetime(df_history["日時"], errors="coerce").dt.date
                    df_history["氏名"] = df_history["氏名"].astype(str).str.replace(r'[\s　]+', '', regex=True)
                    for _, row in df_history.dropna(subset=["日時_date"]).iterrows():
                        h_date = row["日時_date"]
                        s_name = row["氏名"]
                        t_name = str(row["割り当て業務"]).strip()
                        if s_name not in staff_list: continue
                        
                        abs_day = (h_date - datetime.date(2020, 1, 1)).days
                        if t_name in ["宿直", "日勤"]: staff_night_holiday_abs_days[s_name].add(abs_day)
                            
                        if h_date >= fy_start and (h_date.year < target_year or (h_date.year == target_year and h_date.month < target_month)):
                            annual_task_counts[s_name][t_name] = annual_task_counts[s_name].get(t_name, 0) + 1
                            if t_name == "日勤" and is_off_day_func(h_date):
                                annual_task_counts[s_name]["休日出勤"] = annual_task_counts[s_name].get("休日出勤", 0) + 1
                                
                        if h_date.year == target_year and h_date.month == target_month and t_name == "Ｄ":
                            monthly_d_count[s_name] += 1

                monthly_night_count = {s: 0 for s in staff_list}
                monthly_holiday_night_count = {s: 0 for s in staff_list}
                hm_min_counts = {k: safe_int(v) for k, v in edited_hm_min["最低回数"].to_dict().items()}
                monthly_hm_count = {s: 0 for s in staff_list}

                def get_annual_count(staff, task):
                    if task == "心外": return annual_task_counts.get(staff, {}).get("ＨＭ", 0) + annual_task_counts.get(staff, {}).get("Ｈサ", 0)
                    return annual_task_counts.get(staff, {}).get(task, 0)
                    
                def increment_annual_count(staff, task):
                    if staff not in annual_task_counts: annual_task_counts[staff] = {}
                    annual_task_counts[staff][task] = annual_task_counts[staff].get(task, 0) + 1
                    
                def can_do_night_or_holiday(staff, current_abs_day, is_random_assignment=False):
                    for d in staff_night_holiday_abs_days[staff]:
                        if abs(d - current_abs_day) <= 3: return False
                    if is_random_assignment:
                        for d in pre_assigned_shukuchoku[staff]:
                            if d != current_abs_day and abs(d - current_abs_day) <= 3: return False
                    return True

                holiday_block_map = {}
                current_block = []
                for d in range(1, num_days + 1):
                    dt = datetime.date(target_year, target_month, d)
                    if is_off_day_func(dt): current_block.append(d)
                    else:
                        if current_block:
                            for day in current_block: holiday_block_map[day] = current_block
                            current_block = []
                if current_block:
                    for day in current_block: holiday_block_map[day] = current_block
                    
                staff_assigned_holiday_days = {s: set() for s in staff_list}
                def has_worked_in_block(staff, day):
                    block = holiday_block_map.get(day, [])
                    for b_day in block:
                        if b_day in staff_assigned_holiday_days[staff]: return True
                    return False

                df_ope = fetch_data("術式予定", COLS_OPE_SCHEDULE)
                ope_gyomu_dict = {}
                if not df_ope.empty:
                    df_ope["日時_date"] = pd.to_datetime(df_ope["日時"], errors="coerce").dt.date
                    for _, row in df_ope.dropna(subset=["日時_date"]).iterrows():
                        d_key = row["日時_date"].strftime("%Y-%m-%d")
                        gyomu = str(row.get('業務', '')).strip()
                        if d_key not in ope_gyomu_dict: ope_gyomu_dict[d_key] = []
                        if gyomu in ["心外", "アブレーション"]: ope_gyomu_dict[d_key].append(gyomu)

                d_task_assigned_this_month = False
                yesterday_tasks = {}
                
                for d in range(1, num_days + 1):
                    dt = datetime.date(target_year, target_month, d)
                    d_str = dt.strftime("%Y-%m-%d")
                    current_abs_day = (dt - datetime.date(2020, 1, 1)).days
                    is_off_day = is_off_day_func(dt)
                    
                    # 【2. 1日の割り当て処理の完全分離構造】
                    # ① 休みの絶対除外（文字列マッチングで確実に）
                    unavailable = req_unavailable.get(d_str, [])
                    can_work = [s for s in staff_list if s not in unavailable]
                    daily_tasks = {s: [] for s in can_work} # リストでタスクを管理

                    # 必要な業務の設定
                    opes_gyomu = ope_gyomu_dict.get(d_str, [])
                    heart_count = opes_gyomu.count("心外")
                    ablation_count = opes_gyomu.count("アブレーション")
                    
                    c_needed = 0
                    is_ablation_day = False
                    
                    if not is_off_day:
                        week_cols = ["月", "火", "水", "木", "金"]
                        if 0 <= dt.weekday() <= 4:
                            day_str = week_cols[dt.weekday()]
                            if default_settings["心外"][day_str]: heart_count += 1
                            if default_settings["アブレーション"][day_str]: ablation_count += 1
                        
                        is_ablation_day = ablation_count > 0
                        if is_ablation_day: c_needed = 0
                        else: c_needed = 1

                    dummy_counter = 0

                    # 宿直の候補 (休みと非常勤以外)
                    night_candidates = [s for s in can_work if s not in part_time_staff]
                    
                    # 「△」希望者
                    wishers = [s for s in req_night_shift.get(d_str, []) if s in night_candidates]
                    valid_wishers = [s for s in wishers if can_do_night_or_holiday(s, current_abs_day, is_random_assignment=False)]

                    # ② 休日の場合
                    if is_off_day:
                        # 休日は非常勤を除外
                        can_work_holiday = [s for s in night_candidates]
                        
                        night_assigned = False
                        valid_wishers.sort(key=lambda s: (monthly_holiday_night_count.get(s, 0), get_annual_count(s, "休日出勤"), get_annual_count(s, "宿直")))
                        
                        for wisher in valid_wishers:
                            night_staff = wisher
                            daily_tasks[night_staff].append("宿直\n日勤")
                            increment_annual_count(night_staff, "宿直")
                            increment_annual_count(night_staff, "日勤")
                            increment_annual_count(night_staff, "休日出勤")
                            staff_night_holiday_abs_days[night_staff].add(current_abs_day)
                            staff_assigned_holiday_days[night_staff].add(d)
                            monthly_night_count[night_staff] += 1
                            monthly_holiday_night_count[night_staff] += 1
                            night_assigned = True
                            break # 必ず1名で抜ける
                            
                        if not night_assigned:
                            valid_night = [s for s in can_work_holiday if can_do_night_or_holiday(s, current_abs_day, is_random_assignment=True) and monthly_night_count.get(s, 0) < 4]
                            
                            valid_night_holiday = [s for s in valid_night if monthly_holiday_night_count.get(s, 0) < 1 and not has_worked_in_block(s, d)]
                            if not valid_night_holiday:
                                valid_night_holiday = [s for s in can_work_holiday if can_do_night_or_holiday(s, current_abs_day, is_random_assignment=True) and monthly_night_count.get(s, 0) < 4 and monthly_holiday_night_count.get(s, 0) < 2 and not has_worked_in_block(s, d)]
                            if not valid_night_holiday:
                                valid_night_holiday = [s for s in can_work_holiday if can_do_night_or_holiday(s, current_abs_day, is_random_assignment=True) and monthly_night_count.get(s, 0) < 4 and not has_worked_in_block(s, d)]
                            valid_night = valid_night_holiday
                            
                            if not valid_night:
                                valid_night = [s for s in can_work_holiday if monthly_night_count.get(s, 0) < 4]
                            if not valid_night:
                                valid_night = can_work_holiday
                                
                            valid_night.sort(key=lambda s: (monthly_holiday_night_count.get(s, 0), get_annual_count(s, "休日出勤"), get_annual_count(s, "宿直")))
                            
                            for cand in valid_night:
                                night_staff = cand
                                daily_tasks[night_staff].append("宿直\n日勤")
                                increment_annual_count(night_staff, "宿直")
                                increment_annual_count(night_staff, "日勤")
                                increment_annual_count(night_staff, "休日出勤")
                                staff_night_holiday_abs_days[night_staff].add(current_abs_day)
                                staff_assigned_holiday_days[night_staff].add(d)
                                monthly_night_count[night_staff] += 1
                                monthly_holiday_night_count[night_staff] += 1
                                night_assigned = True
                                break # 必ず1名で抜ける

                    # ③ 平日の場合
                    else:
                        available_for_tasks = can_work.copy()

                        # 1. 宿直(1名)の決定
                        night_assigned = False
                        valid_wishers.sort(key=lambda s: (monthly_night_count.get(s, 0), get_annual_count(s, "宿直")))
                        
                        for wisher in valid_wishers:
                            night_staff = wisher
                            daily_tasks[night_staff].append("宿直")
                            increment_annual_count(night_staff, "宿直")
                            staff_night_holiday_abs_days[night_staff].add(current_abs_day)
                            monthly_night_count[night_staff] += 1
                            night_assigned = True
                            break
                            
                        if not night_assigned:
                            valid_night = [s for s in night_candidates if can_do_night_or_holiday(s, current_abs_day, is_random_assignment=True) and monthly_night_count.get(s, 0) < 4]
                            if not valid_night:
                                valid_night = [s for s in night_candidates if monthly_night_count.get(s, 0) < 4]
                            if not valid_night:
                                valid_night = night_candidates
                                
                            valid_night.sort(key=lambda s: (monthly_night_count.get(s, 0), get_annual_count(s, "宿直")))
                            for cand in valid_night:
                                night_staff = cand
                                daily_tasks[night_staff].append("宿直")
                                increment_annual_count(night_staff, "宿直")
                                staff_night_holiday_abs_days[night_staff].add(current_abs_day)
                                monthly_night_count[night_staff] += 1
                                night_assigned = True
                                break

                        # 2. 非常勤の固定
                        required_tasks = ["Ｉ", "Ｏ", "Ｍ", "Ｄ", "Ｒ"]
                        if d_task_assigned_this_month:
                            if "Ｄ" in required_tasks:
                                required_tasks.remove("Ｄ")

                        available_part_timers = [s for s in available_for_tasks if s in part_time_staff]
                        if available_part_timers and "Ｍ" in required_tasks:
                            pt_staff = random.choice(available_part_timers)
                            daily_tasks[pt_staff].append("Ｍ")
                            increment_annual_count(pt_staff, "Ｍ")
                            available_for_tasks.remove(pt_staff)
                            required_tasks.remove("Ｍ")

                        # 3. アブレーション / 心外 / カテ の決定
                        # アブレーション
                        ablation_slots = 2 if is_ablation_day else 0
                        for _ in range(ablation_slots):
                            cands = [s for s in available_for_tasks if safe_int(staff_angio_dict.get(s, 0)) >= 3]
                            if cands:
                                cands.sort(key=lambda s: get_annual_count(s, "Ａ"))
                                chosen = cands[0]
                                daily_tasks[chosen].append("Ａ")
                                increment_annual_count(chosen, "Ａ")
                                available_for_tasks.remove(chosen)
                            else:
                                dummy_name = f"スタッフ{chr(65 + dummy_counter)}"
                                if dummy_name not in daily_tasks: daily_tasks[dummy_name] = []
                                daily_tasks[dummy_name].append("Ａ")
                                dummy_counter += 1

                        # 心外
                        for _ in range(heart_count):
                            hm_cands = [s for s in available_for_tasks if staff_ope_dict.get(s, 'A') >= 'C' and hm_min_counts.get(s, 0) > 0]
                            hm_staff = None
                            if hm_cands:
                                hm_cands.sort(key=lambda s: (
                                    1 if monthly_hm_count.get(s, 0) >= hm_min_counts.get(s, 0) else 0,
                                    get_annual_count(s, "心外")
                                ))
                                hm_staff = hm_cands[0]
                                
                            picked_h_subs = []
                            if hm_staff:
                                hm_rank = staff_ope_dict.get(hm_staff, 'A')
                                if hm_rank == 'C':
                                    d_cands = [s for s in available_for_tasks if staff_ope_dict.get(s, 'A') >= 'D' and s != hm_staff]
                                    d_cands.sort(key=lambda s: get_annual_count(s, "心外"))
                                    if d_cands: picked_h_subs.append(d_cands[0])
                                    
                                    other_cands = [s for s in available_for_tasks if staff_ope_dict.get(s, 'A') >= 'C' and s != hm_staff and s not in picked_h_subs]
                                    other_cands.sort(key=lambda s: get_annual_count(s, "心外"))
                                    if other_cands: picked_h_subs.append(other_cands[0])
                                else:
                                    sub_cands = [s for s in available_for_tasks if staff_ope_dict.get(s, 'A') >= 'C' and s != hm_staff]
                                    sub_cands.sort(key=lambda s: get_annual_count(s, "心外"))
                                    if sub_cands:
                                        first_sub = sub_cands[0]
                                        picked_h_subs.append(first_sub)
                                        if staff_ope_dict.get(first_sub, 'A') == 'C':
                                            other_cands = [s for s in available_for_tasks if staff_ope_dict.get(s, 'A') >= 'C' and s != hm_staff and s not in picked_h_subs]
                                            other_cands.sort(key=lambda s: get_annual_count(s, "心外"))
                                            if other_cands: picked_h_subs.append(other_cands[0])
                            
                            if hm_staff:
                                daily_tasks[hm_staff].append("ＨＭ")
                                increment_annual_count(hm_staff, "ＨＭ")
                                monthly_hm_count[hm_staff] += 1
                                available_for_tasks.remove(hm_staff)
                                
                                for h_sub in picked_h_subs:
                                    daily_tasks[h_sub].append("Ｈサ")
                                    increment_annual_count(h_sub, "Ｈサ")
                                    available_for_tasks.remove(h_sub)
                            else:
                                dummy_name_1 = f"スタッフ{chr(65 + dummy_counter)}"
                                if dummy_name_1 not in daily_tasks: daily_tasks[dummy_name_1] = []
                                daily_tasks[dummy_name_1].append("ＨＭ")
                                dummy_counter += 1
                                
                                dummy_name_2 = f"スタッフ{chr(65 + dummy_counter)}"
                                if dummy_name_2 not in daily_tasks: daily_tasks[dummy_name_2] = []
                                daily_tasks[dummy_name_2].append("Ｈサ")
                                dummy_counter += 1

                        # カテ
                        for _ in range(c_needed):
                            cands = [s for s in available_for_tasks if safe_int(staff_angio_dict.get(s, 0)) >= 2]
                            if cands:
                                cands.sort(key=lambda s: get_annual_count(s, "カ"))
                                chosen = cands[0]
                                daily_tasks[chosen].append("カ")
                                increment_annual_count(chosen, "カ")
                                available_for_tasks.remove(chosen)
                            else:
                                dummy_name = f"スタッフ{chr(65 + dummy_counter)}"
                                if dummy_name not in daily_tasks: daily_tasks[dummy_name] = []
                                daily_tasks[dummy_name].append("カ")
                                dummy_counter += 1

                        # 4. 基本業務 (I,O,R,D) の穴埋め
                        def is_high_skill(s):
                            return staff_ope_dict.get(s, 'A') >= 'C' or safe_int(staff_angio_dict.get(s, 0)) >= 2

                        random.shuffle(required_tasks)
                        for task in required_tasks:
                            if available_for_tasks:
                                cands = available_for_tasks.copy()
                                if task == "Ｄ":
                                    valid_d = [s for s in cands if monthly_d_count.get(s, 0) == 0]
                                    if valid_d: cands = valid_d
                                    
                                filtered_cands = [s for s in cands if yesterday_tasks.get(s) != task]
                                if not filtered_cands: filtered_cands = cands
                                
                                filtered_cands.sort(key=lambda s: (is_high_skill(s), get_annual_count(s, task)))
                                
                                chosen = filtered_cands[0]
                                daily_tasks[chosen].append(task)
                                increment_annual_count(chosen, task)
                                available_for_tasks.remove(chosen)
                                
                                if task == "Ｄ":
                                    d_task_assigned_this_month = True
                                    monthly_d_count[chosen] += 1
                            else:
                                dummy_name = f"スタッフ{chr(65 + dummy_counter)}"
                                if dummy_name not in daily_tasks: daily_tasks[dummy_name] = []
                                daily_tasks[dummy_name].append(task)
                                dummy_counter += 1
                                if task == "Ｄ": d_task_assigned_this_month = True

                        # フリー
                        for leftover in available_for_tasks:
                            daily_tasks[leftover].append("フリー")
                            increment_annual_count(leftover, "フリー")

                    # ④ 最終結合
                    for staff, tasks in daily_tasks.items():
                        if not tasks:
                            continue
                        
                        sorted_tasks = []
                        shukuchoku_task = None
                        for t in tasks:
                            if "宿直" in t:
                                shukuchoku_task = t
                            else:
                                sorted_tasks.append(t)
                                if not staff.startswith("スタッフ") and t not in ["フリー"]:
                                    yesterday_tasks[staff] = t
                                    
                        if shukuchoku_task:
                            sorted_tasks.append(shukuchoku_task)
                            
                        joined_task = "\n".join(sorted_tasks)
                        draft_data.append([d_str, staff, joined_task])

                if draft_data:
                    df_shift = fetch_data("確定勤務表", COLS_SHIFT)
                    month_prefix = f"{target_year}-{target_month:02d}"
                    if not df_shift.empty:
                        df_shift["_date_obj"] = pd.to_datetime(df_shift["日時"], errors="coerce").dt.date
                        df_shift["_date_str"] = df_shift["_date_obj"].apply(lambda x: x.strftime("%Y-%m") if pd.notnull(x) else "")
                        df_remain = df_shift[df_shift["_date_str"] != month_prefix].drop(columns=["_date_obj", "_date_str"], errors="ignore")
                    else: df_remain = pd.DataFrame(columns=COLS_SHIFT)
                        
                    df_new_month = pd.DataFrame(draft_data, columns=COLS_SHIFT)
                    df_final = pd.concat([df_remain, df_new_month], ignore_index=True)
                    res = overwrite_data("確定勤務表", df_final, COLS_SHIFT)
                    if res: st.success(f"{target_year}年{target_month}月の1ヶ月分ドラフトを一括作成し、保存しました！")
                else: st.warning("生成対象のデータがありませんでした。")
    st.markdown('</div>', unsafe_allow_html=True)

# --- メイン処理 ---
def main():
    st.sidebar.markdown("<h2>🏥 ME勤務表管理</h2>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    
    if 'is_admin' not in st.session_state:
        st.session_state['is_admin'] = False

    if st.session_state['is_admin']:
        pages = {
            "① 勤務表確認・編集": lambda: page_home(is_admin=True),
            "② 予定・希望休・マスタ管理": page_schedule_task,
            "③ スタッフマスタ管理": page_staff,
            "④ 勤務表作成（自動生成）": page_shift_creation
        }
    else:
        pages = {
            "① 勤務表の確認（閲覧のみ）": lambda: page_home(is_admin=False),
            "② 希望休入力": render_leave_request_ui
        }
        
    selection = st.sidebar.radio("メニュー", list(pages.keys()))
    
    st.sidebar.markdown("---")
    if not st.session_state['is_admin']:
        with st.sidebar.expander("🔑 管理者ログイン"):
            pwd = st.text_input("パスワード", type="password")
            if st.button("ログイン"):
                if pwd == "ME12345":
                    st.session_state['is_admin'] = True
                    st.rerun()
                else:
                    st.error("パスワードが違います。")
    else:
        st.sidebar.success("🔑 管理者としてログイン中")
        if st.sidebar.button("ログアウト"):
            st.session_state['is_admin'] = False
            st.rerun()
            
    st.sidebar.markdown("---")
    st.sidebar.caption("システムステータス:")
    client, error_msg = get_gspread_client()
    if client is None:
        st.sidebar.error("🔴 DB未接続 (認証エラー)")
        st.sidebar.error(f"エラー詳細:\n{error_msg}")
        st.sidebar.info("現在はプロトタイプ（UIデモ）として動作しています。")
    else:
        if "spreadsheet_id" not in st.secrets: st.sidebar.error("🔴 DB未接続 (ID未設定)")
        else:
            try:
                sheet_id = st.secrets["spreadsheet_id"]
                client.open_by_key(sheet_id)
                st.sidebar.success("🟢 DB接続済 (Google Sheets)")
            except Exception: st.sidebar.error("🔴 DB未接続 (アクセス失敗)")
            
    pages[selection]()

if __name__ == "__main__":
    main()
