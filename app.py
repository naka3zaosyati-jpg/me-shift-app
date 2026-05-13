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
        
        base_bg = "#FFFFFF"
        if is_sun_or_hol: base_bg = HOLIDAY_COLORS["日祝"]
        elif is_sat: base_bg = HOLIDAY_COLORS["土"]
        
        # 本日は最優先でハイライト
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

# --- ヘルパー関数 ---
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

# --- Google Sheets API 接続 ---
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

def append_rows_batch(sheet_name, rows_data):
    client, _ = get_gspread_client()
    if not client: return False
    try:
        sheet_id = st.secrets.get("spreadsheet_id", "")
        sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
        res = sheet.append_rows(rows_data, value_input_option="USER_ENTERED", insert_data_option="INSERT_ROWS", table_range="A1")
        st.cache_data.clear()
        return res
    except Exception as e:
        st.error(f"シート「{sheet_name}」への一括書き込みでエラーが発生しました。\n詳細: {e}")
        return False

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

# --- カラム定義 ---
COLS_STAFF = ["氏名", "役職", "OPE習熟度", "アンギオ習熟度", "総合コード", "人工心肺メイン回数", "人工心肺サブ回数", "アブレーション回数", "カテ回数", "雇用形態"]
COLS_REQUEST = ["日時", "氏名", "区分", "コメント"]
COLS_OPE_MASTER = ["術式名", "術式レベル"]
COLS_OPE_SCHEDULE = ["日時", "術式", "業務"]
COLS_TASK_MASTER = ["略語", "業務名"]
COLS_SHIFT = ["日時", "氏名", "割り当て業務"]
COLS_PROFICIENCY_MASTER = ["区分", "ランク", "定義"]

# --- ページUI コンポーネント ---

def page_home():
    st.markdown('<div class="card"><h2>① ホーム（月間勤務表・シフト）</h2><p>確定した勤務表や術式予定を確認・編集します。</p></div>', unsafe_allow_html=True)
    
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
        for _, row in df_shift.iterrows():
            d_key = parse_date(row["日時"])
            if d_key:
                staff_name = str(row['氏名']).strip()
                staff_info = f"{staff_name} ({row['割り当て業務']})"
                if d_key not in shift_dict: shift_dict[d_key] = []
                shift_dict[d_key].append((staff_name, staff_info))
                
    # カレンダーの表示順を staff_list の順番に合わせる
    def get_staff_idx(name):
        try: return staff_list.index(name)
        except ValueError: return 999
        
    for k in shift_dict:
        shift_dict[k].sort(key=lambda x: get_staff_idx(x[0]))
        shift_dict[k] = [x[1] for x in shift_dict[k]] # 名前を除いて表示用文字列のみにする

    ope_dict = {}
    if not df_ope.empty:
        for _, row in df_ope.iterrows():
            d_key = parse_date(row["日時"])
            if d_key:
                ope_info = str(row['術式'])
                if d_key not in ope_dict: ope_dict[d_key] = []
                ope_dict[d_key].append(ope_info)
                
    # --- カレンダー描画 ---
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
    
    # --- 横型シフト表 ＆ 集計表 (st.data_editor化) ---
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
            ope_str = "\n".join(opes) if opes else ""
            
            # ヘッダーを1つの文字列に結合（改行あり）
            header_str = f"{d}日\n{weekday_str}"
            if ope_str: header_str += f"\n{ope_str}"
            cols_strings.append(header_str)
            
        shift_matrix = pd.DataFrame(index=staff_list, columns=cols_strings)
        shift_matrix.fillna("", inplace=True)
        
        if not df_shift.empty:
            for _, row in df_shift.iterrows():
                d_key = parse_date(row["日時"])
                if d_key:
                    try:
                        dt = datetime.datetime.strptime(d_key, "%Y-%m-%d")
                        if dt.year == selected_year and dt.month == selected_month:
                            staff_name = str(row["氏名"]).strip()
                            duty = str(row["割り当て業務"]).strip()
                            header_col = cols_strings[dt.day - 1]
                            
                            if staff_name in shift_matrix.index or staff_name.startswith("スタッフ"):
                                if staff_name not in shift_matrix.index:
                                    shift_matrix.loc[staff_name] = ""
                                curr_val = shift_matrix.at[staff_name, header_col]
                                if curr_val: shift_matrix.at[staff_name, header_col] = curr_val + "\n" + duty
                                else: shift_matrix.at[staff_name, header_col] = duty
                    except Exception:
                        pass

        # プルダウン（SelectboxColumn）用の選択肢を生成
        base_options = ["", "フリー", "ＨＭ", "Ｈサ", "Ａ", "カ", "Ｉ", "Ｏ", "Ｍ", "Ｄ", "Ｒ", "日勤", "宿直"]
        combos = [o + "\n宿直" for o in base_options if o and o != "宿直"]
        all_options = base_options + combos
        
        # すべてのカラムにプルダウン設定を適用
        column_config = {}
        for col in cols_strings:
            column_config[col] = st.column_config.SelectboxColumn(
                options=all_options,
                default=""
            )

        styled_matrix = shift_matrix.style.apply(lambda df: style_shift_matrix(df, today_date), axis=None)
            
        edited_matrix = st.data_editor(
            styled_matrix,
            use_container_width=True,
            height=600,
            key="month_shift_editor",
            column_config=column_config
        )
        
        if st.button("💾 編集内容をスプレッドシートに保存"):
            with st.spinner("スプレッドシートに保存中..."):
                new_draft = []
                # アンピボットして1次元配列に戻す
                for staff in edited_matrix.index:
                    for d in range(1, num_days + 1):
                        dt = datetime.date(selected_year, selected_month, d)
                        d_str = dt.strftime("%Y-%m-%d")
                        header_str = cols_strings[d-1]
                        
                        val = str(edited_matrix.at[staff, header_str]).strip()
                        if val and val != "nan" and val != "None":
                            tasks = val.split('\n')
                            for t in tasks:
                                t = t.strip()
                                if t: new_draft.append([d_str, staff, t])
                                
                if new_draft:
                    df_shift_all = fetch_data("確定勤務表", COLS_SHIFT)
                    month_prefix = f"{selected_year}-{selected_month:02d}"
                    if not df_shift_all.empty:
                        if "_date_str" not in df_shift_all.columns:
                            df_shift_all["_date_str"] = df_shift_all["日時"].apply(parse_date)
                        df_remain = df_shift_all[~df_shift_all["_date_str"].astype(str).str.startswith(month_prefix)].drop(columns=["_date_str"], errors="ignore")
                    else:
                        df_remain = pd.DataFrame(columns=COLS_SHIFT)
                        
                    df_new_month = pd.DataFrame(new_draft, columns=COLS_SHIFT)
                    df_final = pd.concat([df_remain, df_new_month], ignore_index=True)
                    
                    res = overwrite_data("確定勤務表", df_final, COLS_SHIFT)
                    if res: st.success(f"{selected_year}年{selected_month}月の編集内容を一括保存しました！")
                else:
                    st.warning("保存するデータがありません。")
                    
        try:
            buffer = io.BytesIO()
            try:
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer: shift_matrix.to_excel(writer, sheet_name=f"{selected_month}月シフト")
            except Exception:
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer: shift_matrix.to_excel(writer, sheet_name=f"{selected_month}月シフト")
            st.download_button(label="📥 Excel形式でダウンロード (.xlsx)", data=buffer.getvalue(), file_name=f"shift_{selected_year}_{selected_month}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception:
            csv = shift_matrix.to_csv(encoding="utf-8-sig")
            st.download_button(label="📥 CSV形式でダウンロード (.csv)", data=csv, file_name=f"shift_{selected_year}_{selected_month}.csv", mime="text/csv")

    with tab_summary:
        st.write(f"### 📊 {selected_year}年{selected_month}月 業務割り当て総合回数")
        if not df_shift.empty:
            df_month = df_shift.copy()
            df_month["日時"] = df_month["日時"].apply(parse_date)
            month_prefix = f"{selected_year}-{selected_month:02d}"
            df_month = df_month[df_month["日時"].astype(str).str.startswith(month_prefix)]
            if not df_month.empty:
                try:
                    summary_df = pd.pivot_table(df_month, index="氏名", columns="割り当て業務", aggfunc="size", fill_value=0)
                    summary_df["合計"] = summary_df.sum(axis=1)
                    st.dataframe(summary_df, use_container_width=True)
                except Exception as e: st.warning(f"集計処理中にエラーが発生しました: {e}")
            else: st.info(f"{selected_year}年{selected_month}月の勤務表データがありません。")
        else: st.info("集計する勤務表データがありません。")
    st.markdown('</div>', unsafe_allow_html=True)

def page_schedule_task():
    st.markdown('<div class="card"><h2>② 予定・マスタ管理</h2><p>術式予定、希望休入力、各種マスタの管理を行います。</p></div>', unsafe_allow_html=True)
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
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 希望休入力")
        with st.form("request_form", clear_on_submit=True):
            req_date = st.date_input("希望日時")
            df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
            staff_names = df_staff["氏名"].tolist() if not df_staff.empty else ["テスト太郎", "テスト花子"]
            req_name = st.selectbox("氏名", staff_names)
            req_type = st.radio("区分", ["× (不可)", "△ (宿直希望)", "年 (年休希望)"])
            req_comment = st.text_input("コメント")
            if st.form_submit_button("希望休を登録"):
                val_type = "×" if "×" in req_type else ("年" if "年" in req_type else "△")
                res = append_data("希望入力", [str(req_date), req_name, val_type, req_comment])
                if res: st.success("スプレッドシートに希望休を書き込みました。")
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 登録済み希望一覧")
        st.dataframe(fetch_data("希望入力", COLS_REQUEST), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

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
            else: st.info("データがありません。")
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
            else: st.info("データがありません。")
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
    
    # 習熟度マスタデータの取得
    df_prof = fetch_data("習熟度マスタ", COLS_PROFICIENCY_MASTER)
    ope_options = ["A", "B", "C", "D"] # デフォルト
    angio_options = ["1", "2", "3", "4"] # デフォルト
    ope_defs = ""
    angio_defs = ""
    
    if not df_prof.empty:
        df_ope_prof = df_prof[df_prof["区分"] == "OPE"]
        if not df_ope_prof.empty:
            ope_options = df_ope_prof["ランク"].astype(str).tolist()
            ope_defs = "\n".join([f"- **{r['ランク']}**: {r['定義']}" for _, r in df_ope_prof.iterrows()])
            
        df_angio_prof = df_prof[df_prof["区分"] == "アンギオ"]
        if not df_angio_prof.empty:
            angio_options = df_angio_prof["ランク"].astype(str).tolist()
            angio_defs = "\n".join([f"- **{r['ランク']}**: {r['定義']}" for _, r in df_angio_prof.iterrows()])

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
                if ope_defs: st.caption("【OPE習熟度の定義】\n" + ope_defs)
                
                angio_level = st.selectbox("アンギオ習熟度", angio_options)
                if angio_defs: st.caption("【アンギオ習熟度の定義】\n" + angio_defs)
                
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
                    if ope_defs: st.caption("【OPE習熟度の定義】\n" + ope_defs)
                    
                    a_idx = angio_options.index(str(curr["アンギオ習熟度"])) if str(curr["アンギオ習熟度"]) in angio_options else 0
                    angio_level_upd = st.selectbox("アンギオ習熟度", angio_options, index=a_idx)
                    if angio_defs: st.caption("【アンギオ習熟度の定義】\n" + angio_defs)
                
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
                    row_upd = [staff_name_upd, staff_role_upd, ope_level_upd, angio_level_upd, total_code_upd, int(cpb_main_upd), int(cpb_sub_upd), int(ablation_upd), int(catha_upd), employment_type_upd]
                    res = update_data("スタッフマスタ", 1, staff_name_upd, row_upd)
                    if res: st.success(f"「{staff_name_upd}」さんの情報を上書き更新しました。（新しい総合コード: {total_code_upd}）")
        else: st.info("登録されているスタッフがいません。")
            
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### スタッフ一覧")
    df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
    st.dataframe(df_staff, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

def page_shift_creation():
    st.markdown('<div class="card"><h2>④ 勤務表作成</h2><p>勤務表の「1ヶ月一括作成」を行います。（微調整はホーム画面の横型シフト表から直接編集・保存できます）</p></div>', unsafe_allow_html=True)
    
    with st.expander("現在の自動割り当てアルゴリズム（条件）"):
        st.markdown("""
- **優先割り当て**: 「非常勤」のスタッフは、平日は最優先で「Ｍ（ME業務）」に配置されます。
- **必要枠の生成**: 平日は基本業務（カ, Ｉ, Ｏ, Ｍ, Ｄ, Ｒ）各1名を配置します。曜日別設定の他、「術式予定」の業務項目と連動し、心外枠・アブレーション枠が追加されます。
- **宿直（平日）**: 宿直枠は独立させず、その日の日勤者の中から選出され兼任します（週1回制限あり）。
- **休日・祝日**: 1名体制とし、その1名が日勤と宿直を兼任します。平日の通常業務は配置しません（連休中は1回まで）。
- **特殊条件**: Ｄ（DC確認）業務は月に1回のみ配置されます。
- **心外枠のスキル制限**: スタッフマスタの「OPE習熟度」が C 以降のスタッフのみ選出。D以降が最低1名含まれます（Cが含まれる場合は3名体制）。
- **アブレーション枠のスキル制限**: 「アンギオ習熟度」が 3 以降のスタッフのみ選出されます。
- **公平性の担保**: 全ての業務において、年度ごとの年間実績回数が最も少ないスタッフを優先的に選出し均等化を図ります。
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
            with cols_heart[i]:
                new_heart[d] = st.checkbox(d, value=default_settings["心外"][d], key=f"heart_{d}")
                
        st.write("#### アブレーション枠（Ａ）")
        cols_ab = st.columns(5)
        new_ab = {}
        for i, d in enumerate(["月", "火", "水", "木", "金"]):
            with cols_ab[i]:
                new_ab[d] = st.checkbox(d, value=default_settings["アブレーション"][d], key=f"ab_{d}")
                
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
    st.info("指定した月の1日〜月末までのシフトを自動生成し、スプレッドシートに保存します。\n※既に該当月のデータがある場合は、対象月のデータがすべて「洗い替え（上書き）」されます。")
    
    col1, col2 = st.columns(2)
    today = datetime.date.today()
    with col1: target_year = st.number_input("対象年", min_value=2020, max_value=2050, value=today.year, step=1)
    with col2: target_month = st.number_input("対象月", min_value=1, max_value=12, value=today.month, step=1)
    
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
        if "_date_str" not in df_history_ui.columns:
            df_history_ui["_date_str"] = df_history_ui["日時"].apply(parse_date)
            
        for _, row in df_history_ui.iterrows():
            if pd.isna(row["_date_str"]) or not row["_date_str"]: continue
            h_date = pd.to_datetime(row["_date_str"])
            s_name = str(row["氏名"]).strip()
            t_name = str(row["割り当て業務"]).strip()
            
            if s_name not in staff_list_summary: continue
            
            # Annual (include target_month for UI reflection after draft)
            if h_date.date() >= fy_start_ui and (h_date.year < target_year or (h_date.year == target_year and h_date.month <= target_month)):
                if t_name in ["ＨＭ", "Ｈサ"]: annual_summary[s_name]["心外(HM/Hサ)"] += 1
                elif t_name == "Ａ": annual_summary[s_name]["A"] += 1
                elif t_name == "カ": annual_summary[s_name]["カ"] += 1
                elif t_name == "Ｉ": annual_summary[s_name]["I"] += 1
                elif t_name == "Ｏ": annual_summary[s_name]["O"] += 1
                elif t_name == "Ｍ": annual_summary[s_name]["M"] += 1
                elif t_name == "Ｄ": annual_summary[s_name]["D"] += 1
                elif t_name == "Ｒ": annual_summary[s_name]["R"] += 1
                elif t_name == "宿直": annual_summary[s_name]["宿直"] += 1
                
            # Monthly
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
                    df_staff = df_staff[df_staff["氏名"].astype(str).str.strip() != ""]
                    df_staff = df_staff[~df_staff["氏名"].astype(str).str.lower().isin(["none", "nan"])]
                
                df_request = fetch_data("希望入力", COLS_REQUEST)
                
                staff_list = df_staff["氏名"].astype(str).str.strip().tolist() if not df_staff.empty else []
                
                part_time_staff = []
                staff_ope_dict = {}
                staff_angio_dict = {}
                if "雇用形態" in df_staff.columns:
                    df_staff["雇用形態"] = df_staff["雇用形態"].astype(str).str.strip()
                    df_staff["氏名"] = df_staff["氏名"].astype(str).str.strip()
                    part_time_staff = df_staff[df_staff["雇用形態"] == "非常勤"]["氏名"].tolist()
                
                if not df_staff.empty:
                    for _, row in df_staff.iterrows():
                        sname = str(row["氏名"]).strip()
                        staff_ope_dict[sname] = str(row["OPE習熟度"]).strip()
                        staff_angio_dict[sname] = str(row["アンギオ習熟度"]).strip()
                
                req_unavailable = {}
                req_night_shift = {}
                if not df_request.empty:
                    for _, row in df_request.iterrows():
                        r_date = parse_date(row["日時"])
                        if not r_date: continue
                        kubun = str(row["区分"])
                        staff_n = str(row["氏名"]).strip()
                        
                        if "×" in kubun or "年" in kubun:
                            if r_date not in req_unavailable: req_unavailable[r_date] = []
                            req_unavailable[r_date].append(staff_n)
                        elif "△" in kubun:
                            if r_date not in req_night_shift: req_night_shift[r_date] = []
                            req_night_shift[r_date].append(staff_n)
                                
                _, num_days = calendar.monthrange(target_year, target_month)
                draft_data = []
                
                # --- 年間を通じた回数均等化の準備 ---
                fiscal_year = target_year if target_month >= 4 else target_year - 1
                fy_start = datetime.date(fiscal_year, 4, 1)
                
                annual_task_counts = {s: {} for s in staff_list}
                monthly_d_count = {s: 0 for s in staff_list}
                staff_night_holiday_abs_days = {s: set() for s in staff_list}
                
                df_history = fetch_data("確定勤務表", COLS_SHIFT)
                if not df_history.empty:
                    if "_date_str" not in df_history.columns:
                        df_history["_date_str"] = df_history["日時"].apply(parse_date)
                    for _, row in df_history.iterrows():
                        if pd.isna(row["_date_str"]) or not row["_date_str"]: continue
                        h_date = pd.to_datetime(row["_date_str"])
                        s_name = str(row["氏名"]).strip()
                        t_name = str(row["割り当て業務"]).strip()
                        
                        if s_name not in staff_list: continue
                        
                        abs_day = (h_date.date() - datetime.date(2020, 1, 1)).days
                        
                        if t_name in ["宿直", "日勤"]:
                            staff_night_holiday_abs_days[s_name].add(abs_day)
                            
                        if h_date.date() >= fy_start and (h_date.year < target_year or (h_date.year == target_year and h_date.month < target_month)):
                            annual_task_counts[s_name][t_name] = annual_task_counts[s_name].get(t_name, 0) + 1
                            if t_name == "日勤" and is_off_day_func(h_date.date()):
                                annual_task_counts[s_name]["休日出勤"] = annual_task_counts[s_name].get("休日出勤", 0) + 1
                                
                        if h_date.year == target_year and h_date.month == target_month and t_name == "Ｄ":
                            monthly_d_count[s_name] += 1

                monthly_night_count = {s: 0 for s in staff_list}
                monthly_holiday_night_count = {s: 0 for s in staff_list}

                def get_annual_count(staff, task):
                    if task == "心外": return annual_task_counts.get(staff, {}).get("ＨＭ", 0) + annual_task_counts.get(staff, {}).get("Ｈサ", 0)
                    return annual_task_counts.get(staff, {}).get(task, 0)
                def increment_annual_count(staff, task):
                    if staff not in annual_task_counts: annual_task_counts[staff] = {}
                    annual_task_counts[staff][task] = annual_task_counts[staff].get(task, 0) + 1
                    
                def can_do_night_or_holiday(staff, current_abs_day):
                    for d in staff_night_holiday_abs_days[staff]:
                        if abs(d - current_abs_day) <= 3: return False
                    return True

                holiday_block_map = {}
                current_block = []
                for d in range(1, num_days + 1):
                    dt = datetime.date(target_year, target_month, d)
                    if is_off_day_func(dt):
                        current_block.append(d)
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
                    for _, row in df_ope.iterrows():
                        d_key = parse_date(row["日時"])
                        if d_key:
                            gyomu = str(row.get('業務', '')).strip()
                            if d_key not in ope_gyomu_dict: ope_gyomu_dict[d_key] = []
                            if gyomu in ["心外", "アブレーション"]:
                                ope_gyomu_dict[d_key].append(gyomu)

                d_task_assigned_this_month = False
                
                for d in range(1, num_days + 1):
                    dt = datetime.date(target_year, target_month, d)
                    d_str = dt.strftime("%Y-%m-%d")
                    current_abs_day = (dt - datetime.date(2020, 1, 1)).days
                    
                    is_off_day = is_off_day_func(dt)
                    
                    required_tasks = []
                    opes_gyomu = ope_gyomu_dict.get(d_str, [])
                    heart_count = opes_gyomu.count("心外")
                    ablation_count = opes_gyomu.count("アブレーション")
                    
                    if is_off_day:
                        required_tasks = ["日勤"]
                    else:
                        required_tasks = ["カ", "Ｉ", "Ｏ", "Ｍ", "Ｄ", "Ｒ"]
                        week_cols = ["月", "火", "水", "木", "金"]
                        if 0 <= dt.weekday() <= 4:
                            day_str = week_cols[dt.weekday()]
                            if default_settings["心外"][day_str]: heart_count += 1
                            if default_settings["アブレーション"][day_str]: ablation_count += 1
                    
                    for _ in range(heart_count): required_tasks.extend(["ＨＭ", "Ｈサ"])
                    for _ in range(ablation_count): required_tasks.extend(["Ａ", "Ａ"])
                        
                    if d_task_assigned_this_month and "Ｄ" in required_tasks:
                        required_tasks.remove("Ｄ")
                            
                    unavailable = req_unavailable.get(d_str, [])
                    available_staff = [s for s in staff_list if s not in unavailable]
                    
                    if is_off_day:
                        available_staff = [s for s in available_staff if s not in part_time_staff]
                        
                    assigned_today_staffs = []
                    dummy_counter = 0
                    
                    # 1. Ｍ (非常勤固定)
                    if "Ｍ" in required_tasks:
                        available_part_timers = [s for s in available_staff if s in part_time_staff]
                        if available_part_timers:
                            random.shuffle(available_part_timers)
                            pt_staff = available_part_timers[0]
                            draft_data.append([d_str, pt_staff, "Ｍ"])
                            assigned_today_staffs.append(pt_staff)
                            increment_annual_count(pt_staff, "Ｍ")
                            available_staff.remove(pt_staff)
                            required_tasks.remove("Ｍ")
                    
                    # 2. アブレーション枠 (Ａ)
                    a_needed = required_tasks.count("Ａ")
                    for _ in range(a_needed):
                        required_tasks.remove("Ａ")
                        candidates = [s for s in available_staff if safe_int(staff_angio_dict.get(s, 0)) >= 3]
                        if not candidates:
                            assigned_staff = f"スタッフ{chr(65 + dummy_counter)}"
                            dummy_counter += 1
                        else:
                            candidates.sort(key=lambda s: get_annual_count(s, "Ａ"))
                            assigned_staff = candidates[0]
                            increment_annual_count(assigned_staff, "Ａ")
                            available_staff.remove(assigned_staff)
                            assigned_today_staffs.append(assigned_staff)
                        draft_data.append([d_str, assigned_staff, "Ａ"])

                    # 3. 心外枠 (ＨＭ, Ｈサ)
                    hm_needed = required_tasks.count("ＨＭ")
                    while hm_needed > 0:
                        required_tasks.remove("ＨＭ")
                        if "Ｈサ" in required_tasks: required_tasks.remove("Ｈサ")
                        hm_needed -= 1
                        
                        cands = [s for s in available_staff if staff_ope_dict.get(s, 'A') >= 'C']
                        cands.sort(key=lambda s: get_annual_count(s, "心外"))
                        
                        picked = []
                        d_cands = [s for s in cands if staff_ope_dict.get(s, 'A') >= 'D']
                        if d_cands:
                            first = d_cands[0]
                            picked.append(first)
                            cands.remove(first)
                        else:
                            if cands:
                                first = cands[0]
                                picked.append(first)
                                cands.remove(first)
                                
                        if cands:
                            second = cands[0]
                            picked.append(second)
                            cands.remove(second)
                            
                        has_c = any(staff_ope_dict.get(s, 'A') == 'C' for s in picked)
                        if has_c and cands:
                            third = cands[0]
                            picked.append(third)
                            cands.remove(third)
                            
                        if picked:
                            picked.sort(key=lambda s: staff_ope_dict.get(s, 'A'), reverse=True)
                            hm_staff = picked[0]
                            draft_data.append([d_str, hm_staff, "ＨＭ"])
                            increment_annual_count(hm_staff, "ＨＭ")
                            available_staff.remove(hm_staff)
                            assigned_today_staffs.append(hm_staff)
                            
                            for h_sub in picked[1:]:
                                draft_data.append([d_str, h_sub, "Ｈサ"])
                                increment_annual_count(h_sub, "Ｈサ")
                                available_staff.remove(h_sub)
                                assigned_today_staffs.append(h_sub)
                                
                    # 4. カテ枠 (カ)
                    c_needed = required_tasks.count("カ")
                    for _ in range(c_needed):
                        required_tasks.remove("カ")
                        candidates = [s for s in available_staff if safe_int(staff_angio_dict.get(s, 0)) >= 2]
                        if not candidates:
                            assigned_staff = f"スタッフ{chr(65 + dummy_counter)}"
                            dummy_counter += 1
                        else:
                            candidates.sort(key=lambda s: get_annual_count(s, "カ"))
                            assigned_staff = candidates[0]
                            increment_annual_count(assigned_staff, "カ")
                            available_staff.remove(assigned_staff)
                            assigned_today_staffs.append(assigned_staff)
                        draft_data.append([d_str, assigned_staff, "カ"])

                    # 5. 宿直の決定 (平日のみ、休日は「日勤」が兼ねるため後で処理)
                    night_staff = None
                    if not is_off_day:
                        night_cands_all = [s for s in (assigned_today_staffs + available_staff) if s in staff_list and s not in part_time_staff]
                        night_cands_all = list(dict.fromkeys(night_cands_all))
                        
                        valid_night = [s for s in night_cands_all if can_do_night_or_holiday(s, current_abs_day) and monthly_night_count.get(s, 0) < 4]
                        if not valid_night:
                            valid_night = [s for s in night_cands_all if monthly_night_count.get(s, 0) < 4]
                        if not valid_night:
                            valid_night = night_cands_all
                            
                        wishers = [s for s in req_night_shift.get(d_str, []) if s in valid_night]
                        if wishers:
                            night_staff = wishers[0]
                        else:
                            if valid_night:
                                valid_night.sort(key=lambda s: (monthly_night_count.get(s, 0), get_annual_count(s, "宿直")))
                                night_staff = valid_night[0]
                                
                        if night_staff:
                            increment_annual_count(night_staff, "宿直")
                            draft_data.append([d_str, night_staff, "宿直"])
                            staff_night_holiday_abs_days[night_staff].add(current_abs_day)
                            monthly_night_count[night_staff] += 1
                            
                    # 高スキル資格者判定（基本業務の優先度用）
                    def is_high_skill(s):
                        return staff_ope_dict.get(s, 'A') >= 'C' or safe_int(staff_angio_dict.get(s, 0)) >= 2

                    # 6. 基本業務の穴埋め (I, O, M, D, R, 日勤)
                    for task in required_tasks:
                        assigned_staff = None
                        if available_staff:
                            candidates_for_task = available_staff.copy()
                            
                            if task == "Ｄ":
                                valid_d = [s for s in candidates_for_task if monthly_d_count.get(s, 0) == 0]
                                if valid_d: candidates_for_task = valid_d
                                
                            if not candidates_for_task:
                                assigned_staff = f"スタッフ{chr(65 + dummy_counter)}"
                                dummy_counter += 1
                                if task == "Ｄ": d_task_assigned_this_month = True
                            else:
                                if is_off_day and task == "日勤":
                                    valid_cands = [s for s in candidates_for_task if monthly_night_count.get(s, 0) < 4 and monthly_holiday_night_count.get(s, 0) < 1]
                                    if not valid_cands:
                                        valid_cands = [s for s in candidates_for_task if monthly_night_count.get(s, 0) < 4 and monthly_holiday_night_count.get(s, 0) < 2]
                                    if not valid_cands:
                                        valid_cands = candidates_for_task
                                    
                                    valid_cands2 = [s for s in valid_cands if can_do_night_or_holiday(s, current_abs_day) and not has_worked_in_block(s, d)]
                                    if not valid_cands2: valid_cands2 = [s for s in valid_cands if not has_worked_in_block(s, d)]
                                    if not valid_cands2: valid_cands2 = valid_cands
                                        
                                    candidates_for_task = valid_cands2
                                        
                                    wishers = [s for s in req_night_shift.get(d_str, []) if s in candidates_for_task]
                                    if wishers:
                                        candidates_for_task = wishers
                                    else:
                                        # 年間休日出勤回数と年間宿直回数で均等化
                                        candidates_for_task.sort(key=lambda s: (monthly_holiday_night_count.get(s, 0), monthly_night_count.get(s, 0), get_annual_count(s, "休日出勤"), get_annual_count(s, "宿直")))
                                else:
                                    # 平日の基本業務：高スキルを持たない人を優先（Falseが先）、その後、累計回数
                                    candidates_for_task.sort(key=lambda s: (is_high_skill(s), get_annual_count(s, task)))
                                    
                                assigned_staff = candidates_for_task[0]
                                increment_annual_count(assigned_staff, task)
                                available_staff.remove(assigned_staff)
                                
                                if is_off_day and task == "日勤":
                                    staff_night_holiday_abs_days[assigned_staff].add(current_abs_day)
                                    staff_assigned_holiday_days[assigned_staff].add(d)
                                    monthly_night_count[assigned_staff] += 1
                                    monthly_holiday_night_count[assigned_staff] += 1
                                    increment_annual_count(assigned_staff, "休日出勤")
                                    
                                    # 休日日勤は宿直も兼ねる
                                    increment_annual_count(assigned_staff, "宿直")
                                    draft_data.append([d_str, assigned_staff, "宿直"])
                                    
                                if task == "Ｄ":
                                    d_task_assigned_this_month = True
                                    monthly_d_count[assigned_staff] += 1
                        else:
                            assigned_staff = f"スタッフ{chr(65 + dummy_counter)}"
                            dummy_counter += 1
                            if task == "Ｄ": d_task_assigned_this_month = True
                            
                        draft_data.append([d_str, assigned_staff, task])
                        if assigned_staff in staff_list:
                            assigned_today_staffs.append(assigned_staff)
                            
                    # フリーの割り当て
                    if not is_off_day:
                        for leftover_staff in available_staff:
                            draft_data.append([d_str, leftover_staff, "フリー"])
                            increment_annual_count(leftover_staff, "フリー")
                        
                if draft_data:
                    df_shift = fetch_data("確定勤務表", COLS_SHIFT)
                    month_prefix = f"{target_year}-{target_month:02d}"
                    if not df_shift.empty:
                        if "_date_str" not in df_shift.columns:
                            df_shift["_date_str"] = df_shift["日時"].apply(parse_date)
                        df_remain = df_shift[~df_shift["_date_str"].astype(str).str.startswith(month_prefix)].drop(columns=["_date_str"], errors="ignore")
                    else:
                        df_remain = pd.DataFrame(columns=COLS_SHIFT)
                        
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
    pages = {
        "① ホーム（月間勤務表・シフト）": page_home,
        "② 予定・マスタ管理": page_schedule_task,
        "③ スタッフマスタ管理": page_staff,
        "④ 勤務表作成": page_shift_creation
    }
    selection = st.sidebar.radio("メニュー", list(pages.keys()))
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
            except Exception:
                st.sidebar.error("🔴 DB未接続 (アクセス失敗)")
    pages[selection]()

if __name__ == "__main__":
    main()
