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
    /* 全体的なライトテーマ */
    .stApp {
        background-color: #F8F9FA;
        color: #212529;
    }
    
    /* 見出しの色を明るいアクセントカラーに */
    h1, h2, h3 {
        color: #0056B3 !important;
    }
    
    /* カード型UIの定義 (ライトモード用、影を柔らかく) */
    .card {
        background-color: #FFFFFF;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        border: 1px solid #E9ECEF;
        margin-bottom: 1.5rem;
    }
    
    /* ボタンのスタイル */
    .stButton>button {
        background-color: #007BFF;
        color: #FFFFFF;
        font-weight: bold;
        border-radius: 8px;
        border: none;
        width: 100%;
        padding: 0.6rem;
        transition: all 0.2s;
    }
    .stButton>button:hover {
        background-color: #0056B3;
        color: #FFFFFF;
        transform: translateY(-2px);
    }

    /* データフレーム（表）のヘッダー色 */
    th {
        background-color: #F1F3F5 !important;
        color: #495057 !important;
    }
    
    /* モバイル向けに入力フィールド等を角丸に調整 */
    input, select, textarea {
        border-radius: 6px !important;
    }
</style>
""", unsafe_allow_html=True)

# --- ヘルパー関数 ---
def safe_int(val):
    """NaNや空文字を安全に0に変換するヘルパー関数"""
    if pd.isna(val) or val == "":
        return 0
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return 0

def parse_date(date_val):
    """日付文字列を YYYY-MM-DD 形式に統一して返す（バグ回避用）"""
    if pd.isna(date_val) or str(date_val).strip() == "":
        return None
    d_str = str(date_val).strip()
    try:
        # pandasのto_datetimeを使用して多様な形式（YYYY/MM/DD等）を柔軟にパース
        dt = pd.to_datetime(d_str)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return None

# --- Google Sheets API 接続 ---
def get_gspread_client():
    """
    gspreadクライアントと、エラー発生時のメッセージを返す
    """
    try:
        if "gcp_service_account" in st.secrets:
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            info = dict(st.secrets["gcp_service_account"])
            credentials = Credentials.from_service_account_info(
                info,
                scopes=scopes
            )
            client = gspread.authorize(credentials)
            return client, None
        else:
            return None, "st.secrets に 'gcp_service_account' が設定されていません。"
    except Exception as e:
        return None, str(e)


# --- CRUD ラッパー関数 ---

@st.cache_data(ttl=60)
def _fetch_records_cached(sheet_name):
    """
    API呼び出しをキャッシュして高速化しつつ、更新時にクリアできるようにする内部関数
    """
    client, _ = get_gspread_client()
    if not client:
        return None
    try:
        sheet_id = st.secrets.get("spreadsheet_id", "")
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet(sheet_name)
        return sheet.get_all_records()
    except Exception as e:
        st.error(f"データ取得エラー ({sheet_name}): {e}")
        return None

def fetch_data(sheet_name, expected_columns):
    """
    キャッシュされたデータを取得し、DataFrameに整形する
    """
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
    """
    スプレッドシートへの直接書き込みとキャッシュクリアを実行する。
    """
    client, _ = get_gspread_client()
    if not client:
        st.error(f"DB未接続のため、シート「{sheet_name}」への書き込みができません。")
        return False
        
    try:
        sheet_id = st.secrets.get("spreadsheet_id", "")
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet(sheet_name)
        
        res = sheet.append_row(
            row_data, 
            value_input_option="USER_ENTERED",
            insert_data_option="INSERT_ROWS",
            table_range="A1"
        )
        st.cache_data.clear()
        return res
    except Exception as e:
        st.error(f"シート「{sheet_name}」への書き込みでエラーが発生しました。\n詳細: {e}")
        return False

def append_rows_batch(sheet_name, rows_data):
    """
    スプレッドシートへの複数行の一括書き込みを実行する。（ドラフト作成用）
    """
    client, _ = get_gspread_client()
    if not client:
        st.error(f"DB未接続のため、シート「{sheet_name}」への一括書き込みができません。")
        return False
        
    try:
        sheet_id = st.secrets.get("spreadsheet_id", "")
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet(sheet_name)
        
        res = sheet.append_rows(
            rows_data, 
            value_input_option="USER_ENTERED",
            insert_data_option="INSERT_ROWS",
            table_range="A1"
        )
        st.cache_data.clear()
        return res
    except Exception as e:
        st.error(f"シート「{sheet_name}」への一括書き込みでエラーが発生しました。\n詳細: {e}")
        return False

def update_data(sheet_name, search_col_index, search_value, row_data):
    """
    指定した列(search_col_index: 1始まり)から search_value を検索し、
    見つかった行を row_data で上書き更新する。
    """
    client, _ = get_gspread_client()
    if not client:
        st.error(f"DB未接続のため、シート「{sheet_name}」の更新ができません。")
        return False
        
    try:
        sheet_id = st.secrets.get("spreadsheet_id", "")
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet(sheet_name)
        
        try:
            cell = sheet.find(str(search_value), in_column=search_col_index)
        except gspread.exceptions.CellNotFound:
            st.error(f"指定されたデータ（{search_value}）がスプレッドシート上で見つかりませんでした。")
            return False
            
        row_num = cell.row
        
        end_col_chr = chr(ord('A') + len(row_data) - 1)
        range_str = f"A{row_num}:{end_col_chr}{row_num}"
        
        try:
            res = sheet.update(range_name=range_str, values=[row_data], value_input_option="USER_ENTERED")
        except TypeError:
            res = sheet.update(range_str, [row_data], value_input_option="USER_ENTERED")
            
        st.cache_data.clear()
        
        return res
    except Exception as e:
        st.error(f"シート「{sheet_name}」の更新処理で予期せぬエラーが発生しました。\n詳細: {e}")
        return False

def overwrite_data(sheet_name, df, columns):
    """
    シートの内容をクリアし、DataFrameの内容で全体を上書きする。（洗い替え保存用）
    """
    client, _ = get_gspread_client()
    if not client:
        st.error(f"DB未接続のため、シート「{sheet_name}」の上書きができません。")
        return False
    try:
        sheet_id = st.secrets.get("spreadsheet_id", "")
        spreadsheet = client.open_by_key(sheet_id)
        sheet = spreadsheet.worksheet(sheet_name)
        
        sheet.clear()
        
        # 欠損値を空文字に置換
        df = df.fillna("")
        
        # ヘッダーとデータを結合
        data = [columns] + df[columns].values.tolist()
        
        try:
            res = sheet.update(range_name="A1", values=data, value_input_option="USER_ENTERED")
        except TypeError:
            res = sheet.update("A1", data, value_input_option="USER_ENTERED")
            
        st.cache_data.clear()
        return res
    except Exception as e:
        st.error(f"シート「{sheet_name}」の上書き処理でエラーが発生しました。\n詳細: {e}")
        return False

# --- カラム定義（スプレッドシートの列名） ---
COLS_STAFF = ["氏名", "役職", "OPE習熟度", "アンギオ習熟度", "総合コード", "人工心肺メイン回数", "人工心肺サブ回数", "アブレーション回数", "カテ回数"]
COLS_REQUEST = ["日時", "氏名", "区分", "コメント"]
COLS_OPE_MASTER = ["術式名", "術式レベル"]
COLS_OPE_SCHEDULE = ["日時", "術式"]
COLS_TASK_MASTER = ["略語", "業務名"]
COLS_SHIFT = ["日時", "氏名", "割り当て業務"]

# --- ページUI コンポーネント ---

def page_home():
    st.markdown('<div class="card"><h2>① ホーム（月間勤務表・シフト）</h2><p>確定した勤務表や術式予定を確認します。</p></div>', unsafe_allow_html=True)
    
    df_shift = fetch_data("確定勤務表", COLS_SHIFT)
    df_ope = fetch_data("術式予定", COLS_OPE_SCHEDULE)
    
    st.markdown('<div class="card">', unsafe_allow_html=True)
    
    # ----------------------------------------------------
    # 年月の選択機能
    # ----------------------------------------------------
    today_date = datetime.date.today()
    years = list(range(today_date.year - 1, today_date.year + 2))
    months = list(range(1, 13))
    
    col1, col2, _ = st.columns([1, 1, 4])
    with col1:
        selected_year = st.selectbox("年を選択", years, index=years.index(today_date.year))
    with col2:
        selected_month = st.selectbox("月を選択", months, index=months.index(today_date.month))
    
    # Pythonのcalendarモジュールを使用してカレンダーを生成（日曜始まり）
    cal = calendar.Calendar(firstweekday=6)
    month_days = cal.monthdatescalendar(selected_year, selected_month)
    
    # ----------------------------------------------------
    # データの前処理 (日付 YYYY-MM-DD をキーにした辞書を作成)
    # ----------------------------------------------------
    shift_dict = {}
    if not df_shift.empty:
        for _, row in df_shift.iterrows():
            d_key = parse_date(row["日時"])
            if d_key:
                staff_info = f"{row['氏名']} ({row['割り当て業務']})"
                if d_key not in shift_dict:
                    shift_dict[d_key] = []
                shift_dict[d_key].append(staff_info)

    ope_dict = {}
    if not df_ope.empty:
        for _, row in df_ope.iterrows():
            d_key = parse_date(row["日時"])
            if d_key:
                ope_info = str(row['術式'])
                if d_key not in ope_dict:
                    ope_dict[d_key] = []
                ope_dict[d_key].append(ope_info)
                
    # ----------------------------------------------------
    # CSS (カレンダー用グリッドデザイン)
    # ----------------------------------------------------
    calendar_css = """
    <style>
        .calendar-container {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 8px;
            margin-top: 15px;
        }
        .calendar-header {
            text-align: center;
            font-weight: bold;
            background-color: #F8F9FA;
            padding: 10px;
            border-radius: 8px;
            color: #495057;
            border: 1px solid #E9ECEF;
        }
        .calendar-day {
            background-color: #FFFFFF;
            border: 1px solid #E9ECEF;
            border-radius: 8px;
            min-height: 140px;
            padding: 8px;
            display: flex;
            flex-direction: column;
            box-shadow: 0 2px 4px rgba(0,0,0,0.02);
            transition: transform 0.1s;
        }
        .calendar-day:hover {
            transform: scale(1.02);
            box-shadow: 0 4px 8px rgba(0,0,0,0.05);
            z-index: 10;
        }
        /* 土曜日の背景（淡い青） */
        .calendar-day.saturday {
            background-color: #F0F8FF;
        }
        /* 日曜・祝日の背景（淡い赤） */
        .calendar-day.holiday {
            background-color: #FFF0F5;
        }
        /* 本日の背景（淡い黄色） */
        .calendar-day.today {
            background-color: #FFF3CD !important;
            border: 2px solid #FFC107 !important;
        }
        /* その他の月 */
        .calendar-day.other-month {
            opacity: 0.5;
            background-color: #F8F9FA;
        }
        
        .day-number {
            font-size: 1.3rem;
            font-weight: 800;
            color: #0056B3;
            text-align: left;
            margin-bottom: 0px;
        }
        /* 土曜・休日の日付色上書き */
        .calendar-day.saturday .day-number { color: #0D6EFD; }
        .calendar-day.holiday .day-number { color: #DC3545; }
        .calendar-day.today .day-number { color: #856404; }
        
        .day-weekday {
            font-size: 0.75rem;
            color: #6C757D;
            text-align: left;
            margin-bottom: 8px;
            border-bottom: 1px solid #E9ECEF;
            padding-bottom: 4px;
            font-weight: bold;
        }
        .day-staff {
            font-size: 0.85rem;
            color: #212529;
            flex-grow: 1;
            white-space: pre-wrap;
            line-height: 1.5;
        }
        .day-ope {
            font-size: 0.8rem;
            color: #D63384;
            font-weight: bold;
            margin-top: 8px;
            background-color: #FFF0F6;
            padding: 4px 6px;
            border-radius: 4px;
            white-space: pre-wrap;
            border-left: 3px solid #D63384;
        }
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
            if is_other_month:
                class_name += " other-month"
            elif is_holiday or is_sunday:
                class_name += " holiday"
            elif is_saturday:
                class_name += " saturday"
                
            if is_today:
                class_name += " today"
            
            d_str = d.strftime("%Y-%m-%d")
            day_num = d.day
            
            holiday_name = jpholiday.is_holiday_name(d)
            # Pythonのd.weekday()は月曜=0。正しいリストを使用。
            weekday_list_mapped = ["月", "火", "水", "木", "金", "土", "日"]
            weekday_str = f"{weekday_list_mapped[d.weekday()]}曜日"
            if holiday_name:
                weekday_str += f" <span style='color:#DC3545;'>({holiday_name})</span>"
                
            staffs = shift_dict.get(d_str, [])
            opes = ope_dict.get(d_str, [])
            
            # 「宿直」を紫色にハイライト
            staff_html_parts = []
            for s in staffs:
                if "宿直" in s:
                    staff_html_parts.append(f'<span style="color: #800080; font-weight: bold;">{s}</span>')
                else:
                    staff_html_parts.append(s)
            
            staff_html = "<br>".join(staff_html_parts) if staff_html_parts else ""
            ope_html = "<br>".join(opes) if opes else ""
            
            html += f'<div class="{class_name}">'
            html += f'<div class="day-number">{day_num}</div>'
            html += f'<div class="day-weekday">{weekday_str}</div>'
            html += f'<div class="day-staff">{staff_html}</div>'
            if ope_html:
                html += f'<div class="day-ope">{ope_html}</div>'
            html += '</div>'
            
    html += '</div>'
    
    st.markdown(html, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    # ----------------------------------------------------
    # 横型シフト表 ＆ 集計表
    # ----------------------------------------------------
    st.markdown('<div class="card">', unsafe_allow_html=True)
    tab_shift, tab_summary = st.tabs(["📋 横型シフト表", "📊 業務割り当て集計表"])
    
    with tab_shift:
        st.write(f"### 📋 {selected_year}年{selected_month}月 横型シフト表")
        
        df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
        staff_list = df_staff["氏名"].tolist() if not df_staff.empty else []
        
        _, num_days = calendar.monthrange(selected_year, selected_month)
        
        # マルチインデックスの列を生成： (日にち, 曜日, 術式)
        cols_tuples = []
        weekday_list_mapped = ["月", "火", "水", "木", "金", "土", "日"]
        
        for d in range(1, num_days + 1):
            dt = datetime.date(selected_year, selected_month, d)
            d_str = dt.strftime("%Y-%m-%d")
            
            day_str = f"{d}日"
            wd = dt.weekday()
            is_hol = jpholiday.is_holiday(dt)
            weekday_str = weekday_list_mapped[wd]
            if is_hol:
                weekday_str += "(祝)"
                
            opes = ope_dict.get(d_str, [])
            ope_str = "\n".join(opes) if opes else ""
            
            cols_tuples.append((day_str, weekday_str, ope_str))
            
        shift_matrix = pd.DataFrame(index=staff_list, columns=pd.MultiIndex.from_tuples(cols_tuples, names=["日にち", "曜日", "術式"]))
        shift_matrix.fillna("", inplace=True)
        
        # 確定勤務表データを埋め込む
        if not df_shift.empty:
            for _, row in df_shift.iterrows():
                d_key = parse_date(row["日時"])
                if d_key:
                    try:
                        dt = datetime.datetime.strptime(d_key, "%Y-%m-%d")
                        if dt.year == selected_year and dt.month == selected_month:
                            staff_name = str(row["氏名"])
                            duty = str(row["割り当て業務"])
                            
                            wd = dt.weekday()
                            is_hol = jpholiday.is_holiday(dt)
                            weekday_str = weekday_list_mapped[wd]
                            if is_hol:
                                weekday_str += "(祝)"
                            
                            opes = ope_dict.get(d_key, [])
                            ope_str = "\n".join(opes) if opes else ""
                            
                            col_key = (f"{dt.day}日", weekday_str, ope_str)
                            
                            if staff_name in shift_matrix.index or staff_name.startswith("スタッフ"):
                                if staff_name not in shift_matrix.index:
                                    shift_matrix.loc[staff_name] = ""
                                curr_val = shift_matrix.at[staff_name, col_key]
                                if curr_val:
                                    shift_matrix.at[staff_name, col_key] = curr_val + "\n" + duty
                                else:
                                    shift_matrix.at[staff_name, col_key] = duty
                    except Exception:
                        pass
        
        # HTML横型シフト表の構築（縦書き・ハイライト対応）
        html_table = """
        <style>
            .shift-table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
            .shift-table th, .shift-table td { border: 1px solid #DEE2E6; padding: 4px; }
            .shift-table thead th { position: sticky; top: 0; z-index: 1; }
        </style>
        <div style="overflow-x: auto; max-width: 100%;">
        <table class="shift-table">
        """
        # Row 1: Date
        html_table += '<thead><tr><th style="background-color: #F1F3F5; min-width: 100px;">日にち</th>'
        for col in cols_tuples:
            date_str, _, _ = col
            is_today = (date_str == f"{today_date.day}日" and selected_year == today_date.year and selected_month == today_date.month)
            bg_color = "#FFF3CD" if is_today else "#F1F3F5"
            html_table += f'<th style="background-color: {bg_color}; text-align: center;">{date_str}</th>'
        html_table += '</tr>'
        
        # Row 2: Weekday
        html_table += '<tr><th style="background-color: #F1F3F5;">曜日</th>'
        for col in cols_tuples:
            date_str, wd_str, _ = col
            bg_color = "#F1F3F5"
            text_color = "#495057"
            is_today = (date_str == f"{today_date.day}日" and selected_year == today_date.year and selected_month == today_date.month)
            if "土" in wd_str:
                bg_color = "#EBF5FB"
                text_color = "#0D6EFD"
            elif "日" in wd_str or "祝" in wd_str:
                bg_color = "#FDEDEC"
                text_color = "#DC3545"
            if is_today:
                bg_color = "#FFF3CD"
            html_table += f'<th style="background-color: {bg_color}; color: {text_color}; text-align: center;">{wd_str}</th>'
        html_table += '</tr>'
        
        # Row 3: Operation (縦書き)
        html_table += '<tr><th style="background-color: #F1F3F5;">術式</th>'
        for col in cols_tuples:
            date_str, wd_str, ope_str = col
            bg_color = "#F8F9FA"
            if "土" in wd_str: bg_color = "#EBF5FB"
            elif "日" in wd_str or "祝" in wd_str: bg_color = "#FDEDEC"
            is_today = (date_str == f"{today_date.day}日" and selected_year == today_date.year and selected_month == today_date.month)
            if is_today: bg_color = "#FFF3CD"
            ope_html = str(ope_str).replace("\n", "<br>")
            html_table += f'<th style="background-color: {bg_color}; vertical-align: top; text-align: center; writing-mode: vertical-rl; text-orientation: upright; padding: 10px 5px; height: 120px; line-height: 1.5;">{ope_html}</th>'
        html_table += '</tr></thead><tbody>'
        
        # Data Rows
        for staff in shift_matrix.index:
            html_table += f'<tr><td style="font-weight: bold; background-color: #F8F9FA; white-space: nowrap;">{staff}</td>'
            for col in cols_tuples:
                val = shift_matrix.at[staff, col]
                date_str, wd_str, _ = col
                bg_color = "#FFFFFF"
                if "土" in wd_str: bg_color = "#F0F8FF"
                elif "日" in wd_str or "祝" in wd_str: bg_color = "#FFF0F5"
                is_today = (date_str == f"{today_date.day}日" and selected_year == today_date.year and selected_month == today_date.month)
                if is_today: bg_color = "#FFFAEB"
                
                val_html = str(val).replace("\n", "<br>")
                # 宿直の文字を紫色にする
                if "宿直" in val_html:
                    val_html = val_html.replace("宿直", '<span style="color: #800080; font-weight: bold;">宿直</span>')
                html_table += f'<td style="background-color: {bg_color}; text-align: center; vertical-align: middle;">{val_html}</td>'
            html_table += '</tr>'
        html_table += '</tbody></table></div>'
        
        # UIにHTMLテーブルを描画
        st.markdown(html_table, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        # ダウンロード用のExcel生成
        try:
            buffer = io.BytesIO()
            try:
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    shift_matrix.to_excel(writer, sheet_name=f"{selected_month}月シフト")
            except Exception:
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    shift_matrix.to_excel(writer, sheet_name=f"{selected_month}月シフト")
            
            st.download_button(
                label="📥 Excel形式でダウンロード (.xlsx)",
                data=buffer.getvalue(),
                file_name=f"shift_{selected_year}_{selected_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception:
            st.warning("Excel出力用ライブラリが見つからないため、CSV形式で出力します。")
            csv = shift_matrix.to_csv(encoding="utf-8-sig")
            st.download_button(
                label="📥 CSV形式でダウンロード (.csv)",
                data=csv,
                file_name=f"shift_{selected_year}_{selected_month}.csv",
                mime="text/csv"
            )

    with tab_summary:
        st.write(f"### 📊 {selected_year}年{selected_month}月 業務割り当て総合回数")
        if not df_shift.empty:
            df_month = df_shift.copy()
            df_month["日時"] = df_month["日時"].apply(parse_date)
            month_prefix = f"{selected_year}-{selected_month:02d}"
            df_month = df_month[df_month["日時"].astype(str).str.startswith(month_prefix)]
            
            if not df_month.empty:
                try:
                    summary_df = pd.pivot_table(
                        df_month, 
                        index="氏名", 
                        columns="割り当て業務", 
                        aggfunc="size", 
                        fill_value=0
                    )
                    summary_df["合計"] = summary_df.sum(axis=1)
                    st.dataframe(summary_df, use_container_width=True)
                except Exception as e:
                    st.warning(f"集計処理中にエラーが発生しました: {e}")
            else:
                st.info(f"{selected_year}年{selected_month}月の勤務表データがありません。")
        else:
            st.info("集計する勤務表データがありません。")
            
    st.markdown('</div>', unsafe_allow_html=True)


def page_schedule_task():
    st.markdown('<div class="card"><h2>② 予定・マスタ管理</h2><p>術式予定、希望休入力、各種マスタの管理を行います。</p></div>', unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4 = st.tabs(["術式予定", "希望休入力", "業務マスタ管理", "術式マスタ管理"])
    
    with tab1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 術式予定の登録")
        with st.form("schedule_form", clear_on_submit=True):
            sched_date = st.date_input("手術日時")
            
            df_ope_master = fetch_data("術式マスタ", COLS_OPE_MASTER)
            ope_names = df_ope_master["術式名"].tolist() if not df_ope_master.empty else ["CABG", "AVR", "PCI", "アブレーション"]
            
            sched_ope = st.selectbox("術式", ope_names)
            
            if st.form_submit_button("予定追加"):
                try:
                    res = append_data("術式予定", [str(sched_date), sched_ope])
                    if res:
                        st.success("スプレッドシートに予定を書き込みました。")
                except Exception as e:
                    st.error(f"予期せぬエラーが発生しました: {e}")
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
            req_type = st.radio("区分", ["× (不可)", "△ (要相談)"])
            req_comment = st.text_input("コメント")
            
            if st.form_submit_button("希望休を登録"):
                val_type = "×" if "×" in req_type else "△"
                try:
                    res = append_data("希望入力", [str(req_date), req_name, val_type, req_comment])
                    if res:
                        st.success("スプレッドシートに希望休を書き込みました。")
                except Exception as e:
                    st.error(f"予期せぬエラーが発生しました: {e}")
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
                    try:
                        res = append_data("業務マスタ", [task_abbr, task_name])
                        if res:
                            st.success("スプレッドシートに業務マスタを書き込みました。")
                    except Exception as e:
                        st.error(f"予期せぬエラーが発生しました: {e}")
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
                        if res:
                            st.success(f"業務マスタ「{task_abbr_upd}」を更新しました。")
            else:
                st.info("データがありません。")
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
                    try:
                        res = append_data("術式マスタ", [ope_name, ope_level])
                        if res:
                            st.success("スプレッドシートに術式マスタを追加しました。")
                    except Exception as e:
                        st.error(f"予期せぬエラー: {e}")
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
                        if res:
                            st.success(f"「{ope_name_upd}」の情報を更新しました。")
            else:
                st.info("データがありません。")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 術式マスタ一覧")
        st.dataframe(fetch_data("術式マスタ", COLS_OPE_MASTER), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)


def page_staff():
    st.markdown('<div class="card"><h2>③ スタッフマスタ管理</h2><p>スタッフの基本情報や各分野の習熟度を管理します。</p></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="card">', unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["新規登録", "情報更新"])
    
    with tab1:
        st.write("### 新規スタッフ登録")
        with st.form("staff_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                staff_name = st.text_input("氏名")
                staff_role = st.selectbox("役職", ["技士長", "副技士長", "主任", "一般", "新人"])
            with col2:
                ope_level = st.selectbox("OPE習熟度", ["A", "B", "C", "D"])
                angio_level = st.selectbox("アンギオ習熟度", ["1", "2", "3", "4"])
            
            st.markdown("---")
            st.write("#### 経験回数・実績")
            col3, col4 = st.columns(2)
            with col3:
                cpb_main = st.number_input("人工心肺 メイン回数", min_value=0, step=1)
                cpb_sub = st.number_input("人工心肺 サブ回数", min_value=0, step=1)
            with col4:
                ablation_count = st.number_input("アブレーション回数", min_value=0, step=1)
                catha_count = st.number_input("カテ回数", min_value=0, step=1)
                
            total_code = f"{ope_level}-{angio_level}"
            st.info(f"💡 OPE・アンギオ習熟度から生成される総合コード: **{total_code}**")
            
            if st.form_submit_button("スタッフ登録"):
                row = [
                    staff_name, staff_role, ope_level, angio_level, total_code, 
                    int(cpb_main), int(cpb_sub), int(ablation_count), int(catha_count)
                ]
                try:
                    res = append_data("スタッフマスタ", row)
                    if res:
                        st.success(f"スプレッドシートに「{staff_name}」さんのデータを新規登録しました。")
                except Exception as e:
                    st.error(f"スタッフ登録処理中にエラーが発生しました。\n詳細: {e}")
                    
    with tab2:
        st.write("### 登録済みスタッフの情報更新")
        df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
        if not df_staff.empty:
            staff_names = df_staff["氏名"].astype(str).tolist()
            selected_staff = st.selectbox("更新するスタッフを選択", staff_names)
            
            # 選択されたスタッフの現在のデータを取得
            curr = df_staff[df_staff["氏名"].astype(str) == selected_staff].iloc[0]
            
            with st.form("staff_update_form"):
                col1, col2 = st.columns(2)
                with col1:
                    # 検索のキーとするため、名前は変更不可（disabled）にする
                    staff_name_upd = st.text_input("氏名（検索キーのため変更不可）", value=curr["氏名"], disabled=True)
                    
                    role_options = ["技士長", "副技士長", "主任", "一般", "新人"]
                    r_idx = role_options.index(curr["役職"]) if curr["役職"] in role_options else 3
                    staff_role_upd = st.selectbox("役職", role_options, index=r_idx)
                with col2:
                    ope_options = ["A", "B", "C", "D"]
                    o_idx = ope_options.index(curr["OPE習熟度"]) if curr["OPE習熟度"] in ope_options else 0
                    ope_level_upd = st.selectbox("OPE習熟度", ope_options, index=o_idx)
                    
                    angio_options = ["1", "2", "3", "4"]
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
                    
                total_code_upd = f"{ope_level_upd}-{angio_level_upd}"
                st.info(f"💡 新しい総合コード: **{total_code_upd}**")
                
                if st.form_submit_button("情報を更新"):
                    row_upd = [
                        staff_name_upd,         # 1: 氏名 (検索キー)
                        staff_role_upd,         # 2: 役職
                        ope_level_upd,          # 3: OPE習熟度
                        angio_level_upd,        # 4: アンギオ習熟度
                        total_code_upd,         # 5: 総合コード
                        int(cpb_main_upd),      # 6: 人工心肺メイン回数
                        int(cpb_sub_upd),       # 7: 人工心肺サブ回数
                        int(ablation_upd),      # 8: アブレーション回数
                        int(catha_upd)          # 9: カテ回数
                    ]
                    
                    try:
                        # 氏名（第1列）を検索キーとして上書き更新処理を実行
                        res = update_data("スタッフマスタ", 1, staff_name_upd, row_upd)
                        if res:
                            st.success(f"スプレッドシートの「{staff_name_upd}」さんの情報を上書き更新しました。")
                    except Exception as e:
                        st.error(f"更新処理中に予期せぬエラーが発生しました。\n詳細: {e}")
        else:
            st.info("登録されているスタッフがいません。")
            
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### スタッフ一覧")
    # キャッシュクリアされているため、ここで最新データが取得され一覧に反映される
    df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
    st.dataframe(df_staff, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)


def page_shift_creation():
    st.markdown('<div class="card"><h2>④ 勤務表作成</h2><p>勤務表の「1ヶ月一括作成」と、作成後の「日ごとの微調整」を行います。</p></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 機能A：1ヶ月一括ドラフト作成（ベース作り）")
    st.info("指定した月の1日〜月末までのシフトを自動生成し、スプレッドシートに保存します。\n※既に該当月のデータがある場合は、対象月のデータがすべて「洗い替え（上書き）」されます。")
    
    with st.form("bulk_draft_form"):
        col1, col2 = st.columns(2)
        today = datetime.date.today()
        with col1:
            target_year = st.number_input("対象年", min_value=2020, max_value=2050, value=today.year, step=1)
        with col2:
            target_month = st.number_input("対象月", min_value=1, max_value=12, value=today.month, step=1)
            
        if st.form_submit_button("1ヶ月分の一括ドラフトを作成"):
            with st.spinner(f"{target_year}年{target_month}月のドラフトを作成中..."):
                df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
                df_request = fetch_data("希望入力", COLS_REQUEST)
                
                staff_list = df_staff["氏名"].tolist() if not df_staff.empty else []
                
                # 希望休の抽出（日をキーにした×スタッフのリストを作成）
                req_dict = {}
                if not df_request.empty:
                    for _, row in df_request.iterrows():
                        if "×" in str(row["区分"]):
                            r_date = parse_date(row["日時"])
                            if r_date:
                                if r_date not in req_dict:
                                    req_dict[r_date] = []
                                req_dict[r_date].append(str(row["氏名"]))
                                
                _, num_days = calendar.monthrange(target_year, target_month)
                draft_data = []
                
                # 均等化アルゴリズム用：スタッフごとの各業務カウントを保持
                staff_task_counts = {s: {} for s in staff_list}
                
                # ----------------------------------------------------
                # 年間・累積の均等化を実現するため、過去の確定勤務表データを読み込む
                # ----------------------------------------------------
                df_history = fetch_data("確定勤務表", COLS_SHIFT)
                if not df_history.empty:
                    df_history["_date_str"] = df_history["日時"].apply(parse_date)
                    for _, row in df_history.iterrows():
                        if pd.isna(row["_date_str"]) or not row["_date_str"]: continue
                        h_date = pd.to_datetime(row["_date_str"])
                        # 今作成しようとしている月より「前」のデータを累積カウントとして初期値に設定
                        if h_date.year < target_year or (h_date.year == target_year and h_date.month < target_month):
                            s_name = str(row["氏名"])
                            t_name = str(row["割り当て業務"])
                            if s_name in staff_task_counts:
                                staff_task_counts[s_name][t_name] = staff_task_counts[s_name].get(t_name, 0) + 1
                                
                def get_count(staff, task):
                    return staff_task_counts.get(staff, {}).get(task, 0)
                    
                def increment_count(staff, task):
                    if staff not in staff_task_counts:
                        staff_task_counts[staff] = {}
                    staff_task_counts[staff][task] = staff_task_counts[staff].get(task, 0) + 1
                
                # 1日から月末までループ処理
                for d in range(1, num_days + 1):
                    dt = datetime.date(target_year, target_month, d)
                    d_str = dt.strftime("%Y-%m-%d")
                    
                    is_weekend = dt.weekday() >= 5
                    is_holiday = jpholiday.is_holiday(dt)
                    is_new_year = (dt.month == 12 and dt.day >= 29) or (dt.month == 1 and dt.day <= 3)
                    is_off_day = is_weekend or is_holiday or is_new_year
                    
                    # 曜日・休日ごとの必要業務ルール
                    required_tasks = []
                    if is_off_day:
                        # 土日祝・年末年始は日勤1名のみ（あとでこの人が宿直も兼ねる）
                        required_tasks = ["日勤"]
                    else:
                        # 平日ベース（宿直は後で日勤者から選ぶためここには含めない）
                        required_tasks = ["カ", "Ｉ", "Ｏ", "Ｍ", "Ｄ", "Ｒ"]
                        if dt.weekday() == 0: # 月曜日
                            required_tasks.extend(["ＨＭ", "Ｈサ"])
                        elif dt.weekday() == 3: # 木曜日
                            required_tasks.extend(["ＨＭ", "Ｈサ", "Ａ", "Ａ"])
                            
                    # 本日出勤可能なスタッフ
                    unavailable = req_dict.get(d_str, [])
                    available_staff = [s for s in staff_list if s not in unavailable]
                    
                    # 1. まず日中帯の通常業務（日勤等）を割り当てる
                    assigned_today_staffs = []
                    dummy_counter = 0
                    for task in required_tasks:
                        if available_staff:
                            # 割り当て回数が同じ人がいる場合にランダムになるよう事前にシャッフル
                            random.shuffle(available_staff)
                            # 過去からの該当業務の累積回数が少ない順にソート（昇順）
                            available_staff.sort(key=lambda s: get_count(s, task))
                            
                            # 最も回数が少ないスタッフをアサイン
                            assigned_staff = available_staff.pop(0)
                            increment_count(assigned_staff, task)
                        else:
                            # スタッフが枯渇した場合はダミー
                            assigned_staff = f"スタッフ{chr(65 + dummy_counter)}"
                            dummy_counter += 1
                            
                        draft_data.append([d_str, assigned_staff, task])
                        assigned_today_staffs.append(assigned_staff)
                        
                    # 2. 宿直の割り当て（日勤者の中から重複割り当て）
                    if assigned_today_staffs:
                        # ダミースタッフを除外して、実在のスタッフから候補を絞る
                        real_assigned = [s for s in assigned_today_staffs if s in staff_list]
                        candidates = real_assigned if real_assigned else assigned_today_staffs
                        
                        random.shuffle(candidates)
                        # 宿直の累積回数が少ない人を優先してソート
                        candidates.sort(key=lambda s: get_count(s, "宿直"))
                        
                        night_staff = candidates[0]
                        increment_count(night_staff, "宿直")
                        
                        # 同一人物・同一日付で「宿直」行を追加
                        draft_data.append([d_str, night_staff, "宿直"])
                        
                if draft_data:
                    # 既存の確定勤務表データを取得し、該当月のデータを洗い替えする
                    df_shift = fetch_data("確定勤務表", COLS_SHIFT)
                    month_prefix = f"{target_year}-{target_month:02d}"
                    
                    if not df_shift.empty:
                        if "_date_str" not in df_shift.columns:
                            df_shift["_date_str"] = df_shift["日時"].apply(parse_date)
                        # 対象月以外のデータを残す
                        df_remain = df_shift[~df_shift["_date_str"].astype(str).str.startswith(month_prefix)].drop(columns=["_date_str"], errors="ignore")
                    else:
                        df_remain = pd.DataFrame(columns=COLS_SHIFT)
                        
                    # 生成したドラフトデータをDataFrame化して結合
                    df_draft_month = pd.DataFrame(draft_data, columns=COLS_SHIFT)
                    df_new = pd.concat([df_remain, df_draft_month], ignore_index=True)
                    
                    res = overwrite_data("確定勤務表", df_new, COLS_SHIFT)
                    if res:
                        st.success(f"{target_year}年{target_month}月の1ヶ月分ドラフトを一括作成し、保存しました！")
                else:
                    st.warning("生成対象のデータがありませんでした。")
                    
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 機能B：日ごとの微調整・変更（手動調整）")
    st.info("一括作成されたデータなどを呼び出し、担当者の変更やダミー枠の修正を行います。")
    
    target_date = st.date_input("編集する日付を選択", value=datetime.date.today())
    d_str = target_date.strftime("%Y-%m-%d")
    
    df_shift = fetch_data("確定勤務表", COLS_SHIFT)
    if not df_shift.empty:
        df_shift["_date_str"] = df_shift["日時"].apply(parse_date)
        df_day = df_shift[df_shift["_date_str"] == d_str].drop(columns=["_date_str"])
    else:
        df_day = pd.DataFrame(columns=COLS_SHIFT)
        df_shift = pd.DataFrame(columns=COLS_SHIFT)
        df_shift["_date_str"] = []
        
    st.write(f"#### 📝 {d_str} の勤務データ編集")
    if df_day.empty:
        st.warning("この日のデータはまだありません。機能Aで一括作成するか、行を追加して作成してください。")
        
    # データエディタでインタラクティブに編集
    edited_df = st.data_editor(
        df_day if not df_day.empty else pd.DataFrame(columns=COLS_SHIFT),
        num_rows="dynamic",
        use_container_width=True,
        key="daily_editor"
    )
    
    if st.button("この日の勤務表を上書き保存"):
        with st.spinner("スプレッドシートに保存中..."):
            # 対象日の古いデータを削除 (洗い替え)
            if not df_shift.empty:
                df_remain = df_shift[df_shift["_date_str"] != d_str].drop(columns=["_date_str"])
            else:
                df_remain = pd.DataFrame(columns=COLS_SHIFT)
                
            # 編集後のデータを結合
            df_new = pd.concat([df_remain, edited_df], ignore_index=True)
            
            # 確定勤務表を上書き保存
            res = overwrite_data("確定勤務表", df_new, COLS_SHIFT)
            
            if res:
                st.success(f"{d_str} の勤務データを洗い替え保存しました！")
                
    st.markdown('</div>', unsafe_allow_html=True)


# --- メイン処理（サイドバー・ナビゲーション） ---
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
    
    # --- 接続ステータスの表示とエラーハンドリング ---
    st.sidebar.caption("システムステータス:")
    
    client, error_msg = get_gspread_client()
    
    if client is None:
        st.sidebar.error("🔴 DB未接続 (認証エラー)")
        st.sidebar.error(f"エラー詳細:\n{error_msg}")
        st.sidebar.info("現在はプロトタイプ（UIデモ）として動作しています。")
    else:
        if "spreadsheet_id" not in st.secrets:
            st.sidebar.error("🔴 DB未接続 (ID未設定)")
            st.sidebar.error("エラー詳細:\nst.secrets に 'spreadsheet_id' が設定されていません。")
        else:
            try:
                sheet_id = st.secrets["spreadsheet_id"]
                client.open_by_key(sheet_id)
                st.sidebar.success("🟢 DB接続済 (Google Sheets)")
            except Exception as e:
                st.sidebar.error("🔴 DB未接続 (アクセス失敗)")
                st.sidebar.error(f"エラー詳細:\nスプレッドシートが開けません。IDが間違っているか、サービスアカウントに共有されていません。\nException: {e}")

    # 選択されたページを描画
    pages[selection]()

if __name__ == "__main__":
    main()
