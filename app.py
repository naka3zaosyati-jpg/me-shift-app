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
    today = datetime.date.today()
    years = list(range(today.year - 1, today.year + 2))
    months = list(range(1, 13))
    
    col1, col2, _ = st.columns([1, 1, 4])
    with col1:
        selected_year = st.selectbox("年を選択", years, index=years.index(today.year))
    with col2:
        selected_month = st.selectbox("月を選択", months, index=months.index(today.month))
    
    # Pythonのcalendarモジュールを使用してカレンダーを生成（日曜始まり）
    cal = calendar.Calendar(firstweekday=6)
    month_days = cal.monthdatescalendar(selected_year, selected_month)
    
    # ----------------------------------------------------
    # データの前処理 (日付 YYYY-MM-DD をキーにした辞書を作成)
    # バグ修正: parse_date関数を用いて表記揺れを吸収
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
    # CSS (カレンダー用グリッドデザイン・休日対応版)
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
            flex-grow: 1; /* 余白を埋めることで下の day-ope を押し下げる */
            white-space: pre-wrap;
            line-height: 1.5;
        }
        .day-ope {
            font-size: 0.8rem;
            color: #D63384;
            font-weight: bold;
            margin-top: 8px; /* 一番下に配置させるためのマージン */
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
    
    # HTML生成: ヘッダー部分
    html = '<div class="calendar-container">'
    for wd in weekdays:
        color_style = ""
        if wd == "日": color_style = "color: #DC3545;"
        elif wd == "土": color_style = "color: #0D6EFD;"
        html += f'<div class="calendar-header" style="{color_style}">{wd}</div>'
    
    # HTML生成: 日付のマス部分
    for week in month_days:
        for d in week:
            is_other_month = d.month != selected_month
            is_saturday = d.weekday() == 5
            is_sunday = d.weekday() == 6
            is_holiday = jpholiday.is_holiday(d)
            
            class_name = "calendar-day"
            if is_other_month:
                class_name += " other-month"
            elif is_holiday or is_sunday:
                class_name += " holiday"
            elif is_saturday:
                class_name += " saturday"
            
            d_str = d.strftime("%Y-%m-%d")
            day_num = d.day
            
            # 曜日テキストと祝日名の取得
            weekday_list = ["月", "火", "水", "木", "金", "土", "日"]
            holiday_name = jpholiday.is_holiday_name(d)
            weekday_str = f"{weekday_list[d.weekday()]}曜日"
            if holiday_name:
                weekday_str += f" <span style='color:#DC3545;'>({holiday_name})</span>"
                
            staffs = shift_dict.get(d_str, [])
            opes = ope_dict.get(d_str, [])
            
            staff_html = "<br>".join(staffs) if staffs else ""
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
    # 横型シフト表 ＆ 集計表 (タブで切り替え表示)
    # ----------------------------------------------------
    st.markdown('<div class="card">', unsafe_allow_html=True)
    tab_shift, tab_summary = st.tabs(["📋 横型シフト表", "📊 業務割り当て集計表"])
    
    with tab_shift:
        st.write(f"### 📋 {selected_year}年{selected_month}月 横型シフト表")
        
        # スタッフ一覧を取得（縦軸のインデックスとして使用）
        df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
        staff_list = df_staff["氏名"].tolist() if not df_staff.empty else []
        
        # 選択された月の日数分の列を作成（例: '1日', '2日', ...）
        _, num_days = calendar.monthrange(selected_year, selected_month)
        days_cols = [f"{d}日" for d in range(1, num_days + 1)]
        
        # 空のクロス集計表(DataFrame)を生成
        shift_matrix = pd.DataFrame(index=staff_list, columns=days_cols)
        shift_matrix.fillna("", inplace=True)
        
        # 確定勤務表データを埋め込む
        if not df_shift.empty:
            for _, row in df_shift.iterrows():
                d_key = parse_date(row["日時"])
                if d_key:
                    try:
                        dt = datetime.datetime.strptime(d_key, "%Y-%m-%d")
                        # 選択された年月に該当するデータのみ処理
                        if dt.year == selected_year and dt.month == selected_month:
                            staff_name = str(row["氏名"])
                            duty = str(row["割り当て業務"])
                            if staff_name in shift_matrix.index:
                                day_str = f"{dt.day}日"
                                curr_val = shift_matrix.at[staff_name, day_str]
                                # 同じ日に複数業務がある場合は改行して追記
                                if curr_val:
                                    shift_matrix.at[staff_name, day_str] = curr_val + "\n" + duty
                                else:
                                    shift_matrix.at[staff_name, day_str] = duty
                            # ダミースタッフ等マスタに無いものは末尾に追記
                            elif staff_name.startswith("スタッフ"):
                                if staff_name not in shift_matrix.index:
                                    shift_matrix.loc[staff_name] = ""
                                day_str = f"{dt.day}日"
                                curr_val = shift_matrix.at[staff_name, day_str]
                                if curr_val:
                                    shift_matrix.at[staff_name, day_str] = curr_val + "\n" + duty
                                else:
                                    shift_matrix.at[staff_name, day_str] = duty
                    except Exception:
                        pass
                        
        st.dataframe(shift_matrix, use_container_width=True)
        
        # Excelダウンロード機能
        try:
            buffer = io.BytesIO()
            # xlsxwriter がインストールされているか確認しつつ書き込み
            try:
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    shift_matrix.to_excel(writer, sheet_name=f"{selected_month}月シフト")
            except Exception:
                # 無い場合は openpyxl で試行
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    shift_matrix.to_excel(writer, sheet_name=f"{selected_month}月シフト")
            
            st.download_button(
                label="📥 Excel形式でダウンロード (.xlsx)",
                data=buffer.getvalue(),
                file_name=f"shift_{selected_year}_{selected_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.warning("Excel出力用ライブラリ(xlsxwriter/openpyxl)が見つからないため、代替としてCSV形式で出力します。")
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
            # parse_date を適用してフォーマットを統一
            df_month["日時"] = df_month["日時"].apply(parse_date)
            # 選択年月のデータのみに絞り込む
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
    st.markdown('<div class="card"><h2>② 予定・マスタ・自動ドラフト管理</h2><p>術式予定、希望休入力、各種マスタの管理、およびシフト自動生成を行います。</p></div>', unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["術式予定", "希望休入力", "業務マスタ管理", "術式マスタ管理", "✨ 日別シフト作成"])
    
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

    with tab5:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### ✨ 日別シフト自動割り当て（ドラフト作成）")
        st.info("日付を選択してドラフトを作成後、画面上で微調整し、「確定して保存」ボタンでスプレッドシートに反映します。\n※既にその日のデータがある場合、確定時に「洗い替え（上書き）」されます。")
        
        target_date = st.date_input("対象日を選択", value=datetime.date.today())
        
        if st.button("ドラフトを作成する"):
            with st.spinner("ドラフトを作成中..."):
                df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
                df_request = fetch_data("希望入力", COLS_REQUEST)
                df_task = fetch_data("業務マスタ", COLS_TASK_MASTER)
                
                staff_list = df_staff["氏名"].tolist() if not df_staff.empty else []
                task_abbrs = df_task["略語"].tolist() if not df_task.empty else []
                
                d_str = target_date.strftime("%Y-%m-%d")
                
                # 希望休の抽出（対象日の「×」のスタッフ）
                unavailable = []
                if not df_request.empty:
                    for _, row in df_request.iterrows():
                        if "×" in str(row["区分"]):
                            if parse_date(row["日時"]) == d_str:
                                unavailable.append(str(row["氏名"]))
                                
                # 休日・特別期間の判定
                is_weekend = target_date.weekday() >= 5
                is_holiday = jpholiday.is_holiday(target_date)
                is_new_year = (target_date.month == 12 and target_date.day >= 29) or (target_date.month == 1 and target_date.day <= 3)
                is_off_day = is_weekend or is_holiday or is_new_year
                
                # 必要業務枠の算出
                required_tasks = []
                if is_off_day:
                    required_tasks = ["ME業務"]
                else:
                    base_tasks = [t for t in task_abbrs if t not in ["ヘルツ", "アブレーション"]]
                    required_tasks.extend(base_tasks)
                    if target_date.weekday() in [0, 3]: # 月・木
                        required_tasks.extend(["ヘルツ", "ヘルツ"])
                    if target_date.weekday() == 3: # 木
                        required_tasks.extend(["アブレーション", "アブレーション"])
                        
                # 出勤可能スタッフ
                available_staff = [s for s in staff_list if s not in unavailable]
                random.shuffle(available_staff)
                
                # 割り当ての実行
                draft_data = []
                assigned_count = 0
                for task in required_tasks:
                    if assigned_count < len(available_staff):
                        assigned_staff = available_staff[assigned_count]
                    else:
                        dummy_idx = assigned_count - len(available_staff)
                        assigned_staff = f"スタッフ{chr(65 + dummy_idx)}"
                    
                    draft_data.append({"日時": d_str, "氏名": assigned_staff, "割り当て業務": task})
                    assigned_count += 1
                
                # セッションステートに保存
                st.session_state["draft_df"] = pd.DataFrame(draft_data, columns=COLS_SHIFT)
                st.session_state["target_date"] = target_date
                
        # ドラフトデータの編集と確定UI
        if "draft_df" in st.session_state and st.session_state.get("target_date") == target_date:
            st.markdown("---")
            st.write(f"#### 📝 {target_date.strftime('%Y-%m-%d')} のドラフト編集")
            st.caption("※表のセルをクリックして氏名や業務を直接修正できます。行の追加・削除も可能です。")
            
            # データエディタでインタラクティブに編集
            edited_df = st.data_editor(
                st.session_state["draft_df"],
                num_rows="dynamic",
                use_container_width=True,
                key="shift_data_editor"
            )
            
            if st.button("確定して保存"):
                with st.spinner("スプレッドシートに保存中..."):
                    # 現在の確定勤務表を全取得
                    df_shift = fetch_data("確定勤務表", COLS_SHIFT)
                    d_str = target_date.strftime("%Y-%m-%d")
                    
                    # 対象日の古いデータを削除 (洗い替え)
                    if not df_shift.empty:
                        df_shift["_date_str"] = df_shift["日時"].apply(parse_date)
                        df_remain = df_shift[df_shift["_date_str"] != d_str].drop(columns=["_date_str"])
                    else:
                        df_remain = pd.DataFrame(columns=COLS_SHIFT)
                        
                    # 編集後のデータを結合
                    df_new = pd.concat([df_remain, edited_df], ignore_index=True)
                    
                    # 確定勤務表を上書き保存
                    res = overwrite_data("確定勤務表", df_new, COLS_SHIFT)
                    
                    if res:
                        st.success(f"{d_str} の勤務データを確定し、スプレッドシートに保存しました！")
                        # 完了後にセッションステートをクリアしてエディタを隠す
                        del st.session_state["draft_df"]
                        
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


# --- メイン処理（サイドバー・ナビゲーション） ---
def main():
    st.sidebar.markdown("<h2>🏥 ME勤務表管理</h2>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    
    pages = {
        "① ホーム（月間勤務表・シフト）": page_home,
        "② 予定・マスタ・自動ドラフト管理": page_schedule_task,
        "③ スタッフマスタ管理": page_staff
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
