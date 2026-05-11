import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime

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
        
        # 指定列から値を検索 (1列目の場合は in_column=1)
        try:
            cell = sheet.find(str(search_value), in_column=search_col_index)
        except gspread.exceptions.CellNotFound:
            st.error(f"指定されたデータ（{search_value}）がスプレッドシート上で見つかりませんでした。")
            return False
            
        row_num = cell.row
        
        # 列の終端文字を計算 (A〜Zを想定、列数が最大26までの簡易計算)
        end_col_chr = chr(ord('A') + len(row_data) - 1)
        range_str = f"A{row_num}:{end_col_chr}{row_num}"
        
        # 更新処理 (gspreadのバージョン互換を考慮して try-except を使用)
        try:
            res = sheet.update(range_name=range_str, values=[row_data], value_input_option="USER_ENTERED")
        except TypeError:
            # 古いgspread向けの書き方
            res = sheet.update(range_str, [row_data], value_input_option="USER_ENTERED")
            
        # キャッシュをクリアして次回fetch時に最新状態を反映
        st.cache_data.clear()
        
        return res
    except Exception as e:
        st.error(f"シート「{sheet_name}」の更新処理で予期せぬエラーが発生しました。\n詳細: {e}")
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
    st.markdown('<div class="card"><h2>① ホーム（確定勤務表の確認）</h2><p>確定した勤務表や本日のシフト、術式予定を確認します。</p></div>', unsafe_allow_html=True)
    
    df_shift = fetch_data("確定勤務表", COLS_SHIFT)
    df_ope = fetch_data("術式予定", COLS_OPE_SCHEDULE)
    
    st.markdown('<div class="card">', unsafe_allow_html=True)
    col1, col2 = st.columns([1, 3])
    with col1:
        date_filter = st.date_input("表示日を選択", datetime.date.today())
    
    date_str = str(date_filter)
    
    st.write("### 📅 確定勤務表")
    if not df_shift.empty:
        filtered_shift = df_shift[df_shift["日時"].astype(str).str.contains(date_str)]
        if not filtered_shift.empty:
            st.dataframe(filtered_shift, use_container_width=True)
        else:
            st.info(f"{date_str} の勤務表データがありません。")
    else:
        st.info("確定された勤務表データがありません。")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 🏥 本日の術式予定")
    if not df_ope.empty:
        filtered_ope = df_ope[df_ope["日時"].astype(str).str.contains(date_str)]
        if not filtered_ope.empty:
            st.dataframe(filtered_ope, use_container_width=True)
        else:
            st.info(f"{date_str} の術式予定はありません。")
    else:
        st.info("術式予定データがありません。")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 📊 業務割り当て総合回数（集計表）")
    if not df_shift.empty:
        try:
            summary_df = pd.pivot_table(
                df_shift, 
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
        st.info("集計する勤務表データがありません。")
    st.markdown('</div>', unsafe_allow_html=True)


def page_schedule_task():
    st.markdown('<div class="card"><h2>② 予定・各種マスタ管理</h2><p>術式予定、希望入力の追加、および各マスタの管理（追加・更新）を行います。</p></div>', unsafe_allow_html=True)
    
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
    df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
    st.dataframe(df_staff, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)


# --- メイン処理（サイドバー・ナビゲーション） ---
def main():
    st.sidebar.markdown("<h2>🏥 ME勤務表管理</h2>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    
    pages = {
        "① ホーム（確定勤務表の確認）": page_home,
        "② 予定・業務管理": page_schedule_task,
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
