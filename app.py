ort streamlit as st
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

# --- カスタムCSS（ダークモード＆ミニマルカードUI） ---
st.markdown("""
<style>
    /* 全体的なダークテーマと高コントラスト */
    .stApp {
        background-color: #121212;
        color: #E0E0E0;
    }
    
    /* 見出しの色を高コントラストなアクセントカラーに */
    h1, h2, h3 {
        color: #BB86FC !important;
    }
    
    /* カード型UIの定義 (スマホ操作を意識したミニマルデザイン) */
    .card {
        background-color: #1E1E1E;
        padding: 1.5rem;
        border-radius: 16px;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.4);
        border: 1px solid #333333;
        margin-bottom: 1.5rem;
    }
    
    /* ボタンのスタイル (タップしやすいサイズ) */
    .stButton>button {
        background-color: #03DAC6;
        color: #000000;
        font-weight: bold;
        border-radius: 12px;
        border: none;
        width: 100%;
        padding: 0.75rem;
        transition: all 0.2s;
    }
    .stButton>button:hover {
        background-color: #00E676;
        color: #000000;
        transform: translateY(-2px);
    }

    /* データフレーム（表）のヘッダー色 */
    th {
        background-color: #2D2D2D !important;
        color: #BB86FC !important;
    }
    
    /* モバイル向けに入力フィールド等を調整 */
    input, select, textarea {
        border-radius: 8px !important;
    }
</style>
""", unsafe_allow_html=True)

# --- Google Sheets API 接続 ---
@st.cache_resource
def get_gspread_client():
    try:
        # st.secretsからGCPサービスアカウント情報を取得
        if "gcp_service_account" in st.secrets:
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            credentials = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=scopes
            )
            return gspread.authorize(credentials)
        else:
            return None
    except Exception as e:
        st.error(f"認証エラー: {e}")
        return None

def get_sheet(sheet_name):
    client = get_gspread_client()
    if not client:
        return None
    try:
        # スプレッドシートIDをsecretsから取得
        sheet_id = st.secrets.get("spreadsheet_id", "")
        if not sheet_id:
            return None
        spreadsheet = client.open_by_key(sheet_id)
        return spreadsheet.worksheet(sheet_name)
    except Exception as e:
        # シートが存在しない場合などはNoneを返す
        return None

# --- CRUD ラッパー関数 ---
def fetch_data(sheet_name, expected_columns):
    sheet = get_sheet(sheet_name)
    if sheet:
        records = sheet.get_all_records()
        if records:
            return pd.DataFrame(records)
        else:
            return pd.DataFrame(columns=expected_columns)
    else:
        # API未設定時のプロトタイプ用モックデータ
        return pd.DataFrame(columns=expected_columns)

def append_data(sheet_name, row_data):
    sheet = get_sheet(sheet_name)
    if sheet:
        sheet.append_row(row_data)
        return True
    return False

# --- カラム定義（スプレッドシートの列名） ---
COLS_STAFF = ["氏名", "役職", "OPE習熟度", "アンギオ習熟度", "総合コード", "人工心肺メイン回数", "人工心肺サブ回数", "アブレーション回数", "カテ回数"]
COLS_REQUEST = ["日時", "氏名", "区分", "コメント"]
COLS_OPE_MASTER = ["術式名", "デフォルト必要レベル"]
COLS_OPE_SCHEDULE = ["日時", "術式", "必要レベル"]
COLS_TASK_MASTER = ["略語", "業務名"]
COLS_SHIFT = ["日時", "氏名", "割当業務"]

# --- ページUI コンポーネント ---

def page_home():
    st.markdown('<div class="card"><h2>① ホーム（シフト確認）</h2><p>確定した勤務表や本日のシフトを確認します。</p></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 確定勤務表")
    df_shift = fetch_data("確定勤務表", COLS_SHIFT)
    
    if not df_shift.empty:
        date_filter = st.date_input("表示日を選択", datetime.date.today())
        st.dataframe(df_shift, use_container_width=True)
    else:
        st.info("確定された勤務表データがありません。")
    st.markdown('</div>', unsafe_allow_html=True)


def page_request():
    st.markdown('<div class="card"><h2>② 希望休入力</h2><p>休みの希望（×、△）を入力します。</p></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="card">', unsafe_allow_html=True)
    with st.form("request_form"):
        req_date = st.date_input("希望日時")
        
        # スタッフリストを取得（未接続時はモック名）
        df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
        staff_names = df_staff["氏名"].tolist() if not df_staff.empty else ["テスト太郎", "テスト花子"]
        
        req_name = st.selectbox("氏名", staff_names)
        req_type = st.radio("区分", ["× (不可)", "△ (要相談)"])
        req_comment = st.text_input("コメント")
        
        if st.form_submit_button("希望休を登録"):
            val_type = "×" if "×" in req_type else "△"
            if append_data("希望入力", [str(req_date), req_name, val_type, req_comment]):
                st.success("希望休を登録しました。")
            else:
                st.warning("登録をシミュレートしました（API未接続）。")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 登録済み希望一覧")
    st.dataframe(fetch_data("希望入力", COLS_REQUEST), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)


def page_schedule_task():
    st.markdown('<div class="card"><h2>③ 予定・業務管理</h2><p>術式の予定入力および業務マスタの管理を行います。</p></div>', unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["術式予定", "業務マスタ管理"])
    
    with tab1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 術式予定の登録")
        with st.form("schedule_form"):
            sched_date = st.date_input("手術日時")
            
            df_ope_master = fetch_data("術式マスタ", COLS_OPE_MASTER)
            ope_names = df_ope_master["術式名"].tolist() if not df_ope_master.empty else ["CABG", "AVR", "PCI", "アブレーション"]
            
            sched_ope = st.selectbox("術式", ope_names)
            sched_level = st.selectbox("必要レベル", ["A", "B", "C", "D"])
            
            if st.form_submit_button("予定追加"):
                if append_data("術式予定", [str(sched_date), sched_ope, sched_level]):
                    st.success("予定を追加しました。")
                else:
                    st.warning("予定追加をシミュレートしました。")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 登録済み予定")
        st.dataframe(fetch_data("術式予定", COLS_OPE_SCHEDULE), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with tab2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 業務マスタ登録")
        with st.form("task_form"):
            task_abbr = st.text_input("略語 (例: HM, A)")
            task_name = st.text_input("業務名 (例: 血液浄化, 血管造影)")
            if st.form_submit_button("マスタ追加"):
                if append_data("業務マスタ", [task_abbr, task_name]):
                    st.success("業務マスタを追加しました。")
                else:
                    st.warning("追加をシミュレートしました。")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 業務マスタ一覧")
        st.dataframe(fetch_data("業務マスタ", COLS_TASK_MASTER), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)


def page_staff():
    st.markdown('<div class="card"><h2>④ スタッフマスタ</h2><p>スタッフの基本情報や各分野の習熟度を管理します。</p></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 新規スタッフ登録 / 習熟度更新")
    with st.form("staff_form"):
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
            
        # 総合コードの自動生成 (例: A-1)
        total_code = f"{ope_level}-{angio_level}"
        
        st.info(f"💡 OPE・アンギオ習熟度から生成される総合コード: **{total_code}**")
        
        if st.form_submit_button("スタッフ登録"):
            row = [
                staff_name, staff_role, ope_level, angio_level, total_code, 
                cpb_main, cpb_sub, ablation_count, catha_count
            ]
            if append_data("スタッフマスタ", row):
                st.success(f"{staff_name}さんを登録しました。")
            else:
                st.warning("登録をシミュレートしました（API未接続）。")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### スタッフ一覧")
    df_staff = fetch_data("スタッフマスタ", COLS_STAFF)
    st.dataframe(df_staff, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)


# --- メイン処理（サイドバー・ナビゲーション） ---
def main():
    # スマホでも見やすいようにサイドバーのタイトルを大きめに
    st.sidebar.markdown("<h2>🏥 ME勤務表管理</h2>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    
    pages = {
        "① ホーム（シフト確認）": page_home,
        "② 希望休入力": page_request,
        "③ 予定・業務管理": page_schedule_task,
        "④ スタッフマスタ": page_staff
    }
    
    selection = st.sidebar.radio("メニュー", list(pages.keys()))
    
    st.sidebar.markdown("---")
    
    # 接続ステータスの表示
    st.sidebar.caption("システムステータス:")
    if get_gspread_client() is None:
        st.sidebar.error("🔴 DB未接続\n\n(st.secrets を設定してください)")
        st.sidebar.info("現在はプロトタイプ（UIデモ）として動作しています。")
    else:
        st.sidebar.success("🟢 DB接続済 (Google Sheets)")
        
    # 選択されたページを描画
    pages[selection]()

if __name__ == "__main__":
    main()
