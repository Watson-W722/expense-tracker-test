import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date, timedelta, timezone
import time
import os
import hashlib

# --- 頁面設定 ---
st.set_page_config(page_title="我的記帳本 Pro", layout="wide", page_icon="💰")

# ==========================================
# [設定區]
# ==========================================
TEMPLATE_URL = "https://docs.google.com/spreadsheets/d/1j7WM4A6bgRr1S-0BvHYPw9Xp5oXs0Ikp969-Ys65JL0/copy" 
TRIAL_DAYS = 30 

# ==========================================
# 0. UI 美化
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Noto Sans TC', sans-serif;
        background-color: #f8f9fa;
        color: #2c3e50;
    }
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 5rem !important;
    }
    #MainMenu {visibility: hidden;}
    .metric-container {
        display: flex;
        flex-wrap: wrap;
        gap: 15px;
        margin: 10px 0 20px 0;
    }
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 15px 20px;
        flex: 1;
        min-width: 140px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        border: 1px solid #eef0f2;
        display: flex;
        flex-direction: column;
        align-items: flex-start;
    }
    .metric-label { font-size: 0.85rem; color: #888; font-weight: 500; margin-bottom: 5px; }
    .metric-value { font-size: 1.6rem; font-weight: 700; color: #2c3e50; }
    .val-green { color: #2ecc71; }
    .val-red { color: #e74c3c; }
    div.stButton > button { border-radius: 8px; font-weight: 600; }
    .stTabs { position: sticky; top: 0; background-color: #f8f9fa; z-index: 999; padding-top: 10px; margin-top: -20px; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { background-color: white; border-radius: 8px 8px 0 0; border: 1px solid #dee2e6; border-bottom: none; }
    .stTabs [aria-selected="true"] { border-top: 3px solid #0d6efd; color: #0d6efd !important; }
    .login-container { max-width: 500px; margin: 50px auto; padding: 40px; background: white; border-radius: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); text-align: center; }
    .step-text { text-align: left; margin-bottom: 10px; font-size: 0.95rem; }
    .vip-badge { background-color: #FFD700; color: #000; padding: 2px 8px; border-radius: 10px; font-size: 0.8em; font-weight: bold; }
    .trial-badge { background-color: #87CEEB; color: #000; padding: 2px 8px; border-radius: 10px; font-size: 0.8em; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. 核心連線模組 (含金鑰自動修復)
# ==========================================
@st.cache_resource
def get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = None
    try:
        # 優先嘗試從 Secrets 讀取 (雲端環境)
        if "gcp_service_account" in st.secrets:
            # [關鍵修復] 將 Secrets 轉為普通字典，並修正 private_key 的換行符號
            creds_dict = dict(st.secrets["gcp_service_account"])
            if "private_key" in creds_dict:
                creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
            
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    except Exception as e:
        print(f"Secret loading error: {e}")
        pass

    # 如果 Secrets 失敗，嘗試讀取本地檔案 (本地開發環境)
    if creds is None:
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
        except FileNotFoundError:
            return None
            
    return gspread.authorize(creds)

def open_spreadsheet(client, source_str):
    if source_str.startswith("http"):
        return client.open_by_url(source_str)
    else:
        return client.open(source_str)

def hash_password(password):
    return hashlib.sha256(str(password).encode('utf-8')).hexdigest()

# ==========================================
# [核心] 使用者權限與訂閱管理
# ==========================================
def handle_user_login(email, password, user_sheet_name=None, is_register=False):
    client = get_gspread_client()
    if not client: return False, "API Error (請檢查 Secrets)"

    # [檢查] 確保 admin_sheet_url 存在
    admin_url = st.secrets.get("admin_sheet_url")
    if not admin_url:
        return True, {"Plan": "Dev", "Status": "Active"} 

    try:
        admin_book = client.open_by_url(admin_url)
        users_sheet = admin_book.worksheet("Users")
        records = users_sheet.get_all_records()
        df_users = pd.DataFrame(records)
        
        user_row = df_users[df_users["Email"] == email]
        pwd_hash = hash_password(password)
        today = datetime.now().date()

        if user_row.empty:
            if is_register:
                expire_date = today + timedelta(days=TRIAL_DAYS)
                new_user = {
                    "Email": email,
                    "Sheet_Name": user_sheet_name if user_sheet_name else "",
                    "Join_Date": str(today),
                    "Password_Hash": pwd_hash,
                    "Status": "Active",
                    "Expire_Date": str(expire_date),
                    "Plan": "Trial"
                }
                row_data = [
                    new_user["Email"], new_user["Sheet_Name"], new_user["Join_Date"], 
                    new_user["Password_Hash"], new_user["Status"], new_user["Expire_Date"], new_user["Plan"]
                ]
                users_sheet.append_row(row_data)
                return True, new_user
            else:
                return False, "User not found"
        else:
            user_info = user_row.iloc[0].to_dict()
            stored_hash = str(user_info.get("Password_Hash", ""))
            
            if stored_hash != pwd_hash:
                return False, "Password Incorrect"

            if user_info["Plan"] == "VIP":
                return True, user_info
            
            try:
                expire_dt = datetime.strptime(user_info["Expire_Date"], "%Y-%m-%d").date()
                if today > expire_dt:
                    return False, "Expired"
                else:
                    return True, user_info
            except:
                return False, "Date Error"
                
    except Exception as e:
        return False, f"Login Error: {e}"

# ==========================================
# 登入介面邏輯 (已修改)
# ==========================================
def login_flow():
    if "is_logged_in" in st.session_state and st.session_state.is_logged_in:
        return st.session_state.user_info["Sheet_Name"], "我的記帳本"

    if "login_mode" not in st.session_state: st.session_state.login_mode = "login"

    st.markdown("""<div class="login-container"><h2>👋 歡迎使用記帳本</h2>""", unsafe_allow_html=True)
    
    col_tab1, col_tab2 = st.columns(2)
    with col_tab1:
        if st.button("登入", use_container_width=True, type="primary" if st.session_state.login_mode == "login" else "secondary"):
            st.session_state.login_mode = "login"
            st.rerun()
    with col_tab2:
        if st.button("註冊新帳號", use_container_width=True, type="primary" if st.session_state.login_mode == "register" else "secondary"):
            st.session_state.login_mode = "register"
            st.rerun()
    
    # ------------------ 修改開始: 設定說明區域 ------------------
    st.info("💡 新用戶請先設定您的記帳本")
    with st.expander("👉 點此查看設定步驟 (含圖文教學)"):
        st.markdown(f"""
        **步驟 1：建立記帳本副本**  
        請點擊連結建立一份屬於您的 Google Sheet：  
        👉 [**[點此建立記帳本副本（下載後可更名）]**]({TEMPLATE_URL})
        """)
        #st.markdown("---")        
        st.markdown("**步驟 2：共用權限給機器人**")
        st.write("請將您的記帳本「共用」給以下機器人 Email (權限設為 **編輯者/Editor**)，系統才能寫入資料。")
        
        if "gcp_service_account" in st.secrets:
            st.code(st.secrets["gcp_service_account"]["client_email"], language="text")
        else:
            st.warning("⚠️ 系統尚未設定 Secrets，無法顯示機器人 Email")
        with st. expander("**操作示意圖：**"):
          # 圖片處理：
          # 1. 使用「內嵌 Expander」作為縮圖機制
          # 2. 只有使用者點擊展開時，才顯示完整寬度的圖片 (use_container_width=True)
          # 3. 這樣電腦版不會佔滿畫面，手機版點開後又能清晰查看
          if os.path.exists("guide.png"):
              with st.markdown("📷 點擊查看操作圖解 (點擊展開圖片)"):
                  st.image("guide.png", caption="請參照圖中紅框處共用給機器人", use_container_width=True)
          else:
              # 若無圖片，僅提示
              st.caption("🚫 (提示：將 guide.png 放入專案資料夾即可顯示圖解)")
    # ------------------ 修改結束 ------------------

    with st.container():
        email_input = st.text_input("Email", placeholder="name@example.com").strip()
        password_input = st.text_input("密碼", type="password", placeholder="設定您的密碼")
        
        if st.session_state.login_mode == "register":
            sheet_input = st.text_input("Google Sheet 網址/名稱")
            
            if st.button("✨ 註冊並登入", type="primary", use_container_width=True):
                if email_input and password_input and sheet_input:
                    with st.spinner("註冊中..."):
                        success, result = handle_user_login(email_input, password_input, sheet_input, is_register=True)
                        if success:
                            st.session_state.is_logged_in = True
                            st.session_state.user_info = result
                            st.success("註冊成功！")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error(f"註冊失敗：{result}")
                else:
                    st.warning("請填寫所有欄位")

        else:
            if st.button("🚀 登入", type="primary", use_container_width=True):
                if email_input and password_input:
                    with st.spinner("驗證中..."):
                        success, result = handle_user_login(email_input, password_input, is_register=False)
                        if success:
                            st.session_state.is_logged_in = True
                            st.session_state.user_info = result
                            st.success("登入成功！")
                            time.sleep(0.5)
                            st.rerun()
                        else:
                            if result == "Password Incorrect": st.error("❌ 密碼錯誤")
                            elif result == "User not found": st.error("❌ 帳號不存在，請先註冊")
                            elif result == "Expired": st.error("⛔ 您的訂閱已過期，請續費")
                            else: st.error(f"登入失敗: {result}")
                else:
                    st.warning("請輸入 Email 和密碼")

    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

CURRENT_SHEET_SOURCE, DISPLAY_TITLE = login_flow()

# ==========================================
# (以下為主程式邏輯，與之前版本相同)
# ==========================================

def open_spreadsheet(client, source_str):
    if source_str.startswith("http"): return client.open_by_url(source_str)
    else: return client.open(source_str)

@st.cache_data(ttl=300)
def get_data(worksheet_name, source_str):
    client = get_gspread_client()
    try:
        sheet = open_spreadsheet(client, source_str)
        worksheet = sheet.worksheet(worksheet_name)
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        if worksheet_name == "Settings":
            for col in ["Main_Category", "Sub_Category", "Payment_Method", "Currency", "Default_Currency"]:
                if col not in df.columns: df[col] = ""
        if worksheet_name == "Recurring":
            for col in ["Day", "Type", "Main_Category", "Sub_Category", "Payment_Method", "Currency", "Amount_Original", "Note", "Last_Run_Month"]:
                if col not in df.columns: df[col] = ""
        if not df.empty: df = df.dropna(how='all')
        return df
    except: return pd.DataFrame()

@st.cache_data(ttl=300)
def get_all_transactions(source_str):
    client = get_gspread_client()
    all_data = []
    try:
        sheet = open_spreadsheet(client, source_str)
        for ws in sheet.worksheets():
            if ws.title.startswith("Transactions"):
                data = ws.get_all_records()
                if data: all_data.extend(data)
        df = pd.DataFrame(all_data)
        if not df.empty:
            df = df.dropna(how='all')
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df['Amount_Def'] = pd.to_numeric(df['Amount_Def'], errors='coerce').fillna(0)
            df['Year'] = df['Date'].dt.year
            df['Month'] = df['Date'].dt.strftime('%Y-%m')
        return df
    except: return pd.DataFrame()

def append_data(worksheet_name, row_data, source_str):
    client = get_gspread_client()
    try:
        sheet = open_spreadsheet(client, source_str)
        worksheet = sheet.worksheet(worksheet_name)
        worksheet.append_row(row_data)
        return True
    except: return False

def save_settings_data(new_settings_df, source_str):
    client = get_gspread_client()
    try:
        sheet = open_spreadsheet(client, source_str)
        worksheet = sheet.worksheet("Settings")
        worksheet.clear()
        new_settings_df = new_settings_df.fillna("")
        data_to_write = [new_settings_df.columns.values.tolist()] + new_settings_df.values.tolist()
        worksheet.update(values=data_to_write)
        return True
    except: return False

def update_recurring_last_run(row_index, month_str, source_str):
    client = get_gspread_client()
    try:
        sheet = open_spreadsheet(client, source_str)
        worksheet = sheet.worksheet("Recurring")
        worksheet.update_cell(row_index + 2, 9, month_str)
        return True
    except: return False

def delete_recurring_rule(row_index, source_str):
    client = get_gspread_client()
    try:
        sheet = open_spreadsheet(client, source_str)
        worksheet = sheet.worksheet("Recurring")
        worksheet.delete_rows(row_index + 2)
        return True
    except: return False

def get_user_date(offset_hours):
    tz = timezone(timedelta(hours=offset_hours))
    return datetime.now(tz).date()

@st.cache_data(ttl=3600)
def get_exchange_rates():
    url = "https://rate.bot.com.tw/xrt?Lang=zh-TW"
    try:
        dfs = pd.read_html(url)
        df = dfs[0]
        df = df.iloc[:, 0:5]
        df.columns = ["Currency_Name", "Cash_Buy", "Cash_Sell", "Spot_Buy", "Spot_Sell"]
        df["Currency"] = df["Currency_Name"].str.extract(r'\(([A-Z]+)\)')
        rates = df.dropna(subset=['Currency']).copy()
        rates["Spot_Sell"] = pd.to_numeric(rates["Spot_Sell"], errors='coerce')
        rate_dict = rates.set_index("Currency")["Spot_Sell"].to_dict()
        rate_dict["TWD"] = 1.0
        return rate_dict
    except: return {}

def calculate_exchange(amount, input_currency, target_currency, rates):
    if input_currency == target_currency: return amount, 1.0
    try:
        rate_in = rates.get(input_currency)
        rate_target = rates.get(target_currency)
        if not rate_in or not rate_target: return amount, 0
        conversion_factor = rate_in / rate_target
        exchanged_amount = amount * conversion_factor
        return round(exchanged_amount, 2), conversion_factor
    except: return amount, 0

# --- 側邊欄 ---
with st.sidebar:
    st.header("🌍 地區與帳號")
    user_info = st.session_state.get("user_info", {})
    plan = user_info.get("Plan", "Trial")
    
    # ======== 修改開始 ========
    # 從 user_info 字典中讀取 Email，而不是讀取 st.session_state.user_email
    current_email = user_info.get("Email", "訪客")
    
    if plan == "VIP":
        st.markdown(f"👤 **{current_email}** <span class='vip-badge'>VIP</span>", unsafe_allow_html=True)
    else:
        expire = user_info.get("Expire_Date", "未知")
        st.markdown(f"👤 **{current_email}** <span class='trial-badge'>{plan}</span>", unsafe_allow_html=True)
        st.caption(f"到期日: {expire}")
    # ======== 修改結束 ========
    
    sheet_title = st.session_state.user_info.get("Sheet_Name", "未命名")
    st.success(f"📘 帳本：{sheet_title}")
    
    if st.button("🚪 登出"):
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.query_params.clear()
        st.rerun()
        
    st.divider()
    tz_options = {"台灣/北京 (UTC+8)": 8, "日本/韓國 (UTC+9)": 9, "泰國 (UTC+7)": 7, "美東 (UTC-4)": -4, "歐洲 (UTC+1)": 1}
    selected_tz_label = st.selectbox("當前位置時區", list(tz_options.keys()), index=0)
    user_offset = tz_options[selected_tz_label]
    st.info(f"日期：{get_user_date(user_offset)}")

rates = get_exchange_rates()

# --- 讀取設定 ---
settings_df = get_data("Settings", CURRENT_SHEET_SOURCE)
cat_mapping = {}     
payment_list = []
currency_list_custom = []
default_currency_setting = "TWD" 

if not settings_df.empty:
    if "Main_Category" in settings_df.columns and "Sub_Category" in settings_df.columns:
        valid_cats = settings_df[["Main_Category", "Sub_Category"]].astype(str)
        valid_cats = valid_cats[valid_cats["Main_Category"] != ""]
        for _, row in valid_cats.iterrows():
            main = row["Main_Category"]
            sub = row["Sub_Category"]
            if main not in cat_mapping: cat_mapping[main] = []
            if sub and sub != "" and sub not in cat_mapping[main]: cat_mapping[main].append(sub)
    if "Payment_Method" in settings_df.columns:
        payment_list = settings_df[settings_df["Payment_Method"] != ""]["Payment_Method"].unique().tolist()
    if "Currency" in settings_df.columns:
        currency_list_custom = settings_df[settings_df["Currency"] != ""]["Currency"].unique().tolist()
    if "Default_Currency" in settings_df.columns:
        saved = settings_df[settings_df["Default_Currency"] != ""]["Default_Currency"].unique().tolist()
        if saved: default_currency_setting = saved[0]

if not cat_mapping: cat_mapping = {"收入": ["薪資"], "食": ["早餐"]}
elif "收入" not in cat_mapping: cat_mapping["收入"] = ["薪資"]
if not payment_list: payment_list = ["現金"]
if not currency_list_custom: currency_list_custom = ["TWD"]
if default_currency_setting not in currency_list_custom: default_currency_setting = currency_list_custom[0]
main_cat_list = list(cat_mapping.keys())

# --- Callback ---
def save_all_to_sheet():
    rows = []
    if 'temp_cat_map' in st.session_state:
        for m, subs in st.session_state.temp_cat_map.items():
            if not subs: rows.append({"Main_Category": m, "Sub_Category": ""})
            else:
                for s in subs: rows.append({"Main_Category": m, "Sub_Category": s})
    df_cat = pd.DataFrame(rows)
    list_pay = st.session_state.get('temp_pay_list', payment_list)
    list_curr = st.session_state.get('temp_curr_list', currency_list_custom)
    max_len = max(len(df_cat), len(list_pay), len(list_curr)) if len(df_cat)>0 or len(list_pay)>0 or len(list_curr)>0 else 1
    final_df = pd.DataFrame()
    if not df_cat.empty:
        final_df["Main_Category"] = df_cat["Main_Category"].reindex(range(max_len)).fillna("")
        final_df["Sub_Category"] = df_cat["Sub_Category"].reindex(range(max_len)).fillna("")
    else:
        final_df["Main_Category"] = [""]*max_len
        final_df["Sub_Category"] = [""]*max_len
    final_df["Payment_Method"] = pd.Series(list_pay).reindex(range(max_len)).fillna("")
    final_df["Currency"] = pd.Series(list_curr).reindex(range(max_len)).fillna("")
    final_df["Default_Currency"] = ""
    if len(final_df) > 0: final_df.at[0, "Default_Currency"] = st.session_state.get('temp_default_curr', default_currency_setting)
    if save_settings_data(final_df, CURRENT_SHEET_SOURCE):
        st.toast("✅ 設定已儲存！", icon="💾")
        st.cache_data.clear()

def add_sub_callback(main_cat, key):
    new_val = st.session_state[key]
    if new_val:
        if new_val not in st.session_state.temp_cat_map[main_cat]:
            st.session_state.temp_cat_map[main_cat].append(new_val)
        st.session_state[key] = "" 
def add_pay_callback(key):
    new_val = st.session_state[key]
    if new_val:
        if new_val not in st.session_state.temp_pay_list:
            st.session_state.temp_pay_list.append(new_val)
        st.session_state[key] = ""
def add_curr_callback(key):
    new_val = st.session_state[key]
    if new_val:
        if new_val not in st.session_state.temp_curr_list:
            st.session_state.temp_curr_list.append(new_val)
        st.session_state[key] = ""

def check_and_run_recurring():
    if 'recurring_checked' in st.session_state: return 
    rec_df = get_data("Recurring", CURRENT_SHEET_SOURCE)
    if rec_df.empty: return
    sys_tz = timezone(timedelta(hours=8))
    today = datetime.now(sys_tz)
    current_month_str = today.strftime("%Y-%m")
    current_day = today.day
    executed = 0
    for idx, row in rec_df.iterrows():
        try:
            last_run = str(row['Last_Run_Month']).strip()
            scheduled_day = int(row['Day'])
            if last_run != current_month_str and current_day >= scheduled_day:
                amt_org = float(row['Amount_Original'])
                curr = row['Currency']
                amt_target, _ = calculate_exchange(amt_org, curr, default_currency_setting, rates)
                tx_date = today.strftime("%Y-%m-%d")
                tx_row = [tx_date, row['Type'], row['Main_Category'], row['Sub_Category'], row['Payment_Method'], curr, amt_org, amt_target, f"(自動) {row['Note']}", str(datetime.now(sys_tz))]
                if append_data("Transactions", tx_row, CURRENT_SHEET_SOURCE):
                    update_recurring_last_run(idx, current_month_str, CURRENT_SHEET_SOURCE)
                    executed += 1
        except: continue
    if executed > 0:
        st.toast(f"🤖 自動補登了 {executed} 筆固定收支！", icon="✅")
        st.cache_data.clear()
        time.sleep(1)
        st.rerun()
    st.session_state['recurring_checked'] = True
check_and_run_recurring()

# --- 頁籤 ---
tab1, tab2, tab3 = st.tabs(["📝 每日記帳", "📊 收支分析", "⚙️ 系統設定"])

# ================= Tab 1: 每日記帳 =================
with tab1:
    if st.session_state.get('should_clear_input'):
        st.session_state.form_amount_org = 0.0
        st.session_state.form_amount_def = 0.0
        st.session_state.form_note = ""
        st.session_state.should_clear_input = False

    if 'form_currency' not in st.session_state: st.session_state.form_currency = default_currency_setting
    if 'form_amount_org' not in st.session_state: st.session_state.form_amount_org = 0.0
    if 'form_amount_def' not in st.session_state: st.session_state.form_amount_def = 0.0
    if 'form_note' not in st.session_state: st.session_state.form_note = ""

    def on_input_change():
        c = st.session_state.form_currency
        a = st.session_state.form_amount_org
        val, _ = calculate_exchange(a, c, default_currency_setting, rates)
        st.session_state.form_amount_def = val

    user_today = get_user_date(user_offset)
    current_month_str = user_today.strftime("%Y-%m")
    
    tx_df = get_data("Transactions", CURRENT_SHEET_SOURCE)
    total_income = 0
    total_expense = 0
    
    if not tx_df.empty and 'Date' in tx_df.columns:
        tx_df['Date'] = pd.to_datetime(tx_df['Date'], errors='coerce')
        mask = (tx_df['Date'].dt.strftime('%Y-%m') == current_month_str)
        month_tx = tx_df[mask]
        month_tx['Amount_Def'] = pd.to_numeric(month_tx['Amount_Def'], errors='coerce').fillna(0)
        
        if 'Type' in month_tx.columns:
            total_income = month_tx[month_tx['Type'] == '收入']['Amount_Def'].sum()
            total_expense = month_tx[month_tx['Type'] != '收入']['Amount_Def'].sum()
    
    balance = total_income - total_expense
    balance_class = "val-green" if balance >= 0 else "val-red"

    st.markdown(f"""
    <div class="metric-container">
        <div class="metric-card">
            <span class="metric-label">本月總收入 ({default_currency_setting})</span>
            <span class="metric-value">${total_income:,.2f}</span>
        </div>
        <div class="metric-card">
            <span class="metric-label">已支出 ({default_currency_setting})</span>
            <span class="metric-value">${total_expense:,.2f}</span>
        </div>
        <div class="metric-card">
            <span class="metric-label">剩餘可用</span>
            <span class="metric-value {balance_class}">${balance:,.2f}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.container():
        st.markdown("##### ✍️ 新增交易")
        c1, c2 = st.columns([1, 1])
        with c1: date_input = st.date_input("日期", user_today)
        with c2: payment = st.selectbox("付款方式", payment_list)
        c3, c4 = st.columns([1, 1])
        with c3: main_cat = st.selectbox("大類別", main_cat_list, key="input_main_cat")
        with c4: sub_cat = st.selectbox("次類別", cat_mapping.get(main_cat, []))
        
        with st.container(border=True): 
            st.caption("💰 金額設定")
            c5, c6, c7 = st.columns([1.5, 2, 2])
            try: curr_index = currency_list_custom.index(default_currency_setting)
            except: curr_index = 0
            with c5: currency = st.selectbox("幣別", currency_list_custom, index=curr_index, key="form_currency", on_change=on_input_change)
            with c6: amount_org = st.number_input(f"金額 ({currency})", step=1.0, key="form_amount_org", on_change=on_input_change)
            with c7: 
                amount_def = st.number_input(f"折合 {default_currency_setting}", step=0.1, key="form_amount_def")
                if currency != default_currency_setting and amount_org != 0:
                     _, rate_used = calculate_exchange(100, currency, default_currency_setting, rates)
                     if rate_used > 0: st.caption(f"匯率: {rate_used:.4f}")
        
        note = st.text_input("備註", max_chars=20, placeholder="輸入消費內容 (限20字)...", key="form_note")
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("確認送出記帳", type="primary", use_container_width=True):
            if amount_def == 0: st.error("金額不能為 0")
            else:
                with st.spinner('📡 資料寫入中...'):
                    tx_type = "收入" if main_cat == "收入" else "支出"
                    sys_now = datetime.now()
                    row = [str(date_input), tx_type, main_cat, sub_cat, payment, currency, amount_org, amount_def, note, str(sys_now)]
                    if append_data("Transactions", row, CURRENT_SHEET_SOURCE):
                        st.success(f"✅ {tx_type}已記錄 ${amount_def:,.2f}！")
                        st.session_state['should_clear_input'] = True
                        st.cache_data.clear()
                        time.sleep(1)
                        st.rerun()
                    else: st.error("❌ 寫入失敗")

# ================= Tab 2: 收支分析 =================
with tab2:
    st.markdown("##### 📊 收支狀況")
    df_all = get_all_transactions(CURRENT_SHEET_SOURCE)
    if df_all.empty:
        st.info("尚無交易資料")
    else:
        av_years = sorted(df_all['Year'].dropna().unique().tolist())
        with st.expander("📅 篩選年度區間", expanded=True):
            if len(av_years)>0:
                mn, mx = int(min(av_years)), int(max(av_years))
                sel_y = st.slider("年份", mn, mx, (mn, mx)) if mn != mx else (mn, mx)
                df_y = df_all[(df_all['Year']>=sel_y[0]) & (df_all['Year']<=sel_y[1])]
                exp_t = df_y[df_y['Type']!='收入'].groupby('Year')['Amount_Def'].sum().reset_index()
                exp_t['Type']='支出'
                inc_t = df_y[df_y['Type']=='收入'].groupby('Year')['Amount_Def'].sum().reset_index()
                inc_t['Type']='收入'
                chart = pd.concat([exp_t, inc_t])
                if not chart.empty:
                    import plotly.express as px
                    fig = px.bar(chart, x="Year", y="Amount_Def", color="Type", barmode="group", color_discrete_map={"收入":"#2ecc71","支出":"#ff6b6b"})
                    st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        am = sorted(df_all['Month'].unique(), reverse=True)
        tm = st.selectbox("🗓️ 查看詳細月份", am)
        md = df_all[df_all['Month']==tm]
        mi = md[md['Type']=='收入']['Amount_Def'].sum()
        me = md[md['Type']!='收入']['Amount_Def'].sum()
        
        st.markdown(f"""
        <div class="metric-container">
            <div class="metric-card" style="border-left: 5px solid #2ecc71;"><span class="metric-label">總收入</span><span class="metric-value">${mi:,.2f}</span></div>
            <div class="metric-card" style="border-left: 5px solid #ff6b6b;"><span class="metric-label">總支出</span><span class="metric-value">${me:,.2f}</span></div>
            <div class="metric-card"><span class="metric-label">結餘</span><span class="metric-value">${mi-me:,.2f}</span></div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.expander("🔍 檢視明細"):
            debug = md[['Date','Main_Category','Sub_Category','Amount_Original','Currency','Amount_Def','Note']].sort_values(by='Date', ascending=False)
            st.dataframe(debug, use_container_width=True)

        ed = md[md['Type']!='收入']
        if not ed.empty:
            pd_pie = ed.groupby("Main_Category")["Amount_Def"].sum().reset_index()
            pd_pie = pd_pie[pd_pie["Amount_Def"]>0]
            if not pd_pie.empty:
                fig_pie = px.pie(pd_pie, values="Amount_Def", names="Main_Category", hole=0.5, color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_pie, use_container_width=True)

# ================= Tab 3: 設定管理 =================
with tab3:
    st.markdown("##### ⚙️ 系統資料庫")
    if 'temp_cat_map' not in st.session_state: st.session_state.temp_cat_map = cat_mapping
    if 'temp_pay_list' not in st.session_state: st.session_state.temp_pay_list = payment_list
    if 'temp_curr_list' not in st.session_state: st.session_state.temp_curr_list = currency_list_custom
    if 'temp_default_curr' not in st.session_state: st.session_state.temp_default_curr = default_currency_setting

    with st.expander("🔄 每月固定收支", expanded=True):
        with st.popover("➕ 新增固定規則", use_container_width=True):
            if 'rec_currency' not in st.session_state: st.session_state.rec_currency = default_currency_setting
            if 'rec_amount_org' not in st.session_state: st.session_state.rec_amount_org = 0.0
            def on_rec_change():
                c = st.session_state.rec_currency
                a = st.session_state.rec_amount_org
                val, _ = calculate_exchange(a, c, default_currency_setting, rates)
                st.session_state.rec_amount_def = val
            rec_day = st.number_input("每月幾號執行?", 1, 31, 5)
            c1, c2 = st.columns(2)
            with c1: rec_main = st.selectbox("大類別", main_cat_list, key="rec_main")
            with c2: rec_sub = st.selectbox("次類別", cat_mapping.get(rec_main, []), key="rec_sub")
            rec_pay = st.selectbox("付款方式", payment_list, key="rec_pay")
            c1, c2, c3 = st.columns([1.5, 2, 2])
            with c1: rec_curr = st.selectbox("幣別", currency_list_custom, key="rec_currency", on_change=on_rec_change)
            with c2: rec_amt_org = st.number_input("原幣", step=1.0, key="rec_amount_org", on_change=on_rec_change)
            with c3: rec_amt_def = st.number_input(f"折合 {default_currency_setting}", step=0.1, key="rec_amount_def")
            rec_note = st.text_input("備註", key="rec_note")
            if st.button("儲存規則", type="primary", use_container_width=True):
                rt = "收入" if rec_main == "收入" else "支出"
                if append_data("Recurring", [rec_day, rt, rec_main, rec_sub, rec_pay, rec_curr, rec_amt_org, rec_note, "New", "Active"], CURRENT_SHEET_SOURCE):
                    st.success("規則已新增")
                    st.cache_data.clear()
                    time.sleep(1)
                    st.rerun()
        st.markdown("---")
        rec_df = get_data("Recurring", CURRENT_SHEET_SOURCE)
        if not rec_df.empty:
            for idx, row in rec_df.iterrows():
                with st.expander(f"📅 每月 {row['Day']} 號 - {row['Main_Category']} > {row['Sub_Category']} > {row['Amount_Original']} {row['Currency']}"):
                    c1, c2 = st.columns([4,1])
                    with c1: st.write(f"📝 {row['Note']} ({row['Payment_Method']})")
                    with c2: 
                        if st.button("🗑️", key=f"del_{idx}"):
                             if delete_recurring_rule(idx, CURRENT_SHEET_SOURCE):
                                 st.toast("已刪除"); st.cache_data.clear(); time.sleep(1); st.rerun()

    with st.expander("📂 類別與子類別"):
        with st.popover("➕ 新增大類", use_container_width=True):
            nm = st.text_input("類別名稱")
            if st.button("確認"):
                if nm and nm not in st.session_state.temp_cat_map:
                    st.session_state.temp_cat_map[nm] = []
                    save_all_to_sheet()
                    st.rerun()
        for idx, main in enumerate(st.session_state.temp_cat_map.keys()):
            with st.container():
                with st.expander(f"📁 {main}"):
                    curr_subs = st.session_state.temp_cat_map[main]
                    st.multiselect("子類", curr_subs, default=curr_subs, key=f"ms_{main}", on_change=lambda m=main, k=f"ms_{main}": [st.session_state.temp_cat_map.update({m: st.session_state[k]}), save_all_to_sheet()])
                    c1, c2 = st.columns([3,1])
                    sk = f"new_sub_{main}"
                    if sk not in st.session_state: st.session_state[sk]=""
                    with c1: st.text_input("add", key=sk, label_visibility="collapsed")
                    with c2: st.button("加入", key=f"b_{main}", on_click=add_sub_callback, args=(main, sk))
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button(f"🗑️ 刪除 {main}", key=f"dm_{main}"):
                        del st.session_state.temp_cat_map[main]
                        save_all_to_sheet()
                        st.rerun()

    with st.expander("💳 付款與幣別"):
        pays = st.session_state.temp_pay_list
        st.multiselect("付款方式", pays, default=pays, key="mp_pay", on_change=lambda: [st.session_state.update(temp_pay_list=st.session_state.mp_pay), save_all_to_sheet()])
        c1, c2 = st.columns([3,1])
        with c1: 
            if "np" not in st.session_state: st.session_state.np = ""
            st.text_input("np", key="np", label_visibility="collapsed")
        with c2: st.button("加入", key="bp", on_click=add_pay_callback, args=("np",))
        
        st.divider()
        curs = st.session_state.temp_curr_list
        st.multiselect("常用幣別", curs, default=curs, key="mp_cur", on_change=lambda: [st.session_state.update(temp_curr_list=st.session_state.mp_cur), save_all_to_sheet()])
        c1, c2 = st.columns([3,1])
        with c1: 
            if "nc" not in st.session_state: st.session_state.nc = ""
            st.text_input("nc", key="nc", label_visibility="collapsed")
        with c2: st.button("加入", key="bc", on_click=add_curr_callback, args=("nc",))
        
        st.markdown("<br>", unsafe_allow_html=True)
        try: di = st.session_state.temp_curr_list.index(st.session_state.temp_default_curr)
        except: di = 0
        nd = st.selectbox("預設幣別", st.session_state.temp_curr_list, index=di, key="sel_def")
        if nd != st.session_state.temp_default_curr:
            st.session_state.temp_default_curr = nd
            save_all_to_sheet()
            st.toast("已更新")
    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("💾 儲存所有設定", type="primary", use_container_width=True):
        save_all_to_sheet()
        st.rerun()