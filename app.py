import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date, timedelta, timezone
import time
import os
import hashlib
import smtplib
from email.mime.text import MIMEText
import random
import string
import re
import requests

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
        padding-top: 4rem !important;
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
    /* 按鈕樣式 */
    div.stButton > button { border-radius: 8px; font-weight: 600; }
    
    /* Tab 樣式微調 */
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: white;
        border-radius: 8px 8px 0 0;
        gap: 1px;
        padding: 10px 20px;
        font-size: 1.1rem;
        font-weight: 600;
        color: #6c757d;
        border: 1px solid #dee2e6;
        border-bottom: none;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff;
        color: #0d6efd !important;
        border-top: 3px solid #0d6efd;
    }
    .login-container { max-width: 500px; margin: 30px auto; padding: 40px; background: white; border-radius: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); text-align: center; }
    .vip-badge { background-color: #FFD700; color: #000; padding: 2px 8px; border-radius: 10px; font-size: 0.8em; font-weight: bold; }
    .trial-badge { background-color: #87CEEB; color: #000; padding: 2px 8px; border-radius: 10px; font-size: 0.8em; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. 核心連線與工具函式
# ==========================================
@st.cache_resource
def get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = None
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            if "private_key" in creds_dict:
                creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    except: pass
    if creds is None:
        try: creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
        except: return None
    return gspread.authorize(creds)

def open_spreadsheet(client, source_str):
    if source_str.startswith("http"): return client.open_by_url(source_str)
    else: return client.open(source_str)

def get_sheet_title_safe(source_str):
    client = get_gspread_client()
    try:
        sh = open_spreadsheet(client, source_str)
        return sh.title
    except: return "我的記帳本"

def hash_password(password):
    return hashlib.sha256(str(password).encode('utf-8')).hexdigest()

def is_valid_email(email):
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

def mask_email(email):
    try:
        if "@" not in email: return email
        name, domain = email.split("@")
        if len(name) <= 3: return f"{name[0]}***@{domain}"
        return f"{name[:3]}***@{domain}"
    except: return "******"

# --- 匯率相關函式 ---
@st.cache_data(ttl=3600)
def get_exchange_rates():
    """使用 Frankfurter API 獲取穩定匯率"""
    default_rates = {"TWD": 1.0, "USD": 32.3, "HKD": 4.12, "JPY": 0.21, "SGD": 24.1, "CNY": 4.5, "EUR": 34.5}
    fetch_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        url = "https://api.frankfurter.app/latest?from=TWD"
        response = requests.get(url, timeout=10)
        data = response.json()
        if response.status_code == 200 and "rates" in data:
            api_rates = data["rates"]
            processed_rates = {"TWD": 1.0}
            for curr, val in api_rates.items():
                if val != 0: processed_rates[curr] = round(1 / val, 4)
            return {"rates": processed_rates, "time": fetch_time, "source": "Frankfurter API"}
        else: raise Exception("API Error")
    except:
        return {"rates": default_rates, "time": f"API連線失敗，使用預設匯率 ({fetch_time})", "source": "系統預設"}

def calculate_exchange(amount, input_currency, target_currency, rates_data):
    # 支援傳入字典或包裝過的字典
    rates = rates_data["rates"] if isinstance(rates_data, dict) and "rates" in rates_data else rates_data
    if input_currency == target_currency: return amount, 1.0
    try:
        rate_in = rates.get(input_currency)
        rate_target = rates.get(target_currency)
        if not rate_in or not rate_target: return amount, 1.0
        conversion_factor = rate_in / rate_target
        return round(amount * conversion_factor, 2), conversion_factor
    except: return amount, 1.0

# --- 資料讀取與寫入 ---
@st.cache_data(ttl=300)
def get_data(worksheet_name, source_str):
    client = get_gspread_client()
    try:
        sheet = open_spreadsheet(client, source_str)
        worksheet = sheet.worksheet(worksheet_name)
        data = worksheet.get_all_records()
        return pd.DataFrame(data)
    except: return pd.DataFrame()

@st.cache_data(ttl=60) # 為了 Dashboard 準確，縮短交易資料快取
def get_all_transactions(source_str):
    client = get_gspread_client()
    all_data = []
    try:
        sheet = open_spreadsheet(client, source_str)
        for ws in sheet.worksheets():
            # 自動識別所有包含 Transaction 的分頁 (如 Transaction, Transaction_history)
            if "Transaction" in ws.title:
                data = ws.get_all_records()
                if data: all_data.extend(data)
        df = pd.DataFrame(all_data)
        if not df.empty:
            df = df.dropna(how='all')
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df['Amount_Def'] = pd.to_numeric(df['Amount_Def'], errors='coerce').fillna(0)
            df['Year'] = df['Date'].dt.year
            df['Month'] = df['Date'].dt.strftime('%Y-%m')
            if "Type" not in df.columns: df["Type"] = "支出"
        return df
    except: return pd.DataFrame()

def append_data(worksheet_name, row_data, source_str):
    client = get_gspread_client()
    try:
        sheet = open_spreadsheet(client, source_str)
        # 如果 worksheet_name 是 Transactions 但分頁叫 Transaction，自動校正
        try:
            worksheet = sheet.worksheet(worksheet_name)
        except:
            if worksheet_name == "Transactions": worksheet = sheet.worksheet("Transaction")
            else: raise Exception("Worksheet not found")
            
        if "Transaction" in worksheet_name:
            recorder = st.session_state.user_info.get("Nickname", st.session_state.user_info.get("Email"))
            row_data.append(recorder)
        worksheet.append_row(row_data)
        return True
    except: return False

# ==========================================
# 2. 登入與權限管理 (省略部分重複邏輯以節省空間，保留核心流程)
# ==========================================
# ... (這裡保留你原本代碼中的 send_otp_email, send_invitation_email, handle_user_login, add_binding 等函式) ...
# [由於篇幅限制，以下直接進入主邏輯，請確保保留你原本檔案中所有關於 User 與 Email 的 Function]

# --- 這裡插入你原本程式碼中從 send_otp_email 到 login_flow 的所有函數 ---
# (請參考 all code-4.txt 的 114行 - 425行)

def send_otp_email(to_email, code, subject="【記帳本】驗證碼"):
    if "email" not in st.secrets: return False, "尚未設定 Email Secrets"
    sender = st.secrets["email"]["sender"]; pwd = st.secrets["email"]["password"]
    msg = MIMEText(f"{subject}：{code}\n\n請在頁面上輸入此驗證碼以完成操作。")
    msg['Subject'] = subject; msg['From'] = sender; msg['To'] = to_email
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender, pwd); server.sendmail(sender, to_email, msg.as_string())
        return True, "驗證碼已發送"
    except Exception as e: return False, f"寄信失敗: {e}"

def handle_user_login(email, password, user_sheet_name=None, nickname=None, is_register=False):
    client = get_gspread_client()
    admin_url = st.secrets.get("admin_sheet_url")
    try:
        admin_book = client.open_by_url(admin_url)
        users_sheet = admin_book.worksheet("Users")
        bindings_sheet = admin_book.worksheet("Book_Bindings")
        records = users_sheet.get_all_records(); df_users = pd.DataFrame(records)
        pwd_hash = hash_password(password)
        user_row = df_users[df_users["Email"] == email] if not df_users.empty else pd.DataFrame()
        if is_register:
            if not user_row.empty: return False, "帳號已存在"
            expire_date = datetime.now().date() + timedelta(days=TRIAL_DAYS)
            new_user = [email, user_sheet_name, str(datetime.now().date()), pwd_hash, "Active", str(expire_date), "Trial", nickname]
            users_sheet.append_row(new_user)
            bindings_sheet.append_row([email, user_sheet_name, "我的記帳本", "Owner"])
            return True, {"Email": email, "Nickname": nickname, "Plan": "Trial", "Books": [{"name": "我的記帳本", "url": user_sheet_name, "role": "Owner"}]}
        if user_row.empty: return False, "找不到用戶"
        user_info = user_row.iloc[0].to_dict()
        if user_info["Password_Hash"] != pwd_hash and user_info["Password_Hash"] != "RESET_REQUIRED": return False, "密碼錯誤"
        b_records = bindings_sheet.get_all_records(); df_bind = pd.DataFrame(b_records)
        user_books = df_bind[df_bind["Email"] == email].to_dict('records')
        user_info["Books"] = [{"name": b["Book_Name"], "url": b["Sheet_URL"], "role": b.get("Role", "Member")} for b in user_books]
        return True, user_info
    except Exception as e: return False, str(e)

def login_flow():
    if "is_logged_in" in st.session_state and st.session_state.is_logged_in:
        return st.session_state.current_book_url, st.session_state.current_book_name
    # (此處為簡化後的登入邏輯示範，請以你原本的完整 login_flow 為主)
    st.title("👋 歡迎使用記帳本")
    email = st.text_input("Email"); pwd = st.text_input("密碼", type="password")
    if st.button("登入"):
        success, res = handle_user_login(email, pwd)
        if success:
            st.session_state.is_logged_in = True; st.session_state.user_info = res
            st.session_state.current_book_url = res["Books"][0]["url"]; st.session_state.current_book_name = res["Books"][0]["name"]
            st.rerun()
    st.stop()

# ==========================================
# 3. 主程式執行
# ==========================================

# 啟動登入 (這裡會擋住直到登入成功)
try:
    CURRENT_SHEET_SOURCE = st.session_state.current_book_url
    DISPLAY_TITLE = st.session_state.current_book_name
except:
    CURRENT_SHEET_SOURCE, DISPLAY_TITLE = login_flow()

# 獲取匯率資訊 (全域使用)
rates_info = get_exchange_rates()
rates = rates_info["rates"]

# 讀取設定 (幣別、類別等)
settings_df = get_data("Settings", CURRENT_SHEET_SOURCE)
default_currency_setting = "TWD"
cat_mapping = {"收入": ["薪資"], "食": ["早餐", "午餐", "晚餐"]}
payment_list = ["現金", "信用卡"]
currency_list_custom = ["TWD", "USD", "SGD"]

if not settings_df.empty:
    if "Default_Currency" in settings_df.columns:
        dc = settings_df[settings_df["Default_Currency"] != ""]["Default_Currency"].tolist()
        if dc: default_currency_setting = dc[0]
    # (類別與付款方式讀取邏輯同你原本的代碼...)

# --- 側邊欄 ---
with st.sidebar:
    st.header("🌍 帳號設定")
    st.write(f"👤 {st.session_state.user_info.get('Nickname')}")
    if st.button("🚪 登出"):
        st.session_state.clear(); st.rerun()

# --- Tabs ---
tab1, tab2, tab3 = st.tabs(["📝 每日記帳", "📊 收支分析", "⚙️ 系統設定"])

with tab1:
    # --- 修正後的 Dashboard 計算 ---
    df_all_raw = get_all_transactions(CURRENT_SHEET_SOURCE)
    total_inc = 0; total_exp = 0
    today_dt = datetime.now(); current_month_str = today_dt.strftime("%Y-%m")

    if not df_all_raw.empty:
        mask = (df_all_raw['Date'].dt.strftime('%Y-%m') == current_month_str)
        mtx = df_all_raw[mask].copy()
        if not mtx.empty:
            total_inc = mtx[mtx['Type'] == '收入']['Amount_Def'].sum()
            total_exp = mtx[mtx['Type'] != '收入']['Amount_Def'].sum()
    
    bal = total_inc - total_exp
    b_cls = "val-green" if bal >= 0 else "val-red"

    st.markdown(f"""
    <div class="metric-container">
        <div class="metric-card"><span class="metric-label">本月總收入 ({default_currency_setting})</span><span class="metric-value">${total_inc:,.2f}</span></div>
        <div class="metric-card"><span class="metric-label">已支出 ({default_currency_setting})</span><span class="metric-value">${total_exp:,.2f}</span></div>
        <div class="metric-card"><span class="metric-label">剩餘可用</span><span class="metric-value {b_cls}">${bal:,.2f}</span></div>
    </div>""", unsafe_allow_html=True)

    # --- 新增交易表單 ---
    with st.form("add_tx_form", clear_on_submit=True):
        st.markdown("##### ✍️ 新增交易")
        c1, c2 = st.columns(2)
        with c1: tx_date = st.date_input("日期", date.today())
        with c2: tx_pay = st.selectbox("付款方式", payment_list)
        c3, c4 = st.columns(2)
        with c3: tx_main = st.selectbox("大類別", list(cat_mapping.keys()))
        with c4: tx_sub = st.selectbox("次類別", cat_mapping.get(tx_main, ["無"]))
        
        c5, c6 = st.columns(2)
        with c5: tx_curr = st.selectbox("幣別", currency_list_custom, index=currency_list_custom.index(default_currency_setting) if default_currency_setting in currency_list_custom else 0)
        with c6: tx_amt = st.number_input("金額", min_value=0.0, step=1.0)
        
        tx_note = st.text_input("備註")
        if st.form_submit_button("確認送出", use_container_width=True):
            amt_def, _ = calculate_exchange(tx_amt, tx_curr, default_currency_setting, rates_info)
            tx_type = "收入" if tx_main == "收入" else "支出"
            new_row = [str(tx_date), tx_type, tx_main, tx_sub, tx_pay, tx_curr, tx_amt, amt_def, tx_note, str(datetime.now())]
            if append_data("Transactions", new_row, CURRENT_SHEET_SOURCE):
                st.success("記帳成功！"); st.cache_data.clear(); time.sleep(1); st.rerun()

with tab2:
    st.markdown("##### 📊 收支狀況")
    if df_all_raw.empty:
        st.info("尚無資料")
    else:
        # 使用一致的 df_all_raw 進行圖表分析 (省略具體 Plotly 程式碼)
        st.dataframe(df_all_raw.sort_values("Date", ascending=False), use_container_width=True)

with tab3:
    st.markdown("##### ⚙️ 系統設定")
    # ... (此處保留你原本的類別設定、帳本管理邏輯) ...

    st.markdown("---")
    st.markdown(f"##### 💱 即時匯率參考")
    st.caption(f"資料來源：{rates_info.get('source')} | 更新時間：{rates_info.get('time')}")
    with st.expander("查看當前匯率清單"):
        sorted_rates = dict(sorted(rates.items(), key=lambda item: item[1], reverse=True))
        df_rates = pd.DataFrame(list(sorted_rates.items()), columns=['幣別', f'折合 {default_currency_setting}'])
        st.dataframe(df_rates, use_container_width=True, height=300)