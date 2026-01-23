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
import plotly.express as px  # 確保引入 plotly

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
    except Exception as e:
        print(f"Secret loading error: {e}")
        pass
    if creds is None:
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
        except FileNotFoundError:
            return None
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

# --- Email 相關函式 ---
def send_otp_email(to_email, code, subject="【記帳本】驗證碼"):
    if "email" not in st.secrets: return False, "尚未設定 Email Secrets"
    sender = st.secrets["email"]["sender"]
    pwd = st.secrets["email"]["password"]
    msg = MIMEText(f"{subject}：{code}\n\n請在頁面上輸入此驗證碼以完成操作。")
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = to_email
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, to_email, msg.as_string())
        return True, "驗證碼已發送"
    except Exception as e: return False, f"寄信失敗: {e}"

# [修改] 發送邀請通知信函式 (已加入個資遮罩、標題改用暱稱)
def send_invitation_email(to_email, inviter_email, book_name, inviter_nickname=None):
    if "email" not in st.secrets: return False, "尚未設定 Email Secrets"
    
    # ⚠️ 請確認這裡的網址是您正確的 App 連結
    APP_URL = "https://expense-tracker-test.streamlit.app" 
    
    sender = st.secrets["email"]["sender"]
    pwd = st.secrets["email"]["password"]
    
    # --- 1. 決定顯示名稱 (有暱稱用暱稱，沒暱稱用遮罩 Email) ---
    if inviter_nickname:
        display_name = inviter_nickname
    else:
        display_name = mask_email(inviter_email)
        
    masked_to = mask_email(to_email)
    
    # --- 2. 標題與內容 ---
    # 標題改用 display_name (暱稱)
    subject = f"【我的記帳本】您收到來自 {display_name} 的共用邀請"
    
    body = f"""
    您好！

    使用者 {display_name} ({mask_email(inviter_email)}) 邀請您共同管理記帳本：「{book_name}」。

    --------------------------------------------------
    🔗 App 連結：{APP_URL}
    --------------------------------------------------

    👉 如果您已有帳號：
    請點擊上方連結登入 App，您將在「切換帳本」選單中看到此新帳本。

    👉 如果您尚未註冊 / 初次使用：
    您的帳號已預先建立。請前往 App 首頁：
    1. 點擊「🔑 忘記密碼 / 啟用帳號」
    2. 輸入您的 Email ({masked_to}) 
    3. 收取驗證碼並設定您的密碼與暱稱
    --------------------------------------------------

    祝記帳愉快！
    """
    
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = to_email
    
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, to_email, msg.as_string())
        return True, "邀請信已發送"
    except Exception as e:
        print(f"Mail Error: {e}")
        return False, f"寄信失敗: {e}"

def reset_user_password(email, new_password, new_nickname=None):
    """重設密碼，並處理試用期重置與暱稱更新"""
    client = get_gspread_client()
    try:
        admin_book = client.open_by_url(st.secrets["admin_sheet_url"])
        users_sheet = admin_book.worksheet("Users")
        
        # 尋找使用者 Row
        cell = users_sheet.find(email)
        if not cell: return False, "找不到使用者"
        
        row = cell.row
        old_hash = users_sheet.cell(row, 4).value
        new_hash = hash_password(new_password)
        
        updates = []
        updates.append({'range': f'D{row}', 'values': [[new_hash]]}) # 更新密碼
        
        # 如果是初次啟用 (RESET_REQUIRED)，重置加入日期與到期日
        if old_hash == "RESET_REQUIRED":
            today = datetime.now().date()
            expire_date = today + timedelta(days=TRIAL_DAYS)
            updates.append({'range': f'C{row}', 'values': [[str(today)]]}) # Join_Date
            updates.append({'range': f'F{row}', 'values': [[str(expire_date)]]}) # Expire_Date
        
        if new_nickname:
            updates.append({'range': f'H{row}', 'values': [[new_nickname]]})
            
        users_sheet.batch_update(updates)
        return True, "密碼更新成功 (若是首次啟用，試用期已重置)"
    except Exception as e: return False, f"資料庫錯誤: {e}"

def update_user_nickname(email, new_nickname):
    """更新使用者暱稱"""
    client = get_gspread_client()
    try:
        admin_book = client.open_by_url(st.secrets["admin_sheet_url"])
        users_sheet = admin_book.worksheet("Users")
        cell = users_sheet.find(email)
        if not cell: return False, "找不到使用者"
        users_sheet.update_cell(cell.row, 8, new_nickname)
        return True, "暱稱更新成功"
    except Exception as e: return False, f"Error: {e}"

@st.cache_data(ttl=600)
def get_all_users_nickname_map():
    """回傳 {email: nickname} 的字典，用於顯示"""
    client = get_gspread_client()
    try:
        admin_book = client.open_by_url(st.secrets["admin_sheet_url"])
        users_sheet = admin_book.worksheet("Users")
        records = users_sheet.get_all_records()
        return {row["Email"]: row.get("Nickname", "") for row in records}
    except: return {}

# ==========================================
# [新增] 寫入系統日誌 (Audit Log)
# ==========================================
def write_system_log(operator, action, target_email, book_name, sheet_url):
    client = get_gspread_client()
    try:
        admin_book = client.open_by_url(st.secrets["admin_sheet_url"])
        try: log_sheet = admin_book.worksheet("System_Logs")
        except: log_sheet = admin_book.add_worksheet("System_Logs", 1000, 6); log_sheet.append_row(["Timestamp", "Operator", "Action", "Target_Email", "Book_Name", "Sheet_URL"])
        
        tz_tw = timezone(timedelta(hours=8))
        now_str = datetime.now(tz_tw).strftime("%Y-%m-%d %H:%M:%S")
        log_sheet.append_row([now_str, operator, action, target_email, book_name, sheet_url])
        return True
    except Exception as e:
        print(f"Log Error: {e}")
        return False

# ==========================================
# [新增] 註冊前置檢查 (防呆檢查)
# ==========================================
def validate_registration_pre_check(email, sheet_url):
    client = get_gspread_client()
    if not client: return False, "API Error"
    admin_url = st.secrets.get("admin_sheet_url")
    
    try:
        admin_book = client.open_by_url(admin_url)
        users_sheet = admin_book.worksheet("Users")
        try: cell = users_sheet.find(email); 
        except: cell = None
        if cell: return False, "❌ 此 Email 已存在系統中。請直接「登入」。"

        try:
            bindings_sheet = admin_book.worksheet("Book_Bindings")
            b_records = bindings_sheet.get_all_records()
            df_bind = pd.DataFrame(b_records)
            if not df_bind.empty and "Sheet_URL" in df_bind.columns:
                conflict = df_bind[df_bind["Sheet_URL"] == sheet_url]
                if not conflict.empty:
                    owner_email = conflict.iloc[0]["Email"]
                    owner_nickname = ""
                    try:
                        records_u = users_sheet.get_all_records()
                        df_u = pd.DataFrame(records_u)
                        o_row = df_u[df_u["Email"] == owner_email]
                        if not o_row.empty: owner_nickname = o_row.iloc[0]["Nickname"]
                    except: pass
                    display_name = owner_nickname if owner_nickname else mask_email(owner_email)
                    return False, f"❌ 此帳本已被 **{display_name}** 綁定為擁有者。請聯繫他邀請您加入。"
        except: pass
        return True, "OK"
    except Exception as e: return False, f"系統檢查失敗: {e}"

# ==========================================
# [核心] 使用者與多帳本管理
# ==========================================
def handle_user_login(email, password, user_sheet_name=None, nickname=None, is_register=False):
    client = get_gspread_client()
    if not client: return False, "API Error"
    admin_url = st.secrets.get("admin_sheet_url")
    if not admin_url: return True, {"Plan": "Dev", "Status": "Active", "Nickname": "Dev"} 

    try:
        admin_book = client.open_by_url(admin_url)
        users_sheet = admin_book.worksheet("Users")
        try: bindings_sheet = admin_book.worksheet("Book_Bindings")
        except: bindings_sheet = admin_book.add_worksheet("Book_Bindings", 100, 4); bindings_sheet.append_row(["Email", "Sheet_URL", "Book_Name", "Role"])
        
        records = users_sheet.get_all_records()
        if not records:
            df_users = pd.DataFrame(columns=["Email", "Sheet_Name", "Join_Date", "Password_Hash", "Status", "Expire_Date", "Plan", "Nickname"])
        else:
            df_users = pd.DataFrame(records)
            if "Nickname" not in df_users.columns: df_users["Nickname"] = ""

        user_row = df_users[df_users["Email"] == email]
        pwd_hash = hash_password(password)
        today = datetime.now().date()

        if is_register:
            if not user_row.empty: return False, "帳號已存在"
            expire_date = today + timedelta(days=TRIAL_DAYS)
            final_nickname = nickname if nickname else email.split("@")[0]
            new_user = {"Email": email, "Sheet_Name": user_sheet_name, "Join_Date": str(today), "Password_Hash": pwd_hash, "Status": "Active", "Expire_Date": str(expire_date), "Plan": "Trial", "Nickname": final_nickname}
            row_data = [new_user["Email"], new_user["Sheet_Name"], new_user["Join_Date"], new_user["Password_Hash"], new_user["Status"], new_user["Expire_Date"], new_user["Plan"], new_user["Nickname"]]
            users_sheet.append_row(row_data)
            book_title = get_sheet_title_safe(user_sheet_name)
            bindings_sheet.append_row([email, user_sheet_name, book_title, "Owner"])
            write_system_log(email, "註冊並建立帳本(Owner)", email, book_title, user_sheet_name)
            return True, new_user

        if is_register: 
             records = users_sheet.get_all_records()
             df_users = pd.DataFrame(records)
             user_row = df_users[df_users["Email"] == email]

        if user_row.empty: return False, "User not found"

        user_info = user_row.iloc[0].to_dict()
        stored_hash = str(user_info.get("Password_Hash", ""))
        
        if stored_hash != "RESET_REQUIRED" and stored_hash != pwd_hash:
            return False, "Password Incorrect"
        
        if pd.isna(user_info.get("Nickname")) or user_info.get("Nickname") == "":
            user_info["Nickname"] = email.split("@")[0]

        b_records = bindings_sheet.get_all_records()
        df_bind = pd.DataFrame(b_records)
        user_books = df_bind[df_bind["Email"] == email]
        
        books_list = []
        if not user_books.empty:
            for _, row in user_books.iterrows():
                role = row.get("Role", row.get("Owner", "Member"))
                books_list.append({"name": row["Book_Name"], "url": row["Sheet_URL"], "role": role})
        else:
            books_list.append({"name": "我的記帳本", "url": user_info.get("Sheet_Name", ""), "role": "Owner"})
        
        user_info["Books"] = books_list
        
        if user_info["Plan"] == "VIP": return True, user_info
        try:
            expire_dt = datetime.strptime(user_info["Expire_Date"], "%Y-%m-%d").date()
            if today > expire_dt: return False, "Expired"
            else: return True, user_info
        except: return False, "Date Error"

    except Exception as e: return False, f"Login Error: {e}"

def add_binding(target_email, sheet_url, book_name, role="Member", operator_email=None):
    client = get_gspread_client()
    try:
        admin_book = client.open_by_url(st.secrets["admin_sheet_url"])
        users_sheet = admin_book.worksheet("Users")
        bindings_sheet = admin_book.worksheet("Book_Bindings")
        
        # 1. 檢查使用者是否存在
        try: cell = users_sheet.find(target_email)
        except: cell = None

        if not cell:
            today = str(datetime.now().date())
            row = [target_email, "", today, "RESET_REQUIRED", "Pending", today, "Trial", target_email.split("@")[0]]
            users_sheet.append_row(row)
        
        # 2. 檢查是否已經綁定
        existing = bindings_sheet.get_all_records()
        df = pd.DataFrame(existing)
        if not df.empty:
            check = df[(df["Email"] == target_email) & (df["Sheet_URL"] == sheet_url)]
            if not check.empty: return True, "該使用者已經在此帳本中，無需重複邀請"
        
        # 3. 檢查 Owner 唯一性
        if role == "Owner":
            if not df.empty:
                owner_check = df[(df["Sheet_URL"] == sheet_url) & (df["Role"] == "Owner")]
                if not owner_check.empty: return False, "❌ 此帳本已經有擁有者"

        # 4. 寫入綁定
        bindings_sheet.append_row([target_email, sheet_url, book_name, role])
        
        # 5. 寫入 Log
        op = operator_email if operator_email else "System"
        action = "新增綁定" if role == "Owner" else "邀請成員"
        write_system_log(op, action, target_email, book_name, sheet_url)
        
        # 6. [修改] 執行寄信 (抓取暱稱)
        status_msg = "綁定成功！"
        
        if role == "Member":
            if operator_email:
                # 嘗試從 Session State 抓取當前操作者的暱稱
                current_nick = None
                if "user_info" in st.session_state:
                    # 確保 Session 中的人就是操作者 (通常是的)
                    if st.session_state.user_info.get("Email") == operator_email:
                        current_nick = st.session_state.user_info.get("Nickname")
                
                # 呼叫寄信函式，傳入暱稱
                is_sent, mail_msg = send_invitation_email(target_email, operator_email, book_name, inviter_nickname=current_nick)
                
                if is_sent:
                    status_msg += " (邀請信已寄出 ✅)"
                else:
                    status_msg += f" (但寄信失敗 ❌: {mail_msg})"
            else:
                status_msg += " (未寄信: 缺少操作者 Email)"
        
        return True, status_msg

    except Exception as e: return False, f"系統錯誤: {e}"

def remove_binding_from_db(target_email, sheet_url, operator_email=None, book_name="Unknown"):
    client = get_gspread_client()
    try:
        admin_book = client.open_by_url(st.secrets["admin_sheet_url"])
        bindings_sheet = admin_book.worksheet("Book_Bindings")
        records = bindings_sheet.get_all_records()
        row_to_delete = None
        for i, row in enumerate(records):
            if row["Email"] == target_email and row["Sheet_URL"] == sheet_url:
                row_to_delete = i + 2; book_name = row["Book_Name"]; break
        if row_to_delete:
            bindings_sheet.delete_rows(row_to_delete)
            op = operator_email if operator_email else target_email
            write_system_log(op, "解除綁定/移除成員", target_email, book_name, sheet_url)
            return True, "解除綁定成功"
        else: return False, "找不到該綁定資料"
    except Exception as e: return False, f"刪除失敗: {e}"

# [新增] 移轉擁有權函式
def transfer_book_ownership(sheet_url, old_owner_email, new_owner_email, book_name="Unknown"):
    client = get_gspread_client()
    try:
        admin_book = client.open_by_url(st.secrets["admin_sheet_url"])
        bindings_sheet = admin_book.worksheet("Book_Bindings")
        records = bindings_sheet.get_all_records()
        
        row_old = None
        row_new = None
        
        # 尋找兩位的資料列 (Gspread index 從 1 開始，標題是 1，資料從 2 開始)
        for i, row in enumerate(records):
            if row["Sheet_URL"] == sheet_url:
                if row["Email"] == old_owner_email:
                    row_old = i + 2
                elif row["Email"] == new_owner_email:
                    row_new = i + 2
        
        if row_old and row_new:
            # 假設 Role 是第 4 欄 (D)
            bindings_sheet.update_cell(row_old, 4, "Member")
            bindings_sheet.update_cell(row_new, 4, "Owner")
            
            write_system_log(old_owner_email, "移轉擁有權", new_owner_email, book_name, sheet_url)
            return True, "移轉成功！您已成為成員。"
        else:
            return False, "資料庫讀取錯誤，找不到成員資料"
            
    except Exception as e: return False, f"移轉失敗: {e}"

def get_book_members(sheet_url):
    client = get_gspread_client()
    try:
        admin_book = client.open_by_url(st.secrets["admin_sheet_url"])
        bindings_sheet = admin_book.worksheet("Book_Bindings")
        records = bindings_sheet.get_all_records()
        df = pd.DataFrame(records)
        if not df.empty:
            members = df[df["Sheet_URL"] == sheet_url]
            return members.to_dict('records')
        return []
    except: return []

# ==========================================
# 登入流程 (含 OTP 註冊驗證)
# ==========================================
def login_flow():
    if "is_logged_in" in st.session_state and st.session_state.is_logged_in:
        user_books = st.session_state.user_info.get("Books", [])
        if "current_book_url" not in st.session_state:
            if user_books:
                st.session_state.current_book_url = user_books[0]["url"]
                st.session_state.current_book_name = user_books[0]["name"]
            else:
                st.session_state.current_book_url = st.session_state.user_info["Sheet_Name"]
                st.session_state.current_book_name = "我的記帳本"
        return st.session_state.current_book_url, st.session_state.current_book_name

    if "login_mode" not in st.session_state: st.session_state.login_mode = "login"
    if "reset_stage" not in st.session_state: st.session_state.reset_stage = 1
    if "reg_stage" not in st.session_state: st.session_state.reg_stage = 1
    if "otp_code" not in st.session_state: st.session_state.otp_code = ""
    if "reset_email" not in st.session_state: st.session_state.reset_email = ""
    
    if "reg_data" not in st.session_state: st.session_state.reg_data = {}

    st.markdown("""<div class="login-container"><h2>👋 歡迎使用記帳本</h2>""", unsafe_allow_html=True)
    
    if st.session_state.login_mode == "reset":
        if st.button("⬅️ 返回登入", use_container_width=True):
            st.session_state.login_mode = "login"; st.rerun()
        st.markdown("#### 🔒 重設密碼 / 啟用帳號")
    elif st.session_state.login_mode == "register":
         if st.button("⬅️ 返回登入", use_container_width=True):
            st.session_state.login_mode = "login"; st.rerun()
    else:
        c1, c2 = st.columns(2)
        with c1:
            if st.button("登入", use_container_width=True, type="primary" if st.session_state.login_mode == "login" else "secondary"):
                st.session_state.login_mode = "login"; st.rerun()
        with c2:
            if st.button("註冊", use_container_width=True, type="primary" if st.session_state.login_mode == "register" else "secondary"):
                st.session_state.login_mode = "register"; st.session_state.reg_stage = 1; st.rerun()

    with st.container():
        # === 忘記密碼 / 啟用帳號 ===
        if st.session_state.login_mode == "reset":
            if st.session_state.reset_stage == 1:
                st.info("請輸入 Email，我們將發送驗證碼給您。")
                email_reset = st.text_input("註冊信箱", key="reset_input_email").strip()
                if st.button("📩 發送驗證碼", type="primary", use_container_width=True):
                    if not email_reset: st.warning("請輸入 Email")
                    else:
                        code = ''.join(random.choices(string.digits, k=6))
                        st.session_state.otp_code = code; st.session_state.reset_email = email_reset
                        with st.spinner("寄送中..."):
                            ok, msg = send_otp_email(email_reset, code)
                            if ok: st.session_state.reset_stage = 2; st.success("✅ 已發送！"); time.sleep(1); st.rerun()
                            else: st.error(msg)
            elif st.session_state.reset_stage == 2:
                st.success(f"驗證碼已寄至 {st.session_state.reset_email}")
                otp_input = st.text_input("輸入 6 位數驗證碼", key="otp_input")
                new_pwd = st.text_input("設定新密碼", type="password", key="reset_new_pwd")
                new_nick = st.text_input("設定您的暱稱 (若為初次啟用請填寫)", key="reset_new_nick")
                
                if st.button("🔄 確認重設", type="primary", use_container_width=True):
                    if otp_input == st.session_state.otp_code and new_pwd:
                        ok, msg = reset_user_password(st.session_state.reset_email, new_pwd, new_nickname=new_nick)
                        if ok: 
                            st.success("🎉 帳號設定成功，請重新登入")
                            st.session_state.login_mode = "login"
                            st.session_state.reset_stage = 1
                            time.sleep(2); st.rerun()
                        else: st.error(msg)
                    else: st.error("驗證碼錯誤或密碼為空")

        # === 註冊 (含 OTP) ===
        elif st.session_state.login_mode == "register":
            if st.session_state.reg_stage == 1:
                st.info("💡 新用戶請先設定您的記帳本 (需 Email 驗證)")
                with st.expander("👉 點此查看設定步驟 (含圖文教學)"):
                    st.markdown(f"**步驟 1：建立記帳本副本** 👉 [**[點此建立]**]({TEMPLATE_URL})")
                    st.markdown("---")        
                    st.markdown("**步驟 2：共用權限給機器人**")
                    st.write("請共用給以下 Email (權限設為 **編輯者/Editor**)")
                    if "gcp_service_account" in st.secrets:
                        st.code(st.secrets["gcp_service_account"]["client_email"], language="text")
                    if os.path.exists("guide.png"):
                        with st.expander("📷 操作示意圖"): st.image("guide.png", caption="共用設定示意圖", use_container_width=True)

                email_in = st.text_input("Email", key="reg_email").strip()
                pwd_in = st.text_input("密碼", type="password", key="reg_pwd")
                nick_in = st.text_input("暱稱 (用於交易記錄)", key="reg_nick")
                sheet_in = st.text_input("Google Sheet 網址", key="reg_sheet")
                
                if st.button("📩 驗證 Email 並下一步", type="primary", use_container_width=True):
                    if email_in and pwd_in and sheet_in and nick_in:
                        if not is_valid_email(email_in):
                            st.error("❌ Email 格式不正確")
                        else:
                            st.cache_data.clear() 
                            with st.spinner("檢查帳戶狀態中..."):
                                is_valid, msg = validate_registration_pre_check(email_in, sheet_in)
                            if not is_valid: st.error(msg)
                            else:
                                code = ''.join(random.choices(string.digits, k=6))
                                st.session_state.otp_code = code
                                st.session_state.reg_data = {"email": email_in, "pwd": pwd_in, "nick": nick_in, "sheet": sheet_in}
                                with st.spinner("寄送驗證碼中..."):
                                    ok, msg = send_otp_email(email_in, code, subject="【記帳本】註冊驗證碼")
                                    if ok: st.session_state.reg_stage = 2; st.success("✅ 驗證碼已發送！"); time.sleep(1); st.rerun()
                                    else: st.error(msg)
                    else: st.warning("請填寫所有欄位")
            
            elif st.session_state.reg_stage == 2:
                reg_d = st.session_state.reg_data
                st.success(f"驗證碼已發送至：{reg_d['email']}")
                otp_input = st.text_input("輸入 6 位數驗證碼", key="reg_otp_input")
                
                if st.button("✨ 確認註冊", type="primary", use_container_width=True):
                    if otp_input == st.session_state.otp_code:
                        with st.spinner("建立帳戶中..."):
                            success, result = handle_user_login(reg_d["email"], reg_d["pwd"], reg_d["sheet"], nickname=reg_d["nick"], is_register=True)
                            if success: st.session_state.is_logged_in = True; st.session_state.user_info = result; st.success("註冊成功！"); time.sleep(1); st.rerun()
                            else: st.error(f"註冊失敗：{result}")
                    else: st.error("❌ 驗證碼錯誤")
                if st.button("返回修改資料"): st.session_state.reg_stage = 1; st.rerun()

        # === 登入 ===
        else:
            email_in = st.text_input("Email", key="login_email").strip()
            pwd_in = st.text_input("密碼", type="password", key="login_pwd")
            if st.button("🚀 登入", type="primary", use_container_width=True):
                if email_in and pwd_in:
                    with st.spinner("登入中..."):
                        success, result = handle_user_login(email_in, pwd_in, is_register=False)
                        if success: st.session_state.is_logged_in = True; st.session_state.user_info = result; st.rerun()
                        else: st.error(f"登入失敗: {result}")
            if st.button("🔑 忘記密碼？ (或啟用被邀請的帳號)", type="tertiary"):
                st.session_state.login_mode = "reset"; st.session_state.reset_stage = 1; st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

CURRENT_SHEET_SOURCE, DISPLAY_TITLE = login_flow()

# ============ Header ============
c_logo, c_title = st.columns([1, 15]) 
with c_logo:
    if os.path.exists("logo.png"): st.image("logo.png", width=60) 
    else: st.write("💰")
with c_title:
    st.markdown("<h2 style='margin-bottom: 0; padding-top: 10px;'>我的記帳本</h2>", unsafe_allow_html=True)

# ... (Data Functions) ...
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
        # [修正] 同時讀取 Transactions 與 Transactions_History
        target_sheets = ["Transactions", "Transactions_History"]

        for ws in sheet.worksheets():
            # 只要是開頭符合的都抓 (或是明確指定那兩個)
            if ws.title in target_sheets or ws.title.startswith("Transactions"):
                data = ws.get_all_records()
                if data: all_data.extend(data)

        df = pd.DataFrame(all_data)
        if not df.empty:
            df = df.dropna(how='all')
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df['Amount_Def'] = pd.to_numeric(df['Amount_Def'], errors='coerce').fillna(0)
            df['Year'] = df['Date'].dt.year; df['Month'] = df['Date'].dt.strftime('%Y-%m')
            if "Recorder" not in df.columns: df["Recorder"] = ""
        return df
    except: return pd.DataFrame()

def append_data(worksheet_name, row_data, source_str):
    client = get_gspread_client()
    try:
        sheet = open_spreadsheet(client, source_str)
        worksheet = sheet.worksheet(worksheet_name)
        if worksheet_name == "Transactions":
            recorder = st.session_state.user_info.get("Nickname", st.session_state.user_info.get("Email"))
            row_data.append(recorder)
        worksheet.append_row(row_data)
        return True
    except: return False

def save_settings_data(new_settings_df, source_str):
    client = get_gspread_client()
    try:
        sheet = open_spreadsheet(client, source_str)
        worksheet = sheet.worksheet("Settings"); worksheet.clear()
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
        dfs = pd.read_html(url); df = dfs[0]; df = df.iloc[:, 0:5]
        df.columns = ["Currency_Name", "Cash_Buy", "Cash_Sell", "Spot_Buy", "Spot_Sell"]
        df["Currency"] = df["Currency_Name"].str.extract(r'\(([A-Z]+)\)')
        rates = df.dropna(subset=['Currency']).copy()
        rates["Spot_Sell"] = pd.to_numeric(rates["Spot_Sell"], errors='coerce')
        rate_dict = rates.set_index("Currency")["Spot_Sell"].to_dict(); rate_dict["TWD"] = 1.0
        return rate_dict
    except: return {}

def calculate_exchange(amount, input_currency, target_currency, rates):
    if input_currency == target_currency: return amount, 1.0
    try:
        rate_in = rates.get(input_currency); rate_target = rates.get(target_currency)
        if not rate_in or not rate_target: return amount, 0
        conversion_factor = rate_in / rate_target
        exchanged_amount = amount * conversion_factor
        return round(exchanged_amount, 2), conversion_factor
    except: return amount, 0

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

def add_sub_callback(main_cat, key):
    new_val = st.session_state[key]
    if new_val:
        if new_val not in st.session_state.temp_cat_map[main_cat]: st.session_state.temp_cat_map[main_cat].append(new_val)
        st.session_state[key] = "" 
def add_pay_callback(key):
    new_val = st.session_state[key]
    if new_val and new_val not in st.session_state.temp_pay_list: st.session_state.temp_pay_list.append(new_val); st.session_state[key] = ""
def add_curr_callback(key):
    new_val = st.session_state[key]
    if new_val and new_val not in st.session_state.temp_curr_list: st.session_state.temp_curr_list.append(new_val); st.session_state[key] = ""

# --- 側邊欄 ---
with st.sidebar:
    st.header("🌍 地區與帳號")
    user_info = st.session_state.get("user_info", {})
    plan = user_info.get("Plan", "Trial")
    nickname_display = user_info.get("Nickname", "")
    if not nickname_display: nickname_display = user_info.get("Email", "訪客").split("@")[0]
    
    tz_options = {"台灣/北京 (UTC+8)": 8, "日本/韓國 (UTC+9)": 9, "泰國 (UTC+7)": 7, "美東 (UTC-4)": -4, "歐洲 (UTC+1)": 1}
    selected_tz_label = st.selectbox("當前位置時區", list(tz_options.keys()), index=0)
    user_offset = tz_options[selected_tz_label]; today_date = get_user_date(user_offset)
    st.info(f"日期：{today_date}")

    user_books = user_info.get("Books", [])
    if len(user_books) > 0:
        book_names = [b["name"] for b in user_books]
        try: curr_idx = next(i for i, v in enumerate(user_books) if v["url"] == CURRENT_SHEET_SOURCE)
        except: curr_idx = 0
        selected_book_name = st.selectbox("📘 切換帳本", book_names, index=curr_idx)
        new_url = next(b["url"] for b in user_books if b["name"] == selected_book_name)
        if new_url != CURRENT_SHEET_SOURCE:
            st.session_state.current_book_url = new_url; st.session_state.current_book_name = selected_book_name
            st.cache_data.clear(); st.rerun()
    else: st.success(f"📘 帳本：{DISPLAY_TITLE}")

    if plan == "VIP": st.markdown(f"👤 **{nickname_display}** <span class='vip-badge'>  VIP</span>", unsafe_allow_html=True)
    else:
        expire_str = user_info.get("Expire_Date", str(today_date))
        try: expire_dt = datetime.strptime(expire_str, "%Y-%m-%d").date(); days_left = (expire_dt - today_date).days
        except: days_left = 0
        st.markdown(f"👤 **{nickname_display}** <span class='trial-badge'>  {plan}</span>", unsafe_allow_html=True)
        if days_left > 0: st.caption(f"⏳ 試用倒數：**{days_left}** 天"); st.progress(min(days_left / 30, 1.0))
        else: st.error(f"⛔ 試用期已結束")

    if plan != "VIP":
        #st.info("##### 🚀 升級持續使用")
        if st.button("💎 升級 VIP 持續使用", type="primary", use_container_width=True): st.toast("🚧 金流功能開發中")
    st.divider()
    if st.button("🚪 登出"):
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.query_params.clear(); st.rerun()

rates = get_exchange_rates()

# --- 讀取設定 ---
settings_df = get_data("Settings", CURRENT_SHEET_SOURCE)
cat_mapping = {}; payment_list = []; currency_list_custom = []; default_currency_setting = "TWD" 
if not settings_df.empty:
    if "Main_Category" in settings_df.columns and "Sub_Category" in settings_df.columns:
        valid_cats = settings_df[["Main_Category", "Sub_Category"]].astype(str)
        valid_cats = valid_cats[valid_cats["Main_Category"] != ""]
        for _, row in valid_cats.iterrows():
            main = row["Main_Category"]; sub = row["Sub_Category"]
            if main not in cat_mapping: cat_mapping[main] = []
            if sub and sub != "" and sub not in cat_mapping[main]: cat_mapping[main].append(sub)
    if "Payment_Method" in settings_df.columns: payment_list = settings_df[settings_df["Payment_Method"] != ""]["Payment_Method"].unique().tolist()
    if "Currency" in settings_df.columns: currency_list_custom = settings_df[settings_df["Currency"] != ""]["Currency"].unique().tolist()
    if "Default_Currency" in settings_df.columns:
        saved = settings_df[settings_df["Default_Currency"] != ""]["Default_Currency"].unique().tolist()
        if saved: default_currency_setting = saved[0]
if not cat_mapping: cat_mapping = {"收入": ["薪資"], "食": ["早餐"]}
if not payment_list: payment_list = ["現金"]
if not currency_list_custom: currency_list_custom = ["TWD"]
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
    else: final_df["Main_Category"] = [""]*max_len; final_df["Sub_Category"] = [""]*max_len
    final_df["Payment_Method"] = pd.Series(list_pay).reindex(range(max_len)).fillna("")
    final_df["Currency"] = pd.Series(list_curr).reindex(range(max_len)).fillna("")
    final_df["Default_Currency"] = ""
    if len(final_df) > 0: final_df.at[0, "Default_Currency"] = st.session_state.get('temp_default_curr', default_currency_setting)
    if save_settings_data(final_df, CURRENT_SHEET_SOURCE): st.toast("✅ 設定已儲存！", icon="💾"); st.cache_data.clear()

check_and_run_recurring()

# --- Tabs Content ---
tab1, tab2, tab3 = st.tabs(["📝 每日記帳", "📊 收支分析", "⚙️ 系統設定"])

with tab1:
    if st.session_state.get('should_clear_input'):
        st.session_state.form_amount_org = 0.0; st.session_state.form_amount_def = 0.0; st.session_state.form_note = ""; st.session_state.should_clear_input = False
    if 'form_currency' not in st.session_state: st.session_state.form_currency = default_currency_setting
    if 'form_amount_org' not in st.session_state: st.session_state.form_amount_org = 0.0
    if 'form_amount_def' not in st.session_state: st.session_state.form_amount_def = 0.0
    
    def on_input_change():
        c = st.session_state.form_currency; a = st.session_state.form_amount_org
        val, _ = calculate_exchange(a, c, default_currency_setting, rates)
        st.session_state.form_amount_def = val

    user_today = today_date 
    current_month_str = user_today.strftime("%Y-%m")

    tx_df = get_data("Transactions", CURRENT_SHEET_SOURCE)
    total_inc = 0; total_exp = 0
    if not tx_df.empty and 'Date' in tx_df.columns:
        tx_df['Date'] = pd.to_datetime(tx_df['Date'], errors='coerce')
        mask = (tx_df['Date'].dt.strftime('%Y-%m') == current_month_str)
        mtx = tx_df[mask]; mtx['Amount_Def'] = pd.to_numeric(mtx['Amount_Def'], errors='coerce').fillna(0)
        total_inc = mtx[mtx['Type']=='收入']['Amount_Def'].sum()
        total_exp = mtx[mtx['Type']!='收入']['Amount_Def'].sum()
    bal = total_inc - total_exp; b_cls = "val-green" if bal >= 0 else "val-red"

    st.markdown(f"""
    <div class="metric-container">
        <div class="metric-card"><span class="metric-label">本月總收入 ({default_currency_setting})</span><span class="metric-value">${total_inc:,.2f}</span></div>
        <div class="metric-card"><span class="metric-label">已支出 ({default_currency_setting})</span><span class="metric-value">${total_exp:,.2f}</span></div>
        <div class="metric-card"><span class="metric-label">剩餘可用</span><span class="metric-value {b_cls}">${bal:,.2f}</span></div>
    </div>""", unsafe_allow_html=True)

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
            try: ci = currency_list_custom.index(default_currency_setting)
            except: ci = 0
            with c5: currency = st.selectbox("幣別", currency_list_custom, index=ci, key="form_currency", on_change=on_input_change)
            with c6: amount_org = st.number_input(f"金額 ({currency})", step=1.0, key="form_amount_org", on_change=on_input_change)
            with c7: amount_def = st.number_input(f"折合 {default_currency_setting}", step=0.1, key="form_amount_def")
        
        note = st.text_input("備註", max_chars=20, key="form_note"); st.markdown("<br>", unsafe_allow_html=True)
        if st.button("確認送出記帳", type="primary", use_container_width=True):
            if amount_def == 0: st.error("金額不能為 0")
            else:
                with st.spinner('📡 資料寫入中...'):
                    tx_type = "收入" if main_cat == "收入" else "支出"
                    row = [str(date_input), tx_type, main_cat, sub_cat, payment, currency, amount_org, amount_def, note, str(datetime.now())]
                    if append_data("Transactions", row, CURRENT_SHEET_SOURCE):
                        st.success(f"✅ 已記錄！"); st.session_state['should_clear_input'] = True; st.cache_data.clear(); time.sleep(1); st.rerun()
                    else: st.error("❌ 寫入失敗")

# ================= Tab 2: 收支分析 =================
with tab2:
    st.markdown("##### 📊 收支狀況")
    df_tx = get_data("Transactions", CURRENT_SHEET_SOURCE)

    if df_tx.empty:
        st.info("尚無交易資料")
    else:
        df_tx['Date'] = pd.to_datetime(df_tx['Date'], errors='coerce')
        df_tx['Amount_Def'] = pd.to_numeric(df_tx['Amount_Def'], errors='coerce').fillna(0)
        df_tx['Month'] = df_tx['Date'].dt.strftime('%Y-%m')
        df_tx['Year'] = df_tx['Date'].dt.year
        
        all_months = sorted(df_tx['Month'].unique())

       # 1. 年度趨勢比較圖 (包含 Transactions + Transactions_History)
        all_years = sorted(df_tx['Year'].dropna().unique().astype(int))
        
        with st.expander("📅 篩選年度區間 (比較總收入/總支出)", expanded=True):
            if len(all_years) > 0:
                mn, mx = int(min(all_years)), int(max(all_years))
                # 使用 slider 選擇區間
                sel_y = st.slider("選擇年份範圍", mn, mx, (mn, mx))
                
                # 篩選資料
                df_trend = df_tx[(df_tx['Year'] >= sel_y[0]) & (df_tx['Year'] <= sel_y[1])]
                
                # 分組計算
                trend_group = df_trend.groupby(['Year', 'Type'])['Amount_Def'].sum().reset_index()
                
                # 為了顯示漂亮，可以把 Type 重新命名或排序
                if not trend_group.empty:
                    fig_trend = px.bar(
                        trend_group, 
                        x="Year", 
                        y="Amount_Def", 
                        color="Type", 
                        barmode="group",
                        title=f"{sel_y[0]} - {sel_y[1]} 收支趨勢比較",
                        labels={"Amount_Def": f"金額 ({default_currency_setting})", "Year": "年份"},
                        color_discrete_map={"收入": "#2ecc71", "支出": "#ff6b6b"}
                    )
                    fig_trend.update_layout(xaxis=dict(tickmode='linear')) # 強制顯示所有年份
                    st.plotly_chart(fig_trend, use_container_width=True)
                else:
                    st.info("選定區間無數據")


        # st.markdown("---")
        with st.expander("🗓️ 查看詳細月份", expanded=True):
            target_month = st.selectbox("選擇月份", sorted(all_months, reverse=True))
            
            month_data = df_tx[df_tx['Month'] == target_month]
            monthly_income = month_data[month_data['Type'] == '收入']['Amount_Def'].sum()
            monthly_expense = month_data[month_data['Type'] != '收入']['Amount_Def'].sum()
            
            st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card" style="border-left: 5px solid #2ecc71;">
                    <span class="metric-label">總收入 ({default_currency_setting})</span>
                    <span class="metric-value">${monthly_income:,.2f}</span>
                </div>
                <div class="metric-card" style="border-left: 5px solid #ff6b6b;">
                    <span class="metric-label">總支出 ({default_currency_setting})</span>
                    <span class="metric-value">${monthly_expense:,.2f}</span>
                </div>
                <div class="metric-card">
                    <span class="metric-label">結餘</span>
                    <span class="metric-value">${monthly_income - monthly_expense:,.2f}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            expense_only_data = month_data[month_data['Type'] != '收入']
            if not expense_only_data.empty:
                pie_data = expense_only_data.groupby("Main_Category")["Amount_Def"].sum().reset_index()
                pie_data = pie_data[pie_data["Amount_Def"] > 0]
                
                if not pie_data.empty:
                    fig_pie = px.pie(pie_data, values="Amount_Def", names="Main_Category", hole=0.5,
                                    color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig_pie.update_layout(margin=dict(t=20, b=20, l=20, r=20))
                    st.plotly_chart(fig_pie, use_container_width=True)
                else:
                    st.info("本月支出相抵後無正向金額，無法顯示圓餅圖。")
                
        # [新增] 除錯用明細表
        with st.expander("🔍 檢視本月明細 (除錯用)"):
            debug_df = month_data[['Date', 'Main_Category', 'Sub_Category', 'Amount_Original', 'Currency', 'Amount_Def', 'Note']].sort_values(by='Date', ascending=False)
            st.dataframe(debug_df, use_container_width=True)

# ================= Tab 3: 設定管理 =================
with tab3:
    st.markdown("##### ⚙️ 系統資料庫")
    if 'temp_cat_map' not in st.session_state: st.session_state.temp_cat_map = cat_mapping
    if 'temp_pay_list' not in st.session_state: st.session_state.temp_pay_list = payment_list
    if 'temp_curr_list' not in st.session_state: st.session_state.temp_curr_list = currency_list_custom
    if 'temp_default_curr' not in st.session_state: st.session_state.temp_default_curr = default_currency_setting

    with st.expander("📚 帳本與成員管理", expanded=True):
        
        # 0. 個人資料設定
        st.markdown("###### 👤 個人資料設定")
        c_nick_in, c_nick_btn = st.columns([3, 1])
        current_nick = st.session_state.user_info.get("Nickname", "")
        new_nick_val = c_nick_in.text_input("修改顯示暱稱", value=current_nick, label_visibility="collapsed")
     
        if c_nick_btn.button("💾 儲存暱稱"):
            if new_nick_val and new_nick_val != current_nick:
                with st.spinner("更新中..."):
                    ok, msg = update_user_nickname(st.session_state.user_info["Email"], new_nick_val)
                    if ok:
                        # [修正] 1. 更新 Session State
                        st.session_state.user_info["Nickname"] = new_nick_val
                        # [修正] 2. 清除成員列表快取，確保下方列表顯示新暱稱
                        get_all_users_nickname_map.clear()
                        # [修正] 3. 清除全域資料快取，以防其他依賴
                        st.cache_data.clear()
                        st.success(msg)
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg)

        user_books = st.session_state.user_info.get("Books", [])
        
        if not user_books:
            st.info("目前尚無綁定任何帳本")
        else:
            c_sel, c_btn = st.columns([3, 1])
            with c_sel:
                book_names = [b["name"] for b in user_books]
                try: default_idx = next(i for i, b in enumerate(user_books) if b["url"] == CURRENT_SHEET_SOURCE)
                except: default_idx = 0
                selected_manage_book_name = st.selectbox("選擇要管理的帳本", book_names, index=default_idx, key="manage_book_sel")
            target_book = next((b for b in user_books if b["name"] == selected_manage_book_name), None)
            target_role = target_book.get("role", "Member")
            target_url = target_book.get("url", "")
            with c_btn:
                st.write(""); st.write("") 
                is_owner = (target_role == "Owner")
                btn_label = "無法解除" if is_owner else "❌ 解除綁定"
                btn_help = "擁有者無法解除綁定，請聯絡管理員" if is_owner else "退出此帳本"
                
                if st.button(btn_label, key="top_unbind_btn", disabled=is_owner, type="secondary", help=btn_help, use_container_width=True):
                    with st.spinner("處理中..."):
                        ok, msg = remove_binding_from_db(
                            st.session_state.user_info["Email"], 
                            target_url, 
                            operator_email=st.session_state.user_info["Email"], 
                            book_name=selected_manage_book_name
                        )
                        if ok:
                            st.success(f"已退出 {selected_manage_book_name}")
                            time.sleep(1)
                            st.cache_data.clear()
                            if target_url == st.session_state.get("current_book_url"):
                                del st.session_state["current_book_url"]
                            st.rerun()
                        else:
                            st.error(msg)

            members = get_book_members(target_url)
            nickname_map = get_all_users_nickname_map()

            if members:
                st.caption(f"共 {len(members)} 位成員")
                my_email = st.session_state.user_info["Email"]

                for idx, m in enumerate(members):
                    # 【UI 重點】使用 container(border=True) 建立卡片感
                    with st.container(border=True):
                        # 將卡片分為：[左側資訊區 (70%)] [右側操作區 (30%)]
                        c_info, c_action = st.columns([0.7, 0.3])
                        
                        # --- 左側：資訊區 ---
                        with c_info:
                            is_me = (m["Email"] == my_email)
                            nick = nickname_map.get(m["Email"], "-")
                            role = m.get("Role", "Member")
                            
                            # 第一行：暱稱 + 角色圖示
                            if role == "Owner":
                                st.markdown(f"**{nick}** <span style='background:#FFF3CD; color:#856404; padding:2px 6px; border-radius:4px; font-size:0.8em;'>👑 擁有者</span>", unsafe_allow_html=True)
                            else:
                                st.markdown(f"**{nick}**", unsafe_allow_html=True)
                            
                            # 第二行：Email (使用 caption 縮小字體，適合手機閱讀)
                            display_email = f"{mask_email(m['Email'])} (自己)" if is_me else mask_email(m["Email"])
                            st.caption(f"📧 {display_email}")

                        # --- 右側：操作區 (收納進 Popover) ---
                        with c_action:
                            # 垂直置中調整 (Streamlit 小技巧)
                            st.write("") 
                            
                            # 判斷權限
                            # 只有 Owner 可以管理其他人
                            if target_role == "Owner":
                                if not is_me:
                                    # 使用 Popover 收納按鈕，解決手機版按鈕過大問題
                                    with st.popover("⚙️ 管理", use_container_width=True):
                                        st.write(f"對 {nick} 執行操作：")
                                        
                                        # 移除按鈕
                                        if st.button("🚫 移除成員", key=f"kick_{idx}", use_container_width=True):
                                            ok, msg = remove_binding_from_db(m["Email"], target_url, operator_email=my_email, book_name=selected_manage_book_name)
                                            if ok: st.toast("移除成功"); time.sleep(1); st.rerun()
                                            else: st.error(msg)
                                        
                                        # 移轉按鈕
                                        with st.expander("👑 移轉擁有權"):
                                            st.warning("移轉後您將變為普通成員！")
                                            if st.button("確認移轉", key=f"transfer_{idx}", use_container_width=True):
                                                with st.spinner("處理中..."):
                                                    ok, msg = transfer_book_ownership(target_url, my_email, m["Email"], book_name=selected_manage_book_name)
                                                    if ok:
                                                        st.success(msg)
                                                        st.cache_data.clear()
                                                        time.sleep(2)
                                                        st.rerun()
                                                    else:
                                                        st.error(msg)
                                else:
                                    # 自己是 Owner
                                    st.caption("您是擁有者")

                            elif target_role == "Member":
                                if is_me:
                                    if st.button("🚪 退出", key=f"leave_{idx}", type="primary", use_container_width=True):
                                        ok, msg = remove_binding_from_db(my_email, target_url, operator_email=my_email, book_name=selected_manage_book_name)
                                        if ok: 
                                            st.success("已退出"); time.sleep(1); st.cache_data.clear()
                                            if target_url == st.session_state.get("current_book_url"): del st.session_state["current_book_url"]
                                            st.rerun()
                                        else: st.error(msg)
                                else:
                                    # Member 看別人 -> 無權限
                                    st.caption("成員")

            else:
                st.caption("無法讀取成員列表")
        
        c_inv, c_book = st.columns(2)
        with c_inv:
            with st.popover("➕ 邀請成員加入此帳本", use_container_width=True):
                st.write("請輸入對方的註冊 Email")
                invite_email = st.text_input("對方 Email")
                if st.button("發送邀請"):
                    target_book_invite = next((b for b in user_books if b["name"] == selected_manage_book_name), None)
                    if target_book_invite:
                        if invite_email:
                            ok, msg = add_binding(invite_email, target_book_invite["url"], selected_manage_book_name, role="Member", operator_email=st.session_state.user_info["Email"])
                            if ok: 
                                st.success(msg)
                                time.sleep(1)
                                st.rerun()
                            else: st.error(msg)
                        else: st.warning("請輸入 Email")
        with c_book:
            with st.popover("➕ 綁定其他帳本", use_container_width=True):
                st.write("輸入 Google Sheet 網址以新增帳本")
                new_sheet_url = st.text_input("Google Sheet 網址")
                new_book_name = st.text_input("帳本名稱")
                if st.button("確認綁定"):
                    if new_sheet_url and new_book_name:
                        ok, msg = add_binding(st.session_state.user_info["Email"], new_sheet_url, new_book_name, "Owner", operator_email=st.session_state.user_info["Email"])
                        if ok: 
                            st.success("綁定成功！請重新登入生效"); time.sleep(2); st.cache_data.clear(); st.rerun()
                        else: st.error(msg)
    
    with st.expander("🔄 每月固定收支"):
        with st.popover("➕ 新增固定規則", use_container_width=True):
            if 'rec_currency' not in st.session_state: st.session_state.rec_currency = default_currency_setting
            if 'rec_amount_org' not in st.session_state: st.session_state.rec_amount_org = 0.0
            def on_rec_change():
                c = st.session_state.rec_currency; a = st.session_state.rec_amount_org
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
                    st.success("規則已新增"); st.cache_data.clear(); time.sleep(1); st.rerun()
        st.markdown("---")
        rec_df = get_data("Recurring", CURRENT_SHEET_SOURCE)
        if not rec_df.empty:
            for idx, row in rec_df.iterrows():
                with st.expander(f"📅 每月 {row['Day']} 號 - {row['Main_Category']} > {row['Sub_Category']} > {row['Amount_Original']} {row['Currency']}"):
                    c1, c2 = st.columns([4,1])
                    with c1: st.write(f"📝 {row['Note']} ({row['Payment_Method']})")
                    with c2: 
                        if st.button("🗑️", key=f"del_{idx}"):
                             if delete_recurring_rule(idx, CURRENT_SHEET_SOURCE): st.toast("已刪除"); st.cache_data.clear(); time.sleep(1); st.rerun()

    with st.expander("📂 類別與子類別"):
        with st.popover("➕ 新增大類", use_container_width=True):
            nm = st.text_input("類別名稱")
            if st.button("確認"):
                if nm and nm not in st.session_state.temp_cat_map: st.session_state.temp_cat_map[nm] = []; save_all_to_sheet(); st.rerun()
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
                    if st.button(f"🗑️ 刪除 {main}", key=f"dm_{main}"): del st.session_state.temp_cat_map[main]; save_all_to_sheet(); st.rerun()

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
        if nd != st.session_state.temp_default_curr: st.session_state.temp_default_curr = nd; save_all_to_sheet(); st.toast("已更新")
    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("💾 儲存所有設定", type="primary", use_container_width=True): save_all_to_sheet(); st.rerun()