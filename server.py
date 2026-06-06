from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel
from google import genai
from datetime import date, datetime, timedelta
import uvicorn
import json
import os
import hashlib
import uuid
import re
import requests

# ======================================================
# GPT MINI PREMIUM SERVER
# KHÔNG DÁN API KEY THẬT VÀO FILE NÀY.
# Trên Render thêm Environment Variable:
# GEMINI_API_KEY = API key Gemini của bạn
# ADMIN_PASSWORD = mật khẩu admin nếu muốn đổi
# ======================================================

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

SUPABASE_URL = (os.getenv("SUPABASE_URL") or "").rstrip("/")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")

SUPABASE_TABLE = os.getenv("SUPABASE_TABLE", "users")
COL_USERNAME = os.getenv("SUPABASE_COL_USERNAME", "username")
COL_EMAIL = os.getenv("SUPABASE_COL_EMAIL", "email")
COL_PHONE = os.getenv("SUPABASE_COL_PHONE", "phone")
COL_PASSWORD = os.getenv("SUPABASE_COL_PASSWORD", "password")
COL_PLAN = os.getenv("SUPABASE_COL_PLAN", "plan")
COL_TOKEN = os.getenv("SUPABASE_COL_TOKEN", "token")
COL_USED = os.getenv("SUPABASE_COL_USED", "used")
COL_DATE = os.getenv("SUPABASE_COL_DATE", "date")
COL_REGISTERED_AT = os.getenv("SUPABASE_COL_REGISTERED_AT", "registered_at")
COL_EXPIRES_AT = os.getenv("SUPABASE_COL_EXPIRES_AT", "expires_at")

app = FastAPI(title="GPT Mini Premium")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

USERS_FILE = "users.json"
SUPPORT_FILE = "support_messages.json"
FREE_LIMIT = 10
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "admin123")

SUPPORT_ZALO = "036 338 2629"
SUPPORT_EMAIL = "gptminipro@gmail.com"

PLAN_LABELS = {
    "free": "FREE",
    "basic": "CƠ BẢN",
    "pro": "CHUYÊN NGHIỆP",
    "business": "DOANH NGHIỆP",
    "lifetime": "VĨNH VIỄN",
    "premium": "PREMIUM"
}

PREMIUM_CODES = {
    "OPTIMA-BASIC-99000": "basic",
    "OPTIMA-PRO-199000": "pro",
    "OPTIMA-BUSINESS-399000": "business",
    "OPTIMA-LIFETIME-599000": "lifetime",
    "GPTMINI-BASIC-99000": "basic",
    "GPTMINI-PRO-199000": "pro",
    "GPTMINI-BUSINESS-399000": "business",
    "GPTMINI-LIFETIME-599000": "lifetime",
    "GPTMINI-VIP-2026": "business"
}

chat_memory = {}

class RegisterRequest(BaseModel):
    username: str
    password: str
    email: str = ""
    phone: str = ""

class LoginRequest(BaseModel):
    username: str
    password: str

class GoogleLoginRequest(BaseModel):
    name: str = ""
    email: str

class ChatRequest(BaseModel):
    token: str
    message: str

class ActivateRequest(BaseModel):
    token: str
    code: str

class SupportRequest(BaseModel):
    token: str
    message: str

class AdminReplyRequest(BaseModel):
    username: str
    admin_password: str
    message: str

class AdminCloseTicketRequest(BaseModel):
    username: str
    admin_password: str

class AdminApproveRequest(BaseModel):
    username: str
    admin_password: str
    plan: str = "business"

def today() -> str:
    return str(date.today())

def add_days(days: int) -> str:
    return str(date.today() + timedelta(days=days))

def parse_date(value: str):
    if not value:
        return None
    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()
    except Exception:
        return None

def is_expired(user: dict) -> bool:
    plan = user.get("plan", "free")
    if plan in ["free", "lifetime"]:
        return False

    expires_at = parse_date(user.get("expires_at"))
    if not expires_at:
        return False

    return date.today() > expires_at

def days_left(user: dict):
    plan = user.get("plan", "free")
    if plan == "lifetime":
        return "Vĩnh viễn"
    expires_at = parse_date(user.get("expires_at"))
    if not expires_at:
        return ""
    return max(0, (expires_at - date.today()).days)

def safe_username(text: str) -> str:
    return text.strip().lower().replace(" ", "_")

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def password_error_message(password: str) -> str:
    if len(password) < 8:
        return "Mật khẩu phải có tối thiểu 8 ký tự."
    if not re.search(r"[A-Z]", password):
        return "Mật khẩu phải chứa ít nhất 1 chữ in hoa."
    if not re.search(r"[a-z]", password):
        return "Mật khẩu phải chứa ít nhất 1 chữ thường."
    if not re.search(r"[0-9]", password):
        return "Mật khẩu phải chứa ít nhất 1 số."
    if not re.search(r"[@!#$%&*]", password):
        return "Mật khẩu phải chứa ít nhất 1 ký tự đặc biệt như @ ! # $ % & *."
    return ""

def load_json(file_path: str):
    if not os.path.exists(file_path):
        return {}
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_json(file_path: str, data):
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def supabase_enabled() -> bool:
    return bool(SUPABASE_URL and SUPABASE_KEY)

def supabase_headers(prefer: str = ""):
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json"
    }
    if prefer:
        headers["Prefer"] = prefer
    return headers

def supabase_table_url() -> str:
    return f"{SUPABASE_URL}/rest/v1/{SUPABASE_TABLE}"

def normalize_supabase_user(row: dict):
    if not row:
        return None
    return {
        "id": row.get("id") or row.get("nhận dạng"),
        "username": row.get("username") or row.get(COL_USERNAME) or row.get("tên người dùng"),
        "email": row.get("email") or row.get(COL_EMAIL) or row.get("e-mail"),
        "phone": row.get("phone") or row.get(COL_PHONE) or row.get("điện thoại"),
        "password": row.get("password") or row.get(COL_PASSWORD) or row.get("mật khẩu"),
        "token": row.get(COL_TOKEN) or row.get("token"),
        "plan": row.get(COL_PLAN) or row.get("plan") or "free",
        "used": row.get(COL_USED) or row.get("used") or 0,
        "date": row.get(COL_DATE) or row.get("date") or today(),
        "registered_at": row.get(COL_REGISTERED_AT) or row.get("registered_at") or row.get("created_at") or today(),
        "expires_at": row.get(COL_EXPIRES_AT) or row.get("expires_at"),
        "_raw": row
    }

def supabase_find_user_by_username(username: str):
    if not supabase_enabled():
        return None
    try:
        res = requests.get(
            supabase_table_url(),
            headers=supabase_headers(),
            params={"select": "*", COL_USERNAME: f"eq.{username}"},
            timeout=12
        )
        if res.status_code >= 400:
            return None
        data = res.json()
        return normalize_supabase_user(data[0]) if data else None
    except Exception:
        return None

def supabase_email_or_phone_exists(email: str, clean_phone: str):
    if not supabase_enabled():
        return False, ""
    try:
        res = requests.get(supabase_table_url(), headers=supabase_headers(), params={"select": "*"}, timeout=12)
        if res.status_code >= 400:
            return False, ""
        for row in res.json():
            user = normalize_supabase_user(row)
            if not user:
                continue
            if (user.get("email") or "").lower() == email:
                return True, "Gmail này đã được dùng để đăng ký tài khoản khác."
            if (user.get("phone") or "").replace(" ", "").replace("-", "") == clean_phone:
                return True, "Số điện thoại này đã được dùng để đăng ký tài khoản khác."
        return False, ""
    except Exception:
        return False, ""

def supabase_insert_user(user: dict):
    if not supabase_enabled():
        return False, "Chưa cấu hình Supabase."
    payload_full = {
        COL_USERNAME: user["username"],
        COL_EMAIL: user.get("email", ""),
        COL_PHONE: user.get("phone", ""),
        COL_PASSWORD: user.get("password", ""),
        COL_TOKEN: user.get("token", ""),
        COL_PLAN: user.get("plan", "free"),
        COL_USED: user.get("used", 0),
        COL_DATE: user.get("date", today()),
        COL_REGISTERED_AT: user.get("registered_at", today()),
        COL_EXPIRES_AT: user.get("expires_at")
    }
    payload_basic = {
        COL_USERNAME: user["username"],
        COL_EMAIL: user.get("email", ""),
        COL_PHONE: user.get("phone", ""),
        COL_PASSWORD: user.get("password", ""),
        COL_PLAN: user.get("plan", "free"),
        COL_EXPIRES_AT: user.get("expires_at")
    }
    try:
        res = requests.post(supabase_table_url(), headers=supabase_headers("return=minimal"), json=payload_full, timeout=12)
        if res.status_code < 400:
            return True, ""
        res2 = requests.post(supabase_table_url(), headers=supabase_headers("return=minimal"), json=payload_basic, timeout=12)
        if res2.status_code < 400:
            return True, ""
        return False, res2.text
    except Exception as e:
        return False, str(e)

def supabase_update_user(username: str, updates: dict):
    if not supabase_enabled():
        return False
    mapping = {
        "token": COL_TOKEN, "plan": COL_PLAN, "used": COL_USED, "date": COL_DATE,
        "registered_at": COL_REGISTERED_AT, "expires_at": COL_EXPIRES_AT, "password": COL_PASSWORD,
        "email": COL_EMAIL, "phone": COL_PHONE, "username": COL_USERNAME
    }
    payload = {mapping.get(k, k): v for k, v in updates.items()}
    try:
        res = requests.patch(
            supabase_table_url(),
            headers=supabase_headers("return=minimal"),
            params={COL_USERNAME: f"eq.{username}"},
            json=payload,
            timeout=12
        )
        return res.status_code < 400
    except Exception:
        return False

def load_users():
    if supabase_enabled():
        try:
            res = requests.get(supabase_table_url(), headers=supabase_headers(), params={"select": "*"}, timeout=12)
            if res.status_code < 400:
                users = {}
                for row in res.json():
                    user = normalize_supabase_user(row)
                    if user and user.get("username"):
                        users[user["username"]] = user
                return users
        except Exception:
            pass
    return load_json(USERS_FILE)

def save_users(users):
    save_json(USERS_FILE, users)

def find_user_by_token(users, token):
    for username, user in users.items():
        if user.get("token") == token:
            return username, user
    return None, None

def reset_daily_if_needed(user):
    if user.get("date") != today():
        user["date"] = today()
        user["used"] = 0

def is_paid_plan(plan: str) -> bool:
    return plan in ["basic", "pro", "business", "lifetime", "premium"]

def remaining_text(user):
    if is_paid_plan(user.get("plan", "free")):
        return "Không giới hạn"
    return max(0, FREE_LIMIT - int(user.get("used", 0)))

def smart_fallback_reply(message: str):
    text = message.lower()

    if any(k in text for k in ["content", "bài viết", "facebook", "caption", "bán hàng"]):
        return """📘 MẪU HỖ TRỢ CONTENT BÁN HÀNG

Để tôi viết nội dung chính xác hơn, anh/chị vui lòng gửi:

◆ Sản phẩm/dịch vụ đang bán
◆ Khách hàng mục tiêu
◆ Giá hoặc ưu đãi hiện tại
◆ Mục tiêu bài viết: tăng inbox, chốt đơn hay xây dựng thương hiệu

Gợi ý nhanh:
Anh/chị có thể nhập: "Viết bài Facebook bán [sản phẩm] cho [khách hàng mục tiêu] với ưu đãi [ưu đãi]."

📞 Cần hỗ trợ triển khai chuyên sâu: Zalo 036 338 2629
✉️ Email hỗ trợ: gptminipro@gmail.com"""

    if any(k in text for k in ["quảng cáo", "ads", "facebook ads", "tiktok", "chạy ads"]):
        return """📈 TƯ VẤN QUẢNG CÁO FACEBOOK / TIKTOK

Để tư vấn đúng chiến dịch, anh/chị vui lòng cung cấp:

◆ Ngành hàng/sản phẩm
◆ Ngân sách dự kiến mỗi ngày
◆ Khu vực muốn chạy quảng cáo
◆ Mục tiêu: inbox, đơn hàng, tăng follow hay nhận diện thương hiệu

Gợi ý:
Với shop mới, nên bắt đầu bằng chiến dịch tin nhắn/inbox ngân sách nhỏ để test tệp khách hàng trước.

📞 Cần hỗ trợ chạy quảng cáo trực tiếp: Zalo 036 338 2629
✉️ Email hỗ trợ: gptminipro@gmail.com"""

    if any(k in text for k in ["premium", "nâng cấp", "kích hoạt", "gói", "thanh toán"]):
        return """💎 HỖ TRỢ PREMIUM

GPT Mini Premium hiện có các gói:

◆ Cơ Bản: 99.000đ/tháng
◆ Chuyên Nghiệp: 199.000đ/tháng
◆ Doanh Nghiệp: 399.000đ/tháng
◆ Vĩnh Viễn: 599.000đ

Nếu anh/chị đã thanh toán, vui lòng gửi:
◆ Tên tài khoản đăng nhập
◆ Gói đã thanh toán
◆ Mã giao dịch hoặc nội dung chuyển khoản

📞 Zalo hỗ trợ kích hoạt nhanh: 036 338 2629
✉️ Email hỗ trợ: gptminipro@gmail.com"""

    if any(k in text for k in ["website", "landing", "trang bán hàng"]):
        return """🌐 TƯ VẤN WEBSITE / LANDING PAGE

Để xây dựng landing page bán hàng hiệu quả, anh/chị nên chuẩn bị:

◆ Tên sản phẩm/dịch vụ
◆ Ưu điểm nổi bật
◆ Hình ảnh/video sản phẩm
◆ Chính sách bảo hành/ưu đãi
◆ Số điện thoại/Zalo nhận khách

GPT Mini Premium có thể hỗ trợ lên bố cục, nội dung và kịch bản chuyển đổi.

📞 Hỗ trợ trực tiếp: Zalo 036 338 2629
✉️ Email hỗ trợ: gptminipro@gmail.com"""

    return """⚠️ Hệ thống đang xử lý nhiều yêu cầu.

Vui lòng thử lại sau ít phút.

Cần hỗ trợ nhanh:
📱 Zalo: 036 338 2629
📧 Email: gptminipro@gmail.com

Khi liên hệ, vui lòng gửi tên tài khoản, Gmail đăng ký và ảnh lỗi nếu có."""

def friendly_ai_error(message: str = ""):
    return smart_fallback_reply(message)

def expired_plan_message(user=None):
    expires_at = user.get("expires_at") if user else ""
    return f"""🔒 GÓI PREMIUM ĐÃ HẾT HẠN

Gói dịch vụ của tài khoản đã hết hạn{f" vào ngày {expires_at}" if expires_at else ""}.

━━━━━━━━━━━━━━━━━━━━━━

💎 VUI LÒNG GIA HẠN ĐỂ TIẾP TỤC SỬ DỤNG

◆ Tiếp tục truy cập AI không giới hạn
◆ Duy trì tốc độ phản hồi ưu tiên
◆ Mở lại toàn bộ tính năng Premium
◆ Được hỗ trợ kỹ thuật khi cần

━━━━━━━━━━━━━━━━━━━━━━

📞 Zalo hỗ trợ gia hạn: 036 338 2629
✉️ Email hỗ trợ: gptminipro@gmail.com

Sau khi thanh toán, vui lòng gửi tên tài khoản và ảnh giao dịch để được kích hoạt nhanh nhất."""

def free_limit_message():
    return """🔒 TÀI KHOẢN ĐÃ HẾT LƯỢT SỬ DỤNG

Bạn đã sử dụng toàn bộ lượt truy cập miễn phí được cấp cho tài khoản hiện tại.

━━━━━━━━━━━━━━━━━━━━━━

💎 NÂNG CẤP PREMIUM ĐỂ TIẾP TỤC SỬ DỤNG

◆ Truy cập AI ổn định và liên tục
◆ Hạn mức sử dụng cao hơn
◆ Tốc độ phản hồi nhanh hơn
◆ Ưu tiên tài nguyên hệ thống
◆ Hỗ trợ kỹ thuật ưu tiên
◆ Trải nghiệm đầy đủ các công cụ AI nâng cao

━━━━━━━━━━━━━━━━━━━━━━

📞 Zalo hỗ trợ kích hoạt nhanh: 036 338 2629
✉️ Email hỗ trợ: gptminipro@gmail.com"""

@app.get("/")
def home():
    if os.path.exists("index.html"):
        return FileResponse("index.html")
    return {"message": "GPT Mini Premium server is running."}

@app.get("/health")
def health():
    return {"success": True, "message": "GPT Mini Premium is live."}

@app.post("/register")
def register(req: RegisterRequest):
    username = safe_username(req.username)
    password = req.password.strip()
    email = req.email.strip().lower()
    phone = req.phone.strip()

    if len(username) < 3:
        return {"success": False, "message": "Tên đăng nhập phải có ít nhất 3 ký tự."}
    pass_msg = password_error_message(password)
    if pass_msg:
        return {"success": False, "message": pass_msg + " Ví dụ mật khẩu hợp lệ: Aa@078912"}
    if not email.endswith("@gmail.com") or "@" not in email:
        return {"success": False, "message": "Vui lòng nhập đúng địa chỉ Gmail để đăng ký."}

    clean_phone = phone.replace(" ", "").replace("-", "")
    if not (clean_phone.startswith("0") or clean_phone.startswith("+84")) or len(clean_phone) < 9:
        return {"success": False, "message": "Vui lòng nhập đúng số điện thoại/Zalo để đăng ký."}

    if supabase_enabled():
        if supabase_find_user_by_username(username):
            return {"success": False, "message": "Tài khoản đã tồn tại."}
        exists, msg = supabase_email_or_phone_exists(email, clean_phone)
        if exists:
            return {"success": False, "message": msg}

        token = str(uuid.uuid4())
        user = {
            "username": username,
            "email": email,
            "phone": phone,
            "password": hash_password(password),
            "token": token,
            "login_type": "password",
            "plan": "free",
            "used": 0,
            "date": today(),
            "registered_at": today(),
            "expires_at": None
        }
        ok, err = supabase_insert_user(user)
        if ok:
            return {"success": True, "message": "Đăng ký thành công. Thông tin đã được lưu vào Supabase.", "token": token, "username": username, "plan": "free", "plan_label": "FREE", "remaining": FREE_LIMIT}
        return {"success": False, "message": "Supabase chưa ghi được dữ liệu. Kiểm tra tên bảng/cột hoặc chính sách RLS."}

    users = load_json(USERS_FILE)
    if username in users:
        return {"success": False, "message": "Tài khoản đã tồn tại."}
    for _, existing in users.items():
        if existing.get("email", "").lower() == email:
            return {"success": False, "message": "Gmail này đã được dùng để đăng ký tài khoản khác."}
        if existing.get("phone", "").replace(" ", "").replace("-", "") == clean_phone:
            return {"success": False, "message": "Số điện thoại này đã được dùng để đăng ký tài khoản khác."}

    token = str(uuid.uuid4())
    users[username] = {
        "username": username, "email": email, "phone": phone,
        "password": hash_password(password), "token": token,
        "login_type": "password", "plan": "free", "used": 0,
        "date": today(), "registered_at": today(), "expires_at": None
    }
    save_json(USERS_FILE, users)
    return {"success": True, "message": "Đăng ký thành công. Thông tin Gmail và số điện thoại đã được ghi nhận để hỗ trợ tài khoản.", "token": token, "username": username, "plan": "free", "plan_label": "FREE", "remaining": FREE_LIMIT}

@app.post("/login")
def login(req: LoginRequest):
    username = safe_username(req.username)
    password = req.password.strip()

    if supabase_enabled():
        user = supabase_find_user_by_username(username)
        if not user:
            return {"success": False, "message": "Tài khoản không tồn tại."}
        if user.get("password") != hash_password(password):
            return {"success": False, "message": "Mật khẩu không đúng."}

        reset_daily_if_needed(user)
        if not user.get("token"):
            user["token"] = str(uuid.uuid4())
        supabase_update_user(username, {
            "token": user.get("token"),
            "date": user.get("date", today()),
            "used": user.get("used", 0),
            "plan": user.get("plan", "free")
        })

        plan = user.get("plan", "free")
        return {"success": True, "message": "Đăng nhập thành công.", "token": user["token"], "username": username, "plan": plan, "plan_label": PLAN_LABELS.get(plan, plan.upper()), "remaining": remaining_text(user)}

    users = load_json(USERS_FILE)
    if username not in users:
        return {"success": False, "message": "Tài khoản không tồn tại."}
    user = users[username]
    if user.get("password") != hash_password(password):
        return {"success": False, "message": "Mật khẩu không đúng."}

    reset_daily_if_needed(user)
    save_json(USERS_FILE, users)
    plan = user.get("plan", "free")
    return {"success": True, "message": "Đăng nhập thành công.", "token": user["token"], "username": username, "plan": plan, "plan_label": PLAN_LABELS.get(plan, plan.upper()), "remaining": remaining_text(user)}

@app.post("/google-login")
def google_login(req: GoogleLoginRequest):
    email = req.email.strip().lower()

    if not email or "@" not in email:
        return {"success": False, "message": "Email Google không hợp lệ."}

    username = safe_username(email.split("@")[0])

    if supabase_enabled():
        user = supabase_find_user_by_username(username)
        if not user:
            token = str(uuid.uuid4())
            user = {
                "username": username,
                "email": email,
                "phone": "",
                "password": "",
                "token": token,
                "login_type": "google-demo",
                "plan": "free",
                "used": 0,
                "date": today(),
                "registered_at": today(),
                "expires_at": None
            }
            ok, err = supabase_insert_user(user)
            if not ok:
                return {"success": False, "message": "Chưa lưu được tài khoản Google vào Supabase. Kiểm tra cột dữ liệu hoặc RLS."}
        reset_daily_if_needed(user)
        if not user.get("token"):
            user["token"] = str(uuid.uuid4())
        supabase_update_user(username, {"token": user["token"], "date": user.get("date", today()), "used": user.get("used", 0), "plan": user.get("plan", "free")})
        plan = user.get("plan", "free")
        return {"success": True, "message": "Đăng nhập Google demo thành công.", "token": user["token"], "username": username, "plan": plan, "plan_label": PLAN_LABELS.get(plan, plan.upper()), "remaining": remaining_text(user)}

    users = load_json(USERS_FILE)

    if username not in users:
        users[username] = {
            "username": username,
            "email": email,
            "name": req.name.strip(),
            "password": "",
            "token": str(uuid.uuid4()),
            "login_type": "google-demo",
            "plan": "free",
            "used": 0,
            "date": today()
        }

    user = users[username]
    reset_daily_if_needed(user)
    save_json(USERS_FILE, users)

    plan = user.get("plan", "free")
    return {"success": True, "message": "Đăng nhập Google demo thành công.", "token": user["token"], "username": username, "plan": plan, "plan_label": PLAN_LABELS.get(plan, plan.upper()), "remaining": remaining_text(user)}

@app.get("/status/{token}")
def status(token: str):
    users = load_users()
    username, user = find_user_by_token(users, token)
    if not user:
        return {"success": False, "message": "Chưa đăng nhập."}

    reset_daily_if_needed(user)
    if supabase_enabled() and username:
        supabase_update_user(username, {"date": user.get("date", today()), "used": user.get("used", 0)})
    save_users(users)
    plan = user.get("plan", "free")

    if is_expired(user):
        user["plan"] = "free"
        if supabase_enabled() and username:
            supabase_update_user(username, {"plan": "free"})
        save_users(users)
        return {
            "success": True,
            "username": username,
            "plan": "expired",
            "plan_label": "HẾT HẠN",
            "used": user.get("used", 0),
            "remaining": 0,
            "expired": True,
            "expires_at": user.get("expires_at"),
            "message": "Gói Premium đã hết hạn. Vui lòng gia hạn để tiếp tục sử dụng đầy đủ tính năng."
        }

    return {"success": True, "username": username, "plan": plan, "plan_label": PLAN_LABELS.get(plan, plan.upper()), "used": user.get("used", 0), "remaining": remaining_text(user), "expires_at": user.get("expires_at"), "days_left": days_left(user)}

@app.post("/activate")
def activate(req: ActivateRequest):
    users = load_users()
    username, user = find_user_by_token(users, req.token)
    if not user:
        return {"success": False, "message": "Bạn cần đăng nhập trước."}

    code = req.code.strip()
    if code not in PREMIUM_CODES:
        return {"success": False, "message": "Mã Premium không đúng. Vui lòng kiểm tra lại hoặc liên hệ Zalo 036 338 2629 / Email gptminipro@gmail.com."}

    plan = PREMIUM_CODES[code]
    user["plan"] = plan
    user["used"] = 0
    if plan == "lifetime":
        user["expires_at"] = None
    else:
        user["expires_at"] = add_days(30)

    if supabase_enabled() and username:
        supabase_update_user(username, {"plan": plan, "used": 0, "expires_at": user.get("expires_at")})
    save_users(users)

    expire_msg = "vĩnh viễn" if plan == "lifetime" else f"đến ngày {user.get('expires_at')}"
    return {"success": True, "message": f"Kích hoạt thành công gói {PLAN_LABELS.get(plan, plan)} {expire_msg}.", "plan": plan, "plan_label": PLAN_LABELS.get(plan, plan), "remaining": "Không giới hạn", "expires_at": user.get("expires_at"), "days_left": days_left(user)}

@app.post("/chat")
def chat(req: ChatRequest):
    users = load_users()
    username, user = find_user_by_token(users, req.token)

    if not user:
        return {"reply": "Bạn cần đăng nhập để sử dụng GPT Mini Premium.", "need_login": True}

    reset_daily_if_needed(user)
    plan = user.get("plan", "free")

    if is_expired(user):
        user["plan"] = "free"
        if supabase_enabled() and username:
            supabase_update_user(username, {"plan": "free"})
        save_users(users)
        return {
            "reply": expired_plan_message(user),
            "expired": True,
            "limit_reached": True,
            "plan": "expired",
            "remaining": 0
        }

    if not is_paid_plan(plan):
        if int(user.get("used", 0)) >= FREE_LIMIT:
            return {"reply": free_limit_message(), "limit_reached": True, "plan": plan, "remaining": 0}

        user["used"] = int(user.get("used", 0)) + 1
        if supabase_enabled() and username:
            supabase_update_user(username, {"used": user.get("used", 0), "date": user.get("date", today())})
        save_users(users)

    if not GEMINI_API_KEY:
        return {"reply": "🔔 Hệ thống AI chưa được cấu hình API Key. Vui lòng liên hệ quản trị viên.", "config_error": True}

    try:
        client = genai.Client(api_key=GEMINI_API_KEY)

        if username not in chat_memory:
            chat_memory[username] = []

        user_message = req.message.strip()
        chat_memory[username].append({"role": "user", "text": user_message})

        history_text = ""
        for msg in chat_memory[username][-10:]:
            if msg["role"] == "user":
                history_text += f"Người dùng: {msg['text']}\\n"
            else:
                history_text += f"AI: {msg['text']}\\n"

        prompt = f"""
Bạn là GPT Mini Premium.

Nhiệm vụ:
- Hỗ trợ Marketing, Content bán hàng, Facebook Ads, TikTok Ads.
- Hỗ trợ thiết kế hình ảnh, Website, Landing Page, AI Automation.
- Tư vấn kinh doanh online ngắn gọn, thực tế, dễ áp dụng.

QUY TẮC TRẢ LỜI:
1. Luôn trả lời bằng tiếng Việt.
2. Trả lời trực tiếp vào câu hỏi, không lan man.
3. Ưu tiên câu trả lời dưới 10 dòng nếu người dùng không yêu cầu chi tiết.
4. Không tự giới thiệu lại ở mỗi tin nhắn.
5. Không dùng các câu mở đầu dài dòng như:
   - Tuyệt vời!
   - Rất vui được hỗ trợ bạn.
   - Tôi là GPT Mini Premium.
   - Cảm ơn bạn đã liên hệ.
   - Thiết kế hình ảnh đóng vai trò rất quan trọng...
6. Chỉ hỏi thêm thông tin khi thật sự cần.
7. Nếu không đủ thông tin, trả lời:
   "Vui lòng cung cấp thêm thông tin để tôi hỗ trợ chính xác."

GỢI Ý TRẢ LỜI THEO TÌNH HUỐNG:
- Nếu hỏi thiết kế ảnh: hỏi loại ảnh, nội dung, kích thước, màu chủ đạo.
- Nếu hỏi quảng cáo: hỏi sản phẩm, ngân sách, khu vực, mục tiêu.
- Nếu hỏi content: hỏi sản phẩm, khách hàng mục tiêu, ưu đãi.
- Nếu hỏi website: ưu tiên tốc độ, bảo mật, SEO, trải nghiệm người dùng.
- Nếu hỏi Premium: tư vấn gói phù hợp ngắn gọn.

THÔNG TIN HỖ TRỢ:
- Zalo: 036 338 2629
- Email: gptminipro@gmail.com

LỊCH SỬ NGẮN:
{history_text}

CÂU HỎI MỚI NHẤT:
{user_message}

TRẢ LỜI NGẮN GỌN, CHUYÊN NGHIỆP:
"""

        response = client.models.generate_content(model="gemini-2.5-flash-lite", contents=prompt)
        reply = response.text.strip() if response and response.text else "AI không có phản hồi."

        chat_memory[username].append({"role": "ai", "text": reply})
        return {"reply": reply, "plan": user.get("plan", "free"), "remaining": remaining_text(user)}

    except Exception as e:
        # Ẩn mọi lỗi kỹ thuật: 429 quota, RESOURCE_EXHAUSTED, API error...
        return {
            "reply": """⚠️ Hệ thống đang xử lý nhiều yêu cầu.

Vui lòng thử lại sau ít phút.

📧 Email hỗ trợ:
gptminipro@gmail.com

📱 Zalo:
036 338 2629""",
            "ai_overload": True
        }

@app.post("/support")
def support(req: SupportRequest):
    users = load_users()
    username, user = find_user_by_token(users, req.token)
    if not user:
        return {"success": False, "message": "Bạn cần đăng nhập để gửi hỗ trợ."}

    messages = load_json(SUPPORT_FILE)
    if username not in messages:
        messages[username] = []

    if isinstance(messages.get(username), dict):
        ticket = messages[username]
    else:
        ticket = {
            "status": "open",
            "created_at": today(),
            "messages": messages.get(username, [])
        }

    ticket["status"] = "open"
    ticket["updated_at"] = today()
    ticket.setdefault("messages", [])
    ticket["messages"].append({
        "from": "customer",
        "sender": username,
        "message": req.message.strip(),
        "date": today()
    })

    messages[username] = ticket
    save_json(SUPPORT_FILE, messages)

    return {"success": True, "message": "Đã ghi nhận hỗ trợ. Bộ phận kỹ thuật sẽ kiểm tra trong thời gian sớm nhất."}


@app.get("/support-thread/{token}")
def support_thread(token: str):
    users = load_users()
    username, user = find_user_by_token(users, token)
    if not user:
        return {"success": False, "message": "Bạn cần đăng nhập để xem hỗ trợ."}

    support = load_json(SUPPORT_FILE)
    ticket = support.get(username)

    if not ticket:
        return {"success": True, "username": username, "status": "empty", "messages": []}

    if isinstance(ticket, dict):
        return {
            "success": True,
            "username": username,
            "status": ticket.get("status", "open"),
            "messages": ticket.get("messages", [])
        }

    return {"success": True, "username": username, "status": "open", "messages": ticket}

@app.get("/support-messages")
def support_messages():
    return load_json(SUPPORT_FILE)


@app.get("/supabase-test")
def supabase_test():
    if not supabase_enabled():
        return {"success": False, "message": "Chưa có SUPABASE_URL hoặc SUPABASE_KEY trong Render Environment."}
    try:
        res = requests.get(
            supabase_table_url(),
            headers=supabase_headers(),
            params={"select": "*", "limit": "1"},
            timeout=12
        )
        return {
            "success": res.status_code < 400,
            "status_code": res.status_code,
            "table": SUPABASE_TABLE,
            "url": supabase_table_url(),
            "response": res.text[:500]
        }
    except Exception as e:
        return {"success": False, "message": str(e)}

@app.get("/admin", response_class=HTMLResponse)
def admin_page():
    return """
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<title>GPT Mini Premium Admin</title>
<style>
*{box-sizing:border-box}
body{font-family:Arial;background:#020617;color:white;margin:0;padding:26px}
h1{color:#38bdf8;margin-top:0}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.card{background:#0f172a;border:1px solid #334155;border-radius:18px;padding:18px;margin:14px 0}
button{padding:12px 18px;border:none;border-radius:12px;background:#2563eb;color:white;font-weight:bold;cursor:pointer;margin:5px}
.btn-green{background:#22c55e}.btn-red{background:#ef4444}
input,select,textarea{width:100%;padding:13px;border-radius:12px;border:1px solid #334155;background:#020617;color:white;margin:6px 0}
textarea{min-height:86px;resize:vertical}
.badge{padding:5px 10px;border-radius:999px;background:#2563eb;font-size:13px}
.badge-open{background:#f97316}.badge-closed{background:#334155}.badge-pro{background:#16a34a}.badge-free{background:#64748b}
.small{color:#94a3b8;font-size:13px;line-height:1.6}
.msg{padding:12px;border-radius:14px;margin:8px 0;line-height:1.5}
.customer{background:#1e293b;border-left:4px solid #38bdf8}
.admin{background:#052e16;border-left:4px solid #22c55e}
.quick{display:grid;grid-template-columns:repeat(2,1fr);gap:8px}
.topbar{display:flex;gap:10px;align-items:center;margin-bottom:16px}.topbar input{max-width:260px}
@media(max-width:900px){.grid{grid-template-columns:1fr}.quick{grid-template-columns:1fr}}
</style>
</head>
<body>
<h1>⚡ GPT Mini Premium Admin</h1>
<div class="topbar">
<input id="adminPassword" type="password" placeholder="Mật khẩu admin" value="admin123">
<button onclick="loadData()">Tải dữ liệu</button>
</div>
<div class="grid"><div><h2>👥 Danh sách tài khoản</h2><div id="users"></div></div><div><h2>💳 Trung tâm hỗ trợ / Thanh toán</h2><div id="payments"></div></div></div>
<script>
let GLOBAL_DATA = {};
function esc(str){return String(str||"").replace(/[&<>"']/g,s=>({"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;","'":"&#039;"}[s]));}
function ticketMessages(ticket){if(!ticket)return[];if(Array.isArray(ticket))return ticket;if(Array.isArray(ticket.messages))return ticket.messages;return[];}
async function loadData(){
 const res=await fetch("/admin-data"); const data=await res.json(); GLOBAL_DATA=data;
 const usersBox=document.getElementById("users"); usersBox.innerHTML="";
 Object.keys(data.users||{}).forEach(username=>{
  const u=data.users[username]||{}; const plan=u.plan||"free"; const badge=plan==="free"?"badge-free":"badge-pro";
  usersBox.innerHTML+=`<div class="card"><b>👤 ${esc(username)}</b><br><br>
  Gói hiện tại: <span class="badge ${badge}">${esc(plan)}</span><br>
  Hết hạn: <b>${esc(u.expires_at || (plan==="lifetime"?"Vĩnh viễn":"Chưa có"))}</b><br>
  Gmail: ${esc(u.email||"Chưa có")}<br>SĐT/Zalo: ${esc(u.phone||"Chưa có")}<br>
  Ngày đăng ký: ${esc(u.registered_at||"Chưa có")}<br>Đã dùng: ${esc(u.used||0)} lượt<br><br>
  <select id="plan_${esc(username)}"><option value="free">Free</option><option value="basic">Cơ Bản 99k - 30 ngày</option><option value="pro">Pro 199k - 30 ngày</option><option value="business">Doanh Nghiệp 399k - 30 ngày</option><option value="lifetime">Vĩnh Viễn</option></select>
  <button class="btn-green" onclick="approve('${esc(username)}')">✅ Cập nhật gói</button></div>`;
 });
 renderTickets(data);
}
function renderTickets(data){
 const box=document.getElementById("payments"); box.innerHTML="";
 const support=data.support||{}; const names=Object.keys(support);
 if(!names.length){box.innerHTML='<div class="card small">Chưa có yêu cầu hỗ trợ nào.</div>';return;}
 names.reverse().forEach(username=>{
  const ticket=support[username]; const status=ticket.status||"open"; const msgs=ticketMessages(ticket); const user=(GLOBAL_DATA.users||{})[username]||{};
  let msgHtml=msgs.map(m=>{const role=(m.from==="admin")?"admin":"customer"; const who=(m.from==="admin")?"Admin":username; return `<div class="msg ${role}"><b>${esc(who)}</b> <span class="small">${esc(m.date||"")}</span><br>${esc(m.message)}</div>`;}).join("");
  box.innerHTML+=`<div class="card"><b>👤 ${esc(username)}</b> <span class="badge ${status==="closed"?"badge-closed":"badge-open"}">${status==="closed"?"Đã xử lý":"Đang mở"}</span>
  <div class="small">Gmail: ${esc(user.email||"Chưa có")}<br>SĐT: ${esc(user.phone||"Chưa có")}<br>Gói: ${esc(user.plan||"free")} | Hết hạn: ${esc(user.expires_at||"Chưa có")}</div>
  <h3>💬 Hội thoại</h3>${msgHtml||'<div class="small">Chưa có nội dung.</div>'}
  <textarea id="reply_${esc(username)}" placeholder="Nhập nội dung trả lời khách...">Xin chào anh/chị, em đã nhận được yêu cầu hỗ trợ. Em sẽ kiểm tra và phản hồi trong thời gian sớm nhất ạ.</textarea>
  <div class="quick"><button onclick="fillReply('${esc(username)}','payment')">Mẫu đã nhận thanh toán</button><button onclick="fillReply('${esc(username)}','need_image')">Mẫu cần ảnh giao dịch</button><button onclick="fillReply('${esc(username)}','activated')">Mẫu đã kích hoạt</button><button onclick="fillReply('${esc(username)}','support')">Mẫu hỗ trợ kỹ thuật</button></div>
  <button onclick="replyCustomer('${esc(username)}')">📨 Gửi trả lời</button><button class="btn-red" onclick="closeTicket('${esc(username)}')">🗑 Đóng ticket</button>
  <hr style="border-color:#334155;margin:16px 0">
  <div class="quick"><button onclick="quickApprove('${esc(username)}','basic')">Cơ Bản</button><button onclick="quickApprove('${esc(username)}','pro')">Pro</button><button onclick="quickApprove('${esc(username)}','business')">Doanh Nghiệp</button><button class="btn-green" onclick="quickApprove('${esc(username)}','lifetime')">∞ Vĩnh Viễn</button></div></div>`;
 });
}
function fillReply(username,type){const box=document.getElementById("reply_"+username); const t={payment:"Em đã nhận được thông tin thanh toán của anh/chị. Bộ phận kỹ thuật đang kiểm tra giao dịch và sẽ kích hoạt gói trong thời gian sớm nhất.",need_image:"Anh/chị vui lòng gửi thêm ảnh chụp màn hình giao dịch chuyển khoản hoặc nội dung chuyển khoản để bên em kiểm tra và kích hoạt nhanh hơn ạ.",activated:"Gói Premium của anh/chị đã được kích hoạt thành công. Anh/chị vui lòng đăng xuất rồi đăng nhập lại để hệ thống cập nhật gói mới.",support:"Em đã ghi nhận lỗi anh/chị đang gặp. Anh/chị vui lòng gửi thêm ảnh màn hình lỗi, tên tài khoản và thao tác đang thực hiện để bên em kiểm tra chính xác."}; box.value=t[type]||"";}
async function approve(username){const password=document.getElementById("adminPassword").value; const plan=document.getElementById("plan_"+username).value; const res=await fetch("/admin-approve",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({username,admin_password:password,plan})}); const data=await res.json(); alert(data.message); loadData();}
async function quickApprove(username,plan){const password=document.getElementById("adminPassword").value; if(!confirm("Xác nhận nâng cấp "+username+" lên gói "+plan+"?"))return; const res=await fetch("/admin-approve",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({username,admin_password:password,plan})}); const data=await res.json(); alert(data.message); loadData();}
async function replyCustomer(username){const password=document.getElementById("adminPassword").value; const message=document.getElementById("reply_"+username).value; const res=await fetch("/admin-reply",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({username,admin_password:password,message})}); const data=await res.json(); alert(data.message); loadData();}
async function closeTicket(username){const password=document.getElementById("adminPassword").value; const res=await fetch("/admin-close-ticket",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({username,admin_password:password})}); const data=await res.json(); alert(data.message); loadData();}
</script></body></html>
"""


@app.post("/admin-reply")
def admin_reply(req: AdminReplyRequest):
    if req.admin_password != ADMIN_PASSWORD:
        return {"success": False, "message": "Sai mật khẩu admin."}

    username = safe_username(req.username)
    message = req.message.strip()
    if not message:
        return {"success": False, "message": "Nội dung trả lời không được để trống."}

    support = load_json(SUPPORT_FILE)
    if isinstance(support.get(username), dict):
        ticket = support[username]
    else:
        ticket = {"status": "open", "created_at": today(), "messages": support.get(username, [])}

    ticket.setdefault("messages", [])
    ticket["status"] = "open"
    ticket["updated_at"] = today()
    ticket["messages"].append({"from": "admin", "sender": "admin", "message": message, "date": today()})
    support[username] = ticket
    save_json(SUPPORT_FILE, support)
    return {"success": True, "message": "Đã gửi trả lời cho khách trong trung tâm hỗ trợ."}

@app.post("/admin-close-ticket")
def admin_close_ticket(req: AdminCloseTicketRequest):
    if req.admin_password != ADMIN_PASSWORD:
        return {"success": False, "message": "Sai mật khẩu admin."}

    username = safe_username(req.username)
    support = load_json(SUPPORT_FILE)
    if username not in support:
        return {"success": False, "message": "Không tìm thấy ticket."}

    if isinstance(support[username], dict):
        support[username]["status"] = "closed"
        support[username]["closed_at"] = today()
    else:
        support[username] = {"status": "closed", "created_at": today(), "closed_at": today(), "messages": support[username]}

    save_json(SUPPORT_FILE, support)
    return {"success": True, "message": "Đã đóng ticket hỗ trợ."}

@app.get("/admin-data")
def admin_data():
    return {"users": load_users(), "support": load_json(SUPPORT_FILE)}

@app.post("/admin-approve")
def admin_approve(req: AdminApproveRequest):
    if req.admin_password != ADMIN_PASSWORD:
        return {"success": False, "message": "Sai mật khẩu admin."}

    users = load_users()
    username = safe_username(req.username)
    if username not in users:
        return {"success": False, "message": "Không tìm thấy tài khoản."}

    users[username]["plan"] = req.plan
    users[username]["used"] = 0
    if req.plan == "lifetime":
        users[username]["expires_at"] = None
    elif req.plan == "free":
        users[username]["expires_at"] = None
    else:
        users[username]["expires_at"] = add_days(30)

    if supabase_enabled():
        supabase_update_user(username, {"plan": req.plan, "used": 0, "expires_at": users[username].get("expires_at")})
    save_users(users)

    expire_msg = "vĩnh viễn" if req.plan == "lifetime" else (f"đến ngày {users[username].get('expires_at')}" if req.plan != "free" else "")
    return {"success": True, "message": f"Đã cập nhật tài khoản {username} thành gói {PLAN_LABELS.get(req.plan, req.plan)} {expire_msg}."}

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8005))
    uvicorn.run(app, host="0.0.0.0", port=port)
