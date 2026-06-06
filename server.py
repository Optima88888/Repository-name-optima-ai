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

class AdminApproveRequest(BaseModel):
    username: str
    admin_password: str
    plan: str = "business"

class AdminExtendRequest(BaseModel):
    username: str
    admin_password: str
    days: int = 30

class AdminExpireRequest(BaseModel):
    username: str
    admin_password: str

class AdminDeleteRequest(BaseModel):
    username: str
    admin_password: str

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

def supabase_delete_user(username: str):
    if not supabase_enabled():
        return False
    try:
        res = requests.delete(
            supabase_table_url(),
            headers=supabase_headers("return=minimal"),
            params={COL_USERNAME: f"eq.{username}"},
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

    messages[username].append({"from": username, "message": req.message.strip(), "date": today()})
    save_json(SUPPORT_FILE, messages)

    return {"success": True, "message": "Đã ghi nhận hỗ trợ. Nếu cần gấp vui lòng liên hệ Zalo 036 338 2629 hoặc Email gptminipro@gmail.com."}

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
body{font-family:Arial,system-ui;background:#020617;color:white;margin:0;padding:28px}
h1{color:#38bdf8;margin:0 0 8px}
p{color:#cbd5e1}
.header{display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap;margin-bottom:22px}
.panel{background:#0f172a;border:1px solid #334155;border-radius:18px;padding:18px;margin:15px 0}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(270px,1fr));gap:14px}
.card{background:#111827;border:1px solid #334155;border-radius:16px;padding:18px}
input,select{padding:12px;border-radius:10px;border:1px solid #334155;background:#020617;color:white;margin:4px;width:100%}
button{padding:11px 14px;border:none;border-radius:10px;background:#2563eb;color:white;font-weight:bold;cursor:pointer;margin:4px}
button.green{background:#22c55e}
button.orange{background:#f59e0b}
button.red{background:#ef4444}
button.gray{background:#475569}
.badge{display:inline-block;padding:5px 10px;border-radius:999px;background:#2563eb;font-weight:bold}
.badge.free{background:#64748b}
.badge.pro,.badge.basic,.badge.business,.badge.premium{background:#16a34a}
.badge.lifetime{background:#a855f7}
.row{display:flex;gap:8px;flex-wrap:wrap}
.row button{flex:1}
.small{font-size:13px;color:#94a3b8;line-height:1.6}
.search{max-width:420px}
.notice{background:#172554;border:1px solid #2563eb;color:#dbeafe;border-radius:14px;padding:14px;margin-bottom:16px}
</style>
</head>
<body>
<div class="header">
    <div>
        <h1>⚡ GPT Mini Premium Admin</h1>
        <p>Quản lý khách hàng, nâng cấp gói, gia hạn 30 ngày và khóa tài khoản.</p>
    </div>
    <button class="green" onclick="loadData()">🔄 Tải dữ liệu</button>
</div>

<div class="notice">
    <b>Hướng dẫn nhanh:</b> Nhập mật khẩu admin → Tải dữ liệu → Chọn gói hoặc bấm Gia hạn 30 ngày cho khách.
</div>

<div class="panel">
    <label>Mật khẩu admin</label>
    <input id="adminPassword" type="password" placeholder="Nhập mật khẩu admin" value="admin123">
    <label>Tìm khách hàng</label>
    <input class="search" id="searchBox" placeholder="Tìm theo username, email, số điện thoại..." oninput="renderUsers()">
</div>

<h2>👥 Danh sách tài khoản</h2>
<div id="users" class="grid"></div>

<h2>💳 Khách báo thanh toán / hỗ trợ</h2>
<div id="payments" class="grid"></div>

<script>
let allData = {users:{}, support:{}};

function planBadge(plan){
    plan = plan || "free";
    return `<span class="badge ${plan}">${plan.toUpperCase()}</span>`;
}

function expiredText(u){
    if(!u.expires_at) return (u.plan === "lifetime") ? "Vĩnh viễn" : "Chưa có";
    const today = new Date().toISOString().slice(0,10);
    if(u.expires_at < today) return `<b style="color:#ef4444">${u.expires_at} - ĐÃ HẾT HẠN</b>`;
    return `<b style="color:#22c55e">${u.expires_at}</b>`;
}

async function loadData(){
    const res = await fetch("/admin-data");
    allData = await res.json();
    renderUsers();
    renderSupport();
}

function renderUsers(){
    const usersBox = document.getElementById("users");
    const q = (document.getElementById("searchBox").value || "").toLowerCase();
    usersBox.innerHTML = "";

    Object.keys(allData.users || {}).forEach(username=>{
        const u = allData.users[username];
        const text = `${username} ${u.email||""} ${u.phone||""}`.toLowerCase();
        if(q && !text.includes(q)) return;

        usersBox.innerHTML += `
        <div class="card">
            <h3>👤 ${username}</h3>
            <div>Gói: ${planBadge(u.plan || "free")}</div><br>
            <div class="small">
                Gmail: ${u.email || "Chưa có"}<br>
                SĐT/Zalo: ${u.phone || "Chưa có"}<br>
                Ngày đăng ký: ${u.registered_at || u.created_at || "Chưa có"}<br>
                Hết hạn: ${expiredText(u)}<br>
                Đã dùng: ${u.used || 0} lượt
            </div><br>

            <select id="plan_${username}">
                <option value="free">Free</option>
                <option value="basic">Cơ Bản 99k - 30 ngày</option>
                <option value="pro">Chuyên Nghiệp 199k - 30 ngày</option>
                <option value="business">Doanh Nghiệp 399k - 30 ngày</option>
                <option value="lifetime">Vĩnh Viễn 599k</option>
            </select>

            <div class="row">
                <button class="green" onclick="approve('${username}')">💎 Cập nhật gói</button>
                <button class="orange" onclick="extendUser('${username}',30)">📅 Gia hạn 30 ngày</button>
            </div>
            <div class="row">
                <button class="green" onclick="quickApprove('${username}','lifetime')">♾️ Gói vĩnh viễn</button>
                <button class="gray" onclick="expireUser('${username}')">🚫 Cho hết hạn</button>
                <button class="red" onclick="deleteUser('${username}')">🗑️ Xóa</button>
            </div>
        </div>`;
    });
}

function renderSupport(){
    const paymentsBox = document.getElementById("payments");
    paymentsBox.innerHTML = "";
    Object.keys(allData.support || {}).forEach(username=>{
        allData.support[username].slice().reverse().forEach(msg=>{
            paymentsBox.innerHTML += `
            <div class="card">
                <h3>👤 ${username}</h3>
                <div class="small">${msg.message}</div><br>
                <div class="row">
                    <button onclick="quickApprove('${username}','basic')">Cơ Bản</button>
                    <button onclick="quickApprove('${username}','pro')">Pro</button>
                </div>
                <div class="row">
                    <button onclick="quickApprove('${username}','business')">Doanh Nghiệp</button>
                    <button class="green" onclick="quickApprove('${username}','lifetime')">♾️ Vĩnh Viễn</button>
                </div>
            </div>`;
        });
    });
}

async function approve(username){
    const password = document.getElementById("adminPassword").value;
    const plan = document.getElementById("plan_" + username).value;
    const res = await fetch("/admin-approve", {
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({username, admin_password:password, plan})
    });
    const data = await res.json();
    alert(data.message);
    loadData();
}

async function quickApprove(username, plan){
    const password = document.getElementById("adminPassword").value;
    const res = await fetch("/admin-approve", {
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({username, admin_password:password, plan})
    });
    const data = await res.json();
    alert(data.message);
    loadData();
}

async function extendUser(username, days){
    const password = document.getElementById("adminPassword").value;
    const res = await fetch("/admin-extend", {
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({username, admin_password:password, days})
    });
    const data = await res.json();
    alert(data.message);
    loadData();
}

async function expireUser(username){
    if(!confirm("Cho tài khoản này hết hạn ngay?")) return;
    const password = document.getElementById("adminPassword").value;
    const res = await fetch("/admin-expire", {
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({username, admin_password:password})
    });
    const data = await res.json();
    alert(data.message);
    loadData();
}

async function deleteUser(username){
    if(!confirm("Xóa tài khoản này khỏi hệ thống?")) return;
    const password = document.getElementById("adminPassword").value;
    const res = await fetch("/admin-delete", {
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({username, admin_password:password})
    });
    const data = await res.json();
    alert(data.message);
    loadData();
}

loadData();
</script>
</body>
</html>
"""

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

@app.post("/admin-extend")
def admin_extend(req: AdminExtendRequest):
    if req.admin_password != ADMIN_PASSWORD:
        return {"success": False, "message": "Sai mật khẩu admin."}

    users = load_users()
    username = safe_username(req.username)
    if username not in users:
        return {"success": False, "message": "Không tìm thấy tài khoản."}

    user = users[username]
    current_exp = parse_date(user.get("expires_at"))
    base_date = current_exp if current_exp and current_exp >= date.today() else date.today()
    new_exp = str(base_date + timedelta(days=max(1, int(req.days))))

    if user.get("plan", "free") == "free":
        user["plan"] = "pro"

    if user.get("plan") == "lifetime":
        return {"success": True, "message": f"Tài khoản {username} là gói vĩnh viễn, không cần gia hạn."}

    user["expires_at"] = new_exp
    user["used"] = 0

    if supabase_enabled():
        supabase_update_user(username, {"plan": user.get("plan", "pro"), "used": 0, "expires_at": new_exp})
    save_users(users)

    return {"success": True, "message": f"Đã gia hạn {req.days} ngày cho {username}. Hạn mới: {new_exp}."}

@app.post("/admin-expire")
def admin_expire(req: AdminExpireRequest):
    if req.admin_password != ADMIN_PASSWORD:
        return {"success": False, "message": "Sai mật khẩu admin."}

    users = load_users()
    username = safe_username(req.username)
    if username not in users:
        return {"success": False, "message": "Không tìm thấy tài khoản."}

    users[username]["plan"] = "free"
    users[username]["expires_at"] = str(date.today() - timedelta(days=1))
    users[username]["used"] = FREE_LIMIT

    if supabase_enabled():
        supabase_update_user(username, {"plan": "free", "used": FREE_LIMIT, "expires_at": users[username]["expires_at"]})
    save_users(users)

    return {"success": True, "message": f"Đã cho tài khoản {username} hết hạn."}

@app.post("/admin-delete")
def admin_delete(req: AdminDeleteRequest):
    if req.admin_password != ADMIN_PASSWORD:
        return {"success": False, "message": "Sai mật khẩu admin."}

    users = load_users()
    username = safe_username(req.username)
    if username not in users:
        return {"success": False, "message": "Không tìm thấy tài khoản."}

    if supabase_enabled():
        supabase_delete_user(username)

    local_users = load_json(USERS_FILE)
    if username in local_users:
        del local_users[username]
        save_json(USERS_FILE, local_users)

    return {"success": True, "message": f"Đã xóa tài khoản {username}."}

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8005))
    uvicorn.run(app, host="0.0.0.0", port=port)
