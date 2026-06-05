from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel
from google import genai
from datetime import date
import uvicorn
import json
import os
import hashlib
import uuid

# ======================================================
# OPTIMA AI PRO FINAL
# Thay API key Gemini thật của bạn vào dòng dưới đây
# ======================================================
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
client = genai.Client(api_key=GEMINI_API_KEY)

app = FastAPI(title="OPTIMA AI PRO")

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
ADMIN_PASSWORD = "admin123"

# Các mã test để mở khóa gói khi cần
PREMIUM_CODES = {
    "OPTIMA-BASIC-99000": "basic",
    "OPTIMA-PRO-199000": "pro",
    "OPTIMA-BUSINESS-399000": "business",
    "OPTIMA-LIFETIME-599000": "lifetime",
    "OPTIMA-PREMIUM-99000": "basic",
    "OPTIMA-VIP-2026": "business"
}

PLAN_LABELS = {
    "free": "FREE",
    "basic": "CƠ BẢN",
    "pro": "CHUYÊN NGHIỆP",
    "business": "DOANH NGHIỆP",
    "lifetime": "VĨNH VIỄN"
}

chat_memory = {}

class RegisterRequest(BaseModel):
    username: str
    password: str

class LoginRequest(BaseModel):
    username: str
    password: str

class GoogleLoginRequest(BaseModel):
    name: str
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

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()

def today() -> str:
    return str(date.today())

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

def load_users():
    return load_json(USERS_FILE)

def save_users(users):
    save_json(USERS_FILE, users)

def safe_username(text: str) -> str:
    return text.strip().lower().replace(" ", "_")

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

@app.get("/")
def home():
    return FileResponse("index.html")

@app.post("/register")
def register(req: RegisterRequest):
    users = load_users()
    username = safe_username(req.username)
    password = req.password.strip()

    if len(username) < 3:
        return {"success": False, "message": "Tên đăng nhập phải có ít nhất 3 ký tự."}

    if len(password) < 4:
        return {"success": False, "message": "Mật khẩu phải có ít nhất 4 ký tự."}

    if username in users:
        return {"success": False, "message": "Tài khoản đã tồn tại."}

    token = str(uuid.uuid4())
    users[username] = {
        "username": username,
        "email": "",
        "password": hash_password(password),
        "token": token,
        "login_type": "password",
        "plan": "free",
        "used": 0,
        "date": today()
    }
    save_users(users)

    return {
        "success": True,
        "message": "Đăng ký thành công.",
        "token": token,
        "username": username,
        "plan": "free",
        "plan_label": PLAN_LABELS["free"],
        "remaining": FREE_LIMIT
    }

@app.post("/login")
def login(req: LoginRequest):
    users = load_users()
    username = safe_username(req.username)
    password = req.password.strip()

    if username not in users:
        return {"success": False, "message": "Tài khoản không tồn tại."}

    user = users[username]

    if user.get("password") != hash_password(password):
        return {"success": False, "message": "Mật khẩu không đúng."}

    reset_daily_if_needed(user)
    save_users(users)

    plan = user.get("plan", "free")
    return {
        "success": True,
        "message": "Đăng nhập thành công.",
        "token": user["token"],
        "username": username,
        "plan": plan,
        "plan_label": PLAN_LABELS.get(plan, plan.upper()),
        "remaining": remaining_text(user)
    }

@app.post("/google-login")
def google_login(req: GoogleLoginRequest):
    users = load_users()
    email = req.email.strip().lower()
    name = req.name.strip()

    if not email or "@" not in email:
        return {"success": False, "message": "Email Google không hợp lệ."}

    username = safe_username(email.split("@")[0])

    if username not in users:
        token = str(uuid.uuid4())
        users[username] = {
            "username": username,
            "email": email,
            "name": name,
            "password": "",
            "token": token,
            "login_type": "google-demo",
            "plan": "free",
            "used": 0,
            "date": today()
        }

    user = users[username]
    reset_daily_if_needed(user)
    save_users(users)

    plan = user.get("plan", "free")
    return {
        "success": True,
        "message": "Đăng nhập Google demo thành công.",
        "token": user["token"],
        "username": username,
        "plan": plan,
        "plan_label": PLAN_LABELS.get(plan, plan.upper()),
        "remaining": remaining_text(user)
    }

@app.get("/status/{token}")
def status(token: str):
    users = load_users()
    username, user = find_user_by_token(users, token)

    if not user:
        return {"success": False, "message": "Chưa đăng nhập."}

    reset_daily_if_needed(user)
    save_users(users)

    plan = user.get("plan", "free")
    return {
        "success": True,
        "username": username,
        "plan": plan,
        "plan_label": PLAN_LABELS.get(plan, plan.upper()),
        "used": user.get("used", 0),
        "remaining": remaining_text(user)
    }

@app.post("/activate")
def activate(req: ActivateRequest):
    users = load_users()
    username, user = find_user_by_token(users, req.token)

    if not user:
        return {"success": False, "message": "Bạn cần đăng nhập trước."}

    code = req.code.strip()

    if code in PREMIUM_CODES:
        plan = PREMIUM_CODES[code]
        user["plan"] = plan
        user["used"] = 0
        save_users(users)

        return {
            "success": True,
            "message": f"Kích hoạt thành công gói {PLAN_LABELS.get(plan, plan)}.",
            "plan": plan,
            "plan_label": PLAN_LABELS.get(plan, plan),
            "remaining": "Không giới hạn"
        }

    return {
        "success": False,
        "message": "Mã Premium không đúng. Vui lòng kiểm tra lại hoặc liên hệ admin."
    }

@app.post("/chat")
def chat(req: ChatRequest):
    try:
        users = load_users()
        username, user = find_user_by_token(users, req.token)

        if not user:
            return {"reply": "Bạn cần đăng nhập để sử dụng OPTIMA AI.", "need_login": True}

        reset_daily_if_needed(user)
        plan = user.get("plan", "free")

        if not is_paid_plan(plan):
            if int(user.get("used", 0)) >= FREE_LIMIT:
                return {
                    "reply": "LIMIT_REACHED",
                    "limit_reached": True,
                    "plan": plan,
                    "remaining": 0
                }

            user["used"] = int(user.get("used", 0)) + 1
            save_users(users)

        if username not in chat_memory:
            chat_memory[username] = []

        user_message = req.message.strip()
        chat_memory[username].append({"role": "user", "text": user_message})

        history_text = ""
        for msg in chat_memory[username][-40:]:
            if msg["role"] == "user":
                history_text += f"Người dùng: {msg['text']}\n"
            else:
                history_text += f"OPTIMA AI: {msg['text']}\n"

        prompt = f"""
Bạn là OPTIMA AI, trợ lý trí tuệ nhân tạo cao cấp của OPTIMA.

QUY TẮC:
- Luôn trả lời bằng tiếng Việt.
- Trả lời chuyên nghiệp, rõ ràng, dễ hiểu.
- Ghi nhớ thông tin người dùng trong cuộc trò chuyện hiện tại.
- Hỗ trợ mạnh về marketing, quảng cáo, bán hàng, content, AI automation, website.
- Nếu người dùng hỏi tên bạn, hãy nói bạn là OPTIMA AI.
- Nếu người dùng hỏi bảng giá, hãy giới thiệu 4 gói:
  1) Cơ Bản 99.000đ/tháng
  2) Chuyên Nghiệp 199.000đ/tháng
  3) Doanh Nghiệp 399.000đ/tháng
  4) Vĩnh Viễn 599.000đ thanh toán một lần.

LỊCH SỬ HỘI THOẠI:
{history_text}

HÃY TRẢ LỜI:
"""

        response = client.models.generate_content(
            model="gemini-2.5-flash-lite",
            contents=prompt
        )

        reply = response.text.strip()
        chat_memory[username].append({"role": "ai", "text": reply})

        return {
            "reply": reply,
            "plan": user.get("plan", "free"),
            "remaining": remaining_text(user)
        }

    except Exception as e:
        return {"reply": "Lỗi AI: " + str(e)}

@app.post("/support")
def support(req: SupportRequest):
    users = load_users()
    username, user = find_user_by_token(users, req.token)

    if not user:
        return {"success": False, "message": "Bạn cần đăng nhập để gửi hỗ trợ."}

    messages = load_json(SUPPORT_FILE)

    if username not in messages:
        messages[username] = []

    messages[username].append({
        "from": username,
        "message": req.message.strip(),
        "date": today()
    })

    save_json(SUPPORT_FILE, messages)
    return {"success": True, "message": "Đã gửi tin nhắn hỗ trợ. Nhân viên sẽ phản hồi sớm."}

@app.get("/support-messages")
def support_messages():
    return load_json(SUPPORT_FILE)

@app.get("/admin", response_class=HTMLResponse)
def admin_page():
    return """
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<title>OPTIMA AI ADMIN</title>
<style>
body{font-family:Arial;background:#020617;color:white;padding:30px}
h1{color:#38bdf8}
.card{background:#0f172a;border:1px solid #334155;border-radius:16px;padding:20px;margin:15px 0}
button{padding:12px 18px;border:none;border-radius:10px;background:#22c55e;color:white;font-weight:bold;cursor:pointer}
input,select{padding:12px;border-radius:10px;border:1px solid #334155;background:#020617;color:white;margin:5px}
.badge{padding:5px 10px;border-radius:999px;background:#2563eb}
.free{background:#475569}.basic{background:#2563eb}.pro{background:#7c3aed}.business{background:#f59e0b}.lifetime{background:#ef4444}
</style>
</head>
<body>
<h1>⚡ OPTIMA AI ADMIN</h1>
<p>Mật khẩu admin:</p>
<input id="adminPassword" type="password" placeholder="Mật khẩu admin" value="admin123">
<button onclick="loadData()">Tải dữ liệu</button>
<h2>👥 Danh sách tài khoản</h2>
<div id="users"></div>
<h2>💳 Khách báo thanh toán / hỗ trợ</h2>
<div id="payments"></div>
<script>
async function loadData(){
    const res = await fetch("/admin-data");
    const data = await res.json();
    const usersBox = document.getElementById("users");
    usersBox.innerHTML = "";

    Object.keys(data.users).forEach(username=>{
        const u = data.users[username];
        const plan = u.plan || "free";
        usersBox.innerHTML += `
        <div class="card">
            <b>👤 ${username}</b><br><br>
            Gói hiện tại: <span class="badge ${plan}">${plan}</span><br><br>
            Đã dùng: ${u.used || 0} lượt<br><br>
            <select id="plan_${username}">
                <option value="free">Free</option>
                <option value="basic">Cơ Bản 99k</option>
                <option value="pro">Chuyên Nghiệp 199k</option>
                <option value="business">Doanh Nghiệp 399k</option>
                <option value="lifetime">Vĩnh Viễn 599k</option>
            </select>
            <button onclick="approve('${username}')">✅ Cập nhật gói</button>
        </div>`;
    });

    const paymentsBox = document.getElementById("payments");
    paymentsBox.innerHTML = "";

    Object.keys(data.support).forEach(username=>{
        data.support[username].slice().reverse().forEach(msg=>{
            paymentsBox.innerHTML += `
            <div class="card">
                <b>👤 ${username}</b><br><br>
                ${msg.message}<br><br>
                <button onclick="quickApprove('${username}','basic')">Duyệt Cơ Bản</button>
                <button onclick="quickApprove('${username}','pro')">Duyệt Chuyên Nghiệp</button>
                <button onclick="quickApprove('${username}','business')">Duyệt Doanh Nghiệp</button>
                <button onclick="quickApprove('${username}','lifetime')">Duyệt Vĩnh Viễn</button>
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
</script>
</body>
</html>
"""

@app.get("/admin-data")
def admin_data():
    return {
        "users": load_users(),
        "support": load_json(SUPPORT_FILE)
    }

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
    save_users(users)

    return {
        "success": True,
        "message": f"Đã cập nhật tài khoản {username} thành gói {PLAN_LABELS.get(req.plan, req.plan)}."
    }

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8005)
