from flask import Flask, request, render_template_string, jsonify, send_file
from dotenv import load_dotenv
import google.generativeai as genai
from werkzeug.utils import secure_filename
import os
import json
import requests
import sqlite3
import datetime
import threading
import time
import tempfile
import csv
import random
import io
import base64
import shutil

try:
    import openpyxl
except Exception:
    openpyxl = None

load_dotenv()

APP_TITLE = "Mkt Automation Pro V5 Enterprise Seller AI Suite"
DB = "marketing_automation_pro_v11.db"
UPLOAD_DIR = "uploads"
REPORT_DIR = "reports"
BACKUP_DIR = "backups"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(REPORT_DIR, exist_ok=True)
os.makedirs(BACKUP_DIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
PAGES_JSON = os.getenv("PAGES_JSON", "[]").strip()

# PAGES_JSON vẫn được giữ làm dữ liệu dự phòng.
# Fanpage mới sẽ được thêm trực tiếp trong tool và lưu vào SQLite.
def load_env_pages():
    try:
        data = json.loads(PAGES_JSON or "[]")
        if isinstance(data, list):
            return data
        return []
    except Exception:
        return []

PAGES = load_env_pages()

if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

client = genai if GEMINI_API_KEY else None


CONTENT_LIBRARY = {
    "spa": [
        "Làn da đẹp không đến từ may mắn, mà đến từ chăm sóc đúng cách. Inbox để được tư vấn liệu trình phù hợp.",
        "Da xỉn màu, thiếu sức sống? Một liệu trình chăm sóc chuyên sâu có thể giúp bạn tự tin hơn mỗi ngày.",
        "Dành chút thời gian chăm sóc bản thân hôm nay, bạn sẽ thấy mình rạng rỡ hơn vào ngày mai.",
        "Không cần trang điểm quá nhiều khi làn da đã được chăm sóc đúng cách.",
        "Bạn xứng đáng có một làn da khỏe, mịn và đầy sức sống."
    ],
    "nha_khoa": [
        "Nụ cười đẹp giúp bạn tự tin hơn trong giao tiếp mỗi ngày. Inbox để được tư vấn nha khoa.",
        "Răng đều, trắng sáng và khỏe mạnh bắt đầu từ việc thăm khám đúng lúc.",
        "Đừng để vấn đề răng miệng nhỏ trở thành nỗi lo lớn.",
        "Một nụ cười tự tin có thể thay đổi cách người khác nhìn bạn.",
        "Chăm sóc răng miệng định kỳ là đầu tư cho sức khỏe và sự tự tin lâu dài."
    ],
    "bat_dong_san": [
        "Cơ hội sở hữu bất động sản vị trí đẹp, pháp lý rõ ràng, tiềm năng tăng giá tốt.",
        "Bạn đang tìm căn hộ để ở hoặc đầu tư? Nhắn tin để nhận bảng giá mới.",
        "Bất động sản tốt không chờ đợi quá lâu. Liên hệ ngay để được tư vấn.",
        "Vị trí đẹp, tiện ích đầy đủ, pháp lý minh bạch.",
        "Đầu tư bất động sản cần đúng thời điểm và đúng sản phẩm."
    ],
    "facebook_ads": [
        "Chạy quảng cáo không ra đơn không phải lúc nào cũng do sản phẩm.",
        "Inbox cao, đơn thấp? Đã đến lúc xem lại nội dung, tệp khách hàng và hành trình chốt đơn.",
        "Một chiến dịch quảng cáo hiệu quả cần đúng thông điệp, đúng khách hàng và đúng thời điểm.",
        "Đừng để ngân sách quảng cáo bị tiêu hao mà không tạo ra kết quả.",
        "Muốn giảm chi phí tin nhắn và tăng tỷ lệ chốt đơn? Bắt đầu từ nội dung và tệp khách."
    ],
    "proxy": [
        "Cần proxy ổn định cho công việc online? Tham khảo gói proxy tốc độ cao, hỗ trợ nhanh.",
        "Proxy phù hợp giúp công việc quảng cáo, quản lý tài khoản và automation ổn định hơn.",
        "Tìm proxy Việt Nam ổn định, tốc độ tốt, hỗ trợ nhanh? Nhắn tin để nhận bảng giá.",
        "Giải pháp proxy linh hoạt cho cá nhân và doanh nghiệp.",
        "Proxy chất lượng giúp giảm gián đoạn khi làm việc nhiều tài khoản."
    ],
    "tiktok_shop": [
        "Muốn TikTok Shop vận hành hiệu quả hơn? Bắt đầu từ nội dung đúng sản phẩm, đúng khách hàng.",
        "Sản phẩm tốt cần nội dung đủ cuốn hút để khách dừng lại, xem và mua.",
        "TikTok Shop không chỉ cần sản phẩm, mà cần cách trình bày khiến khách muốn mua ngay.",
        "Tối ưu nội dung, hình ảnh và kịch bản bán hàng giúp shop tăng chuyển đổi.",
        "Đang bán TikTok Shop nhưng chưa ra đơn đều? Hãy tối ưu nội dung và phễu bán hàng."
    ],
    "website": [
        "Website chuyên nghiệp giúp khách hàng tin tưởng hơn trước khi quyết định mua hàng.",
        "Một website tốt không chỉ đẹp, mà còn phải rõ thông tin, dễ dùng và hỗ trợ chuyển đổi.",
        "Bạn cần website bán hàng, giới thiệu dịch vụ hoặc landing page chạy quảng cáo?",
        "Đừng để khách rời đi vì website thiếu chuyên nghiệp.",
        "Website là tài sản số quan trọng cho mọi doanh nghiệp."
    ],
    "noi_that": [
        "Không gian đẹp bắt đầu từ thiết kế phù hợp nhu cầu sống và gu thẩm mỹ của bạn.",
        "Một bộ sofa phù hợp có thể thay đổi cảm giác của cả căn phòng.",
        "Nội thất không chỉ để dùng, mà còn tạo nên phong cách sống.",
        "Tối ưu không gian sống với giải pháp nội thất hiện đại, tiện nghi và thẩm mỹ.",
        "Bạn muốn nhà đẹp hơn nhưng chưa biết bắt đầu từ đâu? Inbox để được tư vấn."
    ],
    "oto": [
        "Chăm sóc xe định kỳ giúp xe bền hơn, vận hành ổn định hơn và giữ giá trị tốt hơn.",
        "Xe sạch, bóng, nội thất thơm tho giúp mỗi chuyến đi dễ chịu hơn.",
        "Phụ kiện phù hợp giúp chiếc xe tiện nghi, an toàn và cá tính hơn.",
        "Bạn đang tìm camera hành trình, phụ kiện hoặc dịch vụ chăm sóc xe?",
        "Đừng chờ xe xuống cấp mới chăm sóc. Bảo dưỡng đúng lúc giúp tiết kiệm chi phí."
    ],
    "giao_duc": [
        "Học đúng phương pháp giúp tiết kiệm thời gian và cải thiện kết quả nhanh hơn.",
        "Bạn muốn nâng cấp kỹ năng nhưng chưa biết bắt đầu từ đâu?",
        "Một khóa học tốt không chỉ cung cấp kiến thức, mà còn giúp ứng dụng thực tế.",
        "Đầu tư vào kiến thức là khoản đầu tư có giá trị lâu dài.",
        "Học online linh hoạt, tiết kiệm thời gian, phù hợp người bận rộn."
    ],
    "du_lich": [
        "Một chuyến đi đẹp bắt đầu từ kế hoạch phù hợp.",
        "Bạn muốn du lịch thư giãn nhưng ngại tự lên lịch?",
        "Khám phá điểm đến mới với lịch trình tối ưu, chi phí hợp lý.",
        "Cần đặt tour, khách sạn, vé máy bay hoặc visa?",
        "Đi để nghỉ ngơi, trải nghiệm và nạp lại năng lượng."
    ]
}


# Kho content V9 mở rộng 15 ngành. Demo hiển thị 10 mẫu miễn phí/ngành, phần còn lại khóa Premium.
EXTENDED_CONTENT_LIBRARY = {
    "thoi_trang": [
        "Phong cách đẹp bắt đầu từ outfit phù hợp. Inbox để được tư vấn mẫu mới hôm nay.",
        "Một set đồ chỉn chu giúp bạn tự tin hơn trong mọi cuộc gặp.",
        "Thời trang không chỉ là mặc đẹp, mà còn là cách bạn thể hiện cá tính.",
        "Mẫu mới về liên tục, chất vải đẹp, form dễ mặc. Nhắn tin để xem ảnh thật.",
        "Đổi mới phong cách mỗi ngày với những thiết kế trẻ trung, dễ phối.",
        "Bạn cần outfit đi làm, đi chơi hay dự tiệc? Inbox để được gợi ý.",
        "Hàng đẹp, form chuẩn, giá hợp lý. Số lượng có hạn, nhắn tin ngay.",
        "Tủ đồ của bạn sẽ thú vị hơn với những mẫu mới đang hot.",
        "Đẹp tự nhiên, mặc thoải mái, phù hợp nhiều dáng người.",
        "Inbox để nhận bảng mẫu mới và ưu đãi hôm nay."
    ],
    "giay_dep": [
        "Một đôi giày phù hợp có thể nâng tầm cả outfit của bạn.",
        "Giày đẹp, êm chân, dễ phối đồ cho cả đi làm và đi chơi.",
        "Bạn đang tìm giày bền, đẹp, giá hợp lý? Nhắn tin để xem mẫu.",
        "Mẫu mới về liên tục, size đầy đủ, hỗ trợ tư vấn chọn size.",
        "Đừng để đôi giày không thoải mái làm bạn mất tự tin.",
        "Giày chuẩn form, đi êm, phù hợp nhiều phong cách.",
        "Inbox để nhận ảnh thật và bảng size chi tiết.",
        "Nâng cấp phong cách hằng ngày chỉ từ một đôi giày đẹp.",
        "Chọn đúng giày, mỗi bước đi đều tự tin hơn.",
        "Mẫu hot hôm nay số lượng có hạn, nhắn tin giữ size."
    ],
    "tui_xach": [
        "Một chiếc túi đẹp giúp outfit của bạn nổi bật hơn.",
        "Túi xách thời trang, dễ phối đồ, phù hợp đi làm và đi chơi.",
        "Bạn cần túi sang, đẹp, tiện dụng? Inbox để xem mẫu mới.",
        "Thiết kế tinh tế, chất liệu đẹp, phù hợp nhiều phong cách.",
        "Túi không chỉ để đựng đồ, mà còn tạo điểm nhấn cho phong cách.",
        "Mẫu mới đang có sẵn, nhắn tin để nhận ảnh thật.",
        "Túi đẹp giúp bạn tự tin hơn mỗi khi ra ngoài.",
        "Phối đồ dễ hơn với những mẫu túi thanh lịch, hiện đại.",
        "Số lượng có hạn, inbox để được tư vấn mẫu phù hợp.",
        "Tặng bản thân một chiếc túi mới để làm mới phong cách."
    ],
    "dong_ho": [
        "Đồng hồ đẹp là điểm nhấn tinh tế cho phong cách của bạn.",
        "Một chiếc đồng hồ phù hợp giúp bạn trông chỉn chu hơn mỗi ngày.",
        "Thiết kế sang trọng, dễ phối, phù hợp đi làm và gặp khách hàng.",
        "Bạn cần đồng hồ làm quà tặng? Inbox để được tư vấn mẫu phù hợp.",
        "Đồng hồ không chỉ xem giờ, mà còn thể hiện gu thẩm mỹ.",
        "Mẫu mới về, thiết kế đẹp, giá hợp lý.",
        "Tối giản nhưng nổi bật, phù hợp nhiều phong cách.",
        "Inbox để xem ảnh thật và nhận tư vấn lựa chọn.",
        "Một món phụ kiện nhỏ nhưng tạo khác biệt lớn.",
        "Chọn đồng hồ phù hợp để hoàn thiện vẻ ngoài chuyên nghiệp."
    ],
    "phu_kien": [
        "Phụ kiện nhỏ có thể làm outfit của bạn nổi bật hơn rất nhiều.",
        "Đừng bỏ qua chi tiết, vì phụ kiện tạo nên phong cách riêng.",
        "Mẫu phụ kiện mới, dễ phối, phù hợp nhiều phong cách.",
        "Inbox để xem mẫu hot và nhận tư vấn phối đồ.",
        "Một món phụ kiện đẹp giúp bạn tự tin hơn khi ra ngoài.",
        "Thiết kế tinh tế, giá dễ mua, phù hợp làm quà tặng.",
        "Nâng cấp outfit chỉ với một điểm nhấn nhỏ.",
        "Phụ kiện đẹp giúp phong cách của bạn có dấu ấn riêng.",
        "Mẫu mới cập nhật mỗi ngày, nhắn tin để xem ngay.",
        "Chọn phụ kiện phù hợp để hoàn thiện vẻ ngoài của bạn."
    ],
    "tham_my_vien": [
        "Vẻ đẹp tự tin bắt đầu từ lựa chọn chăm sóc đúng cách.",
        "Bạn muốn cải thiện diện mạo nhưng chưa biết bắt đầu từ đâu? Inbox để được tư vấn.",
        "Thẩm mỹ an toàn, tư vấn rõ ràng, phù hợp từng nhu cầu.",
        "Làm đẹp là đầu tư cho sự tự tin và chất lượng cuộc sống.",
        "Đừng để khuyết điểm nhỏ làm bạn kém tự tin mỗi ngày.",
        "Đội ngũ tư vấn sẽ giúp bạn chọn giải pháp phù hợp nhất.",
        "Inbox để nhận tư vấn liệu trình và ưu đãi mới.",
        "Chăm sóc vẻ ngoài đúng cách giúp bạn rạng rỡ hơn.",
        "Đẹp tự nhiên, an toàn, phù hợp với từng khách hàng.",
        "Đặt lịch tư vấn để hiểu rõ giải pháp phù hợp."
    ],
    "trung_tam_tieng_anh": [
        "Tiếng Anh tốt mở ra nhiều cơ hội học tập và công việc hơn.",
        "Bạn mất gốc tiếng Anh? Bắt đầu lại với lộ trình phù hợp.",
        "Học đúng phương pháp giúp bạn tiến bộ nhanh và bền vững.",
        "Lớp học linh hoạt, giáo viên hỗ trợ sát sao, phù hợp người bận rộn.",
        "Inbox để kiểm tra trình độ và nhận tư vấn lộ trình học.",
        "Đừng học lan man, hãy học theo mục tiêu rõ ràng.",
        "Giao tiếp tự tin hơn với lộ trình học thực tế.",
        "Tiếng Anh không khó nếu bạn có phương pháp phù hợp.",
        "Đăng ký tư vấn miễn phí để chọn lớp phù hợp.",
        "Nâng cấp tiếng Anh hôm nay để mở rộng cơ hội ngày mai."
    ],
    "khoa_hoc_online": [
        "Học online linh hoạt giúp bạn nâng cấp kỹ năng mọi lúc mọi nơi.",
        "Đầu tư vào kiến thức là khoản đầu tư có giá trị lâu dài.",
        "Bạn muốn học nhưng không có nhiều thời gian? Khóa online là lựa chọn phù hợp.",
        "Nội dung dễ hiểu, ứng dụng thực tế, học theo tốc độ của bạn.",
        "Inbox để nhận lộ trình học phù hợp với mục tiêu.",
        "Nâng cấp kỹ năng để tăng cơ hội công việc và thu nhập.",
        "Học đúng thứ bạn cần, tiết kiệm thời gian và chi phí.",
        "Khóa học thiết kế cho người mới bắt đầu và người muốn nâng cao.",
        "Bắt đầu hôm nay, kết quả sẽ khác sau vài tuần.",
        "Nhắn tin để nhận tư vấn khóa học phù hợp."
    ],
    "showroom_oto": [
        "Tìm xe phù hợp nhu cầu và tài chính chưa bao giờ dễ hơn.",
        "Showroom cập nhật nhiều mẫu xe mới, hỗ trợ tư vấn tận tình.",
        "Bạn cần xe gia đình, xe dịch vụ hay xe cá nhân? Inbox để được tư vấn.",
        "Xe đẹp, hồ sơ rõ ràng, hỗ trợ thủ tục nhanh chóng.",
        "Lựa chọn xe đúng giúp bạn yên tâm hơn trên mọi hành trình.",
        "Inbox để nhận báo giá và chương trình ưu đãi mới nhất.",
        "Tư vấn tài chính, trả góp, hồ sơ nhanh gọn.",
        "Mẫu xe hot đang có sẵn, liên hệ để xem xe trực tiếp.",
        "Một chiếc xe phù hợp sẽ đồng hành cùng bạn lâu dài.",
        "Nhắn tin để được tư vấn mẫu xe phù hợp nhất."
    ],
    "phong_kham": [
        "Sức khỏe là tài sản quan trọng nhất, đừng chờ có triệu chứng mới kiểm tra.",
        "Thăm khám định kỳ giúp phát hiện và xử lý vấn đề sớm hơn.",
        "Phòng khám hỗ trợ tư vấn, đặt lịch nhanh, tiết kiệm thời gian.",
        "Bạn cần kiểm tra sức khỏe? Inbox để được hướng dẫn đặt lịch.",
        "Chăm sóc sức khỏe đúng lúc giúp bạn yên tâm hơn mỗi ngày.",
        "Dịch vụ thăm khám chuyên nghiệp, tư vấn rõ ràng.",
        "Đặt lịch trước để được hỗ trợ nhanh và thuận tiện hơn.",
        "Sức khỏe tốt bắt đầu từ sự chủ động của bạn.",
        "Inbox để nhận thông tin gói khám phù hợp.",
        "Đừng trì hoãn việc chăm sóc sức khỏe của bản thân."
    ],
}

CONTENT_LIBRARY.update(EXTENDED_CONTENT_LIBRARY)

INDUSTRY_LABELS = {
    "spa": "Spa / Chăm sóc da",
    "nha_khoa": "Nha khoa",
    "bat_dong_san": "Bất động sản",
    "facebook_ads": "Facebook Ads",
    "proxy": "Proxy",
    "tiktok_shop": "TikTok Shop",
    "website": "Thiết kế website",
    "noi_that": "Nội thất",
    "oto": "Ô tô - Chăm sóc xe",
    "giao_duc": "Giáo dục",
    "du_lich": "Du lịch",
    "thoi_trang": "Thời trang",
    "giay_dep": "Giày dép",
    "tui_xach": "Túi xách",
    "dong_ho": "Đồng hồ",
    "phu_kien": "Phụ kiện thời trang",
    "tham_my_vien": "Thẩm mỹ viện",
    "trung_tam_tieng_anh": "Trung tâm tiếng Anh",
    "khoa_hoc_online": "Khóa học online",
    "showroom_oto": "Showroom ô tô",
    "phong_kham": "Phòng khám"
}

def init_db():
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS posts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        page_name TEXT,
        page_id TEXT,
        content TEXT,
        status TEXT,
        post_id TEXT,
        schedule_time TEXT,
        image_path TEXT,
        campaign TEXT,
        score INTEGER,
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS campaigns (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        industry TEXT,
        goal TEXT,
        note TEXT,
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS crm (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        phone TEXT,
        zalo TEXT,
        source TEXT,
        note TEXT,
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS token_checks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        page_name TEXT,
        page_id TEXT,
        status TEXT,
        detail TEXT,
        checked_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS fanpages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        page_name TEXT,
        page_id TEXT UNIQUE,
        page_token TEXT,
        note TEXT,
        status TEXT DEFAULT 'active',
        created_at TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS page_clusters (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        page_names TEXT,
        note TEXT,
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS subscriptions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        package_name TEXT,
        start_date TEXT,
        end_date TEXT,
        status TEXT,
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS users_trial (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        package_name TEXT DEFAULT 'free',
        trial_start_date TEXT,
        trial_end_date TEXT,
        status TEXT DEFAULT 'active',
        created_at TEXT
    )
    """)

    # V3 Enterprise AI Suite tables
    c.execute("""
    CREATE TABLE IF NOT EXISTS lead_pipeline (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        customer_name TEXT,
        phone TEXT,
        zalo TEXT,
        source TEXT,
        stage TEXT DEFAULT 'Khách mới',
        value INTEGER DEFAULT 0,
        note TEXT,
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS customer_tasks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        customer_name TEXT,
        task TEXT,
        due_date TEXT,
        status TEXT DEFAULT 'pending',
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS notifications (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT,
        detail TEXT,
        level TEXT DEFAULT 'info',
        status TEXT DEFAULT 'new',
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS competitors (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        page_url TEXT,
        analysis TEXT,
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS sales_scripts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        topic TEXT,
        script_type TEXT,
        content TEXT,
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS landing_pages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        service_name TEXT,
        headline TEXT,
        content TEXT,
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS workflows (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        trigger_name TEXT,
        action_name TEXT,
        status TEXT DEFAULT 'draft',
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS user_roles (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT,
        role TEXT DEFAULT 'staff',
        note TEXT,
        created_at TEXT
    )
    """)

    # V5 Seller AI Suite tables
    c.execute("""
    CREATE TABLE IF NOT EXISTS fb_groups (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        group_name TEXT,
        group_id TEXT,
        niche TEXT,
        note TEXT,
        status TEXT DEFAULT 'active',
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS group_schedules (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        group_name TEXT,
        group_id TEXT,
        content TEXT,
        schedule_time TEXT,
        status TEXT DEFAULT 'scheduled',
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS comment_leads (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        customer_name TEXT,
        phone TEXT,
        comment_text TEXT,
        ai_reply TEXT,
        label TEXT,
        crm_status TEXT DEFAULT 'new',
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS messenger_scripts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product TEXT,
        script_type TEXT,
        content TEXT,
        created_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS success_assets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        asset_type TEXT,
        title TEXT,
        content TEXT,
        created_at TEXT
    )
    """)

    # V6 Premium approval center tables
    c.execute("""
    CREATE TABLE IF NOT EXISTS premium_requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        device_id TEXT,
        phone TEXT,
        email TEXT,
        plan_key TEXT,
        plan_name TEXT,
        amount INTEGER DEFAULT 0,
        transaction_note TEXT,
        status TEXT DEFAULT 'pending',
        premium_start TEXT,
        premium_end TEXT,
        created_at TEXT,
        approved_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS customer_devices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        device_id TEXT UNIQUE,
        phone TEXT,
        email TEXT,
        status TEXT DEFAULT 'pending',
        created_at TEXT,
        updated_at TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS support_messages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        device_id TEXT,
        phone TEXT,
        email TEXT,
        sender TEXT DEFAULT 'customer',
        message TEXT,
        admin_reply TEXT,
        status TEXT DEFAULT 'new',
        created_at TEXT,
        replied_at TEXT
    )
    """)

    conn.commit()
    conn.close()

def db():
    return sqlite3.connect(DB)


def ensure_premium_tables():
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS premium_requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        device_id TEXT,
        phone TEXT,
        email TEXT,
        plan_key TEXT,
        plan_name TEXT,
        amount INTEGER DEFAULT 0,
        transaction_note TEXT,
        status TEXT DEFAULT 'pending',
        premium_start TEXT,
        premium_end TEXT,
        created_at TEXT,
        approved_at TEXT
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS customer_devices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        device_id TEXT UNIQUE,
        phone TEXT,
        email TEXT,
        status TEXT DEFAULT 'pending',
        created_at TEXT,
        updated_at TEXT
    )
    """)
    conn.commit()
    conn.close()


def plan_days(plan_key):
    mapping = {
        "monthly": 30,
        "basic": 30,
        "quarterly": 90,
        "pro": 90,
        "halfyear": 180,
        "business": 180,
        "yearly": 365,
        "lifetime": 3650,
        "sellerpro": 3650,
    }
    return mapping.get(str(plan_key or "").lower(), 30)


def create_premium_request(device_id, phone, email, plan_key, plan_name, amount, transaction_note=""):
    ensure_premium_tables()
    device_id = (device_id or "").strip()
    phone = (phone or "").strip()
    email = (email or "").strip().lower()
    plan_key = (plan_key or "monthly").strip()
    plan_name = (plan_name or plan_key).strip()
    transaction_note = (transaction_note or "").strip()

    if not device_id:
        return False, "Thiếu ID thiết bị. Vui lòng tải lại trang rồi gửi lại."
    if not phone:
        return False, "Vui lòng nhập số điện thoại/Zalo."
    if not email or "@" not in email:
        return False, "Vui lòng nhập đúng Gmail/Email."

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    INSERT INTO premium_requests(device_id,phone,email,plan_key,plan_name,amount,transaction_note,status,created_at)
    VALUES(?,?,?,?,?,?,?,?,?)
    """, (device_id, phone, email, plan_key, plan_name, int(amount or 0), transaction_note, "pending", now))
    c.execute("""
    INSERT INTO customer_devices(device_id,phone,email,status,created_at,updated_at)
    VALUES(?,?,?,?,?,?)
    ON CONFLICT(device_id) DO UPDATE SET
        phone=excluded.phone,
        email=excluded.email,
        status='pending',
        updated_at=excluded.updated_at
    """, (device_id, phone, email, "pending", now, now))
    conn.commit()
    conn.close()
    return True, "Đã gửi yêu cầu kích hoạt Premium. Bộ phận quản trị sẽ kiểm tra và duyệt trực tiếp trên hệ thống."


def get_premium_status_by_device(device_id):
    ensure_premium_tables()
    device_id = (device_id or "").strip()
    if not device_id:
        return {"active": False, "status": "missing", "message": "Thiếu ID thiết bị"}

    now = datetime.datetime.now()
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    SELECT plan_key,plan_name,premium_start,premium_end,status,approved_at
    FROM premium_requests
    WHERE device_id=? AND status='approved'
    ORDER BY id DESC LIMIT 1
    """, (device_id,))
    row = c.fetchone()
    conn.close()
    if not row:
        return {"active": False, "status": "pending_or_trial", "message": "Dùng thử / chờ duyệt"}

    plan_key, plan_name, start, end, status, approved_at = row
    if end:
        try:
            end_dt = datetime.datetime.strptime(str(end)[:19], "%Y-%m-%d %H:%M:%S")
            if now > end_dt:
                return {"active": False, "status": "expired", "plan": plan_name, "premium_end": end, "message": "Premium đã hết hạn"}
        except Exception:
            pass
    return {"active": True, "status": "approved", "plan_key": plan_key, "plan": plan_name, "premium_start": start, "premium_end": end, "approved_at": approved_at, "message": "Premium đang hoạt động"}


def current_device_id():
    """Lấy ID thiết bị từ form/query/cookie để backend nhận biết Premium."""
    try:
        return (
            request.form.get("device_id")
            or request.args.get("device_id")
            or request.cookies.get("mkt_device_id")
            or ""
        ).strip()
    except Exception:
        return ""


def current_premium_status():
    did = current_device_id()
    if not did:
        return {"active": False, "status": "missing", "message": "Thiếu ID thiết bị"}
    return get_premium_status_by_device(did)


def is_current_premium_active():
    return bool(current_premium_status().get("active"))


def get_effective_access_status():
    """Trạng thái quyền dùng sau khi gộp trial + Premium đã duyệt."""
    premium = current_premium_status()
    if premium.get("active"):
        return {
            "package_name": "premium",
            "status": "active",
            "days": 999,
            "hours": 0,
            "percent": 100,
            "is_trial": False,
            "is_expired": False,
            "label": "Premium: 👑 GÓI NHÀ BÁN HÀNG CHUYÊN NGHIỆP",
            "note": "Đã mở toàn bộ tính năng Premium cho thiết bị này.",
            "allowed_features": ["ALL"],
            "locked_features": [],
            "premium": premium,
        }
    return None


def get_premium_requests(limit=80):
    ensure_premium_tables()
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    SELECT id,device_id,phone,email,plan_key,plan_name,amount,transaction_note,status,premium_start,premium_end,created_at,approved_at
    FROM premium_requests
    ORDER BY id DESC LIMIT ?
    """, (limit,))
    rows = c.fetchall()
    conn.close()
    return rows


def approve_premium_request(request_id):
    ensure_premium_tables()
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("SELECT id,device_id,plan_key,plan_name FROM premium_requests WHERE id=?", (int(request_id),))
    row = c.fetchone()
    if not row:
        conn.close()
        return False, "Không tìm thấy yêu cầu Premium."
    _, device_id, plan_key, plan_name = row
    start = datetime.datetime.now()
    end = start + datetime.timedelta(days=plan_days(plan_key))
    start_s = start.strftime("%Y-%m-%d %H:%M:%S")
    end_s = end.strftime("%Y-%m-%d %H:%M:%S")
    c.execute("""
    UPDATE premium_requests
    SET status='approved', premium_start=?, premium_end=?, approved_at=?
    WHERE id=?
    """, (start_s, end_s, start_s, int(request_id)))
    c.execute("UPDATE customer_devices SET status='approved', updated_at=? WHERE device_id=?", (start_s, device_id))
    conn.commit()
    conn.close()
    return True, f"Đã mở khóa Premium cho thiết bị {device_id} đến {end_s}."


def ensure_fanpages_table():
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS fanpages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        page_name TEXT,
        page_id TEXT UNIQUE,
        page_token TEXT,
        note TEXT,
        status TEXT DEFAULT 'active',
        created_at TEXT
    )
    """)
    conn.commit()
    conn.close()


def mask_token(token):
    token = str(token or "").strip()
    if not token:
        return "Thiếu token"
    if len(token) <= 14:
        return token[:4] + "..."
    return token[:8] + "..." + token[-6:]


def get_fanpages(include_env=True):
    """Lấy Fanpage từ SQLite, kèm PAGES_JSON cũ làm fallback nếu cần."""
    ensure_fanpages_table()
    pages = []
    seen = set()

    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    SELECT id,page_name,page_id,page_token,note,status,created_at
    FROM fanpages
    ORDER BY id DESC
    """)
    for row in c.fetchall():
        page_id = str(row[2] or "").strip()
        if not page_id:
            continue
        seen.add(page_id)
        pages.append({
            "row_id": row[0],
            "name": row[1] or "Fanpage chưa đặt tên",
            "id": page_id,
            "token": row[3] or "",
            "token_mask": mask_token(row[3]),
            "note": row[4] or "",
            "status": row[5] or "active",
            "created_at": row[6] or "",
            "source": "database"
        })
    conn.close()

    if include_env:
        for p in load_env_pages():
            page_id = str(p.get("id", "")).strip()
            if not page_id or page_id in seen:
                continue
            pages.append({
                "row_id": "",
                "name": p.get("name", "Fanpage .env"),
                "id": page_id,
                "token": p.get("token", ""),
                "token_mask": mask_token(p.get("token", "")),
                "note": "Từ PAGES_JSON trong Environment",
                "status": "env",
                "created_at": "",
                "source": "env"
            })
    return pages


def add_fanpage_token(page_name, page_id, page_token, note=""):
    page_name = (page_name or "").strip()
    page_id = (page_id or "").strip()
    page_token = (page_token or "").strip()
    note = (note or "").strip()

    if not page_name or not page_id or not page_token:
        return False, "Vui lòng nhập đủ Tên Fanpage, Page ID và Page Token."

    ensure_fanpages_table()
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        c.execute("""
        INSERT INTO fanpages(page_name,page_id,page_token,note,status,created_at)
        VALUES(?,?,?,?,?,?)
        ON CONFLICT(page_id) DO UPDATE SET
            page_name=excluded.page_name,
            page_token=excluded.page_token,
            note=excluded.note,
            status='active'
        """, (page_name, page_id, page_token, note, "active", now))
        conn.commit()
        return True, f"Đã lưu Fanpage {page_name}."
    except Exception as e:
        return False, "Lỗi lưu Fanpage: " + str(e)
    finally:
        conn.close()


def delete_fanpage_token(row_id):
    try:
        row_id = int(row_id)
    except Exception:
        return False, "ID Fanpage không hợp lệ."
    ensure_fanpages_table()
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("DELETE FROM fanpages WHERE id=?", (row_id,))
    conn.commit()
    conn.close()
    return True, "Đã xóa Fanpage khỏi Token Center."

def save_post(page_name, page_id, content, status, post_id="", schedule_time="", image_path="", campaign="", score=0):
    conn = db(); c = conn.cursor()
    c.execute("""
    INSERT INTO posts(page_name,page_id,content,status,post_id,schedule_time,image_path,campaign,score,created_at)
    VALUES(?,?,?,?,?,?,?,?,?,?)
    """, (page_name, page_id, content, status, post_id, schedule_time, image_path, campaign, score,
          datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit(); conn.close()

def update_post(row_id, status, post_id=""):
    conn = db(); c = conn.cursor()
    c.execute("UPDATE posts SET status=?, post_id=? WHERE id=?", (status, post_id, row_id))
    conn.commit(); conn.close()

def get_history(limit=60):
    conn = db(); c = conn.cursor()
    c.execute("""
    SELECT id,page_name,content,status,post_id,schedule_time,image_path,campaign,score,created_at
    FROM posts ORDER BY id DESC LIMIT ?
    """, (limit,))
    rows = c.fetchall(); conn.close()
    return rows

def get_campaigns():
    conn = db(); c = conn.cursor()
    c.execute("SELECT id,name,industry,goal,note,created_at FROM campaigns ORDER BY id DESC")
    rows = c.fetchall(); conn.close()
    return rows

def add_campaign(name, industry, goal, note):
    if not name:
        return
    conn = db(); c = conn.cursor()
    c.execute("""
    INSERT OR IGNORE INTO campaigns(name,industry,goal,note,created_at)
    VALUES(?,?,?,?,?)
    """, (name, industry, goal, note, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit(); conn.close()

def get_stats():
    conn = db(); c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM posts"); total = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM posts WHERE status='posted'"); posted = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM posts WHERE status='scheduled'"); scheduled = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM posts WHERE status='error'"); error = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM campaigns"); campaigns = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM crm")
    crm = c.fetchone()[0]
    conn.close()
    return {"total": total, "posted": posted, "scheduled": scheduled, "error": error, "campaigns": campaigns, "crm": crm}

def selected_pages(page_indexes):
    result = []
    pages = get_fanpages()
    for idx in page_indexes:
        try:
            i = int(idx)
            if 0 <= i < len(pages):
                result.append(pages[i])
        except Exception:
            pass
    return result

def score_content(content):
    score = 50
    text = content.lower()
    if len(content) >= 120: score += 10
    if any(x in text for x in ["inbox", "nhắn tin", "liên hệ", "đặt lịch", "zalo"]): score += 15
    if any(x in content for x in ["🔥", "🚀", "📩", "💎", "✨"]): score += 5
    if "#" in content: score += 5
    if any(x in text for x in ["cam kết 100%", "chắc chắn", "đảm bảo ra đơn", "trị khỏi", "hiệu quả tuyệt đối"]): score -= 25
    return max(0, min(100, score))

def policy_check(content):
    warnings = []
    risky = ["cam kết 100%", "chắc chắn", "đảm bảo ra đơn", "trị khỏi", "hiệu quả tuyệt đối", "thần tốc"]
    for word in risky:
        if word in content.lower():
            warnings.append(f"Nên tránh cụm từ: {word}")
    if content.count("#") > 5:
        warnings.append("Hashtag hơi nhiều, nên giảm xuống 2-4 hashtag.")
    if len(content) > 1200:
        warnings.append("Nội dung khá dài, nên rút gọn để dễ đọc hơn.")
    if not warnings:
        warnings.append("Nội dung tương đối an toàn, vẫn nên kiểm tra lại trước khi đăng.")
    return warnings


def friendly_ai_error(error):
    raw = str(error)
    low = raw.lower()

    if "429" in raw or "resource_exhausted" in low or "quota" in low or "rate-limit" in low or "rate limit" in low:
        return (
            "Hệ thống AI đang quá tải hoặc đã đạt giới hạn xử lý tạm thời. "
            "Vui lòng thử lại sau ít phút. Nếu anh/chị cần sử dụng ổn định hơn, "
            "vui lòng nâng cấp Premium để được ưu tiên hỗ trợ và hạn chế gián đoạn."
        )

    if "api_key" in low or "invalid" in low or "permission" in low or "unauthorized" in low:
        return (
            "Hệ thống AI chưa được cấu hình đầy đủ hoặc khóa API chưa hợp lệ. "
            "Vui lòng liên hệ kỹ thuật để được kiểm tra và kích hoạt lại."
        )

    if "timeout" in low or "connection" in low or "network" in low:
        return (
            "Kết nối AI đang chậm hoặc mạng tạm thời không ổn định. "
            "Vui lòng thử lại sau ít phút."
        )

    return (
        "Hệ thống AI đang bận xử lý. Vui lòng thử lại sau ít phút "
        "hoặc liên hệ Zalo 036 338 2629 để được hỗ trợ nhanh."
    )

def safe_ai_generate(prompt, fallback=""):
    if not client:
        if fallback:
            return fallback
        return (
            "Hệ thống AI chưa được kích hoạt. Vui lòng kiểm tra GEMINI_API_KEY "
            "trong file .env hoặc liên hệ kỹ thuật để được hỗ trợ."
        )
    try:
        model = genai.GenerativeModel("gemini-1.5-flash")
        res = model.generate_content(prompt)
        return (res.text or "").strip()
    except Exception as e:
        if fallback:
            return fallback + "\\n\\n" + friendly_ai_error(e)
        return friendly_ai_error(e)

def generate_content(idea, style="gần gũi", length="120", variants=1):
    fallback = f"""
Nội dung gợi ý:

{idea}

Anh/chị có thể dùng nội dung này làm bản nháp, sau đó chỉnh lại theo sản phẩm/dịch vụ thực tế.

CTA:
Inbox để được tư vấn chi tiết.

Hashtag:
#Marketing #KinhDoanhOnline
""".strip()

    prompt = f"""
Viết {variants} bài đăng Facebook bằng tiếng Việt.

Chủ đề: {idea}
Phong cách: {style}
Độ dài mỗi bài: dưới {length} từ

Yêu cầu:
- Tự nhiên, dễ hiểu, phù hợp bán hàng/dịch vụ online
- Có emoji vừa phải
- Có lời kêu gọi inbox/liên hệ
- Không spam hashtag
- Không cam kết quá đà
- Không vi phạm chính sách Facebook
- Nếu viết nhiều phiên bản, đánh số Phiên bản 1, Phiên bản 2...
"""
    return safe_ai_generate(prompt, fallback)

def ai_plan_30_days(industry, goal):
    if not client:
        # fallback không tốn AI
        lines = []
        base = CONTENT_LIBRARY.get(industry, CONTENT_LIBRARY["spa"])
        for i in range(1, 31):
            lines.append(f"Ngày {i}: {random.choice(base)}")
        return "\n".join(lines)
    prompt = f"""
Tạo kế hoạch marketing Facebook 30 ngày bằng tiếng Việt.
Ngành: {industry}
Mục tiêu: {goal}

Yêu cầu:
- Mỗi ngày 1 ý tưởng bài đăng
- Có CTA ngắn
- Có hashtag gợi ý
- Trình bày theo Ngày 1 đến Ngày 30
- Phù hợp chủ shop/doanh nghiệp nhỏ
"""
    return safe_ai_generate(prompt, "\n".join(lines) if 'lines' in locals() else "")

def analyze_fanpage(name, avatar, cover, bio, post_frequency, cta):
    score = 50
    notes = []
    if avatar == "co": score += 10
    else: notes.append("Nên có ảnh đại diện rõ thương hiệu.")
    if cover == "co": score += 10
    else: notes.append("Nên có ảnh bìa thể hiện dịch vụ/sản phẩm chính.")
    if len(bio.strip()) > 50: score += 10
    else: notes.append("Phần mô tả Fanpage còn ngắn, nên bổ sung lợi ích và thông tin liên hệ.")
    if post_frequency in ["hang_ngay", "3_5_bai_tuan"]: score += 10
    else: notes.append("Tần suất đăng chưa đều, nên có lịch đăng cố định.")
    if cta == "co": score += 10
    else: notes.append("Nên có CTA rõ: Inbox, Zalo, Website hoặc đặt lịch.")
    score = min(score, 100)
    if not notes:
        notes.append("Fanpage đang có nền tảng tốt, nên tối ưu thêm nội dung và lịch đăng.")
    return score, notes

def spin_content_local(content):
    endings = [
        "Nhắn tin ngay để được tư vấn chi tiết hơn.",
        "Inbox để nhận thông tin và báo giá phù hợp.",
        "Liên hệ ngay hôm nay để được hỗ trợ nhanh.",
        "Để lại tin nhắn, đội ngũ tư vấn sẽ hỗ trợ bạn.",
        "Bạn cần thêm thông tin? Inbox ngay nhé."
    ]
    text = content.strip().replace("Inbox", "Nhắn tin").replace("Liên hệ", "Kết nối")
    return text + "\n\n" + random.choice(endings)

def post_to_facebook(page, content, image_path=""):
    if not page.get("token"):
        return {"error": {"message": "Thiếu Page Token. Vui lòng vào Token Fanpage để thêm token.", "code": "MISSING_TOKEN"}}
    if not page.get("id"):
        return {"error": {"message": "Thiếu Page ID. Vui lòng kiểm tra lại Fanpage.", "code": "MISSING_PAGE_ID"}}
    if image_path and os.path.exists(image_path):
        ext = os.path.splitext(image_path)[1].lower()
        if ext in [".mp4", ".mov", ".m4v"]:
            url = f"https://graph-video.facebook.com/v23.0/{page['id']}/videos"
            with open(image_path, "rb") as video:
                res = requests.post(url, data={"description": content, "access_token": page["token"]},
                                    files={"source": video}, timeout=300)
            return res.json()
        url = f"https://graph.facebook.com/v23.0/{page['id']}/photos"
        with open(image_path, "rb") as img:
            res = requests.post(url, data={"message": content, "access_token": page["token"]},
                                files={"source": img}, timeout=90)
        return res.json()

    url = f"https://graph.facebook.com/v23.0/{page['id']}/feed"
    res = requests.post(url, data={"message": content, "access_token": page["token"]}, timeout=60)
    return res.json()


def save_uploads(file_list):
    paths = []
    for file_obj in file_list:
        if not file_obj or file_obj.filename == "":
            continue
        filename = secure_filename(file_obj.filename)
        path = os.path.join(UPLOAD_DIR, datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f_") + filename)
        file_obj.save(path)
        paths.append(path)
    return paths

def pick_media_for_job(media_paths, index):
    if not media_paths:
        return ""
    return media_paths[index % len(media_paths)]

def save_upload(file_obj):
    if not file_obj or file_obj.filename == "":
        return ""
    filename = secure_filename(file_obj.filename)
    path = os.path.join(UPLOAD_DIR, datetime.datetime.now().strftime("%Y%m%d_%H%M%S_") + filename)
    file_obj.save(path)
    return path

def scheduler_loop():
    while True:
        try:
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            conn = db(); c = conn.cursor()
            c.execute("""
            SELECT id,page_name,page_id,content,schedule_time,image_path
            FROM posts
            WHERE status='scheduled' AND schedule_time<=?
            """, (now,))
            jobs = c.fetchall(); conn.close()
            for row_id, page_name, page_id, content, schedule_time, image_path in jobs:
                page = next((p for p in get_fanpages() if str(p["id"]) == str(page_id)), None)
                if not page:
                    update_post(row_id, "error", "Không tìm thấy Page trong Token Center")
                    continue
                result = post_to_facebook(page, content, image_path)
                if "id" in result or "post_id" in result:
                    update_post(row_id, "posted", result.get("post_id") or result.get("id"))
                else:
                    update_post(row_id, "error", str(result))
        except Exception as e:
            print("Scheduler error:", e)
        time.sleep(30)

def read_batch_file(file_obj):
    filename = file_obj.filename.lower()
    if filename.endswith(".csv"):
        text = file_obj.read().decode("utf-8-sig", errors="ignore").splitlines()
        return list(csv.DictReader(text))
    if filename.endswith(".xlsx"):
        if openpyxl is None:
            raise RuntimeError("Chưa cài openpyxl. Chạy: pip install openpyxl")
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        file_obj.save(tmp.name)
        wb = openpyxl.load_workbook(tmp.name)
        ws = wb.active
        headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
        rows = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            item = {}
            for i, h in enumerate(headers):
                item[h] = r[i] if i < len(r) else ""
            rows.append(item)
        return rows
    raise RuntimeError("Chỉ hỗ trợ file .xlsx hoặc .csv")

def export_csv():
    rows = get_history(10000)
    path = os.path.join(REPORT_DIR, "report_posts.csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["id","page_name","content","status","post_id","schedule_time","image_path","campaign","score","created_at"])
        writer.writerows(rows)
    return path


def add_crm(name, phone, zalo, source, note):
    if not name and not phone and not zalo:
        return
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    INSERT INTO crm(name,phone,zalo,source,note,created_at)
    VALUES(?,?,?,?,?,?)
    """, (name, phone, zalo, source, note, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()

def get_crm(limit=30):
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("SELECT id,name,phone,zalo,source,note,created_at FROM crm ORDER BY id DESC LIMIT ?", (limit,))
    rows = c.fetchall()
    conn.close()
    return rows


def save_token_check(page_name, page_id, status, detail):
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    INSERT INTO token_checks(page_name,page_id,status,detail,checked_at)
    VALUES(?,?,?,?,?)
    """, (page_name, page_id, status, detail, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()

def get_latest_token_checks(limit=30):
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    SELECT page_name,page_id,status,detail,checked_at
    FROM token_checks
    ORDER BY id DESC
    LIMIT ?
    """, (limit,))
    rows = c.fetchall()
    conn.close()
    return rows

def check_single_page_token(page):
    page_name = page.get("name", "No name")
    page_id = str(page.get("id", ""))
    token = page.get("token", "")

    if not page_id or not token:
        return {"page_name": page_name, "page_id": page_id, "status": "LỖI", "detail": "Thiếu Page ID hoặc Page Token trong file .env."}

    try:
        url = f"https://graph.facebook.com/v23.0/{page_id}"
        params = {"fields": "id,name", "access_token": token}
        res = requests.get(url, params=params, timeout=20)
        data = res.json()

        if res.status_code < 400 and data.get("id"):
            return {"page_name": data.get("name", page_name), "page_id": data.get("id", page_id), "status": "SỐNG", "detail": "Token hoạt động bình thường."}

        err = data.get("error", {})
        code = err.get("code", "")
        subcode = err.get("error_subcode", "")
        message = err.get("message", str(data))
        if str(code) == "190":
            detail = f"Token giới hạn hoặc phiên đã giới hạn. OAuth code 190. Subcode: {subcode}. Nội dung: {message}"
        else:
            detail = f"Lỗi Facebook API. Code: {code}. Subcode: {subcode}. Nội dung: {message}"

        return {"page_name": page_name, "page_id": page_id, "status": "CHẾT", "detail": detail}

    except Exception as e:
        return {"page_name": page_name, "page_id": page_id, "status": "LỖI", "detail": str(e)}

def check_all_page_tokens():
    results = []
    for page in get_fanpages():
        item = check_single_page_token(page)
        save_token_check(item["page_name"], item["page_id"], item["status"], item["detail"])
        results.append(item)
    return results

def add_page_cluster(name, page_names, note):
    if not name:
        return
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    INSERT OR REPLACE INTO page_clusters(name,page_names,note,created_at)
    VALUES(?,?,?,?)
    """, (name, page_names, note, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()

def get_page_clusters():
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("SELECT id,name,page_names,note,created_at FROM page_clusters ORDER BY id DESC")
    rows = c.fetchall()
    conn.close()
    return rows

def get_analytics_summary():
    conn = sqlite3.connect(DB)
    c = conn.cursor()

    def one(query, params=()):
        try:
            c.execute(query, params)
            row = c.fetchone()
            return row[0] if row and row[0] is not None else 0
        except Exception:
            return 0

    def many(query, params=()):
        try:
            c.execute(query, params)
            return c.fetchall()
        except Exception:
            return []

    total_posts = one("SELECT COUNT(*) FROM posts")
    posted = one("SELECT COUNT(*) FROM posts WHERE status='posted'")
    scheduled = one("SELECT COUNT(*) FROM posts WHERE status='scheduled'")
    errors = one("SELECT COUNT(*) FROM posts WHERE status='error'")
    campaigns = one("SELECT COUNT(*) FROM campaigns")
    crm_total = one("SELECT COUNT(*) FROM crm")
    pipeline_total = one("SELECT COUNT(*) FROM lead_pipeline")
    groups_total = one("SELECT COUNT(*) FROM fb_groups")
    group_scheduled = one("SELECT COUNT(*) FROM group_schedules WHERE status='scheduled'")
    comments_total = one("SELECT COUNT(*) FROM comment_leads")
    messenger_total = one("SELECT COUNT(*) FROM messenger_scripts")

    by_page = many("""
        SELECT COALESCE(NULLIF(page_name,''),'Chưa đặt tên') AS name, COUNT(*) AS total
        FROM posts
        GROUP BY COALESCE(NULLIF(page_name,''),'Chưa đặt tên')
        ORDER BY total DESC
        LIMIT 10
    """)

    by_campaign = many("""
        SELECT COALESCE(NULLIF(campaign,''),'Chưa phân loại') AS name, COUNT(*) AS total
        FROM posts
        GROUP BY COALESCE(NULLIF(campaign,''),'Chưa phân loại')
        ORDER BY total DESC
        LIMIT 10
    """)

    by_day = many("""
        SELECT substr(created_at,1,10) AS day, COUNT(*) AS total
        FROM posts
        WHERE created_at IS NOT NULL AND created_at != ''
        GROUP BY substr(created_at,1,10)
        ORDER BY substr(created_at,1,10) DESC
        LIMIT 7
    """)

    by_status = many("""
        SELECT COALESCE(NULLIF(status,''),'unknown') AS status, COUNT(*) AS total
        FROM posts
        GROUP BY COALESCE(NULLIF(status,''),'unknown')
        ORDER BY total DESC
    """)

    group_by_niche = many("""
        SELECT COALESCE(NULLIF(niche,''),'Chưa phân loại') AS niche, COUNT(*) AS total
        FROM fb_groups
        GROUP BY COALESCE(NULLIF(niche,''),'Chưa phân loại')
        ORDER BY total DESC
        LIMIT 10
    """)

    crm_by_source = many("""
        SELECT COALESCE(NULLIF(source,''),'Chưa rõ nguồn') AS source, COUNT(*) AS total
        FROM crm
        GROUP BY COALESCE(NULLIF(source,''),'Chưa rõ nguồn')
        ORDER BY total DESC
        LIMIT 10
    """)

    crm_by_stage = many("""
        SELECT COALESCE(NULLIF(stage,''),'Khách mới') AS stage, COUNT(*) AS total
        FROM lead_pipeline
        GROUP BY COALESCE(NULLIF(stage,''),'Khách mới')
        ORDER BY total DESC
        LIMIT 10
    """)

    total_value = one("SELECT COALESCE(SUM(value),0) FROM lead_pipeline")
    avg_value = one("SELECT COALESCE(AVG(value),0) FROM lead_pipeline WHERE value > 0")
    conn.close()

    conversion_rate = round((posted / total_posts) * 100, 1) if total_posts else 0
    error_rate = round((errors / total_posts) * 100, 1) if total_posts else 0

    return {
        "summary": {
            "total_posts": total_posts,
            "posted": posted,
            "scheduled": scheduled,
            "errors": errors,
            "campaigns": campaigns,
            "crm_total": crm_total,
            "pipeline_total": pipeline_total,
            "groups_total": groups_total,
            "group_scheduled": group_scheduled,
            "comments_total": comments_total,
            "messenger_total": messenger_total,
            "total_value": int(total_value or 0),
            "avg_value": int(avg_value or 0),
            "conversion_rate": conversion_rate,
            "error_rate": error_rate
        },
        "by_page": by_page,
        "by_campaign": by_campaign,
        "by_day": by_day,
        "by_status": by_status,
        "group_by_niche": group_by_niche,
        "crm_by_source": crm_by_source,
        "crm_by_stage": crm_by_stage
    }

def generate_many_contents(idea, count, style="bán hàng tự nhiên", length="80"):
    count = max(1, min(int(count), 50))
    if not client:
        base = [
            f"{idea} - Giải pháp phù hợp cho khách hàng đang cần tối ưu công việc. Inbox để được tư vấn.",
            f"Bạn đang cần {idea}? Hãy chọn giải pháp ổn định, dễ dùng và được hỗ trợ nhanh.",
            f"{idea} giúp công việc online thuận tiện hơn. Nhắn tin để nhận tư vấn chi tiết."
        ]
        return "\\n\\n".join(base[i % len(base)] for i in range(count))

    prompt = f"""
Viết {count} content Facebook bằng tiếng Việt.
Chủ đề: {idea}
Phong cách: {style}
Mỗi bài dưới {length} từ.

Yêu cầu:
- Mỗi content khác nhau
- Có CTA inbox/liên hệ
- Có emoji vừa phải
- Không cam kết quá đà
- Tách mỗi content bằng dòng trống
"""
    return safe_ai_generate(prompt, "\n\n".join(base[i % len(base)] for i in range(count)) if 'base' in locals() else "")

def smart_schedule_times(start_time, count, gap_minutes):
    times = []
    try:
        base = datetime.datetime.strptime(start_time.replace("T", " "), "%Y-%m-%d %H:%M")
    except Exception:
        base = datetime.datetime.now() + datetime.timedelta(minutes=10)
    gap = int(gap_minutes or 60)
    for i in range(count):
        times.append((base + datetime.timedelta(minutes=i * gap)).strftime("%Y-%m-%d %H:%M"))
    return times


def ai_spin_content(content):
    if not content:
        return ""
    if client:
        prompt = f"""
Viết lại content sau khác khoảng 70-80% nhưng giữ ý chính, giọng tự nhiên, không cam kết quá đà.
Thêm CTA riêng và hashtag riêng.
Content gốc:
{content}
"""
        try:
            return safe_ai_generate(prompt, content)
        except Exception as e:
            return friendly_ai_error(e)
    endings = [
        "Inbox để được tư vấn chi tiết và nhận báo giá phù hợp.",
        "Nhắn tin ngay để được hỗ trợ nhanh nhất hôm nay.",
        "Liên hệ để nhận thông tin mới và ưu đãi phù hợp.",
        "Để lại tin nhắn, đội ngũ tư vấn sẽ hỗ trợ bạn.",
    ]
    return content.strip().replace("Inbox", "Nhắn tin").replace("Liên hệ", "Kết nối") + "\n\n" + random.choice(endings) + "\n#Marketing #KinhDoanhOnline"

def auto_cta_hashtag(content, industry="marketing"):
    text = content.lower()
    cta = "Inbox để được tư vấn nhanh."
    if "zalo" in text:
        cta = "Kết nối Zalo để được hỗ trợ chi tiết."
    elif "đặt lịch" in text or "spa" in text or "nha khoa" in text:
        cta = "Đặt lịch tư vấn ngay hôm nay."
    elif "proxy" in text:
        cta = "Nhắn tin để nhận bảng giá proxy phù hợp."
    hashtags = f"#{industry.replace('_','')} #Marketing #KinhDoanhOnline"
    return cta, hashtags

def choose_best_media_for_content(content, media_paths, used_indexes):
    if not media_paths:
        return ""
    # Bản local: ưu tiên không trùng ảnh; nếu hết ảnh thì quay vòng.
    for i, path in enumerate(media_paths):
        if i not in used_indexes:
            used_indexes.add(i)
            return path
    idx = len(used_indexes) % len(media_paths)
    return media_paths[idx]

def backup_database():
    if not os.path.exists(DB):
        return ""
    name = "backup_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + DB
    path = os.path.join(BACKUP_DIR, name)
    shutil.copy2(DB, path)
    return path

def export_pdf_report():
    # PDF tối giản dạng text để tránh phụ thuộc thư viện ngoài.
    rows = get_history(500)
    path = os.path.join(REPORT_DIR, "report_posts.pdf")
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
        c = canvas.Canvas(path, pagesize=A4)
        width, height = A4
        y = height - 40
        c.setFont("Helvetica-Bold", 14)
        c.drawString(40, y, "Marketing Automation Pro V9 - Report")
        y -= 30
        c.setFont("Helvetica", 9)
        for r in rows[:80]:
            line = f"ID {r[0]} | Page: {r[1]} | Status: {r[3]} | Campaign: {r[7]} | Time: {r[9]}"
            c.drawString(40, y, line[:110])
            y -= 14
            if y < 40:
                c.showPage()
                y = height - 40
                c.setFont("Helvetica", 9)
        c.save()
        return path
    except Exception:
        txt_path = os.path.join(REPORT_DIR, "report_posts_pdf_fallback.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("Marketing Automation Pro V9 Report\n\n")
            for r in rows:
                f.write(f"ID {r[0]} | Page: {r[1]} | Status: {r[3]} | Campaign: {r[7]} | Time: {r[9]}\n")
        return txt_path

def premium_visible_items(items, free_limit=10):
    return items[:free_limit], max(0, len(items) - free_limit)


def get_trial_identity():
    try:
        raw = request.cookies.get("mkt_trial_user") or request.remote_addr or "local_user"
    except Exception:
        raw = "local_user"
    return str(raw).replace(" ", "_")[:80]


def get_free_status(username=None):
    """Trial thật 3 ngày.
    Gói dùng thử chỉ mở: Quản lý Fanpage, Quản lý Group, AI Comment.
    Các tính năng còn lại sẽ chuyển sang popup Premium.
    """
    premium_access = get_effective_access_status()
    if premium_access:
        return premium_access

    username = username or get_trial_identity()
    now = datetime.datetime.now()

    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS users_trial (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        package_name TEXT DEFAULT 'trial',
        trial_start_date TEXT,
        trial_end_date TEXT,
        status TEXT DEFAULT 'active',
        created_at TEXT
    )
    """)
    c.execute("SELECT username,package_name,trial_start_date,trial_end_date,status FROM users_trial WHERE username=?", (username,))
    row = c.fetchone()

    if not row:
        start = now
        end = now + datetime.timedelta(days=3)
        c.execute("""
        INSERT INTO users_trial(username,package_name,trial_start_date,trial_end_date,status,created_at)
        VALUES(?,?,?,?,?,?)
        """, (username, "trial", start.strftime("%Y-%m-%d %H:%M:%S"), end.strftime("%Y-%m-%d %H:%M:%S"), "active", now.strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        row = (username, "trial", start.strftime("%Y-%m-%d %H:%M:%S"), end.strftime("%Y-%m-%d %H:%M:%S"), "active")

    package_name = row[1] or "trial"
    trial_start = row[2] or now.strftime("%Y-%m-%d %H:%M:%S")
    trial_end = row[3] or (now + datetime.timedelta(days=3)).strftime("%Y-%m-%d %H:%M:%S")
    status = row[4] or "active"

    try:
        end_dt = datetime.datetime.strptime(trial_end, "%Y-%m-%d %H:%M:%S")
    except Exception:
        end_dt = now

    remaining = end_dt - now
    total_seconds = int(remaining.total_seconds())
    if package_name == "trial" and total_seconds <= 0:
        status = "expired"
        c.execute("UPDATE users_trial SET status=? WHERE username=?", ("expired", username))
        conn.commit()

    conn.close()

    days = max(0, total_seconds // 86400)
    hours = max(0, (total_seconds % 86400) // 3600)
    percent = max(0, min(100, int((total_seconds / (3 * 86400)) * 100))) if package_name == "trial" else 100

    allowed_features = ["Quản lý Fanpage", "Quản lý Group", "AI Comment"]
    locked_features = ["AI Messenger", "CRM Kanban", "AI Marketing Director", "AI Video", "AI Image", "AI Kinh Doanh", "AI Giọng Nói", "AI Livestream"]

    return {
        "package_name": package_name,
        "status": status,
        "trial_start": trial_start,
        "trial_end": trial_end,
        "free_end": trial_end,
        "days": days,
        "hours": hours,
        "percent": percent,
        "is_trial": package_name == "trial",
        "is_expired": status == "expired",
        "label": "🎁 Dùng thử miễn phí 3 ngày" if status != "expired" else "🔒 Dùng thử đã hết hạn",
        "note": "Dùng thử chỉ mở: Quản lý Fanpage • Quản lý Group • AI Comment",
        "allowed_features": allowed_features,
        "locked_features": locked_features
    }


PREMIUM_PACKAGES = {
    "monthly": {"name": "Gói 1 tháng", "price": "159.000đ", "amount": 159000},
    "quarterly": {"name": "Gói 3 tháng", "price": "359.000đ", "amount": 359000},
    "halfyear": {"name": "Gói 6 tháng", "price": "559.000đ", "amount": 559000},
    "yearly": {"name": "Gói 1 năm", "price": "859.000đ", "amount": 859000},
    "sellerpro": {"name": "Gói nhà bán hàng chuyên nghiệp", "price": "1.959.000đ", "amount": 1959000}
}


def plan_required_message(feature_name, plans):
    return (
        f"Tính năng nâng cao: {feature_name}\n\n"
        f"Công cụ này thuộc nhóm Premium.\n"
        f"Gói đề xuất: {plans}\n\n"
        "Anh/chị có thể nâng cấp để mở khóa đầy đủ tính năng, hạn mức cao hơn và hỗ trợ ưu tiên."
    )

def token_manager_report():
    reports = []
    for p in get_fanpages():
        token = p.get("token", "")
        status = "Có token" if token else "Thiếu token"
        source = "SQLite" if p.get("source") == "database" else "Environment"
        reports.append(f"{p.get('name','No name')} | {p.get('id','No ID')} | {status} | Nguồn: {source}")
    if not reports:
        reports.append("Chưa có Fanpage. Hãy thêm trực tiếp trong Token Center.")
    return "\n".join(reports)

def ai_planner_v6(industry, goal, days):
    if not client:
        base = CONTENT_LIBRARY.get(industry, list(CONTENT_LIBRARY.values())[0])
        return "\\n".join([f"Ngày {i}: {random.choice(base)}\\nCTA: Inbox để được tư vấn.\\nHashtag: #{industry} #Marketing" for i in range(1, int(days)+1)])
    prompt = f"""
Tạo kế hoạch marketing Facebook {days} ngày bằng tiếng Việt.
Ngành: {industry}
Mục tiêu: {goal}

Mỗi ngày gồm:
- Chủ đề bài đăng
- Nội dung ngắn
- CTA
- Hashtag
- Gợi ý khung giờ đăng
Trình bày theo Ngày 1 đến Ngày {days}.
"""
    return safe_ai_generate(prompt, "")




def add_fb_group(group_name, group_id, niche, note):
    if not group_name and not group_id:
        return
    conn = db(); c = conn.cursor()
    c.execute("""
    INSERT INTO fb_groups(group_name,group_id,niche,note,status,created_at)
    VALUES(?,?,?,?,?,?)
    """, (group_name, group_id, niche, note, 'active', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    conn.commit(); conn.close()

def get_fb_groups(limit=80):
    conn = db(); c = conn.cursor()
    c.execute("SELECT id,group_name,group_id,niche,note,status,created_at FROM fb_groups ORDER BY id DESC LIMIT ?", (limit,))
    rows = c.fetchall(); conn.close(); return rows

def add_group_schedule(group_name, group_id, content, schedule_time):
    if not content:
        return
    conn = db(); c = conn.cursor()
    c.execute("""
    INSERT INTO group_schedules(group_name,group_id,content,schedule_time,status,created_at)
    VALUES(?,?,?,?,?,?)
    """, (group_name, group_id, content, schedule_time, 'scheduled', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    conn.commit(); conn.close()

def get_group_schedules(limit=50):
    conn = db(); c = conn.cursor()
    c.execute("SELECT id,group_name,group_id,content,schedule_time,status,created_at FROM group_schedules ORDER BY id DESC LIMIT ?", (limit,))
    rows = c.fetchall(); conn.close(); return rows

def hide_phone_preview(text):
    import re
    return re.sub(r'(0|\+84)[0-9\s\.\-]{7,13}', '[ĐÃ ẨN SĐT]', text or '')

def add_comment_lead(customer_name, phone, comment_text, ai_reply, label):
    if not comment_text and not phone and not customer_name:
        return
    conn = db(); c = conn.cursor()
    c.execute("""
    INSERT INTO comment_leads(customer_name,phone,comment_text,ai_reply,label,crm_status,created_at)
    VALUES(?,?,?,?,?,?,?)
    """, (customer_name, phone, comment_text, ai_reply, label, 'new', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    conn.commit(); conn.close()

def get_comment_leads(limit=50):
    conn = db(); c = conn.cursor()
    c.execute("SELECT id,customer_name,phone,comment_text,ai_reply,label,crm_status,created_at FROM comment_leads ORDER BY id DESC LIMIT ?", (limit,))
    rows = c.fetchall(); conn.close(); return rows

def add_messenger_script(product, script_type, content):
    if not product and not content:
        return
    conn = db(); c = conn.cursor()
    c.execute("INSERT INTO messenger_scripts(product,script_type,content,created_at) VALUES(?,?,?,?)", (product, script_type, content, datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    conn.commit(); conn.close()

def get_messenger_scripts(limit=30):
    conn = db(); c = conn.cursor()
    c.execute("SELECT id,product,script_type,content,created_at FROM messenger_scripts ORDER BY id DESC LIMIT ?", (limit,))
    rows = c.fetchall(); conn.close(); return rows

def v5_seed_success_assets():
    conn = db(); c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM success_assets")
    if c.fetchone()[0] == 0:
        now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        samples = [
            ('Case Study','Shop Proxy tăng phản hồi inbox','Trước: đăng bài rời rạc, thiếu lịch. Sau: dùng AI tạo 30 ngày content, 10 ads, 5 kịch bản chốt sale và CRM theo dõi khách.'),
            ('Fanpage','Mẫu Fanpage dịch vụ chuyên nghiệp','Avatar rõ thương hiệu, cover nêu lợi ích, mô tả có Zalo/CTA, ghim bài giới thiệu dịch vụ và bảng giá.'),
            ('Content','Mẫu content thành công','Nêu nỗi đau, đưa giải pháp, bằng chứng mềm, CTA inbox và hashtag vừa đủ.'),
            ('Ads','Mẫu quảng cáo thành công','Hook mạnh, lợi ích rõ, tránh cam kết quá đà, kêu gọi inbox nhận tư vấn.'),
            ('Script','Mẫu kịch bản chốt sale','Chào hỏi - hỏi nhu cầu - tư vấn gói phù hợp - xử lý từ chối - chốt hành động tiếp theo.')
        ]
        c.executemany("INSERT INTO success_assets(asset_type,title,content,created_at) VALUES(?,?,?,?)", [(a,b,cnt,now) for a,b,cnt in samples])
        conn.commit()
    conn.close()

def get_success_assets(limit=20):
    v5_seed_success_assets()
    conn = db(); c = conn.cursor()
    c.execute("SELECT id,asset_type,title,content,created_at FROM success_assets ORDER BY id ASC LIMIT ?", (limit,))
    rows = c.fetchall(); conn.close(); return rows

def v3_ceo_summary():
    s = get_stats()
    token_checks = get_latest_token_checks(200)
    live_tokens = sum(1 for t in token_checks if str(t[2]).upper() == 'SỐNG')
    dead_tokens = sum(1 for t in token_checks if str(t[2]).upper() in ['CHẾT','LỖI'])
    today = datetime.datetime.now().strftime('%Y-%m-%d')
    conn = db(); c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM posts WHERE substr(created_at,1,10)=?", (today,))
    posts_today = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM crm WHERE substr(created_at,1,10)=?", (today,))
    leads_today = c.fetchone()[0]
    conn.close()
    return {
        'fanpages': len(get_fanpages()),
        'posts': s.get('total', 0),
        'posted': s.get('posted', 0),
        'scheduled': s.get('scheduled', 0),
        'crm': s.get('crm', 0),
        'campaigns': s.get('campaigns', 0),
        'posts_today': posts_today,
        'leads_today': leads_today,
        'revenue_month': '0đ',
        'premium': get_free_status().get('label','🎁 Dùng thử miễn phí'),
        'token_live': live_tokens,
        'token_dead': dead_tokens,
        'token_total': len(get_fanpages())
    }

def add_pipeline_lead(customer_name, phone, zalo, source, stage, value, note):
    if not customer_name and not phone and not zalo:
        return
    conn = db(); c = conn.cursor()
    c.execute("""
    INSERT INTO lead_pipeline(customer_name,phone,zalo,source,stage,value,note,created_at)
    VALUES(?,?,?,?,?,?,?,?)
    """, (customer_name, phone, zalo, source, stage or 'Khách mới', int(value or 0), note, datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    conn.commit(); conn.close()

def get_pipeline_leads(limit=120):
    conn = db(); c = conn.cursor()
    c.execute("SELECT id,customer_name,phone,zalo,source,stage,value,note,created_at FROM lead_pipeline ORDER BY id DESC LIMIT ?", (limit,))
    rows = c.fetchall(); conn.close()
    if not rows:
        # fallback dùng dữ liệu CRM cũ để Kanban vẫn có nội dung
        return [(r[0], r[1], r[2], r[3], r[4], 'Khách mới', 0, r[5], r[6]) for r in get_crm(30)]
    return rows

def add_customer_task(customer_name, task, due_date):
    if not task:
        return
    conn = db(); c = conn.cursor()
    c.execute("INSERT INTO customer_tasks(customer_name,task,due_date,status,created_at) VALUES(?,?,?,?,?)", (customer_name, task, due_date, 'pending', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    conn.commit(); conn.close()

def get_customer_tasks(limit=50):
    conn = db(); c = conn.cursor()
    c.execute("SELECT id,customer_name,task,due_date,status,created_at FROM customer_tasks ORDER BY id DESC LIMIT ?", (limit,))
    rows = c.fetchall(); conn.close(); return rows

def add_notification(title, detail, level='info'):
    if not title:
        return
    conn = db(); c = conn.cursor()
    c.execute("INSERT INTO notifications(title,detail,level,status,created_at) VALUES(?,?,?,?,?)", (title, detail, level, 'new', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    conn.commit(); conn.close()

def get_notifications(limit=30):
    conn = db(); c = conn.cursor()
    c.execute("SELECT id,title,detail,level,status,created_at FROM notifications ORDER BY id DESC LIMIT ?", (limit,))
    rows = c.fetchall(); conn.close(); return rows

def v3_ai_tool_prompt(tool, topic, extra=''):
    if tool == 'marketing_director':
        return f"""
Bạn là AI Marketing Director Premium cho chủ shop/doanh nghiệp nhỏ. Lập bộ kế hoạch bằng tiếng Việt cho: {topic}
Yêu cầu xuất đúng cấu trúc sau:
1. Phân tích nhanh sản phẩm/dịch vụ và lợi thế bán hàng
2. Tệp khách hàng mục tiêu chi tiết
3. 30 content Facebook/TikTok ngắn, dễ đăng ngay
4. 10 mẫu quảng cáo Facebook Ads kéo inbox/chuyển đổi
5. 10 caption viral ngắn
6. 5 kịch bản chốt sale/inbox xử lý từ chối
7. Kế hoạch marketing 30 ngày
8. Ngân sách quảng cáo đề xuất theo 3 mức: thấp, vừa, mạnh
9. Việc cần làm mỗi ngày để tăng tỷ lệ chốt đơn
Không cam kết quá đà, không dùng lời hứa chắc chắn ra đơn, không spam hashtag.
Thông tin thêm: {extra}
"""
    if tool == 'facebook_ads_ai':
        return f"""
Bạn là chuyên gia Facebook Ads. Phân tích và tạo bộ quảng cáo cho: {topic}
Gồm: chấm điểm nội dung, 5 mẫu ads inbox, 5 mẫu ads chuyển đổi, target khách hàng, CTA, gợi ý ngân sách, lỗi cần tránh.
Thông tin thêm: {extra}
"""
    if tool == 'competitor_scanner':
        return f"""
Phân tích Fanpage đối thủ dựa trên mô tả/link này: {topic}
Trả về: điểm mạnh, điểm yếu, phong cách nội dung, CTA, ý tưởng vượt đối thủ, 10 nội dung có thể triển khai.
Thông tin thêm: {extra}
"""
    if tool == 'sales_script':
        return f"""
Tạo bộ kịch bản bán hàng cho: {topic}
Gồm: kịch bản inbox, kịch bản chốt sale, xử lý từ chối, chăm sóc lại khách cũ, tin nhắn nhắc thanh toán, tin nhắn hậu mãi.
Thông tin thêm: {extra}
"""
    if tool == 'landing_page':
        return f"""
Tạo nội dung landing page cho dịch vụ/sản phẩm: {topic}
Gồm: tiêu đề, mô tả, ưu điểm, form thu khách, CTA, FAQ, lý do nên chọn, nội dung hero section.
Thông tin thêm: {extra}
"""
    if tool == 'group_content':
        return f"""
Viết bài đăng Group Facebook bằng tiếng Việt cho: {topic}
Yêu cầu: gần gũi, không quá quảng cáo, tạo thảo luận, có CTA mềm, dưới 120 từ, kèm 3 biến thể tiêu đề.
Thông tin thêm: {extra}
"""
    if tool == 'smart_engagement':
        return f"""
Bạn là chuyên gia tăng tương tác Facebook an toàn cho chủ shop/doanh nghiệp nhỏ.
Sản phẩm/dịch vụ hoặc bài viết cần tối ưu: {topic}
Thông tin thêm: {extra}

Tạo kế hoạch Tăng tương tác thông minh gồm:
1. Gợi ý 5 khung giờ đăng tốt
2. 10 caption kéo comment tự nhiên
3. 10 câu hỏi tạo thảo luận thật
4. 5 mẫu phản hồi comment để kéo khách vào inbox
5. Cách nhận diện bài ít tương tác và việc admin cần làm
6. Kịch bản chăm sóc lại khách đã comment
7. Lưu ý an toàn: không auto like hàng loạt, không dùng nhiều tài khoản, không tạo tương tác giả.
Trình bày rõ ràng, dễ copy, không cam kết quá đà.
"""
    if tool == 'comment_reply':
        return f"""
Bạn là Comment Manager AI. Hãy tạo phản hồi comment chuyên nghiệp cho khách hàng.
Comment/Nội dung khách: {topic}
Yêu cầu: trả lời ngắn, thân thiện, kéo khách vào inbox, không lộ số điện thoại, gắn nhãn khách nóng/ấm/lạnh và đề xuất chuyển CRM.
Thông tin thêm: {extra}
"""
    if tool == 'messenger_ai':
        return f"""
Tạo bộ kịch bản Messenger bán hàng cho: {topic}
Gồm: kịch bản inbox, kịch bản chốt sale, xử lý từ chối, chăm sóc khách cũ, tin nhắn nhắc thanh toán, tin nhắn hậu mãi.
Thông tin thêm: {extra}
"""
    if tool == 'prompt_premium':
        return f"""
Tạo 50 prompt Premium cho: {topic}
Chia nhóm: Facebook Ads, TikTok Ads, Seeding, Chốt sale, Content Viral.
"""
    return f"Tạo nội dung marketing chuyên nghiệp cho: {topic}. {extra}"

HTML = """
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{{ title }}</title>

<meta name="theme-color" content="#0F172A">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Mkt Automation Pro">
<link rel="manifest" href="/manifest.json">
<link rel="apple-touch-icon" href="/pwa-icon-192.png?v=gpt-mkt-pro-1">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=Manrope:wght@400;500;600;700;800&display=swap" rel="stylesheet">


<style>
:root{
  --blue:#2563EB;
  --blue2:#3B82F6;
  --purple:#7C3AED;
  --purple2:#A855F7;
  --deep:#1E1B4B;
  --navy:#0F172A;
  --bg:#F8FAFC;
  --soft:#EEF2FF;
  --card:#FFFFFF;
  --text:#111827;
  --muted:#64748B;
  --border:#E5E7EB;
}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{
  margin:0;
  font-family:'Manrope','Inter',Arial,sans-serif;
  background:
    radial-gradient(circle at top left,rgba(124,58,237,.18),transparent 32%),
    radial-gradient(circle at top right,rgba(37,99,235,.16),transparent 30%),
    linear-gradient(135deg,#FFFFFF,#EEF2FF 52%,#F8FAFC);
  color:var(--text);
}
.layout{
  display:grid;
  grid-template-columns:270px 1fr 330px;
  gap:20px;
  max-width:1560px;
  margin:auto;
  padding:20px;
}
.sidebar{
  position:sticky;
  top:20px;
  height:calc(100vh - 40px);
  padding:22px;
  background:linear-gradient(180deg,#111827,#1E1B4B);
  border:1px solid rgba(255,255,255,.12);
  border-radius:28px;
  box-shadow:0 20px 55px rgba(30,27,75,.25);
  overflow-y:auto;
  overflow-x:hidden;
  overscroll-behavior:contain;
  scrollbar-width:thin;
}
.logo{
  font-size:25px;
  font-weight:900;
  color:white;
  line-height:1.12;
}
.logo:after{
  content:"";
  display:block;
  width:86px;
  height:5px;
  margin-top:12px;
  border-radius:999px;
  background:linear-gradient(135deg,var(--blue),var(--purple2));
}
.subtitle{
  font-size:12px;
  color:#CBD5E1;
  margin:12px 0 18px;
}
.nav a{
  display:block;
  padding:13px 14px;
  border-radius:16px;
  color:#F8FAFC;
  text-decoration:none;
  margin:7px 0;
  background:rgba(255,255,255,.06);
  border:1px solid rgba(255,255,255,.08);
}
.nav a:hover{
  background:linear-gradient(135deg,var(--blue),var(--purple));
  color:white;
  transform:translateX(3px);
}
.main{min-width:0}
.panel,.rightbar{
  background:rgba(255,255,255,.94);
  backdrop-filter:blur(16px);
  border:1px solid rgba(226,232,240,.9);
  border-radius:28px;
  box-shadow:0 18px 45px rgba(30,41,59,.10);
}
.panel{
  padding:25px;
  margin-bottom:20px;
}
.rightbar{
  position:sticky;
  top:20px;
  height:calc(100vh - 40px);
  padding:20px;
  overflow:auto;
}
.badge{
  display:inline-block;
  background:linear-gradient(135deg,var(--blue),var(--purple));
  color:white;
  padding:10px 16px;
  border-radius:999px;
  font-weight:900;
  box-shadow:0 10px 26px rgba(124,58,237,.26);
}
h1{
  font-size:42px;
  color:transparent;
  background:linear-gradient(135deg,var(--blue),var(--purple));
  -webkit-background-clip:text;
  background-clip:text;
  margin:16px 0 8px;
  letter-spacing:-.6px;
}
h2{
  margin-top:0;
  color:var(--deep);
}
p{color:var(--muted)}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.grid4{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:14px}
.stat{
  background:#FFFFFF;
  border:1px solid var(--border);
  border-left:5px solid var(--purple);
  border-radius:22px;
  padding:18px;
  box-shadow:0 12px 28px rgba(124,58,237,.08);
}
.stat span{color:var(--muted);font-size:14px}
.stat b{
  font-size:27px;
  color:var(--purple);
}
textarea,select,input{
  width:100%;
  margin:8px 0;
  padding:15px;
  border:1px solid var(--border);
  background:#FFFFFF;
  color:var(--text);
  border-radius:16px;
  font-size:15px;
  outline:none;
}
textarea:focus,select:focus,input:focus{
  border-color:var(--purple);
  box-shadow:0 0 0 4px rgba(124,58,237,.12);
}
button{
  background:linear-gradient(135deg,var(--blue),var(--purple));
  color:white;
  border:none;
  padding:14px 20px;
  border-radius:16px;
  font-weight:900;
  cursor:pointer;
  margin:8px 6px 8px 0;
  box-shadow:0 12px 28px rgba(37,99,235,.20);
}
button:hover{
  transform:translateY(-1px);
  box-shadow:0 16px 32px rgba(124,58,237,.25);
}
.secondary{
  background:linear-gradient(135deg,#E0E7FF,#F3E8FF);
  color:#4C1D95;
  border:1px solid #DDD6FE;
  box-shadow:none;
}
.success{
  color:#047857;
  background:#ECFDF5;
  border:1px solid #A7F3D0;
  padding:12px;
  border-radius:14px;
}
.error{
  color:#B91C1C;
  background:#FEF2F2;
  border:1px solid #FCA5A5;
  padding:12px;
  border-radius:14px;
}
.page-list{
  display:grid;
  grid-template-columns:repeat(2,1fr);
  gap:10px;
  margin:10px 0;
}
.page-item{
  background:#F8FAFC;
  color:var(--text);
  border:1px solid var(--border);
  border-radius:18px;
  padding:13px;
}
.history{
  white-space:pre-wrap;
  background:#F8FAFC;
  color:var(--text);
  border-radius:18px;
  padding:15px;
  margin:10px 0;
  border:1px solid var(--border);
}
.small{font-size:13px;color:var(--muted)}
.preview{
  background:white;
  color:#111827;
  border-radius:22px;
  padding:18px;
  max-width:570px;
  border:1px solid var(--border);
  box-shadow:0 14px 32px rgba(30,41,59,.10);
}
.preview-head{display:flex;align-items:center;gap:10px}
.avatar{
  width:44px;
  height:44px;
  border-radius:50%;
  background:linear-gradient(135deg,var(--blue),var(--purple));
}
.preview-name{font-weight:900}
.preview-content{white-space:pre-wrap;margin-top:12px;line-height:1.48}
.library-grid{
  display:grid;
  grid-template-columns:repeat(2,1fr);
  gap:12px;
}
.template-card{
  background:#FFFFFF;
  color:var(--text);
  border:1px solid var(--border);
  border-top:4px solid var(--purple);
  border-radius:20px;
  padding:16px;
  box-shadow:0 12px 26px rgba(124,58,237,.08);
}
.rightbar h2{
  color:transparent;
  background:linear-gradient(135deg,var(--blue),var(--purple));
  -webkit-background-clip:text;
  background-clip:text;
}
.rightbar hr{border:none;border-top:1px solid var(--border);margin:18px 0}
.pricing-grid{
  display:grid;
  grid-template-columns:repeat(5,1fr);
  gap:14px;
}
.price-card{
  background:#FFFFFF;
  color:#111827;
  border:1px solid var(--border);
  border-radius:22px;
  padding:18px;
  position:relative;
  box-shadow:0 18px 38px rgba(124,58,237,.10);
}
.price-card h3{
  margin:0 0 8px;
  font-size:18px;
  color:var(--deep);
}
.price{
  font-size:22px;
  font-weight:800;
  line-height:1.2;
  color:transparent;
  background:linear-gradient(135deg,var(--blue),var(--purple));
  -webkit-background-clip:text;
  background-clip:text;
}
.price-card ul{
  padding-left:18px;
  line-height:1.7;
  color:#334155;
}
.popular{
  border:2px solid var(--purple);
  transform:translateY(-6px);
}
.ribbon{
  position:absolute;
  top:-13px;
  left:18px;
  background:linear-gradient(135deg,var(--blue),var(--purple));
  color:white;
  padding:6px 12px;
  border-radius:999px;
  font-weight:900;
  font-size:12px;
}
.premium-lock{
  background:linear-gradient(135deg,#EEF2FF,#F5F3FF);
  color:#1E1B4B;
  border-radius:22px;
  padding:20px;
  border:1px solid #DDD6FE;
}
.mobilebar{display:none}
@media(max-width:1100px){
  .layout{grid-template-columns:230px 1fr}
  .rightbar{display:none}
  .grid4{grid-template-columns:repeat(2,1fr)}
  .pricing-grid{grid-template-columns:repeat(2,1fr)}
}
@media(max-width:820px){
  .layout{display:block;padding:12px}
  .sidebar,.rightbar{display:none}
  .grid,.grid4,.page-list,.library-grid,.pricing-grid{grid-template-columns:1fr}
  h1{font-size:30px}
  .panel{padding:18px;border-radius:22px}
  button{width:100%;font-size:16px}
  textarea{min-height:160px}
  .popular{transform:none}
  .mobilebar{
    display:flex;
    position:fixed;
    bottom:0;
    left:0;
    right:0;
    background:#FFFFFF;
    border-top:1px solid var(--border);
    z-index:999;
    box-shadow:0 -8px 25px rgba(30,41,59,.10);
  }
  .mobilebar a{
    flex:1;
    text-align:center;
    color:#1E1B4B;
    text-decoration:none;
    padding:10px 4px;
    font-size:12px;
  }
  .mobilebar a:hover{color:var(--purple)}
  body{padding-bottom:68px}
}

.analytics-grid{
  display:grid;
  grid-template-columns:repeat(3,1fr);
  gap:14px;
}
.analytics-box{
  background:#F8FAFC;
  border:1px solid var(--border);
  border-radius:18px;
  padding:16px;
}
.analytics-row{
  display:flex;
  justify-content:space-between;
  gap:10px;
  padding:8px 0;
  border-bottom:1px solid #E5E7EB;
}
.premium-center{
  background:linear-gradient(135deg,#EEF2FF,#F5F3FF);
  border:1px solid #DDD6FE;
  border-radius:22px;
  padding:18px;
}
.token-live{color:#047857;font-weight:900}
.token-dead{color:#B91C1C;font-weight:900}
@media(max-width:820px){.analytics-grid{grid-template-columns:1fr}}


.v9-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}
.v9-card{background:#fff;border:1px solid var(--border);border-radius:20px;padding:16px;box-shadow:0 12px 28px rgba(124,58,237,.08)}
.lock-card{background:linear-gradient(135deg,#1E1B4B,#312E81);color:white;border-radius:22px;padding:20px;border:1px solid #A78BFA}
.modal-price{display:none;position:fixed;inset:0;background:rgba(15,23,42,.55);z-index:9999;align-items:center;justify-content:center;padding:20px}
.modal-inner{background:white;border-radius:28px;max-width:980px;width:100%;max-height:90vh;overflow:auto;padding:22px;color:#111827}
.close-btn{float:right;background:#F3F4F6;color:#111827;box-shadow:none}
@media(max-width:820px){.v9-grid{grid-template-columns:1fr}}


/* V10 Floating Support Bot + Live Trust Bar */
.live-trust-bar{
  position:fixed;
  top:16px;
  left:50%;
  transform:translateX(-50%);
  z-index:9998;
  background:rgba(255,255,255,.96);
  border:1px solid #DDD6FE;
  box-shadow:0 14px 35px rgba(30,41,59,.14);
  border-radius:999px;
  padding:10px 18px;
  display:flex;
  align-items:center;
  gap:10px;
  color:#1E1B4B;
  font-weight:800;
  backdrop-filter:blur(14px);
}
.live-dot{
  width:10px;
  height:10px;
  background:#22C55E;
  border-radius:50%;
  box-shadow:0 0 0 6px rgba(34,197,94,.14);
  animation:pulseLive 1.4s infinite;
}
@keyframes pulseLive{
  0%{transform:scale(1);opacity:1}
  70%{transform:scale(1.35);opacity:.65}
  100%{transform:scale(1);opacity:1}
}
.floating-bot{
  position:fixed;
  right:22px;
  bottom:22px;
  z-index:9999;
  font-family:'Manrope','Inter',Arial,sans-serif;
}
.bot-bubble{
  width:60px;
  height:60px;
  border-radius:50%;
  border:none;
  cursor:pointer;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  box-shadow:0 18px 45px rgba(124,58,237,.35);
  color:white;
  font-size:30px;
  display:flex;
  align-items:center;
  justify-content:center;
  position:relative;
  animation:botFloat 2.2s ease-in-out infinite;
}
.bot-bubble:after{
  content:"AI Online\A Phản hồi trong vài giây";
  white-space:pre;
  position:absolute;
  right:66px;
  bottom:8px;
  min-width:150px;
  background:rgba(15,23,42,.96);
  color:#E0F2FE;
  border:1px solid rgba(34,197,94,.35);
  border-radius:14px;
  padding:9px 11px;
  font-size:12px;
  line-height:1.35;
  text-align:left;
  opacity:0;
  pointer-events:none;
  transform:translateX(8px);
  transition:.18s ease;
  box-shadow:0 14px 34px rgba(15,23,42,.28);
}
.bot-bubble:hover:after{opacity:1;transform:translateX(0)}
@keyframes botFloat{
  0%,100%{transform:translateY(0)}
  50%{transform:translateY(-6px)}
}
.bot-status{
  position:absolute;
  right:3px;
  bottom:4px;
  width:10px;
  height:10px;
  border-radius:50%;
  background:#00ff88;
  border:2px solid white;
  box-shadow:0 0 8px #00ff88,0 0 14px rgba(0,255,136,.85);
  animation:aiOnlinePulse 1.6s infinite;
}
@keyframes aiOnlinePulse{
  0%{transform:scale(1);opacity:1}
  50%{transform:scale(1.35);opacity:.78}
  100%{transform:scale(1);opacity:1}
}
.bot-panel{
  display:none;
  position:absolute;
  right:0;
  bottom:82px;
  width:340px;
  max-width:calc(100vw - 28px);
  background:white;
  border:1px solid #DDD6FE;
  border-radius:24px;
  box-shadow:0 22px 55px rgba(30,41,59,.22);
  overflow:hidden;
}
.bot-head{
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  color:white;
  padding:16px;
  display:flex;
  justify-content:space-between;
  align-items:center;
}
.bot-title{
  font-weight:900;
}
.bot-online{
  font-size:12px;
  opacity:.95;
  display:flex;
  align-items:center;
  gap:6px;
}
.typing-dots{
  display:inline-flex;
  gap:3px;
  margin-left:4px;
}
.typing-dots span{
  width:5px;
  height:5px;
  border-radius:50%;
  background:white;
  display:block;
  animation:typingDot 1s infinite ease-in-out;
}
.typing-dots span:nth-child(2){animation-delay:.15s}
.typing-dots span:nth-child(3){animation-delay:.3s}
@keyframes typingDot{
  0%,80%,100%{transform:translateY(0);opacity:.55}
  40%{transform:translateY(-4px);opacity:1}
}
.bot-close{
  background:rgba(255,255,255,.18);
  border:none;
  color:white;
  box-shadow:none;
  width:32px;
  height:32px;
  border-radius:50%;
  padding:0;
  margin:0;
}
.bot-body{
  padding:16px;
  max-height:360px;
  overflow:auto;
  background:#F8FAFC;
}
.bot-msg{
  background:white;
  border:1px solid #E5E7EB;
  border-radius:18px;
  padding:12px;
  margin-bottom:10px;
  line-height:1.45;
  color:#111827;
}
.bot-msg.ai{
  border-left:4px solid #7C3AED;
}
.bot-actions{
  display:grid;
  grid-template-columns:1fr;
  gap:8px;
  padding:14px 16px 16px;
  background:white;
}
.bot-actions button,.bot-actions a{
  display:block;
  text-align:center;
  text-decoration:none;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  color:white;
  border-radius:14px;
  padding:12px;
  font-weight:900;
  border:none;
  box-shadow:0 10px 22px rgba(37,99,235,.18);
}
.bot-actions .light{
  background:#EEF2FF;
  color:#4C1D95;
  border:1px solid #DDD6FE;
  box-shadow:none;
}
.bot-input{
  display:flex;
  gap:8px;
  padding:0 16px 16px;
  background:white;
}
.bot-input input{
  flex:1;
  margin:0;
}
.bot-input button{
  margin:0;
  padding:12px 14px;
}
@media(max-width:820px){
  .live-trust-bar{
    top:8px;
    width:calc(100vw - 24px);
    justify-content:center;
    font-size:12px;
    padding:9px 10px;
  }
  .floating-bot{right:14px;bottom:78px}
  .bot-panel{width:320px}
}


/* V10 SaaS Center Layout */
.top-hero{
  padding:34px;
  border-radius:32px;
  background:
    radial-gradient(circle at top left,rgba(37,99,235,.18),transparent 34%),
    radial-gradient(circle at top right,rgba(124,58,237,.18),transparent 32%),
    rgba(255,255,255,.96);
  border:1px solid #DDD6FE;
  box-shadow:0 22px 55px rgba(30,41,59,.12);
  margin-bottom:20px;
}
.hero-line{
  display:inline-flex;
  align-items:center;
  gap:8px;
  padding:10px 16px;
  background:#EEF2FF;
  border:1px solid #DDD6FE;
  color:#4C1D95;
  border-radius:999px;
  font-weight:900;
}
.hero-actions{
  display:flex;
  flex-wrap:wrap;
  gap:10px;
  margin-top:18px;
}
.module-hub{
  display:grid;
  grid-template-columns:repeat(4,1fr);
  gap:16px;
  margin-top:18px;
}
.module-card{
  background:#FFFFFF;
  border:1px solid #E5E7EB;
  border-radius:24px;
  padding:20px;
  min-height:158px;
  cursor:pointer;
  box-shadow:0 14px 32px rgba(30,41,59,.08);
  transition:.18s ease;
  position:relative;
  overflow:hidden;
}
.module-card:before{
  content:"";
  position:absolute;
  inset:0;
  background:linear-gradient(135deg,rgba(37,99,235,.08),rgba(124,58,237,.08));
  opacity:0;
  transition:.18s ease;
}
.module-card:hover{
  transform:translateY(-5px);
  border-color:#A78BFA;
  box-shadow:0 20px 45px rgba(124,58,237,.16);
}
.module-card:hover:before{opacity:1}
.module-card .icon{
  font-size:32px;
  margin-bottom:10px;
  position:relative;
}
.module-card h3{
  margin:0 0 8px;
  color:#1E1B4B;
  position:relative;
}
.module-card p{
  margin:0;
  color:#64748B;
  font-size:13px;
  line-height:1.45;
  position:relative;
}
.module-pill{
  display:inline-block;
  margin-top:12px;
  font-size:12px;
  font-weight:900;
  color:#4C1D95;
  background:#F5F3FF;
  border:1px solid #DDD6FE;
  border-radius:999px;
  padding:6px 10px;
  position:relative;
}
.module-section{
  display:none;
}
.module-section.active-module{
  display:block;
}
.section-open-note{
  margin-bottom:14px;
  background:#EEF2FF;
  border:1px solid #DDD6FE;
  border-radius:18px;
  padding:12px 14px;
  color:#4C1D95;
  font-weight:800;
}
.price-bottom-title{
  text-align:center;
  margin:34px 0 18px;
}
.price-bottom-title h2{
  font-size:34px;
  margin-bottom:8px;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  color:transparent;
  -webkit-background-clip:text;
  background-clip:text;
}
@media(max-width:1200px){
  .module-hub{grid-template-columns:repeat(3,1fr)}
}
@media(max-width:820px){
  .module-hub{grid-template-columns:repeat(2,1fr)}
  .top-hero{padding:22px}
  .module-card{min-height:140px;padding:16px}
}
@media(max-width:520px){
  .module-hub{grid-template-columns:1fr}
}


#ai-output-box{
  border:2px solid #DDD6FE;
  box-shadow:0 22px 50px rgba(124,58,237,.14);
}
#aiGeneratedContent{
  font-size:15px;
  line-height:1.65;
  max-height:520px;
  overflow:auto;
  background:#FFFFFF;
}
.pricing-visible{
  display:block !important;
}


/* Premium Pricing Pro */
.premium-pricing-pro{
  background:
    radial-gradient(circle at top left,rgba(37,99,235,.18),transparent 32%),
    radial-gradient(circle at top right,rgba(124,58,237,.18),transparent 30%),
    linear-gradient(135deg,#FFFFFF,#F8FAFC);
  border:1px solid #DDD6FE;
  border-radius:32px;
  padding:30px;
  box-shadow:0 24px 60px rgba(30,41,59,.12);
}
.premium-title{
  text-align:center;
  margin-bottom:24px;
}
.premium-title .mini{
  display:inline-block;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  color:white;
  padding:9px 16px;
  border-radius:999px;
  font-weight:900;
  margin-bottom:10px;
}
.premium-title h2{
  font-size:38px;
  margin:6px 0;
  color:transparent;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  -webkit-background-clip:text;
  background-clip:text;
}
.premium-title p{
  max-width:760px;
  margin:8px auto 0;
}
.premium-grid-pro{
  display:grid;
  grid-template-columns:repeat(5,1fr);
  gap:16px;
  align-items:stretch;
}
.premium-plan{
  background:white;
  border:1px solid #E5E7EB;
  border-radius:26px;
  padding:20px;
  position:relative;
  box-shadow:0 18px 42px rgba(30,41,59,.10);
  transition:.18s ease;
  display:flex;
  flex-direction:column;
}
.premium-plan:hover{
  transform:translateY(-6px);
  border-color:#A78BFA;
  box-shadow:0 24px 55px rgba(124,58,237,.18);
}
.premium-plan.featured{
  border:2px solid #7C3AED;
  background:linear-gradient(180deg,#FFFFFF,#F5F3FF);
}
.plan-ribbon{
  position:absolute;
  top:-14px;
  left:18px;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  color:white;
  padding:7px 13px;
  border-radius:999px;
  font-size:12px;
  font-weight:900;
}
.plan-name{
  font-size:18px;
  font-weight:900;
  color:#1E1B4B;
  margin-top:8px;
}
.plan-price{
  font-size:24px;
  font-weight:900;
  margin:8px 0 4px;
  color:transparent;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  -webkit-background-clip:text;
  background-clip:text;
}
.plan-desc{
  font-size:13px;
  color:#64748B;
  min-height:42px;
}
.benefit-title{
  font-weight:900;
  color:#1E1B4B;
  margin:12px 0 8px;
}
.benefit-list{
  list-style:none;
  padding:0;
  margin:0;
  font-size:13px;
  line-height:1.55;
}
.benefit-list li{
  padding:5px 0;
  border-bottom:1px dashed #EEF2FF;
}
.benefit-list .open:before{content:"✓ ";color:#047857;font-weight:900}
.benefit-list .lock:before{content:"🔒 ";color:#B45309;font-weight:900}
.plan-button{
  margin-top:auto;
  width:100%;
}
.premium-note-box{
  margin-top:20px;
  background:#EEF2FF;
  border:1px solid #DDD6FE;
  color:#4C1D95;
  border-radius:22px;
  padding:16px;
  line-height:1.55;
}
.payment-modal{
  display:none;
  position:fixed;
  inset:0;
  background:rgba(15,23,42,.65);
  z-index:10000;
  align-items:center;
  justify-content:center;
  padding:18px;
}
.payment-inner{
  width:min(980px,96vw);
  max-height:92vh;
  overflow:auto;
  background:white;
  border-radius:32px;
  box-shadow:0 30px 85px rgba(15,23,42,.38);
  border:1px solid #DDD6FE;
  padding:0;
}
.payment-head{
  background:
    radial-gradient(circle at top left,rgba(255,255,255,.25),transparent 35%),
    linear-gradient(135deg,#2563EB,#7C3AED);
  color:white;
  padding:24px;
  display:flex;
  justify-content:space-between;
  gap:14px;
  align-items:flex-start;
}
.payment-head h2{
  color:white;
  margin:0 0 6px;
}
.payment-close{
  background:rgba(255,255,255,.18);
  color:white;
  box-shadow:none;
  border:1px solid rgba(255,255,255,.25);
  width:auto;
  margin:0;
}
.payment-body{
  display:grid;
  grid-template-columns:330px 1fr;
  gap:22px;
  padding:24px;
}
.qr-card{
  background:#F8FAFC;
  border:1px solid #E5E7EB;
  border-radius:26px;
  padding:18px;
  text-align:center;
}
.qr-card img{
  width:250px;
  max-width:100%;
  border-radius:18px;
  border:1px solid #E5E7EB;
  background:white;
  padding:8px;
}
.bank-info{
  margin-top:12px;
  text-align:left;
  background:white;
  border:1px solid #E5E7EB;
  border-radius:18px;
  padding:14px;
  line-height:1.7;
}
.payment-detail{
  background:white;
  border:1px solid #E5E7EB;
  border-radius:26px;
  padding:18px;
}
.payment-detail h3{
  margin-top:0;
  color:#1E1B4B;
}
.pay-amount{
  font-size:34px;
  font-weight:900;
  color:transparent;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  -webkit-background-clip:text;
  background-clip:text;
}
.payment-benefits{
  display:grid;
  grid-template-columns:1fr 1fr;
  gap:10px;
  margin:12px 0;
}
.payment-benefits div{
  background:#F8FAFC;
  border:1px solid #E5E7EB;
  border-radius:14px;
  padding:10px;
  font-size:13px;
}
.payment-alert{
  margin-top:14px;
  background:#FFF7ED;
  border:1px solid #FDBA74;
  color:#9A3412;
  border-radius:18px;
  padding:14px;
  line-height:1.55;
}
.payment-actions{
  display:flex;
  gap:10px;
  flex-wrap:wrap;
  margin-top:14px;
}
.payment-actions a{
  display:inline-block;
  text-decoration:none;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  color:white;
  padding:13px 16px;
  border-radius:14px;
  font-weight:900;
}
.payment-actions .light{
  background:#EEF2FF;
  color:#4C1D95;
  border:1px solid #DDD6FE;
}
.permission-db{
  margin-top:20px;
  background:#0F172A;
  color:#E5E7EB;
  border-radius:22px;
  padding:18px;
  overflow:auto;
  font-size:13px;
}
.permission-db code{
  color:#A7F3D0;
  white-space:pre-wrap;
}
@media(max-width:1200px){
  .premium-grid-pro{grid-template-columns:repeat(2,1fr)}
}
@media(max-width:820px){
  .premium-pricing-pro{padding:18px}
  .premium-grid-pro{grid-template-columns:1fr}
  .payment-body{grid-template-columns:1fr}
  .payment-benefits{grid-template-columns:1fr}
}


/* Free + Premium Stable */
.free-status-card{
  background:linear-gradient(135deg,#EEF2FF,#F5F3FF);
  border:1px solid #DDD6FE;
  color:#1E1B4B;
  border-radius:22px;
  padding:16px;
  margin:14px 0;
}
.free-progress{
  width:100%;
  height:12px;
  background:#E5E7EB;
  border-radius:999px;
  overflow:hidden;
  margin:10px 0;
}
.free-progress span{
  display:block;
  height:100%;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  border-radius:999px;
}
.free-expired{
  background:#FEF2F2;
  border-color:#FCA5A5;
  color:#991B1B;
}
.premium-story{
  margin:20px 0;
  background:linear-gradient(135deg,#0F172A,#1E1B4B);
  color:white;
  border-radius:28px;
  padding:24px;
  box-shadow:0 22px 55px rgba(30,27,75,.20);
}
.premium-story h3{
  margin-top:0;
  color:white;
  font-size:26px;
}
.premium-story p{
  color:#E0E7FF;
  line-height:1.65;
}
.why-premium{
  margin-top:22px;
  display:grid;
  grid-template-columns:repeat(2,1fr);
  gap:12px;
}
.why-item{
  background:white;
  border:1px solid #E5E7EB;
  border-radius:18px;
  padding:14px;
  color:#1E1B4B;
  box-shadow:0 12px 28px rgba(30,41,59,.06);
  font-weight:800;
}
.locked-demo{
  background:#FFF7ED;
  border:1px solid #FDBA74;
  color:#9A3412;
  border-radius:18px;
  padding:14px;
  margin-top:12px;
}
.lock-modal{
  display:none;
  position:fixed;
  inset:0;
  background:rgba(15,23,42,.65);
  z-index:10001;
  align-items:center;
  justify-content:center;
  padding:18px;
}
.lock-inner{
  background:white;
  border-radius:28px;
  max-width:560px;
  width:100%;
  padding:24px;
  box-shadow:0 30px 80px rgba(15,23,42,.35);
  border:1px solid #DDD6FE;
}
.lock-inner h2{
  margin-top:0;
  color:#1E1B4B;
}
@media(max-width:820px){
  .why-premium{grid-template-columns:1fr}
}


/* ===== GPT MINI MENU UPGRADE - PROFESSIONAL SAAS ===== */
.menu-upgrade-note{
  margin-top:14px;
  padding:14px;
  border-radius:18px;
  background:linear-gradient(135deg,rgba(37,99,235,.12),rgba(124,58,237,.12));
  border:1px solid rgba(167,139,250,.45);
  color:#E0E7FF;
  font-size:12px;
  line-height:1.55;
}
.menu-mini-group{
  margin:10px 0;
  border-radius:18px;
  overflow:hidden;
  border:1px solid rgba(255,255,255,.08);
  background:rgba(255,255,255,.045);
}
.menu-mini-title{
  padding:12px 14px;
  color:#FFFFFF;
  font-weight:900;
  cursor:pointer;
  display:flex;
  align-items:center;
  justify-content:space-between;
}
.menu-mini-title small{
  color:#CBD5E1;
  font-weight:700;
}
.menu-mini-title:hover{
  background:linear-gradient(135deg,rgba(37,99,235,.55),rgba(124,58,237,.55));
}
.menu-mini-items{
  display:none;
  padding:6px 8px 10px;
}
.menu-mini-group.open .menu-mini-items{
  display:block;
}
.menu-mini-item{
  display:block;
  padding:10px 12px;
  margin:5px 0;
  border-radius:14px;
  color:#E5E7EB;
  text-decoration:none;
  font-size:13px;
  background:rgba(255,255,255,.05);
  border:1px solid rgba(255,255,255,.06);
}
.menu-mini-item:hover{
  color:white;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  transform:translateX(3px);
}
.clean-hero-pro{
  padding:34px;
  border-radius:32px;
  background:
    radial-gradient(circle at top left,rgba(37,99,235,.20),transparent 34%),
    radial-gradient(circle at top right,rgba(124,58,237,.18),transparent 32%),
    rgba(255,255,255,.97);
  border:1px solid #DDD6FE;
  box-shadow:0 22px 55px rgba(30,41,59,.12);
  margin-bottom:20px;
}
.clean-hero-pro h1{
  margin-top:12px;
}
.clean-pill-row{
  display:flex;
  flex-wrap:wrap;
  gap:10px;
  margin-top:18px;
}
.clean-pill{
  padding:10px 14px;
  border-radius:999px;
  background:#EEF2FF;
  color:#4C1D95;
  font-weight:900;
  border:1px solid #DDD6FE;
  font-size:13px;
}
.free-status-box{
  background:linear-gradient(135deg,#EEF2FF,#F5F3FF);
  border:1px solid #DDD6FE;
  color:#4C1D95;
  border-radius:18px;
  padding:14px;
  font-weight:800;
  line-height:1.55;
}


/* ===== SAFE CLEAN MENU + PRICING ONLY ===== */
.safe-menu-title{
  margin:14px 0 8px;
  padding:0 6px;
  font-size:11px;
  font-weight:900;
  letter-spacing:.08em;
  color:#A5B4FC;
  text-transform:uppercase;
}
.safe-menu-link{
  display:flex!important;
  align-items:center;
  gap:10px;
  padding:13px 14px!important;
  border-radius:16px!important;
  color:#F8FAFC!important;
  text-decoration:none!important;
  margin:7px 0!important;
  background:rgba(255,255,255,.06)!important;
  border:1px solid rgba(255,255,255,.08)!important;
  transition:.18s ease!important;
}
.safe-menu-link:hover{
  background:linear-gradient(135deg,#2563EB,#7C3AED)!important;
  transform:translateX(3px)!important;
}
.safe-menu-link .safe-ico{
  width:26px;
  height:26px;
  display:inline-flex;
  align-items:center;
  justify-content:center;
  border-radius:10px;
  background:rgba(255,255,255,.10);
}
.safe-menu-link .safe-text{
  flex:1;
  font-weight:800;
}
.safe-menu-link .safe-tag{
  font-size:10px;
  padding:4px 7px;
  border-radius:999px;
  color:#DDD6FE;
  background:rgba(124,58,237,.18);
  border:1px solid rgba(221,214,254,.18);
}
.safe-premium-box{
  margin-top:14px;
  padding:14px;
  border-radius:18px;
  background:linear-gradient(135deg,rgba(37,99,235,.15),rgba(124,58,237,.16));
  border:1px solid rgba(221,214,254,.22);
  color:#E0E7FF;
  font-size:12px;
  line-height:1.55;
}
.safe-pricing-action{
  display:block;
  width:100%;
  margin-top:12px;
  text-align:center;
  border-radius:14px;
  padding:12px 14px;
  font-weight:900;
  cursor:pointer;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  color:white;
  text-decoration:none;
  border:0;
}
.price-card,.premium-plan{
  cursor:pointer;
}
.price-card:hover,.premium-plan:hover{
  transform:translateY(-4px);
  border-color:#A78BFA!important;
}


/* ===== MKT AUTOMATION PRO - APP MINI PWA MODE ===== */
:root{
  --app-dark:#0F172A;
  --app-primary:#2563EB;
  --app-premium:#7C3AED;
  --app-accent:#06B6D4;
}
.app-shell-top{
  position:sticky;
  top:0;
  z-index:9000;
  background:rgba(15,23,42,.94);
  backdrop-filter:blur(16px);
  color:white;
  border-bottom:1px solid rgba(255,255,255,.10);
  padding:14px 16px;
  display:none;
}
.app-shell-brand{
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:12px;
}
.app-shell-brand b{
  font-size:17px;
}
.app-shell-status{
  font-size:11px;
  color:#C7D2FE;
  margin-top:4px;
}
.app-quick-grid{
  display:grid;
  grid-template-columns:repeat(2,1fr);
  gap:12px;
  margin:18px 0;
}
.app-quick-card{
  border:1px solid #E5E7EB;
  border-radius:22px;
  background:white;
  padding:18px;
  box-shadow:0 14px 30px rgba(15,23,42,.08);
  cursor:pointer;
}
.app-quick-card:hover{
  transform:translateY(-3px);
  border-color:#A78BFA;
}
.app-quick-card .app-ico{
  width:44px;
  height:44px;
  border-radius:16px;
  display:flex;
  align-items:center;
  justify-content:center;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  color:white;
  font-size:22px;
  margin-bottom:10px;
}
.app-quick-card b{
  color:#0F172A;
  display:block;
  margin-bottom:6px;
}
.app-quick-card span{
  color:#64748B;
  font-size:12px;
  line-height:1.45;
}
.app-bottom-nav{
  display:none;
  position:fixed;
  left:0;
  right:0;
  bottom:0;
  z-index:9998;
  background:rgba(255,255,255,.96);
  backdrop-filter:blur(14px);
  border-top:1px solid #E5E7EB;
  box-shadow:0 -10px 30px rgba(15,23,42,.12);
}
.app-bottom-nav a{
  flex:1;
  text-align:center;
  padding:9px 3px 8px;
  color:#475569;
  text-decoration:none;
  font-size:11px;
  font-weight:800;
}
.app-bottom-nav a span{
  display:block;
  font-size:20px;
  margin-bottom:2px;
}
.app-install-banner{
  display:none;
  background:linear-gradient(135deg,#0F172A,#312E81);
  color:white;
  border-radius:24px;
  padding:18px;
  margin:16px 0;
  box-shadow:0 18px 45px rgba(15,23,42,.20);
}
.app-install-banner button{
  margin-top:10px;
  width:100%;
}
@media(max-width:820px){
  body{
    background:#F8FAFC!important;
    padding-bottom:78px!important;
  }
  .app-shell-top{display:block}
  .layout{
    display:block!important;
    padding:12px!important;
    max-width:100%!important;
  }
  .sidebar,.rightbar,.live-trust-bar{
    display:none!important;
  }
  .panel,.top-hero,.premium-pricing-pro{
    border-radius:24px!important;
    padding:18px!important;
    margin-bottom:14px!important;
  }
  h1{
    font-size:28px!important;
    line-height:1.15!important;
  }
  .hero-actions{
    display:grid!important;
    grid-template-columns:1fr!important;
  }
  .module-hub,.grid4,.pricing-grid,.premium-grid-pro{
    grid-template-columns:1fr!important;
    gap:12px!important;
  }
  .app-bottom-nav{
    display:flex!important;
  }
  .app-install-banner{
    display:block!important;
  }
  .floating-bot{
    bottom:88px!important;
    right:14px!important;
  }
  .mobilebar{
    display:none!important;
  }
}


/* ===== MKT AUTOMATION PRO V2 - SAFE FINAL UI ===== */
:root{
  --v2-dark:#0F172A;
  --v2-card:#FFFFFF;
  --v2-blue:#2563EB;
  --v2-purple:#7C3AED;
  --v2-cyan:#06B6D4;
  --v2-soft:#EEF2FF;
  --v2-line:#E5E7EB;
}
.v2-desktop-shell{
  display:block;
}
.v2-topbar{
  position:sticky;
  top:0;
  z-index:8000;
  display:none;
  padding:14px 16px;
  background:rgba(15,23,42,.95);
  color:white;
  backdrop-filter:blur(16px);
  border-bottom:1px solid rgba(255,255,255,.10);
}
.v2-brand-row{
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:12px;
}
.v2-brand-title{
  font-size:17px;
  font-weight:1000;
}
.v2-brand-sub{
  font-size:11px;
  color:#C7D2FE;
  margin-top:3px;
}
.v2-status-pill{
  font-size:11px;
  font-weight:900;
  padding:7px 10px;
  border-radius:999px;
  background:rgba(34,197,94,.16);
  border:1px solid rgba(34,197,94,.24);
  color:#BBF7D0;
}
.v2-nav-title{
  margin:14px 0 8px;
  padding:0 6px;
  font-size:11px;
  font-weight:900;
  letter-spacing:.08em;
  color:#A5B4FC;
  text-transform:uppercase;
}
.v2-nav-link{
  display:flex!important;
  align-items:center;
  gap:10px;
  padding:13px 14px!important;
  border-radius:16px!important;
  color:#F8FAFC!important;
  text-decoration:none!important;
  margin:7px 0!important;
  background:rgba(255,255,255,.06)!important;
  border:1px solid rgba(255,255,255,.08)!important;
  transition:.18s ease!important;
}
.v2-nav-link:hover{
  background:linear-gradient(135deg,#2563EB,#7C3AED)!important;
  transform:translateX(3px)!important;
}
.v2-nav-ico{
  width:28px;
  height:28px;
  display:inline-flex;
  align-items:center;
  justify-content:center;
  border-radius:10px;
  background:rgba(255,255,255,.10);
}
.v2-nav-text{flex:1;font-weight:850}
.v2-nav-tag{
  font-size:10px;
  padding:4px 7px;
  border-radius:999px;
  color:#DDD6FE;
  background:rgba(124,58,237,.20);
  border:1px solid rgba(221,214,254,.20);
}
.v2-side-card{
  margin-top:14px;
  padding:15px;
  border-radius:20px;
  background:linear-gradient(135deg,rgba(37,99,235,.16),rgba(124,58,237,.18));
  border:1px solid rgba(221,214,254,.22);
  color:#E0E7FF;
  font-size:12px;
  line-height:1.55;
}
.v2-hero{
  background:
    radial-gradient(circle at top left,rgba(37,99,235,.22),transparent 34%),
    radial-gradient(circle at top right,rgba(124,58,237,.20),transparent 34%),
    linear-gradient(135deg,#FFFFFF,#F8FAFC);
  border:1px solid #DDD6FE;
  border-radius:32px;
  padding:28px;
  margin-bottom:20px;
  box-shadow:0 22px 55px rgba(30,41,59,.12);
}
.v2-hero-kicker{
  display:inline-flex;
  align-items:center;
  gap:8px;
  padding:10px 15px;
  border-radius:999px;
  background:#EEF2FF;
  color:#4C1D95;
  font-weight:1000;
  border:1px solid #DDD6FE;
}
.v2-hero h1{
  margin:14px 0 8px;
}
.v2-hero-desc{
  max-width:820px;
  color:#64748B;
  font-size:15px;
  line-height:1.6;
}
.v2-quick-grid{
  display:grid;
  grid-template-columns:repeat(4,1fr);
  gap:14px;
  margin-top:20px;
}
.v2-quick-card{
  background:white;
  border:1px solid #E5E7EB;
  border-radius:24px;
  padding:18px;
  box-shadow:0 14px 32px rgba(15,23,42,.08);
  cursor:pointer;
  transition:.18s ease;
}
.v2-quick-card:hover{
  transform:translateY(-4px);
  border-color:#A78BFA;
  box-shadow:0 20px 45px rgba(124,58,237,.14);
}
.v2-quick-ico{
  width:46px;
  height:46px;
  border-radius:17px;
  display:flex;
  align-items:center;
  justify-content:center;
  color:white;
  font-size:22px;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  margin-bottom:12px;
}
.v2-quick-card b{
  display:block;
  color:#0F172A;
  margin-bottom:6px;
}
.v2-quick-card span{
  color:#64748B;
  font-size:12px;
  line-height:1.45;
}
.v2-install-box{
  display:none;
  margin-top:16px;
  border-radius:22px;
  padding:16px;
  color:white;
  background:linear-gradient(135deg,#0F172A,#312E81);
  box-shadow:0 18px 45px rgba(15,23,42,.18);
}
.v2-bottom-nav{
  display:none;
  position:fixed;
  left:0;
  right:0;
  bottom:0;
  z-index:9999;
  background:rgba(255,255,255,.97);
  backdrop-filter:blur(14px);
  border-top:1px solid #E5E7EB;
  box-shadow:0 -10px 30px rgba(15,23,42,.12);
}
.v2-bottom-nav a{
  flex:1;
  text-align:center;
  padding:9px 3px 8px;
  color:#475569;
  text-decoration:none;
  font-size:11px;
  font-weight:900;
}
.v2-bottom-nav span{
  display:block;
  font-size:20px;
  margin-bottom:2px;
}
.v2-pricing-action{
  display:block;
  width:100%;
  margin-top:12px;
  text-align:center;
  border-radius:14px;
  padding:12px 14px;
  font-weight:1000;
  cursor:pointer;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  color:white;
  text-decoration:none;
  border:0;
}
.price-card,.premium-plan{cursor:pointer}
.price-card:hover,.premium-plan:hover{
  transform:translateY(-5px);
  border-color:#A78BFA!important;
}
@media(max-width:1100px){
  .v2-quick-grid{grid-template-columns:repeat(2,1fr)}
}
@media(max-width:820px){
  body{
    background:#F8FAFC!important;
    padding-bottom:78px!important;
  }
  .v2-topbar{display:block!important}
  .layout{
    display:block!important;
    max-width:100%!important;
    padding:12px!important;
  }
  .sidebar,.rightbar,.live-trust-bar,.mobilebar{
    display:none!important;
  }
  .panel,.top-hero,.premium-pricing-pro,.v2-hero{
    border-radius:24px!important;
    padding:18px!important;
    margin-bottom:14px!important;
  }
  h1{font-size:28px!important;line-height:1.15!important}
  .v2-quick-grid,.module-hub,.grid4,.pricing-grid,.premium-grid-pro{
    grid-template-columns:1fr!important;
    gap:12px!important;
  }
  .v2-install-box{display:block!important}
  .v2-bottom-nav{display:flex!important}
  .floating-bot{
    right:14px!important;
    bottom:88px!important;
  }
}



/* V3 Enterprise Add-on */
.v3-kpi-title{margin-top:18px;background:#EEF2FF;border:1px solid #DDD6FE;color:#4C1D95;padding:12px 16px;border-radius:18px;font-weight:800}
.v3-ceo-grid .stat small{display:block;color:#64748B;margin-top:6px;font-size:12px}

.v5-seller-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:14px;margin:14px 0}
.v5-tool-card{background:#fff;border:1px solid var(--border);border-radius:22px;padding:16px;box-shadow:0 12px 28px rgba(37,99,235,.08)}
.v5-tool-card h3{margin:0 0 8px;color:#1E1B4B}.v5-tool-card ul{padding-left:18px;line-height:1.7;color:#334155}
.v5-status-pill{display:inline-block;border-radius:999px;padding:6px 10px;font-size:12px;font-weight:900;background:#ECFDF5;color:#047857;border:1px solid #A7F3D0}
.v5-warning-pill{display:inline-block;border-radius:999px;padding:6px 10px;font-size:12px;font-weight:900;background:#FEF3C7;color:#92400E;border:1px solid #FCD34D}
.v5-table{width:100%;border-collapse:separate;border-spacing:0 8px}.v5-table th{text-align:left;color:#64748B;font-size:13px}.v5-table td{background:#F8FAFC;border-top:1px solid #E5E7EB;border-bottom:1px solid #E5E7EB;padding:10px}.v5-table td:first-child{border-left:1px solid #E5E7EB;border-radius:14px 0 0 14px}.v5-table td:last-child{border-right:1px solid #E5E7EB;border-radius:0 14px 14px 0}

.v3-feature-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px;margin-top:14px}
.v3-feature-card{background:#fff;border:1px solid var(--border);border-radius:22px;padding:18px;box-shadow:0 12px 28px rgba(30,41,59,.08)}
.v3-feature-card h3{margin:0 0 8px;color:#1E1B4B}.v3-feature-card ul{margin:8px 0 0;padding-left:18px;line-height:1.65;color:#334155}
.v3-premium-badge{display:inline-block;background:linear-gradient(135deg,#2563EB,#7C3AED);color:white;border-radius:999px;padding:6px 10px;font-size:12px;font-weight:900;margin-bottom:8px}
.kanban-board{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-top:14px}.kanban-col{background:#F8FAFC;border:1px solid #E5E7EB;border-radius:18px;padding:12px;min-height:180px}.kanban-col h3{font-size:15px;margin:0 0 10px;color:#1E1B4B}.kanban-card{background:white;border:1px solid #DDD6FE;border-left:4px solid #7C3AED;border-radius:14px;padding:10px;margin-bottom:10px;cursor:grab;box-shadow:0 8px 18px rgba(124,58,237,.08)}
.v3-mini-table{width:100%;border-collapse:collapse}.v3-mini-table td,.v3-mini-table th{border-bottom:1px solid #EEF2FF;padding:10px;text-align:left}.v3-mini-table th{color:#1E1B4B}
@media(max-width:1000px){.kanban-board{grid-template-columns:1fr 1fr}.v3-feature-grid{grid-template-columns:1fr}}
@media(max-width:620px){.kanban-board{grid-template-columns:1fr}}


/* ===== V4 PREMIUM CONVERSION UPGRADE ===== */
.v4-trust-grid,.v4-pricing-stats{
  display:grid;
  grid-template-columns:repeat(3,minmax(0,1fr));
  gap:14px;
  margin:18px 0;
}
.v4-trust-card,.v4-pricing-stats div{
  background:linear-gradient(135deg,#FFFFFF,#F8FAFC);
  border:1px solid #DDD6FE;
  border-radius:22px;
  padding:16px;
  text-align:center;
  box-shadow:0 14px 32px rgba(30,41,59,.08);
}
.v4-trust-card b,.v4-pricing-stats b{
  display:block;
  font-size:26px;
  color:transparent;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  -webkit-background-clip:text;
  background-clip:text;
}
.v4-trust-card span,.v4-pricing-stats span{font-size:13px;color:#64748B;font-weight:800}
.v4-premium-hero{
  background:radial-gradient(circle at top left,rgba(37,99,235,.12),transparent 34%),linear-gradient(135deg,#FFFFFF,#F5F3FF)!important;
  padding:26px!important;
}
.v4-hero-label,.value-badge{
  display:inline-block;
  background:linear-gradient(135deg,#2563EB,#7C3AED);
  color:white;
  padding:8px 13px;
  border-radius:999px;
  font-weight:900;
  font-size:12px;
  margin-bottom:10px;
}
.v4-value-grid{
  display:grid;
  grid-template-columns:repeat(3,1fr);
  gap:12px;
  margin:18px 0;
}
.v4-value-grid div{
  background:white;
  border:1px solid #DDD6FE;
  border-radius:20px;
  padding:14px;
}
.v4-value-grid b{display:block;color:#4C1D95;font-size:22px}
.v4-value-grid span{font-size:13px;color:#64748B;font-weight:800}
.v4-pricing-shell{padding:32px!important}
.v4-premium-grid{
  grid-template-columns:repeat(auto-fit,minmax(250px,1fr))!important;
  align-items:stretch;
}
.v4-plan{
  border-radius:28px!important;
  padding:22px!important;
  overflow:hidden;
}
.v4-plan .plan-name{font-size:17px!important}
.v4-plan .plan-price{
  font-size:24px!important;
  line-height:1.15!important;
  margin:8px 0!important;
}
.price-sub{
  color:#64748B;
  font-size:13px;
  font-weight:900;
  margin-bottom:12px;
}
.v4-yearly{
  transform:scale(1.03);
  border:2px solid #7C3AED!important;
  box-shadow:0 32px 75px rgba(124,58,237,.22)!important;
}
.v4-lifetime{
  border:2px solid #F59E0B!important;
  background:linear-gradient(180deg,#FFFFFF,#FFFBEB)!important;
  box-shadow:0 32px 75px rgba(245,158,11,.18)!important;
}
.v4-save-box{
  background:linear-gradient(135deg,#EEF2FF,#F5F3FF);
  border:1px solid #DDD6FE;
  border-radius:16px;
  padding:12px;
  color:#4C1D95;
  font-weight:900;
  margin:12px 0;
  font-size:13px;
}
.v4-save-box.gold{
  background:linear-gradient(135deg,#FEF3C7,#FFF7ED);
  border-color:#FCD34D;
  color:#92400E;
}
.v4-value-received{
  background:#F8FAFC;
  border:1px solid #E5E7EB;
  border-radius:16px;
  padding:12px;
  margin:12px 0;
}
.v4-value-received b{display:block;color:#1E1B4B;margin-bottom:4px}
.v4-value-received span{font-size:13px;color:#64748B}
.premium-btn{
  background:linear-gradient(135deg,#2563EB,#7C3AED)!important;
  border-radius:16px!important;
  font-weight:900!important;
  box-shadow:0 14px 30px rgba(124,58,237,.20)!important;
}
.v4-outsourcing-box{
  margin-top:22px;
  background:white;
  border:1px solid #DDD6FE;
  border-radius:26px;
  padding:22px;
  box-shadow:0 18px 42px rgba(30,41,59,.08);
}
.v4-outsource-grid{
  display:grid;
  grid-template-columns:repeat(4,1fr);
  gap:12px;
  margin:14px 0;
}
.v4-outsource-grid div{
  background:#F8FAFC;
  border:1px solid #E5E7EB;
  border-radius:18px;
  padding:14px;
}
.v4-outsource-grid b{display:block;color:#7C3AED}
.v4-outsource-grid span{font-size:13px;color:#64748B;font-weight:800}
.v4-success-grid{
  display:grid;
  grid-template-columns:repeat(auto-fit,minmax(220px,1fr));
  gap:14px;
}
.v4-success-card{
  background:white;
  border:1px solid #E5E7EB;
  border-radius:24px;
  padding:18px;
  box-shadow:0 16px 36px rgba(30,41,59,.08);
}
.success-icon{
  width:52px;height:52px;border-radius:18px;
  display:flex;align-items:center;justify-content:center;
  font-size:26px;background:#F5F3FF;border:1px solid #DDD6FE;
  margin-bottom:12px;
}
.v4-proof-box{
  margin-top:16px;
  background:linear-gradient(135deg,#EEF2FF,#F8FAFC);
  border:1px solid #DDD6FE;
  border-radius:24px;
  padding:18px;
}
.v4-lock-hero .benefit-list{
  display:grid;
  grid-template-columns:1fr 1fr;
  gap:8px;
}
.v4-gold-card{
  border:2px solid #F59E0B!important;
  background:linear-gradient(180deg,#FFFFFF,#FFFBEB)!important;
}
.rec-price{font-size:24px!important;line-height:1.15!important}
@media(max-width:820px){
  .v4-trust-grid,.v4-pricing-stats,.v4-value-grid,.v4-outsource-grid,.locked-recommend-grid{grid-template-columns:1fr!important}
  .v4-yearly{transform:none}
  .v4-lock-hero .benefit-list{grid-template-columns:1fr}
}


/* ===== V5 PREMIUM DARK POLISH ===== */
body{
  background:radial-gradient(circle at top left,rgba(124,58,237,.22),transparent 30%),radial-gradient(circle at top right,rgba(37,99,235,.18),transparent 28%),linear-gradient(135deg,#0B1020,#111827 48%,#1E293B)!important;
}
.panel,.rightbar,.top-hero{
  background:rgba(15,23,42,.92)!important;
  color:#F8FAFC!important;
  border:1px solid rgba(148,163,184,.22)!important;
  box-shadow:0 22px 60px rgba(0,0,0,.28)!important;
}
.top-hero h1,h1,.rightbar h2{
  background:linear-gradient(135deg,#38BDF8,#8B5CF6)!important;
  -webkit-background-clip:text!important;
  background-clip:text!important;
}
.panel p,.top-hero p,.rightbar p,.small{color:#CBD5E1!important}
.app-quick-card,.template-card,.stat,.price-card,.preview,.locked-recommend-card{
  background:rgba(255,255,255,.96)!important;
}
.activity-card{
  display:flex;
  justify-content:space-between;
  align-items:center;
  gap:12px;
  padding:14px 15px;
  margin:10px 0;
  border-radius:18px;
  background:linear-gradient(135deg,rgba(37,99,235,.15),rgba(124,58,237,.14));
  border:1px solid rgba(148,163,184,.22);
  color:#E5E7EB;
}
.activity-card span{font-weight:800;color:#CBD5E1}
.activity-card b{font-size:24px;color:#38BDF8}
.v5-focus-box{
  margin-top:14px;
  padding:16px;
  border-radius:20px;
  background:linear-gradient(135deg,rgba(56,189,248,.12),rgba(139,92,246,.16));
  border:1px solid rgba(56,189,248,.28);
  color:#E5E7EB;
  line-height:1.65;
}
.v5-focus-box b{color:#FBBF24}

/* ===== ANALYTICS CENTER V5 POLISH ===== */
.analytics-kpi-grid{
  display:grid;
  grid-template-columns:repeat(4,1fr);
  gap:14px;
  margin:18px 0 22px;
}
.analytics-kpi{
  background:linear-gradient(135deg,rgba(37,99,235,.10),rgba(124,58,237,.10));
  border:1px solid rgba(124,58,237,.22);
  border-radius:20px;
  padding:16px;
  box-shadow:0 12px 28px rgba(37,99,235,.08);
}
.analytics-kpi span{display:block;color:#475569;font-size:13px;font-weight:800}
.analytics-kpi b{display:block;color:#4C1D95;font-size:28px;margin:7px 0}
.analytics-kpi small{display:block;color:#64748B;line-height:1.45}
.analytics-section-title{
  margin:24px 0 12px;
  color:#1E1B4B;
  font-size:17px;
  font-weight:900;
}
.analytics-box.full{grid-column:1/-1}
.analytics-bar{height:8px;background:#E5E7EB;border-radius:999px;overflow:hidden;margin:-2px 0 8px}
.analytics-bar.large{height:12px;margin:2px 0 12px}
.analytics-bar i{display:block;height:100%;background:linear-gradient(135deg,var(--blue),var(--purple));border-radius:999px;min-width:6px}
.empty-analytics{
  background:#F8FAFC;
  border:1px dashed #CBD5E1;
  border-radius:14px;
  color:#64748B;
  padding:14px;
  line-height:1.55;
}
@media(max-width:1100px){.analytics-kpi-grid{grid-template-columns:repeat(2,1fr)}}
@media(max-width:700px){.analytics-kpi-grid{grid-template-columns:1fr}}


/* Premium active status: blue 3D seller badge */
.premium-status-text{
  display:inline-block;
  font-weight:900;
}
body.premium-active .premium-status-text{
  color:#1D4ED8;
  padding:4px 8px;
  border-radius:999px;
  background:linear-gradient(135deg,#DBEAFE,#EFF6FF 45%,#C7D2FE);
  border:1px solid rgba(37,99,235,.25);
  text-shadow:0 1px 0 #FFFFFF, 0 2px 0 rgba(37,99,235,.18), 0 8px 18px rgba(37,99,235,.28);
  box-shadow:inset 0 1px 0 rgba(255,255,255,.95), 0 8px 20px rgba(37,99,235,.22);
  letter-spacing:.2px;
}
body.premium-active .device-status-card{
  border-color:rgba(37,99,235,.45)!important;
  background:linear-gradient(135deg,rgba(219,234,254,.95),rgba(255,255,255,.96))!important;
}



/* ===== PREMIUM FONT UPGRADE - LUXURY SAAS TYPOGRAPHY ===== */
:root{
  --font-heading:'Manrope','Inter',Arial,sans-serif;
  --font-body:'Inter','Manrope',Arial,sans-serif;
  --luxury-blue:#38BDF8;
  --luxury-indigo:#6366F1;
  --luxury-purple:#A855F7;
  --luxury-gold:#FACC15;
}
html,body,button,input,textarea,select{
  font-family:var(--font-body)!important;
  text-rendering:geometricPrecision;
  -webkit-font-smoothing:antialiased;
  -moz-osx-font-smoothing:grayscale;
}
body{font-weight:500;letter-spacing:-.01em;}
h1,h2,h3,.logo,.v2-brand-title,.premium-title h2,.price-bottom-title h2,.rightbar h2,.module-card h3,.app-quick-card h3,.template-card h3,.plan-name,.card-title,.section-title,.analytics-section-title{
  font-family:var(--font-heading)!important;
  letter-spacing:-.045em!important;
}
h1,.top-hero h1{
  font-size:clamp(38px,4.2vw,62px)!important;
  line-height:1.02!important;
  font-weight:800!important;
  background:linear-gradient(90deg,#38BDF8 0%,#60A5FA 28%,#6366F1 58%,#A855F7 100%)!important;
  -webkit-background-clip:text!important;
  background-clip:text!important;
  -webkit-text-fill-color:transparent!important;
  text-shadow:0 18px 55px rgba(99,102,241,.25)!important;
}
.logo{
  font-size:30px!important;
  line-height:1.08!important;
  font-weight:800!important;
  letter-spacing:-.055em!important;
}
.subtitle,.panel p,.top-hero p,.rightbar p,.module-card p,.app-quick-card p,.template-card p,.plan-desc,.price-sub,.small{
  font-family:var(--font-body)!important;
  font-size:15px!important;
  line-height:1.72!important;
  font-weight:500!important;
  letter-spacing:-.01em!important;
}
.nav a,.safe-menu-link,.v2-nav-link{
  font-family:var(--font-heading)!important;
  font-weight:800!important;
  letter-spacing:-.025em!important;
}
.module-card h3,.app-quick-card h3,.template-card h3,.price-card h3,.plan-name{
  font-size:22px!important;
  line-height:1.18!important;
  font-weight:800!important;
  color:#0F172A!important;
}
.module-card p,.app-quick-card p,.template-card p{
  color:#64748B!important;
  font-size:15px!important;
  line-height:1.75!important;
}
.price,.plan-price,.rec-price{
  font-family:var(--font-heading)!important;
  font-size:clamp(28px,2.3vw,38px)!important;
  font-weight:800!important;
  letter-spacing:-.06em!important;
  line-height:1.04!important;
  background:linear-gradient(90deg,#FDE68A 0%,#FACC15 35%,#F59E0B 68%,#FFEFD5 100%)!important;
  -webkit-background-clip:text!important;
  background-clip:text!important;
  -webkit-text-fill-color:transparent!important;
  text-shadow:0 0 28px rgba(250,204,21,.35)!important;
}
.price-card,.premium-plan,.locked-recommend-card{
  border-radius:30px!important;
  box-shadow:0 24px 70px rgba(15,23,42,.15)!important;
}
.price-card:hover,.premium-plan:hover,.module-card:hover,.app-quick-card:hover{
  transform:translateY(-6px)!important;
  box-shadow:0 30px 85px rgba(99,102,241,.22)!important;
}
.premium-status-text,#liveMemberCount,.activity-card b,.stat b{
  font-family:var(--font-heading)!important;
  font-weight:800!important;
  letter-spacing:-.045em!important;
}
.premium-title .mini,.ribbon,.plan-ribbon,.module-pill,.safe-tag{
  font-family:var(--font-heading)!important;
  font-weight:800!important;
  letter-spacing:-.02em!important;
}
button,.btn,.safe-pricing-action,.support-send{
  font-family:var(--font-heading)!important;
  font-weight:800!important;
  letter-spacing:-.025em!important;
}
.device-status-card b,.premium-confirm-box h3{
  font-family:var(--font-heading)!important;
  letter-spacing:-.03em!important;
}
@media(max-width:720px){
  h1,.top-hero h1{font-size:36px!important;letter-spacing:-.05em!important}
  .logo{font-size:26px!important}
  .module-card h3,.app-quick-card h3,.template-card h3{font-size:20px!important}
}

</style>

<script>
function copyText(id){
  const el=document.getElementById(id);
  navigator.clipboard.writeText(el.innerText);
  alert("Đã copy content mẫu");
}

function openPremiumPopup(){
  const m=document.getElementById("premiumPopup");
  if(m){m.style.display="flex";}
}
function closePremiumPopup(){
  const m=document.getElementById("premiumPopup");
  if(m){m.style.display="none";}
}

function toggleFloatingBot(){
  const panel=document.getElementById("floatingBotPanel");
  if(!panel) return;
  const isOpen=panel.style.display==="block";
  panel.style.display=isOpen?"none":"block";
  if(!isOpen){
    setTimeout(()=>appendBotGreeting(),250);
  }
}
function closeFloatingBot(){
  const panel=document.getElementById("floatingBotPanel");
  if(panel) panel.style.display="none";
}
let botGreeted=false;
function appendBotGreeting(){
  if(botGreeted) return;
  botGreeted=true;
  const body=document.getElementById("floatingBotBody");
  if(!body) return;
  body.innerHTML += `
    <div class="bot-msg ai">
      Xin chào anh/chị 👋<br><br>
      Em là trợ lý hỗ trợ của <b>Marketing Automation Pro V10</b>.<br>
      Anh/chị cần tư vấn gói Premium, kích hoạt tài khoản, thanh toán hay hướng dẫn sử dụng tool thì em hỗ trợ ngay ạ.
    </div>
    <div class="bot-msg ai">
      Hiện hệ thống có các gói: <b>1 tháng 159K</b>, <b>3 tháng 359K</b>, <b>6 tháng 559K</b>, <b>1 năm 859K</b> và <b>Nhà bán hàng chuyên nghiệp 1.959K</b>.<br>
      Nếu đã thanh toán mà sau 5 phút chưa kích hoạt, vui lòng liên hệ Zalo <b>036 338 2629</b>.
    </div>`;
  body.scrollTop=body.scrollHeight;
}
function botQuick(text){
  const body=document.getElementById("floatingBotBody");
  if(!body) return;
  body.innerHTML += `<div class="bot-msg"><b>Bạn:</b> ${text}</div>`;
  let reply="Dạ em đã nhận yêu cầu. Nếu hệ thống AI báo quá tải hoặc hết lượt tạm thời, anh/chị vui lòng thử lại sau ít phút hoặc liên hệ Zalo để kỹ thuật hỗ trợ nhanh ạ.";
  const lower=text.toLowerCase();
  if(lower.includes("giá") || lower.includes("gói")){
    reply="Dạ gói Premium hiện có: 1 tháng 159K, 3 tháng 359K, 6 tháng 559K, 1 năm 859K và Nhà bán hàng chuyên nghiệp 1.959K. Gói 1 năm đang phổ biến nhất, còn gói Nhà bán hàng chuyên nghiệp mở toàn bộ tính năng và cập nhật tương lai.";
  }else if(lower.includes("thanh toán")){
    reply="Anh/chị bấm Bảng Giá Premium Seller AI, chọn gói cần mua, hệ thống sẽ mở popup thanh toán kèm QR Agribank, STK 8888363382629 - NGUYEN DANG THI XUAN. Nếu sau 5 phút chưa kích hoạt, liên hệ Zalo 036 338 2629.";
  }else if(lower.includes("tính năng")){
    reply="Tool hỗ trợ AI Content Brain, đăng nhiều Fanpage, lịch đăng, chia content/ảnh, CRM Pro, AI Sales Bot, Marketing Funnel, kho content 50.000+, xuất báo cáo và backup dữ liệu.";
  }
  body.innerHTML += `<div class="bot-msg ai"><b>Bot hỗ trợ:</b><br>${reply}</div>`;
  body.scrollTop=body.scrollHeight;
}
function sendBotInput(){
  const input=document.getElementById("botInputText");
  if(!input || !input.value.trim()) return;
  botQuick(input.value.trim());
  input.value="";
}
function updateLiveTrust(){
  const el=document.getElementById("liveMemberCount");
  if(!el) return;
  let base=Number(localStorage.getItem("v10_premium_count") || "231");
  base = base + 1;
  if(base > 999){ base = 231; }
  localStorage.setItem("v10_premium_count", String(base));
  el.innerText = base;
}
document.addEventListener("DOMContentLoaded",function(){
  const el=document.getElementById("liveMemberCount");
  if(el){ el.innerText = localStorage.getItem("v10_premium_count") || "231"; }
  setInterval(updateLiveTrust,4500);
});



let MKT_DEVICE_ID = "";
let MKT_SELECTED_PLAN = "monthly";
let MKT_SELECTED_PACKAGE = "1THANG";
let MKT_PREMIUM_ACTIVE = false;

function getOrCreateDeviceId(){
  let id = localStorage.getItem("mkt_device_id");
  if(!id){
    const rand = Math.random().toString(16).slice(2,8).toUpperCase();
    const d = new Date();
    const ymd = d.getFullYear().toString() + String(d.getMonth()+1).padStart(2,'0') + String(d.getDate()).padStart(2,'0');
    id = "MP-" + ymd + "-" + rand;
    localStorage.setItem("mkt_device_id", id);
  }
  MKT_DEVICE_ID = id;
  document.querySelectorAll(".device-id-text").forEach(function(el){el.innerText=id;});
  return id;
}

async function checkPremiumStatus(){
  const id = getOrCreateDeviceId();
  try{
    const res = await fetch("/premium_status?device_id=" + encodeURIComponent(id));
    const data = await res.json();
    MKT_PREMIUM_ACTIVE = !!data.active;
    const label = data.active ? "Premium: 👑 GÓI NHÀ BÁN HÀNG CHUYÊN NGHIỆP" : "Dùng thử / chờ duyệt";
    document.querySelectorAll(".premium-status-text").forEach(function(el){el.innerText=label;});
    if(data.active){
      document.body.classList.add("premium-active");
      document.querySelectorAll(".premium-unlock-note").forEach(function(el){el.innerText="Đã mở khóa Premium cho thiết bị này.";});
    }else{
      document.body.classList.remove("premium-active");
    }
  }catch(e){
    document.querySelectorAll(".premium-status-text").forEach(function(el){el.innerText="Dùng thử / chờ duyệt";});
  }
}

document.addEventListener("DOMContentLoaded", function(){
  getOrCreateDeviceId();
  checkPremiumStatus();

  // Gắn device_id vào mọi form để backend nhận biết tài khoản đã Premium.
  document.querySelectorAll("form").forEach(function(form){
    if(!form.querySelector("input[name='device_id']")){
      const hidden=document.createElement("input");
      hidden.type="hidden";
      hidden.name="device_id";
      hidden.value=getOrCreateDeviceId();
      form.appendChild(hidden);
    }
    form.addEventListener("submit", function(){
      const h=form.querySelector("input[name='device_id']");
      if(h){h.value=getOrCreateDeviceId();}
    });
  });
});

function scrollToPaymentConfirm(){
  const el=document.getElementById("premiumConfirmFormBox");
  if(el){el.scrollIntoView({behavior:"smooth", block:"center"});}
}

async function submitPremiumRequest(){
  const deviceId = getOrCreateDeviceId();
  const phone = (document.getElementById("premiumPhone") || {}).value || "";
  const email = (document.getElementById("premiumEmail") || {}).value || "";
  const note = (document.getElementById("premiumTransactionNote") || {}).value || "";
  const statusBox = document.getElementById("premiumRequestStatus");
  if(!phone.trim()){
    if(statusBox) statusBox.innerText="Vui lòng nhập số điện thoại/Zalo.";
    return false;
  }
  if(!email.trim() || !email.includes("@")){
    if(statusBox) statusBox.innerText="Vui lòng nhập đúng Gmail/Email.";
    return false;
  }
  if(statusBox) statusBox.innerText="Đang gửi yêu cầu về Admin...";
  try{
    const plan = premiumPlans[MKT_SELECTED_PLAN] || premiumPlans.monthly;
    const res = await fetch("/premium_request", {
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify({
        device_id: deviceId,
        phone: phone,
        email: email,
        transaction_note: note,
        plan_key: MKT_SELECTED_PLAN,
        plan_name: plan.title || plan.package || MKT_SELECTED_PLAN,
        amount: plan.amount || 0
      })
    });
    const data = await res.json();
    if(statusBox) statusBox.innerText = data.message || "Đã gửi yêu cầu.";
    setTimeout(checkPremiumStatus, 800);
  }catch(e){
    if(statusBox) statusBox.innerText="Chưa gửi được yêu cầu. Vui lòng thử lại hoặc liên hệ Zalo.";
  }
  return false;
}

function openModule(moduleId){
  const trialAllowed = ["dashboard", "facebook_center", "fanpage_manager", "group_marketing", "comment_manager", "post", "facebook_publisher_pro", "token", "premium"];
  const premiumLocked = {
    "messenger_ai": "AI Messenger",
    "crm_sales": "CRM Kanban",
    "marketing_director": "AI Marketing Director",
    "ai_studio": "AI Studio",
    "creative_center": "Creative Center",
    "scheduler": "Content Calendar",
    "plan": "Kế hoạch Marketing 30 ngày",
    "analytics": "Analytics Center",
    "automation_center": "Automation Center",
    "success_center": "Success Center",
    "batch": "Đăng hàng loạt",
    "factory": "Content Factory",
    "clusters": "Page Cluster",
    "campaign": "Campaign Manager",
    "smart_engagement": "Tăng tương tác thông minh"
  };

  if(!trialAllowed.includes(moduleId) && premiumLocked[moduleId] && !MKT_PREMIUM_ACTIVE){
    openLockedFeature(premiumLocked[moduleId], "Gói 1 tháng / 3 tháng / 6 tháng / 1 năm / Nhà bán hàng chuyên nghiệp");
    return false;
  }

  document.querySelectorAll(".module-section").forEach(function(el){
    el.classList.remove("active-module");
  });
  const target=document.getElementById(moduleId);
  if(target){
    target.classList.add("active-module");
    target.scrollIntoView({behavior:"smooth",block:"start"});
  }
  return false;
}
function showAllModules(){
  document.querySelectorAll(".module-section").forEach(function(el){
    el.classList.add("active-module");
  });
}
document.addEventListener("DOMContentLoaded",function(){
  const first=document.getElementById("dashboard");
  if(first){first.classList.add("active-module");}
});


function scrollToPricing(){
  const el=document.getElementById("pricing");
  if(el){
    el.style.display="block";
    el.scrollIntoView({behavior:"smooth",block:"start"});
  }else{
    openPayment('monthly');
  }
}


const premiumPlans = {
  trial:{
    title:"🎁 GÓI TRẢI NGHIỆM",
    price:"3 ngày miễn phí",
    amount:"0",
    package:"TRIAL",
    desc:"Nâng cấp để mở toàn bộ hệ sinh thái AI bán hàng.",
    benefits:["Quản lý Fanpage", "Quản lý Group", "AI Comment"],
    locked:["AI Messenger", "CRM Kanban", "AI Marketing Director"]
  },
  monthly:{
    title:"🚀 GÓI 1 THÁNG",
    price:"159.000đ",
    amount:"159000",
    package:"1THANG",
    desc:"Phù hợp người mới bắt đầu dùng hệ thống để quản lý Fanpage, Group và AI Comment.",
    benefits:["Quản lý Fanpage", "Quản lý Group", "AI Comment", "Đăng bài Facebook", "Lịch đăng cơ bản", "Token Manager", "Hỗ trợ cơ bản"],
    locked:["AI Messenger nâng cao", "CRM Kanban", "AI Marketing Director", "AI Video/Image/Voice"]
  },
  quarterly:{
    title:"⭐ GÓI 3 THÁNG",
    price:"359.000đ",
    amount:"359000",
    package:"3THANG",
    desc:"Tối ưu cho shop đang bán hàng cần thêm Messenger AI và quản lý khách hàng.",
    benefits:["Toàn bộ gói 1 tháng", "AI Messenger", "CRM Kanban", "Kịch bản inbox", "Kịch bản chốt sale", "Quản lý khách hàng", "Ưu tiên hỗ trợ"],
    locked:["AI Marketing Director", "AI Video/Image/Voice", "AI Livestream"]
  },
  halfyear:{
    title:"💎 GÓI 6 THÁNG",
    price:"559.000đ",
    amount:"559000",
    package:"6THANG",
    desc:"Mở thêm AI Marketing Director và các công cụ tạo nội dung nâng cao.",
    benefits:["Toàn bộ gói 3 tháng", "AI Marketing Director", "AI Image", "AI Video", "AI Kinh Doanh", "Content Studio", "Không giới hạn Fanpage/Group theo cấu hình"],
    locked:["AI Giọng Nói nâng cao", "AI Livestream nâng cao", "VIP Support"]
  },
  yearly:{
    title:"🔥 GÓI 1 NĂM",
    price:"859.000đ",
    amount:"859000",
    package:"1NAM",
    desc:"Gói phổ biến nhất cho người bán hàng muốn dùng đầy đủ công cụ AI Marketing trong 1 năm.",
    benefits:["Toàn bộ gói 6 tháng", "AI Giọng Nói", "AI Livestream", "AI Automation", "Kho content Premium", "Ưu tiên xử lý", "Cập nhật miễn phí trong thời gian sử dụng"],
    locked:["Gói nhà bán hàng chuyên nghiệp", "Tư vấn VIP", "Cập nhật trọn đời"]
  },
  lifetime:{
    title:"👑 GÓI NHÀ BÁN HÀNG CHUYÊN NGHIỆP",
    price:"1.959.000đ",
    amount:"1959000",
    package:"NHABANHANGCHUYENNGHIEP",
    desc:"Gói cao nhất dành cho nhà bán hàng, agency và team kinh doanh muốn mở toàn bộ hệ thống.",
    benefits:["Toàn bộ tính năng Premium", "Không giới hạn AI theo cấu hình", "AI Marketing Director Pro", "CRM Pro", "AI Automation Pro", "AI Livestream", "AI Giọng Nói", "Export PDF/Excel", "Backup Database", "Ưu tiên hỗ trợ VIP", "Cập nhật trọn đời"],
    locked:[]
  },
  sellerpro:null,
  basic:null,
  pro:null,
  business:null
};
premiumPlans.sellerpro = premiumPlans.lifetime;
premiumPlans.basic = premiumPlans.monthly;
premiumPlans.pro = premiumPlans.quarterly;
premiumPlans.business = premiumPlans.halfyear;

function openPayment(planKey){
  const plan=premiumPlans[planKey] || premiumPlans.basic;
  const modal=document.getElementById("paymentModal");
  if(!modal) return;

  const amountText = Number(plan.amount).toLocaleString("vi-VN") + " VNĐ";
  MKT_SELECTED_PLAN = planKey || "monthly";
  MKT_SELECTED_PACKAGE = (plan.package || MKT_SELECTED_PLAN).toUpperCase();
  const deviceId = getOrCreateDeviceId();
  const addInfo = "PREMIUM " + MKT_SELECTED_PACKAGE + " " + deviceId;
  const qrUrl = "https://img.vietqr.io/image/970405-8888363382629-compact2.png?amount=" + encodeURIComponent(plan.amount) + "&addInfo=" + encodeURIComponent(addInfo) + "&accountName=" + encodeURIComponent("NGUYEN DANG THI XUAN");

  document.getElementById("payPlanTitle").innerText=plan.title;
  document.getElementById("payPlanPrice").innerText=amountText;
  document.getElementById("payPlanDesc").innerText=plan.desc;
  document.getElementById("payQr").src=qrUrl;
  document.getElementById("payContent").innerText=addInfo;
  document.getElementById("payBenefits").innerHTML=plan.benefits.map(x=>"<div>✓ "+x+"</div>").join("");
  document.getElementById("payLocked").innerHTML=plan.locked.length ? plan.locked.map(x=>"<div>🔒 "+x+"</div>").join("") : "<div>✓ Không khóa tính năng</div>";
  modal.style.display="flex";
}
function closePayment(){
  const modal=document.getElementById("paymentModal");
  if(modal) modal.style.display="none";
}


function openLockedFeature(feature, plans){
  const modal=document.getElementById("lockedFeatureModal");
  if(!modal) {
    scrollToPricing();
    return false;
  }

  document.getElementById("lockedFeatureTitle").innerText="🔒 Tính năng Premium";
  document.getElementById("lockedFeaturePlans").innerHTML =
    `<div class="premium-lock-hero v4-lock-hero">
       <p>Bạn đang muốn mở: <b>${feature}</b></p>
       <p>Gói dùng thử 3 ngày chỉ hỗ trợ:</p>
       <ul class="benefit-list">
         <li class="open">Quản lý Fanpage</li>
         <li class="open">Quản lý Group</li>
         <li class="open">AI Comment</li>
       </ul>
       <p><b>Nâng cấp Premium để mở khóa:</b></p>
       <ul class="benefit-list">
         <li class="open">AI Messenger</li>
         <li class="open">CRM Kanban</li>
         <li class="open">AI Marketing Director</li>
         <li class="open">AI Video • AI Image • AI Kinh Doanh</li>
         <li class="open">AI Giọng Nói • AI Livestream • AI Automation</li>
       </ul>
       <div class="locked-recommend-grid">
         <div class="locked-recommend-card">
           <div class="rec-label">🚀 BẮT ĐẦU</div>
           <h3>Gói 1 tháng</h3>
           <div class="rec-price">159.000đ</div>
           <p>Phù hợp để mở các công cụ quản lý cơ bản và dùng ổn định.</p>
           <button onclick="closeLockedFeature();openPayment('monthly')">Chọn gói 1 tháng</button>
         </div>
         <div class="locked-recommend-card featured">
           <div class="rec-label">⭐ PHỔ BIẾN</div>
           <h3>Gói 1 năm</h3>
           <div class="rec-price">859.000đ</div>
           <p>Mở đầy đủ AI Marketing, CRM, Automation và công cụ bán hàng nâng cao.</p>
           <button onclick="closeLockedFeature();openPayment('yearly')">Xem gói 1 năm</button>
         </div>
         <div class="locked-recommend-card v4-gold-card" style="grid-column:1/-1">
           <div class="rec-label">👑 CAO NHẤT</div>
           <h3>Gói nhà bán hàng chuyên nghiệp</h3>
           <div class="rec-price">1.959.000đ</div>
           <p>Mở toàn bộ tính năng cao cấp, hỗ trợ VIP và cập nhật trọn đời.</p>
           <button onclick="closeLockedFeature();openPayment('lifetime')">Mở gói chuyên nghiệp</button>
         </div>
       </div>
       <button class="secondary" onclick="closeLockedFeature();scrollToPricing()">Xem chi tiết tất cả gói</button>
       <p class="small">Sau khi thanh toán 5 phút chưa kích hoạt, vui lòng gửi ảnh giao dịch qua Zalo 036 338 2629.</p>
     </div>`;
  modal.style.display="flex";
  return false;
}
function closeLockedFeature(){
  const modal=document.getElementById("lockedFeatureModal");
  if(modal) modal.style.display="none";
}

</script>

<style>
.premium-grid-pro{grid-template-columns:repeat(auto-fit,minmax(230px,1fr)) !important;}
.locked-plan-actions{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px;margin:14px 0;}
.premium-trial-stat{background:linear-gradient(135deg,#FFFFFF,#F5F3FF)!important;border-left-color:#F59E0B!important}
.token-status-stat{background:linear-gradient(135deg,#FFFFFF,#ECFDF5)!important;border-left-color:#22C55E!important}
.fb-submenu-pro{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:10px;margin:14px 0 18px}
.fb-submenu-pro button{width:100%;margin:0;box-shadow:none}
.fb-submenu-pro .locked{background:linear-gradient(135deg,#EEF2FF,#F5F3FF);color:#4C1D95;border:1px solid #DDD6FE}
.premium-lock-hero p{line-height:1.55}.locked-recommend-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin:14px 0}.locked-recommend-card{background:white;border:1px solid #E5E7EB;border-radius:22px;padding:16px;box-shadow:0 14px 32px rgba(30,41,59,.10)}.locked-recommend-card.featured{border:2px solid #7C3AED;background:linear-gradient(180deg,#FFFFFF,#F5F3FF)}.rec-label{display:inline-block;background:linear-gradient(135deg,#2563EB,#7C3AED);color:white;border-radius:999px;padding:5px 10px;font-size:11px;font-weight:900}.rec-price{font-size:28px;font-weight:900;color:#7C3AED;margin:6px 0}@media(max-width:720px){.locked-recommend-grid{grid-template-columns:1fr}}
.locked-plan-actions button{width:100%;margin:0;}

.trial-box{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.14);border-radius:16px;padding:12px;margin:12px 0;line-height:1.75;color:#E5E7EB}
.locked-list{background:rgba(239,68,68,.08);border-color:rgba(239,68,68,.25)}
.benefit-list .locked{color:#DC2626;list-style:none;margin:6px 0;font-weight:800}
.v2-nav-link[href="#messenger_ai"] .v2-nav-tag,
.v2-nav-link[href="#crm_sales"] .v2-nav-tag,
.v2-nav-link[href="#marketing_director"] .v2-nav-tag{background:linear-gradient(135deg,#F59E0B,#EF4444)!important;color:white!important}

.device-status-card{
  background:linear-gradient(135deg,rgba(37,99,235,.22),rgba(124,58,237,.22));
  border:1px solid rgba(255,255,255,.16);
  color:#E5E7EB;
  border-radius:16px;
  padding:14px;
  margin:14px 0 18px;
  line-height:1.65;
}
.device-status-card b{color:#FACC15}
.premium-confirm-box{
  margin-top:16px;
  padding:16px;
  border-radius:18px;
  background:linear-gradient(135deg,#EEF2FF,#F5F3FF);
  border:1px solid #DDD6FE;
}
.premium-confirm-box h3{margin:0 0 8px;color:#1E1B4B}
.premium-confirm-box input,.premium-confirm-box textarea{background:#fff!important;color:#111827!important}
.premium-confirm-box button{width:100%;margin-top:8px}
.premium-request-status{margin-top:10px;font-weight:800;color:#047857;line-height:1.5}

/* Trial Experience Premium Card - fixed readable design */
.trial-experience-card{
  position:relative;
  overflow:hidden;
  background:linear-gradient(180deg,#FFFFFF 0%,#F8FAFF 100%)!important;
  border:1px solid rgba(37,99,235,.16)!important;
  box-shadow:0 24px 70px rgba(15,23,42,.13)!important;
}
.trial-ribbon{background:linear-gradient(135deg,#2563EB,#7C3AED)!important;color:#fff!important}
.trial-badge-pro{
  font-size:24px;
  font-weight:900;
  color:#1E1B4B;
  margin:12px 0 6px;
  letter-spacing:-.5px;
}
.trial-days-pro{
  font-size:44px;
  line-height:1.05;
  font-weight:950;
  background:linear-gradient(135deg,#F59E0B,#FACC15);
  -webkit-background-clip:text;
  -webkit-text-fill-color:transparent;
  margin:8px 0 18px;
}
.trial-features-pro{margin:8px 0 18px}
.feature-open-pro,.feature-lock-pro{
  display:flex;
  align-items:center;
  min-height:38px;
  padding:8px 0;
  border-bottom:1px dashed rgba(37,99,235,.16);
  font-size:17px;
  line-height:1.35;
  font-weight:850;
}
.feature-open-pro{color:#047857!important}
.feature-lock-pro{color:#DC2626!important}
.trial-upgrade-pro{
  margin-top:18px;
  padding:14px 15px;
  border-radius:16px;
  text-align:center;
  font-weight:900;
  line-height:1.45;
  color:#1E3A8A;
  background:linear-gradient(135deg,#DBEAFE,#EEF2FF);
  border:1px solid rgba(37,99,235,.16);
}
.trial-upgrade-btn{margin-top:14px!important;width:100%}
.trial-side-card h3{color:#1E1B4B!important;font-weight:950}
.trial-side-days{
  font-size:28px;
  font-weight:950;
  margin:8px 0 10px;
  color:#F59E0B;
}
.trial-side-list{
  padding:12px 0;
  border-bottom:1px dashed rgba(37,99,235,.15);
  font-weight:850;
  line-height:1.8;
}
.trial-side-list.open{color:#047857!important}
.trial-side-list.locked{color:#DC2626!important}
.trial-side-upgrade{
  margin:14px 0;
  padding:12px;
  border-radius:14px;
  background:linear-gradient(135deg,#DBEAFE,#EEF2FF);
  color:#1E3A8A;
  font-weight:900;
  line-height:1.45;
}


/* Facebook Publisher Pro + Smart Engagement */
.fb-pro-badge{
  display:inline-flex;align-items:center;gap:8px;
  padding:8px 14px;border-radius:999px;
  background:linear-gradient(135deg,#DBEAFE,#EEF2FF);
  color:#1E3A8A;font-weight:950;font-size:13px;
  border:1px solid rgba(37,99,235,.18);margin-bottom:12px;
}
.fb-pro-card{
  position:relative;overflow:hidden;
  background:linear-gradient(180deg,rgba(255,255,255,.96),rgba(248,250,252,.92));
  border:1px solid rgba(37,99,235,.12);
  border-radius:24px;padding:22px;
  box-shadow:0 18px 45px rgba(15,23,42,.08);
}
.fb-pro-card h3{margin-top:0;color:#1E1B4B;font-size:21px;font-weight:950;letter-spacing:-.4px}
.fb-pro-card p{color:#64748B;line-height:1.7;font-weight:650}
.fb-pro-card ul{padding-left:18px;color:#334155;line-height:1.9;font-weight:750}
.fb-pro-card li::marker{color:#2563EB}
.fb-safe-note{
  padding:16px 18px;border-radius:18px;
  background:linear-gradient(135deg,#F0FDF4,#ECFDF5);
  border:1px solid rgba(16,185,129,.22);
  color:#065F46;font-weight:850;line-height:1.65;margin:16px 0;
}
.fb-danger-note{
  padding:16px 18px;border-radius:18px;
  background:linear-gradient(135deg,#FEF2F2,#FFF7ED);
  border:1px solid rgba(239,68,68,.18);
  color:#991B1B;font-weight:850;line-height:1.65;margin:16px 0;
}
.smart-chip-row{display:flex;flex-wrap:wrap;gap:10px;margin:14px 0 18px}
.smart-chip{
  border:1px solid rgba(37,99,235,.18);
  background:#fff;color:#1E3A8A;font-weight:900;
  padding:10px 13px;border-radius:999px;font-size:13px;
}

</style>

</head>
<body>

<div class="v2-topbar">
  <div class="v2-brand-row">
    <div>
      <div class="v2-brand-title">🚀 Mkt Automation Pro V5</div>
      <div class="v2-brand-sub">AI Marketing • Facebook • CRM • Automation</div>
    </div>
    <div class="v2-status-pill">Online</div>
  </div>
</div>


<div class="live-trust-bar">
  <span class="live-dot"></span>
  <span>Đã có <b id="liveMemberCount">231</b> khách hàng nâng cấp Premium</span>
</div>

<div class="layout">
<aside class="sidebar">
  <div class="logo">Marketing<br>Automation Pro</div>
  <div class="subtitle">V5 Enterprise Seller AI Suite</div>

  <div class="device-status-card">
    🖥 <b>ID thiết bị</b><br>
    <b class="device-id-text">Đang tạo...</b><br>
    <span>Trạng thái: <span class="premium-status-text">Dùng thử / chờ duyệt</span></span>
  </div>

<div class="nav">
  <div class="v2-nav-title">🏠 DASHBOARD CEO</div>
  <a class="v2-nav-link" href="#dashboard"><span class="v2-nav-ico">🏠</span><span class="v2-nav-text">Dashboard CEO</span><span class="v2-nav-tag">Home</span></a>
  <a class="v2-nav-link" href="#analytics" onclick="openModule('analytics')"><span class="v2-nav-ico">📌</span><span class="v2-nav-text">Thống kê nhanh</span></a>
  <a class="v2-nav-link" href="#history" onclick="openModule('history')"><span class="v2-nav-ico">🔥</span><span class="v2-nav-text">Hoạt động hôm nay</span></a>

  <div class="v2-nav-title">📣 FACEBOOK CENTER</div>
  <a class="v2-nav-link" href="#facebook_center" onclick="openModule('facebook_center')"><span class="v2-nav-ico">📣</span><span class="v2-nav-text">Facebook Center</span><span class="v2-nav-tag">Core</span></a>
  <a class="v2-nav-link" href="#post" onclick="openModule('post')"><span class="v2-nav-ico">📝</span><span class="v2-nav-text">Đăng bài Facebook</span></a>
  <a class="v2-nav-link" href="#fanpage_manager" onclick="openModule('fanpage_manager')"><span class="v2-nav-ico">📄</span><span class="v2-nav-text">Quản lý Fanpage</span><span class="v2-nav-tag">V5</span></a>
  <a class="v2-nav-link" href="#group_marketing" onclick="openModule('group_marketing')"><span class="v2-nav-ico">👥</span><span class="v2-nav-text">Quản lý Group</span><span class="v2-nav-tag">V5</span></a>
  <a class="v2-nav-link" href="#token" onclick="openModule('token')"><span class="v2-nav-ico">🔑</span><span class="v2-nav-text">Token Fanpage</span><span class="v2-nav-tag">Pro</span></a>
  <a class="v2-nav-link" href="#facebook_publisher_pro" onclick="openModule('facebook_publisher_pro')"><span class="v2-nav-ico">🚀</span><span class="v2-nav-text">Đăng bài Fanpage</span><span class="v2-nav-tag">Pro</span></a>
  <a class="v2-nav-link" href="#scheduler" onclick="openModule('scheduler')"><span class="v2-nav-ico">🗓️</span><span class="v2-nav-text">Lên lịch đăng bài</span></a>
  <a class="v2-nav-link" href="#history" onclick="openModule('history')"><span class="v2-nav-ico">📜</span><span class="v2-nav-text">Nhật ký đăng bài</span></a>
  <a class="v2-nav-link" href="#smart_engagement" onclick="openModule('smart_engagement')"><span class="v2-nav-ico">⚡</span><span class="v2-nav-text">Tăng tương tác thông minh</span><span class="v2-nav-tag">Safe</span></a>
  <a class="v2-nav-link" href="#comment_manager" onclick="openModule('comment_manager')"><span class="v2-nav-ico">💬</span><span class="v2-nav-text">AI trả lời comment</span></a>

  <div class="v2-nav-title">🤖 AI BÁN HÀNG</div>
  <a class="v2-nav-link" href="#comment_manager" onclick="openModule('comment_manager')"><span class="v2-nav-ico">🤖</span><span class="v2-nav-text">AI Comment</span><span class="v2-nav-tag">AI</span></a>
  <a class="v2-nav-link" href="#messenger_ai" onclick="openModule('messenger_ai')"><span class="v2-nav-ico">💬</span><span class="v2-nav-text">AI Messenger</span><span class="v2-nav-tag">AI</span></a>
  <a class="v2-nav-link" href="#crm_sales" onclick="openModule('crm_sales')"><span class="v2-nav-ico">📋</span><span class="v2-nav-text">CRM Kanban</span><span class="v2-nav-tag">CRM</span></a>
  <a class="v2-nav-link" href="#messenger_ai" onclick="openModule('messenger_ai')"><span class="v2-nav-ico">🧾</span><span class="v2-nav-text">Kịch bản chốt sale</span></a>
  <a class="v2-nav-link" href="#crm_sales" onclick="openModule('crm_sales')"><span class="v2-nav-ico">♻️</span><span class="v2-nav-text">Chăm sóc khách cũ</span></a>

  <div class="v2-nav-title">🧠 MARKETING DIRECTOR</div>
  <a class="v2-nav-link" href="#marketing_director" onclick="openModule('marketing_director')"><span class="v2-nav-ico">🧠</span><span class="v2-nav-text">AI Marketing Director</span><span class="v2-nav-tag">HOT</span></a>
  <a class="v2-nav-link" href="#plan" onclick="openModule('plan')"><span class="v2-nav-ico">📅</span><span class="v2-nav-text">Kế hoạch 30 ngày</span></a>
  <a class="v2-nav-link" href="#marketing_director" onclick="openModule('marketing_director')"><span class="v2-nav-ico">🧲</span><span class="v2-nav-text">Phễu bán hàng</span></a>
  <a class="v2-nav-link" href="#ai_studio" onclick="openModule('ai_studio')"><span class="v2-nav-ico">🔎</span><span class="v2-nav-text">Phân tích đối thủ</span></a>
  <a class="v2-nav-link" href="#scheduler" onclick="openModule('scheduler')"><span class="v2-nav-ico">🗓️</span><span class="v2-nav-text">Content Calendar</span></a>

  <div class="v2-nav-title">📊 ANALYTICS CENTER</div>
  <a class="v2-nav-link" href="#analytics" onclick="openModule('analytics')"><span class="v2-nav-ico">📊</span><span class="v2-nav-text">Analytics Center</span></a>
  <a class="v2-nav-link" href="#fanpage_manager" onclick="openModule('fanpage_manager')"><span class="v2-nav-ico">📈</span><span class="v2-nav-text">Báo cáo Fanpage</span></a>
  <a class="v2-nav-link" href="#group_marketing" onclick="openModule('group_marketing')"><span class="v2-nav-ico">👥</span><span class="v2-nav-text">Báo cáo Group</span></a>
  <a class="v2-nav-link" href="#crm_sales" onclick="openModule('crm_sales')"><span class="v2-nav-ico">💼</span><span class="v2-nav-text">Báo cáo CRM</span></a>
  <a class="v2-nav-link" href="/export"><span class="v2-nav-ico">📤</span><span class="v2-nav-text">Xuất Excel/CSV</span></a>
  <a class="v2-nav-link" href="/export_pdf"><span class="v2-nav-ico">📄</span><span class="v2-nav-text">Xuất PDF</span></a>

  <div class="v2-nav-title">💎 PREMIUM CENTER</div>
  <a class="v2-nav-link" href="#premium" onclick="openModule('premium')"><span class="v2-nav-ico">💎</span><span class="v2-nav-text">Bảng giá Premium</span><span class="v2-nav-tag">VIP</span></a>
  <a class="v2-nav-link" href="#premium" onclick="openModule('premium')"><span class="v2-nav-ico">💳</span><span class="v2-nav-text">Gửi xác nhận thanh toán</span></a>
  <a class="v2-nav-link" href="#premium" onclick="openModule('premium')"><span class="v2-nav-ico">✅</span><span class="v2-nav-text">Trạng thái kích hoạt</span></a>
  <a class="v2-nav-link" href="/install" target="_blank"><span class="v2-nav-ico">📲</span><span class="v2-nav-text">Cài đặt App</span><span class="v2-nav-tag">App</span></a>

  <div class="v2-nav-title">⚙️ TÀI KHOẢN</div>
  <a class="v2-nav-link" href="#premium" onclick="openModule('premium')"><span class="v2-nav-ico">🖥️</span><span class="v2-nav-text">ID thiết bị</span></a>
  <a class="v2-nav-link" href="#premium" onclick="openModule('premium')"><span class="v2-nav-ico">📧</span><span class="v2-nav-text">Gmail / SĐT</span></a>
  <a class="v2-nav-link" href="#premium" onclick="openModule('premium')"><span class="v2-nav-ico">👑</span><span class="v2-nav-text">Trạng thái gói</span></a>

  <div class="v2-side-card">
    🚀 Mkt Automation Pro V6<br>
    Fanpage • Group • Comment • Messenger • CRM • Marketing Director • Analytics • Premium.
  </div>
</div>
</aside>

<main class="main">

<section class="top-hero" id="dashboard">
  <h1>Mkt Automation Pro V5 Seller AI Suite</h1>

<div class="app-install-banner">
  <b>📲 App Mini đã sẵn sàng</b><br>
  Bấm nút bên dưới để cài Mkt Automation Pro vào điện thoại/máy tính và mở như phần mềm riêng.
  <button onclick="showSmartInstall()">📲 Cài đặt GPT MKT Pro</button>
  <a href="/install" target="_blank" style="display:inline-block;margin-top:8px;font-size:13px;font-weight:900;color:#4C1D95;text-decoration:underline">Mở trang cài đặt GPT MKT Pro</a>
  <div id="installStatus" style="margin-top:8px;font-size:13px;color:#4C1D95;font-weight:800"></div>
</div>

<div class="app-quick-grid">
  <div class="app-quick-card" onclick="openModule('post')">
    <div class="app-ico">📢</div>
    <b>Đăng bài Facebook</b>
    <span>Soạn nội dung, chọn Page, đăng ngay hoặc lên lịch.</span>
  </div>
  <div class="app-quick-card" onclick="openModule('ai_studio')">
    <div class="app-ico">🤖</div>
    <b>Tạo Content AI</b>
    <span>Viết bài, caption, ý tưởng quảng cáo và nội dung bán hàng.</span>
  </div>
  <div class="app-quick-card" onclick="openModule('fanpage_manager')">
    <div class="app-ico">📄</div>
    <b>Quản lý Fanpage</b>
    <span>Kiểm tra Page, Token, quyền đăng bài và trạng thái hoạt động.</span>
  </div>
  <div class="app-quick-card" onclick="openModule('group_marketing')">
    <div class="app-ico">👥</div>
    <b>Quản lý Group</b>
    <span>Lưu group, tạo lịch đăng group và viết bài seeding mềm.</span>
  </div>
  <div class="app-quick-card" onclick="openModule('comment_manager')">
    <div class="app-ico">💬</div>
    <b>AI Comment</b>
    <span>Ẩn số điện thoại, trả lời comment, gắn nhãn và chuyển CRM.</span>
  </div>
  <div class="app-quick-card" onclick="openModule('messenger_ai')">
    <div class="app-ico">📨</div>
    <b>AI Messenger</b>
    <span>Tạo kịch bản inbox, chốt sale, xử lý từ chối và chăm sóc lại.</span>
  </div>
  <div class="app-quick-card" onclick="openModule('crm_sales')">
    <div class="app-ico">📊</div>
    <b>CRM Kanban</b>
    <span>Quản lý khách theo các cột: mới, tư vấn, báo giá, chốt, đã mua.</span>
  </div>
  <div class="app-quick-card" onclick="openModule('marketing_director')">
    <div class="app-ico">🧠</div>
    <b>AI Marketing Director</b>
    <span>Lập kế hoạch 30 ngày, ads, content, funnel và KPI bán hàng.</span>
  </div>
  <div class="app-quick-card" onclick="location.href='#premium'">
    <div class="app-ico">💎</div>
    <b>Premium</b>
    <span>Mở khóa hạn mức cao hơn và tính năng nâng cao.</span>
  </div>
</div>

  <p>AI Marketing Automation Platform: tạo content, đăng Fanpage, chia lịch, quản lý CRM, phân tích đối thủ, tạo funnel và hỗ trợ bán hàng tự động.</p>

  {% if message %}
  <div class="{{ 'success' if ok else 'error' }}"><b>Thông báo:</b> {{ message }}</div>
  {% endif %}

  <div class="v3-kpi-title"><b>KPI tổng quan</b> • Fanpage đang hoạt động • Khách hàng mới hôm nay • Bài đăng hôm nay • Doanh thu tháng • Trạng thái Token</div>
  <div class="grid4 v3-ceo-grid" style="margin-top:18px">
    <div class="stat"><span>Fanpage</span><br><b>{{ v3.fanpages }}</b><small>đang hoạt động/cấu hình</small></div>
    <div class="stat"><span>Bài đăng</span><br><b>{{ v3.posts }}</b><small>hôm nay: {{ v3.posts_today }}</small></div>
    <div class="stat"><span>Khách hàng</span><br><b>{{ v3.crm }}</b><small>mới hôm nay: {{ v3.leads_today }}</small></div>
    <div class="stat"><span>Doanh thu</span><br><b>{{ v3.revenue_month }}</b><small>Chưa phát sinh giao dịch</small></div>
    <div class="stat premium-trial-stat"><span>Premium</span><br><b>🎁 Dùng thử miễn phí</b><small>Còn lại: 3 ngày • 10/10 Content AI • 1/1 Fanpage</small></div>
    <div class="stat token-status-stat"><span>Token</span><br><b>🟢 {{ v3.token_live }} sống</b><small>🔴 {{ v3.token_dead }} lỗi • Tổng Page: {{ v3.token_total }}</small></div>
  </div>

  <div class="v4-trust-grid">
    <div class="v4-trust-card"><b>2.381+</b><span>Khách hàng sử dụng</span></div>
    <div class="v4-trust-card"><b>120.000+</b><span>Content đã tạo</span></div>
    <div class="v4-trust-card"><b>98%</b><span>Đánh giá hài lòng</span></div>
  </div>

  <div class="hero-actions">
    <button onclick="openModule('post')">Bắt đầu đăng bài</button>
    <button class="secondary" onclick="openModule('ai_studio')">Tạo content AI</button>
    <button class="secondary" onclick="scrollToPricing()">Xem bảng giá</button>
  </div>

  <div class="module-hub v3-main-hub">
    <div class="module-card" onclick="openModule('facebook_center')"><div class="icon">📣</div><h3>Facebook Center</h3><p>Trung tâm đăng bài, lên lịch, quản lý Fanpage, Group, Token và lịch sử đăng.</p><span class="module-pill">Mở Facebook Center</span></div>
    <div class="module-card" onclick="openModule('smart_engagement')"><div class="icon">⚡</div><h3>Tăng tương tác thông minh</h3><p>Gợi ý giờ đăng, caption kéo comment, câu hỏi tương tác và nhắc chăm sóc khách đã comment.</p><span class="module-pill">Mở Smart Engagement</span></div>
    <div class="module-card" onclick="openModule('fanpage_manager')"><div class="icon">📄</div><h3>Quản lý Fanpage</h3><p>Kiểm tra Page, Token, quyền đăng bài và trạng thái hoạt động.</p><span class="module-pill">Mở Fanpage</span></div>
    <div class="module-card" onclick="openModule('group_marketing')"><div class="icon">👥</div><h3>Quản lý Group</h3><p>Lưu Group, tạo lịch đăng Group và viết bài seeding mềm.</p><span class="module-pill">Mở Group</span></div>
    <div class="module-card" onclick="openModule('comment_manager')"><div class="icon">🤖</div><h3>AI Comment</h3><p>Ẩn số điện thoại, trả lời bình luận, gắn nhãn khách nóng/ấm/lạnh và chuyển CRM.</p><span class="module-pill">Mở AI Comment</span></div>
    <div class="module-card" onclick="openModule('messenger_ai')"><div class="icon">💬</div><h3>AI Messenger</h3><p>Tạo kịch bản inbox, chốt sale, xử lý từ chối và chăm sóc lại.</p><span class="module-pill">Mở AI Messenger</span></div>
    <div class="module-card" onclick="openModule('crm_sales')"><div class="icon">📋</div><h3>CRM Kanban</h3><p>Quản lý khách theo các cột: mới, tư vấn, báo giá, theo dõi, đã chốt.</p><span class="module-pill">Mở CRM</span></div>
    <div class="module-card" onclick="openModule('marketing_director')"><div class="icon">🧠</div><h3>AI Marketing Director</h3><p>Lập kế hoạch 30 ngày, Ads, Content, Funnel, KPI và chiến lược tăng doanh thu.</p><span class="module-pill">Mở Marketing Director</span></div>
    <div class="module-card" onclick="openModule('ai_studio')"><div class="icon">🎨</div><h3>AI Studio</h3><p>Gộp AI Facebook, AI Image, AI Video, AI Giọng Nói và AI Livestream vào một khu vực.</p><span class="module-pill">Mở AI Studio</span></div>
    <div class="module-card" onclick="openModule('premium')"><div class="icon">💎</div><h3>Premium</h3><p>Mở khóa hạn mức cao hơn, module nâng cao và hỗ trợ ưu tiên.</p><span class="module-pill">Xem gói</span></div>
  </div>
</section>

  {% if content %}
  <div class="panel" id="ai-output-box">
    <h2>Kết quả AI vừa tạo</h2>
    <p class="small">Nội dung AI tạo sẽ hiển thị tại đây để khách xem, copy hoặc đưa sang phần đăng Fanpage.</p>
    <div class="history" id="aiGeneratedContent">{{ content }}</div>
    <button onclick="copyText('aiGeneratedContent')">Copy nội dung</button>
    <button class="secondary" onclick="openModule('post')">Mở phần đăng Fanpage</button>
  </div>
  {% endif %}



<section class="panel module-section" id="facebook_center">
  <div class="section-open-note">Bạn đang mở: Facebook Center Pro.</div>
  <h2>📣 Facebook Center</h2>
  <p class="small">Trung tâm đăng bài Fanpage, lên lịch, quản lý nội dung, nhật ký đăng bài, tăng tương tác thông minh và AI trả lời comment.</p>
  <div class="fb-submenu-pro">
    <button onclick="openModule('facebook_publisher_pro')">🚀 Đăng bài Fanpage</button>
    <button onclick="openModule('scheduler')">🗓️ Lên lịch đăng bài</button>
    <button onclick="openModule('library')">🗂️ Quản lý nội dung</button>
    <button onclick="openModule('history')">📜 Nhật ký đăng bài</button>
    <button onclick="openModule('smart_engagement')">⚡ Tăng tương tác thông minh</button>
    <button onclick="openModule('comment_manager')">💬 AI trả lời comment</button>
  </div>

  <div class="fb-safe-note">
    Hệ thống không dùng auto like hàng loạt, không tương tác giả bằng nhiều tài khoản. Thay vào đó dùng chiến lược tăng tương tác thông minh: giờ đăng tốt, caption kéo comment, câu hỏi tương tác, theo dõi bài yếu và nhắc admin chăm sóc khách thật.
  </div>

  <div class="v3-feature-grid">
    <div class="fb-pro-card"><div class="fb-pro-badge">🚀 Facebook Publisher Pro</div><h3>Đăng bài Fanpage chuyên nghiệp</h3><ul><li>Chọn Fanpage</li><li>Soạn bài</li><li>Upload ảnh/video</li><li>Đăng ngay</li><li>Lên lịch đăng</li><li>Lưu nháp</li><li>Nhật ký đăng bài</li><li>Báo lỗi token/page rõ ràng</li></ul><button onclick="openModule('facebook_publisher_pro')">Mở Publisher Pro</button></div>
    <div class="fb-pro-card"><div class="fb-pro-badge">💬 Engagement Assistant</div><h3>Chăm sóc tương tác thật</h3><ul><li>Gợi ý phản hồi comment</li><li>Lọc comment cần chăm sóc</li><li>Gắn nhãn khách quan tâm</li><li>Chuyển comment sang CRM</li><li>Gợi ý bài nên tương tác lại</li><li>Nhắc admin vào Facebook xử lý tương tác</li></ul><button onclick="openModule('comment_manager')">Mở AI Comment</button></div>
    <div class="fb-pro-card"><div class="fb-pro-badge">⚡ Safe Growth</div><h3>Tăng tương tác thông minh</h3><ul><li>Gợi ý giờ đăng tốt</li><li>Gợi ý caption tăng comment</li><li>Tạo câu hỏi kéo tương tác</li><li>Gợi ý trả lời comment</li><li>Theo dõi bài ít tương tác</li><li>Nhắc chăm sóc khách đã comment</li></ul><button onclick="openModule('smart_engagement')">Mở Smart Engagement</button></div>
  </div>

  <div class="fb-danger-note">
    Không tích hợp: tự động like hàng loạt, tự động like bằng nhiều tài khoản, tự động tương tác để tăng reach giả. Đây là nhóm hành vi dễ checkpoint, giảm uy tín Page và rủi ro khóa tài khoản.
  </div>
</section>

<section class="panel module-section" id="facebook_publisher_pro">
  <div class="section-open-note">Bạn đang mở: Facebook Publisher Pro.</div>
  <h2>🚀 Facebook Publisher Pro</h2>
  <p class="small">Đăng bài Fanpage trực tiếp, tải ảnh/video, lên lịch và theo dõi lỗi token/page rõ ràng.</p>
  <div class="v3-feature-grid">
    <div class="fb-pro-card"><h3>Quy trình đăng bài</h3><ul><li>Chọn Fanpage từ Token Center</li><li>Soạn nội dung hoặc lấy nội dung AI</li><li>Upload ảnh/video</li><li>Đăng ngay hoặc lên lịch</li><li>Lưu lịch sử đăng bài</li></ul><button onclick="openModule('post')">Mở form đăng bài</button></div>
    <div class="fb-pro-card"><h3>Lên lịch & lưu nháp</h3><p>Quản lý lịch đăng, chiến dịch nội dung và kiểm tra bài sắp đăng để tránh lỗi nội dung trùng lặp.</p><button onclick="openModule('scheduler')">Mở lịch đăng</button></div>
    <div class="fb-pro-card"><h3>Báo lỗi rõ ràng</h3><ul><li>Thiếu Page Token</li><li>Thiếu Page ID</li><li>Token chết / hết phiên</li><li>Thiếu quyền pages_manage_posts</li><li>Page chưa đủ quyền đăng</li></ul><button onclick="openModule('token')">Kiểm tra Token</button></div>
  </div>
</section>

<section class="panel module-section" id="smart_engagement">
  <div class="section-open-note">Bạn đang mở: Tăng tương tác thông minh.</div>
  <h2>⚡ Tăng tương tác thông minh</h2>
  <p class="small">Thay thế an toàn cho auto like: tăng tương tác bằng nội dung, giờ đăng, CTA và chăm sóc comment thật.</p>
  <div class="smart-chip-row"><span class="smart-chip">Gợi ý giờ đăng tốt</span><span class="smart-chip">Caption tăng comment</span><span class="smart-chip">Câu hỏi kéo tương tác</span><span class="smart-chip">Gợi ý trả lời comment</span><span class="smart-chip">Theo dõi bài ít tương tác</span><span class="smart-chip">Nhắc chăm sóc khách</span></div>
  <form method="post" action="/v3_ai_tool">
    <input type="hidden" name="tool" value="smart_engagement">
    <textarea name="topic" rows="4" placeholder="Nhập sản phẩm/dịch vụ hoặc dán nội dung bài viết cần tăng tương tác. Ví dụ: Dịch vụ chạy quảng cáo Facebook cho shop online"></textarea>
    <textarea name="extra" rows="3" placeholder="Thông tin thêm: tệp khách, ưu đãi, giờ đăng hiện tại, vấn đề đang gặp..."></textarea>
    <button>AI tạo kế hoạch tăng tương tác</button>
    <button type="button" class="secondary" onclick="openModule('comment_manager')">Mở AI trả lời comment</button>
  </form>
  <div class="fb-safe-note">Mục này chỉ tạo gợi ý và nhắc admin xử lý tương tác thật. Không tự động like hàng loạt, không dùng nhiều tài khoản, không tạo reach giả.</div>
</section>

<section class="panel module-section" id="fanpage_manager">
  <div class="section-open-note">Bạn đang mở: Fanpage Manager Pro.</div>
  <h2>📄 Quản lý Fanpage</h2>
  <p class="small">Trung tâm kiểm tra kết nối Page, Token, quyền và trạng thái hoạt động trước khi đăng bài.</p>
  <div class="v5-seller-grid">
    <div class="v5-tool-card"><h3>Kết nối Fanpage</h3><p>Thêm Fanpage trực tiếp trên web, không cần vào Render sửa PAGES_JSON.</p><button onclick="openModule('token')">Mở Token Center</button></div>
    <div class="v5-tool-card"><h3>Kiểm tra Token</h3><p>Quét toàn bộ Page để phát hiện token chết, thiếu quyền hoặc phiên bị giới hạn.</p><form method="post" action="/check_tokens"><button>Kiểm tra Token ngay</button></form></div>
    <div class="v5-tool-card"><h3>Kiểm tra quyền</h3><ul><li>pages_manage_posts</li><li>pages_read_engagement</li><li>pages_manage_metadata</li></ul><span class="v5-warning-pill">Cần kiểm tra từ Meta App</span></div>
    <div class="v5-tool-card"><h3>Làm mới Token</h3><p>Khi token lỗi, dán token mới vào Token Center rồi bấm lưu, không cần deploy lại.</p><button class="secondary" onclick="openModule('token')">Cập nhật Token</button></div>
  </div>
  <table class="v5-table"><tr><th>Fanpage</th><th>Page ID</th><th>Token</th><th>Nguồn</th><th>Trạng thái</th></tr>
    {% for p in pages %}<tr><td>{{ p.name }}</td><td>{{ p.id }}</td><td>{{ p.token_mask }}</td><td>{{ 'Tool' if p.source == 'database' else '.env' }}</td><td><span class="v5-status-pill">{{ 'Đã lưu' if p.token else 'Thiếu token' }}</span></td></tr>{% endfor %}
  </table>
</section>

<section class="panel module-section" id="group_marketing">
  <div class="section-open-note">Bạn đang mở: Group Marketing.</div>
  <h2>👥 Quản lý Group</h2>
  <p class="small">Quản lý Group, danh sách Group, lịch đăng Group và AI viết bài Group.</p>
  <div class="grid">
    <form method="post" action="/fb_group"><h3>Thêm Group</h3><input name="group_name" placeholder="Tên Group"><input name="group_id" placeholder="Group ID"><input name="niche" placeholder="Ngành / tệp khách"><textarea name="note" rows="2" placeholder="Ghi chú"></textarea><button>Lưu Group</button></form>
    <form method="post" action="/v3_ai_tool"><h3>AI viết bài Group</h3><input type="hidden" name="tool" value="group_content"><textarea name="topic" rows="4" placeholder="Ví dụ: Tôi bán Proxy cho người chạy quảng cáo Facebook"></textarea><button>Tạo bài Group</button></form>
  </div>
  <form method="post" action="/group_schedule"><h3>Lịch đăng Group</h3><div class="grid"><input name="group_name" placeholder="Tên Group"><input name="group_id" placeholder="Group ID"><input name="schedule_time" type="datetime-local"></div><textarea name="content" rows="4" placeholder="Nội dung cần lên lịch đăng Group"></textarea><button>Lưu lịch Group</button></form>
  <h3>Danh sách Group</h3>{% for g in fb_groups %}<div class="history"><b>{{ g[1] }}</b> • {{ g[2] }} • {{ g[3] }}<br>{{ g[4] }}</div>{% endfor %}
  <h3>Lịch Group gần nhất</h3>{% for gs in group_schedules %}<div class="history"><b>{{ gs[1] }}</b> • {{ gs[4] }} • {{ gs[5] }}<br>{{ gs[3] }}</div>{% endfor %}
</section>

<section class="panel module-section" id="comment_manager">
  <div class="section-open-note">Bạn đang mở: Comment Manager AI.</div>
  <h2>💬 AI Comment</h2>
  <p class="small">AI trả lời comment, ẩn SĐT, gắn nhãn khách và chuyển CRM.</p>
  <form method="post" action="/comment_ai">
    <div class="grid"><input name="customer_name" placeholder="Tên khách nếu có"><input name="phone" placeholder="SĐT nếu có"></div>
    <textarea name="comment_text" rows="4" placeholder="Dán comment khách hàng vào đây. Ví dụ: shop còn proxy Việt không, số em 09..."></textarea>
    <div class="grid"><select name="label"><option>Khách nóng</option><option>Khách ấm</option><option>Khách lạnh</option><option>Cần chăm sóc lại</option></select><select name="to_crm"><option value="1">Chuyển sang CRM</option><option value="0">Chỉ lưu comment</option></select></div>
    <button>AI xử lý Comment</button><button type="button" class="secondary" onclick="openLockedFeature('Comment Manager AI','Gói 1 năm / Gói Nhà Bán Hàng Chuyên Nghiệp')">Xem bản tự động Webhook</button>
  </form>
  <h3>Comment đã xử lý</h3>{% for c in comment_leads %}<div class="history"><b>{{ c[1] or 'Khách hàng' }}</b> • {{ c[5] }} • {{ c[7] }}<br>Comment: {{ c[3] }}<br>AI: {{ c[4] }}</div>{% endfor %}
</section>

<section class="panel module-section" id="messenger_ai">
  <div class="section-open-note">Bạn đang mở: Messenger AI.</div>
  <h2>📨 AI Messenger</h2>
  <p class="small">Tạo kịch bản Inbox, kịch bản chốt sale, xử lý từ chối và chăm sóc khách cũ.</p>
  <form method="post" action="/messenger_ai_script">
    <div class="grid"><input name="product" placeholder="Sản phẩm/dịch vụ. Ví dụ: Proxy"><select name="script_type"><option>Kịch bản Inbox</option><option>Kịch bản Chốt Sale</option><option>Xử lý từ chối</option><option>Chăm sóc khách cũ</option></select></div>
    <textarea name="extra" rows="3" placeholder="Thông tin thêm: giá, ưu đãi, tệp khách, số Zalo..."></textarea>
    <button>Tạo kịch bản Messenger</button>
  </form>
  <h3>Kịch bản đã lưu</h3>{% for m in messenger_scripts %}<div class="history"><b>{{ m[1] }}</b> • {{ m[2] }} • {{ m[4] }}<br>{{ m[3] }}</div>{% endfor %}
</section>

<section class="panel module-section" id="ai_studio">
  <div class="section-open-note">Bạn đang mở: AI Studio V3.</div>
  <h2>🤖 AI Studio</h2>
  <div class="v3-feature-grid">
    <div class="v3-feature-card"><h3>AI Content</h3><ul><li>Content Facebook</li><li>Content TikTok</li><li>Caption</li></ul><button onclick="openModule('ai_studio')">Mở AI Content</button></div>
    <div class="v3-feature-card"><h3>Viral Content Lab</h3><ul><li>Content Viral</li><li>Storytelling</li><li>Seeding</li></ul><form method="post" action="/v3_ai_tool"><input type="hidden" name="tool" value="prompt_premium"><textarea name="topic" rows="3" placeholder="Ví dụ: nội dung viral cho dịch vụ proxy"></textarea><button>Tạo ý tưởng viral</button></form></div>
    <div class="v3-feature-card"><h3>Facebook Ads AI</h3><ul><li>Chấm điểm quảng cáo</li><li>Viết quảng cáo</li><li>Target khách hàng</li></ul><form method="post" action="/v3_ai_tool"><input type="hidden" name="tool" value="facebook_ads_ai"><textarea name="topic" rows="3" placeholder="Nhập sản phẩm/dịch vụ cần chạy ads"></textarea><button>Tạo Facebook Ads AI</button></form></div>
    <div class="v3-feature-card"><span class="v3-premium-badge">VIP</span><h3>AI Marketing Director</h3><p>Khách nhập: Tôi bán Proxy. AI trả: 30 Content, 10 Quảng cáo, 10 Caption, 5 Kịch bản chốt sale, 30 ngày Marketing, tệp khách hàng và ngân sách đề xuất.</p><form method="post" action="/v3_ai_tool"><input type="hidden" name="tool" value="marketing_director"><textarea name="topic" rows="3" placeholder="Tôi bán Proxy / mỹ phẩm / dịch vụ quảng cáo..."></textarea><button>Tạo kế hoạch tổng</button></form></div>
    <div class="v3-feature-card"><h3>Competitor Scanner</h3><p>Nhập link Fanpage hoặc mô tả đối thủ. AI phân tích điểm mạnh, điểm yếu, nội dung và CTA.</p><form method="post" action="/v3_ai_tool"><input type="hidden" name="tool" value="competitor_scanner"><input name="topic" placeholder="Link Fanpage đối thủ"><button>Phân tích đối thủ</button></form></div>
    <div class="v3-feature-card"><h3>AI Sales Script</h3><ul><li>Kịch bản chốt sale</li><li>Xử lý từ chối</li><li>Kịch bản inbox</li></ul><form method="post" action="/v3_ai_tool"><input type="hidden" name="tool" value="sales_script"><textarea name="topic" rows="3" placeholder="Sản phẩm/dịch vụ cần tư vấn"></textarea><button>Tạo kịch bản sale</button></form></div>
    <div class="v3-feature-card"><h3>AI Landing Page Builder</h3><p>Nhập: Dịch vụ chạy quảng cáo Facebook. AI sinh tiêu đề, CTA, form và ưu điểm.</p><form method="post" action="/v3_ai_tool"><input type="hidden" name="tool" value="landing_page"><textarea name="topic" rows="3" placeholder="Dịch vụ chạy quảng cáo Facebook"></textarea><button>Tạo Landing Page</button></form></div>
  </div>
</section>

<section class="panel module-section" id="creative_center">
  <div class="section-open-note">Bạn đang mở: Creative Center.</div>
  <h2>🎨 Creative Center</h2>
  <div class="v3-feature-grid">
    <div class="v3-feature-card"><h3>AI Image Center</h3><ul><li>Avatar</li><li>Cover</li><li>Banner</li><li>Poster</li></ul><button class="secondary" onclick="openLockedFeature('AI Image Center','Gói 1 năm / Vĩnh viễn')">Mở khóa</button></div>
    <div class="v3-feature-card"><h3>AI Video Center</h3><ul><li>TikTok</li><li>Reels</li><li>Shorts</li></ul><button class="secondary" onclick="openLockedFeature('AI Video Center','Gói 1 năm / Vĩnh viễn')">Mở khóa</button></div>
    <div class="v3-feature-card"><h3>AI Voice Studio</h3><ul><li>Giọng nữ</li><li>Giọng nam</li><li>MC quảng cáo</li></ul><button class="secondary" onclick="openLockedFeature('AI Voice Studio','Gói 1 năm / Vĩnh viễn')">Mở khóa</button></div>
    <div class="v3-feature-card"><h3>Banner Builder</h3><p>Tạo ý tưởng ảnh quảng cáo, bố cục banner và nội dung poster.</p><form method="post" action="/v3_ai_tool"><input type="hidden" name="tool" value="landing_page"><textarea name="topic" rows="3" placeholder="Banner cho dịch vụ quảng cáo Facebook"></textarea><button>Tạo bố cục banner</button></form></div>
  </div>
</section>

<section class="panel module-section" id="marketing_director">
  <div class="section-open-note">Bạn đang mở: AI Marketing Director.</div>
  <h2>🧠 AI Marketing Director</h2>
  <p class="small">Trợ lý giám đốc Marketing: phân tích sản phẩm, tệp khách hàng, content, quảng cáo, phễu bán hàng, KPI và kế hoạch triển khai 30 ngày.</p>
  <div class="v3-feature-grid">
    <div class="v3-feature-card"><h3>Đầu vào cần nhập</h3><ul><li>Sản phẩm/dịch vụ đang bán</li><li>Khách hàng mục tiêu</li><li>Giá/ưu đãi hiện tại</li><li>Mục tiêu: inbox, đơn hàng, thương hiệu</li></ul></div>
    <div class="v3-feature-card"><h3>Kết quả AI trả về</h3><ul><li>30 content Facebook/TikTok</li><li>10 mẫu quảng cáo</li><li>5 kịch bản chốt sale</li><li>Kế hoạch 30 ngày và KPI</li></ul></div>
  </div>
  <form method="post" action="/v3_ai_tool">
    <input type="hidden" name="tool" value="marketing_director">
    <textarea name="topic" rows="4" placeholder="Ví dụ: Tôi bán proxy Việt Nam cho người chạy Facebook Ads, giá từ 80k/tháng, mục tiêu tăng inbox và chốt khách qua Zalo."></textarea>
    <textarea name="extra" rows="3" placeholder="Thông tin thêm: ngân sách ads, khu vực, ưu đãi, số Zalo, đối thủ, điểm mạnh sản phẩm..."></textarea>
    <button>🧠 Tạo kế hoạch Marketing Director</button>
    <button type="button" class="secondary" onclick="openModule('ai_studio')">Mở AI Studio</button>
    <button type="button" class="secondary" onclick="openModule('crm_sales')">Mở CRM Kanban</button>
  </form>
</section>

<section class="panel module-section" id="crm_sales">
  <div class="section-open-note">Bạn đang mở: CRM Sales V3.</div>
  <h2>📊 CRM Kanban</h2>
  <form method="post" action="/pipeline">
    <div class="grid"><input name="customer_name" placeholder="Tên"><input name="phone" placeholder="SĐT"><input name="zalo" placeholder="Zalo"><input name="source" placeholder="Nguồn"></div>
    <div class="grid"><select name="stage"><option>Khách mới</option><option>Đang tư vấn</option><option>Đã báo giá</option><option>Đang chốt</option><option>Đã mua</option></select><input name="value" type="number" placeholder="Doanh thu dự kiến"></div>
    <textarea name="note" rows="3" placeholder="Ghi chú nhu cầu khách"></textarea><button>Lưu vào Pipeline</button>
  </form>
  <h3>Pipeline Kanban</h3>
  <div class="kanban-board">
    {% for stage in ['Khách mới','Đang tư vấn','Đã báo giá','Đang chốt','Đã mua'] %}
    <div class="kanban-col" ondragover="event.preventDefault()" ondrop="dropKanban(event)"><h3>{{ stage }}</h3>
      {% for r in pipeline_rows %}{% if r[5] == stage %}<div class="kanban-card" draggable="true" ondragstart="dragKanban(event)"><b>{{ r[1] }}</b><br><span class="small">{{ r[2] }} • {{ r[4] }}</span><br>{{ r[7] }}</div>{% endif %}{% endfor %}
    </div>
    {% endfor %}
  </div>
  <h3>Customer Care</h3>
  <form method="post" action="/customer_task"><div class="grid"><input name="customer_name" placeholder="Tên khách"><input name="due_date" type="date"></div><textarea name="task" rows="2" placeholder="Lịch chăm sóc / nhắc lịch"></textarea><button>Tạo nhắc lịch</button></form>
  {% for t in customer_tasks %}<div class="history">{{ t[1] }} | {{ t[2] }} | Hạn: {{ t[3] }} | {{ t[4] }}</div>{% endfor %}
</section>

<section class="panel module-section" id="automation_center">
  <div class="section-open-note">Bạn đang mở: Automation Center.</div>
  <h2>⚙️ Automation</h2>
  <div class="v3-feature-grid">
    <div class="v3-feature-card"><h3>Auto Campaign</h3><p>Tạo 30 ngày content và tự động lên lịch.</p><form method="post" action="/plan_30_days"><select name="industry">{% for key,label in industry_labels.items() %}<option value="{{ key }}">{{ label }}</option>{% endfor %}</select><input name="goal" placeholder="Mục tiêu"><input type="hidden" name="days" value="30"><button>Tạo 30 ngày content</button></form></div>
    <div class="v3-feature-card"><h3>Notification Center</h3><ul><li>Bài sắp đăng</li><li>Token lỗi</li><li>Khách mới</li></ul><form method="post" action="/notification"><input name="title" placeholder="Tiêu đề thông báo"><textarea name="detail" rows="2" placeholder="Nội dung"></textarea><button>Lưu thông báo</button></form></div>
    <div class="v3-feature-card"><h3>Backup Center</h3><ul><li>Backup</li><li>Restore</li></ul><a href="/backup"><button>Backup</button></a><a href="/export_pdf"><button class="secondary">Xuất PDF</button></a><a href="/export"><button class="secondary">Xuất Excel/CSV</button></a></div>
  </div>
  <h3>Thông báo gần nhất</h3>{% for n in notifications %}<div class="history"><b>{{ n[1] }}</b> • {{ n[3] }} • {{ n[5] }}<br>{{ n[2] }}</div>{% endfor %}
</section>

<section class="panel module-section" id="success_center">
  <div class="section-open-note">Bạn đang mở: Success Center V4.</div>
  <h2>🏆 Success Center</h2>
  <p class="small">Tăng niềm tin cho khách trước khi nâng cấp Premium bằng case study, mẫu thành công và tài sản bán hàng đã được chuẩn hóa.</p>

  <div class="v4-success-grid">
    {% for a in success_assets %}
    <div class="v4-success-card">
      <div class="success-icon">🏆</div>
      <h3>{{ a[2] }}</h3>
      <p><b>{{ a[1] }}</b> — {{ a[3] }}</p>
      <button onclick="openLockedFeature('{{ a[1] }} Premium','Gói 1 năm / Gói Nhà Bán Hàng Chuyên Nghiệp')">Mở mẫu Premium</button>
    </div>
    {% endfor %}
  </div>

  <div class="v4-proof-box">
    <b>🔥 Thành tích hệ thống</b>
    <div class="v4-pricing-stats">
      <div><b>2.381+</b><span>Khách hàng sử dụng</span></div>
      <div><b>120.000+</b><span>Content đã tạo</span></div>
      <div><b>98%</b><span>Đánh giá hài lòng</span></div>
    </div>
  </div>
</section>

<section class="panel module-section" id="post">
  <div class="section-open-note">Bạn đang mở: Đăng Fanpage tự chia nội dung.</div>
  <h2>Đăng Fanpage Tự Chia Nội Dung</h2>
  <p class="small">Nhập nhiều content, upload nhiều ảnh/video, chọn nhiều Page. Hệ thống tự chia từng content + từng ảnh cho từng Page để tránh trùng nội dung và trùng hình. Mặc định không thêm chữ, không spin, không tự chèn CTA/hashtag.</p>

  <form method="post" action="/multi_post" enctype="multipart/form-data">
    <div class="page-list">
      {% for p in pages %}
      <label class="page-item">
        <input type="checkbox" name="page_indexes" value="{{ loop.index0 }}">
        {{ p.name }}<br>
        <span class="small">{{ p.id }}</span>
      </label>
      {% endfor %}
    </div>

    <textarea name="bulk_content" rows="12" placeholder="Content 1...

Content 2...

Content 3...

Mỗi content cách nhau bằng một dòng trống. Nếu nhập từng dòng, mỗi dòng cũng được hiểu là một content riêng."></textarea>

    <div class="grid">
      <input type="text" name="campaign" placeholder="Tên chiến dịch, ví dụ: Spa tháng 6">
      <input type="file" name="images" accept="image/*,video/mp4,video/quicktime" multiple>
    </div>
    <label class="page-item" style="display:block;margin-top:10px">
      <input type="checkbox" name="use_ai_enhance" value="1">
      Bật AI tối ưu từng bài: spin content, thêm CTA và hashtag riêng
      <br><span class="small">Mặc định tắt để giữ nguyên 100% content đã nhập.</span>
    </label>

    <div class="premium-center" style="margin-top:12px">
      <h3>Chọn cách đăng</h3>
      <p class="small">Đăng ngay hoặc đặt lịch đều dùng cùng nội dung, cùng Page và cùng ảnh/video đã chọn ở trên.</p>

      <div class="grid">
        <div>
          <h3>Đăng ngay</h3>
          <p class="small">Tự chia content + ảnh/video và đăng ngay lên các Page đã chọn.</p>
          <button type="submit" name="action" value="now">Tự chia và đăng ngay</button>
        </div>

        <div>
          <h3>Đặt lịch đăng</h3>
          <p class="small">Chọn thời gian, hệ thống lưu lịch và tự đăng khi đến giờ.</p>
          <input type="datetime-local" name="schedule_time">
          <button type="submit" name="action" value="schedule" class="secondary">Tự chia và lưu lịch</button>
        </div>
      </div>
    </div>
  </form>
</section>

<section class="panel module-section" id="scheduler">
  <div class="section-open-note">Bạn đang mở: Lịch đăng nâng cao.</div>
  <h2>Lịch đăng nâng cao</h2>
  <p class="small">Nhập nhiều content, upload nhiều ảnh/video, chọn Page và giờ bắt đầu. Hệ thống tự chia lịch đăng theo khung giờ, không trùng content và không trùng ảnh.</p>
  <form method="post" action="/smart_schedule" enctype="multipart/form-data">
    <div class="page-list">
      {% for p in pages %}
      <label class="page-item">
        <input type="checkbox" name="page_indexes" value="{{ loop.index0 }}">
        {{ p.name }}<br><span class="small">{{ p.id }}</span>
      </label>
      {% endfor %}
    </div>
    <textarea name="bulk_content" rows="9" placeholder="Content 1...

Content 2...

Content 3..."></textarea>
    <div class="grid">
      <input type="datetime-local" name="start_time">
      <select name="gap_minutes">
        <option value="30">Cách nhau 30 phút</option>
        <option value="60" selected>Cách nhau 1 giờ</option>
        <option value="120">Cách nhau 2 giờ</option>
        <option value="180">Cách nhau 3 giờ</option>
      </select>
    </div>
    <input name="campaign" placeholder="Tên chiến dịch">
    <input type="file" name="images" accept="image/*,video/mp4,video/quicktime" multiple>
    <label class="page-item" style="display:block;margin-top:10px">
      <input type="checkbox" name="use_ai_enhance" value="1">
      Bật AI tối ưu từng bài: spin content, thêm CTA và hashtag riêng
      <br><span class="small">Mặc định tắt để giữ nguyên 100% content đã nhập.</span>
    </label>
    <button type="submit">Tự chia lịch đăng</button>
  </form>
</section>


<section class="panel module-section" id="studio">
  <div class="section-open-note">Bạn đang mở: AI Content Studio.</div>
  <h2>AI Content Studio</h2>
  <form method="post" action="/generate">
    <textarea name="idea" rows="4" placeholder="Nhập ý tưởng: spa trị nám, proxy chạy ads, nha khoa niềng răng, bất động sản căn hộ..."></textarea>
    <div class="grid">
      <select name="style">
        <option value="gần gũi, tự nhiên, dễ ra inbox">Gần gũi</option>
        <option value="bán hàng mạnh, chốt đơn tốt">Bán hàng mạnh</option>
        <option value="chuyên nghiệp, uy tín">Chuyên nghiệp</option>
        <option value="viral, ngắn gọn, thu hút">Viral ngắn gọn</option>
        <option value="cao cấp, sang trọng">Cao cấp</option>
      </select>
      <select name="length">
        <option value="80">Ngắn dưới 80 từ</option>
        <option value="120">Vừa dưới 120 từ</option>
        <option value="180">Dài dưới 180 từ</option>
        <option value="250">Đầy đủ dưới 250 từ</option>
      </select>
    </div>
    <button type="submit" name="variants" value="1">Tạo nội dung</button>
    <button type="submit" name="variants" value="5" class="secondary">Tạo 5 phiên bản</button>
  </form>
</section>

<section class="panel module-section" id="library">
  <div class="section-open-note">Bạn đang mở: Kho Content 50.000+.</div>
  <h2>Kho Content 10.000+ Dùng Thử</h2>
  <p class="small">Bản demo nạp sẵn một số mẫu. Sau này có thể import 10.000 content thật từ JSON/CSV.</p>
  <form method="get" action="/">
    <select name="industry">
      {% for key,label in industry_labels.items() %}
      <option value="{{ key }}" {% if selected_industry==key %}selected{% endif %}>{{ label }}</option>
      {% endfor %}
    </select>
    <button type="submit">Xem mẫu content</button>
  </form>
  <div class="library-grid">
    {% for item in library_items %}
    <div class="template-card">
      <div id="tpl{{ loop.index }}">{{ item }}</div>
      <button onclick="copyText('tpl{{ loop.index }}')">Copy content</button>
    </div>
    {% endfor %}
  </div>
  <div class="lock-card" style="margin-top:16px">
    <h3>Khóa Premium</h3>
    <p>Bạn đang xem 10 content miễn phí đầu tiên.</p>
    <p>Còn lại khoảng 490 content/ngành cần nâng cấp Premium để mở khóa.</p>
    <button onclick="openPremiumPopup()">Xem bảng giá Premium</button>
  </div>
</section>

<section class="panel module-section" id="fanpage">
  <div class="section-open-note">Bạn đang mở: AI Phân tích Fanpage.</div>
  <h2>AI Phân tích Fanpage</h2>
  <form method="post" action="/analyze_fanpage">
    <input name="page_name" placeholder="Tên Fanpage">
    <div class="grid">
      <select name="avatar"><option value="co">Có avatar rõ</option><option value="khong">Chưa tốt</option></select>
      <select name="cover"><option value="co">Có ảnh bìa tốt</option><option value="khong">Chưa tốt</option></select>
      <select name="post_frequency"><option value="hang_ngay">Đăng hằng ngày</option><option value="3_5_bai_tuan">3-5 bài/tuần</option><option value="it">Ít đăng</option></select>
      <select name="cta"><option value="co">Có CTA rõ</option><option value="khong">Chưa có CTA</option></select>
    </div>
    <textarea name="bio" rows="3" placeholder="Mô tả ngắn về Fanpage, sản phẩm/dịch vụ, thông tin liên hệ..."></textarea>
    <button type="submit">Phân tích Fanpage</button>
  </form>
</section>

<section class="panel module-section" id="plan">
  <div class="section-open-note">Bạn đang mở: AI Marketing Planner.</div>
  <h2>AI Kế hoạch Marketing 30 ngày</h2>
  <form method="post" action="/plan_30_days">
    <select name="industry">
      {% for key,label in industry_labels.items() %}
      <option value="{{ key }}">{{ label }}</option>
      {% endfor %}
    </select>
    <input name="goal" placeholder="Mục tiêu: tăng inbox, tăng đơn, tăng nhận diện, ra mắt sản phẩm...">\n    <select name="days"><option value="7">7 ngày</option><option value="30" selected>30 ngày</option><option value="90">90 ngày</option></select>
    <button type="submit">Tạo kế hoạch</button>
  </form>
</section>

<section class="panel module-section" id="campaign">
  <div class="section-open-note">Bạn đang mở: Campaign Manager.</div>
  <h2>Campaign Manager</h2>
  <form method="post" action="/campaign">
    <input name="name" placeholder="Tên chiến dịch, ví dụ: Proxy tháng 6">
    <select name="industry">
      {% for key,label in industry_labels.items() %}
      <option value="{{ key }}">{{ label }}</option>
      {% endfor %}
    </select>
    <input name="goal" placeholder="Mục tiêu chiến dịch">
    <textarea name="note" rows="3" placeholder="Ghi chú chiến dịch"></textarea>
    <button type="submit">Tạo chiến dịch</button>
  </form>
  {% for c in campaigns %}
  <div class="history">ID: {{ c[0] }} | {{ c[1] }} | Ngành: {{ c[2] }} | Mục tiêu: {{ c[3] }} | Tạo: {{ c[5] }}</div>
  {% endfor %}
</section>

{% if content %}
<section class="panel module-section" id="post_edit">
  <h2>Sửa bài, Preview Facebook, chọn Fanpage</h2>
  <form method="post" action="/post" enctype="multipart/form-data">
    <div class="page-list">
      {% for p in pages %}
      <label class="page-item"><input type="checkbox" name="page_indexes" value="{{ loop.index0 }}"> {{ p.name }}<br><span class="small">{{ p.id }}</span></label>
      {% endfor %}
    </div>
    <textarea name="content" rows="12">{{ content }}</textarea>
    <input type="text" name="campaign" placeholder="Tên chiến dịch">
    <input type="file" name="image" accept="image/*,video/mp4,video/quicktime">
    <button type="submit" name="action" value="now">Đăng ngay nhiều Page</button>
    <h3>Đặt lịch đăng</h3>
    <input type="datetime-local" name="schedule_time">
    <button type="submit" name="action" value="schedule">Lưu lịch đăng</button>
  </form>
  <h3>Preview Facebook</h3>
  <div class="preview">
    <div class="preview-head"><div class="avatar"></div><div><div class="preview-name">Fanpage của bạn</div><div class="small">Vừa xong · Công khai</div></div></div>
    <div class="preview-content">{{ content }}</div>
  </div>
  <p>Điểm content: <b>{{ score }}/100</b></p>
  <ul>{% for w in warnings %}<li>{{ w }}</li>{% endfor %}</ul>
</section>
{% endif %}

{% if analysis %}
<section class="panel"><h2>Kết quả phân tích</h2><div class="history">{{ analysis }}</div></section>
{% endif %}

{% if plan %}
<section class="panel"><h2>AI Marketing Planner</h2><div class="history">{{ plan }}</div></section>
{% endif %}


<section class="panel module-section" id="factory">
  <div class="section-open-note">Bạn đang mở: Content Factory.</div>
  <h2>AI Content Factory</h2>
  <p class="small">Tạo hàng loạt 20-50 content khác nhau để chia cho nhiều Fanpage.</p>
  <form method="post" action="/content_factory">
    <textarea name="idea" rows="4" placeholder="Nhập chủ đề cần tạo hàng loạt, ví dụ: Proxy Việt Nam tốc độ cao cho chạy quảng cáo..."></textarea>
    <div class="grid">
      <select name="count">
        <option value="10">Tạo 10 content</option>
        <option value="20">Tạo 20 content</option>
        <option value="30">Tạo 30 content</option>
        <option value="50">Tạo 50 content</option>
      </select>
      <select name="style">
        <option value="bán hàng tự nhiên">Bán hàng tự nhiên</option>
        <option value="viral ngắn gọn">Viral ngắn gọn</option>
        <option value="chuyên nghiệp">Chuyên nghiệp</option>
        <option value="cao cấp">Cao cấp</option>
      </select>
    </div>
    <button type="submit">Tạo hàng loạt content</button>
  </form>
</section>

<section class="panel module-section" id="clusters">
  <h2>Page Cluster</h2>
  <p class="small">Tạo nhóm Page theo ngành như Proxy, Spa, BĐS, Nha khoa để quản lý dễ hơn.</p>
  <form method="post" action="/page_cluster">
    <input name="name" placeholder="Tên nhóm, ví dụ: Nhóm Proxy">
    <input name="page_names" placeholder="Tên Page, ngăn cách bằng dấu phẩy">
    <textarea name="note" rows="3" placeholder="Ghi chú nhóm Page"></textarea>
    <button type="submit">Lưu nhóm Page</button>
  </form>
  {% for c in clusters %}
  <div class="history">ID: {{ c[0] }} | Nhóm: {{ c[1] }}
Page: {{ c[2] }}
Ghi chú: {{ c[3] }}
Tạo: {{ c[4] }}</div>
  {% endfor %}
</section>

<section class="panel module-section" id="analytics">
  <div class="section-open-note">Bạn đang mở: Analytics Center V5.</div>
  <h2>📊 Analytics Center</h2>
  <p class="small">Trung tâm phân tích tổng hợp Fanpage, Group, bài đăng, chiến dịch và CRM để theo dõi hiệu suất bán hàng.</p>

  <div class="analytics-kpi-grid">
    <div class="analytics-kpi"><span>📌 Tổng bài</span><b>{{ analytics.summary.total_posts }}</b><small>Toàn bộ bài đã tạo</small></div>
    <div class="analytics-kpi"><span>✅ Đã đăng</span><b>{{ analytics.summary.posted }}</b><small>Tỷ lệ đăng: {{ analytics.summary.conversion_rate }}%</small></div>
    <div class="analytics-kpi"><span>⏰ Chờ đăng</span><b>{{ analytics.summary.scheduled }}</b><small>Bài đang lên lịch</small></div>
    <div class="analytics-kpi"><span>⚠️ Lỗi đăng</span><b>{{ analytics.summary.errors }}</b><small>Tỷ lệ lỗi: {{ analytics.summary.error_rate }}%</small></div>
    <div class="analytics-kpi"><span>👥 Lead CRM</span><b>{{ analytics.summary.crm_total + analytics.summary.pipeline_total }}</b><small>Tổng khách hàng ghi nhận</small></div>
    <div class="analytics-kpi"><span>💰 Pipeline</span><b>{{ "{:,}".format(analytics.summary.total_value).replace(",", ".") }}đ</b><small>Giá trị cơ hội bán hàng</small></div>
    <div class="analytics-kpi"><span>👥 Group</span><b>{{ analytics.summary.groups_total }}</b><small>{{ analytics.summary.group_scheduled }} lịch đăng Group</small></div>
    <div class="analytics-kpi"><span>🤖 AI xử lý</span><b>{{ analytics.summary.comments_total + analytics.summary.messenger_total }}</b><small>Comment + Messenger</small></div>
  </div>

  <div class="analytics-section-title">📣 Fanpage / Campaign Analytics</div>
  <div class="analytics-grid">
    <div class="analytics-box">
      <h3>Top Fanpage có bài đăng</h3>
      {% if analytics.by_page %}
        {% for r in analytics.by_page %}
        <div class="analytics-row"><span>{{ r[0] }}</span><b>{{ r[1] }}</b></div>
        <div class="analytics-bar"><i style="width:{{ 100 if loop.first else (r[1] / (analytics.by_page[0][1] if analytics.by_page[0][1] else 1) * 100)|round(0) }}%"></i></div>
        {% endfor %}
      {% else %}
        <div class="empty-analytics">Chưa có dữ liệu Fanpage.</div>
      {% endif %}
    </div>
    <div class="analytics-box">
      <h3>Hiệu suất chiến dịch</h3>
      {% if analytics.by_campaign %}
        {% for r in analytics.by_campaign %}
        <div class="analytics-row"><span>{{ r[0] }}</span><b>{{ r[1] }}</b></div>
        <div class="analytics-bar"><i style="width:{{ 100 if loop.first else (r[1] / (analytics.by_campaign[0][1] if analytics.by_campaign[0][1] else 1) * 100)|round(0) }}%"></i></div>
        {% endfor %}
      {% else %}
        <div class="empty-analytics">Chưa có dữ liệu chiến dịch.</div>
      {% endif %}
    </div>
    <div class="analytics-box">
      <h3>Trạng thái bài đăng</h3>
      {% if analytics.by_status %}
        {% for r in analytics.by_status %}
        <div class="analytics-row"><span>{{ r[0] }}</span><b>{{ r[1] }}</b></div>
        {% endfor %}
      {% else %}
        <div class="empty-analytics">Chưa có trạng thái bài đăng.</div>
      {% endif %}
    </div>
  </div>

  <div class="analytics-section-title">👥 Group / CRM Analytics</div>
  <div class="analytics-grid">
    <div class="analytics-box">
      <h3>Group theo ngành</h3>
      {% if analytics.group_by_niche %}
        {% for r in analytics.group_by_niche %}
        <div class="analytics-row"><span>{{ r[0] }}</span><b>{{ r[1] }}</b></div>
        {% endfor %}
      {% else %}
        <div class="empty-analytics">Chưa có dữ liệu Group.</div>
      {% endif %}
    </div>
    <div class="analytics-box">
      <h3>Lead theo nguồn</h3>
      {% if analytics.crm_by_source %}
        {% for r in analytics.crm_by_source %}
        <div class="analytics-row"><span>{{ r[0] }}</span><b>{{ r[1] }}</b></div>
        {% endfor %}
      {% else %}
        <div class="empty-analytics">Chưa có dữ liệu nguồn lead.</div>
      {% endif %}
    </div>
    <div class="analytics-box">
      <h3>Pipeline CRM</h3>
      {% if analytics.crm_by_stage %}
        {% for r in analytics.crm_by_stage %}
        <div class="analytics-row"><span>{{ r[0] }}</span><b>{{ r[1] }}</b></div>
        {% endfor %}
      {% else %}
        <div class="empty-analytics">Chưa có dữ liệu Kanban.</div>
      {% endif %}
    </div>
  </div>

  <div class="analytics-section-title">📅 Xu hướng 7 ngày</div>
  <div class="analytics-box full">
    {% if analytics.by_day %}
      {% for r in analytics.by_day %}
      <div class="analytics-row"><span>{{ r[0] }}</span><b>{{ r[1] }} bài</b></div>
      <div class="analytics-bar large"><i style="width:{{ 100 if loop.first else (r[1] / (analytics.by_day[0][1] if analytics.by_day[0][1] else 1) * 100)|round(0) }}%"></i></div>
      {% endfor %}
    {% else %}
      <div class="empty-analytics">Chưa có dữ liệu theo ngày. Khi bắt đầu đăng bài, biểu đồ sẽ tự cập nhật tại đây.</div>
    {% endif %}
  </div>
</section>

<section class="panel module-section" id="premium">
  <h2>💎 Premium Center V5</h2>
  <div class="premium-center v4-premium-hero">
    <div class="v4-hero-label">AI MARKETING PREMIUM</div>
    <h3>Biến công cụ đăng bài thành trợ lý Marketing tự động</h3>
    <p>Khách không chỉ mua phần mềm, khách mua thời gian, nội dung, quy trình bán hàng và hệ thống hỗ trợ tăng tốc kinh doanh online.</p>
    <div class="v4-value-grid">
      <div><b>159.000đ</b><span>Gói Khởi Động<br>chỉ 5.300đ/ngày</span></div>
      <div><b>859.000đ</b><span>Gói 1 năm<br>phổ biến nhất</span></div>
      <div><b>1.959.000đ</b><span>Gói trọn đời<br>không phí gia hạn</span></div>
    </div>
    <p><b>Giá trị nhận được:</b> tiết kiệm thời gian viết content, quản lý Fanpage, CRM khách hàng, AI Sales Script, Marketing Director và kho Content 50.000+.</p>
    <button onclick="scrollToPricing()">Xem bảng giá chi tiết</button>
    <button class="secondary" onclick="openPayment('yearly')">Chọn gói phổ biến nhất</button>
  </div>
</section>

<section class="panel module-section" id="v9center">
  <h2>V3 Enterprise Center</h2>
  <div class="v9-grid">
    <div class="v9-card">
      <h3>AI Spin 70-80%</h3>
      <p>Mỗi bài đăng được viết lại khác nhau, thêm CTA và hashtag riêng để giảm trùng lặp.</p>
    </div>
    <div class="v9-card">
      <h3>AI Chọn ảnh phù hợp</h3>
      <p>Bản local ưu tiên chia ảnh không trùng. Có thể nâng tiếp để AI phân loại ảnh theo nội dung.</p>
    </div>
    <div class="v9-card">
      <h3>Auto 100 Content → 100 Page</h3>
      <p>Hệ thống tự chia content, media và Page theo vòng quay, phù hợp đăng hàng loạt.</p>
    </div>
    <div class="v9-card">
      <h3>Token Guard</h3>
      <p>Kiểm tra token trước khi đăng ngay. Nếu token chết, hệ thống dừng để tránh lỗi hàng loạt.</p>
    </div>
  </div>
  <div style="margin-top:14px">
    <a href="/backup"><button>Backup Database</button></a>
    <a href="/export_pdf"><button class="secondary">Xuất báo cáo PDF</button></a>
    <a href="/export"><button class="secondary">Xuất báo cáo Excel/CSV</button></a>
  </div>
</section>


<section class="panel module-section" id="crm">
  <div class="section-open-note">Bạn đang mở: CRM Mini.</div>
  <h2>CRM Mini</h2>
  <p class="small">Lưu khách hàng tiềm năng từ comment, inbox, Zalo hoặc data thủ công.</p>
  <form method="post" action="/crm">
    <div class="grid">
      <input name="name" placeholder="Tên khách hàng">
      <input name="phone" placeholder="Số điện thoại">
      <input name="zalo" placeholder="Zalo">
      <input name="source" placeholder="Nguồn: Facebook, Zalo, TikTok...">
    </div>
    <textarea name="note" rows="3" placeholder="Ghi chú nhu cầu khách hàng"></textarea>
    <button type="submit">Lưu khách hàng</button>
  </form>
  {% for r in crm_rows %}
  <div class="history">ID: {{ r[0] }} | {{ r[1] }} | SĐT: {{ r[2] }} | Zalo: {{ r[3] }} | Nguồn: {{ r[4] }} | {{ r[6] }}
Ghi chú: {{ r[5] }}</div>
  {% endfor %}
</section>

<section class="panel module-section" id="token">
  <div class="section-open-note">Bạn đang mở: Token Manager.</div>
  <h2>Token Center Pro</h2>
  <p class="small">Thêm, cập nhật và kiểm tra Page Token trực tiếp trong tool. Không cần vào Render Environment để sửa PAGES_JSON nữa.</p>

  <div class="grid">
    <form method="post" action="/fanpage_token_add">
      <h3>➕ Thêm / cập nhật Fanpage</h3>
      <input name="page_name" placeholder="Tên Fanpage, ví dụ: GPT Mini Premium">
      <input name="page_id" placeholder="Page ID">
      <textarea name="page_token" rows="4" placeholder="Dán Page Access Token tại đây"></textarea>
      <input name="note" placeholder="Ghi chú, ví dụ: Page chính / Page ads / Page dự phòng">
      <button type="submit">Lưu Fanpage Token</button>
    </form>

    <div class="history">
      <h3>📌 Trạng thái cấu hình</h3>
      {{ token_report }}
      <form method="post" action="/check_tokens" style="margin-top:12px">
        <button type="submit">Kiểm tra toàn bộ Page Token</button>
      </form>
    </div>
  </div>

  <h3>Danh sách Fanpage đang kết nối</h3>
  <table class="v5-table">
    <tr><th>Fanpage</th><th>Page ID</th><th>Token</th><th>Nguồn</th><th>Ghi chú</th><th>Hành động</th></tr>
    {% for p in pages %}
    <tr>
      <td>{{ p.name }}</td>
      <td>{{ p.id }}</td>
      <td>{{ p.token_mask }}</td>
      <td>{{ 'Tool' if p.source == 'database' else '.env' }}</td>
      <td>{{ p.note }}</td>
      <td>
        {% if p.source == 'database' %}
        <form method="post" action="/fanpage_token_delete" onsubmit="return confirm('Xóa Fanpage này khỏi Token Center?')">
          <input type="hidden" name="row_id" value="{{ p.row_id }}">
          <button class="secondary" type="submit">Xóa</button>
        </form>
        {% else %}
        <span class="small">Sửa trong Environment</span>
        {% endif %}
      </td>
    </tr>
    {% endfor %}
  </table>

  <h3>Kết quả kiểm tra gần nhất</h3>
  {% for t in token_checks %}
  <div class="history">
Page: {{ t[0] }}
ID: {{ t[1] }}
Trạng thái: {{ t[2] }}
Chi tiết: {{ t[3] }}
Thời gian: {{ t[4] }}
  </div>
  {% endfor %}

  <div class="premium-center">
    <h3>Hướng dẫn khi token giới hạn</h3>
    <p>1. Vào Meta Graph API Explorer.</p>
    <p>2. Generate User Token có quyền pages_show_list, pages_read_engagement, pages_manage_posts, pages_manage_metadata.</p>
    <p>3. Gọi /me/accounts để lấy Page Access Token mới.</p>
    <p>4. Dán token mới vào form Thêm / cập nhật Fanpage ở Token Center.</p>
    <p>5. Bấm Lưu Fanpage Token rồi bấm kiểm tra token, không cần deploy lại.</p>
  </div>
</section>

<section class="panel module-section" id="batch">
  <div class="section-open-note">Bạn đang mở: Excel / CSV hàng loạt.</div>
  <h2>Đăng hàng loạt Excel / CSV</h2>
  <p class="small">Cột hỗ trợ: idea, content, page_names, schedule_time, campaign, image_path hoặc media_path. Nếu không điền page_names, hệ thống tự chia bài lần lượt qua các Page để tránh trùng.</p>
  <form method="post" action="/batch" enctype="multipart/form-data">
    <input type="file" name="batch_file" accept=".xlsx,.csv">
    <button type="submit">Nhập file và lưu lịch</button>
  </form>
  <a href="/export"><button class="secondary">Xuất báo cáo CSV</button></a>
</section>

<section class="panel module-section" id="history">
  <h2>Lịch sử bài đăng / lịch hẹn</h2>
  {% for h in history %}
  <div class="history">
ID: {{ h[0] }}
Page: {{ h[1] }}
Trạng thái: {{ h[3] }}
Post ID / lỗi: {{ h[4] }}
Lịch đăng: {{ h[5] }}
Chiến dịch: {{ h[7] }}
Điểm content: {{ h[8] }}
Thời gian tạo: {{ h[9] }}

{{ h[2] }}
  </div>
  {% endfor %}
</section>

<section class="panel pricing-visible" id="pricing">
  <div class="premium-pricing-pro v4-pricing-shell">
    <div class="premium-title">
      <span class="mini">V5 SELLER AI PREMIUM</span>
      <h2>Bảng Giá Premium</h2>
      <p>Thiết kế theo giá trị nhận được: tiết kiệm thời gian, giảm chi phí thuê ngoài và mở khóa AI Marketing chuyên nghiệp.</p>
    </div>

    <div class="v4-pricing-stats">
      <div><b>2.381+</b><span>Khách hàng sử dụng</span></div>
      <div><b>120.000+</b><span>Content đã tạo</span></div>
      <div><b>98%</b><span>Đánh giá hài lòng</span></div>
    </div>

    <div class="premium-grid-pro v4-premium-grid">
      <div class="premium-plan free-plan v4-plan trial-experience-card">
        <div class="plan-ribbon trial-ribbon">DÙNG THỬ</div>
        <div class="trial-badge-pro">🎁 Gói Trải Nghiệm</div>
        <div class="trial-days-pro">3 ngày miễn phí</div>

        <div class="trial-features-pro">
          <div class="feature-open-pro">✓ Quản lý Fanpage</div>
          <div class="feature-open-pro">✓ Quản lý Group</div>
          <div class="feature-open-pro">✓ AI Comment</div>

          <div class="feature-lock-pro">🔒 AI Messenger</div>
          <div class="feature-lock-pro">🔒 CRM Kanban</div>
          <div class="feature-lock-pro">🔒 AI Marketing Director</div>
        </div>

        <div class="trial-upgrade-pro">
          📈 Nâng cấp để mở toàn bộ hệ sinh thái AI bán hàng
        </div>
        <button class="plan-button premium-btn trial-upgrade-btn" onclick="openPayment('monthly')">Nâng cấp Premium</button>
      </div>

      <div class="premium-plan v4-plan">
        <div class="value-badge">🚀 GÓI 1 THÁNG</div>
        <div class="plan-name">Gói 1 tháng</div>
        <div class="plan-price">159.000đ</div>
        <div class="price-sub">Chỉ 5.300đ/ngày</div>
        <div class="benefit-title">Phù hợp</div>
        <ul class="benefit-list">
          <li class="open">Chủ shop mới</li>
          <li class="open">Cá nhân kinh doanh online</li>
          <li class="open">Người mới chạy quảng cáo</li>
        </ul>
        <div class="benefit-title">Giá trị nhận được</div>
        <div class="v4-save-box">🔥 Tiết kiệm 2-3 giờ mỗi ngày</div>
        <ul class="benefit-list">
          <li class="open">AI Content Studio</li>
          <li class="open">Tạo 5 content/lần</li>
          <li class="open">Đăng Fanpage, ảnh, Video, Reel</li>
          <li class="open">Upload Excel / CSV</li>
          <li class="open">Lịch đăng cơ bản</li>
          <li class="open">Dashboard cơ bản</li>
          <li class="open">Token Manager</li>
        </ul>
        <button class="plan-button premium-btn" onclick="openPayment('monthly')">Mở khóa gói 1 tháng</button>
      </div>

      <div class="premium-plan v4-plan">
        <div class="plan-name">Gói 3 tháng</div>
        <div class="plan-price">359.000đ</div>
        <div class="price-sub">Tối ưu hơn gói tháng</div>
        <div class="benefit-title">Bao gồm toàn bộ gói 1 tháng</div>
        <ul class="benefit-list">
          <li class="open">Campaign Manager</li>
          <li class="open">AI Marketing Planner</li>
          <li class="open">Kho Content Premium</li>
          <li class="open">Token Center nâng cao</li>
          <li class="open">Báo cáo CSV</li>
          <li class="open">Lịch đăng nâng cao</li>
        </ul>
        <div class="v4-value-received">
          <b>🎁 Giá trị nhận được</b>
          <span>Viết content + lập kế hoạch + báo cáo cơ bản cho shop nhỏ.</span>
        </div>
        <button class="plan-button premium-btn" onclick="openPayment('quarterly')">Đăng ký 3 tháng</button>
      </div>

      <div class="premium-plan v4-plan">
        <div class="plan-name">Gói 6 tháng</div>
        <div class="plan-price">559.000đ</div>
        <div class="price-sub">Mở CRM và Sales Bot</div>
        <div class="benefit-title">Bao gồm toàn bộ gói 3 tháng</div>
        <ul class="benefit-list">
          <li class="open">CRM Pro</li>
          <li class="open">AI Sales Bot</li>
          <li class="open">Comment Manager</li>
          <li class="open">Auto Tag khách hàng</li>
          <li class="open">Quản lý khách hàng</li>
          <li class="open">Chuyển khách sang CRM</li>
        </ul>
        <div class="v4-value-received">
          <b>🎁 Giá trị nhận được</b>
          <span>Tối ưu quy trình tư vấn, gom khách và chăm sóc khách hàng.</span>
        </div>
        <button class="plan-button premium-btn" onclick="openPayment('halfyear')">Đăng ký 6 tháng</button>
      </div>

      <div class="premium-plan featured v4-plan v4-yearly">
        <div class="plan-ribbon">⭐ PHỔ BIẾN NHẤT</div>
        <div class="plan-name">Gói 1 năm</div>
        <div class="plan-price">859.000đ</div>
        <div class="price-sub">Chỉ 2.300đ/ngày</div>
        <div class="benefit-title">Giá trị thực tế</div>
        <ul class="benefit-list">
          <li class="open">AI Content Brain</li>
          <li class="open">AI Marketing Director</li>
          <li class="open">AI Ads Chuyên Gia</li>
          <li class="open">CRM Pro</li>
          <li class="open">Kho 50.000+ Content</li>
          <li class="open">Automation Marketing</li>
        </ul>
        <div class="v4-save-box">🚀 Tiết kiệm ~12.000.000đ/năm chi phí thuê nhân sự</div>
        <button class="plan-button premium-btn" onclick="openPayment('yearly')">Chọn gói phổ biến nhất</button>
      </div>

      <div class="premium-plan featured v4-plan v4-lifetime">
        <div class="plan-ribbon">👑 NHÀ BÁN HÀNG PRO</div>
        <div class="plan-name">Gói nhà bán hàng chuyên nghiệp</div>
        <div class="plan-price">1.959.000đ</div>
        <div class="price-sub">Gói cao nhất cho nhà bán hàng chuyên nghiệp</div>
        <div class="benefit-title">Bao gồm</div>
        <ul class="benefit-list">
          <li class="open">Toàn bộ AI hiện tại</li>
          <li class="open">Toàn bộ AI tương lai</li>
          <li class="open">Cập nhật miễn phí</li>
          <li class="open">Ưu tiên hỗ trợ</li>
          <li class="open">Không phí gia hạn</li>
          <li class="open">AI Image Center</li>
          <li class="open">AI Video Center</li>
          <li class="open">Dashboard Enterprise</li>
          <li class="open">Export PDF / Excel</li>
          <li class="open">Backup Database</li>
        </ul>
        <div class="v4-save-box gold">💰 Tiết kiệm hơn 10.000.000đ chi phí sử dụng lâu dài</div>
        <button class="plan-button premium-btn" onclick="openPayment('lifetime')">Mở khóa trọn đời</button>
      </div>
    </div>

    <div class="v4-outsourcing-box">
      <h3>🎁 Giá trị nhận được nếu thuê ngoài</h3>
      <div class="v4-outsource-grid">
        <div><b>2.000.000đ/tháng</b><span>Viết content</span></div>
        <div><b>3.000.000đ/tháng</b><span>Marketing</span></div>
        <div><b>1.000.000đ/tháng</b><span>CRM</span></div>
        <div><b>6.000.000đ+</b><span>Tổng giá trị</span></div>
      </div>
      <p>Bạn chỉ trả từ <b>159.000đ</b> để có hệ thống AI Marketing hỗ trợ tạo nội dung, quản lý khách và tối ưu bán hàng.</p>
    </div>

    <div class="premium-note-box">
      Sau khi thanh toán, nếu 5 phút chưa kích hoạt tự động, vui lòng gửi ảnh giao dịch và nội dung thanh toán qua Zalo 036 338 2629 để được hỗ trợ nhanh.
    </div>
  </div>
</section>

<div class="payment-modal" id="paymentModal">
  <div class="payment-inner">
    <div class="payment-head">
      <div>
        <h2 id="payPlanTitle">GÓI PREMIUM</h2>
        <div id="payPlanDesc">Thông tin gói Premium</div>
      </div>
      <button class="payment-close" onclick="closePayment()">Đóng</button>
    </div>

    <div class="payment-body">
      <div class="qr-card">
        <h3>Quét mã QR để thanh toán</h3>
        <img id="payQr" src="https://img.vietqr.io/image/970405-8888363382629-compact2.png?amount=159000&addInfo=PREMIUM%201THANG&accountName=NGUYEN%20DANG%20THI%20XUAN" alt="QR Agribank">
        <div class="bank-info">
          <b>Ngân hàng:</b> Agribank<br>
          <b>STK:</b> 8888363382629<br>
          <b>Chủ TK:</b> NGUYEN DANG THI XUAN<br>
          <b>Nội dung CK:</b> <span id="payContent">PREMIUM 1THANG</span>
        </div>
      </div>

      <div class="payment-detail">
        <h3>Số tiền cần thanh toán</h3>
        <div class="pay-amount" id="payPlanPrice">159.000 VNĐ</div>

        <h3>Quyền lợi gói này</h3>
        <div class="payment-benefits" id="payBenefits"></div>

        <h3>Tính năng chưa mở ở gói này</h3>
        <div class="payment-benefits" id="payLocked"></div>

        <div class="payment-alert">
          Khi chuyển khoản, vui lòng giữ đúng <b>Nội dung CK</b> có kèm ID thiết bị.
          Sau khi thanh toán, nhập <b>Số điện thoại</b> và <b>Gmail</b> bên dưới để gửi yêu cầu kích hoạt Premium. Bộ phận quản trị sẽ kiểm tra và duyệt trên hệ thống.
        </div>

        <div class="premium-confirm-box" id="premiumConfirmFormBox">
          <h3>Gửi yêu cầu kích hoạt Premium</h3>
          <div class="small">ID thiết bị: <b class="device-id-text">Đang tạo...</b></div>
          <input id="premiumPhone" placeholder="Số điện thoại/Zalo đã thanh toán">
          <input id="premiumEmail" placeholder="Gmail/Email đăng ký">
          <textarea id="premiumTransactionNote" placeholder="Mã giao dịch hoặc ghi chú chuyển khoản (không bắt buộc)"></textarea>
          <button onclick="submitPremiumRequest()">Gửi về Admin duyệt kích hoạt</button>
          <div class="premium-request-status" id="premiumRequestStatus"></div>
        </div>

        <div class="payment-actions">
          <a href="https://zalo.me/0363382629" target="_blank">Liên hệ Zalo hỗ trợ</a>
          <a class="light" href="javascript:void(0)" onclick="scrollToPaymentConfirm()">Tôi đã thanh toán</a>
        </div>
      </div>
    </div>
  </div>
</div>

<div class="payment-modal" id="lockedFeatureModal">
  <div class="payment-inner" style="max-width:760px">
    <div class="payment-head">
      <div>
        <h2 id="lockedFeatureTitle">🔒 Tính năng Premium</h2>
        <div>Tính năng nâng cao cần nâng cấp Premium để sử dụng.</div>
      </div>
      <button class="payment-close" onclick="closeLockedFeature()">Đóng</button>
    </div>
    <div class="payment-detail">
      <div id="lockedFeaturePlans"></div>
    </div>
  </div>
</div>

</main>

<aside class="rightbar">
  <h2>🔥 Hoạt động hôm nay</h2>
  <div class="activity-card">
    <span>📌 Tổng bài</span><b>{{ s.total }}</b>
  </div>
  <div class="activity-card">
    <span>✅ Đã đăng</span><b>{{ s.posted }}</b>
  </div>
  <div class="activity-card">
    <span>⏰ Chờ đăng</span><b>{{ s.scheduled }}</b>
  </div>
  <div class="activity-card">
    <span>👥 Lead CRM</span><b>{{ s.crm }}</b>
  </div>
  <div class="activity-card">
    <span>📊 Chiến dịch</span><b>{{ s.campaigns }}</b>
  </div>

  <div class="{{ 'free-status-card free-expired' if free_status.is_expired else 'free-status-card' }} trial-side-card">
    <h3>🎁 Gói Trải Nghiệm</h3>
    {% if free_status.is_expired %}
      <b>Trạng thái:</b> Đã hết dùng thử<br>
      Vui lòng nâng cấp Premium để mở toàn bộ hệ sinh thái AI bán hàng.
    {% else %}
      <div class="trial-side-days">3 ngày miễn phí</div>
      <b>Còn lại:</b> {{ free_status.days }} ngày {{ free_status.hours }} giờ
      <div class="free-progress"><span style="width:{{ free_status.percent }}%"></span></div>
      <div class="trial-side-list open">
        ✓ Quản lý Fanpage<br>
        ✓ Quản lý Group<br>
        ✓ AI Comment
      </div>
      <div class="trial-side-list locked">
        🔒 AI Messenger<br>
        🔒 CRM Kanban<br>
        🔒 AI Marketing Director
      </div>
      <div class="trial-side-upgrade">📈 Nâng cấp để mở toàn bộ hệ sinh thái AI bán hàng</div>
      <button onclick="scrollToPricing()">Xem chi tiết gói</button>
      <button onclick="openPayment('monthly')">Nâng cấp Premium</button>
    {% endif %}
  </div>

  <hr>
  <h2>⚡ Trạng thái hệ thống</h2>
  <p style="white-space:pre-line">{{ token_warning }}</p>

  <div class="v5-focus-box">
    <b>V5 Seller AI Suite</b><br>
    Fanpage • Group • AI Comment • AI Messenger • CRM Kanban • Marketing Director
  </div>
</aside>
</div>

<nav class="mobilebar">
  <a href="#dashboard">🏠<br>Home</a>
  <a href="#post">✍️<br>Đăng</a>
  <a href="#library">📚<br>Kho</a>
  <a href="#plan">🎯<br>Plan</a>
  <a href="#history">📊<br>Lịch sử</a>
</nav>

<script>
function toggleMenuGroup(el){
  const box = el.closest('.menu-mini-group');
  if(!box) return;
  box.classList.toggle('open');
}
</script>


<script>
(function(){
  function planKeyFromText(text){
    text = (text || '').toLowerCase();
    if(text.includes('1.959') || text.includes('1959') || text.includes('vĩnh') || text.includes('life')) return 'lifetime';
    if(text.includes('859') || text.includes('1 năm') || text.includes('nam') || text.includes('year')) return 'yearly';
    if(text.includes('559') || text.includes('6 tháng') || text.includes('business')) return 'halfyear';
    if(text.includes('359') || text.includes('3 tháng') || text.includes('pro')) return 'quarterly';
    if(text.includes('159') || text.includes('1 tháng') || text.includes('basic')) return 'monthly';
    return 'yearly';
  }
  function bindSafePricing(){
    document.querySelectorAll('.price-card,.premium-plan').forEach(function(card){
      if(card.dataset.safePricingReady) return;
      card.dataset.safePricingReady = '1';
      card.title = 'Bấm để xem chi tiết gói';
      card.addEventListener('click', function(e){
        if(e.target && (e.target.tagName === 'A' || e.target.tagName === 'BUTTON')) return;
        var key = planKeyFromText(card.innerText);
        if(typeof openPremiumCheckout === 'function') openPremiumCheckout(key);
        else if(typeof openPayment === 'function') openPayment(key);
        else if(typeof openPremium === 'function') openPremium();
      });
      var action = card.querySelector('.plan-button,.safe-pricing-action');
      if(action) {
        action.innerText = action.innerText.replace('','Xem chi tiết gói');
      }
    });
  }
  document.addEventListener('DOMContentLoaded', bindSafePricing);
  setTimeout(bindSafePricing, 600);
  setTimeout(bindSafePricing, 1600);
})();
</script>


<script>
function showInstallGuide(){
  alert("Cách cài App Mini:\\n\\nAndroid Chrome: bấm dấu 3 chấm → Thêm vào màn hình chính.\\n\\niPhone Safari: bấm Chia sẻ → Thêm vào MH chính.\\n\\nSau đó mở Mkt Automation Pro như một app trên điện thoại.");
}

let deferredInstallPrompt = null;
window.addEventListener('beforeinstallprompt', function(e){
  e.preventDefault();
  deferredInstallPrompt = e;
});
</script>


<script>
if ('serviceWorker' in navigator) {
  window.addEventListener('load', function(){
    navigator.serviceWorker.register('/service-worker.js').catch(function(err){
      console.log('Service worker registration failed:', err);
    });
  });
}
</script>


<script>
function showInstallGuide(){
  alert("Cách cài App Mini:\\n\\nAndroid Chrome: bấm dấu 3 chấm → Thêm vào màn hình chính.\\n\\niPhone Safari: bấm Chia sẻ → Thêm vào MH chính.\\n\\nSau đó mở Mkt Automation Pro V2 như một app.");
}
(function(){
  function planKeyFromText(text){
    text = (text || '').toLowerCase();
    if(text.includes('1.959') || text.includes('1959') || text.includes('vĩnh') || text.includes('life')) return 'lifetime';
    if(text.includes('859') || text.includes('1 năm') || text.includes('nam') || text.includes('year')) return 'yearly';
    if(text.includes('559') || text.includes('6 tháng') || text.includes('business')) return 'halfyear';
    if(text.includes('359') || text.includes('3 tháng') || text.includes('pro')) return 'quarterly';
    if(text.includes('159') || text.includes('1 tháng') || text.includes('basic')) return 'monthly';
    return 'yearly';
  }
  function bindV2Pricing(){
    document.querySelectorAll('.price-card,.premium-plan').forEach(function(card){
      if(card.dataset.v2PricingReady) return;
      card.dataset.v2PricingReady = '1';
      card.title = 'Bấm để xem chi tiết gói';
      card.addEventListener('click', function(e){
        if(e.target && (e.target.tagName === 'A' || e.target.tagName === 'BUTTON')) return;
        var key = planKeyFromText(card.innerText);
        if(typeof openPremiumCheckout === 'function') openPremiumCheckout(key);
        else if(typeof openPayment === 'function') openPayment(key);
        else if(typeof openPremium === 'function') openPremium();
      });
      card.querySelectorAll('button').forEach(function(btn){
        if((btn.innerText || '').includes('Xem chi tiết gói')) btn.innerText = 'Xem chi tiết gói';
      });
    });
  }
  document.addEventListener('DOMContentLoaded', bindV2Pricing);
  setTimeout(bindV2Pricing, 800);
})();
if ('serviceWorker' in navigator) {
  window.addEventListener('load', function(){
    navigator.serviceWorker.register('/service-worker.js').catch(function(err){
      console.log('Service worker failed:', err);
    });
  });
}
</script>


<script>
let draggedKanbanCard=null;
function dragKanban(ev){ draggedKanbanCard=ev.target; }
function dropKanban(ev){ ev.preventDefault(); const col=ev.currentTarget; if(draggedKanbanCard){ col.appendChild(draggedKanbanCard); draggedKanbanCard=null; } }
</script>



<script>
window.mktDeferredInstallPrompt = window.mktDeferredInstallPrompt || null;
window.addEventListener('beforeinstallprompt', function(e){
  e.preventDefault();
  window.mktDeferredInstallPrompt = e;
  var st=document.getElementById('installStatus');
  if(st){st.innerText='Thiết bị này đã sẵn sàng cài đặt ứng dụng.';}
});
function showInstallGuide(){
  var st=document.getElementById('installStatus');
  var isStandalone = (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) || window.navigator.standalone;
  if(isStandalone){
    if(st) st.innerText='App đã được cài đặt trên thiết bị này.';
    alert('GPT MKT Pro đã được cài đặt trên thiết bị này.');
    return;
  }
  var title='Cài đặt GPT MKT Pro';
  var intro=`✔ Dùng như app trên điện thoại
✔ Không cần mở trình duyệt
✔ Nhận thông báo nhanh
✔ Truy cập chỉ 1 chạm`;
  if(window.mktDeferredInstallPrompt){
    if(confirm(`${title}

${intro}

Bấm OK để cài đặt ngay.`)){
      window.mktDeferredInstallPrompt.prompt();
      window.mktDeferredInstallPrompt.userChoice.then(function(choice){
        if(st){st.innerText = choice.outcome === 'accepted' ? 'Đã gửi yêu cầu cài đặt ứng dụng.' : 'Anh/chị có thể bấm cài đặt lại bất kỳ lúc nào.';}
        window.mktDeferredInstallPrompt = null;
      });
    }
    return;
  }
  var isIOS=/iphone|ipad|ipod/i.test(navigator.userAgent);
  var msg=isIOS
    ? 'iPhone/iPad: mở bằng Safari → bấm Chia sẻ → Thêm vào Màn hình chính.'
    : 'Chrome/Edge: bấm biểu tượng cài đặt trên thanh địa chỉ hoặc menu ⋮ → Cài đặt ứng dụng / Thêm vào màn hình chính.';
  if(st){st.innerText='Trình duyệt chưa bật hộp cài đặt tự động. Làm theo hướng dẫn vừa hiển thị.';}
  alert(`${title}

${intro}

${msg}`);
}
window.addEventListener('appinstalled', function(){
  var st=document.getElementById('installStatus');
  if(st){st.innerText='Đã cài đặt GPT MKT Pro thành công.';}
});
</script>


<script>
function showSmartInstall(){
  var isStandalone = (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) || window.navigator.standalone;
  if(isStandalone){ alert('Ứng dụng đã được cài đặt trên thiết bị này.'); return; }
  if(window.mktDeferredInstallPrompt){ showInstallGuide(); return; }
  window.location.href='/install';
}
</script>

<!-- Mini Chat Support - lưu tin nhắn để Admin trả lời trong /admin -->
<style>
.support-float{position:fixed;right:18px;bottom:18px;z-index:9999;font-family:Arial,sans-serif}.support-btn{width:60px;height:60px;padding:0;border:0;border-radius:50%;cursor:pointer;position:relative;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#2563eb,#7c3aed);box-shadow:0 0 0 5px rgba(37,99,235,.12),0 14px 34px rgba(37,99,235,.38);animation:supportBotFloat 2.2s ease-in-out infinite}.support-robot{font-size:27px;line-height:1}.support-online-dot{position:absolute;right:5px;bottom:6px;width:10px;height:10px;border-radius:50%;background:#00ff88;border:2px solid white;box-shadow:0 0 8px #00ff88,0 0 15px rgba(0,255,136,.85);animation:supportOnlinePulse 1.5s infinite}.support-tooltip{position:absolute;right:66px;bottom:7px;min-width:150px;background:rgba(15,23,42,.96);color:#E0F2FE;border:1px solid rgba(34,197,94,.35);border-radius:14px;padding:9px 11px;font-size:12px;line-height:1.35;text-align:left;opacity:0;pointer-events:none;transform:translateX(8px);transition:.18s ease;box-shadow:0 14px 34px rgba(15,23,42,.28)}.support-btn:hover{transform:scale(1.08)}.support-btn:hover .support-tooltip{opacity:1;transform:translateX(0)}@keyframes supportBotFloat{0%,100%{transform:translateY(0)}50%{transform:translateY(-5px)}}@keyframes supportOnlinePulse{0%{transform:scale(1);opacity:1}50%{transform:scale(1.35);opacity:.78}100%{transform:scale(1);opacity:1}}.support-panel{display:none;width:340px;max-width:calc(100vw - 30px);background:#0f172a;color:#e5e7eb;border:1px solid #334155;border-radius:20px;box-shadow:0 18px 60px rgba(0,0,0,.45);overflow:hidden}.support-panel.open{display:block}.support-head{background:#1e1b4b;padding:12px 14px;font-weight:900;color:#bfdbfe;display:flex;align-items:center;justify-content:space-between}.support-close{background:#020617;color:white;border:1px solid #334155;border-radius:10px;width:28px;height:28px;cursor:pointer}.support-body{padding:12px}.support-mini-menu{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;margin-bottom:8px}.support-mini-menu button{border:1px solid #334155;background:#111827;color:#dbeafe;border-radius:12px;padding:8px 5px;font-size:12px;font-weight:900;cursor:pointer}.support-mini-menu button:hover{background:#1e40af}.support-log{height:170px;overflow-y:auto;background:#020617;border:1px solid #1f2937;border-radius:14px;padding:10px;margin-bottom:10px;font-size:13px}.support-log .me{background:#1d4ed8;margin:6px 0 6px 35px;padding:8px;border-radius:12px}.support-log .ad{background:#14532d;margin:6px 35px 6px 0;padding:8px;border-radius:12px}.support-body input,.support-body textarea{width:100%;background:#020617;color:white;border:1px solid #334155;border-radius:12px;padding:10px;margin:5px 0}.support-body textarea{height:78px}.support-send{width:100%;background:#22c55e;color:white;border:0;border-radius:12px;padding:11px;font-weight:900;cursor:pointer}.support-note{font-size:12px;color:#94a3b8;margin-top:8px}.compact-actions{display:grid!important;grid-template-columns:repeat(3,1fr);gap:6px!important}.compact-actions button,.compact-actions a{font-size:12px!important;padding:8px 6px!important;border-radius:12px!important;text-align:center!important}

/* SaaS cleanup: keep only one AI support bot */
.floating-bot{display:none!important}
.support-float{right:18px!important;bottom:18px!important}
.support-online-dot{width:10px!important;height:10px!important;right:5px!important;bottom:6px!important}
.support-btn{width:60px!important;height:60px!important;transition:transform .18s ease,box-shadow .18s ease}


/* Premium support bot typography */
.support-float,.support-panel,.support-panel *{
  font-family:'Manrope','Inter',Arial,sans-serif!important;
  letter-spacing:-.015em;
}
.support-head{font-weight:800!important;letter-spacing:-.03em!important;}
.support-tooltip{font-family:'Inter','Manrope',Arial,sans-serif!important;font-weight:700!important;}
.support-mini-menu button,.support-close,.support-send{font-family:'Manrope','Inter',Arial,sans-serif!important;font-weight:800!important;}
</style>
<div class="support-float">
  <button class="support-btn" title="AI Online" onclick="toggleSupportChat()"><span class="support-robot">🤖</span><span class="support-online-dot"></span><span class="support-tooltip">AI Online<br>Phản hồi trong vài giây</span></button>
  <div class="support-panel" id="supportPanel">
    <div class="support-head">
      <span>💬 Hỗ trợ trực tiếp</span>
      <button type="button" onclick="toggleSupportChat()" class="support-close">×</button>
    </div>
    <div class="support-body">
      <div class="support-mini-menu">
        <button type="button" onclick="quickSupportText('Tôi cần kích hoạt Premium')">👑 Premium</button>
        <button type="button" onclick="quickSupportText('Tôi đã thanh toán cần hỗ trợ')">💳 Thanh toán</button>
        <button type="button" onclick="quickSupportText('Tôi bị lỗi đăng bài Fanpage')">📣 Lỗi đăng</button>
      </div>
      <div class="support-log" id="supportLog"><div class="ad">Admin sẵn sàng hỗ trợ. Anh/chị để lại SĐT/Email và nội dung cần xử lý.</div></div>
      <input id="supportPhone" placeholder="SĐT/Zalo của anh/chị">
      <input id="supportEmail" placeholder="Email/Gmail">
      <textarea id="supportMessage" placeholder="Nhập nội dung cần hỗ trợ..."></textarea>
      <button class="support-send" onclick="sendSupportMessage()">Gửi cho Admin</button>
      <div class="support-note" id="supportNote">Tin nhắn sẽ hiển thị trong Web Admin để kỹ thuật trả lời.</div>
    </div>
  </div>
</div>
<script>
function getMktDeviceId(){
  let id=localStorage.getItem('mkt_device_id');
  if(!id){id='MP-'+new Date().toISOString().slice(0,10).replaceAll('-','')+'-'+Math.random().toString(16).slice(2,8).toUpperCase();localStorage.setItem('mkt_device_id',id)}
  return id;
}
let lastSupportId=0;
let shownAdminReplies = new Set();
function toggleSupportChat(){document.getElementById('supportPanel').classList.toggle('open');pollSupportReplies();}
function quickSupportText(text){
  const box=document.getElementById('supportMessage');
  if(box){box.value=text;box.focus();}
}
function addSupportBubble(type,text){const log=document.getElementById('supportLog');const div=document.createElement('div');div.className=type;div.innerText=text;log.appendChild(div);log.scrollTop=log.scrollHeight;}
async function sendSupportMessage(){
  const msg=document.getElementById('supportMessage').value.trim();
  if(!msg){alert('Vui lòng nhập nội dung cần hỗ trợ.');return;}
  addSupportBubble('me',msg);document.getElementById('supportMessage').value='';
  const payload={device_id:getMktDeviceId(),phone:document.getElementById('supportPhone').value,email:document.getElementById('supportEmail').value,message:msg};
  const res=await fetch('/support_message',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}).then(r=>r.json()).catch(()=>({success:false,message:'Không gửi được tin nhắn, vui lòng thử lại.'}));
  document.getElementById('supportNote').innerText=res.message||'Đã gửi.';
  pollSupportReplies();
}
async function pollSupportReplies(){
  const data=await fetch('/support_poll?device_id='+encodeURIComponent(getMktDeviceId())+'&after_id=0').then(r=>r.json()).catch(()=>({messages:[]}));
  (data.messages||[]).forEach(function(m){
    lastSupportId=Math.max(lastSupportId,m.id||0);
    if(m.admin_reply && !shownAdminReplies.has(m.id)){
      shownAdminReplies.add(m.id);
      addSupportBubble('ad','Admin: '+m.admin_reply);
    }
  });
}
setInterval(function(){if(document.getElementById('supportPanel')&&document.getElementById('supportPanel').classList.contains('open')) pollSupportReplies();},3000);
</script>
</body>
</html>
"""

def current_library(selected_industry):
    return CONTENT_LIBRARY.get(selected_industry, CONTENT_LIBRARY["spa"])

def render(content="", message="", ok=True, selected_industry="spa", analysis="", plan=""):
    score = score_content(content) if content else 0
    warnings = policy_check(content) if content else []
    pages_current = get_fanpages()
    
    if pages_current and GEMINI_API_KEY:
        token_warning = f"🟢 Hệ thống sẵn sàng\nGemini AI: Đã kết nối\nFanpage: {len(pages_current)} Fanpage\nPremium: Hoạt động"
    else:
        token_warning = "🟡 Hệ thống cần cấu hình thêm\nGemini AI hoặc Fanpage chưa sẵn sàng\nVui lòng vào Token Fanpage để kiểm tra."
    return render_template_string(
        HTML, title=APP_TITLE, pages=pages_current, content=content, message=message, ok=ok,
        history=get_history(), campaigns=get_campaigns(), s=get_stats(), crm_rows=get_crm(), token_report=token_manager_report(), token_checks=get_latest_token_checks(), clusters=get_page_clusters(), analytics=get_analytics_summary(), free_status=get_free_status(),
        industry_labels=INDUSTRY_LABELS, selected_industry=selected_industry,
        library_items=current_library(selected_industry)[:10], locked_count=max(0, 500 - len(current_library(selected_industry)[:10])),
        score=score, warnings=warnings, token_warning=token_warning,
        analysis=analysis, plan=plan, v3=v3_ceo_summary(), pipeline_rows=get_pipeline_leads(), customer_tasks=get_customer_tasks(), notifications=get_notifications(), fb_groups=get_fb_groups(), group_schedules=get_group_schedules(), comment_leads=get_comment_leads(), messenger_scripts=get_messenger_scripts(), success_assets=get_success_assets()
    )

@app.route("/")
def home():
    selected_industry = request.args.get("industry", "spa")
    return render(selected_industry=selected_industry)

@app.route("/generate", methods=["POST"])
def generate():
    idea = request.form.get("idea", "").strip()
    style = request.form.get("style", "gần gũi")
    length = request.form.get("length", "120")
    variants = int(request.form.get("variants", "1"))
    if not idea:
        return render(message="Vui lòng nhập ý tưởng bài viết.", ok=False)
    try:
        content = generate_content(idea, style, length, variants=variants)
        return render(content=content, message="Đã tạo nội dung bằng Gemini Flash.", ok=True)
    except Exception as e:
        return render(message="Lỗi Gemini: " + str(e), ok=False)

@app.route("/analyze_fanpage", methods=["POST"])
def analyze_route():
    name = request.form.get("page_name", "")
    avatar = request.form.get("avatar", "")
    cover = request.form.get("cover", "")
    bio = request.form.get("bio", "")
    freq = request.form.get("post_frequency", "")
    cta = request.form.get("cta", "")
    score, notes = analyze_fanpage(name, avatar, cover, bio, freq, cta)
    analysis = f"Fanpage: {name}\nĐiểm đánh giá: {score}/100\n\nĐề xuất:\n- " + "\n- ".join(notes)
    return render(analysis=analysis, message="Đã phân tích Fanpage.", ok=True)

@app.route("/plan_30_days", methods=["POST"])
def plan_route():
    industry = request.form.get("industry", "spa")
    goal = request.form.get("goal", "tăng inbox")
    try:
        plan = ai_planner_v6(industry, goal, request.form.get("days", "30"))
        return render(plan=plan, message="Đã tạo kế hoạch marketing 30 ngày.", ok=True)
    except Exception as e:
        return render(message=friendly_ai_error(e), ok=False)

@app.route("/campaign", methods=["POST"])
def campaign_route():
    add_campaign(
        request.form.get("name","").strip(),
        request.form.get("industry","").strip(),
        request.form.get("goal","").strip(),
        request.form.get("note","").strip()
    )
    return render(message="Đã tạo chiến dịch.", ok=True)


def split_bulk_contents(text):
    raw = (text or "").replace("\r\n", "\n").strip()
    if not raw:
        return []

    # Ưu tiên tách theo dòng trống: mỗi đoạn là 1 content
    parts = [p.strip() for p in raw.split("\n\n") if p.strip()]

    # Nếu không có dòng trống, mỗi dòng là 1 content
    if len(parts) <= 1:
        lines = [x.strip() for x in raw.split("\n") if x.strip()]
        if len(lines) > 1:
            parts = lines

    return parts

def distribute_content_to_pages(contents, pages):
    jobs = []
    if not contents or not pages:
        return jobs

    # Ví dụ 10 content + 10 page => content 1 page 1, content 2 page 2...
    # Nếu content nhiều hơn page thì quay vòng page.
    for i, content in enumerate(contents):
        page = pages[i % len(pages)]
        jobs.append((page, content))
    return jobs


@app.route("/multi_post", methods=["POST"])
def multi_post():
    free = get_free_status()
    if free.get("is_expired"):
        return render(message="Phiên miễn phí đã giới hạn. Quý khách vui lòng nâng cấp Premium để tiếp tục đăng Fanpage.", ok=False)
    bulk_content = request.form.get("bulk_content", "").strip()
    action = request.form.get("action", "now")
    schedule_time = request.form.get("schedule_time", "").replace("T", " ")
    campaign = request.form.get("campaign", "").strip()
    pages = selected_pages(request.form.getlist("page_indexes"))
    if get_free_status().get("is_trial") and len(pages) > 1:
        return render(message="Gói miễn phí chỉ được đăng tối đa 1 Fanpage. Vui lòng nâng cấp Premium để đăng nhiều Fanpage cùng lúc.", ok=False)
    use_ai_enhance = request.form.get("use_ai_enhance") == "1"

    media_paths = save_uploads(request.files.getlist("images"))
    if not media_paths:
        single = save_upload(request.files.get("image"))
        media_paths = [single] if single else []

    contents = split_bulk_contents(bulk_content)

    if not contents:
        return render(message="Chưa nhập content để đăng.", ok=False)
    if not pages:
        return render(message="Chưa chọn Fanpage.", ok=False)

    # V9: kiểm tra token trước khi đăng ngay
    token_errors = []
    for page in pages:
        token_status = check_single_page_token(page)
        if token_status["status"] != "SỐNG":
            token_errors.append(f"{page.get('name')}: {token_status['detail']}")
    if token_errors and action == "now":
        return render(message="Token lỗi, hệ thống dừng đăng để tránh fail hàng loạt:\n" + "\n".join(token_errors), ok=False)

    jobs = distribute_content_to_pages(contents, pages)
    messages = []
    used_media = set()

    for idx, (page, content) in enumerate(jobs):
        # Mặc định giữ nguyên 100% content khách nhập.
        # Chỉ spin/thêm CTA/hashtag khi khách bật checkbox AI tối ưu.
        final_content = content
        if use_ai_enhance:
            final_content = ai_spin_content(content) if idx > 0 else content
            cta, hashtags = auto_cta_hashtag(final_content, "marketing")
            if cta.lower() not in final_content.lower():
                final_content += "\n\n" + cta
            if "#" not in final_content:
                final_content += "\n" + hashtags

        content_score = score_content(final_content)
        image_path = choose_best_media_for_content(final_content, media_paths, used_media)

        if action == "schedule":
            if not schedule_time:
                return render(message="Chưa chọn thời gian đặt lịch.", ok=False)
            save_post(page["name"], page["id"], final_content, "scheduled", "", schedule_time, image_path, campaign, content_score)
            media_note = f" | Media: {os.path.basename(image_path)}" if image_path else ""
            messages.append(f"Đã lưu lịch cho {page['name']}{media_note}")
        else:
            result = post_to_facebook(page, final_content, image_path)
            if "id" in result or "post_id" in result:
                post_id = result.get("post_id") or result.get("id")
                save_post(page["name"], page["id"], final_content, "posted", post_id, "", image_path, campaign, content_score)
                media_note = f" | Media: {os.path.basename(image_path)}" if image_path else ""
                messages.append(f"Đăng thành công {page['name']}: {post_id}{media_note}")
            else:
                save_post(page["name"], page["id"], final_content, "error", str(result), "", image_path, campaign, content_score)
                messages.append(f"Lỗi {page['name']}: {result}")

    return render(message="\n".join(messages), ok=True)

@app.route("/post", methods=["POST"])
def post():
    content = request.form.get("content", "").strip()
    action = request.form.get("action", "now")
    schedule_time = request.form.get("schedule_time", "").replace("T", " ")
    campaign = request.form.get("campaign", "").strip()
    pages = selected_pages(request.form.getlist("page_indexes"))
    image_path = save_upload(request.files.get("image"))
    if not content:
        return render(message="Nội dung trống.", ok=False)
    if not pages:
        return render(content=content, message="Chưa chọn Fanpage.", ok=False)

    if action == "now":
        token_errors = []
        for page in pages:
            token_status = check_single_page_token(page)
            if token_status["status"] != "SỐNG":
                token_errors.append(f"{page.get('name')}: {token_status['detail']}")
        if token_errors:
            return render(content=content, message="Token/Page chưa đủ quyền nên chưa đăng được:\n" + "\n".join(token_errors) + "\n\nCách xử lý: vào Token Fanpage kiểm tra token, đảm bảo Page Token còn sống và có quyền pages_manage_posts.", ok=False)

    messages = []
    content_score = score_content(content)
    for i, page in enumerate(pages):
        final_content = content if i == 0 else spin_content_local(content)
        if action == "schedule":
            if not schedule_time:
                return render(content=content, message="Chưa chọn thời gian đặt lịch.", ok=False)
            save_post(page["name"], page["id"], final_content, "scheduled", "", schedule_time, image_path, campaign, content_score)
            messages.append(f"Đã lưu lịch {schedule_time} cho {page['name']}")
        else:
            result = post_to_facebook(page, final_content, image_path)
            if "id" in result or "post_id" in result:
                post_id = result.get("post_id") or result.get("id")
                save_post(page["name"], page["id"], final_content, "posted", post_id, "", image_path, campaign, content_score)
                messages.append(f"Đăng thành công {page['name']}: {post_id}")
            else:
                save_post(page["name"], page["id"], final_content, "error", str(result), "", image_path, campaign, content_score)
                messages.append(f"Lỗi {page['name']}: {result}")
    return render(content=content, message="\\n".join(messages), ok=True)

@app.route("/batch", methods=["POST"])
def batch():
    file_obj = request.files.get("batch_file")
    if not file_obj:
        return render(message="Chưa chọn file Excel/CSV.", ok=False)
    try:
        rows = read_batch_file(file_obj)
        count = 0
        for row in rows:
            idea = str(row.get("idea", "") or "").strip()
            content = str(row.get("content", "") or "").strip()
            page_names = str(row.get("page_names", "") or "").strip()
            schedule_time = str(row.get("schedule_time", "") or "").strip()
            campaign = str(row.get("campaign", "") or "").strip()
            if not content and idea:
                content = generate_content(idea, "chuyên nghiệp", "120", variants=1)
            if not content or not schedule_time:
                continue
            all_pages = get_fanpages()
            target_pages = []
            for name in [x.strip().lower() for x in page_names.split(",") if x.strip()]:
                for p in all_pages:
                    if name in p["name"].lower():
                        target_pages.append(p)
            if not target_pages and all_pages:
                # Nếu file không có page_names, tự chia mỗi dòng content sang Page kế tiếp để tránh trùng bài
                target_pages = [all_pages[count % len(all_pages)]]

            # File Excel/CSV có thể thêm cột image_path để gắn ảnh riêng từng bài.
            image_path = str(row.get("image_path", "") or row.get("media_path", "") or "").strip()
            if image_path and not os.path.exists(image_path):
                image_path = ""

            for page in target_pages:
                save_post(page["name"], page["id"], content, "scheduled", "", schedule_time, image_path, campaign, score_content(content))
                count += 1
        return render(message=f"Đã nhập và lưu lịch {count} bài từ file.", ok=True)
    except Exception as e:
        return render(message="Lỗi đọc file: " + str(e), ok=False)


@app.route("/fanpage_token_add", methods=["POST"])
def fanpage_token_add_route():
    ok, msg = add_fanpage_token(
        request.form.get("page_name", ""),
        request.form.get("page_id", ""),
        request.form.get("page_token", ""),
        request.form.get("note", "")
    )
    return render(message=msg, ok=ok)


@app.route("/fanpage_token_delete", methods=["POST"])
def fanpage_token_delete_route():
    ok, msg = delete_fanpage_token(request.form.get("row_id", ""))
    return render(message=msg, ok=ok)


@app.route("/check_tokens", methods=["POST"])
def check_tokens_route():
    if not get_fanpages():
        return render(message="Chưa có Fanpage. Hãy thêm trực tiếp trong Token Center.", ok=False)

    results = check_all_page_tokens()
    alive = sum(1 for x in results if x["status"] == "SỐNG")
    dead = len(results) - alive
    message = f"Đã kiểm tra {len(results)} Page Token. Sống: {alive}. Lỗi/Giới hạn: {dead}."
    return render(message=message, ok=(dead == 0))

@app.route("/content_factory", methods=["POST"])
def content_factory_route():
    idea = request.form.get("idea", "").strip()
    count = request.form.get("count", "20")
    style = request.form.get("style", "bán hàng tự nhiên")

    if not idea:
        return render(message="Chưa nhập chủ đề để tạo content hàng loạt.", ok=False)

    try:
        content = generate_many_contents(idea, count, style)
        return render(content=content, message=f"Đã tạo {count} content. Nội dung đã hiển thị trong khung kết quả AI.", ok=True)
    except Exception as e:
        return render(message=friendly_ai_error(e), ok=False)

@app.route("/smart_schedule", methods=["POST"])
def smart_schedule_route():
    bulk_content = request.form.get("bulk_content", "").strip()
    start_time = request.form.get("start_time", "")
    gap_minutes = request.form.get("gap_minutes", "60")
    campaign = request.form.get("campaign", "").strip()
    pages = selected_pages(request.form.getlist("page_indexes"))
    media_paths = save_uploads(request.files.getlist("images"))
    use_ai_enhance = request.form.get("use_ai_enhance") == "1"

    contents = split_bulk_contents(bulk_content)

    if not contents:
        return render(message="Chưa nhập content để chia lịch.", ok=False)
    if not pages:
        return render(message="Chưa chọn Fanpage.", ok=False)

    jobs = distribute_content_to_pages(contents, pages)
    times = smart_schedule_times(start_time, len(jobs), gap_minutes)
    used_media = set()

    for i, (page, content) in enumerate(jobs):
        # Mặc định giữ nguyên 100% content khách nhập.
        final_content = content
        if use_ai_enhance:
            final_content = ai_spin_content(content) if i > 0 else content
            cta, hashtags = auto_cta_hashtag(final_content, "marketing")
            if cta.lower() not in final_content.lower():
                final_content += "\n\n" + cta
            if "#" not in final_content:
                final_content += "\n" + hashtags

        image_path = choose_best_media_for_content(final_content, media_paths, used_media)
        save_post(page["name"], page["id"], final_content, "scheduled", "", times[i], image_path, campaign, score_content(final_content))

    msg = f"Đã tự chia lịch {len(jobs)} bài. Content được giữ nguyên theo nội dung đã nhập."
    if use_ai_enhance:
        msg = f"Đã tự chia lịch {len(jobs)} bài và bật AI tối ưu từng bài. Khoảng cách: {gap_minutes} phút."
    return render(message=msg, ok=True)

@app.route("/page_cluster", methods=["POST"])
def page_cluster_route():
    add_page_cluster(
        request.form.get("name", "").strip(),
        request.form.get("page_names", "").strip(),
        request.form.get("note", "").strip()
    )
    return render(message="Đã lưu nhóm Page.", ok=True)


@app.route("/crm", methods=["POST"])
def crm_route():
    add_crm(
        request.form.get("name","").strip(),
        request.form.get("phone","").strip(),
        request.form.get("zalo","").strip(),
        request.form.get("source","").strip(),
        request.form.get("note","").strip()
    )
    return render(message="Đã lưu khách hàng vào CRM Mini.", ok=True)


@app.route("/v3_ai_tool", methods=["POST"])
def v3_ai_tool_route():
    tool = request.form.get("tool", "marketing_director")
    topic = request.form.get("topic", "").strip()
    extra = request.form.get("extra", "").strip()
    if not topic:
        return render(message="Vui lòng nhập nội dung cần AI xử lý.", ok=False)
    prompt = v3_ai_tool_prompt(tool, topic, extra)
    result = safe_ai_generate(prompt, fallback=f"Bản demo cho {topic}:\n\n- 30 content\n- 10 quảng cáo\n- Tệp khách hàng mục tiêu\n- CTA và kế hoạch triển khai\n\nVui lòng cấu hình GEMINI_API_KEY để dùng AI đầy đủ.")
    return render(content=result, message="Đã tạo nội dung bằng AI Studio V3.", ok=True)

@app.route("/pipeline", methods=["POST"])
def pipeline_route():
    add_pipeline_lead(
        request.form.get("customer_name", "").strip(),
        request.form.get("phone", "").strip(),
        request.form.get("zalo", "").strip(),
        request.form.get("source", "").strip(),
        request.form.get("stage", "Khách mới").strip(),
        request.form.get("value", "0").strip() or 0,
        request.form.get("note", "").strip()
    )
    return render(message="Đã thêm khách vào CRM Sales Pipeline.", ok=True)

@app.route("/customer_task", methods=["POST"])
def customer_task_route():
    add_customer_task(request.form.get("customer_name", "").strip(), request.form.get("task", "").strip(), request.form.get("due_date", "").strip())
    return render(message="Đã tạo lịch chăm sóc khách hàng.", ok=True)

@app.route("/notification", methods=["POST"])
def notification_route():
    add_notification(request.form.get("title", "").strip(), request.form.get("detail", "").strip(), "info")
    return render(message="Đã thêm thông báo vào Notification Center.", ok=True)


@app.route("/fb_group", methods=["POST"])
def fb_group_route():
    add_fb_group(request.form.get("group_name", "").strip(), request.form.get("group_id", "").strip(), request.form.get("niche", "").strip(), request.form.get("note", "").strip())
    return render(message="Đã lưu Group vào Group Marketing.", ok=True)

@app.route("/group_schedule", methods=["POST"])
def group_schedule_route():
    add_group_schedule(request.form.get("group_name", "").strip(), request.form.get("group_id", "").strip(), request.form.get("content", "").strip(), request.form.get("schedule_time", "").replace("T", " "))
    return render(message="Đã lưu lịch đăng Group.", ok=True)

@app.route("/comment_ai", methods=["POST"])
def comment_ai_route():
    customer_name = request.form.get("customer_name", "").strip()
    phone = request.form.get("phone", "").strip()
    comment_text = request.form.get("comment_text", "").strip()
    label = request.form.get("label", "Khách nóng").strip()
    if not comment_text:
        return render(message="Vui lòng nhập comment khách hàng.", ok=False)
    clean_comment = hide_phone_preview(comment_text)
    prompt = v3_ai_tool_prompt('comment_reply', clean_comment, f"Nhãn khách: {label}")
    fallback = f"Dạ em đã nhận thông tin. Anh/chị vui lòng inbox để bên em tư vấn chi tiết và hỗ trợ nhanh hơn.\n\nNhãn gợi ý: {label}\nHành động: Chuyển CRM"
    ai_reply = safe_ai_generate(prompt, fallback=fallback)
    add_comment_lead(customer_name, phone, clean_comment, ai_reply, label)
    if request.form.get("to_crm", "1") == "1":
        add_pipeline_lead(customer_name or "Khách từ Comment", phone, "", "Comment Facebook", "Khách mới", 0, clean_comment)
    return render(content=ai_reply, message="Đã xử lý comment, ẩn SĐT và chuyển CRM nếu được chọn.", ok=True)

@app.route("/messenger_ai_script", methods=["POST"])
def messenger_ai_script_route():
    product = request.form.get("product", "").strip()
    script_type = request.form.get("script_type", "Kịch bản Inbox").strip()
    extra = request.form.get("extra", "").strip()
    if not product:
        return render(message="Vui lòng nhập sản phẩm/dịch vụ để tạo kịch bản Messenger.", ok=False)
    prompt = v3_ai_tool_prompt('messenger_ai', product, f"Loại kịch bản: {script_type}. {extra}")
    fallback = f"{script_type} cho {product}:\n1. Chào khách tự nhiên.\n2. Hỏi nhu cầu chính.\n3. Tư vấn gói phù hợp.\n4. Xử lý băn khoăn về giá.\n5. Chốt hành động: inbox/Zalo/thanh toán."
    content = safe_ai_generate(prompt, fallback=fallback)
    add_messenger_script(product, script_type, content)
    return render(content=content, message="Đã tạo kịch bản Messenger AI.", ok=True)

@app.route("/export")
def export_route():
    path = export_csv()
    return send_file(path, as_attachment=True)


@app.route("/backup")
def backup_route():
    path = backup_database()
    if not path:
        return render(message="Không tìm thấy database để backup.", ok=False)
    return send_file(path, as_attachment=True)

@app.route("/export_pdf")
def export_pdf_route():
    path = export_pdf_report()
    return send_file(path, as_attachment=True)


@app.route("/api/templates")
def api_templates():
    industry = request.args.get("industry", "spa")
    return jsonify(current_library(industry))




def ensure_support_tables():
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS support_messages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        device_id TEXT,
        phone TEXT,
        email TEXT,
        sender TEXT DEFAULT 'customer',
        message TEXT,
        admin_reply TEXT,
        status TEXT DEFAULT 'new',
        created_at TEXT,
        replied_at TEXT
    )
    """)
    conn.commit()
    conn.close()


def add_support_message(device_id, phone, email, message):
    ensure_support_tables()
    device_id = (device_id or '').strip()
    phone = (phone or '').strip()
    email = (email or '').strip().lower()
    message = (message or '').strip()
    if not message:
        return False, 'Vui lòng nhập nội dung cần hỗ trợ.'
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    INSERT INTO support_messages(device_id,phone,email,sender,message,status,created_at)
    VALUES(?,?,?,?,?,?,?)
    """, (device_id, phone, email, 'customer', message, 'new', now))
    conn.commit()
    conn.close()
    return True, 'Đã gửi tin nhắn đến Admin. Vui lòng chờ phản hồi trong khung chat này.'


def get_support_messages(limit=120):
    ensure_support_tables()
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    SELECT id,device_id,phone,email,sender,message,admin_reply,status,created_at,replied_at
    FROM support_messages
    ORDER BY id DESC LIMIT ?
    """, (limit,))
    rows = c.fetchall()
    conn.close()
    return rows


def reply_support_message(message_id, admin_reply):
    ensure_support_tables()
    admin_reply = (admin_reply or '').strip()
    if not admin_reply:
        return False, 'Vui lòng nhập nội dung trả lời.'
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    UPDATE support_messages
    SET admin_reply=?, status='replied', replied_at=?
    WHERE id=?
    """, (admin_reply, now, int(message_id or 0)))
    ok = c.rowcount > 0
    conn.commit()
    conn.close()
    return ok, 'Đã trả lời khách hàng.' if ok else 'Không tìm thấy tin nhắn.'


def get_support_replies_for_device(device_id, after_id=0):
    ensure_support_tables()
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""
    SELECT id,message,admin_reply,status,created_at,replied_at
    FROM support_messages
    WHERE device_id=? AND id>?
    ORDER BY id ASC LIMIT 50
    """, ((device_id or '').strip(), int(after_id or 0)))
    rows = c.fetchall()
    conn.close()
    return rows


@app.post('/support_message')
def support_message_route():
    data = request.get_json(silent=True) or {}
    ok, msg = add_support_message(data.get('device_id',''), data.get('phone',''), data.get('email',''), data.get('message',''))
    return jsonify({'success': ok, 'message': msg})


@app.get('/support_poll')
def support_poll_route():
    device_id = request.args.get('device_id','')
    after_id = request.args.get('after_id','0')
    rows = get_support_replies_for_device(device_id, after_id)
    return jsonify({'messages': [
        {'id': r[0], 'message': r[1], 'admin_reply': r[2], 'status': r[3], 'created_at': r[4], 'replied_at': r[5]}
        for r in rows
    ]})


@app.post("/premium_request")
def premium_request_route():
    data = request.get_json(silent=True) or {}
    ok, msg = create_premium_request(
        data.get("device_id", ""),
        data.get("phone", ""),
        data.get("email", ""),
        data.get("plan_key", "monthly"),
        data.get("plan_name", "Gói Premium"),
        data.get("amount", 0),
        data.get("transaction_note", "")
    )
    return jsonify({"success": ok, "message": msg})


@app.get("/premium_status")
def premium_status_route():
    device_id = request.args.get("device_id", "")
    return jsonify(get_premium_status_by_device(device_id))


@app.route("/admin", methods=["GET", "POST"])
def admin_premium_page():
    admin_password = os.getenv("ADMIN_PASSWORD", "admin123")
    password = request.args.get("password") or request.form.get("password") or ""
    if password != admin_password:
        return """
        <html><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Admin Login</title></head>
        <body style='font-family:Arial;background:#0f172a;color:white;padding:30px'>
        <h2>Đăng nhập Web Admin</h2>
        <form method='get'>
          <input name='password' type='password' placeholder='Mật khẩu admin' style='padding:12px;border-radius:10px;width:260px;background:#111827;color:white;border:1px solid #334155'>
          <button style='padding:12px 18px;border-radius:10px;background:#7c3aed;color:white;border:0;font-weight:bold'>Vào Admin</button>
        </form>
        </body></html>
        """

    rows = get_premium_requests(500)
    chats = get_support_messages(300)
    token_checks = get_latest_token_checks(120)
    history = get_history(500)
    stats = get_stats()
    fanpages = get_fanpages()

    pending_count = sum(1 for r in rows if r[8] == 'pending')
    approved_count = sum(1 for r in rows if r[8] == 'approved')
    chat_new_count = sum(1 for r in chats if r[7] == 'new')
    error_posts = [r for r in history if str(r[3]).lower() == 'error']
    posted_count = sum(1 for r in history if str(r[3]).lower() == 'posted')
    scheduled_count = sum(1 for r in history if str(r[3]).lower() == 'scheduled')
    total_revenue = sum(int(r[6] or 0) for r in rows if r[8] == 'approved')
    month_now = datetime.datetime.now().strftime('%Y-%m')
    month_revenue = sum(int(r[6] or 0) for r in rows if r[8] == 'approved' and str(r[12] or '').startswith(month_now))

    plan_counts = {}
    for r in rows:
        if r[8] == 'approved':
            name = r[5] or r[4] or 'Gói Premium'
            plan_counts[name] = plan_counts.get(name, 0) + 1

    def esc(x):
        return str(x or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', '&quot;')

    def money(v):
        try:
            return f"{int(v or 0):,}đ".replace(',', '.')
        except Exception:
            return '0đ'

    now_dt = datetime.datetime.now()
    expiring = []
    for r in rows:
        if r[8] == 'approved' and r[10]:
            try:
                end_dt = datetime.datetime.strptime(str(r[10])[:19], '%Y-%m-%d %H:%M:%S')
                if 0 <= (end_dt - now_dt).days <= 7:
                    expiring.append(r)
            except Exception:
                pass

    html_rows = ""
    payment_rows = ""
    premium_customer_rows = ""
    for r in rows:
        status = r[8] or "pending"
        badge_class = "badge-ok" if status == "approved" else "badge-warn"
        approve_btn = ""
        if status != "approved":
            approve_btn = f"""
            <form method='post' action='/admin/approve_premium' style='margin:0'>
              <input type='hidden' name='password' value='{esc(password)}'>
              <input type='hidden' name='request_id' value='{r[0]}'>
              <button class='btn success'>Kích hoạt</button>
            </form>
            """
        else:
            approve_btn = "<span class='muted'>Đã duyệt</span>"
        plan_name = esc(r[5] or r[4] or 'Gói Premium')
        amount = money(r[6])
        html_rows += f"""
        <tr>
          <td>{r[0]}</td><td><b>{esc(r[1])}</b></td><td>{esc(r[2])}</td><td>{esc(r[3])}</td>
          <td><b class='plan-name'>👑 {plan_name}</b><br><small>{amount}</small></td>
          <td>{esc(r[7])}</td><td><span class='badge {badge_class}'>{esc(status)}</span><br><small>{esc(r[11])}</small></td><td>{esc(r[10] or '')}</td><td>{approve_btn}</td>
        </tr>
        """
        payment_rows += f"""
        <tr><td>{r[0]}</td><td>{esc(r[2])}</td><td>{esc(r[3])}</td><td>{plan_name}</td><td><b>{amount}</b></td><td>{esc(r[7])}</td><td><span class='badge {badge_class}'>{esc(status)}</span></td><td>{esc(r[12] or '')}</td></tr>
        """
        if status == 'approved':
            premium_customer_rows += f"""
            <tr><td>{r[0]}</td><td><b>{esc(r[1])}</b></td><td>{esc(r[2])}</td><td>{esc(r[3])}</td><td><b class='plan-name'>{plan_name}</b></td><td>{amount}</td><td>{esc(r[10] or '')}</td></tr>
            """

    plan_cards = ''.join(f"<div class='mini-card'><span>{esc(k)}</span><b>{v}</b></div>" for k,v in sorted(plan_counts.items())) or "<div class='mini-card'><span>Chưa có gói đã duyệt</span><b>0</b></div>"

    chat_rows = ""
    for m in chats:
        status = m[7] or 'new'
        badge_class = 'badge-warn' if status == 'new' else 'badge-ok'
        customer_key = esc(m[2] or m[3] or m[1] or 'Khách chưa có thông tin')
        current_reply = esc(m[6] or '')
        reply_box = ''
        if current_reply:
            reply_box += f"<div class='reply-done'><b>Admin đã trả lời gần nhất:</b><br>{current_reply}<br><small>{esc(m[9])}</small></div>"
        reply_box += f"""
            <form method='post' action='/admin/reply_support' class='reply-form'>
              <input type='hidden' name='password' value='{esc(password)}'>
              <input type='hidden' name='message_id' value='{m[0]}'>
              <textarea name='admin_reply' placeholder='Nhập nội dung trả lời hoặc cập nhật câu trả lời...' required>{current_reply}</textarea>
              <button type='submit' class='btn success'>Gửi / cập nhật trả lời</button>
            </form>
            """
        chat_rows += f"""
        <div class='chat-card'>
          <div class='chat-head'><b>#{m[0]} - {customer_key}</b><span class='badge {badge_class}'>{esc(status)}</span></div>
          <small>Device: {esc(m[1])} • SĐT: {esc(m[2])} • Email: {esc(m[3])} • {esc(m[8])}</small>
          <div class='customer-msg'>{esc(m[5])}</div>
          {reply_box}
        </div>
        """

    token_rows = ''.join(
        f"<tr><td>{esc(t[0])}</td><td>{esc(t[1])}</td><td><span class='badge {'badge-ok' if str(t[2]).upper()=='SỐNG' else 'badge-danger'}'>{esc(t[2])}</span></td><td>{esc(t[3])}</td><td>{esc(t[4])}</td></tr>" for t in token_checks
    )
    error_rows = ''.join(f"<tr><td>{e[0]}</td><td>{esc(e[1])}</td><td><span class='badge badge-danger'>{esc(e[3])}</span></td><td>{esc(e[4])}</td><td>{esc(e[9])}</td></tr>" for e in error_posts[:120])
    expiring_rows = ''.join(f"<tr><td>{r[0]}</td><td>{esc(r[2])}</td><td>{esc(r[3])}</td><td>{esc(r[5] or r[4])}</td><td>{esc(r[10])}</td></tr>" for r in expiring)

    return f"""
    <html><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Admin Center</title>
    <style>
      *{{box-sizing:border-box}} html{{scroll-behavior:smooth}} body{{font-family:Arial,system-ui;background:#0f172a;color:#e5e7eb;margin:0}} a{{color:inherit;text-decoration:none}}
      .layout{{display:flex;min-height:100vh}} .admin-side{{width:282px;background:#08111f;border-right:1px solid #1e293b;padding:20px;position:sticky;top:0;height:100vh;overflow-y:auto}}
      .brand{{font-size:26px;font-weight:900;color:#38bdf8;text-shadow:0 2px 0 #075985,0 0 18px rgba(56,189,248,.35);margin-bottom:18px}}
      .menu-title{{color:#94a3b8;font-size:12px;letter-spacing:.12em;text-transform:uppercase;margin:18px 0 8px}}
      .admin-side a{{display:block;padding:13px 14px;border-radius:14px;margin:7px 0;background:#111827;border:1px solid #1f2937;font-weight:700}}
      .admin-side a:hover{{background:#1e293b;color:#38bdf8;transform:translateX(3px)}} .main{{flex:1;padding:26px;overflow:auto}}
      h1{{color:#38bdf8;font-size:38px;text-shadow:0 2px 0 #075985,0 0 18px rgba(56,189,248,.35);margin:0 0 8px}} h2{{color:#c4b5fd;margin-top:38px;font-size:28px}}
      .sub{{color:#94a3b8;margin-bottom:18px}} .cards{{display:grid;grid-template-columns:repeat(6,minmax(140px,1fr));gap:14px;margin:18px 0}}
      .card{{background:linear-gradient(135deg,#1e293b,#111827);border:1px solid #334155;border-radius:20px;padding:18px;min-height:104px;box-shadow:0 18px 50px rgba(0,0,0,.18)}} .card span{{color:#cbd5e1;font-weight:700}} .card b{{display:block;font-size:30px;color:#fbbf24;margin-top:8px}}
      .grid2{{display:grid;grid-template-columns:1.1fr .9fr;gap:16px}} .panel{{background:#111827;border:1px solid #334155;border-radius:20px;padding:16px;margin:16px 0}}
      .mini-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:12px}} .mini-card{{background:#0b1224;border:1px solid #263449;border-radius:16px;padding:14px}} .mini-card span{{display:block;color:#cbd5e1}} .mini-card b{{font-size:26px;color:#38bdf8}}
      table{{width:100%;border-collapse:collapse;background:#111827;border-radius:18px;overflow:hidden;margin-bottom:20px}} th,td{{border-bottom:1px solid #334155;padding:12px;text-align:left;vertical-align:top}} th{{background:#1e1b4b;color:#c4b5fd;position:sticky;top:0}}
      .btn{{border:0;padding:10px 14px;border-radius:12px;font-weight:900;cursor:pointer}} .success{{background:#22c55e;color:white}} .primary{{background:#2563eb;color:white}} small,.muted{{color:#94a3b8}}
      .badge{{display:inline-block;color:white;padding:6px 10px;border-radius:999px;font-weight:900;font-size:13px}} .badge-ok{{background:#22c55e}} .badge-warn{{background:#f59e0b}} .badge-danger{{background:#ef4444}}
      .plan-name{{color:#38bdf8;text-shadow:0 1px 0 #075985,0 0 13px rgba(59,130,246,.65)}} .tools{{display:flex;gap:10px;flex-wrap:wrap;margin:8px 0 16px}}
      .chat-card{{background:#111827;border:1px solid #334155;border-radius:18px;padding:16px;margin:12px 0}} .chat-head{{display:flex;justify-content:space-between;gap:10px;align-items:center}}
      .customer-msg{{background:#0b1224;border:1px solid #1f2937;border-radius:14px;padding:12px;margin:10px 0;white-space:pre-wrap;line-height:1.5}} .reply-form textarea{{width:100%;height:90px;background:#020617;color:white;border:1px solid #334155;border-radius:12px;padding:12px;margin:8px 0}} .reply-done{{background:#052e16;border:1px solid #166534;border-radius:14px;padding:12px;margin-top:10px}}
      @media(max-width:1100px){{.cards{{grid-template-columns:repeat(2,1fr)}}.grid2{{grid-template-columns:1fr}}}} @media(max-width:760px){{.layout{{display:block}}.admin-side{{width:100%;height:auto;position:relative}}.cards{{grid-template-columns:1fr}}}}
    </style></head><body>
      <div class='layout'>
        <aside class='admin-side'>
          <div class='brand'>Admin Center</div>
          <div class='menu-title'>CEO</div><a href='#overview'>🏠 Dashboard CEO</a>
          <div class='menu-title'>Premium Center</div><a href='#premium'>👑 Yêu cầu kích hoạt</a><a href='#customers'>💎 Khách Premium</a><a href='#expiring'>⏳ Hết hạn sắp tới</a><a href='#payments'>💳 Doanh thu & Thanh toán</a>
          <div class='menu-title'>Chat Center</div><a href='#chat'>💬 Chat khách hàng <span class='badge badge-warn'>{chat_new_count}</span></a>
          <div class='menu-title'>Facebook Center</div><a href='#errors'>📣 Lỗi đăng Fanpage</a><a href='#tokens'>🔑 Token Fanpage</a>
          <div class='menu-title'>Cài đặt App</div><a href='#install_app'>📲 Cài app nhanh</a><a href='/install' target='_blank'>📲 Cài đặt GPT MKT Pro</a>
          <div class='menu-title'>System</div><a href='/admin?password={esc(password)}'>🔄 Tải lại Admin</a>
        </aside>
        <main class='main'>
          <h1 id='overview'>Dashboard CEO</h1><div class='sub'>Quản lý Premium, thanh toán, chat khách hàng, lỗi đăng bài và token Fanpage.</div>
          <div class='cards'>
            <div class='card'><span>Chờ duyệt</span><b>{pending_count}</b></div><div class='card'><span>Đã kích hoạt</span><b>{approved_count}</b></div><div class='card'><span>Doanh thu tháng</span><b>{money(month_revenue)}</b></div><div class='card'><span>Tổng doanh thu</span><b>{money(total_revenue)}</b></div><div class='card'><span>Tin chat mới</span><b>{chat_new_count}</b></div><div class='card'><span>Lỗi đăng bài</span><b>{len(error_posts)}</b></div>
          </div>
          <div class='grid2'>
            <div class='panel'><h3>👑 Thống kê gói Premium</h3><div class='mini-grid'>{plan_cards}</div></div>
            <div class='panel'><h3>📣 Tình trạng hệ thống</h3><div class='mini-grid'><div class='mini-card'><span>Fanpage đã lưu</span><b>{len(fanpages)}</b></div><div class='mini-card'><span>Đã đăng</span><b>{posted_count}</b></div><div class='mini-card'><span>Đang hẹn lịch</span><b>{scheduled_count}</b></div><div class='mini-card'><span>CRM</span><b>{stats.get('crm',0)}</b></div></div></div>
          </div>
          <h2 id='install_app'>📲 Cài đặt GPT MKT Pro</h2>
          <div class='grid2'>
            <div class='panel'>
              <h3>📱 Thêm ứng dụng vào màn hình chính</h3>
              <div style='line-height:1.8;color:#cbd5e1'>
                <b style='color:#fbbf24'>Không dùng QR, khách chỉ cần bấm nút cài đặt.</b><br>
                Android Chrome/Edge: nếu đủ điều kiện, hệ thống tự mở popup cài app.<br>
                iPhone/iPad: hệ thống hiện đúng hướng dẫn Safari → Chia sẻ → Thêm vào Màn hình chính.<br>
                <a class='btn primary' href='/install' target='_blank' style='display:inline-block;margin-top:12px'>📲 Cài đặt GPT MKT Pro</a>
              </div>
            </div>
            <div class='panel'>
              <h3>⚡ Trải nghiệm khách hàng</h3>
              <div class='mini-grid'>
                <div class='mini-card'><span>Android</span><b>Popup</b><span>Bấm nút là mở hộp cài nếu Chrome/Edge đủ điều kiện.</span></div>
                <div class='mini-card'><span>iPhone</span><b>3 bước</b><span>Safari → Chia sẻ → Thêm vào màn hình chính.</span></div>
                <div class='mini-card'><span>Sau cài</span><b>Icon</b><span>GPT MKT Pro xuất hiện ngoài màn hình chính.</span></div>
              </div>
            </div>
          </div>

          <h2 id='premium'>👑 Yêu cầu Premium</h2>
          <table><tr><th>ID</th><th>Device ID</th><th>SĐT</th><th>Email</th><th>Gói</th><th>Ghi chú</th><th>Trạng thái</th><th>Hết hạn</th><th>Duyệt</th></tr>{html_rows or '<tr><td colspan="9">Chưa có yêu cầu Premium.</td></tr>'}</table>
          <h2 id='customers'>💎 Khách đang Premium</h2>
          <table><tr><th>ID</th><th>Device ID</th><th>SĐT</th><th>Email</th><th>Gói</th><th>Số tiền</th><th>Hết hạn</th></tr>{premium_customer_rows or '<tr><td colspan="7">Chưa có khách Premium.</td></tr>'}</table>
          <h2 id='expiring'>⏳ Premium hết hạn trong 7 ngày</h2>
          <table><tr><th>ID</th><th>SĐT</th><th>Email</th><th>Gói</th><th>Hết hạn</th></tr>{expiring_rows or '<tr><td colspan="5">Chưa có khách sắp hết hạn.</td></tr>'}</table>
          <h2 id='payments'>💳 Lịch sử thanh toán</h2>
          <table><tr><th>ID</th><th>SĐT</th><th>Email</th><th>Gói</th><th>Số tiền</th><th>Nội dung CK</th><th>Trạng thái</th><th>Ngày duyệt</th></tr>{payment_rows or '<tr><td colspan="8">Chưa có lịch sử thanh toán.</td></tr>'}</table>
          <h2 id='chat'>💬 Chat khách hàng</h2>
          <div class='panel'><b>Quy trình:</b> Khách nhắn trong Mini Chat Support → lưu vào Admin → Admin trả lời → khách nhận trong khung chat.</div>
          {chat_rows or '<div class="panel">Chưa có tin nhắn hỗ trợ.</div>'}
          <h2 id='errors'>📣 Lỗi đăng Fanpage</h2>
          <table><tr><th>ID</th><th>Fanpage</th><th>Trạng thái</th><th>Chi tiết lỗi</th><th>Thời gian</th></tr>{error_rows or '<tr><td colspan="5">Chưa có lỗi đăng bài.</td></tr>'}</table>
          <h2 id='tokens'>🔑 Token Fanpage</h2>
          <div class='tools'><form method='post' action='/admin/check_tokens'><input type='hidden' name='password' value='{esc(password)}'><button class='btn primary'>Kiểm tra tất cả Token</button></form></div>
          <table><tr><th>Fanpage</th><th>Page ID</th><th>Trạng thái</th><th>Chi tiết</th><th>Thời gian</th></tr>{token_rows or '<tr><td colspan="5">Chưa có dữ liệu kiểm tra token.</td></tr>'}</table>
        </main>
      </div>
    </body></html>
    """


@app.post("/admin/approve_premium")
def admin_approve_premium_route():
    admin_password = os.getenv("ADMIN_PASSWORD", "admin123")
    password = request.form.get("password", "")
    if password != admin_password:
        return "Sai mật khẩu admin.", 403
    ok, msg = approve_premium_request(request.form.get("request_id", "0"))
    return f"<meta charset='UTF-8'><script>alert({json.dumps(msg)});location.href='/admin?password={password}';</script>"



@app.post('/admin/reply_support')
def admin_reply_support_route():
    admin_password = os.getenv('ADMIN_PASSWORD', 'admin123')
    password = request.form.get('password', '')
    if password != admin_password:
        return 'Sai mật khẩu admin.', 403
    ok, msg = reply_support_message(request.form.get('message_id', '0'), request.form.get('admin_reply', ''))
    from urllib.parse import quote
    return f"<meta charset='UTF-8'><script>alert({json.dumps(msg)});location.href='/admin?password={quote(password)}#chat';</script>"



@app.get('/install_qr')
def install_qr():
    """QR cài đặt nhanh trỏ về trang /install."""
    try:
        import qrcode
        install_url = request.host_url.rstrip('/') + '/install'
        qr = qrcode.QRCode(version=1, box_size=8, border=2)
        qr.add_data(install_url)
        qr.make(fit=True)
        img = qr.make_image(fill_color='#0f172a', back_color='white').convert('RGB')
        bio = io.BytesIO()
        img.save(bio, 'PNG')
        bio.seek(0)
        return send_file(bio, mimetype='image/png', download_name='mkt-automation-pro-install-qr.png')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.get('/install')
def install_page():
    """Trang cài đặt PWA thông minh, bỏ QR: Android gọi prompt; iPhone hướng dẫn thêm màn hình chính."""
    app_url = request.host_url.rstrip('/')
    return f"""
<!doctype html>
<html lang='vi'>
<head>
  <meta charset='utf-8'>
  <meta name='viewport' content='width=device-width, initial-scale=1'>
  <title>Cài đặt GPT MKT Pro</title>
  <link rel='manifest' href='/manifest.json'>
  <meta name='theme-color' content='#2563eb'>
  <link rel='apple-touch-icon' href='/pwa-icon-192.png?v=gpt-mkt-pro-1'>
  <link rel='preconnect' href='https://fonts.googleapis.com'>
  <link rel='preconnect' href='https://fonts.gstatic.com' crossorigin>
  <link href='https://fonts.googleapis.com/css2?family=Inter:wght@500;600;700;800;900&family=Manrope:wght@500;600;700;800&display=swap' rel='stylesheet'>
  <style>
    *{{box-sizing:border-box}}
    body{{margin:0;min-height:100vh;font-family:'Manrope','Inter',system-ui;background:radial-gradient(circle at top,#1e3a8a 0,#0f172a 42%,#020617 100%);color:#e5e7eb;display:flex;align-items:center;justify-content:center;padding:18px}}
    .wrap{{width:min(760px,100%);}}
    .hero{{background:rgba(15,23,42,.78);border:1px solid rgba(148,163,184,.22);box-shadow:0 30px 90px rgba(0,0,0,.38);backdrop-filter:blur(20px);border-radius:32px;padding:28px;text-align:left}}
    .badge{{display:inline-flex;gap:8px;align-items:center;background:rgba(34,197,94,.12);border:1px solid rgba(34,197,94,.32);color:#86efac;border-radius:999px;padding:8px 12px;font-weight:900;font-size:13px}}
    h1{{font-size:clamp(34px,8vw,60px);line-height:1.03;margin:18px 0 12px;font-weight:900;letter-spacing:-2px;background:linear-gradient(90deg,#38bdf8,#818cf8,#c084fc,#facc15);-webkit-background-clip:text;-webkit-text-fill-color:transparent}}
    .lead{{color:#cbd5e1;font-size:17px;line-height:1.8;margin-bottom:18px}}
    .benefits{{display:grid;gap:10px;margin:20px 0}}
    .benefits div{{background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.09);border-radius:16px;padding:12px 14px;font-weight:800;color:#f8fafc}}
    .btn{{width:100%;border:0;border-radius:18px;padding:17px 18px;font-size:17px;font-weight:900;cursor:pointer;background:linear-gradient(135deg,#22c55e,#16a34a);color:white;box-shadow:0 18px 40px rgba(34,197,94,.25);margin-top:10px}}
    .btn.secondary{{background:linear-gradient(135deg,#2563eb,#7c3aed)}}
    .btn:disabled{{opacity:.65;cursor:not-allowed}}
    .status{{margin-top:14px;line-height:1.65;color:#bfdbfe;font-weight:800;background:rgba(37,99,235,.12);border:1px solid rgba(147,197,253,.18);padding:12px;border-radius:16px}}
    .steps{{display:none;margin-top:16px;background:#020617;border:1px solid #334155;border-radius:20px;padding:16px;line-height:1.9;color:#dbeafe}}
    .steps.show{{display:block}}
    .steps b{{color:#facc15}}
    .tiny{{font-size:13px;color:#94a3b8;line-height:1.7;margin-top:14px;text-align:center}}
    .installed{{background:linear-gradient(135deg,#0f766e,#14b8a6)!important}}
  </style>
</head>
<body>
  <div class='wrap'>
    <section class='hero'>
      <div class='badge'>🟢 App Mini sẵn sàng</div>
      <h1>Cài đặt GPT MKT Pro</h1>
      <div class='lead'>Thêm ứng dụng vào màn hình chính để mở nhanh như app riêng, không cần tìm lại trình duyệt.</div>
      <div class='benefits'>
        <div>✔ Hoạt động như ứng dụng điện thoại</div>
        <div>✔ Không cần mở trình duyệt mỗi lần sử dụng</div>
        <div>✔ Truy cập chỉ 1 chạm từ màn hình chính</div>
        <div>✔ Phù hợp khách dùng Premium hằng ngày</div>
      </div>
      <button id='installBtn' class='btn' onclick='installApp()'>📲 Cài đặt GPT MKT Pro</button>
      <button class='btn secondary' onclick='showManualGuide()'>📘 Xem hướng dẫn thêm vào màn hình chính</button>
      <div id='installStatus' class='status'>Đang kiểm tra thiết bị...</div>
      <div id='iosSteps' class='steps'>
        <b>📱 Cài trên iPhone/iPad</b><br>
        1. Mở trang này bằng Safari.<br>
        2. Bấm nút Chia sẻ ở dưới màn hình.<br>
        3. Chọn “Thêm vào Màn hình chính”.<br>
        4. Bấm “Thêm”.<br>
        📱 Biểu tượng GPT MKT Pro sẽ xuất hiện ngoài màn hình chính.
      </div>
      <div id='androidSteps' class='steps'>
        <b>🤖 Cài trên Android</b><br>
        1. Mở bằng Chrome hoặc Edge.<br>
        2. Bấm “Cài đặt GPT MKT Pro”.<br>
        3. Xác nhận “Cài đặt” / “Thêm vào màn hình chính”.<br>
        📱 Sau đó icon GPT MKT Pro sẽ xuất hiện ngoài màn hình chính.
      </div>
      <div class='tiny'>{app_url}</div>
    </section>
  </div>
<script>
let deferredPrompt=null;
const statusEl=document.getElementById('installStatus');
const installBtn=document.getElementById('installBtn');
function isStandalone(){{ return (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) || window.navigator.standalone; }}
function isIOS(){{ return /iphone|ipad|ipod/i.test(navigator.userAgent); }}
function isAndroid(){{ return /android/i.test(navigator.userAgent); }}
function markInstalled(){{
  statusEl.innerText='✅ GPT MKT Pro đã được cài đặt. Biểu tượng ứng dụng đã có trên màn hình chính.';
  installBtn.innerText='✅ Đã cài đặt GPT MKT Pro';
  installBtn.classList.add('installed');
  installBtn.disabled=true;
}}
window.addEventListener('beforeinstallprompt', function(e){{
  e.preventDefault();
  deferredPrompt=e;
  statusEl.innerText='Thiết bị này hỗ trợ cài đặt tự động. Bấm “Cài đặt GPT MKT Pro”.';
}});
async function installApp(){{
  if(isStandalone()){{ markInstalled(); return; }}
  if(deferredPrompt){{
    statusEl.innerText='Đang mở hộp cài đặt...';
    deferredPrompt.prompt();
    const choice = await deferredPrompt.userChoice;
    if(choice.outcome === 'accepted'){{
      statusEl.innerText='Đang cài đặt GPT MKT Pro... Sau khi cài xong, icon sẽ xuất hiện ngoài màn hình chính.';
    }} else {{
      statusEl.innerText='Anh/chị có thể bấm cài đặt lại bất cứ lúc nào.';
    }}
    deferredPrompt=null;
    return;
  }}
  showManualGuide();
}}
function showManualGuide(){{
  document.getElementById('iosSteps').classList.toggle('show', isIOS());
  document.getElementById('androidSteps').classList.toggle('show', !isIOS());
  if(isIOS()){{
    statusEl.innerText='iPhone/iPad cần cài bằng Safari: Chia sẻ → Thêm vào Màn hình chính.';
  }} else if(isAndroid()){{
    statusEl.innerText='Nếu chưa hiện popup, hãy mở bằng Chrome/Edge rồi chọn menu ⋮ → Cài đặt ứng dụng / Thêm vào màn hình chính.';
  }} else {{
    statusEl.innerText='Trên máy tính, dùng Chrome/Edge và bấm biểu tượng cài đặt trên thanh địa chỉ.';
  }}
}}
window.addEventListener('appinstalled', function(){{ markInstalled(); }});
if('serviceWorker' in navigator){{ window.addEventListener('load', function(){{ navigator.serviceWorker.register('/service-worker.js').catch(function(){{}}); }}); }}
window.addEventListener('load', function(){{
  if(isStandalone()){{ markInstalled(); return; }}
  setTimeout(function(){{
    if(deferredPrompt){{ statusEl.innerText='Thiết bị này sẵn sàng cài đặt. Bấm nút xanh để thêm vào màn hình chính.'; return; }}
    if(isIOS()){{ showManualGuide(); }}
    else {{ statusEl.innerText='Bấm nút xanh để cài đặt. Nếu chưa hiện popup, hệ thống sẽ hiện hướng dẫn phù hợp.'; }}
  }}, 900);
}});
</script>
</body>
</html>
"""

@app.get("/manifest.json")
def pwa_manifest():
    return jsonify({
        "name": "GPT MKT Pro",
        "short_name": "GPT MKT Pro",
        "description": "GPT MKT Pro - AI Marketing, Facebook, CRM và Automation",
        "start_url": "/?source=pwa",
        "scope": "/",
        "id": "/",
        "display": "standalone",
        "display_override": ["standalone", "minimal-ui", "browser"],
        "background_color": "#FFFFFF",
        "theme_color": "#2563EB",
        "orientation": "portrait",
        "categories": ["business", "productivity", "marketing"],
        "icons": [
            {"src": "/pwa-icon-192.png?v=gpt-mkt-pro-1", "sizes": "192x192", "type": "image/png", "purpose": "any maskable"},
            {"src": "/pwa-icon-512.png?v=gpt-mkt-pro-1", "sizes": "512x512", "type": "image/png", "purpose": "any maskable"}
        ]
    })

@app.get("/service-worker.js")
def pwa_service_worker():
    js = """
const CACHE_NAME = 'gpt-mkt-pro-icon-v1';
const ASSETS = ['/', '/install', '/manifest.json', '/pwa-icon-192.png?v=gpt-mkt-pro-1', '/pwa-icon-512.png?v=gpt-mkt-pro-1'];
self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS)).catch(() => null));
  self.skipWaiting();
});
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key)))).then(() => self.clients.claim())
  );
});
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  event.respondWith(fetch(event.request).catch(() => caches.match(event.request)));
});
"""
    return app.response_class(js, mimetype="application/javascript")

# Icon PWA GPT MKT Pro nhúng trực tiếp trong file để khi deploy lên Render không bị mất logo.
# Thiết kế: nền trắng, GPT MKT PRO, bot tím, bo góc chuẩn app điện thoại.
PWA_ICON_192_BASE64 = """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lqR7797TddEHsD01aCSzc5r52SoqaFLwPJIkJknsFJFJFEIJVKIII0Wj5tNspOzI5Wqdudk6R4/N8vIrx1BmlMCvcsGOnRQK48wvLuA4NsPJEHHcYHg07aDGcYLtpCCAbUs8R2HZHQnEOEp1VTdNOzz/eJXmEiRSEweiu97JaYdqncqpqCQthJM45VitnI3xG4axzZLqco2VlZBG06e63KQ8VKBYtnM+BYOwCEOvS3vHHCW/Qw9OXfJNyzyG29UAzqn3iVzDqosW3140omtmRMpUtyfdXESb/dri9wgheO655/ijP/oj/uzP/owoirBkikK+luB/7fVLD0LUwlk9z+Md73gHP/kTP8ktt9zC+MQEVs50QbSp0aarg0pv4bOBF01+5zKS+UXFcrWOScKMCWllzEyrA4EqQ6LSfDdJNHGGjNRqIUtVnxMnD2GMwJYwNjZNwStQKrt4nkOpUMYojV0QuI6V8X/SeqB1c1piuq2cXiA4erjGqQOat/2z85BljUlWuat1ArLV19BZLZANLPvLmv2PV7nwWgfhhUSBJooTkkThFVx27NpEsWQPdnH534XlbHi+uy95uusU6rVxasWWMYbFhQUe/cEP+PKXv8xf/81f02w0/5d2/X/UBdB6EtnzYi655BJuv/123nrHW7n62muYmpyiWCzw/39snP79/9WPMIw4e/Ys+/fv5/777+Vb3/oWzz+/r4Pdn2Oh+799AeQLo37Nh927d3Pddddx/fXXs3fPXrZu28rU1DTlcrkjLS5FZ/ppgFle1/HdZwftPUna/QdB95COyCExebs706Fn9A7ztGs+0UcDf9Wwduf5dVbYWlZHc7PfTtiNrphV2LjOzPzy50fX+JEYgKBs4J71GzZfnbv37v75Xs5qjlanFsnDkVkdknX0tVJoYwjDgNnZOU6dOsmLL77I0888w9NPPcXBgwe7zFs20tg614//F0WSEirX2v9VAAAAAElFTkSuQmCC"""
PWA_ICON_512_BASE64 = """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"""

def _pwa_icon_bytes(size=192):
    data = PWA_ICON_512_BASE64 if int(size) == 512 else PWA_ICON_192_BASE64
    return base64.b64decode(data)

@app.get("/pwa-icon-192.png")
def pwa_icon_192():
    resp = app.response_class(_pwa_icon_bytes(192), mimetype="image/png")
    resp.headers["Cache-Control"] = "public, max-age=86400"
    return resp

@app.get("/pwa-icon-512.png")
def pwa_icon_512():
    resp = app.response_class(_pwa_icon_bytes(512), mimetype="image/png")
    resp.headers["Cache-Control"] = "public, max-age=86400"
    return resp


# Khởi tạo database khi import app để chạy ổn cả python app.py và gunicorn app:app.
try:
    init_db()
except Exception as e:
    print("Init DB error:", e)

if __name__ == "__main__":
    threading.Thread(target=scheduler_loop, daemon=True).start()
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=False)
