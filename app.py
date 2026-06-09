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
<link rel="apple-touch-icon" href="/pwa-icon-192.png">
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
  const trialAllowed = ["dashboard", "facebook_center", "fanpage_manager", "group_marketing", "comment_manager", "post", "token", "premium"];
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
    "campaign": "Campaign Manager"
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
    title:"🎁 DÙNG THỬ MIỄN PHÍ – 3 NGÀY",
    price:"0Đ",
    amount:"0",
    package:"TRIAL",
    desc:"Dùng thử 3 ngày sau khi đăng ký. Chỉ mở Quản lý Fanpage, Quản lý Group và AI Comment.",
    benefits:["Quản lý Fanpage", "Quản lý Group", "AI Comment", "Xem giao diện hệ thống", "Trải nghiệm quy trình bán hàng"],
    locked:["AI Messenger", "CRM Kanban", "AI Marketing Director", "AI Video", "AI Image", "AI Kinh Doanh", "AI Giọng Nói", "AI Livestream"]
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
  <a class="v2-nav-link" href="/install" target="_blank"><span class="v2-nav-ico">🚀</span><span class="v2-nav-text">Thêm vào màn hình</span><span class="v2-nav-tag">APP</span></a>

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
  <b>🚀 App Mini đã sẵn sàng</b><br>
  Bấm nút bên dưới để thêm Mkt Automation Pro vào màn hình chính và mở như phần mềm riêng.
  <button onclick="showSmartInstall()">🚀 THÊM VÀO MÀN HÌNH CHÍNH</button>
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
  <div class="section-open-note">Bạn đang mở: Facebook Center V3.</div>
  <h2>📢 Facebook Center</h2>
  <p class="small">Gộp toàn bộ công cụ Facebook vào một trung tâm: đăng bài, Scheduler, Fanpage Manager, Group Marketing, Comment Manager và Messenger AI.</p>
  <div class="fb-submenu-pro">
    <button onclick="openModule('post')">📢 Đăng bài</button>
    <button onclick="openModule('scheduler')">📅 Scheduler</button>
    <button onclick="openModule('fanpage_manager')">📄 Fanpage Manager</button>
    <button onclick="openModule('group_marketing')">👥 Group Marketing</button>
    <button onclick="openModule('comment_manager')">💬 Comment Manager AI</button>
    <button onclick="openModule('messenger_ai')">🤖 Messenger AI</button>
  </div>
  <div class="v3-feature-grid">
    <div class="v3-feature-card"><h3>Đăng bài</h3><ul><li>Đăng ngay</li><li>Đăng hàng loạt</li><li>Đăng nhiều Page</li><li>Đăng ảnh/video</li></ul><button onclick="openModule('post')">Mở công cụ đăng bài</button></div>
    <div class="v3-feature-card"><h3>Scheduler</h3><ul><li>Lên lịch tự động</li><li>Đăng chiến dịch</li><li>Chia khung giờ</li><li>Tự lưu lịch</li></ul><button onclick="openModule('scheduler')">Mở lịch đăng</button></div>
    <div class="v3-feature-card"><h3>Quản lý Fanpage</h3><ul><li>Kết nối Fanpage</li><li>Kiểm tra Token</li><li>Kiểm tra quyền</li><li>Trạng thái hoạt động</li><li>Làm mới Token</li></ul><button onclick="openModule('fanpage_manager')">Mở Fanpage Manager</button></div>
    <div class="v3-feature-card"><h3>Tiếp thị nhóm</h3><ul><li>Quản lý Group</li><li>Danh sách Group</li><li>Lịch đăng Group</li><li>AI viết bài Group</li></ul><button onclick="openModule('group_marketing')">Mở Group Marketing</button></div>
    <div class="v3-feature-card"><h3>Trình quản lý bình luận</h3><ul><li>AI trả lời comment</li><li>Ẩn SĐT</li><li>Gắn nhãn khách</li><li>Chuyển CRM</li></ul><button onclick="openModule('comment_manager')">Mở Comment AI</button></div>
    <div class="v3-feature-card"><h3>Trí tuệ nhân tạo Messenger</h3><ul><li>Kịch bản Inbox</li><li>Kịch bản Chốt Sale</li><li>Xử lý từ chối</li><li>Chăm sóc khách cũ</li></ul><button onclick="openModule('messenger_ai')">Mở Messenger AI</button></div>
  </div>
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
      <div class="premium-plan free-plan v4-plan">
        <div class="plan-ribbon">DÙNG THỬ</div>
        <div class="plan-name">Dùng thử miễn phí</div>
        <div class="plan-price">3 ngày</div>
        <div class="plan-desc">Dùng thử 3 ngày sau khi đăng ký, chỉ mở 3 công cụ chính để khách trải nghiệm trước khi nâng cấp.</div>
        <div class="benefit-title">Quyền lợi dùng thử</div>
        <ul class="benefit-list">
          <li class="open">Quản lý Fanpage</li>
          <li class="open">Quản lý Group</li>
          <li class="open">AI Comment</li>
          <li class="locked">AI Messenger cần Premium</li>
          <li class="locked">CRM Kanban cần Premium</li>
          <li class="locked">AI Marketing Director cần Premium</li>
        </ul>
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

  <div class="{{ 'free-status-card free-expired' if free_status.is_expired else 'free-status-card' }}">
    <h3>🎁 Gói dùng thử 3 ngày</h3>
    {% if free_status.is_expired %}
      <b>Trạng thái:</b> Đã hết dùng thử<br>
      Vui lòng nâng cấp Premium để tiếp tục sử dụng các công cụ.
    {% else %}
      <b>Còn lại:</b> {{ free_status.days }} ngày {{ free_status.hours }} giờ
      <div class="free-progress"><span style="width:{{ free_status.percent }}%"></span></div>
      <div class="trial-box">
        <b>Được sử dụng:</b><br>
        ✓ Quản lý Fanpage<br>
        ✓ Quản lý Group<br>
        ✓ AI Comment
      </div>
      <div class="trial-box locked-list">
        <b>Chưa mở khóa:</b><br>
        🔒 AI Messenger<br>
        🔒 CRM Kanban<br>
        🔒 AI Marketing Director<br>
        🔒 AI Video • AI Image<br>
        🔒 AI Giọng Nói • AI Livestream
      </div>
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
    alert('Mkt Automation Pro đã được cài đặt trên thiết bị này.');
    return;
  }
  var title='Cài đặt Mkt Automation Pro';
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
  if(st){st.innerText='Đã cài đặt Mkt Automation Pro thành công.';}
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
          <div class='menu-title'>Cài đặt App</div><a href='#install_app'>🚀 Thêm vào màn hình</a><a href='/install' target='_blank'>📲 Mở trang cài đặt</a>
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


          <h2 id='install_app'>🚀 Thêm ứng dụng vào màn hình chính</h2>
          <div class='grid2'>
            <div class='panel'>
              <h3>📱 Cài Mkt Automation Pro nhanh cho khách</h3>
              <div style='line-height:1.9;color:#cbd5e1'>
                <b style='color:#86efac'>Không dùng QR nữa.</b><br>
                Khách chỉ cần bấm nút bên dưới. Hệ thống tự nhận diện Android, iPhone hoặc máy tính và hiện đúng hướng dẫn cài đặt.
                <br><br>
                <a class='btn primary' href='/install' target='_blank' style='display:inline-block;margin-top:10px'>🚀 Mở trang thêm vào màn hình chính</a>
              </div>
            </div>
            <div class='panel'>
              <h3>⚡ Cách hoạt động</h3>
              <div class='mini-grid'>
                <div class='mini-card'><span>Android</span><b>Popup cài</b><span>Chrome/Edge hỗ trợ sẽ hiện nút Cài đặt.</span></div>
                <div class='mini-card'><span>iPhone</span><b>Safari</b><span>Hiện hướng dẫn Chia sẻ → Thêm vào Màn hình chính.</span></div>
                <div class='mini-card'><span>Máy tính</span><b>App riêng</b><span>Cài như phần mềm riêng trên Chrome/Edge.</span></div>
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



@app.get('/install')
def install_page():
    """Trang cài đặt PWA thông minh không dùng QR: Android gọi prompt, iPhone hướng dẫn thêm màn hình chính."""
    app_url = request.host_url.rstrip('/')
    return f"""
<!doctype html>
<html lang='vi'>
<head>
  <meta charset='utf-8'>
  <meta name='viewport' content='width=device-width, initial-scale=1, viewport-fit=cover'>
  <meta name='theme-color' content='#0F172A'>
  <meta name='apple-mobile-web-app-capable' content='yes'>
  <meta name='apple-mobile-web-app-status-bar-style' content='black-translucent'>
  <meta name='apple-mobile-web-app-title' content='Mkt Pro'>
  <title>Thêm Mkt Automation Pro vào màn hình</title>
  <link rel='manifest' href='/manifest.json'>
  <link rel='apple-touch-icon' href='/pwa-icon-192.png'>
  <link rel='preconnect' href='https://fonts.googleapis.com'>
  <link rel='preconnect' href='https://fonts.gstatic.com' crossorigin>
  <link href='https://fonts.googleapis.com/css2?family=Inter:wght@500;600;700;800;900&family=Manrope:wght@500;600;700;800;900&display=swap' rel='stylesheet'>
  <style>
    *{{box-sizing:border-box}}
    body{{margin:0;min-height:100vh;font-family:'Manrope','Inter',system-ui;background:radial-gradient(circle at top,#1e3a8a 0,#0f172a 42%,#020617 100%);color:#e5e7eb;display:flex;align-items:center;justify-content:center;padding:18px}}
    .wrap{{width:min(720px,100%);background:rgba(15,23,42,.78);border:1px solid rgba(148,163,184,.24);box-shadow:0 32px 90px rgba(0,0,0,.42);backdrop-filter:blur(22px);border-radius:32px;padding:28px;position:relative;overflow:hidden}}
    .wrap:before{{content:'';position:absolute;inset:-120px auto auto -80px;width:260px;height:260px;background:rgba(59,130,246,.28);filter:blur(45px);border-radius:999px}}
    .badge{{display:inline-flex;gap:8px;align-items:center;background:rgba(34,197,94,.13);border:1px solid rgba(34,197,94,.36);color:#86efac;border-radius:999px;padding:8px 12px;font-weight:900;font-size:13px;position:relative}}
    h1{{font-size:clamp(34px,8vw,58px);line-height:1.03;margin:18px 0 12px;font-weight:900;letter-spacing:-2px;background:linear-gradient(90deg,#38bdf8,#818cf8,#c084fc,#facc15);-webkit-background-clip:text;-webkit-text-fill-color:transparent;position:relative}}
    .lead{{color:#cbd5e1;font-size:17px;line-height:1.8;margin-bottom:18px;position:relative}}
    .benefits{{display:grid;gap:10px;margin:20px 0;position:relative}}
    .benefits div{{background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.09);border-radius:16px;padding:12px 14px;font-weight:800;color:#f8fafc}}
    .btn{{width:100%;border:0;border-radius:20px;padding:17px 18px;font-size:17px;font-weight:900;cursor:pointer;background:linear-gradient(135deg,#22c55e,#16a34a);color:white;box-shadow:0 18px 42px rgba(34,197,94,.28);position:relative}}
    .btn.secondary{{background:linear-gradient(135deg,#2563eb,#7c3aed);margin-top:10px}}
    .status{{margin-top:14px;line-height:1.65;color:#bfdbfe;font-weight:800;min-height:22px}}
    .steps{{display:none;margin-top:18px;background:#020617;border:1px solid #334155;border-radius:22px;padding:18px;line-height:1.9;color:#dbeafe}}
    .steps.show{{display:block}}
    .steps b{{color:#facc15}}
    .tip{{font-size:13px;color:#94a3b8;line-height:1.7;margin-top:16px;text-align:center}}
    .device-pill{{display:inline-flex;background:rgba(59,130,246,.15);border:1px solid rgba(96,165,250,.28);color:#bfdbfe;border-radius:999px;padding:8px 12px;font-size:13px;font-weight:900;margin:4px 4px 14px 0}}
  </style>
</head>
<body>
  <main class='wrap'>
    <div class='badge'>🟢 App Mini sẵn sàng</div>
    <h1>Thêm vào màn hình chính</h1>
    <div class='lead'>Mkt Automation Pro sẽ mở như một ứng dụng riêng trên điện thoại, truy cập nhanh chỉ 1 chạm.</div>
    <div id='deviceHint' class='device-pill'>Đang nhận diện thiết bị...</div>
    <div class='benefits'>
      <div>✔ Mở như app trên điện thoại</div>
      <div>✔ Không cần tìm lại trình duyệt</div>
      <div>✔ Truy cập nhanh từ màn hình chính</div>
      <div>✔ Phù hợp khách dùng Premium mỗi ngày</div>
    </div>
    <button id='installBtn' class='btn' onclick='installApp()'>🚀 THÊM VÀO MÀN HÌNH CHÍNH</button>
    <button class='btn secondary' onclick='showManualGuide()'>📘 Xem hướng dẫn cài đặt</button>
    <div id='installStatus' class='status'></div>
    <div id='iosSteps' class='steps'>
      <b>📱 iPhone/iPad</b><br>
      1. Mở trang này bằng Safari.<br>
      2. Bấm nút Chia sẻ ở dưới màn hình.<br>
      3. Chọn “Thêm vào Màn hình chính”.<br>
      4. Bấm “Thêm”.
    </div>
    <div id='androidSteps' class='steps'>
      <b>🤖 Android Chrome/Edge</b><br>
      1. Bấm nút “THÊM VÀO MÀN HÌNH CHÍNH”.<br>
      2. Nếu popup hiện ra, chọn “Cài đặt”.<br>
      3. Nếu chưa hiện, bấm menu ⋮ của Chrome → “Thêm vào màn hình chính”.
    </div>
    <div id='desktopSteps' class='steps'>
      <b>💻 Máy tính Chrome/Edge</b><br>
      1. Bấm nút cài đặt trên thanh địa chỉ nếu có.<br>
      2. Hoặc bấm menu ⋮ → Cài đặt Mkt Automation Pro.<br>
      3. Sau khi cài, app sẽ mở như phần mềm riêng.
    </div>
    <div class='tip'>Nếu trình duyệt chưa hiện popup cài, nguyên nhân thường là chưa mở bằng Chrome/Edge Android hoặc iPhone chưa mở bằng Safari.</div>
  </main>
<script>
let deferredPrompt=null;
const statusEl=document.getElementById('installStatus');
const deviceHint=document.getElementById('deviceHint');
const installBtn=document.getElementById('installBtn');
function isStandalone(){{ return (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) || window.navigator.standalone; }}
function isIOS(){{ return /iphone|ipad|ipod/i.test(navigator.userAgent); }}
function isAndroid(){{ return /android/i.test(navigator.userAgent); }}
function isDesktop(){{ return !isIOS() && !isAndroid(); }}
function setDeviceText(){{
  if(isStandalone()){{ deviceHint.innerText='✅ Ứng dụng đã được cài'; installBtn.innerText='✅ ĐÃ CÀI ĐẶT'; statusEl.innerText='Anh/chị có thể mở app từ màn hình chính.'; return; }}
  if(isIOS()) deviceHint.innerText='📱 Đã phát hiện iPhone/iPad';
  else if(isAndroid()) deviceHint.innerText='🤖 Đã phát hiện Android';
  else deviceHint.innerText='💻 Đã phát hiện máy tính';
}}
window.addEventListener('beforeinstallprompt', function(e){{
  e.preventDefault();
  deferredPrompt=e;
  statusEl.innerText='Thiết bị này hỗ trợ cài đặt tự động. Bấm nút xanh để cài.';
  installBtn.innerText='🚀 CÀI ĐẶT NGAY';
}});
function hideSteps(){{ ['iosSteps','androidSteps','desktopSteps'].forEach(id=>document.getElementById(id).classList.remove('show')); }}
function installApp(){{
  if(isStandalone()){{ statusEl.innerText='Ứng dụng đã được cài trên thiết bị này.'; return; }}
  if(deferredPrompt){{
    deferredPrompt.prompt();
    deferredPrompt.userChoice.then(function(choice){{
      statusEl.innerText = choice.outcome === 'accepted' ? 'Đang cài đặt ứng dụng...' : 'Anh/chị có thể bấm cài lại bất cứ lúc nào.';
      deferredPrompt=null;
    }});
    return;
  }}
  showManualGuide();
}}
function showManualGuide(){{
  hideSteps();
  if(isIOS()){{ document.getElementById('iosSteps').classList.add('show'); statusEl.innerText='iPhone cần thêm vào màn hình chính bằng Safari.'; }}
  else if(isAndroid()){{ document.getElementById('androidSteps').classList.add('show'); statusEl.innerText='Nếu chưa thấy popup, dùng menu ⋮ của Chrome/Edge để thêm vào màn hình chính.'; }}
  else {{ document.getElementById('desktopSteps').classList.add('show'); statusEl.innerText='Máy tính có thể cài app bằng Chrome/Edge nếu trình duyệt hỗ trợ.'; }}
}}
window.addEventListener('appinstalled', function(){{ statusEl.innerText='Đã cài đặt Mkt Automation Pro thành công.'; installBtn.innerText='✅ ĐÃ CÀI ĐẶT'; }});
if('serviceWorker' in navigator){{ window.addEventListener('load', function(){{ navigator.serviceWorker.register('/service-worker.js').catch(function(err){{ statusEl.innerText='Service Worker chưa sẵn sàng. Vui lòng tải lại trang.'; }}); }}); }}
setDeviceText();
setTimeout(function(){{ if(!deferredPrompt && !isStandalone()) showManualGuide(); }}, 1400);
</script>
</body>
</html>
"""

@app.get("/manifest.json")
def pwa_manifest():
    return jsonify({
        "name": "Mkt Automation Pro",
        "short_name": "Mkt Pro",
        "description": "Mkt Automation Pro - AI Marketing, Facebook, CRM và Automation",
        "start_url": "/",
        "scope": "/",
        "display": "standalone",
        "display_override": ["standalone", "minimal-ui"],
        "background_color": "#0F172A",
        "theme_color": "#0F172A",
        "orientation": "portrait",
        "categories": ["business", "productivity", "utilities"],
        "prefer_related_applications": False,
        "icons": [
            {"src": "/pwa-icon-192.png", "sizes": "192x192", "type": "image/png"},
            {"src": "/pwa-icon-512.png", "sizes": "512x512", "type": "image/png"}
        ]
    })

@app.get("/service-worker.js")
def pwa_service_worker():
    js = """
const CACHE_NAME = 'mkt-automation-pro-v2';
const ASSETS = ['/', '/install', '/manifest.json', '/pwa-icon-192.png', '/pwa-icon-512.png'];
self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => Promise.allSettled(ASSETS.map(url => cache.add(url)))));
  self.skipWaiting();
});
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))).then(() => self.clients.claim())
  );
});
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  event.respondWith(fetch(event.request).then(response => {
    const copy = response.clone();
    caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy)).catch(() => {});
    return response;
  }).catch(() => caches.match(event.request)));
});
"""
    return app.response_class(js, mimetype="application/javascript")

def _solid_png(size=192):
    import struct, zlib
    size = int(size)
    # PNG RGBA nền xanh tím đơn giản, dùng làm icon PWA nếu chưa có file icon.
    raw = b''.join([b'\x00' + bytes([37, 99, 235, 255]) * size for _ in range(size)])
    def chunk(tag, data):
        return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', zlib.crc32(tag + data) & 0xffffffff)
    png = b'\x89PNG\r\n\x1a\n'
    png += chunk(b'IHDR', struct.pack('>IIBBBBB', size, size, 8, 6, 0, 0, 0))
    png += chunk(b'IDAT', zlib.compress(raw, 9))
    png += chunk(b'IEND', b'')
    return png

@app.get("/pwa-icon-192.png")
def pwa_icon_192():
    if os.path.exists("pwa-icon-192.png"):
        return send_file("pwa-icon-192.png", mimetype="image/png")
    return app.response_class(_solid_png(192), mimetype="image/png")

@app.get("/pwa-icon-512.png")
def pwa_icon_512():
    if os.path.exists("pwa-icon-512.png"):
        return send_file("pwa-icon-512.png", mimetype="image/png")
    return app.response_class(_solid_png(512), mimetype="image/png")


# Khởi tạo database khi import app để chạy ổn cả python app.py và gunicorn app:app.
try:
    init_db()
except Exception as e:
    print("Init DB error:", e)

if __name__ == "__main__":
    threading.Thread(target=scheduler_loop, daemon=True).start()
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=False)
