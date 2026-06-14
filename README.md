# Facebook Automation Platform 2026

Kiến trúc: React PWA + Node.js/Express + PostgreSQL + Redis/BullMQ + Playwright.

Lưu ý sử dụng an toàn: hệ thống chỉ phục vụ đăng nội dung hợp pháp do chính người dùng sở hữu/quản trị. Không bypass captcha, không spam, không né kiểm duyệt, không tự động thu thập dữ liệu cá nhân. Ưu tiên Meta Graph API khi có quyền chính thức; Playwright dùng cho workflow có người dùng đăng nhập thủ công và cần kiểm tra lại trước khi chạy thật.

## Chạy local
```bash
cp .env.example .env
npm install
npm run dev
```

Hoặc chạy từng phần:
```bash
cd server && npm install && npm run migrate && npm run dev
cd worker && npm install && npm run dev
cd web && npm install && npm run dev
```
