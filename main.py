"""FMP IT Ticket System — Vietnamese/English, FastAPI + SQLite. v20: add IT department, all-ticket default filters, polished account metadata."""
import hashlib
import html
import io
import json
import math
import os
import re
import secrets
import sqlite3
import time
from contextlib import asynccontextmanager, closing
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request as URLRequest, urlopen

from dotenv import load_dotenv
from fastapi import BackgroundTasks, Depends, FastAPI, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from starlette.middleware.sessions import SessionMiddleware

load_dotenv()
DATA_DIR = Path(os.getenv("DATA_DIR", "/data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = os.getenv("DB_PATH", str(DATA_DIR / "tickets.db"))
ADMIN_USER = os.getenv("ADMIN_USER", "admin").strip().lower()
ADMIN_PASS = os.getenv("ADMIN_PASS", "")
ADMIN_NAME = os.getenv("ADMIN_NAME", "IT Administrator")
SECRET_KEY = os.getenv("SECRET_KEY", "")
TZ = ZoneInfo("Asia/Ho_Chi_Minh")
STATUSES = ["Mới", "Đang xử lý", "Đã đóng"]
SEVERITIES = ["Thấp", "Trung bình", "Cao"]
DEVICES = ["Máy tính", "Máy in", "Máy scan", "Điện thoại", "Máy photocopy", "Mạng/Wi-Fi", "Khác"]
ISSUE_TYPES = ["Phần cứng", "Phần mềm", "Kết nối mạng", "Tài khoản & phân quyền", "Máy in / Scan", "Email & ứng dụng", "Điện thoại", "Thiết bị ngoại vi", "Bảo mật", "Khác"]
DEPARTMENTS = ["Lễ tân", "Bác sĩ", "Điều dưỡng", "Xét nghiệm", "CĐHA", "Dược", "Kế toán", "Hành chính", "IT", "Khác"]
# Database values stay canonical Vietnamese; the UI translates them according to the selected language.
DEPARTMENT_EN = {"Lễ tân":"Reception", "Bác sĩ":"Doctors", "Điều dưỡng":"Nurse", "Xét nghiệm":"Laboratory", "CĐHA":"Medical Imaging", "Dược":"Pharmacy", "Kế toán":"Accounting", "Hành chính":"Administration", "IT":"IT", "Khác":"Other"}
DEPARTMENT_ALIASES = {
    # Reception
    "Lễ tân":"Lễ tân", "Reception":"Lễ tân", "Lễ tân / Reception":"Lễ tân", "Reception / Lễ tân":"Lễ tân",
    # Doctors
    "Bác sĩ":"Bác sĩ", "Doctor":"Bác sĩ", "Doctors":"Bác sĩ", "Bác sĩ / Doctors":"Bác sĩ", "Doctors / Bác sĩ":"Bác sĩ",
    # Nursing
    "Điều dưỡng":"Điều dưỡng", "Y tá":"Điều dưỡng", "Nurse":"Điều dưỡng", "Nurses":"Điều dưỡng", "Nursing":"Điều dưỡng", "Y tá / Nurses":"Điều dưỡng", "Điều dưỡng / Nurse":"Điều dưỡng",
    # Laboratory
    "Xét nghiệm":"Xét nghiệm", "Laboratory":"Xét nghiệm", "Lab":"Xét nghiệm", "Xét nghiệm / Laboratory":"Xét nghiệm", "Lab / Xét nghiệm":"Xét nghiệm",
    # Medical Imaging (CĐHA): legacy X-ray and Ultrasound values converge here.
    "CĐHA":"CĐHA", "Chẩn đoán hình ảnh":"CĐHA", "Chan doan hinh anh":"CĐHA",
    "X-quang / Siêu âm":"CĐHA", "Siêu âm / X-quang":"CĐHA",
    "X-ray / Ultrasound":"CĐHA", "Ultrasound / X-ray":"CĐHA",
    "Siêu âm":"CĐHA", "Ultrasound":"CĐHA", "Siêu âm / Ultrasound":"CĐHA", "Ultrasound / Siêu âm":"CĐHA",
    "X-quang":"CĐHA", "X-ray":"CĐHA", "X-Ray":"CĐHA", "X-ray / X-quang":"CĐHA", "X-quang / X-ray":"CĐHA",
    "Medical Imaging":"CĐHA", "Radiology":"CĐHA", "Radiology / Ultrasound":"CĐHA",
    # Pharmacy
    "Dược":"Dược", "Pharmacy":"Dược", "Dược / Pharmacy":"Dược", "Pharmacy / Dược":"Dược",
    # Accounting
    "Kế toán":"Kế toán", "Accounting":"Kế toán", "Kế toán / Accounting":"Kế toán", "Accounting / Kế toán":"Kế toán",
    # Administration
    "Hành chính":"Hành chính", "Administration":"Hành chính", "Hành chính / Administration":"Hành chính", "Administration / Hành chính":"Hành chính",
    # IT department and legacy IT support names share one canonical department.
    "IT":"IT", "IT Department":"IT", "IT Support":"IT", "IT support":"IT", "Hỗ trợ CNTT":"IT", "Hỗ trợ IT":"IT", "Công nghệ thông tin":"IT",
    "IT Support / Hỗ trợ CNTT":"IT", "Hỗ trợ CNTT / IT Support":"IT", "Công nghệ thông tin / IT":"IT",
    "Information Technology":"IT", "Information Technology (IT)":"IT", "IT Department / IT Support":"IT",
    "IT Dept":"IT", "IT Dept.":"IT", "IT Helpdesk":"IT", "IT Help Desk":"IT", "IT Support Department":"IT",
    "IT (Công nghệ thông tin)":"IT", "IT - Công nghệ thông tin":"IT", "Bộ phận IT":"IT", "Phòng IT":"IT",
    # Removed/unsupported departments and unknown legacy values are normalized to Other.
    "Khác":"Khác", "Other":"Khác", "Khác / Other":"Khác", "":"Khác",
}

E = html.escape

LANG = {
"vi": {
"title":"FMP IT Service Desk","home":"Trang chủ","my":"Danh sách ticket","ticket_list":"Danh sách ticket","new":"Tạo ticket","admin":"Danh sách ticket","users":"Tài khoản","settings":"Cài đặt","logout":"Đăng xuất","login":"Đăng nhập","username":"Tài khoản","password":"Mật khẩu","submit":"Gửi ticket","device":"Thiết bị","device_code":"Mã thiết bị","issue_type":"Loại sự cố","dept":"Bộ phận","location":"Vị trí","issue":"Mô tả sự cố","severity":"Mức độ","reporter":"Người báo","phone":"Điện thoại","code":"Mã ticket","time":"Thời gian tạo","status":"Trạng thái","handler":"Người xử lý","note":"Ghi chú","closed":"Thời gian đóng","received":"Thời gian nhận xử lý","take":"Nhận xử lý","taken":"Đã nhận xử lý","close":"Đóng ticket","closed_done":"Đã đóng ticket","must_take_first":"Vui lòng nhận xử lý ticket trước khi đóng.","save":"Lưu","back":"Quay lại","all":"Tất cả","filter":"Lọc","export":"Xuất Excel","fullname":"Họ tên","role":"Vai trò","add":"Thêm tài khoản","initial":"Mật khẩu ban đầu","active":"Hoạt động","locked":"Đã khóa","reset":"Đặt lại mật khẩu","lock":"Khóa/Mở khóa","current":"Mật khẩu hiện tại","newpw":"Mật khẩu mới (ít nhất 6 ký tự)","changepw":"Đổi mật khẩu","hint6":"Mật khẩu phải có ít nhất 6 ký tự.","no_tickets":"Chưa có ticket nào.","invalid_login":"Tên đăng nhập hoặc mật khẩu không đúng.","too_many":"Quá nhiều lần đăng nhập. Vui lòng thử lại sau 5 phút.","password_changed":"Đã đổi mật khẩu.","user_added":"Đã thêm tài khoản.","user_updated":"Đã cập nhật tài khoản.","password_reset":"Đã đặt lại mật khẩu.","username_exists":"Tên đăng nhập đã tồn tại.","invalid_data":"Dữ liệu không hợp lệ.","not_found":"Không tìm thấy ticket/người dùng.","admin_only":"Chỉ quản trị viên mới được phép thực hiện.","ticket_closed":"Ticket đã đóng.","invalid_action":"Thao tác không hợp lệ.","cannot_self_lock":"Không thể khóa tài khoản của chính bạn.","created":"Đã tạo ticket.","role_user":"Người dùng","role_admin":"Quản trị viên","date_format":"Định dạng ngày: Ngày/Tháng/Năm (DD/MM/YYYY). Giờ hiển thị theo giờ Việt Nam (UTC+7).","pending_received":"Chưa nhận xử lý","pending_closed":"Chưa đóng","clinic_support":"Công nghệ hỗ trợ chăm sóc sức khỏe","assigned_none":"Chưa phân công","today":"Hôm nay","low":"Thấp","medium":"Trung bình","high":"Cao","clinic_tagline":"Sức khỏe tốt hơn · Tương lai tươi sáng","delete_user":"Xóa tài khoản","confirm_delete_user":"Bạn có chắc muốn xóa tài khoản này? Ticket và lịch sử vẫn được giữ lại.","user_deleted":"Đã xóa tài khoản; các ticket và lịch sử vẫn được giữ lại.",},
"en": {
"title":"FMP IT Service Desk","home":"Home","my":"My Tickets","ticket_list":"Ticket List","new":"Create Ticket","admin":"Ticket List","users":"Accounts","settings":"Settings","logout":"Logout","login":"Login","username":"Username","password":"Password","submit":"Submit Ticket","device":"Device","device_code":"Device Code","issue_type":"Issue Type","dept":"Department","location":"Location","issue":"Issue Details","severity":"Severity","reporter":"Reporter","phone":"Phone","code":"Ticket ID","time":"Created At","status":"Status","handler":"Assigned To","note":"Note","closed":"Closed At","received":"Received At","take":"Take Ticket","taken":"Accepted for processing","close":"Close Ticket","closed_done":"Ticket closed","must_take_first":"Please accept the ticket before closing it.","save":"Save","back":"Back","all":"All","filter":"Filter","export":"Export Excel","fullname":"Full Name","role":"Role","add":"Add User","initial":"Initial Password","active":"Active","locked":"Locked","reset":"Reset Password","lock":"Lock/Unlock","current":"Current Password","newpw":"New Password (at least 6 characters)","changepw":"Change Password","hint6":"Password must be at least 6 characters.","no_tickets":"No tickets found.","invalid_login":"Invalid username or password.","too_many":"Too many login attempts. Please try again in 5 minutes.","password_changed":"Password changed successfully.","user_added":"User added successfully.","user_updated":"User updated successfully.","password_reset":"Password reset successfully.","username_exists":"Username already exists.","invalid_data":"Invalid data.","not_found":"Ticket or user not found.","admin_only":"Administrator access required.","ticket_closed":"This ticket is already closed.","invalid_action":"Invalid action.","cannot_self_lock":"You cannot lock your own account.","created":"Ticket created successfully.","role_user":"User","role_admin":"Administrator","date_format":"Date format: Day/Month/Year (DD/MM/YYYY). Times are displayed in Vietnam time (UTC+7).","pending_received":"Not yet accepted","pending_closed":"Not closed","clinic_support":"Technology supporting better healthcare","assigned_none":"Unassigned","today":"Today","low":"Low","medium":"Medium","high":"High","clinic_tagline":"Better Health · Brighter Future","delete_user":"Delete account","confirm_delete_user":"Delete this account? Tickets and history will be retained.","user_deleted":"Account deleted; tickets and history were retained.",}}
DEVICE_EN = ["Computer","Printer","Scanner","Phone","Photocopier","Network/Wi-Fi","Other"]
ISSUE_EN = ["Hardware", "Software", "Network connectivity", "Account & permissions", "Printer / Scanner", "Email & applications", "Phone", "Peripherals", "Security", "Other"]
STATUS_EN = {"Mới":"New","Đang xử lý":"In Progress","Đã đóng":"Closed"}
SEVERITY_EN = {"Thấp":"Low","Trung bình":"Medium","Cao":"High"}
DEVICE_MAP = dict(zip(DEVICES, DEVICE_EN))
ISSUE_MAP = dict(zip(ISSUE_TYPES, ISSUE_EN))
# Map older detailed issue values to the new high-level incident category for legacy tickets.
LEGACY_ISSUE_CATEGORY = {
    "Máy tính không vào mạng":"Kết nối mạng", "Máy tính không khởi động":"Phần cứng",
    "Máy in không in được":"Máy in / Scan", "Máy in báo lỗi/kẹt giấy":"Máy in / Scan",
    "Scan không được":"Máy in / Scan", "Điện thoại không nghe/gọi được":"Điện thoại",
    "Wi-Fi không kết nối":"Kết nối mạng", "Mạng chậm/mất mạng":"Kết nối mạng",
    "Phần mềm bị lỗi":"Phần mềm", "Thiết bị hỏng":"Phần cứng"
}
ISSUE_CATEGORY_ICONS = {
    "Phần cứng": '<rect x="3" y="4" width="18" height="13" rx="2"/><path d="M8 21h8M12 17v4"/>',
    "Phần mềm": '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8M8 12h5M8 16h8"/>',
    "Kết nối mạng": '<circle cx="12" cy="12" r="2"/><path d="M5 9a10 10 0 0 1 14 0M2 6a15 15 0 0 1 20 0M5 15a10 10 0 0 0 14 0"/>',
    "Tài khoản & phân quyền": '<circle cx="9" cy="8" r="3"/><path d="M3 20v-2a6 6 0 0 1 12 0v2M16 11l2 2 4-4"/>',
    "Máy in / Scan": '<path d="M6 8V3h12v5M6 17H4a2 2 0 0 1-2-2v-4a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v4a2 2 0 0 1-2 2h-2M6 14h12v7H6z"/>',
    "Email & ứng dụng": '<rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 7 9 6 9-6"/>',
    "Điện thoại": '<rect x="7" y="2" width="10" height="20" rx="2"/><path d="M10 5h4M11 19h2"/>',
    "Thiết bị ngoại vi": '<rect x="3" y="4" width="18" height="13" rx="2"/><path d="M8 21h8M12 17v4M18 8h.01"/>',
    "Bảo mật": '<path d="M12 22s8-4 8-11V5l-8-3-8 3v6c0 7 8 11 8 11Z"/><path d="m9 12 2 2 4-4"/>',
    "Khác": '<circle cx="12" cy="12" r="9"/><path d="M12 8v4M12 16h.01"/>'
}

def tr(req, key):
    return LANG.get(req.session.get("lang", "vi"), LANG["vi"]).get(key, key)

def label(req, value, mapping):
    return mapping.get(value, value) if req.session.get("lang") == "en" else value

def device_label(req, value): return label(req, value, DEVICE_MAP)

def device_icon_svg(value):
    """Small consistent line icon for each device type in ticket lists."""
    icons = {
        "Máy tính": '<rect x="3" y="4" width="18" height="13" rx="2"/><path d="M8 21h8M12 17v4"/>',
        "Máy in": '<path d="M6 9V3h12v6M6 17H4a2 2 0 0 1-2-2v-4a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v4a2 2 0 0 1-2 2h-2"/><path d="M6 14h12v7H6zM17 11h.01"/>',
        "Máy scan": '<path d="M4 4h16v5H4zM6 9v11h12V9M8 13h8M8 16h8"/>',
        "Điện thoại": '<rect x="7" y="2" width="10" height="20" rx="2"/><path d="M10 5h4M11 19h2"/>',
        "Máy photocopy": '<path d="M6 8V3h12v5M6 17H4a2 2 0 0 1-2-2v-4a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v4a2 2 0 0 1-2 2h-2"/><path d="M6 14h12v7H6z"/>',
        "Mạng/Wi-Fi": '<path d="M3 9a14 14 0 0 1 18 0M6 12a9 9 0 0 1 12 0M9 15a4 4 0 0 1 6 0M12 19h.01"/>',
        "Khác": '<rect x="3" y="4" width="18" height="14" rx="2"/><path d="M8 21h8M12 18v3"/>',
    }
    path = icons.get(value, icons["Khác"])
    return f'<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{path}</svg>'
def issue_label(req, value):
    canonical = LEGACY_ISSUE_CATEGORY.get(value, value)
    return label(req, canonical, ISSUE_MAP)

def issue_category_icon_svg(value):
    canonical = LEGACY_ISSUE_CATEGORY.get(value, value)
    path = ISSUE_CATEGORY_ICONS.get(canonical, ISSUE_CATEGORY_ICONS["Khác"])
    return f'<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{path}</svg>' 
def status_label(req, value): return label(req, value, STATUS_EN)

def status_icon_svg(value, size=12):
    """Consistent lightweight status icons: pending, in progress, completed."""
    paths = {
        "Mới": '<circle cx="12" cy="12" r="8.2"/><circle cx="12" cy="12" r="2.2" fill="currentColor" stroke="none"/>',
        "Đang xử lý": '<circle cx="12" cy="12" r="8.2"/><path d="M12 7.5v5l3.2 2"/>',
        "Đã đóng": '<circle cx="12" cy="12" r="8.2"/><path d="m8.2 12.2 2.5 2.5 5.2-5.3"/>',
    }
    path = paths.get(value, paths["Mới"])
    return f'<svg class="status-icon-svg" viewBox="0 0 24 24" width="{int(size)}" height="{int(size)}" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{path}</svg>'

def sev_label(req, value): return label(req, value, SEVERITY_EN)
def severity_icon_svg(value):
    """One consistent minimalist signal-bars icon for all severity levels."""
    level = {"Thấp": 1, "Low": 1, "Trung bình": 2, "Medium": 2, "Cao": 3, "High": 3}.get(value, 2)
    bars = []
    for i, (x, y, h) in enumerate(((3, 14, 6), (9, 9, 11), (15, 4, 16))):
        opacity = "1" if i < level else ".24"
        bars.append(f'<rect x="{x}" y="{y}" width="4" height="{h}" rx="1.4" fill="currentColor" opacity="{opacity}"/>')
    return '<svg class="severity-icon-svg" viewBox="0 0 22 22" width="14" height="14" aria-hidden="true" focusable="false">' + ''.join(bars) + '</svg>'
def normalize_department(value):
    """Normalize every legacy, translated, or unknown department to the fixed canonical set."""
    raw = str(value or "").strip()
    if not raw:
        return "Khác"
    if raw in DEPARTMENT_ALIASES:
        return DEPARTMENT_ALIASES[raw]
    folded = raw.casefold()
    for alias, canonical in DEPARTMENT_ALIASES.items():
        if alias.casefold() == folded:
            return canonical
    return "Khác"

def department_label(req, value):
    canonical = normalize_department(value)
    if req.session.get("lang") == "en":
        return DEPARTMENT_EN.get(canonical, "Other")
    return "CĐ Hình Ảnh" if canonical == "CĐHA" else canonical
def now_dt():
    """Current server time explicitly converted to Vietnam time (UTC+7)."""
    return datetime.now(TZ)

def now():
    return now_dt().strftime("%d/%m/%Y %H:%M:%S")

def current_year():
    return now_dt().year

def db():
    c = sqlite3.connect(DB_PATH, timeout=30)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    return c

def query(sql, args=(), one=False, commit=False):
    with closing(db()) as c:
        cur = c.execute(sql, args)
        if commit:
            c.commit()
            return cur.lastrowid
        rows = cur.fetchall()
        return (rows[0] if rows else None) if one else rows


def department_options():
    """Return the single approved department list used by every page and form."""
    return list(DEPARTMENTS)

# Notification integrations are stored in the existing SQLite database so NAS users
# do not need to add environment variables or install extra Python packages.
NOTIFICATION_EVENT_KEYS = {
    "new_ticket": "notify_new_ticket",
    "high_ticket": "notify_high_ticket",
    "ticket_taken": "notify_ticket_taken",
    "ticket_closed": "notify_ticket_closed",
    "ticket_updated": "notify_ticket_updated",
}


def get_integration_setting(key, default=""):
    try:
        row = query("SELECT value FROM integration_settings WHERE key=?", (key,), one=True)
        return str(row["value"]) if row else str(default)
    except sqlite3.OperationalError:
        # The table is created at application startup. This also keeps utility
        # functions safe during initial app setup or a database migration.
        return str(default)


def set_integration_setting(key, value):
    query("INSERT INTO integration_settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
          (key, str(value)), commit=True)


def integration_credentials_ready(service):
    if service == "teams":
        return bool(get_integration_setting("teams_webhook_url", "").strip())
    if service == "telegram":
        return bool(get_integration_setting("telegram_bot_token", "").strip() and
                    get_integration_setting("telegram_chat_id", "").strip())
    if service == "discord":
        url = get_integration_setting("discord_webhook_url", "").strip()
        return url.startswith("https://discord.com/api/webhooks/") or url.startswith("https://discordapp.com/api/webhooks/")
    return False


def integration_is_connected(service):
    return (get_integration_setting(f"{service}_enabled", "0") == "1" and
            integration_credentials_ready(service) and
            get_integration_setting(f"{service}_status", "not_connected") == "connected")


def set_integration_status(service, connected, error=""):
    set_integration_setting(f"{service}_status", "connected" if connected else "not_connected")
    set_integration_setting(f"{service}_last_checked", now())
    set_integration_setting(f"{service}_last_error", (str(error or "")[:350]))


def _post_json(url, payload, timeout=8):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = URLRequest(url, data=body, headers={"Content-Type": "application/json", "Accept": "application/json", "User-Agent": "FMP-IT-Service-Desk/1.0"}, method="POST")
    try:
        with urlopen(request, timeout=timeout) as response:
            status_code = getattr(response, "status", 200)
            response_body = response.read(8192).decode("utf-8", errors="replace")
            if not 200 <= status_code < 300:
                raise RuntimeError(f"HTTP {status_code}")
            return status_code, response_body
    except HTTPError as exc:
        # Do not store or display endpoint URLs or credentials in diagnostics.
        try:
            response_body = exc.read(2048).decode("utf-8", errors="replace")
        except Exception:
            response_body = ""
        detail = ""
        discord_code = ""
        try:
            parsed = json.loads(response_body)
            detail = str(parsed.get("description") or parsed.get("message") or "")
            if parsed.get("code") is not None:
                discord_code = f" (API code {parsed['code']})"
        except Exception:
            # Some gateways return an HTML/plain-text 403 page instead of Discord JSON.
            # Never echo response HTML or the webhook URL into the UI/logs.
            if response_body and not response_body.lstrip().startswith(("<", "<!DOCTYPE")):
                detail = re.sub(r"\s+", " ", response_body).strip()[:120]
        if exc.code == 403 and not detail:
            detail = "Forbidden: Discord or an outbound network gateway rejected the request. Verify the webhook URL and whether the NAS/network can reach discord.com."
        raise RuntimeError(f"HTTP {exc.code}{discord_code}" + (f": {detail[:220]}" if detail else "")) from None
    except (URLError, TimeoutError, OSError) as exc:
        reason = getattr(exc, "reason", None)
        if isinstance(reason, TimeoutError):
            raise RuntimeError("Connection timed out") from None
        raise RuntimeError("Unable to reach the service. Check the NAS internet connection and URL.") from None


def send_integration_message(service, message):
    if service not in ("teams", "telegram", "discord"):
        raise RuntimeError("Unknown integration")
    if get_integration_setting(f"{service}_enabled", "0") != "1":
        raise RuntimeError("Integration is turned off in Settings")
    try:
        if service == "discord":
            webhook_url = get_integration_setting("discord_webhook_url", "").strip()
            if not (webhook_url.startswith("https://discord.com/api/webhooks/") or
                    webhook_url.startswith("https://discordapp.com/api/webhooks/")):
                raise RuntimeError("Save a valid Discord webhook URL first")
            # Discord limits message content to 2,000 characters. Suppress accidental @everyone/@here pings.
            _post_json(webhook_url, {"content": str(message)[:1900], "allowed_mentions": {"parse": []}})
        elif service == "teams":
            webhook_url = get_integration_setting("teams_webhook_url", "").strip()
            if not webhook_url or not webhook_url.lower().startswith("https://"):
                raise RuntimeError("Save a valid HTTPS Teams Workflows webhook URL first")
            # Microsoft Teams Workflows accepts a simple {"text": "..."} JSON body.
            _post_json(webhook_url, {"text": message})
        else:
            token = get_integration_setting("telegram_bot_token", "").strip()
            chat_id = get_integration_setting("telegram_chat_id", "").strip()
            if not token or not chat_id:
                raise RuntimeError("Save the Telegram Bot Token and Chat ID first")
            if not token.replace(":", "").replace("_", "").replace("-", "").isalnum():
                raise RuntimeError("Telegram Bot Token format looks invalid")
            url = f"https://api.telegram.org/bot{token}/sendMessage"
            _, response_body = _post_json(url, {"chat_id": chat_id, "text": message, "disable_web_page_preview": True})
            try:
                api_result = json.loads(response_body)
            except json.JSONDecodeError:
                api_result = {}
            if api_result.get("ok") is not True:
                description = str(api_result.get("description") or "Telegram API rejected the message")
                raise RuntimeError(description[:250])
        set_integration_status(service, True, "")
    except Exception as exc:
        set_integration_status(service, False, str(exc))
        raise


def get_telegram_chat_id():
    token = get_integration_setting("telegram_bot_token", "").strip()
    if not token:
        raise RuntimeError("Save the Telegram Bot Token first")
    url = f"https://api.telegram.org/bot{token}/getUpdates"
    # getUpdates is a GET method in the Telegram Bot API; only retrieve a few recent updates.
    request = URLRequest(url, headers={"Accept": "application/json"}, method="GET")
    try:
        with urlopen(request, timeout=8) as response:
            payload = json.loads(response.read(65536).decode("utf-8", errors="replace"))
    except HTTPError as exc:
        raise RuntimeError(f"Telegram API HTTP {exc.code}") from None
    except (URLError, TimeoutError, OSError):
        raise RuntimeError("Unable to reach Telegram. Check the NAS internet connection.") from None
    if payload.get("ok") is not True:
        raise RuntimeError(str(payload.get("description") or "Telegram could not read recent bot messages")[:250])
    updates = payload.get("result") or []
    for item in reversed(updates):
        message = item.get("message") or item.get("edited_message") or item.get("channel_post") or item.get("edited_channel_post") or {}
        chat = message.get("chat") or {}
        if chat.get("id") is not None:
            return str(chat["id"]), str(chat.get("title") or chat.get("username") or chat.get("first_name") or "Telegram chat")
    raise RuntimeError("No message found yet. Open the bot in Telegram, press Start and send it a message, then try again.")


def notification_message(ticket, event_key):
    labels = {
        "new_ticket": "🆕 TICKET MỚI",
        "high_ticket": "🚨 TICKET MỨC ĐỘ CAO",
        "ticket_taken": "🛠 TICKET ĐÃ ĐƯỢC NHẬN XỬ LÝ",
        "ticket_closed": "✅ TICKET ĐÃ ĐÓNG",
        "ticket_updated": "✏️ TICKET ĐÃ CẬP NHẬT",
    }
    title = labels.get(event_key, "🔔 CẬP NHẬT TICKET")
    return "\n".join([
        f"{title} — FMP IT Service Desk",
        f"Mã ticket: {ticket.get('code') or ('#' + str(ticket.get('id', '')))}",
        f"Trạng thái: {ticket.get('status') or 'Mới'}",
        f"Mức độ: {ticket.get('severity') or '—'}",
        f"Thiết bị: {ticket.get('device') or '—'}" + (f" ({ticket.get('device_code')})" if ticket.get('device_code') else ""),
        f"Bộ phận: {ticket.get('department') or '—'} | Vị trí: {ticket.get('location') or '—'}",
        f"Người báo: {ticket.get('reporter') or '—'} | Điện thoại: {ticket.get('phone') or '—'}",
        f"Loại sự cố: {ticket.get('issue_type') or '—'}",
        f"Mô tả: {ticket.get('issue') or '—'}",
        f"Người xử lý: {ticket.get('handler') or 'Chưa phân công'}",
        f"Thời gian: {now()}",
        *([f"⭐ Đánh giá dịch vụ: {ticket.get('feedback_url')}" ] if event_key == "ticket_closed" and ticket.get("feedback_url") else []),
    ])


def dispatch_ticket_notification(ticket, event_key):
    setting_key = NOTIFICATION_EVENT_KEYS.get(event_key)
    if not setting_key or get_integration_setting(setting_key, "1") != "1":
        return
    message = notification_message(dict(ticket), event_key)
    for service in ("teams", "telegram", "discord"):
        if get_integration_setting(f"{service}_enabled", "0") != "1":
            continue
        if not integration_credentials_ready(service):
            set_integration_status(service, False, "Configuration is incomplete")
            continue
        try:
            send_integration_message(service, message)
        except Exception:
            # Service error details are stored by send_integration_message; a failed
            # notification must never break ticket creation or status changes.
            pass


def integration_banner_html(req):
    """Show an online notification provider as the primary banner item; list other providers below."""
    english = req.session.get("lang") == "en"
    teams_connected = integration_is_connected("teams")
    telegram_connected = integration_is_connected("telegram")
    discord_connected = integration_is_connected("discord")
    connected_label = "Connected" if english else "Đã kết nối"
    disconnected_label = "Not connected" if english else "Chưa kết nối"
    no_online_label = "No service online" if english else "Không có dịch vụ online"
    # Inline brand-inspired SVG marks keep the header crisp without external image requests.
    teams_icon = '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path fill="#6264A7" d="M15.2 5.2h5.1c.9 0 1.7.8 1.7 1.7v9.2c0 .9-.8 1.7-1.7 1.7h-5.1z"/><circle cx="16.9" cy="3.9" r="2.1" fill="#8B8CC7"/><path fill="#464775" d="M2 7.2h12.6v11.1c0 1-.8 1.7-1.7 1.7H3.7c-1 0-1.7-.8-1.7-1.7z"/><path fill="#fff" d="M5 10h6.6v1.8H9.3v5.5H7.4v-5.5H5z"/></svg>'
    telegram_icon = '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><circle cx="12" cy="12" r="12" fill="#229ED9"/><path fill="#fff" d="m5.1 11.6 12.8-4.9c.6-.2 1.1.2.9.9l-2.2 10.3c-.1.7-.6.8-1.1.5l-3.4-2.5-1.7 1.7c-.2.2-.5.2-.6-.2l.3-3.5 6.4-5.8c.3-.3-.1-.4-.4-.2l-7.9 5-3.4-1.1c-.7-.2-.7-.7.3-1.1z"/></svg>'
    discord_icon = '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path fill="#5865F2" d="M19.7 5.2A18 18 0 0 0 15.3 4l-.5 1a16 16 0 0 0-5.6 0l-.5-1a18 18 0 0 0-4.4 1.2C1.5 9.4.7 13.5 1.1 17.6a18 18 0 0 0 5.4 2.7l1.2-1.9-1.9-.9.5-.4c3.7 1.7 7.7 1.7 11.4 0l.5.4-1.9.9 1.2 1.9a18 18 0 0 0 5.4-2.7c.5-4.7-.8-8.7-3.2-12.4ZM8.7 14.5c-1 0-1.8-.9-1.8-2s.8-2 1.8-2 1.8.9 1.8 2-.8 2-1.8 2Zm6.6 0c-1 0-1.8-.9-1.8-2s.8-2 1.8-2 1.8.9 1.8 2-.8 2-1.8 2Z"/></svg>'
    neutral_icon = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M18 8a6 6 0 0 0-12 0c0 7-3 7-3 9h18c0-2-3-2-3-9"/><path d="M10 21h4"/></svg>'
    chevron = '<svg class="integration-header-chevron" viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6"/></svg>'
    service_rows = [
        ("Microsoft Teams", teams_icon, "teams", teams_connected),
        ("Telegram", telegram_icon, "telegram", telegram_connected),
        ("Discord", discord_icon, "discord", discord_connected),
    ]
    # Put connected services first. The first connected service replaces the fixed Teams primary label.
    service_rows.sort(key=lambda item: (not item[3], ["teams", "telegram", "discord"].index(item[2])))
    primary = next((item for item in service_rows if item[3]), None)
    if primary:
        primary_name, primary_icon, primary_service, primary_connected = primary
        remaining_rows = [item for item in service_rows if item[2] != primary_service]
        primary_status = connected_label
        primary_status_cls = "connected"
    else:
        # Do not keep an offline provider pinned as the primary item when none are online.
        primary_name, primary_icon, primary_service = no_online_label, neutral_icon, "none"
        remaining_rows = service_rows
        primary_status = disconnected_label
        primary_status_cls = "disconnected"
    title = ("Notification status; online service first, offline services below" if english
             else "Trạng thái thông báo; dịch vụ online hiển thị trước, dịch vụ offline ở bên dưới")
    settings_label = "Notification settings" if english else "Cài đặt thông báo"
    service_items_html = "".join(
        f'<div class="integration-header-item" role="status">'
        f'<span class="integration-header-item-icon {service}">{icon_svg}</span>'
        f'<span class="integration-header-item-copy"><strong>{name}</strong>'
        f'<small class="integration-header-state {"connected" if connected else "disconnected"}"><i aria-hidden="true"></i>{connected_label if connected else disconnected_label}</small></span></div>'
        for name, icon_svg, service, connected in remaining_rows
    )
    return (
        '<details class="integration-header-dropdown">'
        f'<summary class="integration-header-trigger" title="{E(title)}" aria-label="{E(title)}">'
        f'<span class="integration-header-main-icon {E(primary_service)}">{primary_icon}</span>'
        f'<span class="integration-header-main-copy"><strong>{E(primary_name)}</strong>'
        f'<small class="integration-header-state {primary_status_cls}"><i aria-hidden="true"></i>{primary_status}</small></span>'
        f'{chevron}</summary>'
        f'<div class="integration-header-menu" role="menu">{service_items_html}'
        f'<a class="integration-menu-settings" href="/settings?tab=notifications" role="menuitem">'
        f'<span aria-hidden="true">⚙</span><span>{E(settings_label)}</span><span aria-hidden="true">›</span></a>'
        '</div></details>'
    )


def pw_hash(pw):
    salt = secrets.token_hex(16)
    return salt + "$" + hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), 300000).hex()

def pw_check(pw, stored):
    try:
        salt, digest = stored.split("$", 1)
        got = hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), 300000).hex()
        return secrets.compare_digest(got, digest)
    except (ValueError, TypeError, AttributeError):
        return False

def init_db():
    query("CREATE TABLE IF NOT EXISTS integration_settings(key TEXT PRIMARY KEY, value TEXT NOT NULL)", commit=True)
    query("""CREATE TABLE IF NOT EXISTS users(
      id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
      full_name TEXT, department TEXT, phone TEXT, role TEXT NOT NULL DEFAULT 'user', active INTEGER NOT NULL DEFAULT 1)""", commit=True)
    query("""CREATE TABLE IF NOT EXISTS tickets(
      id INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT, created_at TEXT, user_id INTEGER,
      device TEXT, device_code TEXT, department TEXT, location TEXT, issue TEXT, issue_type TEXT, severity TEXT,
      reporter TEXT, phone TEXT, status TEXT DEFAULT 'Mới', handler TEXT, note TEXT, public_note TEXT, received_at TEXT, closed_at TEXT,
      created_by_user_id INTEGER)""", commit=True)
    query("""CREATE TABLE IF NOT EXISTS ticket_feedback_tokens(
      ticket_id INTEGER PRIMARY KEY, token TEXT UNIQUE NOT NULL, created_at TEXT NOT NULL)""", commit=True)
    query("""CREATE TABLE IF NOT EXISTS ticket_feedback(
      id INTEGER PRIMARY KEY AUTOINCREMENT, ticket_id INTEGER NOT NULL UNIQUE,
      rating INTEGER NOT NULL CHECK(rating BETWEEN 1 AND 5), comment TEXT NOT NULL DEFAULT '',
      reviewer_user_id INTEGER, reviewer_name TEXT, created_at TEXT NOT NULL)""", commit=True)
    user_cols = [r["name"] for r in query("PRAGMA table_info(users)")]
    if "passwordless_login" not in user_cols:
        query("ALTER TABLE users ADD COLUMN passwordless_login INTEGER NOT NULL DEFAULT 0", commit=True)
    cols = [r["name"] for r in query("PRAGMA table_info(tickets)")]
    if "issue_type" not in cols:
        query("ALTER TABLE tickets ADD COLUMN issue_type TEXT", commit=True)
    if "received_at" not in cols:
        query("ALTER TABLE tickets ADD COLUMN received_at TEXT", commit=True)
    if "public_note" not in cols:
        query("ALTER TABLE tickets ADD COLUMN public_note TEXT", commit=True)
    if "created_by_user_id" not in cols:
        # Additive migration only: keep all existing ticket IDs and data intact.
        query("ALTER TABLE tickets ADD COLUMN created_by_user_id INTEGER", commit=True)
    # Remove legacy handling-note data. Keep the column only for SQLite compatibility with existing databases.
    if "note" in cols:
        query("UPDATE tickets SET note=NULL WHERE note IS NOT NULL", commit=True)
    # Administrator accounts must remain active and cannot be locked from account management.
    query("UPDATE users SET active=1 WHERE role='admin' AND active<>1", commit=True)
    # Keep the SQLite data aligned with the approved bilingual department list.
    # Legacy/removed/custom department names are retained as records but normalized to "Khác".
    for table in ("users", "tickets"):
        for row in query(f"SELECT id, department FROM {table}"):
            canonical = normalize_department(row["department"])
            if (row["department"] or "") != canonical:
                query(f"UPDATE {table} SET department=? WHERE id=?", (canonical, row["id"]), commit=True)
    # In-place indexes for department filtering and feedback lookups; existing records stay intact.
    query("CREATE INDEX IF NOT EXISTS idx_tickets_department ON tickets(department)", commit=True)
    query("CREATE INDEX IF NOT EXISTS idx_tickets_user_department ON tickets(user_id, department)", commit=True)
    query("CREATE INDEX IF NOT EXISTS idx_feedback_ticket_rating ON ticket_feedback(ticket_id, rating)", commit=True)
    # Small indexes reduce common status/device/date filtering and owner ticket lookups.
    query("CREATE INDEX IF NOT EXISTS idx_tickets_status ON tickets(status)", commit=True)
    query("CREATE INDEX IF NOT EXISTS idx_tickets_device ON tickets(device)", commit=True)
    query("CREATE INDEX IF NOT EXISTS idx_tickets_user_id ON tickets(user_id)", commit=True)
    query("CREATE INDEX IF NOT EXISTS idx_tickets_feedback_ticket ON ticket_feedback(ticket_id)", commit=True)
    # A revision counter lets User/Admin detect live changes without reading and hashing
    # every ticket every few seconds. Triggers preserve sync across all existing write paths.
    query("CREATE TABLE IF NOT EXISTS app_revision (id INTEGER PRIMARY KEY CHECK(id=1), revision INTEGER NOT NULL DEFAULT 0)", commit=True)
    query("INSERT OR IGNORE INTO app_revision(id, revision) VALUES(1, 0)", commit=True)
    for table, action in (("tickets", "INSERT"), ("tickets", "UPDATE"), ("tickets", "DELETE"),
                          ("ticket_feedback", "INSERT"), ("ticket_feedback", "UPDATE"), ("ticket_feedback", "DELETE")):
        trigger = f"trg_revision_{table}_{action.lower()}"
        query(f"CREATE TRIGGER IF NOT EXISTS {trigger} AFTER {action} ON {table} BEGIN UPDATE app_revision SET revision=revision+1 WHERE id=1; END", commit=True)
    # One-time-compatible normalization: old detailed issue choices become categories.
    # The detailed narrative remains in tickets.issue and is not changed.
    for old_issue, category in LEGACY_ISSUE_CATEGORY.items():
        query("UPDATE tickets SET issue_type=? WHERE issue_type=?", (category, old_issue), commit=True)
    if not query("SELECT id FROM users LIMIT 1", one=True):
        query("INSERT INTO users(username,password_hash,full_name,department,role) VALUES(?,?,?,?,?)",
              (ADMIN_USER, pw_hash(ADMIN_PASS), ADMIN_NAME, "Khác", "admin"), commit=True)

def page(req, title, body, user=None):
    lang = req.session.get("lang", "vi")
    toggle = f'<a class="lang-switch" href="/lang/{"en" if lang=="vi" else "vi"}">{"EN" if lang=="vi" else "VI"}</a>'
    def nav_icon(kind):
        paths = {
            "home": '<path d="m3 10 9-7 9 7v10a1 1 0 0 1-1 1h-5v-7H9v7H4a1 1 0 0 1-1-1Z"/>',
            "tickets": '<rect x="4" y="4" width="16" height="16" rx="3"/><path d="M8 9h8M8 13h8M8 17h5"/>',
            "new": '<circle cx="12" cy="12" r="9"/><path d="M12 8v8M8 12h8"/>',
            "users": '<circle cx="9" cy="8" r="3"/><path d="M3 20v-1a6 6 0 0 1 12 0v1M16 5.5a3 3 0 0 1 0 5.8M18 14a5 5 0 0 1 3 4.5V20"/>',
            "settings": '<circle cx="12" cy="12" r="3"/><path d="m19.4 15 .1.1 1.4 1.1-1.4 2.4-1.7-.6a8 8 0 0 1-1.8 1l-.3 1.8h-2.8l-.3-1.8a8 8 0 0 1-1.8-1l-1.7.6-1.4-2.4 1.4-1.1a7 7 0 0 1 0-2l-1.4-1.1 1.4-2.4 1.7.6a8 8 0 0 1 1.8-1l.3-1.8h2.8l.3 1.8a8 8 0 0 1 1.8 1l1.7-.6 1.4 2.4-1.4 1.1a7 7 0 0 1 0 1.9Z"/>',
            "logout": '<path d="M10 17l5-5-5-5M15 12H3"/><path d="M12 3h6a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-6"/>'
        }
        return f'<svg viewBox="0 0 24 24" width="19" height="19" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{paths[kind]}</svg>'
    links = []
    if user:
        # Keep all existing navigation destinations; use a consistent modern line-icon set.
        if user["role"] == "admin":
            links.append(f'<a class="side-link {"active" if req.url.path == "/" else ""}" href="/"><span class="side-icon">{nav_icon("home")}</span><span>{tr(req,"home")}</span></a>')
        if user["role"] == "admin":
            links.append(f'<a class="side-link {"active" if req.url.path == "/admin" else ""}" href="/admin"><span class="side-icon">{nav_icon("tickets")}</span><span>{tr(req,"ticket_list")}</span></a>')
        else:
            links.append(f'<a class="side-link {"active" if req.url.path == "/" else ""}" href="/"><span class="side-icon">{nav_icon("tickets")}</span><span>{tr(req,"my")}</span></a>')
        links.append(f'<a class="side-link {"active" if req.url.path == "/new" else ""}" href="/new"><span class="side-icon">{nav_icon("new")}</span><span>{tr(req,"new")}</span></a>')
        if user["role"] == "admin":
            links.append(f'<a class="side-link {"active" if req.url.path == "/users" else ""}" href="/users"><span class="side-icon">{nav_icon("users")}</span><span>{tr(req,"users")}</span></a>')
            links.append(f'<a class="side-link {"active" if req.url.path == "/settings" else ""}" href="/settings"><span class="side-icon">{nav_icon("settings")}</span><span>{tr(req,"settings")}</span></a>')
        links.append(f'<a class="side-link" href="/logout"><span class="side-icon">{nav_icon("logout")}</span><span>{tr(req,"logout")}</span></a>')
    nav = "".join(links)
    if user:
        notification = ""
        if user["role"] == "admin":
            new_count_row = query("SELECT COUNT(*) AS n FROM tickets WHERE status='Mới'", one=True)
            new_count = int(new_count_row["n"] if new_count_row else 0)
            count_badge = f'<span class="notification-count">{new_count if new_count < 100 else "99+"}</span>' if new_count else ""
            bell_svg = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M18 8a6 6 0 0 0-12 0c0 7-3 7-3 9h18c0-2-3-2-3-9"/><path d="M10 21h4"/></svg>'
            notification = f'<a class="top-notification" href="/admin?all_dates=1&amp;status=M%E1%BB%9Bi" title="Ticket mới" aria-label="Ticket mới">{bell_svg}{count_badge}</a>'
        integration_indicator = integration_banner_html(req) if user["role"] == "admin" else ""
        top_divider = '<span class="top-divider"></span>' if user["role"] == "admin" else ""
        hotline_banner = ""
        if user["role"] == "user":
            hotline_phone = get_integration_setting("hotline_phone", "").strip()
            hotline_email = get_integration_setting("hotline_email", "").strip()
            hotline_label = "IT Support" if lang == "en" else "Hỗ trợ IT"
            hotline_missing = "Not configured" if lang == "en" else "Chưa cập nhật số"
            hotline_number_html = f'<span class="hotline-value">{E(hotline_phone)}</span>' if hotline_phone else f'<span class="hotline-value hotline-not-configured">{E(hotline_missing)}</span>'
            hotline_email_html = f'<span class="hotline-value">{E(hotline_email)}</span>' if hotline_email else ''
            hotline_rows = f'<div class="hotline-contact-row"><span class="hotline-contact-label">Hotline:</span>{hotline_number_html}</div>'
            if hotline_email_html:
                hotline_rows += f'<div class="hotline-contact-row"><span class="hotline-contact-label">Email:</span>{hotline_email_html}</div>'
            headset_svg = '<svg viewBox="0 0 48 48" width="23" height="23" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 25v-4a16 16 0 0 1 32 0v4"/><path d="M8 24h7v12h-3a4 4 0 0 1-4-4zM40 24h-7v12h3a4 4 0 0 0 4-4z"/><path d="M33 37a11 11 0 0 1-9 5h-4"/></svg>'
            hotline_banner = f'<div class="user-hotline-banner"><span class="hotline-it-icon" aria-hidden="true">{headset_svg}</span><div class="hotline-support-copy"><strong>{E(hotline_label)}</strong>{hotline_rows}</div></div>'
            sidebar_footer = ''
        else:
            sidebar_footer = ''
        shell = f"""<header class="topbar"><div class="brand"><span class="brand-mark" aria-hidden="true"><svg viewBox="0 0 48 48" width="34" height="34" fill="none" stroke="currentColor" stroke-width="2.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="24" cy="24" r="21" stroke="rgba(255,255,255,.72)" stroke-width="1.2"/><path d="M10 25v-3a14 14 0 0 1 28 0v3"/><rect x="8" y="23" width="7" height="12" rx="3.5" fill="currentColor" stroke="none"/><rect x="33" y="23" width="7" height="12" rx="3.5" fill="currentColor" stroke="none"/><path d="M35 36a12 12 0 0 1-10 5h-3"/></svg></span><strong>IT SERVICE DESK</strong></div>{hotline_banner}<div class="topbar-right">{notification}{integration_indicator}{top_divider}<span class="user-mark role-mark role-{E(user['role'])}" title="{tr(req,'role_admin') if user['role']=='admin' else tr(req,'role_user')}" aria-label="{tr(req,'role_admin') if user['role']=='admin' else tr(req,'role_user')}"><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="8" r="3.5"/><path d="M5 20c.5-4 3-6 7-6s6.5 2 7 6"/></svg></span><span class="user-name">{E(user['full_name'] or user['username'])}</span>{toggle}</div></header><aside class="sidebar"><nav>{nav}</nav>{sidebar_footer}</aside><main class="app-main">{body}</main><footer class="app-credit">DESIGNED BY <strong>IT.anhnvn</strong> · IT SERVICE DESK</footer><script>(function(){{function paintSeveritySelect(sel){{var v=(sel.value||"").trim();var map={{"Cao":"#b42332","High":"#b42332","Trung bình":"#94610f","Medium":"#94610f","Thấp":"#1d4ed8","Low":"#1d4ed8"}};var c=map[v]||"#19324f";sel.style.color=c;sel.style.fontWeight="800";Array.prototype.forEach.call(sel.options,function(o){{var oc=map[(o.value||"").trim()]||"#19324f";o.style.color=oc;o.style.fontWeight="800";}});}}document.querySelectorAll("select[name=severity]").forEach(function(sel){{paintSeveritySelect(sel);sel.addEventListener("change",function(){{paintSeveritySelect(sel);}});}});}})();</script>"""
    else:
        login_hotline_phone = get_integration_setting("hotline_phone", "").strip()
        login_hotline_email = get_integration_setting("hotline_email", "").strip()
        login_hotline_label = "IT Support" if lang == "en" else "Hỗ trợ IT"
        login_hotline_missing = "Not configured" if lang == "en" else "Chưa cập nhật số"
        login_hotline_number_html = f'<span class="hotline-value">{E(login_hotline_phone)}</span>' if login_hotline_phone else f'<span class="hotline-value hotline-not-configured">{E(login_hotline_missing)}</span>'
        login_hotline_rows = f'<div class="hotline-contact-row"><span class="hotline-contact-label">Hotline:</span>{login_hotline_number_html}</div>'
        if login_hotline_email:
            login_hotline_rows += f'<div class="hotline-contact-row"><span class="hotline-contact-label">Email:</span><span class="hotline-value">{E(login_hotline_email)}</span></div>'
        login_headset_svg = '<svg viewBox="0 0 48 48" width="23" height="23" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 25v-4a16 16 0 0 1 32 0v4"/><path d="M8 24h7v12h-3a4 4 0 0 1-4-4zM40 24h-7v12h3a4 4 0 0 0 4-4z"/><path d="M33 37a11 11 0 0 1-9 5h-4"/></svg>'
        login_hotline_banner = f'<div class="user-hotline-banner"><span class="hotline-it-icon" aria-hidden="true">{login_headset_svg}</span><div class="hotline-support-copy"><strong>{E(login_hotline_label)}</strong>{login_hotline_rows}</div></div>'
        shell = f"""<header class="topbar"><div class="brand"><span class="brand-mark" aria-hidden="true"><svg viewBox="0 0 48 48" width="34" height="34" fill="none" stroke="currentColor" stroke-width="2.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="24" cy="24" r="21" stroke="rgba(255,255,255,.72)" stroke-width="1.2"/><path d="M10 25v-3a14 14 0 0 1 28 0v3"/><rect x="8" y="23" width="7" height="12" rx="3.5" fill="currentColor" stroke="none"/><rect x="33" y="23" width="7" height="12" rx="3.5" fill="currentColor" stroke="none"/><path d="M35 36a12 12 0 0 1-10 5h-3"/></svg></span><strong>IT SERVICE DESK</strong></div>{login_hotline_banner}{toggle}</header><main class="login-main">{body}</main>"""
    return HTMLResponse(f"""<!doctype html><html lang="{lang}"><head><meta charset="utf-8"><title>{E(title)}</title>
    <style>
    body{{font-family:Inter,Segoe UI,Arial,sans-serif;background:linear-gradient(135deg,#f3f8fd 0%,#eef5fb 100%);color:#19324f;margin:0}}header{{background:linear-gradient(110deg,#fff 0%,#f5faff 100%);color:#123d70;padding:17px 3%;display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;border-bottom:1px solid #dce9f6;box-shadow:0 3px 14px rgba(30,75,120,.06)}}header a{{color:white}}main{{max-width:1450px;margin:24px auto;padding:0 18px}}.box{{background:rgba(255,255,255,.96);padding:22px;border:1px solid #e1ebf5;border-radius:14px;margin:14px 0;overflow-x:auto;box-shadow:0 5px 18px rgba(30,70,110,.045)}}input,select,textarea,button{{font:inherit;padding:9px;max-width:100%;box-sizing:border-box}}input,select,textarea{{border:1px solid #bbc4d0;border-radius:8px}}label{{display:block;margin:12px 0 5px;font-size:13px;font-weight:650;color:#36516f}}.w{{width:100%}}button,.btn{{background:linear-gradient(135deg,#087fc2,#1769aa);color:white;border:0;border-radius:8px;padding:9px 13px;cursor:pointer;text-decoration:none;display:inline-block}}table{{border-collapse:collapse;width:100%;background:white}}td,th{{padding:9px;border-bottom:1px solid #ddd;text-align:left;vertical-align:top}}th{{background:#e8edf4}}.ticket-table{{border-collapse:separate;border-spacing:0;width:100%;min-width:720px;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden}}.ticket-table thead th{{background:#f1f5f9;color:#475569;font-size:12px;text-transform:uppercase;letter-spacing:.035em;font-weight:700;border-bottom:1px solid #dbe3ed;padding:13px 12px}}.ticket-table tbody td{{padding:13px 12px;border-bottom:1px solid #e8edf3;border-right:1px solid #edf2f8;transition:background .15s ease}}.ticket-table tbody td:last-child{{border-right:0}}.ticket-table tbody tr:last-child td{{border-bottom:0}}.ticket-table tbody tr:hover td{{filter:brightness(.985)}}.ticket-row-new td{{background:#f8fbff}}.ticket-row-new td:first-child{{box-shadow:inset 3px 0 #60a5fa}}.ticket-row-progress td{{background:#fbfaff}}.ticket-row-progress td:first-child{{box-shadow:inset 3px 0 #a78bfa}}.ticket-row-closed td{{background:#f5fcf7}}.ticket-row-closed td:first-child{{box-shadow:inset 3px 0 #4ade80}}.ticket-code{{font-weight:700;text-decoration:none;color:#174a7e;white-space:nowrap}}.ticket-code:hover{{text-decoration:underline}}.ticket-workflow{{display:block;margin-top:6px;padding:4px 7px;border-radius:6px;font-size:11px;line-height:1.45;font-weight:600;max-width:250px;overflow-wrap:anywhere}}.ticket-workflow-received{{background:#fff7dd;color:#92400e;border:1px solid #fde68a}}.ticket-workflow-closed{{background:#dcfce7;color:#166534;border:1px solid #bbf7d0}}.ticket-table .pill{{padding:4px 8px;font-size:11px}}.grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px}}.msg{{color:#a21d1d}}small,.help{{color:#555;font-size:13px}}a{{color:#145b9a}}header a{{color:white}}.pill{{display:inline-block;padding:4px 9px;border-radius:999px;font-size:12px;font-weight:700;white-space:nowrap}}.sev-low{{background:#dcfce7;color:#166534}}.sev-medium{{background:#fef3c7;color:#92400e}}.sev-high{{background:#fee2e2;color:#991b1b}}.st-new{{background:#dbeafe;color:#1d4ed8}}.st-progress{{background:#ede9fe;color:#6d28d9}}.st-closed{{background:#d1fae5;color:#065f46}}.filter-grid,.user-edit-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:10px;align-items:end}}.user-card{{border:1px solid #d9e0e8;border-radius:10px;padding:16px;margin:14px 0;background:#fff}}.user-card-head,.user-actions{{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}}.user-meta{{color:#596579;font-size:13px}}.btn-secondary{{background:#5b6573}}.btn-take{{background:#2563eb;color:#fff}}.btn-close{{background:#15803d;color:#fff}}.btn-done{{background:#dcfce7;color:#166534;border:1px solid #86efac;border-radius:6px;padding:9px 13px;font-weight:700;cursor:default}}.workflow-time{{display:block;font-size:12px;font-weight:400;margin-top:4px;opacity:.95}}.detail-time{{display:inline-flex;align-items:center;gap:8px;padding:8px 11px;border-radius:9px;font-size:15px;font-weight:700;line-height:1.4;white-space:normal;overflow-wrap:anywhere}}.detail-time-mark{{font-size:17px;line-height:1;flex-shrink:0}}.detail-time-progress{{background:#fef3c7;color:#92400e;border:1px solid #fcd34d}}.detail-time-closed{{background:#dcfce7;color:#166534;border:1px solid #86efac}}.detail-time-pending{{background:#f1f5f9;color:#64748b;border:1px solid #e2e8f0;font-size:13px;font-weight:500}}.detail-field-received_at,.detail-field-closed_at{{padding:12px;border-radius:10px}}.detail-field-received_at{{background:#fffbeb;border:1px solid #fde68a}}.detail-field-closed_at{{background:#f0fdf4;border:1px solid #bbf7d0}}.detail-field-received_at strong{{color:#92400e}}.detail-field-closed_at strong{{color:#166534}}.workflow-panel{{display:block;margin:22px 0 8px;padding:22px;background:linear-gradient(135deg,#f8fafc 0%,#f1f5f9 100%);border:1px solid #dbe3ed;border-radius:14px}}
    .workflow-heading{{display:flex;align-items:center;gap:10px;margin:0 0 16px;color:#243449;font-size:18px;font-weight:700}}
    .workflow-heading .workflow-mark{{display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:10px;background:#dbeafe;color:#1d4ed8;font-size:18px}}
    .workflow-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:16px;align-items:stretch}}
    .workflow-card{{display:flex;flex-direction:column;gap:12px;min-width:0;padding:18px;background:#fff;border:1px solid #dbe3ed;border-radius:12px;box-shadow:0 2px 5px rgba(15,23,42,.04)}}
    .workflow-card-title{{display:flex;align-items:center;gap:9px;font-weight:700;color:#334155}}
    .workflow-icon{{display:inline-flex;align-items:center;justify-content:center;width:32px;height:32px;border-radius:9px;flex-shrink:0}}
    .workflow-icon-blue{{background:#dbeafe;color:#1d4ed8}}.workflow-icon-amber{{background:#fef3c7;color:#92400e}}.workflow-icon-green{{background:#dcfce7;color:#166534}}
    .workflow-time{{display:block;font-size:13px;font-weight:500;margin-top:5px;color:#166534;line-height:1.5;overflow-wrap:anywhere}}
    .workflow-meta{{font-size:13px;color:#64748b;line-height:1.5}}
    .workflow-panel form{{margin:0}}.workflow-panel label{{margin:0 0 7px;font-size:14px;font-weight:600;color:#334155}}
    .workflow-panel textarea{{width:100%;min-height:100px;resize:vertical;background:#fff}}
    .workflow-panel .btn-take,.workflow-panel .btn-close{{width:100%;padding:11px 16px;font-weight:700;border-radius:8px;transition:filter .15s ease}}
    .workflow-panel .btn-take:hover,.workflow-panel .btn-close:hover{{filter:brightness(.94)}}
    .workflow-panel .btn-done{{display:block;width:100%;text-align:left;cursor:default;border-radius:8px;padding:12px 14px}}
    .workflow-complete{{background:#f0fdf4;border-color:#bbf7d0}}
    .workflow-pending{{background:#fff;border-color:#fde68a}}.note-editor-panel{{background:linear-gradient(135deg,#f3f8ff,#f8fbff);border-color:#cfe3ff}}.workflow-close-action{{background:linear-gradient(135deg,#f0fdf4,#fff);border-color:#a7f3d0}}
    .workflow-actions{{display:flex;flex-direction:column;gap:12px;margin-top:auto}}
    .workflow-note-help{{font-size:12px;color:#64748b;margin-top:-5px}}
    /* Compact, consistent ticket detail styling for both admin and regular users */
    .detail-field{{min-width:0;padding:14px 14px;border:1px solid #e3edf7;background:#fff;border-radius:12px;align-self:stretch;box-shadow:0 2px 6px rgba(30,70,110,.025)}}
    .detail-field strong{{display:block;margin-bottom:9px;font-size:13px;color:#36516f;line-height:1.35;letter-spacing:.01em}}
    .detail-field p{{margin:0;line-height:1.5;overflow-wrap:anywhere;min-height:21px}}
    .detail-field-handler .handler-value{{display:inline-flex;align-items:center;gap:8px;max-width:100%;color:#202733;font-weight:600;line-height:1.4;overflow-wrap:anywhere}}
    .handler-avatar{{display:inline-flex;align-items:center;justify-content:center;flex:0 0 28px;width:28px;height:28px;border-radius:8px;background:#dbeafe;color:#1d4ed8;border:1px solid #bfdbfe}}
    .handler-avatar svg{{display:block;width:15px;height:15px}}
    .handler-name{{min-width:0;overflow-wrap:anywhere}}
    .detail-time{{display:inline-flex;align-items:center;gap:6px;padding:4px 8px;border-radius:7px;font-size:12px;font-weight:700;line-height:1.35;white-space:normal;overflow-wrap:anywhere}}
    .detail-time-mark{{font-size:13px}}
    .detail-time-pending{{font-size:12px;font-weight:500;padding:3px 7px}}
    .detail-field-received_at,.detail-field-closed_at,.detail-field-note{{padding:10px 12px;border-radius:9px;align-self:stretch;min-width:0}}
    .detail-field-received_at p,.detail-field-closed_at p,.detail-field-note p{{min-height:22px}}
    .detail-field-received_at{{background:#fffbeb;border:1px solid #fde68a}}
    .detail-field-closed_at{{background:#f0fdf4;border:1px solid #bbf7d0}}
    .detail-field-note{{background:#f0f7ff;border:1px solid #d5e7fb}}
    .detail-field-received_at strong{{color:#92400e}}
    .detail-field-closed_at strong{{color:#166534}}
    .detail-field-note strong{{color:#1d4ed8}}
    .detail-field-note p{{font-size:14px;color:#334155;white-space:pre-wrap}}
    /* Visual hierarchy for operational details: issue, department, location and assigned IT */
    .detail-field-issue{{grid-column:span 2;padding:12px 14px;background:#f8fafc;border:1px solid #cbd5e1;border-left:4px solid #3b82f6;border-radius:10px}}
    .detail-field-issue strong{{color:#1d4ed8}}
    .detail-field-issue p{{font-size:14px;color:#1e293b;white-space:pre-wrap}}
    .detail-tag{{display:inline-flex;align-items:center;gap:6px;padding:4px 9px;border-radius:7px;font-size:13px;font-weight:600;line-height:1.4;overflow-wrap:anywhere;max-width:100%}}
    .detail-tag-dept{{background:#ecfeff;color:#0e7490;border:1px solid #a5f3fc}}
    .detail-tag-location{{background:#f5f3ff;color:#6d28d9;border:1px solid #ddd6fe}}
    /* Keep assigned IT name plain and aligned with the other field values. */
    .detail-field-handler p{{display:block}}
    .detail-field-device_code p,.detail-field-phone p,.detail-field-time p{{font-variant-numeric:tabular-nums}}
    
    .inline{{display:inline-flex;gap:8px;align-items:center;flex-wrap:wrap}}
    
    
    .detail-grid{{grid-template-columns:repeat(4,minmax(0,1fr));align-items:stretch;gap:18px 16px}}
    .detail-field-department,.detail-field-location{{text-align:center;display:flex;flex-direction:column;align-items:center;justify-content:flex-start}}
    .detail-field-department strong,.detail-field-location strong{{width:100%;text-align:center}}
    .detail-field-department p,.detail-field-location p{{display:flex;justify-content:center;width:100%;margin:0}}
    .detail-field-issue{{grid-column:2 / span 2;background:#f4f9ff;border-color:#cfe3fa;border-left:4px solid #3984f5}}
    .detail-field-received_at{{grid-column:3;grid-row:auto;background:#fffaf0;border-color:#ffdf91}}
    .detail-field-closed_at{{grid-column:3;background:#effcf5;border-color:#b4f0d0}}
    .detail-field-note{{grid-column:1 / span 2;background:#f1f7ff;border-color:#cfe3ff;border-left:4px solid #3984f5}}
    .detail-field-handler{{grid-column:4}}
    .detail-field-handler .handler-value{{display:inline-flex;align-items:center;gap:7px;padding:6px 11px;border-radius:9px;background:#eaf4ff;border:1px solid #b9d9ff;color:#075bb5;font-size:14px;font-weight:700}}
    .detail-field-handler .handler-value:before{{content:'IT';display:inline-flex;align-items:center;justify-content:center;width:23px;height:23px;border-radius:6px;background:#dbeafe;color:#1d4ed8;font-size:10px;font-weight:800}}
    .calendar-toolbar{{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:0 0 14px}}
    .quick-filter{{display:inline-flex;align-items:center;gap:6px;padding:7px 10px;border:1px solid #dbe3ed;border-radius:8px;background:#fff;color:#334155;text-decoration:none;font-size:13px;font-weight:700}}
    .quick-filter:hover{{background:#f8fafc}}
    .quick-filter-hot{{background:#fff1f2;border-color:#fecdd3;color:#be123c}}
    .calendar-input{{min-width:155px}}.status-person{{display:block;margin-top:5px;font-size:11px;color:#526b86;line-height:1.4}}.reporter-chip{{display:inline-flex;align-items:center;gap:6px;padding:5px 9px;border:1px solid #b8e7fb;border-radius:8px;background:#eafaff;color:#087da5;font-size:12px;font-weight:700;line-height:1.35;max-width:100%;overflow-wrap:anywhere}}.detail-field-department,.detail-field-location{{border-right:1px solid #e6eef8;border-bottom:1px solid #e6eef8;border-radius:0;box-shadow:none;background:#fbfdff}}.detail-field-department p,.detail-field-location p{{align-items:center;min-height:32px}}.detail-field-reporter p{{display:flex;align-items:center}}.detail-field-note{{white-space:normal}}.detail-field-note p{{white-space:pre-wrap}}.filter-hint{{font-size:12px;color:#71849a;margin-top:5px}}
    .app-credit{{max-width:1100px;margin:0 auto 18px;padding:8px 16px;text-align:center;color:#8491a3;font-size:11px;letter-spacing:.04em}}
    .user-ticket-list{{display:grid;grid-template-columns:repeat(auto-fit,minmax(330px,1fr));gap:14px}}
    .user-ticket-card{{min-width:0;background:#fff;border:1px solid #dfeaf5;border-radius:13px;padding:16px;box-shadow:0 3px 10px rgba(30,70,110,.04);border-left:4px solid #60a5fa}}
    .user-ticket-new{{background:linear-gradient(135deg,#fff 0%,#f7fbff 100%);border-left-color:#60a5fa}}
    .user-ticket-progress{{background:linear-gradient(135deg,#fff 0%,#faf8ff 100%);border-left-color:#a78bfa}}
    .user-ticket-closed{{background:linear-gradient(135deg,#fff 0%,#f4fcf7 100%);border-left-color:#34d399}}
    .user-ticket-top{{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;padding-bottom:12px;border-bottom:1px solid #e7eef6}}
    .user-ticket-main{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px 12px;padding:14px 0}}
    .user-ticket-main>div{{display:flex;flex-direction:column;align-items:flex-start;gap:5px;min-width:0;overflow-wrap:anywhere}}
    .card-label{{font-size:11px;text-transform:uppercase;letter-spacing:.045em;color:#72869c;font-weight:700}}
    .user-ticket-main strong{{font-size:13px;color:#284461;font-weight:650;line-height:1.45}}
    .user-ticket-bottom{{display:flex;align-items:center;gap:8px;flex-wrap:wrap;border-top:1px solid #e7eef6;padding-top:12px}}
    .ticket-event{{display:inline-flex;align-items:center;gap:5px;padding:5px 8px;background:#fff7df;border:1px solid #fde2a0;border-radius:7px;color:#96500c;font-size:11px;font-weight:650;line-height:1.4}}
    .user-ticket-closed .ticket-event{{background:#e7fbef;border-color:#b8efcf;color:#13764a}}
    .ticket-event-pending{{background:#f1f5f9;border-color:#e2e8f0;color:#64748b}}
    .handler-chip{{display:inline-flex;align-items:center;padding:5px 8px;background:#eaf4ff;border:1px solid #c4e0ff;border-radius:7px;color:#075bb5;font-size:11px;font-weight:700}}

    
    
    /* Modern Medical IT Service Desk shell */
    :root{{--navy:#123d70;--blue:#087fc2;--line:#e2edf8;--muted:#7186a3;--surface:#fff}}
    body{{font-family:Inter,"Segoe UI",Arial,sans-serif;background:linear-gradient(135deg,#f7fbff 0%,#eef5fc 100%);color:#18375f;margin:0;font-size:14px}}
    .topbar{{height:68px;box-sizing:border-box;position:fixed;z-index:20;top:0;left:0;right:0;display:flex;align-items:center;justify-content:space-between;padding:0 28px;background:rgba(255,255,255,.94);border-bottom:1px solid #dce9f6;box-shadow:0 2px 12px rgba(27,73,122,.05);gap:18px}}
    .brand,.topbar-right{{display:flex;align-items:center;gap:13px;min-width:0}}
    .brand-mark{{display:inline-flex;width:38px;height:38px;align-items:center;justify-content:center;border-radius:12px;background:linear-gradient(145deg,#19b7f4,#0874d1);color:white;font-size:30px;font-weight:800;box-shadow:0 4px 10px #087fc233}}
    .brand strong{{font-size:20px;letter-spacing:.015em;color:#123d70;white-space:nowrap}}
    .brand-divider,.top-divider{{height:27px;border-left:1px solid #d5e2f0;margin:0 7px}}
    .brand-credit{{font-size:12px;color:#6f8ab2;white-space:nowrap}}
    .topbar-right{{gap:12px;color:#244b7c}}
    .clinic-mark{{font-size:28px;color:#0996d4;line-height:1}}
    .clinic-copy{{display:flex;flex-direction:column;gap:2px;white-space:nowrap}}
    .clinic-copy strong{{font-size:12px;font-weight:650}}.clinic-copy small{{font-size:10px;color:#7b94b8}}
    .user-mark{{display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:50%;background:#e3f1ff;color:#0875c9;font-size:20px;flex:0 0 34px}}.user-mark svg{{display:block;width:20px;height:20px}}
    .user-name{{font-size:12px;max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
    .lang-switch{{padding:6px 9px;border:1px solid #d8e7f7;border-radius:8px;background:#f8fbff;color:#1b65a5!important;font-weight:700;text-decoration:none}}
    .sidebar{{position:fixed;z-index:10;top:68px;bottom:0;left:0;width:230px;box-sizing:border-box;padding:20px 11px 18px;background:rgba(255,255,255,.72);border-right:1px solid #e0ebf7;display:flex;flex-direction:column;justify-content:space-between}}
    .sidebar nav{{display:flex;flex-direction:column;gap:7px}}
    .side-link{{display:flex;align-items:center;gap:13px;padding:12px 14px;border-radius:10px;text-decoration:none;color:#365780;font-weight:550;transition:background .15s,color .15s}}
    .side-link:hover{{background:#f0f7ff;color:#086fc5}}.side-link.active{{background:linear-gradient(100deg,#e7f3ff,#eaf5ff);color:#006bd1;font-weight:700}}
    .side-icon{{display:inline-flex;width:24px;justify-content:center;font-size:22px;line-height:1;color:#426895}}
    .side-link.active .side-icon{{color:#0878d1}}
    .sidebar-footer{{display:flex;flex-direction:column;align-items:center;gap:8px;text-align:center;color:#8198b7;font-size:12px;padding:16px 5px}}
    .medical-doodle{{font-size:40px;color:#a5bbd6;line-height:1}}.sidebar-footer small{{font-size:10px;color:#a2b3c9}}
    main.app-main{{max-width:none;margin:88px 22px 24px 252px;padding:0;min-width:0}}
    main.login-main{{max-width:600px;margin:110px auto 30px;padding:0 18px}}
    .box{{background:rgba(255,255,255,.96);padding:20px;border:1px solid #e2edf8;border-radius:13px;margin:12px 0;overflow-x:auto;box-shadow:0 4px 14px rgba(30,75,120,.045)}}
    .box h2{{color:#123d70;font-size:20px;margin:0 0 16px}}
    input,select,textarea{{border-color:#d3e1f0;background:#fff;color:#203e64;border-radius:8px;min-height:38px}}
    input:focus,select:focus,textarea:focus{{outline:2px solid #bfdefb;border-color:#70b6f3}}
    label{{color:#385779;font-size:12px;font-weight:700}}
    button,.btn{{background:linear-gradient(135deg,#0788d1,#1266b2);box-shadow:0 2px 5px rgba(8,111,190,.12);border-radius:8px;font-weight:650}}
    .btn-secondary{{background:#64748b}}.btn-close{{background:linear-gradient(135deg,#13a879,#07855f)}}.btn-take{{background:linear-gradient(135deg,#0d91d4,#1769bd)}}
    .ticket-table{{border-color:#e2edf8;border-radius:10px;min-width:780px}}
    .ticket-table thead th{{background:#f5f8fc;color:#456384;text-transform:none;font-size:11px;padding:12px 10px}}
    .ticket-table tbody td{{padding:11px 10px;font-size:12px;border-bottom-color:#edf2f8}}
    .ticket-row-new td{{background:#fbfdff}}.ticket-row-progress td{{background:#fdfcff}}.ticket-row-closed td{{background:#f8fdf9}}
    .pill{{font-size:11px;padding:5px 9px}}.st-new{{background:#e2efff;color:#2363b8}}.st-progress{{background:#efe8ff;color:#7139d5}}.st-closed{{background:#d8f8eb;color:#07845e}}
    .sev-low{{background:#def8f1;color:#07836e}}.sev-medium{{background:#fff0d7;color:#a15b00}}.sev-high{{background:#ffe3e5;color:#c52c3b}}
    .grid{{gap:12px}}
    .detail-grid{{grid-template-columns:repeat(4,minmax(0,1fr));gap:0 0;border:1px solid #e7eff8;border-radius:12px;padding:4px 10px}}
    .detail-field{{border:0;border-right:1px solid #e6eef8;border-radius:0;box-shadow:none;padding:14px 15px;background:transparent}}
    .detail-field:nth-child(4n){{border-right:0}}.detail-field strong{{font-size:12px;color:#3a5b83;margin-bottom:7px}}.detail-field p{{font-size:14px;color:#203e64}}
    .detail-field-issue{{grid-column:auto;background:#f1f8ff;border:1px solid #bcdfff;border-left:3px solid #2990ef;border-radius:9px;margin:5px;padding:12px}}
    .detail-field-department,.detail-field-location{{text-align:left;align-items:flex-start;justify-content:center}}
    .detail-field-department strong,.detail-field-location strong{{text-align:left}}
    .detail-field-received_at{{background:#fffaf0;border:1px solid #ffdc91;border-radius:9px;margin:5px}}
    .detail-field-closed_at{{background:#f0fcf7;border:1px solid #a9f0d5;border-radius:9px;margin:5px}}
    .detail-field-note{{background:#f1f8ff;border:1px solid #bfdefe;border-radius:9px;margin:5px}}
    .detail-field-handler{{background:#f2f8ff;border:1px solid #c6e2ff;border-radius:9px;margin:5px}}
    .detail-field-handler .handler-value{{color:#0874c9;font-size:13px;font-weight:700}}
    .handler-avatar{{display:none}}
    .detail-tag{{border-radius:999px;padding:5px 9px;font-size:12px;font-weight:650}}
    .detail-time{{font-size:12px;padding:5px 8px}}
    .workflow-panel{{background:#f8fbff;border-color:#e0ebf7;padding:16px;border-radius:12px}}
    .workflow-grid{{grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px}}
    .workflow-card{{padding:14px;border-color:#e2edf8;box-shadow:none}}
    .calendar-toolbar{{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}}
    .quick-filter{{border:1px solid #d8e7f6;background:#f8fbff;color:#1e5d9b;padding:8px 11px;border-radius:8px;text-decoration:none;font-size:12px;font-weight:650}}
    .quick-filter-hot{{background:#fff0f1;border-color:#ffd1d5;color:#bd3443}}
    .filter-grid{{gap:12px}}
    .user-ticket-list{{grid-template-columns:repeat(auto-fit,minmax(290px,1fr))}}
    .user-ticket-card{{box-shadow:0 3px 12px rgba(30,70,110,.04);border-radius:12px}}
    .app-credit{{margin-left:240px;max-width:none;padding:10px 20px;color:#91a5bf}}
    
    
    
    /* Final alignment pass: department/location cells share the same visual axis. */
    .detail-field-department,.detail-field-location{{display:flex;flex-direction:column;align-items:center;justify-content:flex-start;text-align:center;background:#fbfdff;border-right:1px solid #e2edf8;border-bottom:1px solid #e2edf8;border-radius:0;box-shadow:none}}
    .detail-field-department strong,.detail-field-location strong{{width:100%;text-align:center;margin-bottom:10px}}
    .detail-field-department p,.detail-field-location p{{display:flex;justify-content:center;align-items:center;width:100%;min-height:34px;margin:0}}
    .reporter-chip{{display:inline-flex;align-items:center;gap:6px;padding:5px 9px;border:1px solid #b8e7fb;border-radius:8px;background:#eafaff;color:#087da5;font-size:12px;font-weight:700;line-height:1.35;max-width:100%;overflow-wrap:anywhere}}
    .detail-field-reporter p{{display:flex;align-items:center;min-height:30px}}
    .note-editor-panel{{background:linear-gradient(135deg,#f3f8ff,#f8fbff);border-color:#cfe3ff}}
    .workflow-close-action{{background:linear-gradient(135deg,#f0fdf4,#fff);border-color:#a7f3d0}}

    /* Reference-matched polish: compact medical dashboard, soft blue surfaces and pill controls */
    :root{{color-scheme:light}}
    body{{background:#f2f7fc;color:#193b67}}
    .topbar{{background:rgba(255,255,255,.96);padding:0 28px}}
    .brand-mark{{border-radius:50%;box-shadow:0 3px 9px #168bd329}}.brand-mark svg{{display:block;width:100%;height:100%;padding:2px;box-sizing:border-box}}
    .sidebar{{background:#f9fcff;border-right-color:#e3edf8}}
    .side-link{{min-height:22px}}
    main.app-main{{margin-top:86px;margin-right:22px}}
    .box{{border-radius:12px;border-color:#e4edf7;box-shadow:0 3px 12px rgba(35,82,130,.035)}}
    .box h2{{font-size:18px}}
    .table-scroll{{width:100%;overflow-x:auto;border:1px solid #e7eef7;border-radius:10px}}
    .ticket-table{{width:100%;min-width:960px;border:0;border-radius:0;border-spacing:0}}
    .ticket-table thead th{{white-space:nowrap;background:#f5f8fc;color:#38577d;font-size:11px;font-weight:700;padding:12px 10px;border-bottom:1px solid #e4edf7}}
    .ticket-table tbody td{{padding:10px;font-size:12px;vertical-align:middle;background:#fff;border-bottom:1px solid #edf2f8;border-right:0}}
    .ticket-table tbody tr:hover td{{background:#f8fbff}}
    .ticket-table tbody tr:last-child td{{border-bottom:0}}
    .ticket-code{{color:#103d78;font-weight:750}}
    .row-number{{color:#607995;width:24px}}
    .device-cell{{display:inline-flex;align-items:center;gap:8px;white-space:nowrap}}
    .mini-icon{{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:6px;background:#eef6ff;color:#1269bd;font-size:17px}}
    .issue-cell{{max-width:180px;min-width:90px;overflow-wrap:anywhere}}
    .date-cell{{white-space:nowrap;font-variant-numeric:tabular-nums}}
    .dept-chip{{border:0;border-radius:999px;background:#e7f8ff;color:#087eaa;white-space:nowrap}}
    .handler-chip{{display:inline-flex;align-items:center;gap:5px;border-radius:999px;background:#eaf7ff;color:#0878a5;padding:5px 9px;font-size:11px;font-weight:700;white-space:nowrap}}
    .btn-view{{display:inline-flex;align-items:center;justify-content:center;min-width:52px;padding:6px 12px;border:1px solid #8abaff;border-radius:7px;background:#fff;color:#0869c4;text-decoration:none;font-size:12px;font-weight:700}}
    .btn-view:hover{{background:#edf6ff}}
    .pill{{border-radius:999px;font-size:11px;padding:5px 9px}}
    .sev-low{{background:#dcf8f0;color:#07866d}}
    .sev-medium{{background:#fff2dc;color:#a65b00}}
    .sev-high{{background:#ffe7e8;color:#c62e3b}}
    .st-new{{background:#e3efff;color:#2467bd}}
    .st-progress{{background:#f0e9ff;color:#7139d5}}
    .st-closed{{background:#d9f8eb;color:#087e5c}}
    .empty-state{{display:flex;align-items:center;gap:10px;color:#6b829d;padding:22px}}
    .empty-icon{{display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:9px;background:#eaf4ff;color:#1479d0;font-size:20px}}
    .empty-icon svg{{display:block}}
    .welcome-calendar-main{{margin:0;flex:0 0 36px;width:36px;height:36px}}.welcome-datetime{{display:flex;align-items:center;justify-content:flex-end;gap:16px}}.welcome-date-item{{display:flex;align-items:center;gap:10px;min-height:38px}}.welcome-date-control{{display:flex;align-items:center;min-height:36px}}.welcome-date-text,.welcome-time-text{{display:inline-flex;align-items:center;min-height:32px;font-variant-numeric:tabular-nums;line-height:1.3}}.welcome-time-text{{font-size:17px;font-weight:800;letter-spacing:.035em}}.welcome-date-text{{font-size:14px;font-weight:750}}.welcome-date-item>svg{{display:block;flex:0 0 20px}}
    .welcome-calendar-main:focus-visible{{outline:3px solid #93c5fd;outline-offset:2px}}
    .inline-ticket-edit summary:before{{content:none}}
    .ticket-heading-edit{{display:flex;align-items:center;gap:7px;flex-wrap:wrap}}
    .ticket-delete-inline{{display:inline-flex;margin:0;align-items:center}}
    .ticket-delete-icon{{display:inline-flex;align-items:center;justify-content:center;width:27px;height:27px;padding:0;border:1px solid #f4c8ce;border-radius:7px;background:#fff4f5;color:#c53d4b;box-shadow:0 1px 3px rgba(25,70,110,.06);cursor:pointer}}
    .ticket-delete-icon:hover{{background:#ffe5e8;border-color:#e99aa5}}
    .detail-back-row{{display:flex;justify-content:flex-start;margin:14px 0 4px}}
    .icon-back-button{{display:inline-flex;align-items:center;justify-content:center;width:30px;height:30px;border-radius:50%;background:#f3f8fe;color:#3976ad;border:1px solid #dce9f6;box-shadow:none;text-decoration:none;transition:background .18s,transform .18s}}
    .icon-back-button:hover{{background:#dceeff;transform:translateX(-2px)}}
    .icon-back-button:focus-visible{{outline:3px solid #93c5fd;outline-offset:3px}}
    .detail-field{{padding:13px 14px}}
    .detail-field strong{{color:#36577f}}
    .detail-field-issue,.detail-field-note,.detail-field-handler,.detail-field-received_at,.detail-field-closed_at{{border-radius:9px}}
    .workflow-panel{{background:#fff;border-color:#e3edf8}}
    input,select,textarea{{border-color:#d8e5f3}}
    button,.btn{{border-radius:8px}}
    .user-filter-grid{{grid-template-columns:repeat(5,minmax(112px,1fr)) minmax(150px,1.2fr) auto;gap:8px;align-items:end;min-width:940px}}
    .user-filter-grid .filter-actions{{display:flex;gap:6px;align-items:center;white-space:nowrap}}
    .inline-ticket-edit{{position:relative;display:inline-block;margin-top:5px;vertical-align:top}}
    .inline-ticket-edit summary{{list-style:none;display:inline-flex;align-items:center;justify-content:center;width:25px;height:25px;border:1px solid #c8dced;border-radius:7px;background:#f3f8fe;color:#1769aa;cursor:pointer;font-size:0;box-shadow:0 1px 3px rgba(25,70,110,.08)}}
    .inline-ticket-edit summary::-webkit-details-marker{{display:none}}
    .inline-ticket-edit summary:before{{content:none}}.inline-ticket-edit summary svg{{display:block;width:15px;height:15px}}
    .inline-ticket-edit[open] summary{{background:#e5f2ff;border-color:#8dbce8}}
    .inline-ticket-edit form{{position:absolute;z-index:20;top:31px;left:0;width:min(520px,calc(100vw - 64px));max-height:70vh;overflow:auto;padding:12px;background:#fff;border:1px solid #d7e5f4;border-radius:11px;margin:0;box-shadow:0 12px 32px rgba(28,66,108,.18)}}
    .inline-ticket-edit .grid{{grid-template-columns:repeat(2,minmax(0,1fr));gap:7px}}
    .inline-ticket-edit label{{margin:4px 0;font-size:11px}}
    .inline-ticket-edit input,.inline-ticket-edit select,.inline-ticket-edit textarea{{font-size:12px;min-height:32px;padding:7px}}
    .ticket-heading-copy{{min-width:0}}.pill.sev-high,.priority-pill.sev-high{{background:#ffe7e8!important;color:#c62e3b!important;border:1px solid #ffc9ce!important}}.pill.sev-medium,.priority-pill.sev-medium{{background:#fff2dc!important;color:#945b00!important;border:1px solid #f4d49b!important}}.pill.sev-low,.priority-pill.sev-low{{background:#dcf8f0!important;color:#087e68!important;border:1px solid #b7eadc!important}}.ticket-table .priority-cell .sev-high{{background:#ffe7e8!important;color:#c62e3b!important}}.ticket-table .priority-cell .sev-medium{{background:#fff2dc!important;color:#945b00!important}}.ticket-table .priority-cell .sev-low{{background:#dcf8f0!important;color:#087e68!important}}
    .ticket-heading-edit{{display:flex;align-items:center;gap:8px;margin-top:7px;color:#31577f;font-size:13px;font-weight:750;letter-spacing:.01em}}.ticket-heading-edit>span{{font-weight:800;color:#174a7e}}.inline-ticket-edit{{margin-top:0}}.inline-ticket-edit summary{{width:30px;height:30px;border-radius:9px}}.ticket-delete-icon{{width:30px;height:30px;border-radius:9px}}.ticket-heading-edit .ticket-delete-inline{{margin-left:0}}
    
    


    /* Final visual pass: align the entire application with the supplied FMP service-desk reference. */
    :root{{color-scheme:light;--navy:#123d70;--blue:#0878d1;--line:#e4edf7;--muted:#6f86a7}}
    body{{font-family:Inter,"Segoe UI",Arial,sans-serif;background:#f2f7fc;color:#193b67}}
    .topbar{{position:fixed;z-index:20;inset:0 0 auto;height:68px;box-sizing:border-box;padding:0 28px;background:rgba(255,255,255,.97);border-bottom:1px solid #e0ebf7;box-shadow:0 3px 14px rgba(30,75,120,.05);flex-wrap:nowrap}}
    .brand{{display:flex;align-items:center;gap:13px;min-width:0}}.brand-mark{{display:inline-flex;align-items:center;justify-content:center;width:42px;height:42px;flex:0 0 42px;border-radius:50%;background:linear-gradient(145deg,#0789dd,#35c5ef);color:#fff;font-size:27px;font-weight:800;box-shadow:0 3px 9px #168bd329}}.brand-mark svg{{display:block;width:100%;height:100%;padding:2px;box-sizing:border-box}}
    .brand strong{{font-size:20px;letter-spacing:.015em;color:#123d70;white-space:nowrap}}.brand-divider,.top-divider{{height:27px;border-left:1px solid #d5e2f0;margin:0 7px}}.brand-credit{{font-size:12px;color:#6f8ab2;white-space:nowrap}}
    .topbar-right{{display:flex;align-items:center;gap:12px;color:#244b7c}}.clinic-mark{{font-size:28px;color:#0996d4;line-height:1}}.clinic-copy{{display:flex;flex-direction:column;gap:2px;white-space:nowrap}}.clinic-copy strong{{font-size:12px;font-weight:650}}.clinic-copy small{{font-size:10px;color:#7b94b8}}
    .user-mark{{display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:50%;background:#e3f1ff;color:#0875c9;font-size:20px}}.user-name{{font-size:12px;max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}.lang-switch{{padding:6px 9px;border:1px solid #d8e7f7;border-radius:8px;background:#f8fbff;color:#1b65a5!important;font-weight:700;text-decoration:none}}
    .sidebar{{position:fixed;z-index:10;top:68px;bottom:0;left:0;width:230px;box-sizing:border-box;padding:20px 11px 18px;background:#f9fcff;border-right:1px solid #e3edf8;display:flex;flex-direction:column;justify-content:space-between}}.sidebar nav{{display:flex;flex-direction:column;gap:7px}}.side-link{{display:flex;align-items:center;gap:13px;min-height:22px;padding:12px 14px;border-radius:10px;text-decoration:none;color:#365780;font-weight:550;transition:background .15s,color .15s}}.side-link:hover{{background:#f0f7ff;color:#086fc5}}.side-link.active{{background:linear-gradient(100deg,#e7f3ff,#eaf5ff);color:#006bd1;font-weight:700}}.side-icon{{display:inline-flex;width:24px;justify-content:center;font-size:21px;line-height:1;color:#426895}}.side-link.active .side-icon{{color:#0878d1}}
    .sidebar-footer{{display:flex;flex-direction:column;align-items:center;gap:8px;text-align:center;color:#8198b7;font-size:12px;padding:16px 5px}}.medical-doodle{{font-size:40px;color:#a5bbd6;line-height:1}}.sidebar-footer small{{font-size:10px;color:#a2b3c9}}
    main.app-main{{max-width:none;margin:88px 22px 24px 252px;padding:0;min-width:0}}main.login-main{{max-width:600px;margin:110px auto 30px;padding:0 18px}}.box{{background:rgba(255,255,255,.97);padding:20px;border:1px solid #e2edf8;border-radius:13px;margin:12px 0;overflow-x:auto;box-shadow:0 3px 12px rgba(35,82,130,.035)}}.box h2{{color:#123d70;font-size:18px;margin:0 0 16px}}
    input,select,textarea{{border-color:#d8e5f3;background:#fff;color:#203e64;border-radius:8px;min-height:38px}}input:focus,select:focus,textarea:focus{{outline:2px solid #bfdefb;border-color:#70b6f3}}label{{color:#385779;font-size:12px;font-weight:700}}button,.btn{{background:linear-gradient(135deg,#0788d1,#1266b2);box-shadow:0 2px 5px rgba(8,111,190,.12);border-radius:8px;font-weight:650}}
    .ticket-heading-card{{display:flex;align-items:center;gap:16px;min-height:90px;box-sizing:border-box;padding:18px 22px;background:#fff;border:1px solid #e2edf8;border-radius:13px;box-shadow:0 3px 12px rgba(35,82,130,.035);margin:12px 0}}.ticket-heading-icon{{display:flex;align-items:center;justify-content:center;flex:0 0 52px;width:52px;height:52px;border-radius:15px;background:#e6f2ff;color:#0878d1;font-size:28px}}.ticket-heading-copy{{flex:1;min-width:0}}.ticket-heading-line{{display:flex;align-items:center;gap:12px;flex-wrap:wrap}}.ticket-heading-line h2{{margin:0;color:#123d70;font-size:22px}}.ticket-heading-copy p{{margin:6px 0 0;color:#7890b1;font-size:13px}}.ticket-heading-actions{{display:flex;align-items:center;gap:12px}}.priority-pill{{display:inline-flex;align-items:center;gap:7px;padding:8px 12px;border-radius:999px;font-size:12px;font-weight:750;white-space:nowrap}}.more-button{{display:flex;align-items:center;justify-content:center;width:34px;height:34px;border:1px solid #e5edf7;border-radius:10px;text-decoration:none;color:#244b7c;font-size:22px;background:#fff}}.detail-box{{padding:16px 20px 20px}}.detail-grid{{grid-template-columns:repeat(4,minmax(0,1fr));gap:0;border:0;border-radius:0;padding:0}}.detail-field{{min-width:0;min-height:70px;padding:12px 15px;border:0;border-right:1px solid #e6eef8;border-bottom:1px solid #e6eef8;border-radius:0;background:transparent;box-shadow:none}}.detail-field:nth-child(4n){{border-right:0}}.detail-field strong{{display:block;margin-bottom:7px;font-size:12px;color:#36577f;line-height:1.35}}.detail-field p{{margin:0;line-height:1.5;overflow-wrap:anywhere;font-size:14px;color:#203e64}}.detail-field-issue,.detail-field-note,.detail-field-handler,.detail-field-received_at,.detail-field-closed_at{{margin:5px;border:1px solid #cfe3ff;border-radius:9px;background:#f1f8ff}}.detail-field-issue{{background:#f1f8ff}}.detail-field-handler{{background:#f2f8ff}}.detail-field-received_at{{background:#fffaf0;border-color:#ffdc91}}.detail-field-closed_at{{background:#f0fcf7;border-color:#a9f0d5}}.detail-field-note{{background:#f1f8ff;border-color:#bfdefe}}.detail-field-department,.detail-field-location{{display:flex;flex-direction:column;align-items:flex-start;justify-content:center}}.detail-tag,.reporter-chip,.handler-chip{{display:inline-flex;align-items:center;gap:6px;padding:5px 9px;border-radius:999px;font-size:12px;font-weight:700;max-width:100%;overflow-wrap:anywhere}}.detail-tag-dept,.dept-chip{{background:#e7f8ff;color:#087eaa}}.detail-tag-location{{background:#f0e9ff;color:#7139d5}}.reporter-chip{{border:1px solid #b8e7fb;background:#eafaff;color:#087da5}}.handler-chip{{background:#eaf7ff;color:#0878a5}}.detail-time{{display:inline-flex;align-items:center;gap:6px;padding:5px 8px;border-radius:999px;font-size:12px;font-weight:700;line-height:1.35;overflow-wrap:anywhere}}.detail-time-progress{{background:#fff2dc;color:#a65b00;border:1px solid #ffdc91}}.detail-time-closed{{background:#dcf8f0;color:#07866d;border:1px solid #a7f3d0}}.detail-time-pending{{background:#eaf3fa;color:#52708f;border:0;font-size:12px;font-weight:500}}
    .ticket-table{{width:100%;min-width:960px;border:0;border-radius:0;border-spacing:0}}.table-scroll{{width:100%;overflow-x:auto;border:1px solid #e7eef7;border-radius:10px}}.ticket-table thead th{{white-space:nowrap;background:#f5f8fc;color:#38577d;font-size:11px;font-weight:700;text-transform:none;letter-spacing:0;padding:12px 10px;border-bottom:1px solid #e4edf7}}.ticket-table tbody td{{padding:10px;font-size:12px;vertical-align:middle;background:#fff;border-bottom:1px solid #edf2f8;border-right:0}}.ticket-table tbody tr:hover td{{background:#f8fbff}}.ticket-table tbody tr:last-child td{{border-bottom:0}}.ticket-code{{color:#103d78;font-weight:750}}.row-number{{color:#607995;width:24px}}.device-cell{{display:inline-flex;align-items:center;gap:8px;white-space:nowrap}}.mini-icon{{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:6px;background:#eef6ff;color:#1269bd;font-size:17px}}.issue-cell{{max-width:180px;min-width:90px;overflow-wrap:anywhere}}.date-cell{{white-space:nowrap;font-variant-numeric:tabular-nums}}.btn-view{{display:inline-flex;align-items:center;justify-content:center;min-width:52px;padding:6px 12px;border:1px solid #8abaff;border-radius:7px;background:#fff;color:#0869c4;text-decoration:none;font-size:12px;font-weight:700;box-shadow:none}}.btn-view:hover{{background:#edf6ff}}.pill{{border-radius:999px;font-size:11px;padding:5px 9px}}.sev-low{{background:#dcf8f0;color:#07866d}}.sev-medium{{background:#fff2dc;color:#a65b00}}.sev-high{{background:#ffe7e8;color:#c62e3b}}.st-new{{background:#e3efff;color:#2467bd}}.st-progress{{background:#f0e9ff;color:#7139d5}}.st-closed{{background:#d9f8eb;color:#087e5c}}
    .calendar-toolbar{{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}}.quick-filter{{border:1px solid #d8e7f6;background:#f8fbff;color:#1e5d9b;padding:8px 11px;border-radius:8px;text-decoration:none;font-size:12px;font-weight:650}}.quick-filter-hot{{background:#fff0f1;border-color:#ffd1d5;color:#bd3443}}.filter-grid{{gap:12px}}.filter-grid button,.filter-grid .btn{{min-height:38px}}.filter-hint{{font-size:11px;color:#8aa0ba;margin-top:4px}}.user-card{{border-color:#e2edf8;border-radius:12px}}.app-credit{{margin-left:240px;max-width:none;padding:10px 20px;color:#91a5bf}}
    
    
    

    .ticket-list-head{{display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap;margin-bottom:16px}}.ticket-list-head h2{{display:flex;align-items:center;gap:10px;margin:0;color:#123d70;font-size:18px}}.section-icon{{display:inline-flex;align-items:center;justify-content:center;width:36px;height:36px;border-radius:11px;background:#e5f2ff;color:#0878d1;font-size:22px}}.ticket-search{{display:none}}.top-notification{{position:relative;display:inline-flex;align-items:center;justify-content:center;width:36px;height:36px;border:1px solid #d9e8f8;border-radius:10px;background:#f4f9ff;color:#1769aa;text-decoration:none;font-size:20px}}.top-notification .notification-count{{position:absolute;top:-6px;right:-5px;min-width:16px;height:16px;padding:0 3px;display:flex;align-items:center;justify-content:center;border-radius:99px;background:#e5484d;color:#fff;font-size:10px;font-weight:800;line-height:1}}.calendar-pick-wrap{{display:flex;gap:6px;align-items:center}}.calendar-pick-wrap input[type=text]{{flex:1;min-width:0}}.calendar-pick-button{{flex:0 0 38px;width:38px;height:38px;padding:0!important;display:inline-flex;align-items:center;justify-content:center;background:#eaf4ff!important;color:#1769aa!important;border:1px solid #cfe3fa!important}}


    /* Final compact layout pass aligned with the Modern Medical IT reference */
    body{{overflow-x:hidden;background:#f2f7fc}}
    .topbar{{height:68px;min-height:68px;padding:0 26px;flex-wrap:nowrap;gap:14px}}
    .brand{{gap:12px;min-width:0}}.brand strong{{font-size:20px}}
    .clinic-mark{{display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:50%;background:#e3f6ff;color:#0788bf;font-size:25px}}
    .role-mark{{flex:0 0 34px;width:34px;height:34px;border-radius:50%;font-size:18px}}
    .role-mark.role-admin{{background:#e3efff;color:#075fc2;border:1px solid #b7d7ff;box-shadow:inset 0 0 0 3px #f5faff}}
    .role-mark.role-user{{background:#e3f1ff;color:#0875c9;border:1px solid #c7e1fa;box-shadow:inset 0 0 0 3px #f5faff}}
    /* V13: align User banner controls on the same center line as Admin on desktop. */
    @media (min-width: 769px){{
      .topbar{{display:flex;align-items:center;gap:14px}}
      .topbar .brand,.topbar .topbar-right{{display:flex;align-items:center;min-height:68px}}
      .topbar .topbar-right{{justify-content:flex-end;gap:9px;flex-wrap:nowrap}}
      .topbar .topbar-right>*{{align-self:center;flex-shrink:0}}
      .topbar .user-mark,.topbar .role-mark{{display:inline-flex;align-items:center;justify-content:center;box-sizing:border-box}}
      .topbar .user-name{{display:inline-flex;align-items:center;line-height:1.2}}
      .topbar .top-notification,.topbar .integration-status-link{{align-self:center}}
    }}

    .sidebar{{overflow:hidden;background:#f9fcff}}
    main.app-main{{margin-top:86px;margin-bottom:18px;min-width:0}}
    .box{{overflow:visible;min-width:0;padding:16px 18px;margin:0 0 12px;border-radius:12px;border-color:#e4edf7;box-shadow:0 3px 12px rgba(35,82,130,.035)}}
    .ticket-toolbar-panel{{padding:15px 16px 14px}}.ticket-table-panel{{padding:0;overflow:hidden}}
    .ticket-list-head{{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:nowrap;margin-bottom:12px}}
    .ticket-list-head h2{{display:flex;align-items:center;gap:10px;margin:0;font-size:18px;white-space:nowrap}}
    .section-icon{{display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:10px;background:#e5f2ff;color:#0878d1;font-size:21px;flex:0 0 auto}}
    .toolbar-actions{{display:flex;align-items:center;justify-content:flex-end;gap:8px;min-width:0;flex:1}}
    .ticket-search{{display:none}}
    .btn-export,.btn-reset{{display:inline-flex;align-items:center;justify-content:center;gap:6px;height:36px;box-sizing:border-box;padding:0 12px;white-space:nowrap;background:#fff;color:#1767b2;border:1px solid #cfe1f5;border-radius:8px;font-size:12px;font-weight:650;text-decoration:none}}
    .quick-toolbar{{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:0 0 12px}}
    .quick-filter{{display:inline-flex;align-items:center;gap:6px;min-height:33px;box-sizing:border-box;padding:0 11px;border:1px solid #d8e7f6;background:#f8fbff;color:#1e5d9b;border-radius:8px;text-decoration:none;font-size:12px;font-weight:650}}
    .quick-filter-hot{{background:#fff0f1;border-color:#ffd1d5;color:#bd3443}}
    .admin-quick-toolbar{{flex-wrap:nowrap;gap:8px;overflow-x:auto;overflow-y:hidden;padding:1px 1px 4px;scrollbar-width:thin}}
    .admin-quick-toolbar .quick-filter{{flex:0 0 auto;white-space:nowrap;font-family:inherit;cursor:pointer}}
    .admin-quick-toolbar .refresh-quick{{appearance:none;-webkit-appearance:none}}
    .compact-filter-grid{{display:grid;grid-template-columns:repeat(5,minmax(115px,1fr)) minmax(140px,auto);gap:9px;align-items:end}}
    .compact-filter-grid>div{{min-width:0}}.compact-filter-grid label{{display:block;margin:0 0 5px;font-size:11px;color:#536f90;font-weight:700}}
    .compact-filter-grid select,.compact-filter-grid input{{width:100%;height:36px;min-height:36px;padding:6px 8px;border:1px solid #d8e5f3;border-radius:8px;background:#fff;color:#315477;font-size:12px}}
    .compact-filter-grid .filter-actions{{display:flex;align-items:center;gap:7px;min-width:0}}
    .compact-filter-grid button,.compact-filter-grid .btn{{height:36px;min-height:36px;display:inline-flex;align-items:center;justify-content:center;gap:5px;white-space:nowrap;padding:0 10px;font-size:12px;border-radius:8px}}
    .compact-filter-grid button{{background:#e8f3ff;border:1px solid #c6e0ff;color:#0873d0;box-shadow:none}}
    .filter-hint{{display:none!important}}
    .table-wrap,.table-scroll{{width:100%;max-width:100%;overflow:hidden;border:1px solid #e7eef7;border-radius:10px}}
    .ticket-table{{width:100%;min-width:0!important;table-layout:fixed;border:0;border-collapse:separate;border-spacing:0}}
    .ticket-table thead th{{white-space:normal;background:#f5f8fc;color:#38577d;font-size:11px;font-weight:700;text-transform:none;letter-spacing:0;padding:10px 7px;border-bottom:1px solid #e4edf7;line-height:1.25;overflow-wrap:anywhere}}
    .ticket-table tbody td{{padding:9px 7px;font-size:11px;vertical-align:middle;background:#fff;border-bottom:1px solid #edf2f8;border-right:0;line-height:1.35;overflow-wrap:anywhere}}
    .ticket-table tbody tr:hover td{{background:#f7fbff}}.ticket-table tbody tr:last-child td{{border-bottom:0}}
    .ticket-table th:nth-child(1),.ticket-table td:nth-child(1){{width:3.5%;text-align:center}}
    .ticket-table th:nth-child(2),.ticket-table td:nth-child(2){{width:11%}}
    .ticket-table th:nth-child(3),.ticket-table td:nth-child(3){{width:9%}}
    .ticket-table th:nth-child(4),.ticket-table td:nth-child(4){{width:12%}}
    .ticket-table th:nth-child(5),.ticket-table td:nth-child(5){{width:12%}}
    .ticket-table th:nth-child(6),.ticket-table td:nth-child(6){{width:7%}}
    .ticket-table th:nth-child(7),.ticket-table td:nth-child(7){{width:9%}}
    .ticket-table th:nth-child(8),.ticket-table td:nth-child(8){{width:12%}}
    .ticket-table th:nth-child(9),.ticket-table td:nth-child(9){{width:12%}}
    .ticket-table th:nth-child(10),.ticket-table td:nth-child(10){{width:7.5%}}
    .ticket-code{{white-space:normal;overflow-wrap:anywhere;font-size:11px}}
    .device-cell{{white-space:normal;gap:5px}}.mini-icon{{width:20px;height:20px;flex:0 0 20px;font-size:15px}}
    .reporter-chip,.dept-chip,.handler-chip{{display:inline-flex;align-items:center;justify-content:center;gap:4px;max-width:100%;padding:5px 7px;border-radius:999px;font-size:10px;line-height:1.25;overflow-wrap:anywhere}}
    .ticket-table .pill{{display:inline-flex;align-items:center;justify-content:center;white-space:normal;text-align:center;padding:5px 7px;font-size:10px;line-height:1.2}}
    .btn-view{{min-width:0;width:100%;box-sizing:border-box;padding:6px 5px;font-size:11px}}
    .ticket-heading-card{{display:flex;align-items:center;gap:14px;padding:16px 20px;margin:0 0 12px;background:#fff;border:1px solid #e3edf7;border-radius:12px;box-shadow:0 3px 12px rgba(31,78,126,.035);min-width:0}}
    .ticket-heading-icon{{flex:0 0 48px;width:48px;height:48px;border-radius:15px;background:#e5f2ff;color:#0878d1;font-size:27px}}
    .ticket-heading-copy{{flex:1;min-width:0}}.ticket-heading-line{{display:flex;align-items:center;gap:10px;flex-wrap:wrap}}
    .ticket-heading-line h2{{margin:0;font-size:21px;color:#123d70}}.ticket-heading-copy p{{margin:4px 0 0;font-size:12px;color:#7b91ad}}
    .ticket-heading-actions{{display:flex;align-items:center;gap:8px;flex:0 0 auto}}.more-button{{display:none!important}}
    .priority-pill{{display:inline-flex;align-items:center;gap:6px;padding:7px 11px;border-radius:999px;font-size:12px;font-weight:700;white-space:nowrap}}
    .priority-pill.sev-high{{background:#ffe7e8;color:#c62e3b}}.priority-pill.sev-medium{{background:#fff2dc;color:#a65b00}}.priority-pill.sev-low{{background:#e5f9ff;color:#0780a7}}
    .detail-box{{padding:16px 20px}}.detail-grid{{grid-template-columns:repeat(4,minmax(0,1fr));gap:0;align-items:stretch}}
    .detail-field{{min-width:0;padding:12px 13px;min-height:68px}}.detail-field strong{{font-size:12px;margin-bottom:6px}}.detail-field p{{font-size:14px;line-height:1.4;overflow-wrap:anywhere}}
    .detail-field-issue,.detail-field-note,.detail-field-handler,.detail-field-received_at,.detail-field-closed_at{{margin:5px;border:1px solid #cfe3ff;border-radius:9px;background:#f1f8ff}}
    .detail-field-received_at{{background:#fffaf0;border-color:#ffdc91}}.detail-field-closed_at{{background:#f0fcf7;border-color:#a9f0d5}}
    .create-ticket-box h2{{display:flex;align-items:center;gap:10px;margin-top:0}}.create-ticket-box .grid{{grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}}
    .create-ticket-box input,.create-ticket-box select,.create-ticket-box textarea{{border-color:#d6e4f3;background:#fff;border-radius:8px}}
    
    
    

    
    .pill-icon{{display:inline-flex;align-items:center;justify-content:center;margin-right:4px;font-size:10px;line-height:1}}
    .priority-shield{{font-size:12px}}
    .detail-field strong{{display:flex;align-items:center;gap:7px}}
    .detail-field strong:before{{display:inline-flex;align-items:center;justify-content:center;flex:0 0 25px;width:25px;height:25px;border-radius:8px;background:#eef6ff;color:#28629c;font-size:14px;line-height:1}}
    .detail-field-device strong:before{{content:"▱"}}.detail-field-device_code strong:before{{content:"▣"}}
    .detail-field-dept strong:before{{content:"▦"}}.detail-field-location strong:before{{content:"⌖";background:#f0e9ff;color:#7139d5}}
    .detail-field-issue_type strong:before{{content:"⚠"}}.detail-field-issue strong:before{{content:"▤"}}
    .detail-field-severity strong:before{{content:"▥"}}.detail-field-reporter strong:before{{content:"♙"}}
    .detail-field-phone strong:before{{content:"☎"}}.detail-field-time strong:before{{content:"▦"}}
    .detail-field-handler strong:before{{content:"⚒";background:#e5f2ff;color:#0878d1}}
    .detail-field-received_at strong:before{{content:"◷";background:#fff2dc;color:#a65b00}}
    .detail-field-closed_at strong:before{{content:"✓";background:#dcf8f0;color:#07866d}}
    .detail-field-note strong:before{{content:"✎";background:#e5f2ff;color:#0878d1}}


    /* Final detail layout: precise, balanced, and consistent with the reference image */
    .ticket-heading-card{{display:flex;align-items:center;gap:14px;padding:16px 20px;margin:0 0 12px;background:#fff;border:1px solid #e3edf7;border-radius:12px;box-shadow:0 3px 12px rgba(31,78,126,.035);min-width:0}}
    .ticket-heading-icon{{display:flex;align-items:center;justify-content:center;flex:0 0 48px;width:48px;height:48px;border-radius:15px;background:#e5f2ff;color:#0878d1;font-size:0}}
    .ticket-heading-icon svg{{display:block}}
    .ticket-heading-copy{{flex:1;min-width:0}}.ticket-heading-line{{display:flex;align-items:center;gap:10px;flex-wrap:wrap}}
    .ticket-heading-line h2{{margin:0;font-size:21px;color:#123d70}}.ticket-heading-copy p{{margin:4px 0 0;font-size:12px;color:#7b91ad}}
    .ticket-heading-actions{{display:flex;align-items:center;gap:8px;flex:0 0 auto}}.more-button{{display:none!important}}
    .priority-pill{{display:inline-flex;align-items:center;gap:6px;padding:7px 11px;border-radius:999px;font-size:12px;font-weight:700;white-space:nowrap}}
    .priority-pill.sev-high{{background:#ffe7e8;color:#c62e3b}}.priority-pill.sev-medium{{background:#fff2dc;color:#a65b00}}.priority-pill.sev-low{{background:#e5f9ff;color:#0780a7}}
    .detail-box{{padding:10px 20px 16px;overflow:visible}}
    .detail-grid{{display:grid!important;grid-template-columns:repeat(4,minmax(0,1fr))!important;gap:0!important;align-items:stretch;border:0;border-radius:0;padding:0}}
    .detail-field{{display:flex;flex-direction:column;justify-content:flex-start;min-width:0;min-height:78px;padding:12px 14px;border:0;border-right:1px solid #e6eef8;border-bottom:1px solid #e6eef8;border-radius:0;background:transparent;box-shadow:none;margin:0;text-align:left!important}}
    .detail-field:nth-child(4n){{border-right:0}}
    .detail-field strong:before{{content:none!important;display:none!important}}
    .detail-field strong{{display:flex!important;align-items:center;justify-content:flex-start!important;gap:9px;min-height:28px;margin:0 0 6px;font-size:12px;font-weight:700;color:#36577f;line-height:1.35;text-align:left!important;width:auto!important}}
    .detail-icon{{display:inline-flex;align-items:center;justify-content:center;flex:0 0 28px;width:28px;height:28px;border-radius:9px;background:#f0f6fd;color:#315b8c}}
    .detail-icon svg,.detail-tag svg,.reporter-chip svg,.detail-time svg{{display:block;flex:0 0 auto}}.detail-tag svg,.reporter-chip svg,.detail-time svg{{display:block;flex:0 0 auto}}
    .detail-field p{{display:block!important;min-width:0;margin:0 0 0 37px!important;font-size:14px;line-height:1.4;color:#203e64;overflow-wrap:anywhere;white-space:normal;min-height:21px}}
    .detail-field-department .detail-icon{{background:#e6f8ff;color:#0784aa}}.detail-field-location .detail-icon{{background:#f0e9ff;color:#7139d5}}
    .detail-field-issue .detail-icon{{background:#e6f2ff;color:#0878d1}}.detail-field-severity .detail-icon{{background:#eff6ff;color:#315b8c}}
    .detail-field-department,.detail-field-location{{align-items:stretch!important;justify-content:flex-start!important;background:transparent!important}}
    .detail-field-department p,.detail-field-location p{{display:block!important;justify-content:flex-start!important;width:auto!important;margin-left:37px!important;min-height:21px}}
    .detail-field-issue{{grid-column:auto!important;background:#f1f8ff;border:1px solid #cfe3ff;border-radius:9px;margin:5px;padding:8px 11px;min-height:66px}}
    .detail-field-issue p{{font-size:14px;color:#203e64;white-space:pre-wrap}}
    .detail-field-handler{{grid-column:auto!important;background:#f1f8ff;border:1px solid #cfe3ff;border-radius:9px;margin:5px;padding:8px 11px;min-height:66px}}
    .detail-field-handler strong{{color:#075bb5}}.detail-field-handler p{{margin-left:37px!important}}
    .detail-field-handler .handler-value{{display:inline-flex;align-items:center;gap:7px;max-width:100%;padding:5px 9px;border-radius:999px;background:#eaf4ff;border:1px solid #b9d9ff;color:#075bb5;font-size:12px;font-weight:700;overflow-wrap:anywhere}}
    .detail-field-received_at{{grid-column:auto!important;background:#fffaf0;border:1px solid #ffdc91;border-radius:9px;margin:5px;padding:8px 11px;min-height:66px}}
    .detail-field-received_at strong,.detail-field-received_at .detail-icon{{color:#b85a00}}.detail-field-received_at .detail-icon{{background:#fff2dc}}
    .detail-field-received_at p{{margin-left:37px!important}}
    .detail-field-closed_at{{grid-column:1 / -1!important;display:grid!important;grid-template-columns:minmax(150px,220px) minmax(0,1fr);align-items:center;gap:12px;min-height:64px;margin:7px 0 5px;padding:10px 13px;border:1px solid #a9f0d5;border-radius:9px;background:#f0fcf7}}
    .detail-field-closed_at strong{{margin:0;color:#07866d}}.detail-field-closed_at .detail-icon{{background:#d8f8eb;color:#07866d}}.detail-field-closed_at p{{margin:0!important}}
    .detail-field-note{{grid-column:1 / -1!important;display:grid!important;grid-template-columns:minmax(150px,220px) minmax(0,1fr);align-items:start;gap:12px;min-height:72px;margin:7px 0 4px;padding:10px 13px;border:1px solid #bfdefe;border-radius:9px;background:#f1f8ff}}
    .detail-field-note strong{{margin:0;color:#0878d1}}.detail-field-note .detail-icon{{background:#e0efff;color:#0878d1}}
    .detail-field-note p{{margin:4px 0 0!important;white-space:pre-wrap;color:#203e64}}
    .detail-field-closed_at .detail-time,.detail-field-received_at .detail-time{{font-size:12px;line-height:1.3;padding:5px 9px}}
    
    

    
    /* Final requested refinements */
    .brand-credit{{display:none!important}}
    .date-text-input{{min-height:38px;letter-spacing:.01em;font-variant-numeric:tabular-nums}}
    .date-text-input::placeholder{{color:#8aa0ba;opacity:1}}
    .mini-icon svg{{display:block}}
    .ticket-table thead th,.ticket-table tbody td{{box-sizing:border-box}}
    /* Align Priority / Status headings and their icons consistently in both languages */
    .ticket-table th.priority-heading,.ticket-table th.status-heading{{text-align:center;vertical-align:middle;padding-left:4px;padding-right:4px}}
    .ticket-table td.priority-cell,.ticket-table td.status-cell{{text-align:center;vertical-align:middle;padding-left:4px;padding-right:4px}}
    .ticket-table td.priority-cell .pill,.ticket-table td.status-cell .pill{{display:inline-flex;align-items:center;justify-content:center;gap:4px;min-width:0;max-width:100%;margin:0 auto;white-space:normal;text-align:center;vertical-align:middle}}
    .ticket-table td.priority-cell .pill-icon,.ticket-table td.status-cell .pill-icon{{display:inline-flex;align-items:center;justify-content:center;flex:0 0 12px;width:12px;height:12px;margin:0;font-size:10px;line-height:1}}
    .ticket-table td.priority-cell .pill-label,.ticket-table td.status-cell .pill-label{{display:inline-block;min-width:0;line-height:1.2;text-align:center;overflow-wrap:normal;word-break:normal}}
    .ticket-table thead th:not(:last-child),.ticket-table tbody td:not(:last-child){{border-right:1px solid #e5edf6}}
    .ticket-table tbody tr:not(:last-child) td{{border-bottom:1px solid #e3ecf6}}
    .detail-field-closed_at{{grid-column:span 1!important;display:flex!important;flex-direction:column!important;justify-content:flex-start!important;align-items:stretch!important;gap:4px!important;min-height:66px;margin:5px!important;padding:8px 11px!important}}
    .detail-field-closed_at strong{{display:flex!important;align-items:center!important;margin:0 0 2px!important}}
    .detail-field-closed_at p{{margin:0 0 0 37px!important;min-height:21px}}
    .detail-field-note{{grid-column:span 3!important;display:flex!important;flex-direction:column!important;justify-content:flex-start!important;align-items:stretch!important;gap:4px!important;min-height:66px;margin:5px!important;padding:8px 11px!important}}
    .detail-field-note strong{{display:flex!important;align-items:center!important;margin:0 0 2px!important}}
    .detail-field-note p{{margin:0 0 0 37px!important;min-height:21px}}
    
    

    
    /* Compact desktop ticket table: keep all columns visible and grid lines crisp */
    .ticket-table{{width:100%;min-width:0!important;table-layout:fixed;border-collapse:collapse;border:1px solid #dbe6f2}}
    .ticket-table th,.ticket-table td{{box-sizing:border-box;overflow-wrap:anywhere}}
    .ticket-table thead th{{padding:8px 5px!important;font-size:10px!important;line-height:1.2;text-align:center;border-right:1px solid #dbe6f2;border-bottom:1px solid #dbe6f2}}
    .ticket-table tbody td{{padding:7px 5px!important;font-size:11px!important;line-height:1.25;border-right:1px solid #e1eaf4!important;border-bottom:1px solid #e1eaf4!important;vertical-align:middle}}
    .ticket-table tbody tr:last-child td{{border-bottom:1px solid #dbe6f2!important}}
    .ticket-table th:last-child,.ticket-table td:last-child{{border-right:0!important}}
    .ticket-table th:nth-child(1),.ticket-table td:nth-child(1){{width:3.5%;text-align:center}}
    .ticket-table th:nth-child(2),.ticket-table td:nth-child(2){{width:10%}}
    .ticket-table th:nth-child(3),.ticket-table td:nth-child(3){{width:10%}}
    .ticket-table th:nth-child(4),.ticket-table td:nth-child(4){{width:14%}}
    .ticket-table th:nth-child(5),.ticket-table td:nth-child(5){{width:10%}}
    .ticket-table th:nth-child(6),.ticket-table td:nth-child(6){{width:9%}}
    .ticket-table th:nth-child(7),.ticket-table td:nth-child(7){{width:10%}}
    .ticket-table th:nth-child(8),.ticket-table td:nth-child(8){{width:10%}}
    .ticket-table th:nth-child(9),.ticket-table td:nth-child(9){{width:14%}}
    .ticket-table th:nth-child(10),.ticket-table td:nth-child(10){{width:6.5%;text-align:center}}
    .ticket-table td.priority-cell .pill,.ticket-table td.status-cell .pill{{padding:3px 4px!important;font-size:9px!important;gap:3px;max-width:100%;white-space:normal}}
    .ticket-table td.priority-cell .pill-icon,.ticket-table td.status-cell .pill-icon{{width:9px;height:9px;flex-basis:9px;font-size:8px}}
    /* Ticket details: consistent three-column grid with clear horizontal/vertical separators */
    .detail-grid{{grid-template-columns:repeat(3,minmax(0,1fr))!important;gap:0!important;border:1px solid #dce7f2!important;border-radius:10px;overflow:hidden}}
    .detail-field{{box-sizing:border-box;border:0!important;border-right:1px solid #dce7f2!important;border-bottom:1px solid #dce7f2!important;border-radius:0!important;margin:0!important;min-width:0;padding:12px!important}}
    .detail-field:nth-child(3n){{border-right:0!important}}
    .detail-field-issue,.detail-field-note,.detail-field-handler,.detail-field-received_at,.detail-field-closed_at{{margin:0!important;border-radius:0!important}}
    
    
    .ticket-table td.priority-cell .pill-label,.ticket-table td.status-cell .pill-label{{overflow-wrap:anywhere}}
    

    /* Refined login card: shared credentials and two compact sign-in tabs */
    main.login-main{{max-width:480px!important;margin:70px auto 30px!important;padding:0 16px!important}}
    .login-card{{padding:26px!important;border:1px solid #dce8f5!important;border-radius:18px!important;box-shadow:0 16px 42px rgba(24,75,125,.10)!important;overflow:hidden!important}}
    .login-title-mark{{display:flex;align-items:center;justify-content:center;width:42px;height:42px;border-radius:13px;background:linear-gradient(135deg,#e1f4ff,#e8efff);color:#0878bd;font-size:25px;margin-bottom:12px}}
    .login-card h2{{font-size:23px!important;margin-bottom:8px!important}}
    .login-tabs{{display:grid;grid-template-columns:1fr 1fr;gap:5px;padding:5px;background:#eef4fa;border:1px solid #e1eaf4;border-radius:11px;margin:18px 0}}
    .login-tab{{display:flex;align-items:center;justify-content:center;min-height:40px;padding:8px 9px;border:1px solid transparent;border-radius:8px;background:transparent;color:#59718b;font-size:12px;font-weight:700;line-height:1.25;box-shadow:none;text-align:center}}
    .login-tab.active{{background:#fff;color:#0877bd;border-color:#d4e8fa;box-shadow:0 2px 7px rgba(25,73,120,.08)}}
    .login-card label{{margin:12px 0 6px;font-size:12px;color:#42617f}}
    .login-input{{height:42px;padding:10px 12px!important;border:1px solid #d5e2ef!important;border-radius:9px!important;background:#fbfdff!important;font-size:13px!important;outline:none}}
    .login-input:focus{{border-color:#6ab5eb!important;box-shadow:0 0 0 3px rgba(8,127,194,.10)}}
    .login-submit-row{{margin:18px 0 8px}}
    .login-submit{{width:100%;min-height:42px;border-radius:9px;font-weight:700;font-size:13px;box-shadow:0 5px 12px rgba(8,111,177,.17)}}
    .login-notice{{padding:10px 12px;margin:15px 0 0;border:1px solid #dbeafe;border-radius:9px;background:#f4f9ff;color:#526d89!important;font-size:11px!important;line-height:1.5}}
    .dept-chip{{display:inline-flex!important;align-items:center!important;gap:5px!important;white-space:normal!important}}
    .dept-chip svg{{flex:0 0 13px;display:block}}.dept-chip span{{min-width:0}}
    /* Compact 4-column desktop grid with consistent separators */
    .detail-grid{{grid-template-columns:repeat(4,minmax(0,1fr))!important;grid-auto-flow:row!important;gap:0!important;align-items:stretch!important;border-top:1px solid #dce7f2;border-left:1px solid #dce7f2}}
    .detail-field{{box-sizing:border-box!important;min-width:0!important;border-right:1px solid #dce7f2!important;border-bottom:1px solid #dce7f2!important;padding:10px!important;display:flex!important;flex-direction:column!important;align-items:stretch!important;gap:7px!important;background:#fff}}
    .detail-field strong{{display:flex!important;align-items:center!important;gap:7px!important;line-height:1.25!important;min-height:26px!important;font-size:12px!important}}
    .detail-field p{{margin:0!important;min-width:0!important;overflow-wrap:anywhere!important;line-height:1.4!important;font-size:12px!important}}
    .detail-field-closed_at{{grid-column:span 1!important;display:flex!important;flex-direction:column!important;align-items:stretch!important;justify-content:flex-start!important;min-width:0!important;margin:0!important;padding:12px!important;gap:5px!important}}
    .detail-field-closed_at strong,.detail-field-note strong{{display:flex!important;align-items:center!important;gap:8px!important;margin:0!important;min-width:0}}
    .detail-field-closed_at p{{margin:0 0 0 31px!important;min-width:0;overflow-wrap:anywhere}}
    .detail-field-note{{grid-column:span 3!important;display:flex!important;flex-direction:column!important;align-items:stretch!important;gap:5px!important;min-width:0!important;min-height:76px;margin:0!important;padding:12px!important}}
    .detail-field-note p{{margin:0 0 0 31px!important;min-width:0;white-space:pre-wrap;overflow-wrap:anywhere;line-height:1.45}}
    .detail-field .detail-icon{{display:inline-flex;align-items:center;justify-content:center;flex:0 0 26px;width:26px;height:26px}}
    
    


    /* Admin user management: compact, aligned, modern controls */
    .user-card{{padding:16px!important;margin:12px 0!important;border:1px solid #dce8f4!important;border-radius:14px!important;background:linear-gradient(180deg,#fff,#fbfdff)!important;box-shadow:0 4px 14px rgba(30,75,120,.045)!important;overflow:hidden!important}}
    .user-card-head{{display:flex!important;align-items:center!important;justify-content:space-between!important;gap:12px!important;padding-bottom:12px!important;margin-bottom:12px!important;border-bottom:1px solid #e7eef6!important}}
    .user-card-head h3{{margin:0 0 5px!important;font-size:15px!important;color:#173f68!important}}
    .user-meta{{font-size:12px!important;line-height:1.5!important;color:#718096!important;overflow-wrap:anywhere}}
    .user-card form{{min-width:0}}
    .user-card .user-edit-grid{{grid-template-columns:repeat(4,minmax(0,1fr))!important;gap:10px!important;align-items:end!important}}
    .user-card label{{margin:5px 0 5px!important;font-size:11px!important;color:#60758d!important}}
    .user-card input,.user-card select{{width:100%!important;min-width:0!important;height:36px!important;padding:7px 9px!important;border:1px solid #d7e3ef!important;border-radius:8px!important;background:#fff!important;font-size:12px!important}}
    .user-card button{{min-height:34px!important;padding:7px 11px!important;border-radius:8px!important;font-size:11px!important;font-weight:700!important;line-height:1.25!important;white-space:normal!important;box-shadow:0 2px 5px rgba(30,70,110,.08)!important;transition:filter .15s ease,transform .15s ease}}
    .user-card button:hover{{filter:brightness(.96);transform:translateY(-1px)}}
    .user-card .btn-secondary{{background:#eef3f8!important;color:#405a75!important;border:1px solid #d8e3ef!important;box-shadow:none!important}}
    .user-card .btn-passwordless{{background:#e8f7f2!important;color:#087c63!important;border:1px solid #bce9dc!important;box-shadow:none!important}}
    .user-card .passwordless-action{{display:flex!important;justify-content:flex-start!important;margin:0 0 12px!important}}
    .user-card form[action$="/edit"]{{padding:12px!important;background:#f7faff!important;border:1px solid #e4edf7!important;border-radius:10px!important}}
    .user-card form[action$="/edit"] p{{margin:10px 0 0!important}}
    .user-card .user-actions{{display:flex!important;justify-content:flex-start!important;align-items:end!important;margin-top:12px!important;padding-top:12px!important;border-top:1px solid #e7eef6!important}}
    .user-card .user-actions form{{display:grid!important;grid-template-columns:minmax(220px,1fr) auto!important;align-items:end!important;gap:10px!important;width:min(560px,100%)!important}}
    .user-card .user-actions form>div{{min-width:0}}
    .user-card .user-actions .help{{font-size:10px!important;margin-top:4px!important}}
    .user-card .user-actions button{{background:#fff1f0!important;color:#b42318!important;border:1px solid #ffd0cb!important;box-shadow:none!important}}
    .box>form[action="/users"]{{padding:16px!important;border:1px solid #e0eaf5!important;border-radius:12px!important;background:#f8fbff!important}}
    .box>form[action="/users"] .user-edit-grid{{grid-template-columns:repeat(3,minmax(0,1fr))!important;gap:12px!important}}
    .box>form[action="/users"] p{{margin:12px 0 0!important}}
    .box>form[action="/users"] button{{min-height:36px!important;padding:8px 15px!important}}
    
    


    /* Final polish: compact navigation, clear ticket grid, compact forms */
    .sidebar{{width:204px!important;padding:18px 9px 14px!important}}
    main.app-main{{margin-left:222px!important;margin-right:18px!important}}
    .app-credit{{margin-left:204px!important}}
    .side-link{{gap:10px!important;padding:10px 11px!important;font-size:12px!important;border-radius:9px!important}}
    .side-icon{{width:22px!important;font-size:19px!important}}
    .detail-grid{{display:grid!important;grid-template-columns:repeat(4,minmax(0,1fr))!important;gap:0!important;border-top:1px solid #d7e2ee!important;border-left:1px solid #d7e2ee!important;border-radius:10px!important;overflow:hidden!important}}
    .detail-field,.detail-field:nth-child(4n),.detail-field:nth-child(2n){{border-right:1px solid #d7e2ee!important;border-bottom:1px solid #d7e2ee!important;min-width:0!important;margin:0!important;border-radius:0!important;padding:9px 10px!important;box-sizing:border-box!important}}
    .detail-field:nth-child(4n){{border-right:1px solid #d7e2ee!important}}
    .detail-field strong{{display:flex!important;align-items:center!important;gap:7px!important;font-weight:750!important;line-height:1.25!important;font-size:12px!important}}
    .detail-field p{{font-size:12px!important;line-height:1.4!important;overflow-wrap:anywhere!important}}
    .detail-field-closed_at{{grid-column:span 1!important}}
    .detail-field-note{{grid-column:span 3!important}}
    .detail-field-closed_at,.detail-field-note{{border-right:1px solid #d7e2ee!important;border-bottom:1px solid #d7e2ee!important}}
    .pill,.ticket-workflow,.priority-pill,.detail-time{{font-weight:800!important}}
    .create-ticket-box .grid{{grid-template-columns:repeat(3,minmax(0,1fr))!important;gap:10px!important}}
    .create-ticket-box label{{margin:7px 0 4px!important;font-size:11px!important}}
    .create-ticket-box input,.create-ticket-box select,.create-ticket-box textarea{{padding:8px 9px!important;font-size:12px!important;min-height:34px!important;border-radius:8px!important}}
    .create-ticket-box textarea{{min-height:76px!important}}
    .user-card .account-action-row{{display:flex!important;align-items:center!important;gap:8px!important;flex-wrap:wrap!important;margin:0 0 10px!important}}
    .user-card .account-action-row form{{margin:0!important;display:flex!important;align-items:center!important}}
    .user-card .account-action-row button{{min-height:32px!important;padding:6px 10px!important;font-size:11px!important}}
    .user-card .btn-secondary{{background:#f0f4f8!important;color:#405a75!important;border:1px solid #d6e1ed!important}}
    .user-card .btn-passwordless{{background:linear-gradient(135deg,#dff8ef,#e8f2ff)!important;color:#087c63!important;border:1px solid #b7e8d9!important}}
    .user-card form[action$="/edit"] p{{display:flex!important;justify-content:flex-end!important;margin:8px 0 0!important}}
    .user-card form[action$="/edit"] p button{{min-width:76px!important;background:linear-gradient(135deg,#087fc2,#1769aa)!important}}
    
    

    /* Final detail-grid separators and calm clinical status styling */
    .detail-grid{{border:1px solid #d7e2ee!important;border-right:0!important;border-bottom:0!important;border-radius:11px!important;overflow:hidden!important;background:#fff!important}}
    .detail-field,.detail-field:nth-child(2n),.detail-field:nth-child(4n){{border-right:1px solid #d7e2ee!important;border-bottom:1px solid #d7e2ee!important;background:#fff!important}}
    .detail-field:nth-child(4n){{border-right:1px solid #d7e2ee!important}}
    .detail-field strong{{color:#345b7f!important;font-weight:750!important}}
    .detail-field .detail-icon{{color:#3976a8!important;background:#edf6fc!important;border-radius:7px!important}}
    .detail-field-department,.detail-field-location{{background:#fbfdff!important}}
    .detail-field-reporter{{background:#fcfeff!important}}
    .detail-field-received_at{{background:#fffdf7!important}}
    .detail-field-closed_at{{background:#f7fcfa!important}}
    .detail-field-note{{background:#f7fbff!important}}
    .detail-field-issue{{background:#f7fbff!important}}
    .pill{{display:inline-flex!important;align-items:center!important;justify-content:center!important;gap:5px!important;font-weight:800!important;line-height:1.25!important;border:1px solid transparent!important}}
    .st-new{{background:#edf5ff!important;color:#245d99!important;border-color:#d5e7fb!important}}
    .st-progress{{background:#f2efff!important;color:#6650a8!important;border-color:#e2dcff!important}}
    .st-closed{{background:#eaf8f1!important;color:#237653!important;border-color:#ccebdc!important}}
    .detail-time-pending{{background:#f1f5f9!important;color:#667b90!important;border-color:#e1e8ef!important}}
    .workflow-panel{{background:#f8fbfe!important;border-color:#dce8f3!important;box-shadow:none!important}}
    .workflow-card{{background:#fff!important;border-color:#e0eaf3!important;border-radius:10px!important}}
    .workflow-card-title{{color:#284f75!important;font-weight:750!important}}
    /* Account forms: consistent vertical flow with action buttons at the bottom */
    .user-card form[action$="/edit"]{{display:flex!important;flex-direction:column!important;align-items:stretch!important;gap:0!important}}
    .user-card form[action$="/edit"] .user-edit-grid{{width:100%!important}}
    .user-card form[action$="/edit"] p{{display:flex!important;justify-content:flex-end!important;align-items:center!important;margin:12px 0 0!important;padding-top:10px!important;border-top:1px solid #e3edf6!important}}
    .user-card form[action$="/edit"] p button{{min-width:100px!important}}
    .user-card .user-actions{{display:block!important;margin-top:12px!important;padding-top:12px!important;border-top:1px solid #e3edf6!important}}
    .user-card .user-actions form{{display:flex!important;flex-direction:column!important;align-items:stretch!important;gap:8px!important;width:100%!important}}
    .user-card .user-actions form>div{{width:100%!important}}
    .user-card .user-actions form button{{align-self:flex-end!important;min-width:150px!important;margin-top:2px!important}}
    .user-card .account-action-row{{margin:0 0 12px!important;padding-bottom:10px!important;border-bottom:1px solid #e7eef6!important}}
    .user-card .account-action-row button{{min-height:32px!important}}
    
    


    /* Unified clinical UI polish: SVG sidebar icons, ticket description chips, compact account controls */
    .side-link{{min-height:42px!important;gap:12px!important;padding:10px 12px!important;border:1px solid transparent!important;color:#3b5878!important;font-size:13px!important;font-weight:650!important}}
    .side-icon{{display:inline-flex!important;align-items:center!important;justify-content:center!important;width:30px!important;height:30px!important;flex:0 0 30px!important;border-radius:9px!important;background:#eef5fc!important;color:#537da8!important}}
    .side-icon svg{{display:block!important;width:18px!important;height:18px!important}}
    .side-link:hover .side-icon{{background:#e0efff!important;color:#176cb0!important}}
    .side-link.active{{background:#eaf4ff!important;border-color:#d9eafb!important;color:#075fa8!important}}
    .side-link.active .side-icon{{background:#d9ecff!important;color:#086bb6!important}}
    .detail-field-issue_type p,.detail-field-issue p{{margin-top:2px!important}}
    .issue-type-chip,.issue-description-chip{{display:inline-flex!important;align-items:center!important;gap:7px!important;max-width:100%!important;box-sizing:border-box!important;padding:6px 10px!important;border-radius:9px!important;font-size:12px!important;line-height:1.45!important;font-weight:700!important;overflow-wrap:anywhere!important;white-space:pre-wrap!important}}
    .issue-type-chip{{background:#f0f5ff!important;border:1px solid #d9e5fb!important;color:#3f6090!important}}
    .issue-type-chip svg{{flex:0 0 auto!important;color:#6485b5!important}}
    .issue-description-chip{{background:#eef8ff!important;border:1px solid #d4eafb!important;color:#285d83!important}}
    .detail-field-issue_type .detail-icon{{background:#edf3ff!important;color:#637fba!important}}
    .detail-field-issue .detail-icon{{background:#e8f5ff!important;color:#397ca8!important}}
    .box>form[action="/users"]{{padding:12px 14px!important;background:#f8fbff!important;border-color:#e2edf7!important}}
    .box>form[action="/users"] .user-edit-grid{{grid-template-columns:repeat(3,minmax(0,1fr))!important;gap:8px 12px!important}}
    .box>form[action="/users"] label{{margin:5px 0 4px!important;font-size:11px!important}}
    .box>form[action="/users"] input,.box>form[action="/users"] select{{height:34px!important;padding:6px 9px!important;font-size:12px!important}}
    .box>form[action="/users"] p{{display:flex!important;justify-content:flex-end!important;margin:9px 0 0!important}}
    .box>form[action="/users"] button{{min-height:33px!important;padding:7px 14px!important;font-size:12px!important}}
    .user-card{{position:relative!important;padding:14px!important;border-color:#e0eaf4!important;border-radius:12px!important}}
    .user-card-head{{padding-right:300px!important;min-height:34px!important;align-items:center!important}}
    .user-card-head h3{{margin:0 0 4px!important;color:#23496f!important;font-size:14px!important}}
    .user-card .account-action-row{{position:absolute!important;top:14px!important;right:14px!important;display:flex!important;justify-content:flex-end!important;align-items:center!important;gap:7px!important;flex-wrap:wrap!important;margin:0!important;padding:0!important;border:0!important;max-width:48%!important}}
    .user-card .account-action-row form{{margin:0!important}}
    .user-card .account-action-row button{{min-height:30px!important;padding:5px 9px!important;font-size:10px!important}}
    .user-card form[action$="/edit"]{{margin-top:12px!important;padding:10px!important;background:#f8fbfe!important;border-color:#e6eef7!important}}
    .user-card form[action$="/edit"] .user-edit-grid{{grid-template-columns:repeat(4,minmax(0,1fr))!important;gap:8px!important}}
    .user-card form[action$="/edit"] p{{margin-top:8px!important;padding-top:8px!important}}
    .user-card .user-actions{{margin-top:9px!important;padding-top:9px!important}}
    .user-card .user-actions form{{grid-template-columns:minmax(180px,1fr) auto!important;width:min(520px,100%)!important}}
    
    

    /* Requested final pass: left-side user actions and borderless, color-coded ticket fields */
    .user-card{{display:grid!important;grid-template-columns:minmax(180px,auto) minmax(0,1fr)!important;grid-template-rows:auto auto auto!important;column-gap:18px!important;row-gap:10px!important;align-items:start!important}}
    .user-card-head{{grid-column:2!important;grid-row:1!important;min-width:0!important;padding:0!important;min-height:34px!important;align-self:center!important}}
    .user-card .account-action-row{{position:static!important;grid-column:1!important;grid-row:1!important;display:flex!important;justify-content:flex-start!important;align-items:center!important;align-self:center!important;gap:7px!important;flex-wrap:wrap!important;max-width:none!important;margin:0!important;padding:0!important;border:0!important}}
    .user-card .account-action-row form{{margin:0!important;display:flex!important;align-items:center!important}}
    .user-card .account-action-row button{{min-height:32px!important;padding:6px 10px!important;font-size:11px!important;white-space:normal!important}}
    .user-card form[action$="/edit"]{{grid-column:1/-1!important;grid-row:2!important;width:100%!important;box-sizing:border-box!important;margin:0!important}}
    .user-card .user-actions{{grid-column:1/-1!important;grid-row:3!important}}
    .detail-grid{{gap:10px!important;padding:0!important;border:0!important;border-radius:0!important;overflow:visible!important;background:transparent!important}}
    .detail-field,.detail-field:nth-child(2n),.detail-field:nth-child(4n),.detail-field-received_at,.detail-field-closed_at,.detail-field-note,.detail-field-issue{{border:0!important;border-radius:12px!important;box-shadow:none!important;margin:0!important;min-width:0!important;padding:13px 14px!important;background:#f5f9fd!important;transition:background .15s ease!important}}
    .detail-field strong{{display:flex!important;align-items:center!important;gap:8px!important;margin-bottom:8px!important;color:#365b7d!important;font-size:12px!important;font-weight:750!important}}
    .detail-field .detail-icon{{display:inline-flex!important;align-items:center!important;justify-content:center!important;flex:0 0 30px!important;width:30px!important;height:30px!important;border-radius:9px!important;background:#e4effa!important;color:#3975a7!important}}
    .detail-field .detail-icon svg{{width:16px!important;height:16px!important}}
    .detail-field p{{margin:0!important;color:#244563!important;font-size:13px!important;line-height:1.5!important;overflow-wrap:anywhere!important}}
    .detail-field-device{{background:#eff7ff!important}}.detail-field-device .detail-icon{{background:#dcecff!important;color:#2868a4!important}}
    .detail-field-device_code{{background:#f2f4ff!important}}.detail-field-device_code .detail-icon{{background:#e2e7ff!important;color:#5a63ac!important}}
    .detail-field-department{{background:#edf9fb!important}}.detail-field-department .detail-icon{{background:#d7f1f4!important;color:#167d8b!important}}
    .detail-field-location{{background:#f3f0ff!important}}.detail-field-location .detail-icon{{background:#e6ddff!important;color:#7354b5!important}}
    .detail-field-issue_type{{background:#fff7eb!important}}.detail-field-issue_type .detail-icon{{background:#ffebc9!important;color:#a96a17!important}}
    .detail-field-issue{{background:#f0f8ff!important}}.detail-field-issue .detail-icon{{background:#dceeff!important;color:#3379b3!important}}
    .detail-field-severity{{background:#fff6f3!important}}.detail-field-severity .detail-icon{{background:#ffe5de!important;color:#b85c49!important}}
    .detail-field-reporter{{background:#effbfa!important}}.detail-field-reporter .detail-icon{{background:#d9f4f0!important;color:#218477!important}}
    .detail-field-phone{{background:#f1f8f2!important}}.detail-field-phone .detail-icon{{background:#dff1e2!important;color:#438451!important}}
    .detail-field-time,.detail-field-created_at{{background:#f0f7ff!important}}.detail-field-time .detail-icon,.detail-field-created_at .detail-icon{{background:#dcecff!important;color:#2868a4!important}}
    .detail-field-handler{{background:#f1f4ff!important}}.detail-field-handler .detail-icon{{background:#e0e6ff!important;color:#5267b1!important}}
    .detail-field-received_at{{background:#fff8e9!important}}.detail-field-received_at .detail-icon{{background:#ffedc5!important;color:#a56a14!important}}
    .detail-field-closed_at{{background:#effaf4!important}}.detail-field-closed_at .detail-icon{{background:#d8f2e3!important;color:#258054!important}}
    .detail-field-note{{background:#f7f4ff!important}}.detail-field-note .detail-icon{{background:#e9e0ff!important;color:#7558b4!important}}
    .detail-tag,.reporter-chip,.handler-value,.issue-type-chip,.issue-description-chip,.detail-time{{border:0!important;box-shadow:none!important}}
    .detail-time-created{{display:inline-flex!important;align-items:center!important;gap:6px!important;padding:5px 8px!important;border-radius:999px!important;background:#e6f2ff!important;color:#2868a4!important;font-size:12px!important;font-weight:700!important;line-height:1.35!important;overflow-wrap:anywhere!important}}
    .detail-time-progress{{background:#fff0cf!important;color:#93600b!important;border:0!important}}
    .detail-time-closed{{background:#dff6e9!important;color:#23774f!important;border:0!important}}
    .detail-time-pending{{background:#edf2f7!important;color:#718096!important;border:0!important}}
    .detail-tag-dept{{background:#dff5f7!important;color:#167d8b!important}}
    .detail-tag-location{{background:#ebe5ff!important;color:#7354b5!important}}
    .reporter-chip{{background:#dff6f2!important;color:#218477!important}}
    .issue-type-chip{{background:#fff0d6!important;color:#95601a!important}}
    .issue-description-chip{{background:#dff0ff!important;color:#285d83!important}}
    .st-new{{background:#e4efff!important;color:#285f9b!important;border:0!important}}
    .st-progress{{background:#eee8ff!important;color:#6650a8!important;border:0!important}}
    .st-closed{{background:#ddf5e7!important;color:#237653!important;border:0!important}}
    
    


    /* Final account layout: compact vertical flow, every action aligned to the left */
    .box>form[action="/users"]{{display:flex!important;flex-direction:column!important;align-items:stretch!important;gap:8px!important}}
    .box>form[action="/users"] .user-edit-grid{{display:grid!important;grid-template-columns:repeat(3,minmax(0,1fr))!important;gap:9px 12px!important;align-items:end!important}}
    .box>form[action="/users"] label{{margin:4px 0!important;font-size:12px!important}}
    .box>form[action="/users"] input,.box>form[action="/users"] select{{height:35px!important;padding:7px 9px!important;font-size:12px!important}}
    .box>form[action="/users"] p{{display:flex!important;justify-content:flex-start!important;align-items:center!important;margin:7px 0 0!important;padding:0!important;border:0!important}}
    .box>form[action="/users"] p button{{align-self:flex-start!important;min-width:125px!important}}
    .user-card{{display:flex!important;flex-direction:column!important;align-items:stretch!important;gap:10px!important;padding:14px!important}}
    .user-card-head{{order:0!important;width:100%!important;min-height:0!important;margin:0!important;padding:0 0 9px!important;border-bottom:1px solid #e7eef6!important}}
    .user-card .account-action-row{{order:1!important;position:static!important;display:flex!important;justify-content:flex-start!important;align-items:center!important;gap:7px!important;flex-wrap:wrap!important;width:100%!important;max-width:none!important;margin:0!important;padding:0!important;border:0!important}}
    .user-card .account-action-row form{{display:flex!important;align-items:center!important;margin:0!important}}
    .user-card .account-action-row button{{min-height:31px!important;padding:6px 10px!important;font-size:11px!important}}
    .user-card form[action$="/edit"]{{order:2!important;display:flex!important;flex-direction:column!important;align-items:stretch!important;width:100%!important;margin:0!important;padding:10px!important;background:#f8fbfe!important;border:1px solid #e5eef7!important;border-radius:10px!important;box-sizing:border-box!important}}
    .user-card form[action$="/edit"] .user-edit-grid{{display:grid!important;grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:8px 10px!important;width:100%!important}}
    .user-card form[action$="/edit"] label{{margin:3px 0 4px!important;font-size:11px!important}}
    .user-card form[action$="/edit"] input,.user-card form[action$="/edit"] select{{width:100%!important;height:34px!important;padding:6px 8px!important;font-size:12px!important}}
    .user-card form[action$="/edit"] p{{display:flex!important;justify-content:flex-start!important;align-items:center!important;width:100%!important;margin:9px 0 0!important;padding:8px 0 0!important;border-top:1px solid #e4edf6!important}}
    .user-card form[action$="/edit"] p button{{align-self:flex-start!important;min-width:90px!important}}
    .user-card .user-actions{{order:3!important;display:block!important;width:100%!important;margin:0!important;padding:0!important;border:0!important}}
    .user-card .user-actions form,.user-card .user-actions form.inline{{display:flex!important;flex-direction:column!important;align-items:flex-start!important;gap:7px!important;width:100%!important;margin:0!important;padding:10px!important;background:#fbfdff!important;border:1px solid #e5eef7!important;border-radius:10px!important;box-sizing:border-box!important}}
    .user-card .user-actions form>div{{display:block!important;width:100%!important;max-width:480px!important}}
    .user-card .user-actions label{{margin:0 0 5px!important;font-size:11px!important}}
    .user-card .user-actions input{{display:block!important;width:min(420px,100%)!important;height:34px!important;padding:6px 8px!important;font-size:12px!important}}
    .user-card .user-actions .help{{margin-top:4px!important;font-size:11px!important}}
    .user-card .user-actions form button{{align-self:flex-start!important;min-width:130px!important;margin:0!important}}
    
    


    /* Ticket detail refinement: compact workflow and clearer lifecycle colors */
    .detail-grid{{gap:8px!important}}
    .detail-field,.detail-field:nth-child(2n),.detail-field:nth-child(4n),.detail-field-received_at,.detail-field-closed_at,.detail-field-note,.detail-field-issue{{padding:10px 11px!important;border-radius:10px!important}}
    .detail-field strong{{gap:7px!important;margin-bottom:5px!important;font-size:11.5px!important;min-height:0!important}}
    .detail-field .detail-icon{{width:27px!important;height:27px!important;flex-basis:27px!important;border-radius:8px!important}}
    .detail-field .detail-icon svg{{width:15px!important;height:15px!important}}
    .detail-field p{{font-size:12px!important;line-height:1.4!important}}
    .detail-field-device{{background:#edf6ff!important}}.detail-field-device .detail-icon{{background:#d8eaff!important;color:#2568a7!important}}
    .detail-field-device_code{{background:#f0f1ff!important}}.detail-field-device_code .detail-icon{{background:#dfe3ff!important;color:#575fb0!important}}
    .detail-field-phone{{background:#edf8f0!important}}.detail-field-phone .detail-icon{{background:#d7efde!important;color:#347b4a!important}}
    .detail-field-received_at{{background:#fff3d8!important;color:#80550a!important}}
    .detail-field-received_at strong{{color:#8c5d0a!important}}
    .detail-field-received_at .detail-icon{{background:#ffe3a3!important;color:#925d00!important}}
    .detail-field-closed_at{{background:#e5f8ed!important;color:#176642!important}}
    .detail-field-closed_at strong{{color:#1e754e!important}}
    .detail-field-closed_at .detail-icon{{background:#c9efd9!important;color:#1d7a4e!important}}
    .detail-time-progress{{display:inline-flex!important;align-items:center!important;gap:5px!important;padding:5px 9px!important;border-radius:999px!important;background:#ffedbd!important;color:#8b5a00!important;font-weight:800!important;font-size:11.5px!important}}
    .detail-time-closed{{display:inline-flex!important;align-items:center!important;gap:5px!important;padding:5px 9px!important;border-radius:999px!important;background:#c9f0d9!important;color:#176642!important;font-weight:800!important;font-size:11.5px!important}}
    .workflow-panel{{padding:12px!important;border-radius:12px!important;background:#f7faff!important;border:1px solid #e1eaf4!important;box-shadow:none!important;margin-top:12px!important}}
    .workflow-heading{{font-size:14px!important;margin:0 0 9px!important;gap:7px!important;color:#284f75!important}}
    .workflow-grid{{gap:8px!important}}
    .workflow-card{{padding:10px 11px!important;gap:6px!important;border:0!important;border-radius:10px!important;box-shadow:none!important;min-width:0!important}}
    .workflow-card-title{{font-size:12px!important;gap:7px!important;line-height:1.3!important}}
    .workflow-icon{{width:25px!important;height:25px!important;flex:0 0 25px!important;border-radius:8px!important}}
    .workflow-meta,.workflow-time{{font-size:11px!important;line-height:1.4!important;overflow-wrap:anywhere!important}}
    .workflow-pending{{background:#edf5ff!important}}
    .workflow-pending .workflow-card-title{{color:#285f9b!important}}
    .workflow-complete{{background:#e8f8ef!important}}
    .workflow-complete .workflow-card-title{{color:#237653!important}}
    .workflow-close-action{{background:#e4f7ec!important}}
    .workflow-close-action .workflow-card-title{{color:#176642!important}}
    .btn-take{{background:#3f7fce!important;color:#fff!important;border:0!important;border-radius:8px!important;box-shadow:0 2px 5px #285f9b22!important;font-weight:750!important;padding:7px 12px!important}}
    .btn-close{{background:#21865b!important;color:#fff!important;border:0!important;border-radius:8px!important;box-shadow:0 2px 5px #17664222!important;font-weight:750!important;padding:7px 12px!important}}
    .note-editor-panel{{background:#f7f5ff!important;border-color:#e7e0fb!important;padding:12px!important}}
    .note-editor-panel textarea{{min-height:66px!important;padding:9px 10px!important;font-size:12px!important;border-radius:9px!important}}
    .note-editor-panel label{{font-size:11px!important;margin-bottom:5px!important}}
    .workflow-actions{{margin-top:7px!important}}
    .workflow-actions button{{min-height:31px!important;font-size:11.5px!important}}
    


    /* Final refinement: values use matching icon/color chips */
    .detail-value-chip{{display:inline-flex!important;align-items:center!important;gap:7px!important;max-width:100%!important;box-sizing:border-box!important;padding:6px 10px!important;border-radius:9px!important;font-size:12.5px!important;font-weight:750!important;line-height:1.4!important;overflow-wrap:anywhere!important;vertical-align:middle!important}}
    .detail-value-chip svg{{flex:0 0 15px!important;width:15px!important;height:15px!important}}
    .value-device{{background:#e4f1ff!important;color:#245f98!important}}
    .value-device-code{{background:#eceeff!important;color:#555eaa!important;font-family:ui-monospace,SFMono-Regular,Consolas,monospace!important;letter-spacing:.2px!important}}
    .value-phone{{background:#e4f6e9!important;color:#2d7546!important}}
    .detail-field-device p,.detail-field-device_code p,.detail-field-phone p{{display:flex!important;align-items:flex-start!important}}
    .note-value-chip{{display:block!important;white-space:pre-wrap!important;overflow-wrap:anywhere!important;box-sizing:border-box!important;width:100%!important;padding:10px 12px!important;border-radius:10px!important;background:#e8f4ff!important;color:#194f7d!important;border-left:4px solid #4b91cb!important;font-size:13px!important;font-weight:650!important;line-height:1.6!important;box-shadow:inset 0 0 0 1px #d5e9fa!important}}
    .note-empty{{color:#7b8da0!important;font-size:12px!important;font-style:italic!important}}
    .detail-field-note{{background:#f1f7ff!important}}
    .detail-field-note .detail-icon{{background:#dcecff!important;color:#2868a4!important}}
    .detail-field-note p{{width:100%!important}}
    .detail-field-note p:has(.note-value-chip){{margin-top:2px!important}}
    


    /* Unified ticket detail value styling: one clinical icon/color system */
    .detail-field{{background:#fff!important;min-width:0!important}}
    .detail-field strong{{color:#42617f!important}}
    .detail-field .detail-icon{{width:27px!important;height:27px!important;flex:0 0 27px!important;border-radius:8px!important;background:#edf5fc!important;color:#3976a8!important;border:0!important}}
    .detail-field p{{color:#294866!important;font-size:13px!important;line-height:1.45!important;overflow-wrap:anywhere!important}}
    .detail-field-device .detail-icon{{background:#e8f3ff!important;color:#2873b8!important}}
    .detail-field-device_code .detail-icon{{background:#f0edff!important;color:#6b59b5!important}}
    .detail-field-phone .detail-icon{{background:#e8f8f0!important;color:#24815e!important}}
    .detail-field-time .detail-icon{{background:#eaf6ff!important;color:#2875a8!important}}
    .detail-field-issue_type .detail-icon{{background:#fff3e5!important;color:#a96618!important}}
    .detail-field-severity .detail-icon{{background:#fff0ef!important;color:#bd5151!important}}
    .detail-field-reporter .detail-icon{{background:#e7f8fb!important;color:#16829a!important}}
    .detail-field-handler .detail-icon{{background:#eaf2ff!important;color:#3f69b4!important}}
    .detail-field-department .detail-icon{{background:#e6f8ff!important;color:#0784aa!important}}
    .detail-field-location .detail-icon{{background:#f0eaff!important;color:#7359b7!important}}
    .detail-field-issue{{background:#f5f9ff!important;border:1px solid #dce9f8!important;border-radius:10px!important}}
    .detail-field-note{{background:#f1f8ff!important;border:1px solid #cfe3fb!important;border-radius:10px!important}}
    .detail-field-note p{{padding:10px 12px!important;background:#fff!important;border-left:3px solid #70aee8!important;border-radius:7px!important;color:#244967!important;font-size:13px!important;font-weight:600!important;line-height:1.55!important;box-shadow:0 1px 3px rgba(35,85,130,.04)!important}}
    .detail-field-received_at{{background:#fff9ed!important;border:1px solid #f5dfb1!important;border-radius:10px!important}}
    .detail-field-received_at .detail-icon{{background:#ffefcf!important;color:#a96512!important}}
    .detail-field-closed_at{{background:#effaf4!important;border:1px solid #c8ead7!important;border-radius:10px!important}}
    .detail-field-closed_at .detail-icon{{background:#dcf5e7!important;color:#21805a!important}}
    .detail-field .detail-tag,.detail-field .reporter-chip,.detail-field .handler-chip{{max-width:100%;width:fit-content;box-sizing:border-box;border:1px solid transparent!important;border-radius:8px!important;padding:5px 9px!important;font-size:12px!important;font-weight:700!important;line-height:1.35!important;white-space:normal!important;overflow-wrap:anywhere!important}}
    .detail-field .detail-tag-dept{{background:#e8f7fc!important;color:#167b9c!important;border-color:#d2edf5!important}}
    .detail-field .detail-tag-location{{background:#f0edff!important;color:#6956aa!important;border-color:#e1dcfb!important}}
    .detail-field .reporter-chip{{background:#eaf8fb!important;color:#167b91!important;border-color:#ccecf2!important}}
    .detail-field .handler-chip,.detail-field .handler-value{{background:#eaf3ff!important;color:#2e65a3!important;border-color:#d2e4fb!important}}
    .detail-field .detail-time-progress{{background:#fff1d9!important;color:#92580e!important;border:1px solid #f4d79d!important}}
    .detail-field .detail-time-closed{{background:#e6f7ed!important;color:#21724f!important;border:1px solid #c4e8d3!important}}
    .detail-field .detail-time-pending{{background:#eef3f8!important;color:#647b91!important;border:1px solid #e0e8f0!important}}
    


    /* v9: cohesive clinical ticket status surfaces; single-tone processing note */
    .workflow-panel{{background:#f4f8fc!important;border:0!important;border-radius:14px!important;box-shadow:none!important;padding:12px!important}}
    .workflow-card{{border:0!important;border-radius:12px!important;box-shadow:none!important;padding:12px 13px!important;transition:background .15s ease!important}}
    .workflow-card.workflow-pending{{background:#eaf3ff!important}}
    .workflow-card.workflow-complete{{background:#e9f7ef!important}}
    .workflow-card.workflow-close-action{{background:#fff4df!important}}
    .workflow-card-title{{font-size:13px!important;font-weight:800!important;line-height:1.35!important}}
    .workflow-pending .workflow-card-title{{color:#245f9e!important}}
    .workflow-complete .workflow-card-title{{color:#216d4b!important}}
    .workflow-close-action .workflow-card-title{{color:#8b5a12!important}}
    .workflow-meta{{color:#526b82!important;font-size:11.5px!important;line-height:1.45!important}}
    .workflow-time{{display:inline-flex!important;align-items:center!important;width:fit-content!important;max-width:100%!important;padding:5px 8px!important;border-radius:8px!important;background:#d7f0e1!important;color:#216d4b!important;font-size:12px!important;font-weight:800!important;overflow-wrap:anywhere!important}}
    .workflow-icon-blue{{background:#d5e8ff!important;color:#245f9e!important}}
    .workflow-icon-green{{background:#cdebd9!important;color:#216d4b!important}}
    .workflow-actions{{margin-top:8px!important}}
    .workflow-actions .btn-take{{background:#2f75b8!important;color:#fff!important;border:0!important;border-radius:8px!important;box-shadow:0 2px 5px rgba(47,117,184,.16)!important}}
    .workflow-actions .btn-close{{background:#2c8a61!important;color:#fff!important;border:0!important;border-radius:8px!important;box-shadow:0 2px 5px rgba(44,138,97,.16)!important}}
    .pill.st-new{{background:#e3efff!important;color:#245d99!important;border:0!important;border-radius:999px!important;padding:5px 9px!important}}
    .pill.st-progress{{background:#eee8ff!important;color:#6650a8!important;border:0!important;border-radius:999px!important;padding:5px 9px!important}}
    .pill.st-closed{{background:#dff5e8!important;color:#20764f!important;border:0!important;border-radius:999px!important;padding:5px 9px!important}}
    .pill.sev-low{{background:#e7f7ed!important;color:#23764f!important;border:0!important;border-radius:999px!important;padding:5px 9px!important}}
    .pill.sev-medium{{background:#fff3d8!important;color:#94610f!important;border:0!important;border-radius:999px!important;padding:5px 9px!important}}
    .pill.sev-high{{background:#ffebeb!important;color:#b33b43!important;border:0!important;border-radius:999px!important;padding:5px 9px!important}}
    .detail-field-note{{background:#eaf3ff!important;border:0!important;border-radius:11px!important;box-shadow:none!important}}
    .detail-field-note .detail-icon{{background:#d5e7ff!important;color:#245f9e!important}}
    .detail-field-note p{{background:#eaf3ff!important;border:0!important;border-radius:7px!important;padding:8px 10px!important;color:#193f63!important;font-weight:800!important;font-size:13px!important;line-height:1.55!important;box-shadow:none!important}}
    .detail-field-note .note-value-chip{{background:transparent!important;color:#193f63!important;font-weight:800!important;border:0!important;box-shadow:none!important;padding:0!important}}
    


    /* Clinical color pass: give requested ticket fields clear, harmonious backgrounds */
    .detail-field-device{{background:#edf5ff!important;color:#244e78!important}}
    .detail-field-device .detail-icon{{background:#d8eaff!important;color:#2868a4!important}}
    .detail-field-device strong{{color:#315f8a!important}}
    .detail-field-issue_type{{background:#fff5e6!important;color:#76501d!important}}
    .detail-field-issue_type .detail-icon{{background:#ffe8bf!important;color:#a96a17!important}}
    .detail-field-issue_type strong{{color:#8d611f!important}}
    .detail-field-phone{{background:#edf8f0!important;color:#2d6542!important}}
    .detail-field-phone .detail-icon{{background:#d8efdf!important;color:#347b4b!important}}
    .detail-field-phone strong{{color:#39724b!important}}
    .detail-field-severity{{background:#fff1ef!important;color:#854b42!important}}
    .detail-field-severity .detail-icon{{background:#ffdfd9!important;color:#b85c49!important}}
    .detail-field-severity strong{{color:#a65345!important}}
    .detail-field-handler{{background:#f0f2ff!important;color:#3c4d88!important}}
    .detail-field-handler .detail-icon{{background:#dfe5ff!important;color:#5267b1!important}}
    .detail-field-handler strong{{color:#4d5f9e!important}}
    .detail-field-department{{background:#eaf8fa!important;color:#23656e!important}}
    .detail-field-department .detail-icon{{background:#d1eff2!important;color:#167d8b!important}}
    .detail-field-department strong{{color:#237782!important}}
    /* Keep inner value chips coordinated with their parent field */
    .detail-field-device p,.detail-field-issue_type p,.detail-field-phone p,.detail-field-severity p,.detail-field-handler p,.detail-field-department p{{color:inherit!important}}
    .detail-field-device .detail-tag,.detail-field-issue_type .issue-type-chip,.detail-field-phone .reporter-chip,.detail-field-handler .handler-value,.detail-field-department .detail-tag-dept{{box-shadow:none!important}}
    

    /* Final UI quality pass: clearer priority/status hierarchy, consistent with clinical IT dashboards */
    .ticket-table thead th.priority-heading,
    .ticket-table thead th.status-heading{{
        font-size:12px!important;font-weight:800!important;line-height:1.25!important;
        color:#24496f!important;background:#edf4fb!important;padding:10px 5px!important;
    }}
    .ticket-table tbody td.priority-cell,
    .ticket-table tbody td.status-cell{{padding:9px 5px!important;}}
    .ticket-table td.priority-cell .pill,
    .ticket-table td.status-cell .pill{{
        gap:5px!important;padding:6px 9px!important;min-height:30px;
        border-radius:9px!important;font-size:12px!important;font-weight:800!important;
        line-height:1.25!important;letter-spacing:.05px;
        box-shadow:inset 0 0 0 1px rgba(36,73,111,.06);
    }}
    .ticket-table td.priority-cell .pill-label,
    .ticket-table td.status-cell .pill-label{{
        font-size:12px!important;font-weight:800!important;line-height:1.25!important;
        overflow-wrap:normal!important;word-break:normal!important;
    }}
    .ticket-table td.priority-cell .pill-icon,
    .ticket-table td.status-cell .pill-icon{{
        flex-basis:13px!important;width:13px!important;height:13px!important;
        font-size:11px!important;font-weight:900!important;
    }}
    .ticket-table td.priority-cell .sev-low{{background:#e7f7ed!important;color:#216b49!important}}
    .ticket-table td.priority-cell .sev-medium{{background:#fff2d7!important;color:#8c5a0b!important}}
    .ticket-table td.priority-cell .sev-high{{background:#ffeaeb!important;color:#ad303b!important}}
    .ticket-table td.status-cell .st-new{{background:#e5f0ff!important;color:#205b9b!important}}
    .ticket-table td.status-cell .st-progress{{background:#eee9ff!important;color:#5e45a6!important}}
    .ticket-table td.status-cell .st-closed{{background:#def5e9!important;color:#176d4b!important}}
    .ticket-table tbody tr:hover td.priority-cell,
    .ticket-table tbody tr:hover td.status-cell{{background:#f7fbff!important}}
    .ticket-table .pill-label{{text-wrap:balance}}
    
    /* Shared visual rhythm: clean clinical surfaces, readable forms, and predictable keyboard focus */
    .box{{border-color:#e1eaf4}}
    .ticket-toolbar-panel,.ticket-table-panel,.detail-box,.workflow-panel{{box-shadow:0 5px 18px rgba(27,67,112,.045)}}
    input:focus-visible,select:focus-visible,textarea:focus-visible,button:focus-visible,a:focus-visible{{outline:3px solid rgba(45,132,204,.28)!important;outline-offset:2px}}
    .ticket-table thead th{{letter-spacing:0!important}}
    /* v12: unify every ticket-list column into one calm clinical table system */
    .table-wrap{{border:1px solid #d9e5f2!important;border-radius:14px!important;overflow:auto!important;background:#fff;box-shadow:0 4px 16px rgba(35,76,119,.045)}}
    .ticket-table{{border:0!important;border-radius:14px!important;overflow:hidden;table-layout:fixed}}
    .ticket-table thead th{{background:#edf4fb!important;color:#365574!important;font-family:Inter,'Segoe UI',Arial,sans-serif!important;font-size:11px!important;font-weight:750!important;line-height:1.35!important;letter-spacing:.01em!important;padding:12px 8px!important;border-right:1px solid #dce7f2!important;border-bottom:1px solid #d5e2ef!important;text-align:left;vertical-align:middle}}
    .ticket-table thead th:first-child{{text-align:center}}
    .ticket-table tbody td{{background:#fff!important;color:#344d68!important;font-family:Inter,'Segoe UI',Arial,sans-serif!important;font-size:12px!important;font-weight:450;line-height:1.45!important;padding:12px 9px!important;border-right:1px solid #e6edf5!important;border-bottom:1px solid #e6edf5!important;vertical-align:middle!important;transition:background-color .16s ease,color .16s ease}}
    .ticket-table tbody tr:nth-child(even) td{{background:#f9fbfe!important}}
    .ticket-table tbody tr:hover td{{background:#eef6ff!important}}
    .ticket-table thead th:last-child,.ticket-table tbody td:last-child{{border-right:0!important}}
    .ticket-table tbody tr:last-child td{{border-bottom:0!important}}
    .ticket-table .row-number{{text-align:center!important;color:#8194a9!important;font-size:11px!important;font-weight:650!important;font-variant-numeric:tabular-nums}}
    .ticket-table .ticket-code{{display:inline-block;color:#1c5b91!important;font-size:12px!important;font-weight:750!important;text-decoration:none!important;background:#edf6ff;border:1px solid #d9eafa;border-radius:7px;padding:5px 7px;white-space:normal;overflow-wrap:anywhere;line-height:1.25}}
    .ticket-table .ticket-code:hover{{background:#e1f0ff!important;border-color:#b9d9f5!important}}
    .ticket-table .device-cell{{display:inline-flex;align-items:center;gap:7px;white-space:normal;line-height:1.35;font-weight:550;color:#345b7b}}
    .ticket-table .mini-icon{{width:27px!important;height:27px!important;flex:0 0 27px;border-radius:8px!important;background:#eaf3fc!important;color:#3876aa!important;font-size:16px!important}}
    .ticket-table .issue-cell{{color:#4b6077!important;font-size:11.5px!important;line-height:1.5!important;overflow-wrap:anywhere}}
    .ticket-table .dept-chip,.ticket-table .handler-chip{{display:inline-flex;align-items:center;gap:5px;max-width:100%;padding:5px 8px!important;border-radius:8px!important;background:#f0f6fb!important;border:1px solid #e0eaf3!important;color:#47647f!important;font-size:11px!important;font-weight:600!important;line-height:1.35!important;white-space:normal;overflow-wrap:anywhere}}
    .ticket-table .dept-chip svg{{flex:0 0 13px;color:#6b8cab}}
    .ticket-table .date-cell{{color:#5d7188!important;font-size:11px!important;font-variant-numeric:tabular-nums;white-space:normal;overflow-wrap:anywhere}}
    .ticket-table .btn-view{{min-width:54px!important;padding:7px 10px!important;border:1px solid #c5dff5!important;border-radius:8px!important;background:#f5faff!important;color:#2169a6!important;font-size:11px!important;font-weight:750!important;line-height:1.2!important;text-align:center;box-shadow:none!important;transition:background .15s ease,border-color .15s ease}}
    .ticket-table .btn-view:hover{{background:#e5f2ff!important;border-color:#9ec8ec!important}}
    .ticket-table td.priority-cell,.ticket-table td.status-cell{{text-align:center!important;padding:10px 6px!important}}
    .ticket-table td.priority-cell .pill,.ticket-table td.status-cell .pill{{display:inline-flex;align-items:center;justify-content:center;gap:5px;max-width:100%;min-height:30px;padding:6px 9px!important;border-radius:9px!important;font-size:12px!important;font-weight:750!important;line-height:1.25!important;white-space:normal;text-align:center;box-shadow:inset 0 0 0 1px rgba(36,73,111,.06)}}
    .ticket-table td.priority-cell .pill-label,.ticket-table td.status-cell .pill-label{{font-size:12px!important;font-weight:750!important;line-height:1.25!important;overflow-wrap:anywhere!important}}
    .ticket-table td.priority-cell .pill-icon,.ticket-table td.status-cell .pill-icon{{width:13px!important;height:13px!important;flex:0 0 13px!important;font-size:10px!important;line-height:1!important}}
    .ticket-table td.priority-cell .sev-low{{background:#e7f6ed!important;color:#246b4b!important}}
    .ticket-table td.priority-cell .sev-medium{{background:#fff2d8!important;color:#895a13!important}}
    .ticket-table td.priority-cell .sev-high{{background:#ffeaeb!important;color:#a9333d!important}}
    .ticket-table td.status-cell .st-new{{background:#e6f0ff!important;color:#285d9b!important}}
    .ticket-table td.status-cell .st-progress{{background:#eee9ff!important;color:#6048a3!important}}
    .ticket-table td.status-cell .st-closed{{background:#def4e8!important;color:#1c704e!important}}
    

    /* Department column: stable chip sizing and predictable text wrapping. */
    .ticket-table th:nth-child(5),.ticket-table td.department-cell{{vertical-align:middle!important}}
    .ticket-table td.department-cell{{padding-left:6px!important;padding-right:6px!important}}
    .ticket-table .dept-chip{{display:flex!important;align-items:center!important;justify-content:flex-start!important;gap:6px!important;width:100%!important;max-width:100%!important;min-width:0!important;min-height:32px;box-sizing:border-box!important;padding:6px 7px!important;border:1px solid #ccebf0!important;border-radius:9px!important;background:#effafb!important;color:#246b75!important;white-space:normal!important;line-height:1.25!important;text-align:left!important;overflow:hidden!important;overflow-wrap:normal!important;word-break:normal!important}}
    .ticket-table .dept-chip>svg{{flex:0 0 13px!important;width:13px!important;height:13px!important;align-self:flex-start;margin-top:1px}}
    .ticket-table .dept-chip>span{{display:-webkit-box!important;flex:1 1 auto!important;min-width:0!important;max-width:100%!important;white-space:normal!important;overflow-wrap:anywhere!important;word-break:normal!important;line-height:1.25!important;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden!important;text-overflow:ellipsis}}
    


    /* Harmonized ticket-list rhythm: equal alignment, restrained medical palette, no icons in Department/Priority. */
    .ticket-table thead th{{padding:12px 9px!important;font-size:12px!important;line-height:1.3!important;color:#355879!important;background:#f0f6fc!important;border-bottom:1px solid #dce8f4!important}}
    .ticket-table tbody td{{padding:11px 9px!important;line-height:1.4!important;font-size:12px!important;vertical-align:middle!important;border-bottom:1px solid #e7eef6!important;color:#294560}}
    .ticket-table tbody tr{{height:auto;transition:background-color .16s ease}}
    .ticket-table tbody tr:hover td{{background:#f5faff!important}}
    .ticket-table .ticket-code,.ticket-table .device-cell,.ticket-table .issue-cell,.ticket-table .date-cell{{line-height:1.4}}
    .ticket-table .department-label{{display:inline-flex;align-items:center;justify-content:center;max-width:100%;min-height:27px;box-sizing:border-box;padding:5px 9px!important;border:1px solid #cceaf0;border-radius:8px!important;background:#eaf8fa!important;color:#246b75!important;font-size:11px!important;font-weight:650;line-height:1.3!important;text-align:center;white-space:normal;overflow-wrap:anywhere;word-break:normal}}
    .ticket-table .priority-label{{display:inline-flex;align-items:center;justify-content:center;min-width:64px;min-height:28px;box-sizing:border-box;padding:5px 9px!important;border:1px solid transparent;border-radius:8px!important;font-size:11px!important;font-weight:750;line-height:1.25!important;text-align:center;white-space:normal}}
    .ticket-table .sev-low{{background:#e7f7f0!important;color:#167653!important;border-color:#c5ead9!important}}
    .ticket-table .sev-medium{{background:#fff5df!important;color:#94620b!important;border-color:#f3dfad!important}}
    .ticket-table .sev-high{{background:#ffeded!important;color:#b63745!important;border-color:#f5cfd2!important}}
    .ticket-table .st-new,.ticket-table .st-progress,.ticket-table .st-closed{{display:inline-flex;align-items:center;justify-content:center;min-height:28px;box-sizing:border-box;padding:5px 9px!important;border:1px solid transparent;border-radius:8px!important;font-size:11px!important;font-weight:700;line-height:1.25!important;text-align:center;white-space:normal}}
    .ticket-table .st-new{{background:#eaf2ff!important;color:#2c62a4!important;border-color:#d2e2fb!important}}
    .ticket-table .st-progress{{background:#f1edff!important;color:#674ab0!important;border-color:#ded5fb!important}}
    .ticket-table .st-closed{{background:#e6f7ef!important;color:#187653!important;border-color:#c8ead9!important}}
    .ticket-table .reporter-chip,.ticket-table .handler-chip{{min-height:27px;box-sizing:border-box;border-radius:8px!important;padding:5px 8px!important;line-height:1.3!important;font-size:11px!important}}
    .ticket-table .btn-view{{min-height:28px;border-radius:8px!important;font-size:11px!important;font-weight:700}}
    .ticket-table td:nth-child(1){{font-variant-numeric:tabular-nums;color:#7890a8}}
    


    /* Final compact ticket list: match supplied FMP reference and keep every column aligned. */
    .ticket-table-panel{{padding:10px!important;overflow:visible!important}}
    .table-wrap{{width:100%;overflow-x:auto!important;border:1px solid #dbe8f5!important;border-radius:9px!important;background:#fff!important;box-shadow:none!important}}
    .ticket-table{{width:100%!important;min-width:0!important;table-layout:fixed!important;border:0!important;border-collapse:collapse!important;border-radius:0!important}}
    .ticket-table thead th{{padding:8px 4px!important;font-size:9px!important;line-height:1.2!important;text-transform:none!important;letter-spacing:0!important;white-space:normal!important;text-align:center!important;overflow-wrap:normal!important;word-break:normal!important;background:#edf5ff!important;color:#183b66!important;border-right:1px solid #e1ebf6!important;border-bottom:1px solid #dce8f5!important}}
    .ticket-table tbody td{{padding:7px 4px!important;font-size:9px!important;line-height:1.25!important;vertical-align:middle!important;text-align:center!important;overflow-wrap:normal!important;word-break:normal!important;border-right:1px solid #e7eef7!important;border-bottom:1px solid #e7eef7!important;color:#16365c!important}}
    .ticket-table tbody tr:nth-child(even) td{{background:#fbfdff!important}}
    .ticket-table tbody tr:hover td{{background:#f0f7ff!important}}
    .ticket-table th:nth-child(1),.ticket-table td:nth-child(1){{width:3%!important}}
    .ticket-table th:nth-child(2),.ticket-table td:nth-child(2){{width:7%!important}}
    .ticket-table th:nth-child(3),.ticket-table td:nth-child(3){{width:8%!important}}
    .ticket-table th:nth-child(4),.ticket-table td:nth-child(4){{width:8%!important}}
    .ticket-table th:nth-child(5),.ticket-table td:nth-child(5){{width:9%!important}}
    .ticket-table th:nth-child(6),.ticket-table td:nth-child(6){{width:7%!important}}
    .ticket-table th:nth-child(7),.ticket-table td:nth-child(7){{width:9%!important}}
    .ticket-table th:nth-child(8),.ticket-table td:nth-child(8){{width:8%!important}}
    .ticket-table th:nth-child(9),.ticket-table td:nth-child(9){{width:7%!important}}
    .ticket-table th:nth-child(10),.ticket-table td:nth-child(10){{width:8%!important}}
    .ticket-table th:nth-child(11),.ticket-table td:nth-child(11){{width:10%!important}}
    .ticket-table th:nth-child(12),.ticket-table td:nth-child(12){{width:6%!important}}
    .ticket-table .ticket-code{{font-size:9px!important;padding:3px 4px!important;border-radius:5px!important;white-space:nowrap!important;overflow-wrap:normal!important}}
    .ticket-table .device-cell{{display:inline-flex!important;justify-content:center!important;gap:3px!important;font-size:9px!important;white-space:normal!important;overflow-wrap:anywhere!important}}
    .ticket-table .mini-icon{{display:none!important}}
    .ticket-table .dept-chip,.ticket-table .department-label,.ticket-table .handler-chip,.ticket-table .reporter-chip{{display:inline-flex!important;justify-content:center!important;max-width:100%!important;min-height:0!important;padding:3px 4px!important;font-size:8.5px!important;line-height:1.2!important;border-radius:6px!important;white-space:normal!important;overflow-wrap:anywhere!important;word-break:normal!important}}
    .ticket-table td.priority-cell,.ticket-table td.status-cell{{padding:5px 3px!important}}
    .ticket-table td.priority-cell .pill,.ticket-table td.status-cell .pill,.ticket-table .priority-label,.ticket-table .st-new,.ticket-table .st-progress,.ticket-table .st-closed{{display:inline-flex!important;justify-content:center!important;align-items:center!important;min-height:0!important;padding:4px 5px!important;font-size:8.5px!important;line-height:1.15!important;border-radius:7px!important;white-space:normal!important;gap:2px!important}}
    .ticket-table .pill-icon{{display:none!important}}
    .ticket-table .btn-view{{min-width:0!important;min-height:0!important;padding:4px 8px!important;font-size:9px!important;line-height:1.1!important;border-radius:5px!important;white-space:nowrap!important}}
    .ticket-table .date-cell{{font-size:8.5px!important;white-space:nowrap!important;overflow-wrap:normal!important;font-variant-numeric:tabular-nums}}
    .ticket-pagination{{display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;padding:12px 3px 2px;color:#68809c;font-size:10px}}
    .pagination-buttons{{display:flex;align-items:center;gap:4px;flex-wrap:wrap}}
    .page-button{{display:inline-flex;align-items:center;justify-content:center;min-width:25px;height:27px;padding:0 6px;box-sizing:border-box;border:1px solid #d9e5f3;border-radius:6px;background:#fff;color:#294d78!important;text-decoration:none!important;font-size:10px;font-weight:600;line-height:1}}
    .page-button:hover{{background:#edf6ff;border-color:#a9cdec}}
    .page-button.active{{background:#1478d4!important;border-color:#1478d4!important;color:#fff!important;box-shadow:0 2px 5px rgba(20,120,212,.18)}}
    .page-button[aria-disabled='true']{{opacity:.45;pointer-events:none}}
    .page-arrow{{font-size:15px}}
    .page-ellipsis{{padding:0 2px;color:#8ca0b7}}
    
    
    /* Final typography and icon pass requested by the user. Keep the current compact table geometry. */
    .ticket-table thead th{{font-size:12px!important;font-weight:700!important;line-height:1.25!important;color:#38577d!important;letter-spacing:0!important;text-transform:none!important}}
    .ticket-table tbody td{{font-size:12px!important;line-height:1.35!important}}
    .ticket-table .ticket-code{{font-size:11.5px!important}}
    .ticket-table .dept-chip,.ticket-table .department-label,.ticket-table .handler-chip,.ticket-table .reporter-chip{{font-size:10px!important;line-height:1.25!important}}
    .ticket-table .issue-category-chip{{font-size:10.5px!important;line-height:1.25!important}}
    .ticket-table .date-cell{{font-size:11px!important}}
    .ticket-table .btn-view{{font-size:11px!important}}
    .ticket-table thead th.priority-heading,.ticket-table thead th.status-heading{{font-size:12px!important;font-weight:700!important;line-height:1.25!important;color:#38577d!important;background:#f5f8fc!important;padding:10px 5px!important;white-space:normal!important}}
    .ticket-table td.priority-cell .priority-label,.ticket-table td.status-cell .pill,.ticket-table td.priority-cell .pill-label,.ticket-table td.status-cell .pill-label{{font-size:10.5px!important;font-weight:750!important;line-height:1.2!important}}
    .section-icon svg,.quick-filter svg,.calendar-pick-button svg{{display:block;flex:0 0 auto}}
    .section-icon svg{{width:19px;height:19px}}
    .quick-filter svg{{width:14px;height:14px}}
    .calendar-pick-button svg{{width:16px;height:16px}}
    .quick-filter{{gap:7px}}
    .calendar-pick-button{{border-radius:9px!important}}
    
    /* Previous compact balance pass retained as the base geometry. */
    .ticket-table th.priority-heading,.ticket-table th.status-heading,
    .ticket-table td.priority-cell,.ticket-table td.status-cell{{
      width:7%!important; min-width:0!important; padding-left:3px!important;padding-right:3px!important;
      text-align:center!important; vertical-align:middle!important;
    }}
    .ticket-table th.priority-heading,.ticket-table th.status-heading{{font-size:9px!important;line-height:1.15!important;font-weight:750!important;white-space:normal!important}}
    .ticket-table td.priority-cell .priority-label,.ticket-table td.status-cell .pill{{
      display:inline-flex!important;align-items:center!important;justify-content:center!important;
      width:auto!important;max-width:100%!important;min-width:0!important;min-height:23px!important;
      padding:3px 5px!important;border-radius:7px!important;font-size:9px!important;line-height:1.15!important;
      font-weight:700!important;white-space:normal!important;overflow-wrap:normal!important;word-break:normal!important;
    }}
    .ticket-table td.priority-cell .pill-label,.ticket-table td.status-cell .pill-label{{font-size:9px!important;line-height:1.15!important;font-weight:700!important}}
    .ticket-table td.status-cell .pill-icon{{display:inline-flex!important;align-items:center;justify-content:center;font-size:9px!important;flex:0 0 auto!important}}
    .ticket-table .mini-icon{{display:inline-flex!important;align-items:center;justify-content:center;flex:0 0 13px;width:13px;height:13px;color:#2580c4!important;vertical-align:middle}}
    .ticket-table .mini-icon svg{{display:block;width:13px;height:13px}}
    .ticket-table .device-cell{{align-items:center!important;gap:3px!important}}
    .ticket-table .issue-category-chip{{display:inline-flex;align-items:center;justify-content:center;gap:4px;max-width:100%;box-sizing:border-box;padding:4px 6px;border:1px solid #d9e7f7;border-radius:7px;background:#eef6ff;color:#245f99;font-size:9px;font-weight:650;line-height:1.2;white-space:normal;overflow-wrap:anywhere;text-align:left}}
    .ticket-table .issue-category-icon{{display:inline-flex;align-items:center;justify-content:center;flex:0 0 13px;width:13px;height:13px;color:currentColor}}
    .ticket-table .issue-category-icon svg{{display:block;width:13px;height:13px}}
    .ticket-table .issue-category-Phần-cứng{{background:#edf5ff;color:#2764a5;border-color:#d6e7fb}}
    .ticket-table .issue-category-Phần-mềm{{background:#f2edff;color:#674bb0;border-color:#e1d7ff}}
    .ticket-table .issue-category-Kết-nối-mạng{{background:#eaf5ff;color:#176da8;border-color:#cfe8fb}}
    .ticket-table .issue-category-Tài-khoản-and-phân-quyền{{background:#fff0f2;color:#a83d56;border-color:#f8d6de}}
    .ticket-table .issue-category-Máy-in--Scan{{background:#eafaf1;color:#24764f;border-color:#cceedd}}
    .ticket-table .issue-category-Email-and-ứng-dụng{{background:#fff6e5;color:#95600b;border-color:#f5e3b6}}
    .ticket-table .issue-category-Điện-thoại{{background:#eaf7fb;color:#236e89;border-color:#d0edf5}}
    .ticket-table .issue-category-Thiết-bị-ngoại-vi{{background:#eef2ff;color:#485fa4;border-color:#dbe2ff}}
    .ticket-table .issue-category-Bảo-mật{{background:#fff0ed;color:#a64c35;border-color:#f7d8d0}}
    .ticket-table .issue-category-Khác{{background:#f1f5f9;color:#536579;border-color:#dce5ee}}
    

    /* Last override: balanced, readable ticket table typography without changing column layout. */
    .ticket-table thead th{{font-size:12px!important;font-weight:700!important;line-height:1.25!important;color:#38577d!important;letter-spacing:0!important;text-transform:none!important}}
    .ticket-table tbody td{{font-size:12px!important;line-height:1.35!important}}
    .ticket-table .ticket-code{{font-size:11.5px!important}}
    .ticket-table .dept-chip,.ticket-table .department-label,.ticket-table .handler-chip,.ticket-table .reporter-chip{{font-size:10px!important;line-height:1.25!important}}
    .ticket-table .issue-category-chip{{font-size:10px!important;line-height:1.25!important}}
    .ticket-table .date-cell{{font-size:11px!important}}
    .ticket-table .btn-view{{font-size:11px!important}}
    .ticket-table thead th.priority-heading,.ticket-table thead th.status-heading{{font-size:12px!important;font-weight:700!important;line-height:1.25!important;color:#38577d!important;background:#f5f8fc!important;padding:10px 5px!important;white-space:normal!important}}
    .ticket-table td.priority-cell .priority-label,.ticket-table td.status-cell .pill,.ticket-table td.priority-cell .pill-label,.ticket-table td.status-cell .pill-label{{font-size:10px!important;font-weight:750!important;line-height:1.2!important}}
    .section-icon svg{{width:19px;height:19px;display:block}}
    .quick-filter svg{{width:14px;height:14px;display:block}}
    .calendar-pick-button svg{{width:16px;height:16px;display:block}}
    

    /* Final compact table pass: preserve approved geometry and keep all content on one line. */
    .table-wrap{{width:100%;overflow-x:auto!important;}}
    .ticket-table{{width:100%!important;min-width:960px!important;table-layout:fixed!important;border-collapse:collapse!important;}}
    .ticket-table thead th{{white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important;font-size:11px!important;font-weight:700!important;line-height:1.2!important;padding:9px 6px!important;text-align:center!important;}}
    .ticket-table tbody td{{white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important;font-size:11px!important;line-height:1.25!important;padding:8px 6px!important;vertical-align:middle!important;}}
    .ticket-table .device-cell{{display:block!important;white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important;font-size:11px!important;line-height:1.25!important;font-weight:500!important;}}
    .ticket-table .issue-category-chip{{display:inline-block!important;max-width:100%!important;box-sizing:border-box!important;white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important;vertical-align:middle!important;padding:4px 6px!important;font-size:10px!important;line-height:1.2!important;text-align:center!important;}}
    .ticket-table .issue-category-chip>span{{white-space:nowrap!important;}}
    .ticket-table .ticket-code{{white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important;max-width:100%!important;font-size:10.5px!important;padding:4px 5px!important;}}
    .ticket-table .dept-chip,.ticket-table .department-label,.ticket-table .handler-chip,.ticket-table .reporter-chip{{display:inline-block!important;max-width:100%!important;white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important;font-size:10px!important;line-height:1.2!important;padding:4px 5px!important;vertical-align:middle!important;}}
    .ticket-table td.priority-cell,.ticket-table td.status-cell{{white-space:nowrap!important;text-align:center!important;padding:6px 3px!important;}}
    .ticket-table td.priority-cell .pill,.ticket-table td.status-cell .pill,.ticket-table td.priority-cell .pill-label,.ticket-table td.status-cell .pill-label{{display:inline-flex!important;align-items:center!important;justify-content:center!important;white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important;overflow-wrap:normal!important;word-break:normal!important;font-size:10px!important;line-height:1.15!important;padding:4px 5px!important;min-height:0!important;max-width:100%!important;}}
    .ticket-table td.priority-cell .pill-icon,.ticket-table td.status-cell .pill-icon{{display:none!important;}}
    .ticket-table .date-cell{{white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important;font-size:10px!important;font-variant-numeric:tabular-nums!important;}}
    .ticket-table .btn-view{{white-space:nowrap!important;padding:5px 7px!important;font-size:10px!important;min-width:0!important;}}
    .ticket-table th:nth-child(1),.ticket-table td:nth-child(1){{width:3%!important;text-align:center!important;}}
    .ticket-table th:nth-child(2),.ticket-table td:nth-child(2){{width:8%!important;}}
    .ticket-table th:nth-child(3),.ticket-table td:nth-child(3){{width:8%!important;}}
    .ticket-table th:nth-child(4),.ticket-table td:nth-child(4){{width:9%!important;}}
    .ticket-table th:nth-child(5),.ticket-table td:nth-child(5){{width:9%!important;}}
    .ticket-table th:nth-child(6),.ticket-table td:nth-child(6){{width:7%!important;}}
    .ticket-table th:nth-child(7),.ticket-table td:nth-child(7){{width:10%!important;}}
    .ticket-table th:nth-child(8),.ticket-table td:nth-child(8){{width:6%!important;}}
    .ticket-table th:nth-child(9),.ticket-table td:nth-child(9){{width:7%!important;}}
    .ticket-table th:nth-child(10),.ticket-table td:nth-child(10){{width:8%!important;}}
    .ticket-table th:nth-child(11),.ticket-table td:nth-child(11){{width:10%!important;}}
    .ticket-table th:nth-child(12),.ticket-table td:nth-child(12){{width:7%!important;}}
    .ticket-pagination{{display:flex!important;align-items:center!important;justify-content:space-between!important;gap:12px!important;flex-wrap:wrap!important;}}
    .pagination-counts{{display:flex;align-items:center;gap:12px;flex-wrap:wrap;min-width:0;}}
    .pagination-summary,.pagination-all-count{{display:inline-block;white-space:nowrap;color:#68809c;font-size:10px;font-weight:500;}}
    .pagination-all-count{{padding-left:11px;border-left:1px solid #cbd9e8;color:#315f8e;font-weight:700;}}
    


    /* Final correction: keep every ticket value visible; use horizontal scroll instead of truncation. */
    .table-wrap{{width:100%!important;max-width:100%!important;overflow-x:auto!important;overflow-y:visible!important;}}
    .ticket-table{{width:max-content!important;min-width:100%!important;table-layout:auto!important;border-collapse:collapse!important;}}
    .ticket-table thead th,.ticket-table tbody td{{box-sizing:border-box!important;white-space:nowrap!important;overflow:visible!important;text-overflow:clip!important;overflow-wrap:normal!important;word-break:normal!important;}}
    .ticket-table thead th{{padding:9px 9px!important;font-size:11px!important;}}
    .ticket-table tbody td{{padding:8px 9px!important;font-size:11px!important;}}
    .ticket-table th:nth-child(1),.ticket-table td:nth-child(1){{min-width:38px!important;width:38px!important;text-align:center!important;}}
    .ticket-table th:nth-child(2),.ticket-table td:nth-child(2){{min-width:105px!important;width:105px!important;}}
    .ticket-table th:nth-child(3),.ticket-table td:nth-child(3){{min-width:105px!important;width:105px!important;}}
    .ticket-table th:nth-child(4),.ticket-table td:nth-child(4){{min-width:112px!important;width:112px!important;}}
    .ticket-table th:nth-child(5),.ticket-table td:nth-child(5){{min-width:110px!important;width:110px!important;}}
    .ticket-table th:nth-child(6),.ticket-table td:nth-child(6){{min-width:82px!important;width:82px!important;}}
    .ticket-table th:nth-child(7),.ticket-table td:nth-child(7){{min-width:120px!important;width:120px!important;}}
    .ticket-table th:nth-child(8),.ticket-table td:nth-child(8){{min-width:78px!important;width:78px!important;text-align:center!important;}}
    .ticket-table th:nth-child(9),.ticket-table td:nth-child(9){{min-width:95px!important;width:95px!important;text-align:center!important;}}
    .ticket-table th:nth-child(10),.ticket-table td:nth-child(10){{min-width:145px!important;width:145px!important;text-align:center!important;font-variant-numeric:tabular-nums!important;}}
    .ticket-table th:nth-child(11),.ticket-table td:nth-child(11){{min-width:125px!important;width:125px!important;}}
    .ticket-table th:nth-child(12),.ticket-table td:nth-child(12){{min-width:75px!important;width:75px!important;text-align:center!important;}}
    .ticket-table .device-cell{{display:table-cell!important;min-width:105px!important;white-space:nowrap!important;overflow:visible!important;text-overflow:clip!important;}}
    .ticket-table .issue-category-chip{{display:inline-flex!important;align-items:center!important;justify-content:center!important;max-width:none!important;white-space:nowrap!important;overflow:visible!important;text-overflow:clip!important;padding:4px 7px!important;font-size:10px!important;}}
    .ticket-table .ticket-code,.ticket-table .dept-chip,.ticket-table .department-label,.ticket-table .handler-chip,.ticket-table .reporter-chip{{display:inline-block!important;max-width:none!important;white-space:nowrap!important;overflow:visible!important;text-overflow:clip!important;}}
    .ticket-table .date-cell{{min-width:145px!important;white-space:nowrap!important;overflow:visible!important;text-overflow:clip!important;font-size:10.5px!important;}}
    

    /* Approved no-scroll ticket layout: remove the action column and wrap values naturally. */
    .table-wrap{{width:100%!important;max-width:100%!important;overflow-x:hidden!important;overflow-y:visible!important;}}
    .ticket-table{{width:100%!important;min-width:0!important;max-width:100%!important;table-layout:fixed!important;border-collapse:collapse!important;}}
    .ticket-table thead th,.ticket-table tbody td{{display:table-cell!important;box-sizing:border-box!important;white-space:normal!important;overflow:visible!important;text-overflow:clip!important;overflow-wrap:anywhere!important;word-break:normal!important;min-width:0!important;}}
    .ticket-table thead th{{padding:9px 5px!important;font-size:11px!important;line-height:1.25!important;text-align:center!important;}}
    .ticket-table tbody td{{padding:8px 5px!important;font-size:11px!important;line-height:1.3!important;vertical-align:middle!important;}}
    .ticket-table th:nth-child(1),.ticket-table td:nth-child(1){{width:4%!important;text-align:center!important;}}
    .ticket-table th:nth-child(2),.ticket-table td:nth-child(2){{width:9%!important;}}
    .ticket-table th:nth-child(3),.ticket-table td:nth-child(3){{width:10%!important;}}
    .ticket-table th:nth-child(4),.ticket-table td:nth-child(4){{width:11%!important;}}
    .ticket-table th:nth-child(5),.ticket-table td:nth-child(5){{width:10%!important;}}
    .ticket-table th:nth-child(6),.ticket-table td:nth-child(6){{width:7%!important;}}
    .ticket-table th:nth-child(7),.ticket-table td:nth-child(7){{width:11%!important;}}
    .ticket-table th:nth-child(8),.ticket-table td:nth-child(8){{width:7%!important;text-align:center!important;}}
    .ticket-table th:nth-child(9),.ticket-table td:nth-child(9){{width:8%!important;text-align:center!important;}}
    .ticket-table th:nth-child(10),.ticket-table td:nth-child(10){{width:12%!important;text-align:center!important;font-variant-numeric:tabular-nums!important;}}
    .ticket-table th:nth-child(11),.ticket-table td:nth-child(11){{width:11%!important;}}
    .ticket-table .device-cell{{display:table-cell!important;white-space:normal!important;overflow:visible!important;text-overflow:clip!important;overflow-wrap:anywhere!important;min-width:0!important;}}
    .ticket-table .issue-category-chip{{display:inline-flex!important;align-items:center!important;justify-content:center!important;max-width:100%!important;box-sizing:border-box!important;white-space:normal!important;overflow:visible!important;text-overflow:clip!important;overflow-wrap:anywhere!important;word-break:normal!important;padding:4px 6px!important;font-size:10px!important;line-height:1.25!important;}}
    .ticket-table .issue-category-chip>span{{white-space:normal!important;overflow-wrap:anywhere!important;}}
    .ticket-table .ticket-code,.ticket-table .dept-chip,.ticket-table .department-label,.ticket-table .handler-chip,.ticket-table .reporter-chip{{display:inline-block!important;max-width:100%!important;box-sizing:border-box!important;white-space:normal!important;overflow:visible!important;text-overflow:clip!important;overflow-wrap:anywhere!important;word-break:normal!important;vertical-align:middle!important;}}
    .ticket-table .date-cell{{white-space:normal!important;overflow:visible!important;text-overflow:clip!important;overflow-wrap:normal!important;word-break:normal!important;font-size:10.5px!important;line-height:1.3!important;min-width:0!important;}}
    .ticket-table td.priority-cell,.ticket-table td.status-cell{{white-space:normal!important;text-align:center!important;padding:6px 3px!important;}}
    .ticket-table td.priority-cell .pill,.ticket-table td.status-cell .pill,.ticket-table td.priority-cell .pill-label,.ticket-table td.status-cell .pill-label{{display:inline-flex!important;align-items:center!important;justify-content:center!important;white-space:normal!important;overflow:visible!important;text-overflow:clip!important;overflow-wrap:anywhere!important;word-break:normal!important;font-size:10px!important;line-height:1.2!important;padding:4px 5px!important;min-height:0!important;max-width:100%!important;}}
    .ticket-table td.priority-cell .pill-icon,.ticket-table td.status-cell .pill-icon{{display:none!important;}}
    
    
    /* Compact, aligned topbar: ticket alert and a single combined integration-status control. */
    .brand{{gap:10px;flex-shrink:0}}
    .brand strong{{font-size:19px;letter-spacing:.035em;color:#123d70;white-space:nowrap}}
    .topbar-right{{gap:9px;flex-shrink:0}}
    .top-notification{{flex:0 0 38px;width:38px;height:38px;box-sizing:border-box;gap:0;color:#1769aa!important;line-height:1}}
    .top-notification svg{{display:block;flex:0 0 auto}}
    .integration-status-link{{display:inline-flex;align-items:center;gap:8px;max-width:390px;min-width:0;flex:0 1 auto;padding:4px 9px;border:1px solid #dce9f6;border-radius:11px;background:linear-gradient(135deg,#f9fcff,#f1f7ff);text-decoration:none!important;color:#36577f!important;box-shadow:0 2px 5px rgba(26,82,133,.035)}}
    .integration-status-icon{{display:inline-flex;align-items:center;justify-content:center;flex:0 0 29px;width:29px;height:29px;border-radius:9px;background:#e4f1ff;color:#1477bd}}
    .integration-status-providers{{display:flex;flex-direction:row;align-items:center;gap:8px;min-width:0;flex-wrap:wrap}}
    .provider-status{{display:inline-flex;align-items:center;gap:4px;font-size:10px;line-height:1.2;white-space:nowrap}}
    .provider-status b{{font-size:10px;min-width:0;color:#355475}}
    .provider-status.connected span{{color:#16805d}}
    .provider-status.disconnected span{{color:#a15d28}}
    .settings-hero{{display:flex;align-items:flex-start;gap:13px;margin-bottom:14px}}
    .settings-hero-icon{{display:flex;align-items:center;justify-content:center;flex:0 0 44px;width:44px;height:44px;border-radius:13px;background:#e5f2ff;color:#0878d1}}
    .settings-hero h2{{margin:0 0 5px!important}}
    .settings-hero p{{margin:0;color:#6982a2;font-size:13px;line-height:1.5}}
    .integration-grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin:14px 0}}
    .integration-card{{border:1px solid #dfeaf6;border-radius:13px;background:#fff;padding:16px;min-width:0}}
    .integration-card-head{{display:flex;align-items:center;justify-content:space-between;gap:10px;padding-bottom:12px;margin-bottom:12px;border-bottom:1px solid #edf2f8}}
    .integration-card-title{{display:flex;align-items:center;gap:9px;font-size:15px;font-weight:800;color:#24496f}}
    .integration-card-mark{{display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:10px;background:#edf6ff;color:#1779ba;font-size:16px}}
    .connection-badge{{display:inline-flex;align-items:center;gap:5px;padding:5px 8px;border-radius:999px;font-size:11px;font-weight:750;white-space:nowrap}}
    .connection-badge.connected{{background:#e7f8ef;color:#167449}}
    .connection-badge.disconnected{{background:#fff4e8;color:#a45b19}}
    .integration-card label{{margin-top:9px}}
    .integration-card input[type=text],.integration-card input[type=url],.integration-card input[type=password]{{width:100%;min-width:0}}
    .integration-help{{margin:8px 0 0;color:#7186a1;font-size:11px;line-height:1.5}}
    .integration-check{{display:flex;align-items:flex-start;gap:8px;margin:10px 0!important;font-weight:650!important;color:#365779}}
    .integration-check input{{min-height:auto;margin-top:3px;flex:0 0 auto}}
    .integration-small-actions{{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:13px}}
    .integration-small-actions button,.integration-small-actions .btn{{font-size:12px;padding:8px 11px}}
    .integration-clear{{font-size:11px!important;font-weight:500!important;color:#8a5260!important}}
    .notification-events{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px 14px}}
    .notification-events .integration-check{{padding:10px;border:1px solid #e5edf6;border-radius:9px;background:#f9fbfe;margin:0!important}}
    .settings-save-row{{display:flex;justify-content:flex-end;align-items:center;gap:10px;flex-wrap:wrap;margin-top:14px}}
    .settings-message{{padding:11px 13px;border-radius:9px;background:#edf8f2;border:1px solid #d0eddd;color:#256b4b;font-size:13px;margin:12px 0}}
    .settings-message.error{{background:#fff5f2;border-color:#f4d5cc;color:#9f3d2d}}
    .integration-last-check{{font-size:11px;color:#7186a1;line-height:1.5;margin-top:10px;overflow-wrap:anywhere}}
    
    


    /* Current UI refinements: compact integration dropdown, hotline support, settings tabs, stronger priority. */
    .integration-dropdown{{position:relative;display:inline-block;min-width:0;z-index:30}}
    .integration-dropdown>summary{{list-style:none}}
    .integration-dropdown>summary::-webkit-details-marker{{display:none}}
    .integration-dropdown-trigger{{display:inline-flex;align-items:center;gap:7px;min-height:38px;max-width:250px;padding:4px 8px;border:1px solid #dce9f6;border-radius:11px;background:linear-gradient(135deg,#f9fcff,#f1f7ff);color:#36577f;cursor:pointer;box-shadow:0 2px 5px rgba(26,82,133,.035)}}
    .integration-dropdown-trigger:hover,.integration-dropdown[open] .integration-dropdown-trigger{{background:#eaf4ff;border-color:#bddaf5}}
    .integration-dropdown-trigger .integration-status-icon{{flex:0 0 28px;width:28px;height:28px}}
    .integration-dropdown-summary{{display:flex;min-width:0;flex-direction:column;align-items:flex-start;gap:2px;line-height:1.2}}
    .integration-dropdown-summary strong{{font-size:11px;color:#27496e}}
    .integration-summary-state{{font-size:10px;font-weight:700;white-space:nowrap;color:#9a5a20}}
    .integration-summary-state.connected{{color:#14764c}}.integration-summary-state.partial{{color:#986516}}.integration-summary-state.disconnected{{color:#9b3b39}}
    .integration-dropdown-chevron{{flex:0 0 auto;transition:transform .15s ease}}
    .integration-dropdown[open] .integration-dropdown-chevron{{transform:rotate(180deg)}}
    .integration-dropdown-menu{{position:absolute;right:0;top:calc(100% + 8px);width:min(330px,calc(100vw - 24px));padding:8px;background:#fff;border:1px solid #dce8f4;border-radius:13px;box-shadow:0 12px 34px rgba(27,61,99,.16);z-index:80}}
    .integration-menu-item{{display:flex;align-items:center;gap:9px;padding:10px 9px;color:#24496f!important;text-decoration:none!important;border-radius:9px}}
    .integration-menu-item:hover,.integration-menu-settings:hover{{background:#f2f7fd}}
    .integration-menu-service-icon{{display:inline-flex;align-items:center;justify-content:center;flex:0 0 30px;width:30px;height:30px;border-radius:9px;background:#edf5fd;color:#2975b3}}
    .integration-menu-copy{{display:flex;flex:1;min-width:0;flex-direction:column;gap:3px}}
    .integration-menu-copy strong{{font-size:12px;color:#24496f}}
    .integration-menu-copy small{{font-size:10px;color:#7186a1}}
    .integration-menu-state{{display:inline-flex;align-items:center;gap:4px;font-size:10px;font-weight:700;white-space:nowrap}}
    .integration-menu-state i{{width:7px;height:7px;border-radius:50%;display:inline-block;background:#c56a50}}
    .integration-menu-state.connected{{color:#167449}}.integration-menu-state.connected i{{background:#17a56b}}
    .integration-menu-state.disconnected{{color:#a45b19}}
    .integration-menu-settings{{display:flex;align-items:center;gap:8px;padding:10px;border-top:1px solid #edf2f8;border-radius:8px;color:#1769aa!important;text-decoration:none!important;font-size:12px;font-weight:700}}
    .integration-menu-settings span:last-child{{margin-left:auto;font-size:18px}}
    .settings-tabs{{display:flex;gap:8px;flex-wrap:wrap;padding:6px;margin:14px 0 18px;background:#f4f8fc;border:1px solid #e2ebf5;border-radius:12px}}
    .settings-tab{{display:inline-flex;align-items:center;justify-content:center;gap:8px;padding:10px 14px;border:1px solid transparent;border-radius:9px;color:#46627f!important;text-decoration:none!important;font-size:13px;font-weight:700}}
    .settings-tab:hover{{background:#eaf3fc}}
    .settings-tab.active{{background:#fff;border-color:#cfe2f7;color:#126eb2!important;box-shadow:0 2px 5px rgba(23,91,150,.06)}}
    .settings-tab svg{{width:16px;height:16px;flex-shrink:0}}
    .hotline-support-card{{display:flex;align-items:center;width:100%;box-sizing:border-box;padding:12px 11px;border:1px solid #b9dcf6;border-radius:14px;background:linear-gradient(145deg,#eaf5ff 0%,#ffffff 68%,#f0f8ff 100%);box-shadow:0 6px 17px rgba(26,87,135,.13);text-align:left}}
    .user-hotline-banner{{display:flex;align-items:center;gap:9px;min-width:0;max-width:360px;margin:0 auto;padding:7px 12px;border:1px solid transparent;border-radius:12px;background:transparent;box-shadow:none;flex:0 1 360px}}
    .user-hotline-banner .hotline-it-icon{{flex:0 0 32px;width:32px;height:32px}}
    .user-hotline-banner .hotline-support-copy{{gap:2px}}
    .user-hotline-banner .hotline-support-copy strong{{font-size:12px;line-height:1.25}}
    .user-hotline-banner .hotline-contact-row{{gap:4px;line-height:1.25}}
    .user-hotline-banner .hotline-contact-label{{flex:0 0 auto;font-size:10px}}
    .user-hotline-banner .hotline-value{{font-size:10px;letter-spacing:-.1px}}
    .user-hotline-banner .hotline-contact-row:first-of-type .hotline-value{{font-size:10px;font-weight:800;letter-spacing:-.1px}}
    .user-hotline-banner .hotline-contact-row .hotline-value{{overflow:visible;text-overflow:clip}}
    .user-hotline-banner .hotline-support-copy{{min-width:0}}

    .user-hotline-footer{{align-items:stretch!important;padding:8px 2px 16px!important;gap:0!important;transform:translateY(-9px)}}
    .hotline-it-icon{{display:inline-flex;align-items:center;justify-content:center;flex:0 0 32px;width:32px;height:32px;border-radius:50%;background:linear-gradient(145deg,#148fe0,#1760a5);color:#fff!important;box-shadow:0 3px 8px rgba(8,95,164,.2);border:1px solid rgba(255,255,255,.9)}}
    .hotline-support-copy{{display:flex;flex:1;min-width:0;width:100%;flex-direction:column;gap:6px}}
    .hotline-support-copy strong{{font-size:15px;font-weight:800;color:#174b7b;line-height:1.35;white-space:nowrap;letter-spacing:.1px}}
    .hotline-contact-row{{display:flex;align-items:baseline;gap:3px;min-width:0;width:100%;white-space:nowrap;line-height:1.5}}
    .hotline-contact-label{{flex:0 0 43px;color:#52718f;font-size:10px;font-weight:750;white-space:nowrap;letter-spacing:-.1px}}
    .hotline-value{{display:block;flex:1;min-width:0;overflow:hidden;text-overflow:clip;white-space:nowrap;color:#315b83;text-decoration:none!important;font-size:11px;font-weight:650;line-height:1.5;font-variant-numeric:tabular-nums;letter-spacing:-.25px}}
    .hotline-contact-row:first-of-type .hotline-value{{color:#245b8f;font-size:12px;font-weight:800;letter-spacing:-.3px}}
    .hotline-not-configured{{font-size:9px;color:#8a99a9!important;font-weight:550}}
    .hotline-email{{display:block;color:#52718f;font-size:10px;font-weight:650;line-height:1.5;overflow:hidden;text-overflow:clip;white-space:nowrap;letter-spacing:-.3px}}
    .hotline-not-configured{{font-size:10px;color:#8a99a9!important;font-weight:600}}
    .hotline-config-card{{max-width:760px}}
    .hotline-input-row{{display:grid;grid-template-columns:minmax(0,1fr);gap:6px;max-width:500px}}
    .hotline-input-row input{{width:100%;min-height:44px}}
    .hotline-preview{{display:flex;align-items:center;gap:12px;margin:18px 0;padding:13px;border:1px solid #c9e5f8;border-radius:12px;background:linear-gradient(135deg,#f2f9ff,#fffaf5);max-width:460px}}
    .hotline-preview .hotline-it-icon{{flex-basis:32px;width:32px;height:32px}}
    .hotline-preview-copy{{display:flex;flex-direction:column;gap:5px}}
    .hotline-preview-copy strong{{font-size:13px;color:#184c7f}}
    .hotline-preview-copy .hotline-number{{font-size:14px;font-weight:850;color:#245b8f;overflow-wrap:anywhere}}
    .hotline-save-actions{{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-top:18px}}
    .ticket-table .sev-high,.ticket-table td.priority-cell .sev-high,.pill.sev-high,.priority-pill.sev-high{{background:#ffdfe2!important;color:#b42332!important;border:1px solid #f1a6ae!important;box-shadow:none!important;font-weight:750!important}}
    /* Unified severity icon family: consistent minimalist signal bars for High/Medium/Low. */
    .severity-icon-svg{{display:inline-block!important;width:14px!important;height:14px!important;flex:0 0 14px!important;vertical-align:-2px!important;overflow:visible!important}}
    .priority-shield{{display:inline-flex!important;align-items:center!important;justify-content:center!important;font-size:0!important;line-height:1!important}}
    .priority-shield .severity-icon-svg{{width:15px!important;height:15px!important;flex-basis:15px!important}}
    .pill-icon .severity-icon-svg{{width:13px!important;height:13px!important;flex-basis:13px!important}}
    .sev-high{{background:#ffdfe2!important;color:#b42332!important;border:1px solid #f1a6ae!important}}
    .integration-connection-group{{display:inline-flex;align-items:center;gap:5px;min-width:0}}
    .integration-teams-primary{{display:inline-flex;align-items:center;gap:6px;min-height:34px;padding:4px 8px;border:1px solid #dce9f6;border-radius:10px;background:linear-gradient(135deg,#f9fcff,#f1f7ff);color:#36577f!important;text-decoration:none!important;white-space:nowrap}}
    .integration-teams-primary:hover{{background:#eaf4ff;border-color:#bddaf5}}
    .integration-teams-mark{{display:inline-flex;align-items:center;justify-content:center;flex:0 0 24px;width:24px;height:24px;border-radius:7px;background:#e5f1ff;color:#2870aa}}
    .integration-primary-copy{{display:flex;flex-direction:column;align-items:flex-start;gap:2px;line-height:1.1}}
    .integration-primary-copy strong{{font-size:10px;color:#27496e;white-space:nowrap}}
    .integration-primary-state{{font-size:9px;font-weight:750;white-space:nowrap}}
    .integration-primary-state.connected{{color:#14764c}}.integration-primary-state.disconnected{{color:#9b3b39}}
    .integration-dropdown-trigger{{display:inline-flex;align-items:center;justify-content:center;gap:4px;width:37px;min-width:37px;height:34px;min-height:34px;padding:0;border:1px solid #dce9f6;border-radius:10px;background:#f7fbff;color:#36577f;cursor:pointer}}
    .integration-dropdown-trigger:hover,.integration-dropdown[open] .integration-dropdown-trigger{{background:#eaf4ff;border-color:#bddaf5}}
    .integration-telegram-trigger{{display:inline-flex;align-items:center;justify-content:center;color:#2478b9}}
    .integration-dropdown-menu{{width:min(260px,calc(100vw - 24px));padding:6px;border-radius:11px;top:calc(100% + 5px)}}
    .integration-menu-item{{gap:7px;padding:7px 6px;border-radius:8px}}
    .integration-menu-service-icon{{flex-basis:26px;width:26px;height:26px;border-radius:7px}}
    .integration-menu-copy{{gap:2px}}
    .integration-menu-copy strong{{font-size:11px}}
    .integration-menu-copy small{{font-size:9px}}
    .integration-menu-state{{font-size:9px;gap:3px}}
    .integration-menu-state i{{width:6px;height:6px}}
    .integration-menu-settings{{padding:7px 6px;font-size:11px}}
    .admin-sidebar-footer{{display:flex;align-items:center;justify-content:center;min-height:26px}}
    .admin-sidebar-footer small{{font-size:10px;color:#88a0ba;font-weight:700;letter-spacing:.06em}}
    
    
    


    /* Unified integration dropdown: a single Teams-shaped control with compact Teams/Telegram rows. */
    .integration-header-dropdown{{position:relative;display:inline-block;flex:0 0 auto;min-width:0;z-index:85}}
    .integration-header-dropdown>summary{{list-style:none}}
    .integration-header-dropdown>summary::-webkit-details-marker{{display:none}}
    .integration-header-trigger{{display:inline-flex;align-items:center;gap:7px;min-height:36px;max-width:210px;box-sizing:border-box;padding:4px 8px;border:1px solid #dce9f6;border-radius:10px;background:linear-gradient(135deg,#f9fcff,#f1f7ff);color:#36577f;cursor:pointer;box-shadow:0 2px 5px rgba(26,82,133,.035)}}
    .integration-header-trigger:hover,.integration-header-dropdown[open]>.integration-header-trigger{{background:#eaf4ff;border-color:#bddaf5}}
    .integration-header-main-icon{{display:inline-flex;align-items:center;justify-content:center;flex:0 0 25px;width:25px;height:25px;border-radius:8px;background:#e4f1ff;color:#2870aa}}
    .integration-header-main-icon svg,.integration-header-item-icon svg{{display:block;width:16px;height:16px}}
    .integration-header-main-copy{{display:flex;flex-direction:column;align-items:flex-start;gap:3px;min-width:0;line-height:1.05}}
    .integration-header-main-copy strong{{font-size:10px;white-space:nowrap;color:#27496e}}
    .integration-header-state{{display:inline-flex;align-items:center;gap:4px;font-size:9px;font-weight:750;line-height:1.2;white-space:nowrap}}
    .integration-header-state i{{display:inline-block;width:6px;height:6px;flex:0 0 6px;border-radius:50%;background:#c56a50}}
    .integration-header-state.connected{{color:#14764c}}.integration-header-state.connected i{{background:#17a56b}}
    .integration-header-state.disconnected{{color:#9b3b39}}.integration-header-state.disconnected i{{background:#c45454}}
    .integration-header-chevron{{flex:0 0 auto;transition:transform .15s ease;color:#577797}}
    .integration-header-dropdown[open] .integration-header-chevron{{transform:rotate(180deg)}}
    .integration-header-menu{{position:absolute;top:calc(100% + 4px);right:0;width:190px;max-width:calc(100vw - 18px);box-sizing:border-box;padding:4px;background:#fff;border:1px solid #dce8f4;border-radius:10px;box-shadow:0 8px 20px rgba(27,61,99,.13);z-index:90}}
    .integration-header-item{{display:flex;align-items:center;gap:7px;min-height:33px;box-sizing:border-box;padding:5px 6px;border-radius:7px;color:#24496f!important;text-decoration:none!important}}
    .integration-header-item:hover{{background:#f2f7fd}}
    .integration-header-item+.integration-header-item{{margin-top:2px}}
    .integration-header-item-icon{{display:inline-flex;align-items:center;justify-content:center;flex:0 0 26px;width:26px;height:26px;border-radius:7px;background:#edf5fd;color:#2975b3}}
    .integration-header-item-icon.telegram{{background:#eaf6ff;color:#1689c9}}.integration-header-item-icon.telegram svg{{width:19px;height:19px}}
    .integration-header-item-copy{{display:flex;flex:1;min-width:0;flex-direction:column;align-items:flex-start;gap:3px;line-height:1.1}}
    .integration-header-item-copy strong{{font-size:10px;font-weight:750;color:#24496f;white-space:nowrap}}
    
    
    

    
    /* FMP Admin home: layout closely follows the user's supplied reference */
    .dashboard-ticket-filters{{display:grid;grid-template-columns:repeat(6,minmax(105px,1fr));gap:9px;margin:0 0 14px;padding:12px;background:#f7fbff;border:1px solid #e1edf9;border-radius:10px}}.dashboard-ticket-filters>div{{min-width:0}}.dashboard-ticket-filters label{{display:block;margin:0 0 5px;font-size:11px;font-weight:700;color:#55708e}}.dashboard-ticket-filters select,.dashboard-ticket-filters input{{box-sizing:border-box;width:100%;height:34px;min-width:0;padding:6px 8px;border:1px solid #d5e3f2;border-radius:7px;background:#fff;color:#173e6d;font-size:11px}}.dashboard-filter-actions{{display:flex;align-items:end;gap:7px}}.dashboard-filter-actions button,.dashboard-filter-actions a{{display:inline-flex;align-items:center;justify-content:center;min-height:34px;padding:0 11px;border:1px solid #cfe2f7;border-radius:7px;background:#eaf4ff;color:#1769aa;text-decoration:none;font-size:11px;font-weight:750;white-space:nowrap}}.dashboard-filter-actions a{{background:#fff;color:#657d96}}
    .admin-dashboard{{max-width:1500px;margin:0 auto;padding:0 2px 18px;color:#123b70}}.admin-welcome-banner{{display:flex;align-items:center;gap:20px;min-height:110px;padding:18px 24px;margin:0 0 14px;border-radius:13px;background:linear-gradient(110deg,#eaf5ff,#dcefff 52%,#eff8ff);border:1px solid #d8eaff;position:relative;overflow:hidden}}.admin-welcome-banner:after{{content:"";position:absolute;right:5%;top:-70px;width:240px;height:240px;border-radius:50%;border:28px solid rgba(255,255,255,.28);pointer-events:none}}.welcome-art{{display:flex;align-items:center;justify-content:center;flex:0 0 78px;height:70px;color:#168be8;background:rgba(255,255,255,.6);border:2px solid #9bcfff;border-radius:10px}}.welcome-copy{{min-width:0;flex:1}}.welcome-copy h1{{font-size:21px;line-height:1.3;margin:0 0 7px;color:#073f8b;font-weight:800}}.welcome-copy p{{font-size:14px;color:#38639b;margin:0}}.welcome-datetime{{position:relative;z-index:1;display:grid;gap:12px;min-width:145px;font-size:13px;font-weight:650;color:#12427d}}.admin-stat-grid{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin-bottom:16px}}.admin-stat-card{{display:flex;align-items:center;gap:15px;min-width:0;min-height:108px;padding:15px 17px;background:#fff;border:1px solid #dcecff;border-radius:12px;text-decoration:none;color:#103f80;box-shadow:0 3px 12px rgba(42,105,165,.04)}}.admin-stat-icon{{display:flex;align-items:center;justify-content:center;flex:0 0 54px;width:54px;height:54px;border-radius:18px;background:#e4f4ff;color:#168be8}}.stat-progress .admin-stat-icon{{background:#f0eaff;color:#8055df}}.stat-closed .admin-stat-icon{{background:#e9f8df;color:#399a30}}.admin-stat-copy{{display:flex;flex-direction:column;gap:4px;min-width:0}}.admin-stat-copy>span{{font-size:12px;color:#234c7e}}.admin-stat-copy strong{{font-size:26px;line-height:1.05;color:#073f8b;font-weight:800}}.admin-stat-copy small{{font-size:10px;color:#6586aa}}.admin-dashboard-columns{{display:grid;grid-template-columns:minmax(0,1fr);gap:16px;align-items:start}}.admin-recent-panel{{min-width:0;margin:0;padding:16px 14px 12px;border-color:#dcecff;border-radius:13px}}.admin-panel-heading{{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:14px}}.admin-panel-heading h2{{display:flex;align-items:center;gap:8px;margin:0;font-size:16px;color:#102f60}}.dashboard-heading-icon{{flex:0 0 auto;color:#168be8}}.admin-refresh{{display:inline-flex;align-items:center;gap:5px}}.admin-panel-actions{{display:flex;align-items:center;gap:8px}}.admin-panel-actions .btn{{font-size:12px;padding:8px 12px}}.admin-refresh{{font-size:12px;text-decoration:none;color:#6283a9;white-space:nowrap}}.admin-recent-table{{overflow:visible;min-width:0;width:100%}}.admin-recent-table .table-wrap{{overflow:visible;max-width:100%}}.admin-recent-table .ticket-table{{width:100%;min-width:0;table-layout:auto}}.admin-recent-table .ticket-table thead th{{padding:10px 8px;font-size:11px;text-transform:none;letter-spacing:0;color:#14365f;background:#edf6ff;white-space:normal}}.admin-recent-table .ticket-table tbody td{{padding:10px 8px;font-size:12px;white-space:normal;overflow-wrap:anywhere}}.admin-recent-table .ticket-table .ticket-code{{white-space:normal}}.admin-panel-footer{{display:flex;justify-content:space-between;gap:10px;padding:12px 5px 0;color:#7890ad;font-size:11px}}.admin-panel-footer a{{color:#1c6fc2;text-decoration:none}}.admin-dashboard-aside{{display:grid;gap:14px;min-width:0}}.admin-dashboard-aside .box{{margin:0;padding:15px;border:1px solid #dcecff;border-radius:12px}}.admin-quick-create h2{{display:flex;align-items:center;gap:8px;margin:0 0 12px;font-size:13px;color:#153d70}}.aside-icon{{display:inline-flex;align-items:center;justify-content:center;width:23px;height:23px;border-radius:50%;background:#168be8;color:white;font-size:20px}}.admin-quick-create p,.admin-notification-card p{{margin:0 0 12px;font-size:11px;color:#6985a6;line-height:1.6}}.admin-action-link{{display:flex;align-items:center;justify-content:space-between;padding:10px 11px;border-radius:8px;background:linear-gradient(135deg,#168be8,#176ac8);color:white;text-decoration:none;font-size:12px;font-weight:700}}.quick-create-links{{display:grid;gap:10px;margin-top:15px}}.quick-create-links a{{font-size:11px;color:#31577f;text-decoration:none}}.admin-system-status h3,.admin-notification-card h3{{margin:0 0 12px;font-size:12px;color:#153d70}}.admin-system-status p{{margin:0;font-size:11px;color:#15803d;display:flex;align-items:center;gap:8px}}.system-online-dot{{width:9px;height:9px;border-radius:50%;background:#16a34a;box-shadow:0 0 0 4px #e4f8e9}}.admin-notification-card h3{{display:flex;align-items:center;gap:7px}}.teams-mini-mark{{display:inline-flex;align-items:center;justify-content:center;width:25px;height:25px;background:#ecebff;color:#5952b8;border-radius:7px;font-size:17px}}.admin-notification-card>a{{font-size:11px;color:#176fc3;text-decoration:none}}.admin-brand-card{{display:flex;align-items:flex-start;gap:12px;padding:16px 14px;border-radius:12px;background:linear-gradient(135deg,#e5f5ff,#f1f9ff);color:#17497f}}.admin-brand-heart{{display:flex;align-items:center;justify-content:center;flex:0 0 42px;height:42px;border-radius:50%;background:#cceaff;color:#168be8;font-size:23px}}.admin-brand-card strong{{font-size:11px}}.admin-brand-card p{{font-size:10px;line-height:1.6;margin:6px 0 0;color:#4676a8}}

    
    /* Refined admin dashboard icons, typography, and date/time alignment */
    .welcome-art{{flex:0 0 104px;width:104px;height:82px;border:0;border-radius:18px;background:transparent;box-shadow:none;color:#1684df;overflow:hidden}}
    .welcome-art svg{{width:100%;height:100%;max-width:104px;max-height:82px;filter:drop-shadow(0 3px 2px rgba(22,132,223,.13));stroke-width:2.2}}
    .welcome-copy h1{{font-size:23px;letter-spacing:-.35px;font-weight:800;color:#073f8b}}
    .welcome-copy p{{font-size:14px;line-height:1.55;color:#416b9c}}
    .welcome-datetime{{gap:10px;min-width:150px;font-size:12px;color:#17477f}}
    .welcome-date-item{{display:flex;align-items:center;gap:9px;white-space:nowrap}}
    .welcome-date-item svg{{color:#1679d2;flex:0 0 auto}}
    .welcome-datetime{{position:relative;z-index:1;display:grid;grid-template-columns:minmax(0,1fr);grid-auto-rows:minmax(30px,auto);align-content:center;gap:5px;min-width:275px;padding:0!important;margin:0;background:transparent!important;border:0!important;border-radius:0!important;box-shadow:none!important;color:#17477f}}
    .welcome-date-item{{display:grid;grid-template-columns:20px minmax(0,1fr);align-items:center;gap:10px;min-height:30px;margin:0;font-size:14px;font-weight:750;line-height:1.35;color:#123f78}}
    .welcome-date-item>svg{{display:block;width:20px;height:20px;flex:0 0 20px;color:#1679d2}}
    .welcome-date-control{{display:flex;align-items:center;gap:0;min-width:0;flex:1}}
    .welcome-date-text{{min-width:0;white-space:normal;font-size:14px;font-weight:800;font-variant-numeric:tabular-nums;line-height:1.35}}
    .welcome-calendar-button{{display:inline-flex;align-items:center;justify-content:center;flex:0 0 20px;width:20px;height:20px;min-height:20px;padding:0;border:0!important;border-radius:0!important;background:transparent!important;box-shadow:none!important;color:#1679d2;cursor:pointer}}
    .welcome-calendar-button:hover{{color:#075ca8;background:transparent!important}}
    .welcome-calendar-button svg{{display:block;width:20px;height:20px}}
    .welcome-hidden-date{{position:absolute;width:1px;height:1px;opacity:0;pointer-events:none}}
    .welcome-time-text{{font-size:16px;font-weight:850;letter-spacing:.3px;font-variant-numeric:tabular-nums;color:#073f8b;line-height:1.35}}
    
    .admin-stat-card{{position:relative;gap:14px;transition:transform .18s ease,box-shadow .18s ease,border-color .18s ease}}
    .admin-stat-card:hover{{transform:translateY(-2px);border-color:#b9dcff;box-shadow:0 7px 18px rgba(42,105,165,.10)}}
    .admin-stat-icon{{border-radius:17px;box-shadow:inset 0 0 0 1px rgba(255,255,255,.72)}}
    .stat-total .admin-stat-icon{{background:linear-gradient(145deg,#e4f5ff,#d5edff);color:#168be8}}
    .stat-new .admin-stat-icon{{background:linear-gradient(145deg,#e1f4ff,#cceaff);color:#087bd5}}
    .stat-progress .admin-stat-icon{{background:linear-gradient(145deg,#fff3d6,#ffdf9e);color:#b86a00}}
    .stat-closed .admin-stat-icon{{background:linear-gradient(145deg,#eaf8e3,#d8f0ca);color:#318b2d}}
    .admin-stat-copy>span{{font-size:12px;font-weight:700;color:#294b76;letter-spacing:.05px}}
    .admin-stat-copy strong{{font-size:28px;letter-spacing:-.5px;font-weight:800;color:#083f83;font-variant-numeric:tabular-nums}}
    .stat-new .admin-stat-copy strong{{color:#087bd5}}.stat-progress .admin-stat-copy strong{{color:#b86a00}}.stat-closed .admin-stat-copy strong{{color:#25833b}}
    .stat-new .admin-stat-copy>span{{color:#087bd5}}.stat-progress .admin-stat-copy>span{{color:#b86a00}}.stat-closed .admin-stat-copy>span{{color:#25833b}}
    .admin-stat-copy .stat-insight{{display:flex;align-items:center;gap:5px;font-size:10px;font-weight:750;line-height:1.2}}
    .admin-stat-copy .stat-insight em{{font-style:normal;font-weight:500;color:#8196ae;margin-left:2px}}
    .stat-insight.trend-up{{color:#16834a}}.stat-insight.trend-up svg{{color:#16834a}}
    .stat-insight.trend-down{{color:#d04444}}.stat-insight.trend-down svg{{color:#d04444}}
    .stat-insight.trend-flat{{color:#7188a3}}.stat-insight.trend-flat svg{{color:#7188a3}}
    .admin-stat-copy .stat-insight{{display:flex;align-items:center;gap:5px;font-size:10px;color:#6c89a8;line-height:1.2}}
    .stat-insight svg{{color:#2182cf;flex:0 0 auto}}.stat-insight .trend-symbol{{font-size:13px;font-weight:900;line-height:1}}
    
    
/* v14: one severity palette across all pages and controls */
.sev-high,.priority-pill.sev-high,.pill.sev-high,.priority-label.sev-high{{color:#b42332!important;font-weight:800!important}}
.sev-medium,.priority-pill.sev-medium,.pill.sev-medium,.priority-label.sev-medium{{color:#94610f!important;font-weight:800!important}}
.sev-low,.priority-pill.sev-low,.pill.sev-low,.priority-label.sev-low{{color:#167653!important;font-weight:800!important}}
select[name="severity"]{{font-weight:800!important}}
select[name="severity"] option{{font-weight:800!important}}

    /* v15: refined, centered User IT-support group */
    .topbar:has(.user-hotline-banner) .user-hotline-banner{{
      align-self:center!important; justify-content:center!important; align-items:center!important;
      gap:8px!important; width:max-content; max-width:min(380px,42vw); margin-left:auto!important; margin-right:auto!important;
      padding:4px 8px!important; background:transparent!important; border:0!important; box-shadow:none!important;
    }}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-it-icon{{
      flex:0 0 23px!important; width:23px!important; height:23px!important;
      min-width:23px!important; min-height:23px!important; border:0!important; box-shadow:none!important;
      background:transparent!important; color:#1769aa!important; border-radius:0!important;
    }}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-it-icon svg{{width:21px!important;height:21px!important;stroke-width:2.2!important}}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-support-copy{{
      display:flex!important; flex:0 1 auto!important; width:auto!important; min-width:0!important;
      align-items:flex-start!important; justify-content:center!important; gap:2px!important; text-align:left!important;
    }}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-support-copy strong{{
      color:#244b70!important; font-size:12px!important; line-height:1.2!important; font-weight:800!important; letter-spacing:.15px!important;
    }}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-contact-row{{gap:5px!important;line-height:1.25!important;align-items:baseline!important}}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-contact-label{{font-size:10px!important;font-weight:650!important;color:#657d93!important}}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-value,
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-contact-row:first-of-type .hotline-value{{
      color:#1268b3!important; font-size:15px!important; line-height:1.25!important; font-weight:850!important;
      letter-spacing:.15px!important; font-variant-numeric:tabular-nums!important; text-shadow:0 1px 0 rgba(255,255,255,.65);
    }}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-not-configured{{font-size:11px!important;color:#8293a3!important;font-weight:600!important}}
    
    /* v16: User banner typography and headset alignment only */
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-it-icon{{
      align-self:flex-start!important; margin-top:4px!important;
    }}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-support-copy strong{{
      font-size:14px!important; line-height:1.25!important; font-weight:850!important;
      letter-spacing:.2px!important; color:#244b70!important;
    }}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-contact-row:first-of-type .hotline-value{{
      font-size:12px!important; line-height:1.25!important; font-weight:750!important;
      letter-spacing:0!important; text-shadow:none!important;
    }}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-contact-row:nth-of-type(2) .hotline-contact-label,
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-contact-row:nth-of-type(2) .hotline-value{{
      font-size:10px!important; line-height:1.25!important; font-weight:600!important;
      letter-spacing:0!important; color:#657d93!important; text-shadow:none!important;
    }}
    
    .csat-card-counts{{display:block;font-size:10px;white-space:nowrap;color:#58718d;margin-top:2px}}
    .ticket-code-cell{{overflow:visible!important}}
    .ticket-csat-mini{{position:absolute;right:-5px;top:-5px;display:inline-flex;align-items:center;justify-content:center;min-width:14px;height:14px;padding:0 2px;border:1px solid #fff;border-radius:999px;background:#f8fafc;color:#b77908;font-size:8px;line-height:1;white-space:nowrap;font-weight:800;box-shadow:0 1px 3px rgba(15,23,42,.12);z-index:3;cursor:default}}
    .ticket-code-cell .ticket-code-rating-line{{display:block;min-width:0;line-height:1.4;overflow:visible}}
    .ticket-code-rating-line .ticket-code{{position:relative;display:inline-block;max-width:100%;white-space:nowrap;vertical-align:middle}}
    .ticket-code-cell .ticket-csat-mini .rating-face{{font-size:9px;line-height:1}}
    .ticket-csat-5,.ticket-csat-4{{color:#16815b;background:#e8f7ef}}.ticket-csat-3{{color:#9a6700;background:#fff5d9}}.ticket-csat-2,.ticket-csat-1{{color:#c2413b;background:#fff0ef}}.ticket-csat-pending{{color:#8b9aab}}.row-number{{white-space:nowrap}}
    .csat-modal-backdrop{{padding:12px;overflow-y:auto;align-items:center;background:rgba(15,23,42,.28);backdrop-filter:blur(1px)}}
    .csat-modal{{width:min(420px,100%);max-height:calc(100dvh - 24px);overflow-y:auto;box-sizing:border-box;padding:22px;border:1px solid #dbe7f3;box-shadow:0 14px 42px rgba(23,54,93,.18)}}
    .csat-modal h2{{margin:0 28px 8px 0;font-size:19px;line-height:1.35;color:#244c75}}
    .csat-modal p{{margin:8px 0;color:#5e7188;font-size:13px;line-height:1.5}}
    .csat-stars{{gap:5px;margin:14px 0}}.csat-star{{font-size:30px;padding:1px;color:#d6a12b}}.csat-star.selected{{color:#e3a008}}
    .csat-actions{{flex-wrap:wrap;gap:8px}}.csat-actions button{{font-size:13px;padding:8px 12px}}
    
    .csat-history-panel{{margin:14px 0;padding:14px 16px;border:1px solid #dbe8f5;border-radius:12px;background:linear-gradient(135deg,#fbfdff,#f4f9ff)}}.csat-history-panel h3{{margin:0 0 8px;font-size:14px;color:#244c75}}.csat-history-stars{{font-size:17px;font-weight:800;letter-spacing:1px}}.csat-history-comment{{margin:6px 0 0;white-space:pre-wrap;overflow-wrap:anywhere;color:#52677e}}.csat-history-time{{margin-top:6px;color:#71849a;font-size:11px}}.csat-summary-line{{margin:8px 0 0;color:#62758b;font-size:12px}}.csat-modal-backdrop{{position:fixed;inset:0;z-index:9999;display:flex;align-items:center;justify-content:center;padding:18px;background:rgba(15,23,42,.48);backdrop-filter:blur(3px);animation:csatFade .22s ease}}.csat-modal{{position:relative;width:min(440px,100%);background:var(--card,#fff);color:var(--text,#172033);border-radius:18px;padding:26px;box-shadow:0 24px 80px #0003;animation:csatPop .24s ease}}.csat-x{{position:absolute;right:12px;top:10px;border:0;background:transparent;font-size:26px;color:inherit;cursor:pointer}}.csat-stars{{display:flex;gap:8px;justify-content:center;margin:20px 0}}.csat-star{{font-size:34px;border:0;background:transparent;color:#e9a800;cursor:pointer;padding:2px}}.csat-star.selected{{color:#f5ad00}}.csat-modal textarea,.csat-public textarea{{width:100%;box-sizing:border-box;margin:8px 0 12px;padding:10px;border:1px solid #cbd5e1;border-radius:10px}}.csat-actions{{display:flex;justify-content:flex-end;gap:10px;margin-top:18px}}.csat-actions button{{border-radius:10px}}.csat-result{{margin:16px 0}}.csat-public{{max-width:620px;margin:24px auto}}.csat-user-layout{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;align-items:stretch;margin:14px 0}}.csat-user-layout .box{{margin:0;min-width:0}}.csat-user-summary{{background:linear-gradient(135deg,#f8fbff,#eef6ff)!important;border-color:#d6e6f6!important}}.csat-user-edit{{background:linear-gradient(135deg,#f8fcfa,#edf8f2)!important;border-color:#d5eadf!important}}.csat-history-panel{{background:linear-gradient(135deg,#f8fbff,#eef6ff)!important;border-color:#d6e6f6!important}}.csat-user-layout h3{{display:flex;align-items:center;gap:8px;margin:0 0 10px;color:#244c75;font-size:14px}}.csat-heading-icon{{display:inline-flex;align-items:center;justify-content:center;width:25px;height:25px;flex:0 0 25px;border-radius:8px;background:#e7f2ff;color:#d99a00;font-size:17px}}.csat-edit-icon{{background:#fff0cf;color:#a66a00}}.csat-user-layout .csat-stars{{margin:10px 0;gap:5px}}.csat-user-layout .csat-star{{font-size:27px}}.csat-user-layout label{{margin-top:8px}}.csat-user-layout textarea{{width:100%;box-sizing:border-box}}.csat-score{{font-size:12px;letter-spacing:0;color:#53677f}}.csat-not-rated{{color:#8b9aab;font-size:13px}}.csat-edit-hint{{font-size:12px;color:#71849a;margin:0 0 8px}}.csat-user-layout .csat-actions{{margin-top:10px}}.csat-user-layout .csat-actions button:disabled{{opacity:.55;cursor:not-allowed}}@keyframes csatFade{{from{{opacity:0}}to{{opacity:1}}}}@keyframes csatPop{{from{{opacity:0;transform:scale(.96) translateY(8px)}}to{{opacity:1;transform:scale(1) translateY(0)}}}}
    /* Final visual consistency pass: compact non-blocking rating tips and responsive CSAT dialog. */
    .ticket-code-cell{{overflow:visible!important}}
    .ticket-code-rating-line{{position:relative;overflow:visible!important}}
    .csat-hover-tooltip{{position:fixed;z-index:20000;max-width:min(210px,70vw);box-sizing:border-box;padding:6px 9px;border:1px solid #d7e5f4;border-radius:8px;background:rgba(255,255,255,.98);color:#244563;font-size:11px;font-weight:650;line-height:1.35;white-space:normal;overflow-wrap:anywhere;text-align:left;box-shadow:0 4px 12px rgba(23,54,93,.12);pointer-events:none;opacity:0;transform:translateY(2px);transition:opacity .12s ease,transform .12s ease}}
    .csat-hover-tooltip.is-visible{{opacity:1;transform:translateY(0)}}
    .ticket-code-cell .ticket-code-rating-line{{display:inline-flex!important;align-items:center!important;justify-content:flex-start!important;gap:3px!important;max-width:100%;vertical-align:middle;overflow:visible!important}}.ticket-code-rating-line .ticket-code{{position:static!important;display:inline-block;min-width:0;max-width:100%;vertical-align:middle}}.ticket-code-rating-line .ticket-csat-mini{{position:static!important;inset:auto!important;display:inline-flex!important;align-items:center!important;justify-content:center!important;flex:0 0 14px!important;min-width:14px!important;width:14px!important;height:14px!important;box-sizing:border-box!important;padding:0!important;margin:0!important;border-radius:50%!important;font-size:8px!important;line-height:1!important;vertical-align:middle;z-index:auto;box-shadow:0 1px 3px rgba(15,23,42,.10)}}
        /* V18: keep ticket IDs perfectly aligned; move the rating face into the existing row-number gutter. */
    .ticket-table td.row-number{{position:relative!important;width:24px!important;min-width:24px!important;max-width:24px!important;box-sizing:border-box!important;text-align:left!important;padding-left:7px!important;padding-right:3px!important;white-space:nowrap!important;overflow:visible!important}}
    .ticket-table td.row-number .row-number-value{{display:inline-block;position:relative;left:-2px;font-variant-numeric:tabular-nums}}
    .ticket-table td.row-number > .ticket-csat-mini{{position:absolute!important;left:auto!important;right:1px!important;top:50%!important;transform:translateY(-50%)!important;display:inline-flex!important;align-items:center!important;justify-content:center!important;width:14px!important;min-width:14px!important;height:14px!important;padding:0!important;margin:0!important;box-sizing:border-box!important;border-radius:50%!important;font-size:8px!important;line-height:1!important;z-index:2!important;box-shadow:0 1px 3px rgba(15,23,42,.10)!important;}}
    .ticket-table td.row-number .rating-face{{font-size:9px!important;line-height:1!important}}
    .ticket-table td.ticket-code-cell{{padding-left:10px!important;overflow:visible!important}}
    .ticket-table td.ticket-code-cell .ticket-code{{display:inline-block!important}}
    .ticket-csat-4,.ticket-csat-5{{background:#e8f7ef!important;color:#16815b!important;border-color:#c9ead8!important}}
    .ticket-csat-3{{background:#fff5d9!important;color:#9a6700!important;border-color:#f4df9a!important}}
    .ticket-csat-1,.ticket-csat-2{{background:#fff0ef!important;color:#c2413b!important;border-color:#f5d0cd!important}}
    .csat-history-panel,.csat-user-summary{{background:linear-gradient(135deg,#f8fbff,#eef6ff)!important;border-color:#d6e6f6!important;box-shadow:0 3px 12px rgba(34,76,117,.04)}}
    .csat-user-edit{{background:linear-gradient(135deg,#f8fcfa,#edf8f2)!important;border-color:#d5eadf!important;box-shadow:0 3px 12px rgba(29,105,73,.04)}}
    .csat-heading-icon{{background:#e7f2ff!important;color:#d99a00!important}}
    .csat-edit-icon{{background:#e6f5ed!important;color:#26845c!important}}
    .csat-modal-backdrop{{padding:12px;overflow-y:auto;align-items:center;background:rgba(15,23,42,.28);backdrop-filter:blur(1px)}}
    .csat-modal{{width:min(420px,100%);max-height:calc(100vh - 24px);max-height:calc(100dvh - 24px);overflow-y:auto;box-sizing:border-box;padding:22px;border:1px solid #dbe7f3;box-shadow:0 14px 42px rgba(23,54,93,.18)}}
    .csat-modal h2{{margin:0 28px 8px 0;font-size:19px;line-height:1.35;color:#244c75}}
    .csat-modal p{{margin:8px 0;color:#5e7188;font-size:13px;line-height:1.5}}
    .csat-stars{{gap:5px;margin:14px 0}}.csat-star{{font-size:30px;padding:1px;color:#d6a12b}}.csat-star.selected{{color:#e3a008}}
    .csat-actions{{flex-wrap:wrap;gap:8px}}.csat-actions button{{font-size:13px;padding:8px 12px}}
    
    /* FMP unified palette: neutral surfaces, blue actions, clear status and severity meanings */
    .workflow-panel{{background:#f8fafc!important;border:1px solid #e2e8f0!important;box-shadow:0 4px 14px rgba(27,67,112,.045)!important}}
    .workflow-card{{background:#fff!important;border-color:#e2e8f0!important;box-shadow:0 2px 5px rgba(15,23,42,.035)!important}}
    .workflow-close-action{{background:#f8fafc!important;border-color:#e2e8f0!important}}
    .workflow-close-action .workflow-card-title{{color:#334155!important}}
    .workflow-close-action .workflow-icon-green{{background:#dbeafe!important;color:#1d4ed8!important}}
    .workflow-panel .btn-take,.workflow-panel .btn-close,.btn-take,.btn-close,.workflow-actions .btn-take,.workflow-actions .btn-close{{background:#2563eb!important;color:#fff!important;border:1px solid #2563eb!important;border-radius:8px!important;box-shadow:0 2px 4px rgba(37,99,235,.14)!important;font-weight:750!important;transition:background .16s ease,border-color .16s ease,box-shadow .16s ease,transform .16s ease!important}}
    .workflow-panel .btn-take:hover,.workflow-panel .btn-close:hover,.btn-take:hover,.btn-close:hover,.workflow-actions .btn-take:hover,.workflow-actions .btn-close:hover{{background:#1d4ed8!important;border-color:#1d4ed8!important;box-shadow:0 3px 7px rgba(37,99,235,.2)!important;filter:none!important}}
    .workflow-panel .btn-take:focus-visible,.workflow-panel .btn-close:focus-visible{{outline:3px solid #bfdbfe!important;outline-offset:2px!important}}
    .workflow-complete{{background:#f8fafc!important;border-color:#e2e8f0!important}}
    .workflow-pending{{background:#fff!important;border-color:#e2e8f0!important}}
    .note-editor-panel{{background:#eff6ff!important;border-color:#bfdbfe!important}}
    .detail-field-note{{background:#eff6ff!important;border:1px solid #bfdbfe!important;color:#1e3a5f!important}}
    .detail-field-note strong,.detail-field-note .public-note-label{{color:#1d4ed8!important}}
    .detail-field-note .detail-icon{{background:#dbeafe!important;color:#2563eb!important}}
    .detail-field-note .note-value-chip{{color:#1e3a5f!important}}
    .detail-field-note .note-empty{{color:#64748b!important}}
    .detail-field-note .public-note-form textarea{{border-color:#bfdbfe!important;color:#1e3a5f!important;resize:none!important}}
    .detail-field-note .public-note-save{{background:#2563eb!important;color:#fff!important}}
    .detail-field-note .public-note-save:hover{{background:#1d4ed8!important}}
    .pill.st-new,.ticket-table td.status-cell .st-new,.ticket-table .st-new{{background:#dbeafe!important;color:#1d4ed8!important;border-color:#bfdbfe!important}}
    .pill.st-progress,.ticket-table td.status-cell .st-progress,.ticket-table .st-progress{{background:#fef3c7!important;color:#92400e!important;border-color:#fde68a!important}}
    .pill.st-closed,.ticket-table td.status-cell .st-closed,.ticket-table .st-closed{{background:#dcfce7!important;color:#166534!important;border-color:#bbf7d0!important}}
    .ticket-row-new td{{background:#f8fbff!important}}
    .ticket-row-progress td{{background:#fffcf2!important}}
    .ticket-row-closed td{{background:#f8fdf9!important}}
    .detail-time-progress,.detail-field-received_at{{background:#fffbeb!important;color:#92400e!important;border-color:#fde68a!important}}
    .detail-time-closed,.detail-field-closed_at{{background:#f0fdf4!important;color:#166534!important;border-color:#bbf7d0!important}}
    .pill.sev-low,.priority-pill.sev-low,.ticket-table td.priority-cell .sev-low,.ticket-table .sev-low{{background:#dbeafe!important;color:#1d4ed8!important;border:1px solid #bfdbfe!important}}
    .pill.sev-medium,.priority-pill.sev-medium,.ticket-table td.priority-cell .sev-medium,.ticket-table .sev-medium{{background:#fef3c7!important;color:#92400e!important;border:1px solid #fde68a!important}}
    .pill.sev-high,.priority-pill.sev-high,.ticket-table td.priority-cell .sev-high,.ticket-table .sev-high{{background:#fee2e2!important;color:#b42332!important;border:1px solid #fecaca!important}}
    .ticket-table tbody tr:hover td{{background:#eef6ff!important}}
    /* Final unified palette and fixed-size public handling note card */
    .btn-take,.btn-close{{background:#2563eb!important;color:#fff!important;border:1px solid #2563eb!important}}
    .btn-take:hover,.btn-close:hover{{background:#1d4ed8!important;border-color:#1d4ed8!important}}
    .workflow-panel{{background:#f8fafc!important;border-color:#e2e8f0!important}}
    .workflow-close-action{{background:#f8fafc!important;border-color:#e2e8f0!important}}
    .workflow-note-help{{display:block!important;color:#64748b!important;line-height:1.45!important;font-size:12px!important}}
    .st-new,.pill.st-new,.ticket-table .st-new{{background:#dbeafe!important;color:#1d4ed8!important;border-color:#bfdbfe!important}}
    .st-progress,.pill.st-progress,.ticket-table .st-progress{{background:#fef3c7!important;color:#92400e!important;border-color:#fcd34d!important}}
    .st-closed,.pill.st-closed,.ticket-table .st-closed{{background:#dcfce7!important;color:#166534!important;border-color:#bbf7d0!important}}
    .sev-low,.pill.sev-low,.priority-pill.sev-low,.priority-label.sev-low,.ticket-table .sev-low{{background:#dbeafe!important;color:#1d4ed8!important;border-color:#bfdbfe!important}}
    .sev-medium,.pill.sev-medium,.priority-pill.sev-medium,.priority-label.sev-medium,.ticket-table .sev-medium{{background:#fef3c7!important;color:#92400e!important;border-color:#fcd34d!important}}
    .sev-high,.pill.sev-high,.priority-pill.sev-high,.priority-label.sev-high,.ticket-table .sev-high{{background:#fee2e2!important;color:#991b1b!important;border-color:#fecaca!important}}
    .ticket-table td.status-cell .st-new,.ticket-table .st-new{{background:#dbeafe!important;color:#1d4ed8!important;border-color:#bfdbfe!important}}
    .ticket-table td.status-cell .st-progress,.ticket-table .st-progress{{background:#fef3c7!important;color:#92400e!important;border-color:#fcd34d!important}}
    .ticket-table td.status-cell .st-closed,.ticket-table .st-closed{{background:#dcfce7!important;color:#166534!important;border-color:#bbf7d0!important}}
    .ticket-table td.priority-cell .sev-low,.ticket-table .sev-low{{background:#dbeafe!important;color:#1d4ed8!important;border-color:#bfdbfe!important}}
    .ticket-table td.priority-cell .sev-medium,.ticket-table .sev-medium{{background:#fef3c7!important;color:#92400e!important;border-color:#fcd34d!important}}
    .ticket-table td.priority-cell .sev-high,.ticket-table .sev-high{{background:#fee2e2!important;color:#991b1b!important;border-color:#fecaca!important}}
    .ticket-row-progress td{{background:#fffcf2!important}}.ticket-row-progress td:first-child{{box-shadow:inset 3px 0 #f59e0b!important}}
    .user-ticket-progress{{background:linear-gradient(135deg,#fff 0%,#fffcf2 100%)!important;border-left-color:#f59e0b!important}}
    .detail-field-note{{grid-column:span 3!important;box-sizing:border-box!important;min-height:104px!important;height:auto!important;display:flex!important;flex-direction:column!important;gap:7px!important;padding:10px 12px!important;background:#eff6ff!important;border:1px solid #bfdbfe!important;border-radius:11px!important;color:#1e3a5f!important;overflow:hidden!important}}
    .detail-field-note .note-heading{{display:flex!important;align-items:center!important;justify-content:space-between!important;gap:8px!important;flex:0 0 auto!important;min-width:0!important}}
    .detail-field-note .note-heading strong{{display:flex!important;align-items:center!important;gap:8px!important;margin:0!important;color:#1d4ed8!important;min-width:0!important}}
    .detail-field-note .detail-icon{{display:none!important}}
    .detail-field-note .note-heading strong{{font-size:12px!important;line-height:1.35!important}}
    .note-heading-actions{{display:flex!important;align-items:center!important;justify-content:flex-end!important;gap:4px!important;flex:0 0 auto!important}}
    .note-field-content{{margin:0!important;min-width:0!important;flex:1 1 auto!important;min-height:0!important;display:flex!important;flex-direction:column!important}}
    .public-note-view{{flex:1 1 auto!important;min-height:0!important;overflow:auto!important;white-space:pre-wrap!important;overflow-wrap:anywhere!important;line-height:1.5!important;color:#1e3a5f!important;padding:8px 10px!important;border:1px solid #dbeafe!important;border-radius:8px!important;background:rgba(255,255,255,.78)!important}}
    .note-value-chip{{display:block!important;white-space:pre-wrap!important;overflow-wrap:anywhere!important;color:#1e3a5f!important;font-weight:600!important;line-height:1.5!important}}
    .note-empty{{display:block!important;color:#64748b!important;font-size:12px!important;font-style:italic!important;font-weight:400!important}}
    .public-note-form{{display:flex!important;flex:1 1 auto!important;flex-direction:column!important;gap:7px!important;min-width:0!important;min-height:0!important}}
    .public-note-form textarea{{width:100%!important;min-width:0!important;max-width:100%!important;min-height:0!important;height:100%!important;flex:1 1 auto!important;box-sizing:border-box!important;resize:none!important;overflow:auto!important;border:1px solid #bfdbfe!important;background:#fff!important;border-radius:8px!important;padding:8px 10px!important;color:#1e3a5f!important;font:inherit!important;font-size:13px!important;line-height:1.5!important}}
    .public-note-actions{{display:flex!important;justify-content:flex-end!important;align-items:center!important;gap:6px!important;flex:0 0 auto!important}}
    .note-icon-btn{{display:inline-flex!important;align-items:center!important;justify-content:center!important;width:25px!important;height:25px!important;min-width:25px!important;min-height:25px!important;padding:0!important;border:1px solid #bfdbfe!important;border-radius:7px!important;background:#fff!important;color:#2563eb!important;cursor:pointer!important;box-shadow:none!important;flex:0 0 25px!important;transition:background .15s ease,color .15s ease,border-color .15s ease,opacity .15s ease!important}}
    .note-icon-btn svg{{width:13px!important;height:13px!important;fill:none!important;stroke:currentColor!important;stroke-width:2!important;stroke-linecap:round!important;stroke-linejoin:round!important}}
    .note-icon-btn:hover{{background:#dbeafe!important;color:#1d4ed8!important;border-color:#93c5fd!important}}
    .note-icon-btn.note-save{{background:#2563eb!important;color:#fff!important;border-color:#2563eb!important}}
    .note-icon-btn.note-cancel{{color:#64748b!important;border-color:#d7e2ee!important}}
    .note-icon-btn:disabled{{opacity:.42!important;cursor:default!important;box-shadow:none!important}}
    .note-icon-btn:disabled:hover{{background:#fff!important;color:inherit!important;border-color:#d7e2ee!important}}
    .note-icon-btn.note-save:hover{{background:#1d4ed8!important;border-color:#1d4ed8!important}}
    [hidden]{{display:none!important}}
    
    

    /* V10: shared login support banner + harmonized ticket labels. */
    main.login-main{{max-width:480px!important;margin:54px auto 30px!important;padding:0 16px!important}}
    .login-card{{padding:28px!important;border:1px solid #dce9f6!important;border-radius:20px!important;background:rgba(255,255,255,.98)!important;box-shadow:0 18px 48px rgba(24,75,125,.11)!important}}
    .login-title-mark{{width:58px!important;height:58px!important;border-radius:17px!important;margin:0 auto 14px!important;background:linear-gradient(145deg,#123f78,#087fc2)!important;color:#fff!important;box-shadow:0 8px 18px rgba(8,111,177,.20)!important}}
    .login-title-mark svg{{display:block;width:34px;height:34px}}
    .login-card h2{{text-align:center!important;color:#123d70!important;font-size:24px!important;font-weight:800!important;letter-spacing:-.3px}}
    .login-card>.msg{{text-align:center!important;min-height:18px}}
    .login-support{{display:flex;align-items:center;gap:12px;box-sizing:border-box;width:100%;margin:14px auto 0;padding:13px 16px;border:1px solid #d5e8f8;border-radius:14px;background:linear-gradient(135deg,#eaf5ff 0%,#fff 100%);box-shadow:0 6px 18px rgba(35,82,130,.06);color:#244d73}}
    .login-support-icon{{display:flex;align-items:center;justify-content:center;flex:0 0 38px;width:38px;height:38px;border-radius:12px;background:#123f78;color:#fff}}
    .login-support-icon svg{{width:25px;height:25px}}
    .login-support-content{{display:flex;flex:1;flex-direction:column;gap:4px;min-width:0}}
    .login-support-content>strong{{color:#174b7b;font-size:13px;font-weight:800;line-height:1.3}}
    .topbar:has(.user-hotline-banner){{justify-content:flex-start!important;gap:18px!important}}
    .topbar:has(.user-hotline-banner) .brand{{flex:0 0 auto}}
    .topbar:has(.user-hotline-banner) .user-hotline-banner{{margin:0!important;max-width:420px!important;flex:1 1 300px!important}}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-contact-label{{font-size:10px!important}}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-value{{font-size:11px!important;color:#315d83!important;font-weight:700!important}}
    .topbar:has(.user-hotline-banner) .user-hotline-banner .hotline-support-copy strong{{color:#174b7b!important}}
    @media(max-width:700px){{.topbar:has(.user-hotline-banner){{gap:8px!important}}.topbar:has(.user-hotline-banner) .user-hotline-banner{{order:3;flex-basis:100%!important;max-width:100%!important;padding:5px 0!important}}}}
    .login-support-line{{display:flex;align-items:baseline;gap:7px;min-width:0;font-size:11px;line-height:1.35}}
    .login-support-line span{{flex:0 0 auto;color:#69819a;font-weight:650}}
    .login-support-line strong{{min-width:0;color:#315d83;font-weight:700;overflow-wrap:anywhere}}
    .ticket-table td.priority-cell,.ticket-table td.status-cell{{white-space:nowrap!important;text-align:center!important}}
    .ticket-table td.priority-cell .pill,.ticket-table td.status-cell .pill,.ticket-table td.priority-cell .pill-label,.ticket-table td.status-cell .pill-label{{white-space:nowrap!important;overflow-wrap:normal!important;word-break:keep-all!important;flex-wrap:nowrap!important;line-height:1.15!important}}
    .ticket-table td.priority-cell .pill,.ticket-table td.status-cell .pill{{width:max-content!important;max-width:none!important;display:inline-flex!important;align-items:center!important;justify-content:center!important;padding:5px 8px!important}}
    .ticket-table td.priority-cell .pill-label,.ticket-table td.status-cell .pill-label{{display:inline!important}}
    @media(max-width:600px){{main.login-main{{margin:28px auto 22px!important}}.login-card{{padding:22px 18px!important}}.login-support{{padding:11px 13px}}}}
    /* V14: three-zone header alignment; support centered and language/control group at far right. */
    @media (min-width: 701px){{
      .topbar:has(.user-hotline-banner){{
        display:grid!important;
        grid-template-columns:minmax(0,1fr) minmax(260px,420px) minmax(0,1fr)!important;
        align-items:center!important;
        justify-content:initial!important;
        column-gap:16px!important;
      }}
      .topbar:has(.user-hotline-banner)>.brand{{grid-column:1;justify-self:start;min-width:0}}
      .topbar:has(.user-hotline-banner)>.user-hotline-banner{{
        grid-column:2;justify-self:center;width:100%;max-width:420px!important;
        min-width:0;margin:0!important;box-sizing:border-box;
      }}
      .topbar:has(.user-hotline-banner)>.topbar-right,
      .topbar:has(.user-hotline-banner)>.lang-switch{{
        grid-column:3;justify-self:end;display:flex;align-items:center;
        justify-content:flex-end;gap:9px;min-width:0;margin-left:0!important;
      }}
      .topbar:has(.user-hotline-banner)>.topbar-right>*{{flex-shrink:0;align-self:center}}
      .topbar:has(.user-hotline-banner)>.lang-switch{{width:max-content}}
    }}
    @media(max-width:700px){{
      .topbar:has(.user-hotline-banner){{display:flex!important;flex-wrap:wrap!important;align-items:center!important}}
      .topbar:has(.user-hotline-banner)>.brand{{order:1}}
      .topbar:has(.user-hotline-banner)>.topbar-right,
      .topbar:has(.user-hotline-banner)>.lang-switch{{order:2;margin-left:auto!important}}
      .topbar:has(.user-hotline-banner)>.user-hotline-banner{{order:3;flex:1 0 100%!important;max-width:100%!important;justify-self:stretch}}
    }}

    </style></head><body>
    {shell}
    <script>
    // Close every open dropdown/details menu when clicking outside; Escape also closes it.
    (function(){{
      function closeOpenMenus(except){{
        document.querySelectorAll('details[open]').forEach(function(menu){{
          if(menu !== except && !menu.contains(except)) menu.removeAttribute('open');
        }});
      }}
      document.addEventListener('click', function(event){{
        document.querySelectorAll('details[open]').forEach(function(menu){{
          if(!menu.contains(event.target)) menu.removeAttribute('open');
        }});
      }});
      document.addEventListener('keydown', function(event){{
        if(event.key === 'Escape'){{
          document.querySelectorAll('details[open]').forEach(function(menu){{menu.removeAttribute('open');}});
        }}
      }});
    }})();
    (function(){{
      let tip=null, active=null;
      function hideTip(){{if(tip){{tip.remove();tip=null}}active=null}}
      function showTip(el){{
        const text=el&&el.getAttribute('data-tip');if(!text)return;
        hideTip();active=el;tip=document.createElement('div');tip.className='csat-hover-tooltip';tip.setAttribute('role','tooltip');tip.textContent=text;document.body.appendChild(tip);
        const r=el.getBoundingClientRect(), tr=tip.getBoundingClientRect(), gap=6;
        let left=Math.max(8,Math.min(r.left,window.innerWidth-tr.width-8));
        let top=r.top-tr.height-gap;if(top<8)top=Math.min(window.innerHeight-tr.height-8,r.bottom+gap);
        tip.style.left=left+'px';tip.style.top=Math.max(8,top)+'px';requestAnimationFrame(()=>{{if(tip)tip.classList.add('is-visible')}});
      }}
      document.addEventListener('mouseover',function(e){{const el=e.target.closest('[data-tip]');if(el&&(!e.relatedTarget||!el.contains(e.relatedTarget)))showTip(el)}});
      document.addEventListener('mouseout',function(e){{if(active&&active.contains(e.target)&&(!e.relatedTarget||!active.contains(e.relatedTarget)))hideTip()}});
      document.addEventListener('focusin',function(e){{const el=e.target.closest('[data-tip]');if(el)showTip(el)}});
      document.addEventListener('focusout',function(e){{if(active&&active===e.target)hideTip()}});
      window.addEventListener('scroll',hideTip,true);window.addEventListener('resize',hideTip);
    }})();
    </script>
    <script>
    (function(){{
      const en = document.documentElement.lang === 'en';
      function toast(message, failed){{
        const old=document.getElementById('global-action-toast'); if(old)old.remove();
        const el=document.createElement('div'); el.id='global-action-toast'; el.setAttribute('role',failed?'alert':'status');
        el.style.cssText='position:fixed;z-index:100000;top:20px;right:20px;max-width:min(440px,calc(100vw - 40px));padding:13px 17px;border-radius:12px;box-shadow:0 12px 32px rgba(20,45,75,.2);font:600 13px/1.5 Inter,Segoe UI,Arial,sans-serif;border:1px solid;';
        el.style.background=failed?'#fff5f3':'#edf6ff';el.style.color=failed?'#a5281b':'#174f82';el.style.borderColor=failed?'#f3c9c2':'#bfdbf5';el.textContent=(failed?'⚠ ':'✓ ')+message;
        document.body.appendChild(el);window.setTimeout(()=>el.remove(),5500);
      }}
      try {{
        const pending=sessionStorage.getItem('fmpActionToast');
        if(pending){{
          sessionStorage.removeItem('fmpActionToast');
          if(!document.querySelector('.account-toast')){{
            try{{const saved=JSON.parse(pending);if(saved&&saved.message)toast(String(saved.message),Boolean(saved.failed));else toast(String(pending),false);}}
            catch(e){{toast(String(pending),false);}}
          }}
        }}
      }} catch(e) {{}}
      document.addEventListener('submit', async function(ev){{
        const form=ev.target;
        if(!(form instanceof HTMLFormElement) || (form.method||'get').toLowerCase()!=='post' || form.target || form.dataset.noGlobalToast==='1')return;
        ev.preventDefault();
        if(form.dataset.submitting==='1')return;
        form.dataset.submitting='1';
        const submitter=ev.submitter;
        const buttonLabel=(submitter && (submitter.innerText||submitter.value||submitter.getAttribute('aria-label')||submitter.title) || '').trim();
        if(submitter && 'disabled' in submitter)submitter.disabled=true;
        const payload=new FormData(form);
        if(submitter && submitter.name)payload.append(submitter.name,submitter.value||'');
        const actionUrl=new URL(form.action||location.href,location.href);
        const actionPath=actionUrl.pathname;
        const pendingLabel=en?'Processing your request…':'Đang xử lý yêu cầu…';
        toast(pendingLabel,false);
        try {{
          const opts={{method:'POST',body:payload,credentials:'same-origin',redirect:'follow'}};
          const response=await fetch(actionUrl.href,opts);
          if(!response.ok){{
            let detail='';
            try {{
              const raw=await response.text();
              try{{const payload=JSON.parse(raw);if(payload&&payload.detail)detail=String(payload.detail);}}
              catch(parseError){{const doc=new DOMParser().parseFromString(raw,'text/html');detail=(doc.body&&doc.body.innerText||raw).trim().replace(/\\s+/g,' ').slice(0,260);}}
            }} catch(e) {{}}
            if(response.status===405) detail=en?'This action is not supported at this address (HTTP 405). Refresh the page and try again.':'Địa chỉ thao tác không hỗ trợ phương thức gửi này (HTTP 405). Vui lòng tải lại trang rồi thử lại.';
            else if(response.status===401 || response.status===403) detail=en?'You do not have permission to perform this action, or your session has expired. Please sign in again.':'Bạn không có quyền thực hiện thao tác này hoặc phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.';
            else if(!detail)detail=(en?'The operation failed (HTTP ':'Thao tác thất bại (HTTP ')+response.status+'). '+(en?'Please check the information and try again.':'Vui lòng kiểm tra thông tin rồi thử lại.');
            toast(detail,true);form.dataset.submitting='0';if(submitter && 'disabled' in submitter)submitter.disabled=false;return;
          }}
          let message=''; let failureFlag=false;
          try {{
            const u=new URL(response.url,location.href); const code=u.searchParams.get('msg')||'';
            const known={{'User added':en?'User added successfully':'Thêm tài khoản thành công','User updated':en?'User updated successfully':'Cập nhật tài khoản thành công','Password reset':en?'Password reset successfully':'Đặt lại mật khẩu thành công','User locked':en?'Account locked successfully':'Đã khóa tài khoản thành công','User unlocked':en?'Account unlocked successfully':'Đã mở khóa tài khoản thành công','User deleted; tickets retained':en?'Account deleted; tickets were retained':'Đã xóa tài khoản; ticket được giữ lại','Passwordless enabled':en?'Passwordless login enabled':'Đã bật đăng nhập không cần mật khẩu','Passwordless disabled':en?'Passwordless login disabled':'Đã tắt đăng nhập không cần mật khẩu','saved':en?'Settings saved successfully':'Lưu cài đặt thành công','hotline_saved':en?'Support contact saved successfully':'Lưu thông tin hỗ trợ thành công','telegram_chat_found':en?'Telegram chat found successfully':'Tìm thấy Telegram Chat ID thành công','teams_test_ok':en?'Teams test message sent successfully':'Gửi tin nhắn kiểm tra Teams thành công','telegram_test_ok':en?'Telegram test message sent successfully':'Gửi tin nhắn kiểm tra Telegram thành công','discord_test_ok':en?'Discord test message sent successfully':'Gửi tin nhắn kiểm tra Discord thành công'}};
            const failures={{'Invalid username or password':en?'Invalid username or password':'Tên đăng nhập hoặc mật khẩu không đúng','Too many attempts':en?'Too many attempts. Please wait and try again.':'Quá nhiều lần thử. Vui lòng chờ rồi thử lại.','hotline_invalid':en?'Invalid support phone or email':'Số điện thoại hoặc email hỗ trợ không hợp lệ','invalid_teams_url':en?'Please enter a valid Teams webhook URL':'Vui lòng nhập Teams webhook URL hợp lệ','invalid_discord_url':en?'Please enter a valid Discord webhook URL':'Vui lòng nhập Discord webhook URL hợp lệ','telegram_chat_failed':en?'Could not find Telegram Chat ID':'Không tìm thấy Telegram Chat ID','teams_test_failed':en?'Teams test message failed':'Gửi tin nhắn kiểm tra Teams thất bại','telegram_test_failed':en?'Telegram test message failed':'Gửi tin nhắn kiểm tra Telegram thất bại','discord_test_failed':en?'Discord test message failed':'Gửi tin nhắn kiểm tra Discord thất bại'}};
            if(Object.prototype.hasOwnProperty.call(failures,code)){{message=failures[code];failureFlag=true;}}
            else if(code && (/failed|invalid|error/i.test(code))){{message=en?'The operation failed. Please check the information and try again.':'Thao tác thất bại. Vui lòng kiểm tra lại thông tin.';failureFlag=true;}}
            else if(u.pathname.endsWith('/login') && !form.action.endsWith('/login')){{message=en?'Your session expired or you are not authorized. Please log in again.':'Phiên đăng nhập đã hết hạn hoặc bạn không có quyền thực hiện. Vui lòng đăng nhập lại.';failureFlag=true;}}
            else if(code && /failed|invalid|error|not found|không|thất bại|lỗi|hết hạn/i.test(code)){{message=en?'The operation failed: '+code:'Thao tác thất bại: '+code;failureFlag=true;}}
            else message=known[code]||'';
          }} catch(e) {{}}
          if(!message){{
            const label=buttonLabel.toLowerCase();
            const finalUrl=new URL(response.url,location.href);
            const actionData=payload;
            const workflowAction=String(actionData.get('action')||'').toLowerCase();
            if(finalUrl.pathname.endsWith('/login') && !actionPath.endsWith('/login') && !actionPath.endsWith('/quick-login')){{message=en?'Your session expired or you are not authorized. Please sign in again.':'Phiên đăng nhập đã hết hạn hoặc bạn không có quyền thực hiện. Vui lòng đăng nhập lại.';failureFlag=true;}}
            else if(actionPath==='/new')message=en?'Ticket created successfully':'Tạo ticket thành công';
            
            else if((actionPath.startsWith('/ticket/') && actionPath.endsWith('/edit')))message=en?'Ticket updated successfully':'Cập nhật ticket thành công';
            else if((actionPath.startsWith('/ticket/') && actionPath.endsWith('/delete')))message=en?'Ticket deleted successfully':'Xóa ticket thành công';
            else if((actionPath.startsWith('/ticket/') && (actionPath.endsWith('/update') || actionPath.endsWith('/take') || actionPath.endsWith('/close')))){{
              if(actionPath.endsWith('/take') || workflowAction==='take')message=en?'Ticket accepted for processing':'Đã nhận xử lý ticket';
              else if(actionPath.endsWith('/close') || workflowAction==='close')message=en?'Ticket closed successfully':'Đóng ticket thành công';
              
              else {{message=en?'Ticket updated successfully':'Cập nhật ticket thành công';}}
            }}
            else if(((actionPath.startsWith('/ticket/') && actionPath.endsWith('/feedback')) || (actionPath.startsWith('/feedback/'))))message=en?'Rating submitted successfully':'Gửi đánh giá thành công';
            else if(actionPath==='/settings/hotline')message=en?'Support contact saved successfully':'Lưu thông tin hỗ trợ thành công';
            else if(actionPath==='/settings')message=en?'Notification settings saved successfully':'Lưu cài đặt thông báo thành công';
            else if(actionPath==='/settings/telegram/discover-chat'){{message=en?'Telegram Chat ID found':'Đã tìm thấy Telegram Chat ID';}}
            else if((actionPath.startsWith('/settings/test/') && ['/settings/test/teams','/settings/test/telegram','/settings/test/discord'].includes(actionPath))){{message=en?'Test message sent successfully':'Gửi tin nhắn kiểm tra thành công';}}
            else if(actionPath==='/users')message=en?'User added successfully':'Thêm tài khoản thành công';
            else if((actionPath.startsWith('/users/') && actionPath.endsWith('/edit')))message=en?'User updated successfully':'Cập nhật tài khoản thành công';
            else if((actionPath.startsWith('/users/') && actionPath.endsWith('/reset')))message=en?'Password reset successfully':'Đặt lại mật khẩu thành công';
            else if((actionPath.startsWith('/users/') && actionPath.endsWith('/toggle')))message=en?'Account status updated successfully':'Cập nhật trạng thái tài khoản thành công';
            else if((actionPath.startsWith('/users/') && actionPath.endsWith('/delete')))message=en?'Account deleted; tickets were retained':'Đã xóa tài khoản; ticket được giữ lại';
            else if((actionPath.startsWith('/users/') && actionPath.endsWith('/passwordless')))message=en?'Passwordless login setting updated':'Cập nhật đăng nhập không mật khẩu thành công';
            else if(/delete|xóa/.test(label))message=en?'Deleted successfully':'Xóa thành công';
            else if(/send|gửi|submit|tạo|create/.test(label))message=en?'Submitted successfully':'Gửi yêu cầu thành công';
            else if(/close|đóng/.test(label))message=en?'Updated successfully':'Cập nhật trạng thái thành công';
            else if(/save|lưu|update|cập nhật|edit|sửa|reset|đặt lại/.test(label))message=en?'Saved successfully':'Lưu thành công';
            else message=en?'Operation completed successfully':'Thao tác hoàn tất thành công';
          }}
          try{{if(new URL(response.url,location.href).pathname!=='/users')sessionStorage.setItem('fmpActionToast',JSON.stringify({{message:message,failed:failureFlag}}));}}catch(e){{}}
          if(response.redirected && response.url){{location.assign(response.url);return;}}
          toast(message,failureFlag);form.dataset.submitting='0';if(submitter && 'disabled' in submitter)submitter.disabled=false;
        }} catch(err){{
          toast(en?'Network error. The operation could not be confirmed.':'Lỗi kết nối. Chưa thể xác nhận thao tác đã thành công.',true);
          form.dataset.submitting='0';if(submitter && 'disabled' in submitter)submitter.disabled=false;
        }}
      }},true);
    }})();
    </script>
    </body></html>""")

def current_user(req: Request):
    uid = req.session.get("uid")
    u = query("SELECT * FROM users WHERE id=? AND active=1", (uid,), one=True) if uid else None
    if not u: raise HTTPException(303, headers={"Location": "/login"})
    return u

def admin_user(u=Depends(current_user)):
    if u["role"] != "admin": raise HTTPException(403, "Administrator access required")
    return u

@asynccontextmanager
async def lifespan(app):
    if not SECRET_KEY or len(SECRET_KEY) < 32:
        raise RuntimeError("Set SECRET_KEY to a random value of at least 32 characters in Docker environment.")
    if not ADMIN_PASS or ADMIN_PASS == "CHANGE_ME_NOW":
        raise RuntimeError("Set ADMIN_PASS in Docker environment.")
    if len(ADMIN_PASS) < 6:
        raise RuntimeError("ADMIN_PASS must be at least 6 characters.")
    init_db()
    yield

app = FastAPI(title="FMP IT Ticket System", lifespan=lifespan)
app.add_middleware(SessionMiddleware, secret_key=SECRET_KEY, same_site="lax",
                   https_only=os.getenv("HTTPS_ONLY", "false").lower() == "true", max_age=43200)
FAILS = {}

@app.get("/health")
def health():
    local_now = now_dt()
    utc_now = datetime.now(ZoneInfo("UTC"))
    return {
        "status": "ok",
        "timezone": "Asia/Ho_Chi_Minh",
        "server_time_vietnam": local_now.strftime("%d/%m/%Y %H:%M:%S %Z%z"),
        "server_time_iso": local_now.isoformat(),
        "server_time_utc": utc_now.isoformat(),
    }

@app.get("/lang/{lang}")
def set_lang(req: Request, lang: str):
    if lang in ("vi", "en"): req.session["lang"] = lang
    return RedirectResponse(req.headers.get("referer", "/"), 303)

@app.get("/login", response_class=HTMLResponse)
def login_form(req: Request, msg: str = ""):
    if msg == "Invalid username or password": msg = tr(req, "invalid_login")
    elif msg == "Too many attempts": msg = tr(req, "too_many")
    english = req.session.get("lang") == "en"
    password_tab = "Sign in with password" if english else "Đăng nhập với mật khẩu"
    passwordless_tab = "Sign in without password" if english else "Đăng nhập không cần mật khẩu"
    notice = "Passwordless access is available only for accounts enabled by an administrator." if english else "Đăng nhập không mật khẩu chỉ dùng được với tài khoản đã được quản trị viên cho phép."
    hotline_phone = get_integration_setting("hotline_phone", "").strip()
    hotline_email = get_integration_setting("hotline_email", "").strip()
    support_title = "IT Support" if english else "Hỗ trợ IT"
    support_missing = "Not configured" if english else "Chưa cập nhật"
    support_phone = E(hotline_phone) if hotline_phone else support_missing
    support_email = f'<div class="login-support-line"><span>Email</span><strong>{E(hotline_email)}</strong></div>' if hotline_email else ""
    headset_login_svg = '<svg viewBox="0 0 48 48" width="32" height="32" fill="none" stroke="currentColor" stroke-width="2.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 25v-4a16 16 0 0 1 32 0v4"/><path d="M8 24h7v12h-3a4 4 0 0 1-4-4zM40 24h-7v12h3a4 4 0 0 0 4-4z"/><path d="M33 37a11 11 0 0 1-9 5h-4"/></svg>'
    body = f'''<div class="box login-card"><div class="login-title-mark" aria-hidden="true">{headset_login_svg}</div><h2>{tr(req,"login")}</h2><p class="msg">{E(msg)}</p>
    <div class="login-tabs" role="tablist"><button type="button" class="login-tab active" id="tab-password" role="tab" aria-selected="true" onclick="setLoginMode('password')">{password_tab}</button><button type="button" class="login-tab" id="tab-passwordless" role="tab" aria-selected="false" onclick="setLoginMode('passwordless')">{passwordless_tab}</button></div>
    <form method="post" action="/login" id="login-form"><label>{tr(req,"username")}</label><input class="w login-input" name="username" required autocomplete="username" placeholder="{tr(req,'username')}"><div id="password-field"><label>{tr(req,"password")}</label><input class="w login-input" type="password" name="password" id="login-password" required autocomplete="current-password" placeholder="{tr(req,'password')}"></div><p class="login-submit-row"><button class="login-submit" type="submit" id="login-submit">{password_tab}</button></p></form>
    <p class="help login-notice">ⓘ {notice}</p></div>
    <script>function setLoginMode(mode){{const passwordMode=mode==='password';const form=document.getElementById('login-form');const field=document.getElementById('password-field');const password=document.getElementById('login-password');const a=document.getElementById('tab-password');const b=document.getElementById('tab-passwordless');form.action=passwordMode?'/login':'/quick-login';field.hidden=!passwordMode;password.required=passwordMode;if(!passwordMode)password.value='';a.classList.toggle('active',passwordMode);b.classList.toggle('active',!passwordMode);a.setAttribute('aria-selected',String(passwordMode));b.setAttribute('aria-selected',String(!passwordMode));document.getElementById('login-submit').textContent=passwordMode?{password_tab!r}:{passwordless_tab!r};}}</script>'''
    return page(req, tr(req, "login"), body)

@app.post("/login")
def login(req: Request, username: str = Form(...), password: str = Form(...)):
    key = username.strip().lower()
    n, last = FAILS.get(key, (0, 0))
    if n >= 5 and time.time() - last < 300:
        return RedirectResponse("/login?msg=Too+many+attempts", 303)
    u = query("SELECT * FROM users WHERE username=? AND active=1", (key,), one=True)
    if not u or not pw_check(password, u["password_hash"]):
        FAILS[key] = (n + 1 if time.time() - last < 300 else 1, time.time())
        return RedirectResponse("/login?msg=Invalid+username+or+password", 303)
    FAILS.pop(key, None)
    # Preserve the language selected on the login page. Previously this was
    # forcibly reset to Vietnamese after every successful login, which made
    # device and issue-type dropdowns appear untranslated even when EN was chosen.
    selected_lang = req.session.get("lang", "vi")
    if selected_lang not in ("vi", "en"):
        selected_lang = "vi"
    req.session.clear()
    req.session["uid"] = u["id"]
    req.session["lang"] = selected_lang
    return RedirectResponse("/", 303)

@app.post("/quick-login")
def quick_login(req: Request, username: str = Form(...)):
    key = username.strip().lower()
    u = query("SELECT * FROM users WHERE username=? AND active=1 AND passwordless_login=1 AND role='user'", (key,), one=True)
    if not u:
        return RedirectResponse("/login?msg=Invalid+username+or+password", 303)
    selected_lang = req.session.get("lang", "vi")
    req.session.clear()
    req.session["uid"] = u["id"]
    req.session["lang"] = selected_lang if selected_lang in ("vi", "en") else "vi"
    return RedirectResponse("/", 303)

@app.get("/logout")
def logout(req: Request):
    selected_lang = req.session.get("lang", "vi")
    req.session.clear()
    req.session["lang"] = selected_lang if selected_lang in ("vi", "en") else "vi"
    return RedirectResponse("/login", 303)

def ticket_pagination(req, total, page_num, per_page=10, all_total=None, show_all_count=True):
    """Compact pagination matching the FMP reference table footer."""
    pages = max(1, (total + per_page - 1) // per_page)
    page_num = max(1, min(int(page_num or 1), pages))
    params = dict(req.query_params)
    params.pop("page_num", None)
    def href(n):
        q = dict(params)
        if n > 1:
            q["page_num"] = str(n)
        else:
            q.pop("page_num", None)
        return E(req.url.path + ("?" + urlencode(q) if q else ""))
    start = max(1, page_num - 2)
    end = min(pages, max(5, page_num + 2))
    if end - start < 4:
        start = max(1, end - 4)
    buttons = [f'<a class="page-button page-arrow" href="{href(max(1,page_num-1))}" aria-label="Previous page" {"aria-disabled='true'" if page_num==1 else ""}>‹</a>']
    for n in range(start, end + 1):
        buttons.append(f'<a class="page-button {"active" if n==page_num else ""}" href="{href(n)}" aria-current="{"page" if n==page_num else "false"}">{n}</a>')
    if end < pages:
        if end < pages - 1: buttons.append('<span class="page-ellipsis">…</span>')
        buttons.append(f'<a class="page-button" href="{href(pages)}">{pages}</a>')
    buttons.append(f'<a class="page-button page-arrow" href="{href(min(pages,page_num+1))}" aria-label="Next page" {"aria-disabled='true'" if page_num==pages else ""}>›</a>')
    english = req.session.get("lang") == "en"
    shown = min(total, (page_num-1)*per_page + (1 if total else 0))
    to = min(total, page_num*per_page)
    summary = (f"Showing {shown}–{to} of {total} tickets" if english and total else f"Showing 0 of 0 tickets" if english else f"Hiển thị {((page_num-1)*per_page+1) if total else 0}–{to} / {total} ticket")
    all_total = total if all_total is None else all_total
    all_count_label = (f"All tickets: {all_total}" if english else f"Tổng tất cả ticket: {all_total}")
    all_count_html = f'<span class="pagination-all-count">{all_count_label}</span>' if show_all_count else ""
    return f'<div class="ticket-pagination"><div class="pagination-counts"><span class="pagination-summary">{summary}</span>{all_count_html}</div><nav class="pagination-buttons" aria-label="Ticket pages">{"".join(buttons)}</nav></div>' 

def ticket_rows(req, rows, admin=False, all_tickets=False):
    """Responsive ticket table styled to match the FMP medical service-desk reference."""
    if not rows:
        return f'<p class="empty-state"><span class="empty-icon" aria-hidden="true"><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8M8 12h5"/><circle cx="16.5" cy="16.5" r="3.5"/><path d="m19 19 2 2"/></svg></span>{tr(req,"no_tickets")}</p>'
    out = ""
    for idx, r in enumerate(rows, 1):
        status = r["status"] or "Mới"
        stcls = {"Mới":"st-new", "Đang xử lý":"st-progress", "Đã đóng":"st-closed"}.get(status, "st-new")
        severity_raw = str(r["severity"] or "").strip().casefold()
        sevcls = ("sev-high" if severity_raw in ("cao", "high") else
                  "sev-medium" if severity_raw in ("trung bình", "medium") else
                  "sev-low" if severity_raw in ("thấp", "low") else "sev-medium")
        handler_name = r["handler"] or ("Chưa phân công" if req.session.get("lang") != "en" else "Unassigned")
        english = req.session.get("lang") == "en"
        take_label = "Take over" if english else "Nhận xử lý"
        if admin and status == "Mới":
            return_to = req.url.path + (("?" + req.url.query) if req.url.query else "")
            handler_cell = (f'<td class="handler-action-cell"><form method="post" action="/ticket/{r["id"]}/take" class="inline-take-form">'
                            f'<input type="hidden" name="return_to" value="{E(return_to)}">'
                            f'<button type="submit" class="inline-take-button {sevcls}" title="{E(take_label)}"><span class="inline-take-icon" aria-hidden="true">⚒</span><span>{E(take_label)}</span></button>'
                            '</form></td>')
        elif status == "Đã đóng":
            if r["handler"]:
                closed_label = "Closed" if english else "Đã đóng"
                handler_cell = (f'<td class="handler-action-cell" title="{E(closed_label + ": " + str(r["handler"]))}">'
                                f'<span class="handler-chip handler-chip-closed">⚒ {E(r["handler"])}</span></td>')
            else:
                handler_cell = '<td class="handler-action-cell">—</td>'
        elif status == "Đang xử lý" and r["handler"]:
            progress_label = "In progress" if english else "Đang xử lý"
            handler_cell = (f'<td class="handler-action-cell" title="{E(progress_label + ": " + str(r["handler"]))}">'
                            f'<span class="handler-chip handler-chip-assigned {sevcls}">⚒ {E(r["handler"])}</span></td>')
        else:
            handler_cell = f'<td class="handler-action-cell" title="{E(handler_name)}"><span class="handler-chip">⚒ {E(handler_name)}</span></td>'
        issue_category = r["issue_type"] or "Khác"
        issue_category_label = issue_label(req, issue_category)
        created = r["created_at"] or "—"
        created_short = created  # show the complete created timestamp; never truncate
        location = r["location"] or "—"
        reporter = r["reporter"] or "—"
        detail_url = f"/ticket/{r['id']}" + ("?all_tickets=1" if all_tickets and not admin else "")
        feedback = query("SELECT rating, comment FROM ticket_feedback WHERE ticket_id=?", (r["id"],), one=True) if status == "Đã đóng" else None
        rating_mark = ""
        fb_tip = ""
        if feedback:
            # Compact bilingual tooltip: rating category + face and score only; never show the comment here.
            rating_value = int(feedback["rating"])
            is_good = rating_value >= 4
            face = "😊" if is_good else "😢"
            if req.session.get("lang") == "en":
                rating_group = "Good" if is_good else "Poor"
            else:
                rating_group = "Tốt" if is_good else "Tệ"
            fb_tip = f"{rating_group}: {face}{rating_value}"
            rating_mark = f'<span class="ticket-csat-mini ticket-csat-{feedback["rating"]}" data-tip="{E(fb_tip)}" aria-label="{E(fb_tip)}" tabindex="0"><span class="rating-face" aria-hidden="true">{face}</span></span>'
        elif status == "Đã đóng":
            # Third CSAT state: closed ticket that has not been rated yet.
            fb_tip = "Not rated yet" if req.session.get("lang") == "en" else "Chưa đánh giá"
            rating_mark = f'<span class="ticket-csat-mini ticket-csat-pending" data-tip="{E(fb_tip)}" aria-label="{E(fb_tip)}" tabindex="0"><span class="rating-face" aria-hidden="true">😐</span></span>'
        status_title = E(status_label(req,status))
        # Keep number and rating face together in normal document flow.
        rating_slot = rating_mark or '<span class="ticket-csat-slot" aria-hidden="true"></span>'
        out += f'''<tr class="ticket-row-{"new" if status=="Mới" else "progress" if status=="Đang xử lý" else "closed"} ticket-row-{sevcls}">
          <td class="row-number"><span class="row-number-group"><span class="row-number-value">{idx}</span>{rating_slot}</span></td>
          <td class="ticket-code-cell"><a class="ticket-code" href="{detail_url}" {'title="'+E(r["code"] or "")+'"' if not feedback else ""}>{E(r["code"] or "")}</a></td>
          <td class="device-cell" title="{E(device_label(req,r["device"] or ""))}">{E(device_label(req,r["device"] or ""))}</td>
          <td class="issue-cell issue-category-cell"><span class="issue-category-chip issue-category-{E(LEGACY_ISSUE_CATEGORY.get(issue_category, issue_category).replace(" ", "-").replace("/", "").replace("&", "and"))}" title="{E(issue_category_label)}"><span>{E(issue_category_label)}</span></span></td>
          <td class="department-cell"><span class="reporter-chip dept-chip department-label" title="{E(department_label(req,r["department"] or "—"))}"><span>{E(department_label(req,r["department"] or "—"))}</span></span></td>
          <td class="location-cell" title="{E(location)}">{E(location)}</td>
          <td class="reporter-cell" title="{E(reporter)}">{E(reporter)}</td>
          <td class="priority-cell" title="{E(sev_label(req,r["severity"] or ""))}"><span class="pill {sevcls} priority-label"><span class="pill-label">{E(sev_label(req,r["severity"] or ""))}</span></span></td>
          <td class="status-cell"><span class="pill {stcls}" title="{status_title}" aria-label="{E(status_label(req,status))}"><span class="pill-icon" aria-hidden="true">{status_icon_svg(status, 12)}</span><span class="pill-label">{E(status_label(req,status))}</span></span></td>
          <td class="date-cell" title="{E(created)}">{E(created_short)}</td>
          {handler_cell}
        </tr>'''
    if req.session.get("lang") == "en":
        headers = ["#","Ticket ID","Device","Issue type","Department","Location","Reporter","Severity","Status","Created",("Take over" if admin else "Assigned to")]
    else:
        headers = ["#","Mã ticket","Thiết bị","Loại sự cố","Bộ phận","Vị trí","Người báo","Mức độ","Trạng thái","Thời gian tạo",("Nhận xử lý" if admin else "Người xử lý")]
    heads = "".join(f'<th class="{"priority-heading" if h in ("Severity", "Mức độ") else "status-heading" if h in ("Status", "Trạng thái") else ""}">{h}</th>' for h in headers)
    table_css = """
    <style>
      /* V12: one reusable sizing system for compact chips inside ticket tables only. */
      .ticket-table{
        --ticket-chip-height:25px;
        --ticket-chip-font:10px;
        --ticket-chip-line:1.15;
        --ticket-chip-px:7px;
        --ticket-chip-py:3px;
        --ticket-chip-radius:7px;
      }
      .ticket-table .ticket-code,
      .ticket-table .dept-chip,
      .ticket-table .issue-category-chip,
      .ticket-table td.priority-cell .pill,
      .ticket-table td.status-cell .pill,
      .ticket-table .inline-take-button,
      .ticket-table .handler-chip{
        box-sizing:border-box!important;display:inline-flex!important;align-items:center!important;
        justify-content:center!important;min-height:var(--ticket-chip-height)!important;
        height:var(--ticket-chip-height)!important;padding:var(--ticket-chip-py) var(--ticket-chip-px)!important;
        border-radius:var(--ticket-chip-radius)!important;font-size:var(--ticket-chip-font)!important;
        font-weight:650!important;line-height:var(--ticket-chip-line)!important;
        vertical-align:middle!important;white-space:nowrap!important;
      }
      /* Content chips stay compact and use ellipsis only when their table cell is too narrow. */
      .ticket-table .ticket-code,.ticket-table .issue-category-chip,
      .ticket-table .dept-chip,.ticket-table .handler-chip{
        max-width:100%!important;min-width:0!important;overflow:hidden!important;text-overflow:ellipsis!important;
      }
      .ticket-table .ticket-code,.ticket-table .handler-chip{gap:4px!important}
      .ticket-table .dept-chip{width:auto!important;justify-content:flex-start!important}
      .ticket-table .issue-category-chip{gap:4px!important}
      .ticket-table .issue-category-chip>span{
        min-width:0!important;overflow:hidden!important;text-overflow:ellipsis!important;white-space:nowrap!important;
      }
      .ticket-table .inline-take-button{gap:4px!important;cursor:pointer!important;transition:filter .15s ease!important}
      .ticket-table .inline-take-button:hover{filter:brightness(.95)!important;transform:none!important}
      .ticket-table .inline-take-icon{
        display:inline-flex!important;align-items:center!important;justify-content:center!important;
        flex:0 0 10px!important;width:10px!important;height:10px!important;font-size:10px!important;line-height:1!important;
      }
      /* Severity/status share the same outer geometry while retaining their semantic colors. */
      .ticket-table td.priority-cell,.ticket-table td.status-cell{
        padding:8px 5px!important;text-align:center!important;vertical-align:middle!important;white-space:normal!important;
      }
      .ticket-table td.priority-cell .pill,.ticket-table td.status-cell .pill{
        max-width:100%!important;gap:4px!important;overflow:hidden!important;text-overflow:ellipsis!important;white-space:nowrap!important;
      }
      .ticket-table td.priority-cell .pill-label,.ticket-table td.status-cell .pill-label{
        display:inline!important;min-width:0!important;height:auto!important;min-height:0!important;
        padding:0!important;margin:0!important;border:0!important;border-radius:0!important;background:transparent!important;
        font-size:var(--ticket-chip-font)!important;font-weight:650!important;line-height:var(--ticket-chip-line)!important;
        white-space:nowrap!important;
      }
      .ticket-table td.priority-cell .pill-icon,.ticket-table td.status-cell .pill-icon{
        display:inline-flex!important;align-items:center!important;justify-content:center!important;
        flex:0 0 9px!important;width:9px!important;height:9px!important;font-size:8px!important;line-height:1!important;
      }
      .ticket-table td.department-cell{text-align:left!important}
      .ticket-table th:nth-child(3),.ticket-table td:nth-child(3){width:8.5%!important}
      .ticket-table th:nth-child(4),.ticket-table td:nth-child(4){width:9.5%!important}
      .ticket-table th:nth-child(7),.ticket-table td:nth-child(7){width:9.5%!important}
      .ticket-table th:nth-child(8),.ticket-table td:nth-child(8){width:8%!important}
      .ticket-table th:nth-child(9),.ticket-table td:nth-child(9){width:9%!important}
      .ticket-table td.department-cell .dept-chip{margin-left:0!important}
      /* Vạch màu đầu dòng được đồng bộ theo trạng thái ticket, không theo mức độ. */
      .handler-action-cell{white-space:nowrap;text-align:center}
      .inline-take-form{display:inline-flex;margin:0}
      .inline-take-button.sev-high{color:#b42332;background:#fff0f1;border:1px solid #ef9aa3}
      .inline-take-button.sev-medium{color:#a65b00;background:#fff5e5;border:1px solid #f1c27b}
      .inline-take-button.sev-low{color:#1266a5;background:#edf6ff;border:1px solid #9bc8ed}
      .handler-chip-assigned.sev-high{color:#b42332;background:#fff0f1}
      .handler-chip-assigned.sev-medium{color:#a65b00;background:#fff5e5}
      .handler-chip-assigned.sev-low{color:#1266a5;background:#edf6ff}
      .handler-chip-closed{color:#64748b;background:#f1f3f5;border:1px solid #d9dee5}

      /* V19: keep the rating face strictly inside the # cell and align ticket-code chips. */
      .ticket-table tbody td.row-number{
        position:relative!important;
        width:24px!important;min-width:24px!important;max-width:24px!important;
        box-sizing:border-box!important;
        padding-left:6px!important;padding-right:2px!important;
        text-align:left!important;vertical-align:middle!important;
        white-space:nowrap!important;overflow:hidden!important;
      }
      .ticket-table tbody td.row-number .row-number-value{
        display:inline-block!important;position:relative!important;left:-5px!important;
        margin:0!important;padding:0!important;line-height:1.2!important;
        vertical-align:middle!important;font-variant-numeric:tabular-nums!important;
      }
      .ticket-table tbody td.row-number > .ticket-csat-mini{
        right:1px!important;left:auto!important;top:50%!important;
        transform:translateY(-50%)!important;
      }
      .ticket-table tbody td.ticket-code-cell{
        padding:10px 10px!important;
        text-align:left!important;vertical-align:middle!important;
        line-height:1.25!important;white-space:nowrap!important;
      }
      .ticket-table tbody td.ticket-code-cell > .ticket-code{
        display:inline-flex!important;align-items:center!important;justify-content:flex-start!important;
        position:static!important;float:none!important;clear:none!important;
        box-sizing:border-box!important;min-height:23px!important;
        margin:0!important;vertical-align:middle!important;
        line-height:1.2!important;white-space:nowrap!important;
        transform:none!important;
      }
      /* V20: number + rating face form one compact centered group. */
      .ticket-table tbody td.row-number{
        position:static!important;width:4%!important;min-width:0!important;max-width:none!important;
        box-sizing:border-box!important;padding:7px 2px!important;
        text-align:center!important;vertical-align:middle!important;
        white-space:nowrap!important;overflow:hidden!important;
      }
      .ticket-table tbody td.row-number .row-number-group{
        display:inline-flex!important;align-items:center!important;justify-content:center!important;
        gap:2px!important;max-width:100%!important;white-space:nowrap!important;
        vertical-align:middle!important;line-height:1!important;
      }
      .ticket-table tbody td.row-number .row-number-value{
        position:static!important;left:auto!important;top:auto!important;
        display:inline-block!important;flex:0 0 auto!important;
        margin:0!important;padding:0!important;font-size:10px!important;line-height:1!important;
        font-variant-numeric:tabular-nums!important;vertical-align:middle!important;
      }
      .ticket-table tbody td.row-number .ticket-csat-mini,
      .ticket-table tbody td.row-number .ticket-csat-slot{
        position:static!important;inset:auto!important;
        display:inline-flex!important;align-items:center!important;justify-content:center!important;
        flex:0 0 12px!important;width:12px!important;min-width:12px!important;
        height:12px!important;min-height:12px!important;box-sizing:border-box!important;
        padding:0!important;margin:0!important;transform:none!important;float:none!important;
        border-radius:50%!important;line-height:1!important;vertical-align:middle!important;
      }
      .ticket-table tbody td.row-number .ticket-csat-mini .rating-face{
        display:block!important;font-size:9px!important;line-height:1!important;
      }
      /* V24: normalize table-cell geometry and align all row content on one baseline. */
      .ticket-table tbody td{
        box-sizing:border-box!important;
        vertical-align:middle!important;
        padding-top:10px!important;
        padding-bottom:10px!important;
      }
      /* A TD must remain a table-cell; inline-flex here breaks the table column geometry. */
      .ticket-table tbody td.device-cell{
        display:table-cell!important;
        vertical-align:middle!important;
        white-space:nowrap!important;
        line-height:1.35!important;
      }
      .ticket-table tbody td.ticket-code-cell{
        padding-left:10px!important;padding-right:10px!important;
        text-align:left!important;vertical-align:middle!important;
        white-space:nowrap!important;overflow:hidden!important;
        line-height:1.25!important;
      }
      .ticket-table tbody td.ticket-code-cell > .ticket-code{
        display:inline-flex!important;align-items:center!important;justify-content:flex-start!important;
        box-sizing:border-box!important;
        height:25px!important;min-height:25px!important;max-height:25px!important;
        max-width:100%!important;min-width:0!important;
        margin:0!important;padding:3px 7px!important;
        border-radius:7px!important;
        float:none!important;position:static!important;transform:none!important;
        text-align:left!important;vertical-align:middle!important;
        white-space:nowrap!important;line-height:1.15!important;
        font-size:10px!important;font-weight:700!important;
      }
      /* Keep the chip family at the same 25px height as the ticket ID. */
      .ticket-table .dept-chip,
      .ticket-table .issue-category-chip,
      .ticket-table td.priority-cell .pill,
      .ticket-table td.status-cell .pill,
      .ticket-table .handler-chip,
      .ticket-table .inline-take-button{
        box-sizing:border-box!important;
        height:25px!important;min-height:25px!important;max-height:25px!important;
        line-height:1.15!important;vertical-align:middle!important;
      }
      .ticket-table td.priority-cell,
      .ticket-table td.status-cell{
        padding-top:10px!important;padding-bottom:10px!important;
        vertical-align:middle!important;
      }
      /* V26: center department cell content consistently with other table cells. */
      .ticket-table tbody td.department-cell{
        text-align:center!important;
        vertical-align:middle!important;
      }
      .ticket-table tbody td.department-cell .dept-chip{
        display:inline-flex!important;
        align-items:center!important;
        justify-content:center!important;
        margin:0 auto!important;
        max-width:100%!important;
        min-width:0!important;
        text-align:center!important;
        vertical-align:middle!important;
      }
      .ticket-table tbody td.department-cell .dept-chip > span{
        display:block!important;
        min-width:0!important;
        max-width:100%!important;
        text-align:center!important;
        white-space:nowrap!important;
        overflow:hidden!important;
        text-overflow:ellipsis!important;
      }
      /* V27: make Device typography match the regular ticket-table body text. */
      .ticket-table tbody td.device-cell{
        display:table-cell!important;
        vertical-align:middle!important;
        font-family:Inter,'Segoe UI',Arial,sans-serif!important;
        font-size:12px!important;
        font-weight:500!important;
        line-height:1.4!important;
        color:#344d68!important;
        white-space:normal!important;
        overflow:visible!important;
        text-overflow:clip!important;
        overflow-wrap:break-word!important;
        word-break:normal!important;
      }
      /* V28: unify Device, Location, Reporter and Created typography. */
      .ticket-table tbody td.device-cell,
      .ticket-table tbody td.location-cell,
      .ticket-table tbody td.reporter-cell,
      .ticket-table tbody td.date-cell{
        font-family:Inter,'Segoe UI',Arial,sans-serif!important;
        font-size:12px!important;
        font-weight:500!important;
        line-height:1.4!important;
        color:#344d68!important;
      }
      /* Keep the Created value numeric and visually aligned without changing its format. */
      .ticket-table tbody td.date-cell{
        font-variant-numeric:tabular-nums!important;
        letter-spacing:0!important;
      }
    </style>
    """
    return f'{table_css}<div class="table-wrap"><table class="ticket-table"><thead><tr>{heads}</tr></thead><tbody>{out}</tbody></table></div>'
def admin_home_dashboard(req: Request, u, status: str = "", device: str = "", department: str = "", day: str = "", severity: str = "", issue_type: str = "", high_only: int = 0):
    """Admin landing dashboard following the FMP reference image supplied by the user."""
    english = req.session.get("lang") == "en"
    now = now_dt()
    def count(sql, params=()):
        row = query(sql, params, one=True)
        return int(row["n"] if row else 0)
    total = count("SELECT COUNT(*) AS n FROM tickets")
    new_count = count("SELECT COUNT(*) AS n FROM tickets WHERE COALESCE(status,'Mới')='Mới'")
    progress_count = count("SELECT COUNT(*) AS n FROM tickets WHERE status='Đang xử lý'")
    closed_count = count("SELECT COUNT(*) AS n FROM tickets WHERE status='Đã đóng'")
    # Dashboard preview uses the same filter and robust date/severity/status ordering as the full Admin list.
    if status not in STATUSES: status = ""
    if device not in DEVICES: device = ""
    department = normalize_department(department) if department else ""
    if department and department not in department_options(): department = ""
    import re
    if day:
        try:
            if re.fullmatch(r"\d{2}/\d{2}/\d{4}", day): day = datetime.strptime(day, "%d/%m/%Y").strftime("%Y-%m-%d")
            elif re.fullmatch(r"\d{4}-\d{2}-\d{2}", day): day = datetime.strptime(day, "%Y-%m-%d").strftime("%Y-%m-%d")
            else: day = ""
        except ValueError: day = ""
    if severity not in SEVERITIES: severity = "Cao" if high_only else ""
    if issue_type not in ISSUE_TYPES: issue_type = ""
    day_sql = "substr(created_at,7,4)||'-'||substr(created_at,4,2)||'-'||substr(created_at,1,2)"
    created_day_sort = "substr(created_at,7,4)||substr(created_at,4,2)||substr(created_at,1,2)"
    created_sort = created_day_sort + "||substr(created_at,12,8)"
    severity_order_sql = "CASE WHEN lower(trim(COALESCE(severity,''))) IN ('cao','high') THEN 0 WHEN lower(trim(COALESCE(severity,''))) IN ('trung bình','medium') THEN 1 WHEN lower(trim(COALESCE(severity,''))) IN ('thấp','low') THEN 2 ELSE 3 END"
    status_order_sql = "CASE WHEN lower(trim(COALESCE(status,'Mới'))) IN ('mới','new','pending','chưa xử lý') THEN 0 WHEN lower(trim(COALESCE(status,''))) IN ('đang xử lý','in progress','in-progress') THEN 1 WHEN lower(trim(COALESCE(status,''))) IN ('đã đóng','closed','completed','đã xử lý') THEN 2 ELSE 3 END"
    order_clause = f"{status_order_sql} ASC, {severity_order_sql} ASC, {created_sort} ASC, id ASC"
    # Homepage is a simple preview only: always show the latest tickets without filters.
    # Users can open the full Admin ticket list to filter/search tickets.
    recent_rows = query(f"SELECT * FROM tickets ORDER BY {created_sort} DESC, id DESC LIMIT 6")
    filtered_total = total
    today_iso = now.strftime("%Y-%m-%d")
    english_status_labels = {"Mới": "Pending", "Đang xử lý": "In progress", "Đã đóng": "Completed"}
    all_status_label = "All statuses" if english else "Tất cả trạng thái"
    status_opts = '<option value="">'+all_status_label+'</option>'+''.join(f'<option value="{E(x)}" {"selected" if status==x else ""}>{E(english_status_labels[x] if english else ("Chưa xử lý" if x=="Mới" else "Đang xử lý" if x=="Đang xử lý" else "Đã xử lý"))}</option>' for x in STATUSES)
    device_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(x)}" {"selected" if device==x else ""}>{E(device_label(req,x))}</option>' for x in DEVICES)
    dept_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(x)}" {"selected" if department==x else ""}>{E(department_label(req,x))}</option>' for x in department_options())
    severity_order = ["Cao", "Trung bình", "Thấp"]
    severity_colors = {"Cao":"#b42332", "Trung bình":"#94610f", "Thấp":"#1d4ed8"}
    severity_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(x)}" style="color:{severity_colors[x]};font-weight:700" {"selected" if severity==x else ""}>{E(sev_label(req,x))}</option>' for x in severity_order)
    issue_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(x)}" {"selected" if issue_type==x else ""}>{E(issue_label(req,x))}</option>' for x in ISSUE_TYPES)
    day_display = datetime.strptime(day, "%Y-%m-%d").strftime("%d/%m/%Y") if day else ""
    dashboard_filter_html = f"""<form method="get" class="dashboard-ticket-filters"><div><label>{tr(req,"dept")}</label><select name="department">{dept_opts}</select></div><div><label>{tr(req,"status")}</label><select name="status">{status_opts}</select></div><div><label>{tr(req,"device")}</label><select name="device">{device_opts}</select></div><div><label>{tr(req,"issue_type")}</label><select name="issue_type">{issue_opts}</select></div><div><label>{"Created date" if english else "Ngày tạo"}</label><input name="day" value="{E(day_display)}" placeholder="DD/MM/YYYY" inputmode="numeric"></div><div><label>{tr(req,"severity")}</label><select name="severity">{severity_opts}</select></div><div class="dashboard-filter-actions"><button type="submit">⌕ {tr(req,"filter")}</button><a href="/">{ "Reset" if english else "Đặt lại"}</a></div></form>"""
    title = "Welcome to IT Service Desk" if english else "Chào mừng đến với IT Service Desk"
    subtitle = "Fast support · Effective solutions · Here for you" if english else "Hỗ trợ nhanh chóng · Giải pháp hiệu quả · Đồng hành cùng bạn"
    cards = [("total", "Total tickets" if english else "Tổng số ticket", total, "ticket", None), ("new", "New" if english else "Mới", new_count, "plus", "Mới"), ("progress", "In progress" if english else "Đang xử lý", progress_count, "clock", "Đang xử lý"), ("closed", "Closed" if english else "Đã đóng", closed_count, "check", "Đã đóng")]
    # Compare ticket creation volume over the last 7 days against the preceding 7 days.
    # Timestamps are stored as DD/MM/YYYY HH:MM:SS in Vietnam local time.
    from datetime import timedelta
    today = now.date()
    current_start = today - timedelta(days=6)
    previous_start = today - timedelta(days=13)
    # Aggregate daily ticket totals in SQLite; do not transfer every ticket row to Python.
    all_created = query("SELECT substr(created_at,7,4)||'-'||substr(created_at,4,2)||'-'||substr(created_at,1,2) AS ticket_day, status, COUNT(*) AS n FROM tickets WHERE length(created_at)>=10 GROUP BY ticket_day, status")
    def ticket_day(value):
        try:
            return datetime.strptime(str(value or "")[:10], "%Y-%m-%d").date()
        except Exception:
            return None
    def period_count(start, end, status_filter=None):
        n = 0
        for item in all_created:
            d = ticket_day(item["ticket_day"])
            if d and start <= d <= end and (status_filter is None or item["status"] == status_filter):
                n += int(item["n"] or 0)
        return n
    paths = {"ticket": '<rect x="4" y="5" width="16" height="15" rx="2"/><path d="M8 5V3h8v2M8 10h8M8 14h5"/>', "plus": '<circle cx="12" cy="12" r="9"/><path d="M12 6v6l4 2"/>', "clock": '<path d="M20 7v5h-5"/><path d="M4 17v-5h5"/><path d="M5.6 9A7.5 7.5 0 0 1 18 6l2 6M4 12l2 6a7.5 7.5 0 0 0 12.4-1"/>', "check": '<circle cx="12" cy="12" r="9"/><path d="m7.5 12.2 3 3 6-6.2"/>'}
    stat_html = ""
    for key, label, value, icon, status_filter in cards:
        query_string = "/admin?all_dates=1" if key == "total" else "/admin?all_dates=1&status=" + urlencode({"status": status_filter}).split("=",1)[1]
        current_n = period_count(current_start, today, status_filter)
        previous_n = period_count(previous_start, current_start - timedelta(days=1), status_filter)
        if previous_n == 0:
            trend = 100 if current_n > 0 else 0
            trend_label = ("↑ " if current_n > 0 else "→ ") + ("100%" if current_n > 0 else "0%")
        else:
            trend = round((current_n - previous_n) * 100 / previous_n)
            trend_label = ("↑ " if trend > 0 else "↓ " if trend < 0 else "→ ") + f"{abs(trend)}%"
        trend_class = "trend-up" if trend > 0 else "trend-down" if trend < 0 else "trend-flat"
        csat_counts = ""
        if key == "closed":
            counts = query("SELECT SUM(CASE WHEN f.rating>=4 THEN 1 ELSE 0 END) AS good, SUM(CASE WHEN f.rating BETWEEN 1 AND 3 THEN 1 ELSE 0 END) AS bad, SUM(CASE WHEN f.ticket_id IS NULL THEN 1 ELSE 0 END) AS unrated FROM tickets t LEFT JOIN ticket_feedback f ON f.ticket_id=t.id WHERE t.status='Đã đóng'", one=True)
            # Match the compact custom tooltip used by the rating icons in ticket lists.
            csat_counts = (f'<span class="csat-card-counts">'
                           f'<span class="csat-stat-filter" data-tip="{E("Good: 😊" if english else "Tốt: 😊")}" data-feedback-group="good" role="link" tabindex="0" aria-label="{"Good" if english else "Tốt"}">😊 {counts["good"] or 0}</span>　'
                           f'<span class="csat-stat-filter" data-tip="{E("Poor: 😢" if english else "Tệ: 😢")}" data-feedback-group="poor" role="link" tabindex="0" aria-label="{"Poor" if english else "Tệ"}">😢 {counts["bad"] or 0}</span>　'
                           f'<span class="csat-stat-filter" data-tip="{E("Not rated: 😐" if english else "Chưa đánh giá: 😐")}" data-feedback-group="unrated" role="link" tabindex="0" aria-label="{"Not rated" if english else "Chưa đánh giá"}">😐 {counts["unrated"] or 0}</span></span>')
        stat_html += f'<a class="admin-stat-card stat-{key}" href="{query_string}" title="{"Last 7 days compared with previous 7 days" if english else "7 ngày gần nhất so với 7 ngày trước đó"}"><span class="admin-stat-icon"><svg viewBox="0 0 24 24" width="29" height="29" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{paths[icon]}</svg></span><span class="admin-stat-copy"><span>{label}</span><strong>{value:,}</strong><small class="stat-insight {trend_class}"><span class="trend-symbol">{"↑" if trend > 0 else "↓" if trend < 0 else "→"}</span><span>{trend_label.split(" ",1)[-1]}</span><em>{"vs previous 7 days" if english else "so với 7 ngày trước"}</em></small>{csat_counts}</span></a>'
    # Desktop charts are calculated from the existing SQLite tickets table only.
    created_daily_rows = query("SELECT substr(created_at,7,4)||'-'||substr(created_at,4,2)||'-'||substr(created_at,1,2) AS d, COUNT(*) AS n FROM tickets WHERE length(COALESCE(created_at,''))>=10 GROUP BY d")
    closed_daily_rows = query("SELECT substr(closed_at,7,4)||'-'||substr(closed_at,4,2)||'-'||substr(closed_at,1,2) AS d, COUNT(*) AS n FROM tickets WHERE length(COALESCE(closed_at,''))>=10 GROUP BY d")
    created_daily, closed_daily = {}, {}
    for item in created_daily_rows:
        try: created_daily[datetime.strptime(item["d"], "%Y-%m-%d").date()] = int(item["n"] or 0)
        except Exception: pass
    for item in closed_daily_rows:
        try: closed_daily[datetime.strptime(item["d"], "%Y-%m-%d").date()] = int(item["n"] or 0)
        except Exception: pass
    chart_days = [today - timedelta(days=6-i) for i in range(7)]
    chart_values = [(d, created_daily.get(d, 0), closed_daily.get(d, 0)) for d in chart_days]
    chart_peak = max([a+b for _, a, b in chart_values] or [0])
    chart_y_max = max(50, ((chart_peak + 49) // 50) * 50)
    chart_left, chart_top, chart_height, chart_width = 54, 24, 190, 640
    chart_bottom = chart_top + chart_height
    chart_svg_parts = [f'<svg class="admin-trend-svg" viewBox="0 0 720 278" role="img" aria-label="{"Ticket trend for the last seven days" if english else "Xu hướng ticket trong 7 ngày gần nhất"}" xmlns="http://www.w3.org/2000/svg">']
    for tick_i in range(5):
        tick = int(chart_y_max * tick_i / 4)
        y = chart_bottom - chart_height * tick / chart_y_max
        chart_svg_parts.append(f'<line x1="{chart_left}" y1="{y:.1f}" x2="704" y2="{y:.1f}" stroke="#e5edf6" stroke-width="1"/>')
        chart_svg_parts.append(f'<text x="43" y="{y+4:.1f}" text-anchor="end" fill="#627d9d" font-size="13">{tick:,}</text>')
    day_labels_vi, day_labels_en = ["T2","T3","T4","T5","T6","T7","CN"], ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
    for i, (d, created_n, closed_n) in enumerate(chart_values):
        center_x = chart_left + (i + .5) * (chart_width / 7)
        bar_w = 42
        created_h = chart_height * created_n / chart_y_max
        closed_h = chart_height * closed_n / chart_y_max
        total_day = created_n + closed_n
        total_h = created_h + closed_h
        bar_x = center_x - bar_w / 2
        bar_top = chart_bottom - total_h
        blue_y = chart_bottom - created_h
        green_y = bar_top
        if total_day:
            # A single rounded bar container; the two colored segments meet exactly.
            clip_id = f"ticket-bar-clip-{i}"
            chart_svg_parts.append(f'<defs><clipPath id="{clip_id}"><rect x="{bar_x:.1f}" y="{bar_top:.1f}" width="{bar_w}" height="{total_h:.1f}" rx="5" ry="5"/></clipPath></defs>')
            chart_svg_parts.append(f'<g clip-path="url(#{clip_id})"><rect x="{bar_x:.1f}" y="{blue_y:.1f}" width="{bar_w}" height="{created_h:.1f}" fill="#168be8"/><rect x="{bar_x:.1f}" y="{green_y:.1f}" width="{bar_w}" height="{closed_h:.1f}" fill="#28b878"/></g>')
            chart_svg_parts.append(f'<text x="{center_x:.1f}" y="{max(14,bar_top-8):.1f}" text-anchor="middle" fill="#153b70" font-size="14" font-weight="750">{total_day:,}</text>')
        label = day_labels_en[d.weekday()] if english else day_labels_vi[d.weekday()]
        chart_svg_parts.append(f'<text x="{center_x:.1f}" y="{chart_bottom+25}" text-anchor="middle" fill="#234b78" font-size="13" font-weight="650">{label}</text>')
    chart_svg_parts.append(f'<line x1="{chart_left}" y1="{chart_top}" x2="{chart_left}" y2="{chart_bottom}" stroke="#b8cbe0" stroke-width="1.2"/><line x1="{chart_left}" y1="{chart_bottom}" x2="704" y2="{chart_bottom}" stroke="#b8cbe0" stroke-width="1.2"/>')
    trend_svg = ''.join(chart_svg_parts) + '</svg>'

    status_values = [new_count, progress_count, closed_count]
    status_names = (["New", "In progress", "Closed"] if english else ["Mới", "Đang xử lý", "Đã đóng"])
    status_colors = ["#168be8", "#f59e0b", "#28b878"]
    status_sum = sum(status_values)
    donut_total = total if total else status_sum
    donut_circ = 2 * math.pi * 76
    donut_parts = [f'<svg class="admin-status-donut" viewBox="0 0 220 220" role="img" aria-label="{"Ticket status distribution" if english else "Phân bổ trạng thái ticket"}" xmlns="http://www.w3.org/2000/svg">', '<circle cx="110" cy="110" r="76" fill="none" stroke="#edf2f7" stroke-width="25"/>']
    offset = 0.0
    if status_sum:
        for value, color in zip(status_values, status_colors):
            length = donut_circ * value / status_sum
            if value: donut_parts.append(f'<circle cx="110" cy="110" r="76" fill="none" stroke="{color}" stroke-width="25" stroke-dasharray="{length:.2f} {donut_circ-length:.2f}" stroke-dashoffset="{-offset:.2f}" transform="rotate(-90 110 110)"/>')
            offset += length
    donut_parts.append(f'<text x="110" y="103" text-anchor="middle" fill="#355779" font-size="13">{"Total tickets" if english else "Tổng ticket"}</text><text x="110" y="129" text-anchor="middle" fill="#123f7b" font-size="23" font-weight="800">{donut_total:,}</text></svg>')
    donut_svg = ''.join(donut_parts)
    status_legend = ''.join(f'<div class="admin-status-legend-row"><span class="admin-status-dot" style="background:{color}"></span><span class="admin-status-name">{name}</span><strong>{value:,}</strong><small>{(value * 100 / status_sum):.1f}%</small></div>' for name, value, color in zip(status_names, status_values, status_colors)) if status_sum else f'<div class="admin-chart-empty">{"No ticket data" if english else "Chưa có dữ liệu ticket"}</div>'

    recent_status_text = {"Mới": ("New" if english else "Mới"), "Đang xử lý": ("In progress" if english else "Đang xử lý"), "Đã đóng": ("Closed" if english else "Đã đóng")}
    recent_headers = (["Created", "Ticket Code", "Issue", "Status", "Priority"] if english else ["Ngày tạo", "Mã ticket", "Nội dung sự cố", "Trạng thái", "Mức độ"])
    recent_rows_html = []
    for row in recent_rows:
        row_status = row["status"] or "Mới"
        normalized_status = str(row_status).strip().lower()
        if normalized_status in ("mới", "new", "pending", "chưa xử lý"):
            row_status_class = "new"
        elif normalized_status in ("đang xử lý", "in progress", "in-progress"):
            row_status_class = "progress"
        elif normalized_status in ("đã đóng", "closed", "completed", "đã xử lý"):
            row_status_class = "closed"
        else:
            row_status_class = "other"
        row_severity = row["severity"] or "Trung bình"
        row_severity_class = "high" if str(row_severity).strip().lower() in ("cao", "high") else "medium" if str(row_severity).strip().lower() in ("trung bình", "medium") else "low"
        created_text, issue_text, code_text = (row["created_at"] or "—")[:10], row["issue"] or "—", row["code"] or "—"
        severity_icons = {
            "high": '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M8 2.1v7.1" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/><circle cx="8" cy="12.2" r="1.25" fill="currentColor"/></svg>',
            "medium": '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M8 1.6 15 14H1z" fill="currentColor"/><path d="M8 5.2v4.2M8 11.7v.1" stroke="#fff" stroke-width="1.5" stroke-linecap="round"/></svg>',
            "low": '<svg viewBox="0 0 16 16" aria-hidden="true"><circle cx="8" cy="8" r="7" fill="currentColor"/><path d="M8 7v4M8 4.5v.2" stroke="#fff" stroke-width="1.5" stroke-linecap="round"/></svg>'
        }
        recent_rows_html.append(f'<tr class="recent-row-status-{row_status_class}"><td class="recent-date">{E(created_text)}</td><td class="recent-code"><a href="/ticket/{row["id"]}">{E(code_text)}</a></td><td class="recent-issue" title="{E(issue_text)}">{E(issue_text)}</td><td><span class="recent-status recent-status-{row_status_class}"><span class="recent-status-icon">{status_icon_svg(row_status, 12)}</span>{E(recent_status_text.get(row_status, status_label(req, row_status)))}</span></td><td><span class="recent-priority recent-priority-{row_severity_class}"><i>{severity_icons[row_severity_class]}</i>{E(sev_label(req, row_severity))}</span></td></tr>')
    recent_empty_row = f'<tr><td colspan="5" class="recent-empty">{tr(req, "no_tickets")}</td></tr>'
    recent_table_html = f'<table class="admin-recent-data"><colgroup><col style="width:14%"><col style="width:16%"><col style="width:43%"><col style="width:15%"><col style="width:12%"></colgroup><thead><tr>{"".join(f"<th>{h}</th>" for h in recent_headers)}</tr></thead><tbody>{"".join(recent_rows_html) if recent_rows_html else recent_empty_row}</tbody></table>'

    welcome_svg = '<circle cx="24" cy="24" r="21" fill="rgba(255,255,255,.72)" stroke="currentColor" stroke-width="1.2"/><path d="M10 25v-3a14 14 0 0 1 28 0v3"/><rect x="8" y="23" width="7" height="12" rx="3.5" fill="currentColor" stroke="none"/><rect x="33" y="23" width="7" height="12" rx="3.5" fill="currentColor" stroke="none"/><path d="M35 36a12 12 0 0 1-10 5h-3"/><rect x="20" y="37" width="7" height="4" rx="2" fill="currentColor" stroke="none"/>'
    body = f'''<div class="admin-dashboard">
    <section class="admin-welcome-banner"><div class="welcome-art"><svg viewBox="0 0 48 48" width="58" height="58" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round">{welcome_svg}</svg></div><div class="welcome-copy"><h1>{title}</h1><p>{subtitle}</p></div><div class="welcome-datetime"><div class="welcome-date-item"><button class="welcome-calendar-button welcome-calendar-main" type="button" id="admin-home-calendar-button" title="{'Choose date' if english else 'Chọn ngày'}" aria-label="{'Choose date' if english else 'Chọn ngày'}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18M8 14h.01M12 14h.01M16 14h.01M8 17h.01M12 17h.01M16 17h.01"/></svg></button><div class="welcome-date-control"><span class="welcome-date-text" id="admin-home-date-text">{now.strftime('%A, %d/%m/%Y') if english else ''}</span><input class="welcome-hidden-date" type="date" id="admin-home-date-picker" value="{now.strftime('%Y-%m-%d')}" aria-label="{'Choose date' if english else 'Chọn ngày'}"></div></div><div class="welcome-date-item"><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg><span class="welcome-time-text" id="admin-home-clock">{now.strftime('%H:%M:%S')}</span></div></div></section>
    <section class="admin-stat-grid">{stat_html}</section>
    <section class="admin-analytics-grid">
      <article class="box admin-chart-card admin-trend-card"><div class="admin-panel-heading"><h2><svg class="dashboard-heading-icon" viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 20V11M10 20V4M16 20v-7M22 20V8"/></svg>{'Ticket trend – 7 days' if english else 'Xu hướng ticket 7 ngày'}</h2><div class="admin-chart-legend"><span><i class="legend-created"></i>{'Created' if english else 'Tạo mới'}</span><span><i class="legend-closed"></i>{'Closed' if english else 'Đã đóng'}</span></div></div>{trend_svg}</article>
      <article class="box admin-chart-card admin-status-card"><div class="admin-panel-heading"><h2><svg class="dashboard-heading-icon" viewBox="0 0 24 24" width="20" height="20" fill="currentColor" aria-hidden="true"><path d="M11 2a10 10 0 1 0 10 11H11z"/><path d="M13 2v8h8A10 10 0 0 0 13 2z" opacity=".72"/></svg>{'Ticket status' if english else 'Trạng thái ticket'}</h2></div><div class="admin-status-content">{donut_svg}<div class="admin-status-legend">{status_legend}</div></div></article>
    </section>
    <section class="box admin-recent-panel admin-recent-full"><div class="admin-panel-heading"><h2><svg class="dashboard-heading-icon" viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.8" stroke-linecap="round" aria-hidden="true"><path d="M8 5h13M8 12h13M8 19h13"/><circle cx="3" cy="5" r="1"/><circle cx="3" cy="12" r="1"/><circle cx="3" cy="19" r="1"/></svg>{'Recent tickets' if english else 'Ticket gần đây'}</h2><a class="admin-view-all" href="/admin?all_dates=1">{'View all' if english else 'Xem tất cả'} <span aria-hidden="true">↗</span></a></div><div class="admin-recent-table-wrap">{recent_table_html}</div></section></div>
    <script>(function(){{const picker=document.getElementById('admin-home-date-picker');const dateText=document.getElementById('admin-home-date-text');const button=document.getElementById('admin-home-calendar-button');const clock=document.getElementById('admin-home-clock');const english={'true' if english else 'false'};if(!picker||!dateText||!button)return;function renderDate(value){{if(!value)return;const parts=value.split('-').map(Number);const d=new Date(parts[0],parts[1]-1,parts[2]);const locale=english?'en-US':'vi-VN';const weekday=new Intl.DateTimeFormat(locale,{{weekday:'long'}}).format(d);const formatted=String(parts[2]).padStart(2,'0')+'/'+String(parts[1]).padStart(2,'0')+'/'+parts[0];dateText.textContent=weekday.charAt(0).toUpperCase()+weekday.slice(1)+', '+formatted;}}function openPicker(){{if(picker.showPicker){{try{{picker.showPicker();return;}}catch(e){{}}}}picker.click();}}button.addEventListener('click',openPicker);picker.addEventListener('change',()=>renderDate(picker.value));renderDate(picker.value);function updateClock(){{if(!clock)return;clock.textContent=new Intl.DateTimeFormat('en-GB',{{timeZone:'Asia/Ho_Chi_Minh',hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:false}}).format(new Date());}}updateClock();setInterval(updateClock,1000);}})();</script>'''
    dashboard_desktop_css = '.admin-dashboard{max-width:1500px;margin:0 auto;padding:0 2px 20px;color:#123b70}.admin-analytics-grid{display:grid;grid-template-columns:minmax(0,1.45fr) minmax(420px,1fr);gap:14px;align-items:stretch;margin:0 0 14px}.admin-chart-card{min-width:0;margin:0;padding:15px 14px 10px;border:1px solid #dcecff;border-radius:13px;background:#fff;box-shadow:0 3px 12px rgba(42,105,165,.035)}.admin-chart-card .admin-panel-heading{margin-bottom:8px}.admin-chart-card .admin-panel-heading h2{font-size:15px}.admin-chart-legend{display:flex;align-items:center;gap:15px;white-space:nowrap;color:#456487;font-size:11px;font-weight:650}.admin-chart-legend span{display:inline-flex;align-items:center;gap:5px}.admin-chart-legend i{display:inline-block;width:9px;height:9px;border-radius:50%}.legend-created{background:#168be8}.legend-closed{background:#28b878}.admin-trend-svg{display:block;width:100%;height:auto;max-height:310px;overflow:visible}.admin-status-content{display:flex;align-items:center;justify-content:center;gap:14px;min-height:240px}.admin-status-donut{display:block;flex:0 0 220px;width:220px;height:220px}.admin-status-legend{display:grid;gap:14px;min-width:150px;flex:1}.admin-status-legend-row{display:grid;grid-template-columns:10px minmax(0,1fr) auto;align-items:center;column-gap:8px;color:#294b76;font-size:12px}.admin-status-dot{width:10px;height:10px;border-radius:50%}.admin-status-name{min-width:0}.admin-status-legend-row strong{font-size:12px;color:#143f78;font-variant-numeric:tabular-nums}.admin-status-legend-row small{grid-column:2 / 4;color:#7189a4;font-size:11px;margin-top:2px}.admin-chart-empty{color:#7189a4;font-size:12px}.admin-recent-full{margin:0;padding:15px 14px 12px;border-color:#dcecff;border-radius:13px}.admin-recent-full .admin-panel-heading{margin-bottom:10px}.admin-recent-full .dashboard-heading-icon{color:#f59e0b}.admin-view-all{display:inline-flex;align-items:center;gap:9px;padding:7px 13px;border:1px solid #d4e4f5;border-radius:999px;background:#f8fbff;color:#164a83;text-decoration:none;font-size:12px;font-weight:750;white-space:nowrap}.admin-view-all:hover{background:#edf6ff;border-color:#b8d7f7}.admin-recent-table-wrap{width:100%;overflow:visible}.admin-recent-data{width:100%;table-layout:fixed;border-collapse:separate;border-spacing:0;border:1px solid #e0ebf6;border-radius:10px;overflow:hidden;background:#fff}.admin-recent-data th{padding:10px 11px;background:#edf6ff;color:#14365f;text-align:left;font-size:11px;font-weight:750;border-bottom:1px solid #dce8f5}.admin-recent-data td{padding:9px 11px;border-bottom:1px solid #e8eff7;color:#294765;font-size:12px;vertical-align:middle}.admin-recent-data tr:last-child td{border-bottom:0}.admin-recent-data tbody tr:hover td{background:#f8fbff}.admin-recent-data .recent-date,.admin-recent-data .recent-code{white-space:nowrap;font-variant-numeric:tabular-nums}.recent-code a{color:#174a7e;text-decoration:none;font-weight:750}.recent-code a:hover{text-decoration:underline}.recent-issue{white-space:normal;overflow-wrap:anywhere;word-break:normal;line-height:1.45}.recent-status{display:inline-flex;justify-content:center;align-items:center;min-width:92px;padding:5px 9px;border-radius:999px;font-size:11px;font-weight:750;white-space:nowrap}.recent-status-new{background:#e1f1ff;color:#087bd5}.recent-status-progress{background:#f0eaff;color:#6844c8}.recent-status-closed{background:#e2f6e9;color:#21834a}.recent-priority{display:inline-flex;align-items:center;gap:7px;white-space:nowrap;font-size:11px;font-weight:700}.recent-priority i{display:inline-block;width:8px;height:8px;border-radius:50%}.recent-priority-high{color:#b42332}.recent-priority-high i{background:#d83d4a}.recent-priority-medium{color:#94610f}.recent-priority-medium i{background:#e9a12b}.recent-priority-low{color:#167653}.recent-priority-low i{background:#218bca}.recent-empty{text-align:center!important;color:#7890ad!important;padding:24px!important}.admin-recent-panel{min-width:0}.admin-dashboard .admin-welcome-banner{margin-bottom:14px}.admin-dashboard .admin-stat-grid{margin-bottom:14px}body{min-width:1100px}.pill-icon{display:inline-flex!important;align-items:center!important;justify-content:center!important;flex:0 0 12px!important;width:12px!important;height:12px!important;margin:0!important;line-height:1!important}.pill-icon .status-icon-svg,.recent-status-icon .status-icon-svg{display:block;flex:0 0 auto;width:12px;height:12px;stroke:currentColor;stroke-width:1.9}.ticket-table .status-cell .pill{display:inline-flex!important;align-items:center!important;justify-content:center!important;gap:4px!important;white-space:nowrap!important}.ticket-table .status-cell .pill-label{white-space:nowrap!important}.recent-status{gap:5px!important}.recent-status-icon{display:inline-flex;align-items:center;justify-content:center;flex:0 0 12px;width:12px;height:12px}.admin-stat-card.stat-progress{background:linear-gradient(135deg,#fff8eb 0%,#ffedcc 100%)!important;border-color:#f8d9a5!important;color:#9a5b08!important}.admin-stat-card.stat-progress .admin-stat-icon{background:#f59e0b!important;color:#fff!important;box-shadow:0 4px 10px rgba(245,158,11,.20)!important}.admin-stat-card.stat-progress:after{background:rgba(245,158,11,.11)!important}.admin-stat-card.stat-progress .admin-stat-copy strong,.admin-stat-card.stat-progress .admin-stat-copy>span{color:#9a5b08!important}.admin-recent-full .dashboard-heading-icon{color:#f59e0b!important}.admin-recent-data tbody tr.recent-row-status-progress td{background:#fff8eb!important}.admin-recent-data tbody tr.recent-row-status-progress:hover td{background:#fff0d6!important}.admin-recent-data tbody tr.recent-row-status-progress td:first-child:before{background:#f59e0b!important}.recent-status-progress{background:#ffedcc!important;color:#9a5b08!important}\n.admin-stat-card{position:relative;overflow:hidden;background:linear-gradient(135deg,#f4faff 0%,#e8f4ff 100%);border-color:#cfe6ff;box-shadow:0 4px 14px rgba(24,117,194,.055)}\n.admin-stat-card:after{content:"";position:absolute;right:-18px;bottom:-26px;width:82px;height:82px;border-radius:50%;background:rgba(22,139,232,.08);pointer-events:none}\n.admin-stat-card .admin-stat-icon{position:relative;z-index:1;flex-basis:48px;width:48px;height:48px;border-radius:15px;background:#168be8;color:#fff;box-shadow:0 4px 10px rgba(22,139,232,.16)}\n.admin-stat-card .admin-stat-icon svg{width:25px;height:25px;stroke:#fff;stroke-width:2.15}\n.admin-stat-card.stat-new .admin-stat-icon{border-radius:50%;background:#168be8}\n.admin-stat-card.stat-progress{background:linear-gradient(135deg,#fff8eb 0%,#ffedcc 100%);border-color:#f8d9a5;color:#9a5b08}\n.admin-stat-card.stat-progress .admin-stat-icon{background:#f59e0b;box-shadow:0 4px 10px rgba(245,158,11,.20)}\n.admin-stat-card.stat-progress:after{background:rgba(245,158,11,.11)}\n.admin-stat-card.stat-closed{background:linear-gradient(135deg,#f2fcf6 0%,#e7f8ee 100%);border-color:#d1f0de;color:#176b4c}\n.admin-stat-card.stat-closed .admin-stat-icon{background:#19ae72;box-shadow:0 4px 10px rgba(25,174,114,.18)}\n.admin-stat-card.stat-closed:after{background:rgba(25,174,114,.09)}\n.admin-stat-card.stat-total .admin-stat-icon{border-radius:10px}\n.admin-stat-card.stat-new:after{background:rgba(22,139,232,.10)}\n.admin-chart-card{border-color:#cfe2fa;background:linear-gradient(180deg,#fff 0%,#fcfdff 100%);box-shadow:0 4px 15px rgba(28,105,177,.045)}\n.admin-chart-card .dashboard-heading-icon{color:#20b879}\n.admin-panel-heading h2{color:#102f70}\n.admin-recent-full{border-color:#cfe2fa;background:linear-gradient(180deg,#fff 0%,#fcfdff 100%);box-shadow:0 4px 15px rgba(28,105,177,.045)}\n.admin-recent-data{border-color:#dce9f7}\n.admin-recent-data th{background:linear-gradient(180deg,#eef6ff 0%,#e7f1fc 100%);color:#123c73}\n.admin-recent-data td{transition:background .15s ease}\n.admin-recent-data tbody tr.recent-row-status-new td{background:#f2f8ff}\n.admin-recent-data tbody tr.recent-row-status-progress td{background:#fff8eb}\n.admin-recent-data tbody tr.recent-row-status-closed td{background:#f0faf4}\n.admin-recent-data tbody tr.recent-row-status-other td{background:#fff9ef}\n.admin-recent-data tbody tr td:first-child{position:relative;padding-left:17px}\n.admin-recent-data tbody tr.recent-row-status-new td:first-child:before,\n.admin-recent-data tbody tr.recent-row-status-progress td:first-child:before,\n.admin-recent-data tbody tr.recent-row-status-closed td:first-child:before,\n.admin-recent-data tbody tr.recent-row-status-other td:first-child:before{content:"";position:absolute;left:7px;top:6px;bottom:6px;width:3px;border-radius:4px}\n.admin-recent-data tbody tr.recent-row-status-new td:first-child:before{background:#168be8}\n.admin-recent-data tbody tr.recent-row-status-progress td:first-child:before{background:#f59e0b}\n.admin-recent-data tbody tr.recent-row-status-closed td:first-child:before{background:#20ad70}\n.admin-recent-data tbody tr.recent-row-status-other td:first-child:before{background:#e3a13b}\n.admin-recent-data tbody tr.recent-row-status-new:hover td{background:#e8f3ff}\n.admin-recent-data tbody tr.recent-row-status-progress:hover td{background:#fff0d6}\n.admin-recent-data tbody tr.recent-row-status-closed:hover td{background:#e5f6ec}\n.recent-status-progress{background:#ffedcc;color:#9a5b08}\n.recent-status-closed{background:#ddf4e6;color:#21834a}\n.recent-status-other{background:#fff1d9;color:#9a650c}\n.recent-priority{gap:7px}\n.recent-priority i{display:inline-flex;align-items:center;justify-content:center;width:14px;height:14px;background:transparent}\n.recent-priority i svg{display:block;width:14px;height:14px}\n.recent-priority-high{color:#d72f3c}\n.recent-priority-medium{color:#dc8b0b}\n.recent-priority-low{color:#168be8}\n.admin-view-all{background:#f5f9ff;border-color:#cfe1f5;color:#123f78}\n.admin-view-all:hover{background:#eaf4ff}\n'
    body += f'<style>{dashboard_desktop_css}</style>'
    body += f'<style>.csat-card-counts .csat-stat-filter{{display:inline-block;cursor:pointer;border-radius:0;padding:1px 2px;background:transparent!important;color:inherit!important;text-decoration:none!important;box-shadow:none!important;transition:none!important}}.csat-card-counts .csat-stat-filter:hover,.csat-card-counts .csat-stat-filter:focus,.csat-card-counts .csat-stat-filter:focus-visible{{background:transparent!important;color:inherit!important;text-decoration:none!important;outline:none!important;box-shadow:none!important}}</style><script>(function(){{const initialRevision="{ticket_sync_signature()}";let signature="";async function check(){{try{{const r=await fetch("/api/feedback/summary",{{cache:"no-store"}});if(!r.ok)return;const j=await r.json();const next=JSON.stringify([j.feedback,j.tickets,j.ticket_signature]);if((!signature&&String(j.ticket_signature)!==initialRevision)||(signature&&next!==signature)){{location.reload();return}}signature=next}}catch(e){{}}}}check();setInterval(check,15000);document.addEventListener("click",function(e){{const group=e.target.closest(".csat-stat-filter");if(group){{e.preventDefault();e.stopPropagation();const key=group.dataset.feedbackGroup;if(["good","poor","unrated"].includes(key))location.href="/admin?all_dates=1&status="+encodeURIComponent("Đã đóng")+"&feedback_group="+encodeURIComponent(key)}}}},true);document.addEventListener("keydown",function(e){{const group=e.target.closest(".csat-stat-filter");if(group&&(e.key==="Enter"||e.key===" ")){{e.preventDefault();e.stopPropagation();const key=group.dataset.feedbackGroup;if(["good","poor","unrated"].includes(key))location.href="/admin?all_dates=1&status="+encodeURIComponent("Đã đóng")+"&feedback_group="+encodeURIComponent(key)}}}},true)}})();</script>'
    return page(req, "Home" if english else "Trang chủ", body, u)

@app.get("/", response_class=HTMLResponse)
def home(req: Request, msg: str = "", status: str = "", device: str = "", department: str = "", day: str = "", severity: str = "", issue_type: str = "", high_only: int = 0, all_tickets: int = 1, page_num: int = 1, u=Depends(current_user)):
    if u["role"] == "admin":
        return admin_home_dashboard(req, u, status, device, department, day, severity, issue_type, high_only)
    if status not in STATUSES: status = ""
    if device not in DEVICES: device = ""
    department = normalize_department(department) if department else ""
    if department and department not in department_options(): department = ""
    # Keep the selected scope intact: Department combines with either My Tickets or All Tickets.
    import re
    if day:
        try:
            if re.fullmatch(r"\d{2}/\d{2}/\d{4}", day): day = datetime.strptime(day, "%d/%m/%Y").strftime("%Y-%m-%d")
            elif re.fullmatch(r"\d{4}-\d{2}-\d{2}", day): day = datetime.strptime(day, "%Y-%m-%d").strftime("%Y-%m-%d")
            else: day = ""
        except ValueError: day = ""
    if severity not in SEVERITIES: severity = "Cao" if high_only else ""
    if issue_type not in ISSUE_TYPES: issue_type = ""
    day_sql = "substr(created_at,7,4)||'-'||substr(created_at,4,2)||'-'||substr(created_at,1,2)"
    # Keep User aligned with Admin/home: newest date first; within a date sort by severity,
    # then workflow status, then newest time; accept normalized legacy labels too.
    created_day_sort = "substr(created_at,7,4)||substr(created_at,4,2)||substr(created_at,1,2)"
    created_sort = created_day_sort + "||substr(created_at,12,8)"
    severity_order_sql = "CASE WHEN lower(trim(COALESCE(severity,''))) IN ('cao','high') THEN 0 WHEN lower(trim(COALESCE(severity,''))) IN ('trung bình','medium') THEN 1 WHEN lower(trim(COALESCE(severity,''))) IN ('thấp','low') THEN 2 ELSE 3 END"
    status_order_sql = "CASE WHEN lower(trim(COALESCE(status,'Mới'))) IN ('mới','new','pending','chưa xử lý') THEN 0 WHEN lower(trim(COALESCE(status,''))) IN ('đang xử lý','in progress','in-progress') THEN 1 WHEN lower(trim(COALESCE(status,''))) IN ('đã đóng','closed','completed','đã xử lý') THEN 2 ELSE 3 END"
    user_where = f"WHERE (?=1 OR user_id=?) AND (?='' OR status=?) AND (?='' OR device=?) AND (?='' OR department=?) AND (?='' OR {day_sql}=?) AND (?='' OR severity=?) AND (?='' OR issue_type=?)"
    user_args = (1 if all_tickets else 0, u["id"], status, status, device, device, department, department, day, day, severity, severity, issue_type, issue_type)
    total_row = query(f"SELECT COUNT(*) AS n FROM tickets {user_where}", user_args, one=True)
    total_rows = int(total_row["n"] if total_row else 0)
    all_ticket_row = query("SELECT COUNT(*) AS n FROM tickets", one=True)
    all_ticket_count = int(all_ticket_row["n"] if all_ticket_row else 0)
    page_num = max(1, min(int(page_num or 1), max(1, (total_rows + 9) // 10)))
    rows = query(f"SELECT * FROM tickets {user_where} ORDER BY {status_order_sql} ASC, {severity_order_sql} ASC, {created_sort} DESC, id DESC LIMIT 10 OFFSET ?",
                 user_args + ((page_num - 1) * 10,))
    status_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(x)}" {"selected" if status==x else ""}>{E(status_label(req,x))}</option>' for x in STATUSES)
    device_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(x)}" {"selected" if device==x else ""}>{E(device_label(req,x))}</option>' for x in DEVICES)
    dept_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(d)}" {"selected" if department==d else ""}>{E(department_label(req,d))}</option>' for d in department_options())
    severity_order = ["Cao", "Trung bình", "Thấp"]
    severity_option_colors = {"Cao":"#c62e3b", "Trung bình":"#a15b00", "Thấp":"#087e68"}
    severity_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(x)}" style="color:{severity_option_colors[x]};font-weight:700" {"selected" if severity==x else ""}>{E(sev_label(req,x))}</option>' for x in severity_order)
    issue_filter_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(x)}" {"selected" if issue_type==x else ""}>{E(issue_label(req,x))}</option>' for x in ISSUE_TYPES)
    today_iso = now_dt().strftime("%Y-%m-%d")
    day_display = datetime.strptime(day, "%Y-%m-%d").strftime("%d/%m/%Y") if day else ""
    english = req.session.get("lang") == "en"
    all_label = "All tickets" if english else "Tất cả ticket"
    own_label = "My tickets" if english else "Ticket của tôi"
    day_label = "Created date" if english else "Ngày tạo"
    date_field = f'<div><label>{day_label}</label><div class="calendar-pick-wrap"><input class="w date-text-input" type="text" name="day" id="day-text" value="{E(day_display)}" placeholder="DD/MM/YYYY" inputmode="numeric" pattern="[0-9]{{2}}/[0-9]{{2}}/[0-9]{{4}}" autocomplete="off"><input type="date" id="day-picker" aria-label="{day_label}" style="position:absolute;width:1px;height:1px;opacity:0;pointer-events:none"><button type="button" class="calendar-pick-button" id="open-day-picker" title="{day_label}" aria-label="{day_label}"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18"/><path d="M8 14h.01M12 14h.01M16 14h.01M8 17h.01M12 17h.01"/></svg></button></div></div>'
    body = fr'''<section class="box ticket-toolbar-panel"><div class="ticket-list-head"><h2><span class="section-icon"><svg viewBox="0 0 24 24" width="19" height="19" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8M8 12h8M8 16h5"/></svg></span>{tr(req,"my")}</h2></div><div class="quick-toolbar"><a class="quick-filter" href="/?day={today_iso}&all_tickets=1"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18"/><path d="M8 14h.01M12 14h.01M16 14h.01M8 17h.01M12 17h.01"/></svg> {tr(req,"today")}</a><a class="quick-filter" href="/?all_tickets=0"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="4" width="16" height="16" rx="3"/><path d="M8 9h8M8 13h8M8 17h5"/></svg> {own_label}</a><a class="quick-filter" href="/?all_tickets=1"><svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 6h13M8 12h13M8 18h13"/><path d="m3 6 1 1 2-2M3 12l1 1 2-2M3 18l1 1 2-2"/></svg> {all_label}</a></div><form method="get" class="ticket-filter-form"><input type="hidden" name="all_tickets" value="{1 if all_tickets else 0}"><div class="filter-grid compact-filter-grid user-filter-grid"><div><label>{tr(req,"dept")}</label><select class="w" name="department">{dept_opts}</select></div><div><label>{tr(req,"status")}</label><select class="w" name="status">{status_opts}</select></div><div><label>{tr(req,"device")}</label><select class="w" name="device">{device_opts}</select></div><div><label>{tr(req,"issue_type")}</label><select class="w" name="issue_type">{issue_filter_opts}</select></div>{date_field}<div><label>{tr(req,"severity")}</label><select class="w" name="severity">{severity_opts}</select></div><div class="filter-actions"><button type="submit"><span aria-hidden="true">⌕</span> {tr(req,"filter")}</button><a class="btn btn-reset" href="/?all_tickets={1 if all_tickets else 0}">{tr(req,"all")}</a></div></div></form></section><section class="box ticket-table-panel">{ticket_rows(req,rows,all_tickets=bool(all_tickets))}{ticket_pagination(req,total_rows,page_num,all_total=all_ticket_count,show_all_count=False)}</section><div class="admin-home-bottom"><a href="/" class="admin-home-icon" title="{'Back to home' if req.session.get('lang') == 'en' else 'Quay về trang chủ'}" aria-label="{'Back to home' if req.session.get('lang') == 'en' else 'Quay về trang chủ'}"><svg viewBox="0 0 24 24" width="21" height="21" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m3 10 9-7 9 7"/><path d="M5 9v11h14V9M9 20v-6h6v6"/></svg></a><style>.admin-home-bottom{{display:flex;justify-content:center;align-items:center;margin:24px 0 8px;padding-top:14px;border-top:1px solid #e4edf7}}.admin-home-icon{{display:inline-flex;align-items:center;justify-content:center;width:42px;height:42px;border-radius:50%;background:#edf6ff;color:#176fc3;border:1px solid #d6e8fa;box-shadow:0 3px 10px rgba(30,90,150,.08);text-decoration:none;transition:background .18s,transform .18s}}.admin-home-icon:hover{{background:#dceeff;transform:translateY(-1px)}}.admin-home-icon:focus-visible{{outline:3px solid #93c5fd;outline-offset:3px}}</style><script>(function(){{const text=document.getElementById('day-text'), picker=document.getElementById('day-picker'), button=document.getElementById('open-day-picker');function isoToDmy(v){{if(!v)return '';const p=v.split('-');return p[2]+'/'+p[1]+'/'+p[0];}}function dmyToIso(v){{const m=/^(\d{{2}})\/(\d{{2}})\/(\d{{4}})$/.exec(v);return m?m[3]+'-'+m[2]+'-'+m[1]:'';}}if(text&&picker&&button){{picker.value=dmyToIso(text.value);button.addEventListener('click',function(){{if(picker.showPicker)picker.showPicker();else picker.click();}});picker.addEventListener('change',function(){{text.value=isoToDmy(this.value);}});text.addEventListener('change',function(){{const v=dmyToIso(this.value);if(v)picker.value=v;}});}}}})();</script>'''
    body += '<script>(function(){let signature="";async function syncTickets(){try{const r=await fetch("/api/tickets/sync",{cache:"no-store",credentials:"same-origin"});if(!r.ok)return;const j=await r.json();if(signature&&signature!==j.signature){location.reload();return}signature=j.signature}catch(e){}}syncTickets();setInterval(syncTickets,10000)})();</script>'
    return page(req, tr(req, "my"), (f'<p class="msg">{E(msg)}</p>' if msg else '') + body, u)
def account_option_labels(req, accounts):
    """Build readable bilingual owner labels while keeping real account IDs as option values."""
    prepared = []
    for account in accounts:
        username = (account["username"] or "").strip()
        full_name = (account["full_name"] or "").strip()
        dept = normalize_department(account["department"])
        dept_label = department_label(req, dept)
        base = f"{username} ({dept_label}) - {full_name}" if full_name else f"{username} ({dept_label})"
        prepared.append((account, base, dept))
    counts = {}
    for _, label, _ in prepared:
        counts[label.casefold()] = counts.get(label.casefold(), 0) + 1
    result = []
    for account, base, dept in prepared:
        label = base
        if counts.get(base.casefold(), 0) > 1:
            phone = (account["phone"] or "").strip() if "phone" in account.keys() else ""
            label = f"{base} - {phone}" if phone else f"{base} - ID {account['id']}"
        result.append((account, label, dept))
    return result


@app.get("/new", response_class=HTMLResponse)
def new_form(req: Request, u=Depends(current_user)):
    # Keep the v9 ticket creation form and POST workflow intact; only color the severity choices.
    severity_colors = {"Cao": "#b42332", "Trung bình": "#94610f", "Thấp": "#1d4ed8"}
    opts = "".join(f'<option value="{E(s)}" data-color="{severity_colors.get(s, "#193b67")}" style="color:{severity_colors.get(s, "#193b67")};font-weight:700">{E(sev_label(req,s))}</option>' for s in SEVERITIES)
    device_opts = "".join(f'<option value="{E(s)}">{E(device_label(req,s))}</option>' for s in DEVICES)
    issue_opts = "".join(f'<option value="{E(s)}">{E(issue_label(req,s))}</option>' for s in ISSUE_TYPES)
    def field(key, name, value=""):
        return f'<div><label>{tr(req,key)}</label><input class="w" name="{name}" value="{E(value or "")}" required maxlength="300"></div>'
    user_dept = normalize_department(u["department"])
    dept_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(d)}" {"selected" if user_dept.casefold()==d.casefold() else ""}>{E(department_label(req,d))}</option>' for d in department_options())
    can_create_for_other = u["role"] == "admin" or user_dept == "IT"
    requested_for_html = ""
    if can_create_for_other:
        accounts = query("SELECT id, username, full_name, department, phone FROM users WHERE active=1 ORDER BY department, full_name, username")
        account_opts = []
        for account, label, account_dept in account_option_labels(req, accounts):
            account_opts.append(f'<option value="{account["id"]}" data-name="{E((account["full_name"] or account["username"]).strip())}" data-department="{E(account_dept)}" data-phone="{E(account["phone"] or "")}" {"selected" if int(account["id"]) == int(u["id"]) else ""}>{E(label)}</option>')
        requested_label = "Ticket for account" if req.session.get("lang") == "en" else "Tạo hộ tài khoản"
        requested_help = "Select the account this ticket belongs to. The creator is recorded separately." if req.session.get("lang") == "en" else "Chọn tài khoản sở hữu ticket; hệ thống lưu riêng tài khoản tạo hộ."
        requested_for_html = f'<div class="requested-for-field"><label for="requested_for_user_id">{requested_label}</label><select class="w" id="requested_for_user_id" name="requested_for_user_id">{"".join(account_opts)}</select><small>{requested_help}</small></div>'
    body = f'<div class="box create-ticket-box"><h2><span class="section-icon">＋</span>{tr(req,"new")}</h2><form method="post">{requested_for_html}<div class="grid"><div><label>{tr(req,"device")}</label><select class="w" name="device" required>{device_opts}</select></div>{field("device_code","device_code")}<div><label>{tr(req,"dept")}</label><select class="w" name="department" id="ticket-department" required>{dept_opts}</select></div>{field("location","location")}<div><label>{tr(req,"reporter")}</label><input class="w" id="ticket-reporter" name="reporter" value="{E(u["full_name"] or u["username"])}" required maxlength="300"></div><div><label>{tr(req,"phone")}</label><input class="w" id="ticket-phone" name="phone" value="{E(u["phone"] or "")}" maxlength="50"></div></div><label>{tr(req,"issue_type")}</label><select class="w" name="issue_type" required>{issue_opts}</select><label>{tr(req,"issue")}</label><textarea class="w" name="issue" rows="4" maxlength="2000" required></textarea><label>{tr(req,"severity")}</label><select class="w" name="severity" required style="color:{severity_colors.get(SEVERITIES[0], "#193b67")}" onchange="this.style.color=this.options[this.selectedIndex].dataset.color||&#39;#193b67&#39;">{opts}</select><p><button>{tr(req,"submit")}</button></p></form></div><script>(function(){{const target=document.getElementById("requested_for_user_id");if(!target)return;target.addEventListener("change",function(){{const o=this.options[this.selectedIndex];const reporter=document.getElementById("ticket-reporter"),phone=document.getElementById("ticket-phone"),dept=document.getElementById("ticket-department");if(reporter)reporter.value=o.dataset.name||reporter.value;if(phone)phone.value=o.dataset.phone||"";if(dept&&o.dataset.department){{for(const d of dept.options)if(d.value===o.dataset.department){{dept.value=d.value;break;}}}}}});}})();</script>'
    return page(req, tr(req, "new"), body, u)

@app.post("/new")
def new_submit(req: Request, background_tasks: BackgroundTasks, device: str = Form(...), device_code: str = Form(...), department: str = Form(...), location: str = Form(...), issue: str = Form(...), issue_type: str = Form(...), severity: str = Form(...), reporter: str = Form(...), phone: str = Form(""), requested_for_user_id: int = Form(0), u=Depends(current_user)):
    department = normalize_department(department)
    can_create_for_other = u["role"] == "admin" or normalize_department(u["department"]) == "IT"
    owner_id = int(u["id"])
    if can_create_for_other and requested_for_user_id:
        target_user = query("SELECT id, active FROM users WHERE id=?", (requested_for_user_id,), one=True)
        if not target_user or not target_user["active"]:
            raise HTTPException(400, tr(req, "invalid_data"))
        owner_id = int(target_user["id"])
    if device not in DEVICES or severity not in SEVERITIES or issue_type not in ISSUE_TYPES or department not in department_options():
        raise HTTPException(400, tr(req, "invalid_data"))
    created_dt = now_dt()  # Capture once so the ticket timestamp and ID year agree.
    created_at = created_dt.strftime("%d/%m/%Y %H:%M:%S")
    with closing(db()) as c:
        cur = c.execute("INSERT INTO tickets(created_at,user_id,device,device_code,department,location,issue,issue_type,severity,reporter,phone,created_by_user_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
            (created_at,owner_id,device.strip()[:300],device_code.strip()[:300],department.strip()[:300],location.strip()[:300],issue.strip()[:2000],issue_type,severity,reporter.strip()[:300],phone.strip()[:50],u["id"]))
        tid = cur.lastrowid
        department_initials = {
            "Bác sĩ": "D", "Lễ tân": "R", "Điều dưỡng": "N", "CĐHA": "X",
            "Xét nghiệm": "L", "Hành chính": "A", "Dược": "P", "Kế toán": "K", "IT": "I", "Khác": "O",
        }
        canonical_department = normalize_department(department)
        code_prefix = f"IT{department_initials.get(canonical_department, 'O')}{created_dt.strftime('%d%m')}"
        existing_count = c.execute(
            "SELECT COUNT(*) FROM tickets WHERE code LIKE ? AND id<>?",
            (code_prefix + '-%', tid),
        ).fetchone()[0]
        code = f"{code_prefix}-{existing_count + 1:02d}"
        c.execute("UPDATE tickets SET code=? WHERE id=?", (code, tid))
        c.commit()
    ticket_for_notification = query("SELECT * FROM tickets WHERE id=?", (tid,), one=True)
    if ticket_for_notification:
        ticket_payload = dict(ticket_for_notification)
        background_tasks.add_task(dispatch_ticket_notification, ticket_payload, "new_ticket")
        if str(ticket_payload.get("severity") or "").strip().casefold() in ("cao", "high"):
            background_tasks.add_task(dispatch_ticket_notification, ticket_payload, "high_ticket")
    return RedirectResponse(f"/ticket/{tid}", 303)

@app.get("/ticket/{tid}", response_class=HTMLResponse)
def detail(req: Request, tid: int, all_tickets: int = 0, u=Depends(current_user)):
    t = query("SELECT * FROM tickets WHERE id=?", (tid,), one=True)
    # Users may open other departments/admin-created tickets only from the explicit All tickets view.
    can_view = bool(t) and (u["role"] == "admin" or t["user_id"] == u["id"] or all_tickets == 1)
    if not can_view: raise HTTPException(404, tr(req, "not_found"))
    fields = [("device","device"),("device_code","device_code"),("dept","department"),("location","location"),("issue_type","issue_type"),("issue","issue"),("severity","severity"),("reporter","reporter"),("phone","phone"),("time","created_at"),("handler","handler"),("received","received_at")]
    fields.append(("closed","closed_at"))
    show_user_note = (u["role"] != "admin" and t["user_id"] == u["id"] and bool(t["received_at"]))
    if u["role"] == "admin" or show_user_note:
        fields.append(("note","public_note"))
    stcls = {"Mới":"st-new","Đang xử lý":"st-progress","Đã đóng":"st-closed"}.get(t["status"], "st-new")
    priority_class = {"Thấp":"sev-low", "Trung bình":"sev-medium", "Cao":"sev-high"}.get(t["severity"], "sev-medium")
    created_caption = "Created" if req.session.get("lang") == "en" else "Tạo lúc"
    inline_edit = ""
    if u["role"] == "admin" or t["user_id"] == u["id"]:
        def edit_select(name, options, current, labels=None):
            opts_html = ''.join(f'<option value="{E(opt)}" {"selected" if opt == current else ""}>{E((labels or {}).get(opt, opt))}</option>' for opt in options)
            return f'<div><label>{tr(req, name)}</label><select class="w" name="{name}" required>{opts_html}</select></div>'
        edit_labels = {
            "device": {x: device_label(req, x) for x in DEVICES},
            "department": {x: department_label(req, x) for x in department_options()},
            "issue_type": {x: issue_label(req, x) for x in ISSUE_TYPES},
            "severity": {x: sev_label(req, x) for x in SEVERITIES},
        }
        edit_title = "Edit ticket" if req.session.get("lang") == "en" else "Chỉnh sửa ticket"
        edit_submit = "Save changes" if req.session.get("lang") == "en" else "Lưu thay đổi"
        owner_select_html = ""
        if u["role"] == "admin":
            owner_accounts = query("SELECT id, username, full_name, department, phone FROM users WHERE active=1 ORDER BY department, full_name, username")
            owner_options = []
            for account, account_label, _account_dept in account_option_labels(req, owner_accounts):
                owner_options.append(f'<option value="{account["id"]}" {"selected" if int(account["id"]) == int(t["user_id"] or 0) else ""}>{E(account_label)}</option>')
            owner_label = "Ticket belongs to account" if req.session.get("lang") == "en" else "Tài khoản sở hữu ticket"
            owner_select_html = f'<div class="ticket-owner-select"><label for="requested_for_user_id">{owner_label}</label><select class="w" name="requested_for_user_id" id="requested_for_user_id" required>{"".join(owner_options)}</select></div>'
        delete_confirm = "Delete this ticket permanently?" if req.session.get("lang") == "en" else "Xóa ticket này vĩnh viễn?"
        delete_label = "Delete ticket" if req.session.get("lang") == "en" else "Xóa ticket"
        delete_action = (f'<form class="ticket-delete-inline" method="post" action="/ticket/{tid}/delete" onsubmit="return confirm(&quot;{delete_confirm}&quot;)"><button type="submit" class="ticket-delete-icon" title="{delete_label}" aria-label="{delete_label}"><svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 6h18M8 6V4h8v2M6 6l1 15h10l1-15M10 10v7M14 10v7"/></svg></button></form>' if u["role"] == "admin" else '')
        inline_edit = f'<div class="ticket-heading-edit"><span>{"Edit" if req.session.get("lang") == "en" else "Sửa ticket"}</span><details class="inline-ticket-edit"><summary title="{edit_title}" aria-label="{edit_title}"><svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m14 5 5 5M4 20l4.5-1 11-11a2.1 2.1 0 0 0-3-3l-11 11Z"/></svg></summary><form method="post" action="/ticket/{tid}/edit">{owner_select_html}<div class="grid">{edit_select("device", DEVICES, t["device"] or DEVICES[-1], edit_labels["device"])}<div><label>{tr(req,"device_code")}</label><input class="w" name="device_code" value="{E(t["device_code"] or "")}"></div>{edit_select("department", department_options(), t["department"] or DEPARTMENTS[-1], edit_labels["department"])}<div><label>{tr(req,"location")}</label><input class="w" name="location" value="{E(t["location"] or "")}"></div><div><label>{tr(req,"reporter")}</label><input class="w" name="reporter" value="{E(t["reporter"] or "")}" required></div><div><label>{tr(req,"phone")}</label><input class="w" name="phone" value="{E(t["phone"] or "")}"></div>{edit_select("issue_type", ISSUE_TYPES, t["issue_type"] or ISSUE_TYPES[-1], edit_labels["issue_type"])}{edit_select("severity", SEVERITIES, t["severity"] or SEVERITIES[0], edit_labels["severity"])}</div><label>{tr(req,"issue")}</label><textarea class="w" name="issue" rows="3" maxlength="2000" required>{E(t["issue"] or "")}</textarea><p style="margin:9px 0 0"><button type="submit">{edit_submit}</button></p></form></details>{delete_action}</div>'
    body = f'<div class="ticket-heading-card"><div class="ticket-heading-icon" aria-hidden="true"><svg viewBox="0 0 24 24" width="27" height="27" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 7h8M8 11h8M8 15h5"/></svg></div><div class="ticket-heading-copy"><div class="ticket-heading-line"><h2>{E(t["code"] or "")}</h2><span class="pill {stcls}"><span class="pill-icon" aria-hidden="true">{status_icon_svg(t["status"], 13)}</span>{E(status_label(req,t["status"]))}</span></div><p>{created_caption}: {E(t["created_at"] or "—")}</p>{inline_edit}</div><div class="ticket-heading-actions"><span class="priority-pill {priority_class}"><span class="priority-shield" aria-hidden="true">{severity_icon_svg(t["severity"])}</span> {E(sev_label(req,t["severity"] or ""))}</span></div></div><div class="box detail-box"><div class="grid detail-grid">'
    detail_icons = {
        "device": '<rect x="3" y="4" width="18" height="14" rx="2"/><path d="M8 21h8M12 18v3"/>',
        "device_code": '<rect x="5" y="3" width="14" height="18" rx="2"/><path d="M9 7h6M9 11h3M9 15h6"/>',
        "department": '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M9 7h2v2H9zM13 7h2v2h-2zM9 11h2v2H9zM13 11h2v2h-2zM10 21v-4h4v4"/>',
        "location": '<path d="M20 10c0 5-8 11-8 11S4 15 4 10a8 8 0 1 1 16 0Z"/><circle cx="12" cy="10" r="2.5"/>',
        "issue_type": '<path d="M12 3 22 20H2L12 3Z"/><path d="M12 9v4M12 16h.01"/>',
        "issue": '<path d="M6 3h9l4 4v14H6z"/><path d="M14 3v5h5M9 12h7M9 16h7"/>',
        "severity": '<path d="M4 19V10M10 19V5M16 19v-7M22 19V8"/>',
        "reporter": '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',
        "phone": '<path d="M7 3H4a1 1 0 0 0-1 1c0 9 8 17 17 17a1 1 0 0 0 1-1v-3l-5-2-2 3a14 14 0 0 1-5-5l3-2-2-5Z"/>',
        "created_at": '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M7 3v4M17 3v4M3 10h18"/>',
        "handler": '<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0"/><path d="m16 13 2 2 4-4"/>',
        "received_at": '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
        "closed_at": '<circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>',
        "note": '<path d="m4 16-.8 4.8L8 20l11-11-4-4L4 16Z"/><path d="m13.5 6.5 4 4"/>'
    }
    def detail_icon(name):
        path = detail_icons.get(name, '<circle cx="12" cy="12" r="8"/>')
        return f'<svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{path}</svg>'
    for key, col in fields:
        val = t[col] or ""
        if col == "severity":
            sevcls = {"Thấp":"sev-low","Trung bình":"sev-medium","Cao":"sev-high"}.get(val, "sev-medium")
            val = f'<span class="pill {sevcls}"><span class="pill-icon" aria-hidden="true">{severity_icon_svg(val)}</span>{E(sev_label(req,val))}</span>'
        elif col == "device":
            val = f'<span class="detail-value-chip value-device"><svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="4" width="18" height="13" rx="2"/><path d="M8 21h8M12 17v4"/></svg>{E(device_label(req,val or "—"))}</span>'
        elif col == "device_code":
            val = f'<span class="detail-value-chip value-device-code"><svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8M8 12h4M8 16h8"/></svg>{E(val or "—")}</span>'
        elif col == "phone":
            val = f'<span class="detail-value-chip value-phone"><svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M7 3H4a1 1 0 0 0-1 1c0 9 8 17 17 17a1 1 0 0 0 1-1v-3l-5-2-2 3a14 14 0 0 1-5-5l3-2-2-5Z"/></svg>{E(val or "—")}</span'
        elif col == "issue_type":
            issue_type_text = issue_label(req, val)
            val = f'<span class="issue-type-chip"><svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3 22 20H2L12 3Z"/><path d="M12 9v4M12 16h.01"/></svg>{E(issue_type_text or "—")}</span>'
        elif col == "issue":
            val = f'<span class="issue-description-chip">{E(val or "—")}</span>'
        elif col == "department":
            val = f'<span class="detail-tag detail-tag-dept"><svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M9 7h2v2H9zM13 7h2v2h-2zM9 11h2v2H9zM13 11h2v2h-2zM10 21v-4h4v4"/></svg> {E(department_label(req,val or "—"))}</span>'
        elif col == "location":
            val = f'<span class="detail-tag detail-tag-location"><svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M20 10c0 5-8 11-8 11S4 15 4 10a8 8 0 1 1 16 0Z"/><circle cx="12" cy="10" r="2.5"/></svg> {E(val or "—")}</span>'
        elif col == "reporter":
            val = f'<span class="reporter-chip"><svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/></svg> {E(val or "—")}</span>'
        elif col == "handler":
            handler_name = val or tr(req,"assigned_none")
            val = f'<span class="handler-value">{E(handler_name)}</span>'
        elif col == "created_at":
            if val:
                val = f'<span class="detail-time detail-time-created"><svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M7 3v4M17 3v4M3 10h18"/></svg>{E(val)}</span>'
            else:
                val = f'<span class="detail-time detail-time-pending">—</span>'
        elif col == "received_at":
            if val:
                val = f'<span class="detail-time detail-time-progress"><svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>{E(val)}</span>'
            else:
                val = f'<span class="detail-time detail-time-pending">{tr(req,"pending_received")}</span>'
        elif col == "closed_at":
            if val:
                val = f'<span class="detail-time detail-time-closed"><svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/></svg>{E(val)}</span>'
            else:
                val = f'<span class="detail-time detail-time-pending">{tr(req,"pending_closed")}</span>'
        elif col == "public_note":
            note_text = t["public_note"] or ""
            en_note = req.session.get("lang") == "en"
            note_label = "Handling note" if en_note else "Ghi chú xử lý"
            note_empty = "No handling note yet…" if en_note else "Chưa có ghi chú xử lý…"
            if u["role"] == "admin":
                note_placeholder = "Enter handling note…" if en_note else "Nhập nội dung ghi chú xử lý…"
                edit_tip = "Edit note" if en_note else "Chỉnh sửa ghi chú"
                save_tip = "Save note" if en_note else "Lưu ghi chú"
                cancel_tip = "Cancel editing" if en_note else "Hủy chỉnh sửa"
                edit_icon = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m4 16-.8 4.8L8 20 19 9l-4-4L4 16Z"/><path d="m13.5 6.5 4 4"/></svg>'
                save_icon = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m5 12 4 4L19 6"/></svg>'
                # Use a modern undo/return symbol for cancel; all three compact actions stay visible.
                cancel_icon = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 14 4 9l5-5"/><path d="M4 9h9a7 7 0 0 1 0 14h-2"/></svg>'
                shown_note = f'<span class="note-value-chip">{E(note_text)}</span>' if note_text else f'<span class="note-empty">{note_empty}</span>'
                val = (f'<div class="note-field-content"><div class="public-note-view" id="note-view-{tid}">{shown_note}</div>'
                       f'<form class="public-note-form" id="note-form-{tid}" method="post" action="/ticket/{tid}/public-note" hidden>'
                       f'<textarea id="public-note-{tid}" name="public_note" rows="3" maxlength="2000" data-original="{E(note_text)}" placeholder="{note_placeholder}">{E(note_text)}</textarea></form></div>')
            else:
                shown_note = f'<span class="note-value-chip">{E(note_text)}</span>' if note_text else f'<span class="note-empty">{note_empty}</span>'
                val = f'<div class="note-field-content"><div class="public-note-view">{shown_note}</div></div>'
        else: val = E(val)
        field_class = "detail-field-note" if col == "public_note" else f"detail-field-{col}"
        if col == "public_note":
            if u["role"] == "admin":
                note_action_controls = (f'<div class="note-heading-actions">'
                    f'<button type="button" id="note-edit-btn-{tid}" class="note-icon-btn note-edit" title="{edit_tip}" aria-label="{edit_tip}" onclick="editPublicNote({tid})">{edit_icon}</button>'
                    f'<button type="button" id="note-save-btn-{tid}" class="note-icon-btn note-save" title="{save_tip}" aria-label="{save_tip}" disabled onclick="savePublicNote({tid})">{save_icon}</button>'
                    f'<button type="button" id="note-cancel-btn-{tid}" class="note-icon-btn note-cancel" title="{cancel_tip}" aria-label="{cancel_tip}" disabled onclick="cancelPublicNote({tid})">{cancel_icon}</button></div>')
            else:
                note_action_controls = ""
            # Keep the note title clean; actions are small, aligned horizontally on the same row.
            body += f'<div class="detail-field {field_class}"><div class="note-heading"><strong><span>{note_label}</span></strong>{note_action_controls}</div>{val}</div>'
        else:
            body += f'<div class="detail-field {field_class}"><strong><span class="detail-icon">{detail_icon(col)}</span><span>{tr(req,key)}</span></strong><p>{val}</p></div>'
    body += '</div></div>'
    if u["role"] == "admin" and any(col == "public_note" for _, col in fields):
        body += """<script>
        function setPublicNoteActionState(id,editing){const edit=document.getElementById('note-edit-btn-'+id),save=document.getElementById('note-save-btn-'+id),cancel=document.getElementById('note-cancel-btn-'+id);if(edit)edit.disabled=!!editing;if(save)save.disabled=!editing;if(cancel)cancel.disabled=!editing;}
        function editPublicNote(id){const view=document.getElementById('note-view-'+id),form=document.getElementById('note-form-'+id),area=document.getElementById('public-note-'+id);if(!view||!form||!area)return;area.dataset.original=area.value;view.hidden=true;form.hidden=false;setPublicNoteActionState(id,true);area.focus();}
        function savePublicNote(id){const form=document.getElementById('note-form-'+id);if(form&&!form.hidden&&typeof form.requestSubmit==='function')form.requestSubmit();else if(form&&!form.hidden)form.submit();}
        function cancelPublicNote(id){const view=document.getElementById('note-view-'+id),form=document.getElementById('note-form-'+id),area=document.getElementById('public-note-'+id);if(!view||!form||!area)return;area.value=area.dataset.original||'';form.hidden=true;view.hidden=false;setPublicNoteActionState(id,false);}
        </script>"""
    feedback_history = query("SELECT rating,comment,created_at,reviewer_name FROM ticket_feedback WHERE ticket_id=?", (tid,), one=True)
    if feedback_history and u["role"] == "admin":
        en_history = req.session.get("lang") == "en"
        history_title = "User rating" if en_history else "Đánh giá người dùng"
        history_by = "Rated by" if en_history else "Người đánh giá"
        history_comment = feedback_history["comment"] or ("No comment provided." if en_history else "User không để lại nhận xét.")
        history_color = "#16815b" if feedback_history["rating"] >= 4 else "#b77908" if feedback_history["rating"] == 3 else "#c2413b"
        history_stars = "★" * feedback_history["rating"] + "☆" * (5 - feedback_history["rating"])
        history_time_label = "Rated at" if en_history else "Đánh giá lúc"
        body += f'<section class="csat-history-panel" aria-label="{history_title}"><h3>{history_title}</h3><div class="csat-history-stars" style="color:{history_color}">{history_stars} <span style="font-size:13px;letter-spacing:0">({feedback_history["rating"]}/5)</span></div><p class="csat-history-comment">{E(history_comment)}</p><div class="csat-history-time">{history_by}: {E(feedback_history["reviewer_name"] or "User")} · {history_time_label}: {E(feedback_history["created_at"] or "—")}</div></section>'
    if u["role"] == "admin":
        body += f'<section class="workflow-panel"><h3 class="workflow-heading"><span class="workflow-mark" aria-hidden="true"><svg viewBox="0 0 24 24" width="19" height="19" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M9 6h11M9 12h11M9 18h11"/><path d="m3.5 6 1.2 1.2L7 4.8M3.5 12l1.2 1.2L7 10.8M3.5 18l1.2 1.2L7 16.8"/></svg></span>{tr(req,"admin")} · {tr(req,"status")}</h3><div class="workflow-grid">'
        if t["status"] == "Mới":
            body += f'<div class="workflow-card workflow-pending"><div class="workflow-card-title"><span class="workflow-icon workflow-icon-blue">1</span><span>{tr(req,"take")}</span></div><div class="workflow-meta">{tr(req,"status")}: {E(status_label(req,t["status"]))}</div><div class="workflow-actions"><form method="post" action="/ticket/{tid}/take"><button class="btn-take" type="submit">✓ {tr(req,"take")}</button></form></div></div>'
        elif t["status"] == "Đang xử lý":
            body += f'<div class="workflow-card workflow-complete"><div class="workflow-card-title"><span class="workflow-icon workflow-icon-green">✓</span><span>{tr(req,"taken")}</span></div><div class="workflow-time">{E(t["received_at"] or "—")}</div><div class="workflow-meta">{tr(req,"handler")}: {E(t["handler"] or "—")}</div></div>'
            complete_label = "Complete handling" if req.session.get("lang") == "en" else "Hoàn tất xử lý"
            complete_help = "Finish the work to close this ticket and notify the user." if req.session.get("lang") == "en" else "Hoàn tất công việc để chuyển ticket sang trạng thái Đã đóng và thông báo cho User."
            body += f'<div class="workflow-card workflow-close-action"><div class="workflow-card-title"><span class="workflow-icon workflow-icon-blue">✓</span><span>{complete_label}</span></div><form method="post" action="/ticket/{tid}/close"><div class="workflow-actions"><button class="btn-close" type="submit">✓ {complete_label}</button><small class="workflow-note-help">{complete_help}</small></div></form></div>' 
        else:
            body += f'<div class="workflow-card workflow-complete"><div class="workflow-card-title"><span class="workflow-icon workflow-icon-green">✓</span><span>{tr(req,"closed_done")}</span></div><div class="workflow-time">{E(t["closed_at"] or "—")}</div><div class="workflow-meta">{tr(req,"handler")}: {E(t["handler"] or "—")}</div></div>'
        body += '</div></section>'
    if u["role"] != "admin" and t["user_id"] == u["id"] and t["status"] == "Đã đóng":
        fb = query("SELECT rating,comment,created_at FROM ticket_feedback WHERE ticket_id=?", (tid,), one=True)
        en = req.session.get("lang") == "en"
        rating_label = "Your rating" if en else "Đánh giá của bạn"
        save_label = "Save changes" if en else "Lưu chỉnh sửa"
        optional_label = "Comment (required for ratings below 3)" if en else "Nhận xét (bắt buộc nếu đánh giá dưới 3 sao)"
        edit_title = "Edit your rating" if en else "Chỉnh sửa đánh giá của bạn"
        stars_current = fb["rating"] if fb else 0
        rating_color = "#16815b" if stars_current >= 4 else "#b77908" if stars_current == 3 else "#c2413b"
        stars_display = "★" * stars_current + "☆" * (5 - stars_current)
        stars_edit = "".join(f'<button type="button" class="csat-star {"selected" if i <= stars_current else ""}" aria-label="{i} {"stars" if en else "sao"}" onclick="setEditCsat({i})">{"★" if i <= stars_current else "☆"}</button>' for i in range(1,6))
        shown_comment = (fb["comment"] or ("No comment provided" if en else "Chưa có nhận xét")) if fb else ("You have not rated this ticket yet." if en else "Bạn chưa đánh giá ticket này.")
        rated_line = (("You rated this ticket " if en else "Bạn đã đánh giá ticket này ") + f'{stars_current}/5 ⭐ · {"on" if en else "vào lúc"} {E(fb["created_at"] or "—")}') if fb else ("Please rate your IT support experience." if en else "Mời bạn đánh giá trải nghiệm hỗ trợ IT.")
        if fb:
            body += f'''<div class="csat-user-layout">
          <section class="box csat-user-summary"><h3><span class="csat-heading-icon" aria-hidden="true">★</span>{rating_label}</h3>
          {f'<div class="csat-history-stars" style="color:{rating_color}">{stars_display} <span class="csat-score">({stars_current}/5)</span></div>' if fb else f'<p class="csat-not-rated">{"Not rated yet" if en else "Chưa có đánh giá"}</p>'}
          <p class="csat-history-comment">{E(shown_comment)}</p><p class="csat-summary-line">{rated_line}</p></section>
          <section class="box csat-user-edit"><h3><span class="csat-heading-icon csat-edit-icon" aria-hidden="true">✎</span>{edit_title if fb else ("Rate this ticket" if en else "Đánh giá ticket này")}</h3>
          <p class="csat-edit-hint">{"You can update your rating and comment here." if en else "Bạn có thể chọn lại số sao và chỉnh sửa nhận xét tại đây."}</p>
          <form method="post" action="/ticket/{tid}/feedback" id="csatEditForm"><div class="csat-stars" role="radiogroup" aria-label="{"Rating" if en else "Số sao đánh giá"}">{stars_edit}</div><input type="hidden" name="rating" id="editCsatRating" value="{stars_current}"><label for="editCsatComment">{optional_label}</label><textarea name="comment" id="editCsatComment" rows="3" maxlength="1500" placeholder="{"Comment is optional; blank clears the previous comment" if en else "Nhận xét không bắt buộc; để trống sẽ xóa nhận xét cũ"}"></textarea><div class="csat-actions"><button type="submit" id="csatEditSubmit" {"disabled" if not fb else ""}>{save_label if fb else ("Submit rating" if en else "Gửi đánh giá")}</button></div></form></section></div>
          <script>function setEditCsat(n){{const rating=document.getElementById('editCsatRating');if(!rating)return;rating.value=String(n);document.querySelectorAll('#csatEditForm .csat-star').forEach((b,i)=>{{const selected=i<n;b.textContent=selected?'★':'☆';b.classList.toggle('selected',selected);b.setAttribute('aria-pressed',selected?'true':'false')}});const c=document.getElementById('editCsatComment');if(c)c.required=false;const submit=document.getElementById('csatEditSubmit');if(submit)submit.disabled=false}}document.querySelectorAll('#csatEditForm .csat-star').forEach((b,i)=>{{b.addEventListener('click',()=>setEditCsat(i+1))}});</script>'''
        else:
            body += f'''<div class="csat-modal-backdrop" id="csatModal" role="dialog" aria-modal="true" aria-labelledby="csatTitle"><div class="csat-modal" onclick="event.stopPropagation()"><button class="csat-x" type="button" aria-label="{"Close" if en else "Đóng"}" onclick="closeCsat()">×</button><h2 id="csatTitle">{"Rate IT support" if en else "Đánh giá hỗ trợ IT"}</h2><p>{"How satisfied are you with this ticket?" if en else "Bạn thấy dịch vụ hỗ trợ này thế nào?"}</p><form method="post" action="/ticket/{tid}/feedback" id="csatForm"><div class="csat-stars" role="radiogroup" aria-label="{"Rating" if en else "Số sao đánh giá"}">{''.join(f'<button type="button" class="csat-star" data-rating="{i}" aria-label="{i} {"stars" if en else "sao"}" onclick="setCsat({i})">☆</button>' for i in range(1,6))}</div><input type="hidden" name="rating" id="csatRating" value=""><div id="csatCommentWrap" hidden><label for="csatComment">{"Tell us what went wrong (required for ratings below 3)" if en else "Vui lòng cho biết lý do (bắt buộc khi đánh giá dưới 3 sao)"}</label><textarea name="comment" id="csatComment" rows="3" maxlength="1500"></textarea></div><div class="csat-actions"><button type="button" class="btn-secondary" onclick="closeCsat()">{"Later" if en else "Để sau"}</button><button type="submit" id="csatSubmit" disabled>{"Submit rating" if en else "Gửi đánh giá"}</button></div></form></div></div><script>function closeCsat(){{const m=document.getElementById('csatModal');if(m)m.remove()}}document.getElementById('csatModal').addEventListener('click',function(e){{if(e.target===this)closeCsat()}});function setCsat(n){{document.getElementById('csatRating').value=n;document.querySelectorAll('#csatForm .csat-star').forEach((b,i)=>{{b.textContent=i<n?'★':'☆';b.classList.toggle('selected',i<n)}});const w=document.getElementById('csatCommentWrap'),c=document.getElementById('csatComment');w.hidden=n>=3;c.required=n<3;document.getElementById('csatSubmit').disabled=false}}</script>'''
    back_url = "/admin" if u["role"] == "admin" else ("/?all_tickets=1" if all_tickets == 1 else "/")
    back_label = "Back to home" if req.session.get("lang") == "en" else "Quay lại trang chủ"
    body += '<div class="detail-back-row"><a href="'+back_url+'" class="icon-back-button" title="'+back_label+'" aria-label="'+back_label+'"><svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M19 12H5M12 19l-7-7 7-7"/></svg></a></div>'
    body += f'<script>(function(){{let sig="";async function syncTicket(){{try{{const r=await fetch("/api/ticket/{tid}/state",{{cache:"no-store"}});if(!r.ok)return;const j=await r.json();const next=JSON.stringify(j);if(sig&&sig!==next&&!document.querySelector("#csatForm textarea:focus,#csatEditForm textarea:focus")){{location.reload();return}}sig=next}}catch(e){{}}}}syncTicket();setInterval(syncTicket,10000)}})();</script>'
    return page(req, t["code"] or "Ticket", body, u)

@app.post("/ticket/{tid}/edit")
def edit_ticket_details(req: Request, tid: int, background_tasks: BackgroundTasks,
                        device: str = Form(...), device_code: str = Form(""),
                        department: str = Form(...), location: str = Form(""),
                        reporter: str = Form(...), phone: str = Form(""),
                        issue_type: str = Form(...), severity: str = Form(...),
                        issue: str = Form(...), requested_for_user_id: int = Form(0), u=Depends(current_user)):
    t = query("SELECT id,user_id,severity FROM tickets WHERE id=?", (tid,), one=True)
    if not t:
        raise HTTPException(404, tr(req, "not_found"))
    if u["role"] != "admin" and t["user_id"] != u["id"]:
        raise HTTPException(403, tr(req, "admin_only"))
    department = normalize_department(department)
    if device not in DEVICES or department not in department_options() or issue_type not in ISSUE_TYPES or severity not in SEVERITIES:
        raise HTTPException(400, tr(req, "invalid_data"))
    if u["role"] == "admin" and requested_for_user_id:
        target_user = query("SELECT id, active FROM users WHERE id=?", (requested_for_user_id,), one=True)
        if not target_user or not target_user["active"]:
            raise HTTPException(400, tr(req, "invalid_data"))
        query("UPDATE tickets SET user_id=? WHERE id=?", (target_user["id"], tid), commit=True)
    query("UPDATE tickets SET device=?,device_code=?,department=?,location=?,reporter=?,phone=?,issue_type=?,severity=?,issue=? WHERE id=?",
          (device, device_code.strip()[:300], department, location.strip()[:300], reporter.strip()[:300], phone.strip()[:50], issue_type, severity, issue.strip()[:2000], tid), commit=True)
    updated_ticket = query("SELECT * FROM tickets WHERE id=?", (tid,), one=True)
    if updated_ticket:
        updated_payload = dict(updated_ticket)
        background_tasks.add_task(dispatch_ticket_notification, updated_payload, "ticket_updated")
        old_severity = str(t["severity"] or "").strip().casefold()
        new_severity = str(updated_payload.get("severity") or "").strip().casefold()
        if new_severity in ("cao", "high") and old_severity not in ("cao", "high"):
            background_tasks.add_task(dispatch_ticket_notification, updated_payload, "high_ticket")
    return RedirectResponse(f"/ticket/{tid}", 303)

@app.post("/ticket/{tid}/delete")
def delete_ticket(req: Request, tid: int, u=Depends(admin_user)):
    ticket = query("SELECT id FROM tickets WHERE id=?", (tid,), one=True)
    if not ticket:
        raise HTTPException(404, tr(req, "not_found"))
    query("DELETE FROM tickets WHERE id=?", (tid,), commit=True)
    return RedirectResponse("/admin?all_dates=1", 303)

@app.post("/ticket/{tid}/update")
def update_ticket(req: Request, tid: int, background_tasks: BackgroundTasks, action: str = Form(...), return_to: str = Form(""), u=Depends(admin_user)):
    t = query("SELECT * FROM tickets WHERE id=?", (tid,), one=True)
    if not t: raise HTTPException(404, tr(req, "not_found"))
    if t["status"] == "Đã đóng": raise HTTPException(400, tr(req, "ticket_closed"))
    who = u["full_name"] or u["username"]
    if action == "take":
        if t["status"] != "Mới":
            raise HTTPException(409, tr(req, "invalid_action"))
        received_at = now()
        # Atomic claim prevents two Admins from taking the same ticket.
        with closing(db()) as conn:
            cur = conn.execute(
                "UPDATE tickets SET status='Đang xử lý',handler=?,received_at=? WHERE id=? AND status='Mới'",
                (who, received_at, tid)
            )
            conn.commit()
            claimed = cur.rowcount
        if not claimed:
            raise HTTPException(409, tr(req, "invalid_action"))
    elif action == "close":
        if t["status"] != "Đang xử lý":
            raise HTTPException(400, tr(req, "must_take_first"))
        closed_at = now()
        query("UPDATE tickets SET status='Đã đóng',handler=COALESCE(handler,?),closed_at=? WHERE id=?", (who,closed_at,tid), commit=True)
    else: raise HTTPException(400, tr(req, "invalid_action"))
    event_key = {"take": "ticket_taken", "close": "ticket_closed"}.get(action)
    updated_ticket = query("SELECT * FROM tickets WHERE id=?", (tid,), one=True)
    if event_key and updated_ticket:
        payload = dict(updated_ticket)
        if action == "close":
            token_row = query("SELECT token FROM ticket_feedback_tokens WHERE ticket_id=?", (tid,), one=True)
            token = token_row["token"] if token_row else secrets.token_urlsafe(32)
            if not token_row:
                query("INSERT OR IGNORE INTO ticket_feedback_tokens(ticket_id,token,created_at) VALUES(?,?,?)", (tid, token, now()), commit=True)
                token_row = query("SELECT token FROM ticket_feedback_tokens WHERE ticket_id=?", (tid,), one=True)
                token = token_row["token"]
            base = os.getenv("APP_PUBLIC_URL", str(req.base_url).rstrip("/")).rstrip("/")
            payload["feedback_url"] = f"{base}/feedback/{token}"
        background_tasks.add_task(dispatch_ticket_notification, payload, event_key)
    if action == "take" and return_to.startswith("/") and not return_to.startswith("//"):
        return RedirectResponse(return_to, 303)
    return RedirectResponse(f"/ticket/{tid}", 303)

@app.post("/ticket/{tid}/take")
def take_ticket_route(req: Request, tid: int, background_tasks: BackgroundTasks, return_to: str = Form(""), u=Depends(admin_user)):
    # Dedicated POST endpoint used by detail page and inline Admin list action.
    return update_ticket(req, tid, background_tasks, action="take", return_to=return_to, u=u)

@app.post("/ticket/{tid}/close")
def close_ticket_route(req: Request, tid: int, background_tasks: BackgroundTasks, u=Depends(admin_user)):
    # Dedicated POST endpoint used by the Close Ticket button.
    return update_ticket(req, tid, background_tasks, action="close", u=u)

@app.get("/feedback/{token}", response_class=HTMLResponse)
def public_feedback(req: Request, token: str, msg: str = ""):
    row = query("SELECT t.*, f.rating AS feedback_rating FROM ticket_feedback_tokens k JOIN tickets t ON t.id=k.ticket_id LEFT JOIN ticket_feedback f ON f.ticket_id=t.id WHERE k.token=?", (token,), one=True)
    if not row or row["status"] != "Đã đóng": raise HTTPException(404, "Feedback link is invalid or ticket is not closed")
    en = req.session.get("lang")=="en"
    if row["feedback_rating"] is not None:
        body = f'<div class="box csat-result"><h2>{"Thank you for your feedback" if en else "Cảm ơn bạn đã đánh giá"}</h2><p>{E(row["code"] or "")}</p></div>'
    else:
        body = f'''<div class="box csat-public"><h2>{"Rate IT support" if en else "Đánh giá hỗ trợ IT"}</h2><p>{E(row["code"] or "")}: {E(row["issue"] or "")}</p><form method="post" action="/feedback/{E(token)}" id="publicCsat"><div class="csat-stars" role="radiogroup">{''.join(f'<button type="button" class="csat-star" data-rating="{i}" aria-label="{i} stars" onclick="setPublicCsat({i})">☆</button>' for i in range(1,6))}</div><input type="hidden" name="rating" id="publicRating"><div id="publicCommentWrap" hidden><label for="publicComment">{"Reason (required for ratings below 3)" if en else "Lý do (bắt buộc khi dưới 3 sao)"}</label><textarea name="comment" id="publicComment" rows="4" maxlength="1500"></textarea></div><button type="submit" id="publicSubmit" disabled>{"Submit rating" if en else "Gửi đánh giá"}</button></form></div><script>function setPublicCsat(n){{document.getElementById('publicRating').value=n;document.querySelectorAll('.csat-star').forEach((b,i)=>b.textContent=i<n?'★':'☆');const w=document.getElementById('publicCommentWrap'),c=document.getElementById('publicComment');w.hidden=n>=3;c.required=n<3;document.getElementById('publicSubmit').disabled=false}}</script>'''
    return page(req, "Feedback" if en else "Đánh giá dịch vụ", body)

@app.post("/feedback/{token}")
def submit_public_feedback(req: Request, token: str, rating: int = Form(...), comment: str = Form("")):
    row = query("SELECT t.id,t.status,t.reporter FROM ticket_feedback_tokens k JOIN tickets t ON t.id=k.ticket_id WHERE k.token=?", (token,), one=True)
    comment = comment.strip()[:1500]
    if not row or row["status"] != "Đã đóng": raise HTTPException(404, "Feedback link is invalid" if req.session.get("lang") == "en" else "Liên kết đánh giá không hợp lệ")
    if rating < 1 or rating > 5 or (rating < 3 and not comment): raise HTTPException(400, "Rating must be 1–5; ratings below 3 require a comment" if req.session.get("lang") == "en" else "Đánh giá phải từ 1–5 sao; đánh giá dưới 3 sao bắt buộc có nhận xét")
    query("INSERT OR IGNORE INTO ticket_feedback(ticket_id,rating,comment,reviewer_name,created_at) VALUES(?,?,?,?,?)", (row["id"],rating,comment,row["reporter"],now()), commit=True)
    return RedirectResponse(f"/feedback/{token}",303)


@app.post("/ticket/{tid}/public-note")
def save_public_ticket_note(req: Request, tid: int, public_note: str = Form(""), u=Depends(admin_user)):
    ticket = query("SELECT id FROM tickets WHERE id=?", (tid,), one=True)
    if not ticket:
        raise HTTPException(404, tr(req, "not_found"))
    query("UPDATE tickets SET public_note=? WHERE id=?", (public_note.strip()[:2000], tid), commit=True)
    return RedirectResponse(f"/ticket/{tid}", 303)

@app.post("/ticket/{tid}/feedback")
def submit_ticket_feedback(req: Request, tid: int, rating: int = Form(...), comment: str = Form(""), u=Depends(current_user)):
    row = query("SELECT id,user_id,status,reporter FROM tickets WHERE id=?", (tid,), one=True)
    comment = comment.strip()[:1500]
    if not row or row["user_id"] != u["id"] or row["status"] != "Đã đóng":
        raise HTTPException(403, tr(req, "not_found") if req.session.get("lang") != "en" else "You are not allowed to rate this ticket")
    existing_feedback = query("SELECT ticket_id FROM ticket_feedback WHERE ticket_id=?", (tid,), one=True)
    # First-time low ratings require a comment. Editing an existing rating allows
    # any score from 1–5; an empty comment intentionally clears the previous one.
    if rating < 1 or rating > 5 or (not existing_feedback and rating < 3 and not comment):
        message = ("Rating must be 1–5; first-time ratings below 3 require a comment" if req.session.get("lang") == "en" else "Đánh giá phải từ 1–5 sao; lần đánh giá đầu dưới 3 sao bắt buộc có nhận xét")
        raise HTTPException(400, message)
    query("INSERT INTO ticket_feedback(ticket_id,rating,comment,reviewer_user_id,reviewer_name,created_at) VALUES(?,?,?,?,?,?) ON CONFLICT(ticket_id) DO UPDATE SET rating=excluded.rating, comment=excluded.comment, reviewer_user_id=excluded.reviewer_user_id, reviewer_name=excluded.reviewer_name, created_at=excluded.created_at", (tid,rating,comment,u["id"],u["full_name"] or u["username"],now()), commit=True)
    return RedirectResponse(f"/ticket/{tid}",303)

@app.get("/api/ticket/{tid}/state")
def ticket_state(tid: int, u=Depends(current_user)):
    t = query("SELECT id,user_id,status,handler,received_at,closed_at FROM tickets WHERE id=?", (tid,), one=True)
    if not t or (u["role"] != "admin" and t["user_id"] != u["id"]):
        raise HTTPException(404, "Ticket not found")
    f = query("SELECT rating,comment,created_at FROM ticket_feedback WHERE ticket_id=?", (tid,), one=True)
    return {"status": t["status"], "handler": t["handler"], "received_at": t["received_at"], "closed_at": t["closed_at"], "feedback": dict(f) if f else None}

def ticket_sync_signature():
    # O(1) live-sync check: SQLite triggers increment this revision on ticket/feedback writes.
    row = query("SELECT revision FROM app_revision WHERE id=1", one=True)
    return str(row["revision"] if row else 0)

@app.get("/api/tickets/sync")
def tickets_sync(u=Depends(current_user)):
    return {"signature": ticket_sync_signature()}

@app.get("/api/feedback/summary")
def feedback_summary(u=Depends(admin_user)):
    # This endpoint is polled by the dashboard. Aggregate only; detailed feedback
    # remains available on ticket detail and is not repeatedly transferred to the browser.
    feedback = query("SELECT COUNT(*) AS total, SUM(CASE WHEN rating>=4 THEN 1 ELSE 0 END) AS good, SUM(CASE WHEN rating BETWEEN 1 AND 3 THEN 1 ELSE 0 END) AS poor FROM ticket_feedback", one=True)
    counts = query("SELECT status,COUNT(*) AS n FROM tickets GROUP BY status ORDER BY status")
    return {"feedback": {"total": int(feedback["total"] or 0), "good": int(feedback["good"] or 0), "poor": int(feedback["poor"] or 0)}, "tickets": [{"status":r["status"],"count":r["n"]} for r in counts], "ticket_signature": ticket_sync_signature()}

@app.get("/admin", response_class=HTMLResponse)
def admin_page(req: Request, status: str = "", device: str = "", department: str = "", day: str = "", severity: str = "", issue_type: str = "", high_only: int = 0, all_dates: int = 0, page_num: int = 1, feedback_group: str = "", u=Depends(admin_user)):
    if feedback_group not in ("", "good", "poor", "unrated"): feedback_group = ""
    if status not in STATUSES: status = ""
    if device not in DEVICES: device = ""
    department = normalize_department(department) if department else ""
    if department and department not in department_options(): department = ""
    import re
    today_iso = now_dt().strftime('%Y-%m-%d')
    # Opening the Admin ticket list shows all tickets by default. The Today shortcut
    # passes an explicit day, while All tickets keeps the complete date range.
    if not day and not all_dates:
        # Filters such as Department=IT must search the full history unless the user
        # explicitly chooses a creation date (the Today shortcut supplies `day`).
        all_dates = 1
    if day:
        try:
            if re.fullmatch(r"\d{2}/\d{2}/\d{4}", day): day = datetime.strptime(day, "%d/%m/%Y").strftime("%Y-%m-%d")
            elif re.fullmatch(r"\d{4}-\d{2}-\d{2}", day): day = datetime.strptime(day, "%Y-%m-%d").strftime("%Y-%m-%d")
            else: day = ""
        except ValueError: day = ""
    if not day and not all_dates:
        all_dates = 1
    if severity not in SEVERITIES: severity = "Cao" if high_only else ""
    if issue_type not in ISSUE_TYPES: issue_type = ""
    day_sql = "substr(created_at,7,4)||'-'||substr(created_at,4,2)||'-'||substr(created_at,1,2)"
    # Timestamp is DD/MM/YYYY HH:MM:SS. Admin groups by active workflow first, then severity;
    # within the same status and severity, oldest outstanding tickets are shown first.
    # Trim/case-fold aliases so legacy English/Vietnamese values sort consistently.
    created_day_sort = "substr(created_at,7,4)||substr(created_at,4,2)||substr(created_at,1,2)"
    created_sort = created_day_sort + "||substr(created_at,12,8)"
    severity_order_sql = "CASE WHEN lower(trim(COALESCE(severity,''))) IN ('cao','high') THEN 0 WHEN lower(trim(COALESCE(severity,''))) IN ('trung bình','medium') THEN 1 WHEN lower(trim(COALESCE(severity,''))) IN ('thấp','low') THEN 2 ELSE 3 END"
    status_order_sql = "CASE WHEN lower(trim(COALESCE(status,'Mới'))) IN ('mới','new','pending','chưa xử lý') THEN 0 WHEN lower(trim(COALESCE(status,''))) IN ('đang xử lý','in progress','in-progress') THEN 1 WHEN lower(trim(COALESCE(status,''))) IN ('đã đóng','closed','completed','đã xử lý') THEN 2 ELSE 3 END"
    order_clause = f"{status_order_sql} ASC, {severity_order_sql} ASC, {created_sort} DESC, id DESC"
    feedback_filter_sql = "(?='' OR (?='good' AND status='Đã đóng' AND EXISTS (SELECT 1 FROM ticket_feedback f WHERE f.ticket_id=tickets.id AND f.rating>=4)) OR (?='poor' AND status='Đã đóng' AND EXISTS (SELECT 1 FROM ticket_feedback f WHERE f.ticket_id=tickets.id AND f.rating BETWEEN 1 AND 3)) OR (?='unrated' AND status='Đã đóng' AND NOT EXISTS (SELECT 1 FROM ticket_feedback f WHERE f.ticket_id=tickets.id)))"
    admin_where = f"WHERE (?='' OR status=?) AND (?='' OR device=?) AND (?='' OR department=?) AND (?='' OR {day_sql}=?) AND (?='' OR severity=?) AND (?='' OR issue_type=?) AND {feedback_filter_sql}"
    admin_args = (status,status,device,device,department,department,day,day,severity,severity,issue_type,issue_type,feedback_group,feedback_group,feedback_group,feedback_group)
    total_row = query(f"SELECT COUNT(*) AS n FROM tickets {admin_where}", admin_args, one=True)
    total_rows = int(total_row["n"] if total_row else 0)
    all_ticket_row = query("SELECT COUNT(*) AS n FROM tickets", one=True)
    all_ticket_count = int(all_ticket_row["n"] if all_ticket_row else 0)
    page_num = max(1, min(int(page_num or 1), max(1, (total_rows + 9) // 10)))
    rows = query(f"SELECT * FROM tickets {admin_where} ORDER BY {order_clause} LIMIT 10 OFFSET ?",
                 admin_args + ((page_num - 1) * 10,))
    # Keep the database status values unchanged while presenting the business-friendly labels requested for Admin.
    admin_status_filter_labels = {
        "Mới": "Pending" if req.session.get("lang") == "en" else "Chưa xử lý",
        "Đang xử lý": "In progress" if req.session.get("lang") == "en" else "Đang xử lý",
        "Đã đóng": "Completed" if req.session.get("lang") == "en" else "Đã xử lý",
    }
    all_status_filter_label = "All statuses" if req.session.get("lang") == "en" else "Tất cả trạng thái"
    opts = '<option value="">'+all_status_filter_label+'</option>'+''.join(f'<option value="{E(s)}" {"selected" if s==status else ""}>{E(admin_status_filter_labels[s])}</option>' for s in STATUSES)
    device_filter_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(d)}" {"selected" if d==device else ""}>{E(device_label(req,d))}</option>' for d in DEVICES)
    dept_filter_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(d)}" {"selected" if department==d else ""}>{E(department_label(req,d))}</option>' for d in department_options())
    severity_order = ["Cao", "Trung bình", "Thấp"]
    severity_option_colors = {"Cao":"#c62e3b", "Trung bình":"#a15b00", "Thấp":"#087e68"}
    severity_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(x)}" style="color:{severity_option_colors[x]};font-weight:700" {"selected" if severity==x else ""}>{E(sev_label(req,x))}</option>' for x in severity_order)
    issue_filter_opts = '<option value="">'+tr(req,"all")+'</option>'+''.join(f'<option value="{E(x)}" {"selected" if issue_type==x else ""}>{E(issue_label(req,x))}</option>' for x in ISSUE_TYPES)
    day_display = datetime.strptime(day, "%Y-%m-%d").strftime("%d/%m/%Y") if day else ""
    english = req.session.get("lang") == "en"
    created_label = "Created date" if english else "Ngày tạo"
    all_label = "All" if english else "Tất cả"
    today_label = "Today’s tickets" if english else "Ticket hôm nay"
    all_quick = "All tickets" if english else "Tất cả ticket"
    refresh_label = "Refresh" if english else "Làm mới"
    date_field = f'<div><label>{created_label}</label><div class="calendar-pick-wrap"><input class="w date-text-input" type="text" name="day" id="day-text" value="{E(day_display)}" placeholder="DD/MM/YYYY" inputmode="numeric" pattern="[0-9]{{2}}/[0-9]{{2}}/[0-9]{{4}}" autocomplete="off"><input type="date" id="day-picker" aria-label="{created_label}" style="position:absolute;width:1px;height:1px;opacity:0;pointer-events:none"><button type="button" class="calendar-pick-button" id="open-day-picker" title="{created_label}" aria-label="{created_label}"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18"/><path d="M8 14h.01M12 14h.01M16 14h.01M8 17h.01M12 17h.01"/></svg></button></div></div>'
    body = fr'''<section class="box ticket-toolbar-panel"><div class="ticket-list-head"><h2><span class="section-icon"><svg viewBox="0 0 24 24" width="19" height="19" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8M8 12h8M8 16h5"/></svg></span>{tr(req,"ticket_list")}</h2><div class="toolbar-actions"><details class="export-language-menu"><summary class="btn btn-export"><svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M13 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V10z"/><path d="M13 3v7h7M8 13h8M8 16h8M8 19h5"/><path d="M8 10h2"/></svg> {tr(req,"export")} ▾</summary><div class="export-language-options"><a href="/export?lang=vi">Tiếng Việt</a><a href="/export?lang=en">English</a></div></details></div><style>.export-language-menu{{position:relative}}.export-language-menu summary{{list-style:none;cursor:pointer}}.export-language-menu summary::-webkit-details-marker{{display:none}}.export-language-options{{position:absolute;right:0;top:calc(100% + 6px);z-index:80;min-width:130px;padding:5px;background:#fff;border:1px solid #dbe7f3;border-radius:9px;box-shadow:0 8px 24px #17365d22}}.export-language-options a{{display:block;padding:9px 10px;border-radius:6px;color:#244c75;text-decoration:none;font-size:12px}}.export-language-options a:hover{{background:#edf6ff}}</style></div><div class="quick-toolbar admin-quick-toolbar"><a class="quick-filter" href="/admin?day={today_iso}"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18"/><path d="M8 14h.01M12 14h.01M16 14h.01M8 17h.01M12 17h.01"/></svg> {today_label}</a><a class="quick-filter" href="/admin?all_dates=1"><svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 6h13M8 12h13M8 18h13"/><path d="m3 6 1 1 2-2M3 12l1 1 2-2M3 18l1 1 2-2"/></svg> {all_quick}</a><a class="quick-filter refresh-quick" href="/admin?all_dates=1" title="{refresh_label}" aria-label="{refresh_label}"><svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.13-3.36L23 10M1 14l5.36 4.36A9 9 0 0 0 20.49 15"/></svg> {refresh_label}</a></div><form method="get" class="ticket-filter-form"><input type="hidden" name="all_dates" value="{1 if all_dates else 0}"><input type="hidden" name="feedback_group" value="{E(feedback_group)}"><div class="filter-grid compact-filter-grid"><div><label>{tr(req,"dept")}</label><select class="w" name="department">{dept_filter_opts}</select></div><div><label>{tr(req,"status")}</label><select class="w" name="status">{opts}</select></div><div><label>{tr(req,"device")}</label><select class="w" name="device">{device_filter_opts}</select></div><div><label>{tr(req,"issue_type")}</label><select class="w" name="issue_type">{issue_filter_opts}</select></div>{date_field}<div><label>{tr(req,"severity")}</label><select class="w" name="severity">{severity_opts}</select></div><div class="filter-actions"><button type="submit"><span aria-hidden="true">⌕</span> {tr(req,"filter")}</button><a class="btn btn-reset" href="/admin?all_dates=1">{all_label}</a></div></div></form></section><section class="box ticket-table-panel">{ticket_rows(req,rows,True)}{ticket_pagination(req,total_rows,page_num,all_total=all_ticket_count)}</section><div class="admin-home-bottom"><a href="/" class="admin-home-icon" title="{'Back to home' if english else 'Quay về trang chủ'}" aria-label="{'Back to home' if english else 'Quay về trang chủ'}"><svg viewBox="0 0 24 24" width="21" height="21" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m3 10 9-7 9 7"/><path d="M5 9v11h14V9M9 20v-6h6v6"/></svg></a></div><style>.admin-home-bottom{{display:flex;justify-content:center;align-items:center;margin:24px 0 8px;padding-top:14px;border-top:1px solid #e4edf7}}.admin-home-icon{{display:inline-flex;align-items:center;justify-content:center;width:42px;height:42px;border-radius:50%;background:#edf6ff;color:#176fc3;border:1px solid #d6e8fa;box-shadow:0 3px 10px rgba(30,90,150,.08);text-decoration:none;transition:background .18s,transform .18s}}.admin-home-icon:hover{{background:#dceeff;transform:translateY(-1px)}}.admin-home-icon:focus-visible{{outline:3px solid #93c5fd;outline-offset:3px}}</style><script>(function(){{const text=document.getElementById('day-text'), picker=document.getElementById('day-picker'), button=document.getElementById('open-day-picker');function isoToDmy(v){{if(!v)return '';const p=v.split('-');return p[2]+'/'+p[1]+'/'+p[0];}}function dmyToIso(v){{const m=/^(\d{{2}})\/(\d{{2}})\/(\d{{4}})$/.exec(v);return m?m[3]+'-'+m[2]+'-'+m[1]:'';}}if(text&&picker&&button){{picker.value=dmyToIso(text.value);button.addEventListener('click',function(){{if(picker.showPicker)picker.showPicker();else picker.click();}});picker.addEventListener('change',function(){{text.value=isoToDmy(this.value);}});text.addEventListener('change',function(){{const v=dmyToIso(this.value);if(v)picker.value=v;}});}}}})();</script>'''
    body += f'<script>(function(){{const initialRevision="{ticket_sync_signature()}";let signature="";async function check(){{try{{const r=await fetch("/api/feedback/summary",{{cache:"no-store"}});if(!r.ok)return;const j=await r.json();const next=JSON.stringify([j.feedback,j.tickets,j.ticket_signature]);if((!signature&&String(j.ticket_signature)!==initialRevision)||(signature&&next!==signature)){{location.reload();return}}signature=next}}catch(e){{}}}}check();setInterval(check,15000)}})();</script>'
    return page(req, tr(req, "admin"), body, u)
@app.get("/settings", response_class=HTMLResponse)
def integration_settings_page(req: Request, msg: str = "", tab: str = "notifications", u=Depends(admin_user)):
    tab = "hotline" if tab == "hotline" else "notifications"
    english = req.session.get("lang") == "en"
    text = {
        "title": "Automatic Notifications" if english else "Cài đặt thông báo tự động",
        "intro": "Configure Microsoft Teams, Telegram and Discord notifications. Save your settings, then test each connection." if english else "Cấu hình thông báo Microsoft Teams, Telegram và Discord. Lưu cấu hình, sau đó kiểm tra kết nối từng dịch vụ.",
        "teams": "Microsoft Teams", "telegram": "Telegram", "discord": "Discord",
        "enabled": "Enable notifications for this service" if english else "Bật gửi thông báo qua dịch vụ này",
        "webhook": "Teams Workflows webhook URL" if english else "Đường dẫn Webhook của Teams Workflows",
        "discord_webhook": "Discord Webhook URL",
        "token": "Telegram Bot Token", "chat_id": "Telegram Chat ID",
        "secret_saved": "A value is already saved. Leave blank to keep it; enter a new value to replace it." if english else "Đã có thông tin được lưu. Để trống để giữ nguyên; nhập giá trị mới nếu muốn thay thế.",
        "clear": "Delete saved connection information" if english else "Xóa thông tin kết nối đã lưu",
        "test": "Test connection" if english else "Kiểm tra kết nối",
        "discover": "Find Chat ID automatically" if english else "Tự động lấy Chat ID",
        "events": "When should notifications be sent?" if english else "Gửi thông báo khi nào?",
        "new_event": "A new ticket is created" if english else "Có ticket mới được tạo",
        "high_event": "A ticket is created or changed to High severity" if english else "Ticket mới hoặc được đổi sang mức độ Cao",
        "taken_event": "An admin accepts a ticket" if english else "Admin nhận xử lý ticket",
        "closed_event": "A ticket is closed" if english else "Ticket được đóng",
        "updated_event": "Ticket details are updated" if english else "Thông tin ticket được cập nhật",
        "save": "Save settings" if english else "Lưu cài đặt",
        "teams_help": "Recommended: in Teams, open the target channel → Workflows → choose a webhook alert template → save and copy its URL. Paste the URL here." if english else "Cách làm: trong Teams, mở kênh nhận tin → Workflows → chọn mẫu webhook thông báo → lưu và sao chép đường dẫn. Dán đường dẫn vào đây.",
        "telegram_help": "Create a bot with @BotFather, save its token below, then open the bot and press Start/send it a message. Save the settings and use “Find Chat ID automatically”. For a group, add the bot and send a message in that group first." if english else "Tạo bot bằng @BotFather, lưu Bot Token bên dưới, rồi mở bot và bấm Start/gửi một tin nhắn. Lưu cài đặt, sau đó bấm “Tự động lấy Chat ID”. Nếu dùng nhóm, thêm bot vào nhóm rồi gửi một tin nhắn trong nhóm trước.",
        "discord_help": "Create a free Discord server and a text channel, then open Server Settings → Integrations → Webhooks → New Webhook. Choose the channel and copy the webhook URL. Keep that URL private." if english else "Tạo server Discord miễn phí và một kênh văn bản, sau đó vào Cài đặt máy chủ → Tích hợp (Integrations) → Webhooks → Tạo Webhook mới. Chọn kênh và sao chép URL webhook. Giữ bí mật URL này.",
        "connected": "Connected" if english else "Đã kết nối",
        "not_connected": "Not connected" if english else "Chưa kết nối",
        "last_check": "Last successful connection check" if english else "Lần kiểm tra kết nối gần nhất",
        "not_checked": "Not tested yet" if english else "Chưa kiểm tra",
    }
    notices = {
        "saved": ("Settings saved successfully." if english else "Đã lưu cài đặt thành công.", False),
        "teams_test_ok": ("Teams test message sent successfully." if english else "Đã gửi tin nhắn kiểm tra đến Teams thành công.", False),
        "telegram_test_ok": ("Telegram test message sent successfully." if english else "Đã gửi tin nhắn kiểm tra đến Telegram thành công.", False),
        "discord_test_ok": ("Discord test message sent successfully." if english else "Đã gửi tin nhắn kiểm tra đến Discord thành công.", False),
        "teams_test_failed": ("Teams connection test failed. Review the saved URL and error details below." if english else "Kiểm tra kết nối Teams thất bại. Kiểm tra lại URL và thông tin lỗi bên dưới.", True),
        "telegram_test_failed": ("Telegram connection test failed. Review the saved token/Chat ID and error details below." if english else "Kiểm tra kết nối Telegram thất bại. Kiểm tra lại Bot Token/Chat ID và thông tin lỗi bên dưới.", True),
        "discord_test_failed": ("Discord connection test failed. Check the saved webhook URL and error details below." if english else "Kiểm tra kết nối Discord thất bại. Kiểm tra URL webhook đã lưu và lỗi bên dưới.", True),
        "telegram_chat_found": ("Chat ID found and saved. Test the connection to confirm messaging works." if english else "Đã tìm và lưu Chat ID. Hãy kiểm tra kết nối để xác nhận gửi tin nhắn hoạt động.", False),
        "telegram_chat_failed": ("Could not find a Telegram Chat ID. Follow the instructions in the Telegram card." if english else "Chưa lấy được Chat ID Telegram. Làm theo hướng dẫn trong khung Telegram.", True),
        "invalid_teams_url": ("Teams webhook URL must start with https://." if english else "URL webhook Teams phải bắt đầu bằng https://." , True),
        "invalid_discord_url": ("Enter a valid Discord webhook URL starting with https://discord.com/api/webhooks/." if english else "Nhập URL Discord Webhook hợp lệ, bắt đầu bằng https://discord.com/api/webhooks/.", True),
    }
    settings_label = "Settings" if english else "Cài đặt"
    notifications_tab_label = "Automatic notifications" if english else "Thông báo tự động"
    hotline_tab_label = "IT Support Hotline" if english else "Hotline hỗ trợ"
    settings_tabs_html = (
        '<nav class="settings-tabs" aria-label="Settings sections">'
        f'<a class="settings-tab {"active" if tab == "notifications" else ""}" href="/settings?tab=notifications">'
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M18 8a6 6 0 0 0-12 0c0 7-3 7-3 9h18c0-2-3-2-3-9"/><path d="M10 21h4"/></svg>'
        f'{notifications_tab_label}</a><a class="settings-tab {"active" if tab == "hotline" else ""}" href="/settings?tab=hotline">'
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.4 19.4 0 0 1-6-6A19.8 19.8 0 0 1 2.1 4.2 2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7l.4 2.8a2 2 0 0 1-.6 1.7L7.6 9.5a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 1.7-.6l2.8.4a2 2 0 0 1 1.7 1.9Z"/></svg>'
        f'{hotline_tab_label}</a></nav>'
    )
    if tab == "hotline":
        hotline_phone = get_integration_setting("hotline_phone", "").strip()
        hotline_email = get_integration_setting("hotline_email", "").strip()
        hotline_saved = "Đã lưu thông tin Hotline thành công." if not english else "Hotline details saved successfully."
        hotline_invalid = "Số điện thoại chỉ nên gồm chữ số, dấu +, dấu cách, ngoặc hoặc dấu gạch ngang." if not english else "Use digits, +, spaces, parentheses or hyphens for the phone number."
        hotline_notice_map = {
            "hotline_saved": (hotline_saved, False),
            "hotline_invalid": (hotline_invalid, True),
        }
        hotline_notice_html = ""
        if msg in hotline_notice_map:
            note, is_error = hotline_notice_map[msg]
            hotline_notice_html = f'<div class="settings-message{" error" if is_error else ""}" role="status">{E(note)}</div>'
        hotline_preview_number = f'<span class="hotline-number">{E(hotline_phone)}</span>' if hotline_phone else '<span class="hotline-number hotline-not-configured">' + ("Chưa cập nhật số" if not english else "Not configured") + '</span>'
        hotline_preview_email = f'<span class="hotline-email">{E(hotline_email)}</span>' if hotline_email else ''
        hotline_title = "IT Support Hotline" if english else "Hỗ trợ IT"
        hotline_intro = "Set the phone number and email displayed in the User sidebar. Contact details are shown as text, not call/email links." if english else "Cài đặt số điện thoại và email sẽ hiển thị trong menu bên trái trang User. Thông tin chỉ hiển thị dạng chữ, không mở ứng dụng gọi điện hoặc email."
        hotline_label = "Hotline phone number" if english else "Số điện thoại Hotline"
        hotline_email_label = "IT support email" if english else "Email hỗ trợ IT"
        hotline_email_help = "Example: it@example.com. Leave empty to hide the email on the User page." if english else "Ví dụ: it@example.com. Để trống nếu không muốn hiển thị email ở trang User."
        hotline_help = "Examples: +84 28 1234 5678, 028-1234-5678. Leave empty and save to remove the number." if english else "Ví dụ: +84 28 1234 5678 hoặc 028-1234-5678. Để trống rồi lưu nếu muốn xóa số Hotline."
        save_hotline_label = "Save contact details" if english else "Lưu thông tin hỗ trợ"
        body = f'''<section class="box">
          <div class="settings-hero"><span class="settings-hero-icon"><svg viewBox="0 0 24 24" width="24" height="24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.4 19.4 0 0 1-6-6A19.8 19.8 0 0 1 2.1 4.2 2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7l.4 2.8a2 2 0 0 1-.6 1.7L7.6 9.5a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 1.7-.6l2.8.4a2 2 0 0 1 1.7 1.9Z"/></svg></span><div><h2>{hotline_title}</h2><p>{hotline_intro}</p></div></div>
          {settings_tabs_html}{hotline_notice_html}
          <section class="integration-card hotline-config-card">
            <form method="post" action="/settings/hotline" autocomplete="off">
              <label for="hotline_phone">{hotline_label}</label>
              <div class="hotline-input-row"><input id="hotline_phone" type="tel" name="hotline_phone" value="{E(hotline_phone)}" inputmode="tel" maxlength="32" placeholder="+84 28 1234 5678" autocomplete="tel"></div>
              <p class="integration-help">{hotline_help}</p>
              <label for="hotline_email">{hotline_email_label}</label>
              <div class="hotline-input-row"><input id="hotline_email" type="email" name="hotline_email" value="{E(hotline_email)}" maxlength="254" placeholder="it@example.com" autocomplete="email"></div>
              <p class="integration-help">{hotline_email_help}</p>
              <div class="hotline-preview"><span class="hotline-it-icon"><svg viewBox="0 0 48 48" width="24" height="24" aria-hidden="true"><circle cx="24" cy="24" r="21" fill="none" stroke="currentColor" stroke-width="1.6" opacity=".7"/><path d="M12 25v-3a12 12 0 0 1 24 0v3M12 24h5v10h-3a3 3 0 0 1-3-3v-4a3 3 0 0 1 1-3ZM36 24h-5v10h3a3 3 0 0 0 3-3v-4a3 3 0 0 0-1-3ZM31 35a8 8 0 0 1-7 3h-2" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><text x="24" y="17" text-anchor="middle" font-size="7" font-weight="900" fill="currentColor">24/7</text></svg></span><div class="hotline-preview-copy"><strong>{hotline_title}</strong>{hotline_preview_number}{hotline_preview_email}</div></div>
              <div class="hotline-save-actions"><button type="submit">✓ {save_hotline_label}</button></div>
            </form>
          </section>
        </section>'''
        return page(req, settings_label, body, u)

    notice_html = ""

    if msg in notices:
        notice_text, is_error = notices[msg]
        notice_html = f'<div class="settings-message{" error" if is_error else ""}" role="status">{E(notice_text)}</div>'

    teams_ready = integration_credentials_ready("teams")
    telegram_ready = integration_credentials_ready("telegram")
    discord_ready = integration_credentials_ready("discord")
    teams_connected = integration_is_connected("teams")
    telegram_connected = integration_is_connected("telegram")
    discord_connected = integration_is_connected("discord")
    teams_badge = f'<span class="connection-badge {"connected" if teams_connected else "disconnected"}">● {text["connected"] if teams_connected else text["not_connected"]}</span>'
    telegram_badge = f'<span class="connection-badge {"connected" if telegram_connected else "disconnected"}">● {text["connected"] if telegram_connected else text["not_connected"]}</span>'
    discord_badge = f'<span class="connection-badge {"connected" if discord_connected else "disconnected"}">● {text["connected"] if discord_connected else text["not_connected"]}</span>'
    def check(name):
        return 'checked' if get_integration_setting(name, "1") == "1" else ''
    def last_check(service):
        checked_at = get_integration_setting(f"{service}_last_checked", "")
        error = get_integration_setting(f"{service}_last_error", "")
        parts = [f'{text["last_check"]}: {E(checked_at or text["not_checked"])}']
        if error:
            parts.append(("Last error: " if english else "Lỗi gần nhất: ") + E(error))
        return '<div class="integration-last-check">' + '<br>'.join(parts) + '</div>'

    body = f'''<section class="box">
      <div class="settings-hero"><span class="settings-hero-icon"><svg viewBox="0 0 24 24" width="24" height="24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3v3M5.6 5.6l2.1 2.1M3 12h3M18 12h3M16.3 7.7l2.1-2.1"/><path d="M6.5 12a5.5 5.5 0 0 1 11 0c0 2.5 1 3.3 1.8 4.3H4.7C5.5 15.3 6.5 14.5 6.5 12Z"/><path d="M9.5 19h5"/></svg></span><div><h2>{text["title"]}</h2><p>{text["intro"]}</p></div></div>
      {notice_html}
      {settings_tabs_html}
      <form method="post" action="/settings" autocomplete="off">
        <div class="integration-grid">
          <details id="teams-card" class="integration-card integration-accordion"><summary class="integration-card-head"><span class="integration-card-title"><span class="integration-card-mark">✣</span>{text["teams"]}</span>{teams_badge}</summary><div class="integration-card-content">
            <label class="integration-check"><input type="checkbox" name="teams_enabled" value="1" {check("teams_enabled")}> <span>{text["enabled"]}</span></label>
            <label for="teams_webhook_url">{text["webhook"]}</label><input id="teams_webhook_url" type="url" name="teams_webhook_url" class="w" placeholder="https://... (Teams Workflows)" autocomplete="new-password">
            <p class="integration-help">{text["secret_saved"] if teams_ready else ("Paste the webhook URL copied from Teams Workflows here." if english else "Dán URL webhook vừa sao chép từ Teams Workflows vào đây.")}</p>
            <label class="integration-check integration-clear"><input type="checkbox" name="clear_teams_config" value="1"> <span>{text["clear"]}</span></label>
            <p class="integration-help">{text["teams_help"]}</p>{last_check("teams")}
          </div></details>
          <details id="telegram-card" class="integration-card integration-accordion"><summary class="integration-card-head"><span class="integration-card-title"><span class="integration-card-mark">➤</span>{text["telegram"]}</span>{telegram_badge}</summary><div class="integration-card-content">
            <label class="integration-check"><input type="checkbox" name="telegram_enabled" value="1" {check("telegram_enabled")}> <span>{text["enabled"]}</span></label>
            <label for="telegram_bot_token">{text["token"]}</label><input id="telegram_bot_token" type="password" name="telegram_bot_token" class="w" placeholder="{E(text["secret_saved"] if get_integration_setting("telegram_bot_token", "") else ("Enter bot token" if english else "Nhập Bot Token"))}" autocomplete="new-password">
            <label for="telegram_chat_id">{text["chat_id"]}</label><input id="telegram_chat_id" type="text" name="telegram_chat_id" class="w" placeholder="{E(text["secret_saved"] if get_integration_setting("telegram_chat_id", "") else ("e.g. 123456789 or -100123..."))}" autocomplete="off">
            <label class="integration-check integration-clear"><input type="checkbox" name="clear_telegram_config" value="1"> <span>{text["clear"]}</span></label>
            <p class="integration-help">{text["telegram_help"]}</p>{last_check("telegram")}
          </div></details>
          <details id="discord-card" class="integration-card integration-accordion"><summary class="integration-card-head"><span class="integration-card-title"><span class="integration-card-mark">◉</span>{text["discord"]}</span>{discord_badge}</summary><div class="integration-card-content">
            <label class="integration-check"><input type="checkbox" name="discord_enabled" value="1" {check("discord_enabled")}> <span>{text["enabled"]}</span></label>
            <label for="discord_webhook_url">{text["discord_webhook"]}</label><input id="discord_webhook_url" type="password" name="discord_webhook_url" class="w" placeholder="{E(text["secret_saved"] if discord_ready else "https://discord.com/api/webhooks/...")}" autocomplete="new-password">
            <p class="integration-help">{text["secret_saved"] if discord_ready else text["discord_help"]}</p>
            <label class="integration-check integration-clear"><input type="checkbox" name="clear_discord_config" value="1"> <span>{text["clear"]}</span></label>
            {last_check("discord")}
          </div></details>
        </div>
        <details class="integration-card integration-accordion" open><summary class="integration-card-head"><span class="integration-card-title"><span class="integration-card-mark">☷</span>{text["events"]}</span><span class="accordion-chevron" aria-hidden="true">⌄</span></summary><div class="integration-card-content"><div class="notification-events">
          <label class="integration-check"><input type="checkbox" name="notify_new_ticket" value="1" {check("notify_new_ticket")}> <span>{text["new_event"]}</span></label>
          <label class="integration-check"><input type="checkbox" name="notify_high_ticket" value="1" {check("notify_high_ticket")}> <span>{text["high_event"]}</span></label>
          <label class="integration-check"><input type="checkbox" name="notify_ticket_taken" value="1" {check("notify_ticket_taken")}> <span>{text["taken_event"]}</span></label>
          <label class="integration-check"><input type="checkbox" name="notify_ticket_closed" value="1" {check("notify_ticket_closed")}> <span>{text["closed_event"]}</span></label>
          <label class="integration-check"><input type="checkbox" name="notify_ticket_updated" value="1" {check("notify_ticket_updated")}> <span>{text["updated_event"]}</span></label>
        </div></div></details>
        <div class="settings-save-row"><button type="submit"><svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h12l4 4v12a2 2 0 0 1-2 2Z"/><path d="M17 21v-8H7v8M7 3v5h9"/></svg>{text["save"]}</button></div>
      </form>
      <div class="integration-grid">
        <section class="integration-card"><div class="integration-card-title">{text["teams"]} — {text["test"]}</div><p class="integration-help">{"Save settings first. A real test message will appear in the configured Teams channel/chat." if english else "Lưu cài đặt trước. Tin nhắn kiểm tra thực tế sẽ xuất hiện tại kênh/chat Teams đã cấu hình."}</p><div class="integration-small-actions"><form method="post" action="/settings/test/teams"><button type="submit">✉ {text["test"]}</button></form></div></section>
        <section class="integration-card"><div class="integration-card-title">{text["telegram"]} — {text["test"]}</div><p class="integration-help">{"After saving the bot token and sending it a message in Telegram, you can detect the Chat ID automatically, then test it." if english else "Sau khi lưu Bot Token và gửi bot một tin nhắn trong Telegram, mi có thể tự lấy Chat ID rồi kiểm tra kết nối."}</p><div class="integration-small-actions"><form method="post" action="/settings/telegram/discover-chat"><button class="btn-secondary" type="submit">⌕ {text["discover"]}</button></form><form method="post" action="/settings/test/telegram"><button type="submit">✉ {text["test"]}</button></form></div></section>
        <section class="integration-card"><div class="integration-card-title">{text["discord"]} — {text["test"]}</div><p class="integration-help">{"Save the Discord webhook URL and enable Discord notifications before testing." if english else "Lưu URL Discord Webhook và bật thông báo Discord trước khi kiểm tra."}</p><div class="integration-small-actions"><form method="post" action="/settings/test/discord"><button type="submit">✉ {text["test"]}</button></form></div></section>
      </div>
      <style>
      .integration-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,270px),1fr));gap:12px;align-items:start;margin:12px 0}}
      .integration-accordion{{padding:0!important;overflow:hidden;border:1px solid #dfe8f2!important;border-radius:12px!important;background:var(--panel,#fff);box-shadow:0 2px 8px rgba(24,58,92,.045)}}
      .integration-accordion>summary{{list-style:none;cursor:pointer;display:flex;align-items:center;justify-content:space-between;gap:12px;padding:13px 15px;margin:0;min-height:48px;user-select:none}}
      .integration-accordion>summary::-webkit-details-marker{{display:none}}
      .integration-accordion>summary::after{{content:'⌄';font-size:18px;color:#71859b;transition:transform .18s ease;margin-left:8px;flex-shrink:0}}
      .integration-accordion[open]>summary::after{{transform:rotate(180deg)}}
      .integration-accordion .integration-card-head{{border-bottom:1px solid transparent}}
      .integration-accordion[open]>.integration-card-head{{border-bottom-color:#e8eef5;background:#f8fbff}}
      .integration-card-content{{padding:14px 15px 15px;display:grid;gap:8px}}
      .integration-card-content label:not(.integration-check){{font-size:12px;font-weight:650;color:#52677d;margin-top:3px}}
      .integration-card-content input.w{{min-height:36px;padding:8px 10px;border-radius:8px;font-size:13px;box-sizing:border-box;width:100%}}
      .integration-card-content .integration-help{{font-size:12px;line-height:1.5;margin:2px 0;color:#718096}}
      .integration-card-content .integration-check{{font-size:13px;line-height:1.4;margin:2px 0;display:flex;align-items:flex-start;gap:8px}}
      .integration-card-mark{{display:inline-flex;align-items:center;justify-content:center;width:30px;height:30px;min-width:30px;border-radius:9px;background:#edf5ff;color:#2368a9;font-size:17px}}
      .integration-card-title{{display:inline-flex;align-items:center;gap:10px;font-weight:700;color:#24496f}}
      .accordion-chevron{{display:none}}
      .notification-events{{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,235px),1fr));gap:8px 16px}}
      .settings-save-row{{display:flex;justify-content:flex-start;align-items:center;margin:12px 0 16px}}
      .settings-save-row button{{display:inline-flex;align-items:center;justify-content:center;gap:8px;min-height:38px;padding:8px 16px;border-radius:9px;font-weight:650;box-shadow:0 2px 5px rgba(20,75,120,.10)}}
      .integration-small-actions{{display:flex;flex-wrap:wrap;gap:8px}}
      .integration-small-actions button{{min-height:34px;padding:7px 11px;font-size:12px;border-radius:8px}}
      
      </style>
    </section>'''
    return page(req, text["title"], body, u)


@app.post("/settings/hotline")
def save_hotline_setting(req: Request, hotline_phone: str = Form(""), hotline_email: str = Form(""), u=Depends(admin_user)):
    phone = hotline_phone.strip()
    email = hotline_email.strip()
    if phone and (not re.fullmatch(r"[0-9+().\-\s]{4,32}", phone) or not any(ch.isdigit() for ch in phone)):
        return RedirectResponse("/settings?tab=hotline&msg=hotline_invalid", 303)
    if email and (len(email) > 254 or not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email)):
        return RedirectResponse("/settings?tab=hotline&msg=hotline_invalid", 303)
    set_integration_setting("hotline_phone", phone)
    set_integration_setting("hotline_email", email)
    return RedirectResponse("/settings?tab=hotline&msg=hotline_saved", 303)


@app.post("/settings")
def save_integration_settings(req: Request,
    teams_enabled: str = Form(""), teams_webhook_url: str = Form(""), clear_teams_config: str = Form(""),
    telegram_enabled: str = Form(""), telegram_bot_token: str = Form(""), telegram_chat_id: str = Form(""), clear_telegram_config: str = Form(""),
    discord_enabled: str = Form(""), discord_webhook_url: str = Form(""), clear_discord_config: str = Form(""),
    notify_new_ticket: str = Form(""), notify_high_ticket: str = Form(""), notify_ticket_taken: str = Form(""), notify_ticket_closed: str = Form(""), notify_ticket_updated: str = Form(""),
    u=Depends(admin_user)):
    new_teams_url = teams_webhook_url.strip()
    new_discord_url = discord_webhook_url.strip()
    if new_teams_url and not new_teams_url.lower().startswith("https://"):
        return RedirectResponse("/settings?tab=notifications&msg=invalid_teams_url", 303)
    if new_discord_url and not (new_discord_url.startswith("https://discord.com/api/webhooks/") or new_discord_url.startswith("https://discordapp.com/api/webhooks/")):
        return RedirectResponse("/settings?tab=notifications&msg=invalid_discord_url", 303)

    existing_teams_url = get_integration_setting("teams_webhook_url", "")
    existing_telegram_token = get_integration_setting("telegram_bot_token", "")
    existing_telegram_chat_id = get_integration_setting("telegram_chat_id", "")
    existing_discord_url = get_integration_setting("discord_webhook_url", "")

    if clear_teams_config == "1":
        new_teams_url = ""
    elif not new_teams_url:
        new_teams_url = existing_teams_url
    new_telegram_token = "" if clear_telegram_config == "1" else (telegram_bot_token.strip() or existing_telegram_token)
    new_telegram_chat_id = "" if clear_telegram_config == "1" else (telegram_chat_id.strip() or existing_telegram_chat_id)
    if clear_discord_config == "1":
        new_discord_url = ""
    elif not new_discord_url:
        new_discord_url = existing_discord_url

    set_integration_setting("teams_enabled", "1" if teams_enabled == "1" else "0")
    set_integration_setting("teams_webhook_url", new_teams_url)
    set_integration_setting("telegram_enabled", "1" if telegram_enabled == "1" else "0")
    set_integration_setting("telegram_bot_token", new_telegram_token)
    set_integration_setting("telegram_chat_id", new_telegram_chat_id)
    set_integration_setting("discord_enabled", "1" if discord_enabled == "1" else "0")
    set_integration_setting("discord_webhook_url", new_discord_url)
    for key, value in (("notify_new_ticket", notify_new_ticket), ("notify_high_ticket", notify_high_ticket),
                       ("notify_ticket_taken", notify_ticket_taken), ("notify_ticket_closed", notify_ticket_closed),
                       ("notify_ticket_updated", notify_ticket_updated)):
        set_integration_setting(key, "1" if value == "1" else "0")

    if clear_teams_config == "1" or new_teams_url != existing_teams_url:
        set_integration_status("teams", False, "Not tested yet")
    if clear_telegram_config == "1" or new_telegram_token != existing_telegram_token or new_telegram_chat_id != existing_telegram_chat_id:
        set_integration_status("telegram", False, "Not tested yet")
    if clear_discord_config == "1" or new_discord_url != existing_discord_url:
        set_integration_status("discord", False, "Not tested yet")
    return RedirectResponse("/settings?tab=notifications&msg=saved", 303)


@app.post("/settings/test/{service}")
def test_integration(req: Request, service: str, u=Depends(admin_user)):
    if service not in ("teams", "telegram", "discord"):
        raise HTTPException(404, "Integration not found")
    english = req.session.get("lang") == "en"
    message = "✅ FMP IT Service Desk — Test connection successful. Automatic notifications are ready." if english else "✅ FMP IT Service Desk — Kiểm tra kết nối thành công. Hệ thống sẵn sàng gửi thông báo tự động."
    try:
        send_integration_message(service, message)
        return RedirectResponse(f"/settings?tab=notifications&msg={service}_test_ok", 303)
    except Exception:
        return RedirectResponse(f"/settings?tab=notifications&msg={service}_test_failed", 303)


@app.post("/settings/telegram/discover-chat")
def discover_telegram_chat(req: Request, u=Depends(admin_user)):
    try:
        chat_id, _chat_name = get_telegram_chat_id()
        set_integration_setting("telegram_chat_id", chat_id)
        set_integration_status("telegram", False, "Chat ID saved; please test the connection")
        return RedirectResponse("/settings?tab=notifications&msg=telegram_chat_found", 303)
    except Exception:
        return RedirectResponse("/settings?tab=notifications&msg=telegram_chat_failed", 303)


@app.get("/export")
def export(req: Request, u=Depends(admin_user)):
    rows = query("SELECT * FROM tickets ORDER BY id")
    requested_lang = (req.query_params.get("lang") or "").lower()
    english = requested_lang == "en" if requested_lang in ("en", "vi") else req.session.get("lang") == "en"
    wb = Workbook()
    ws = wb.active
    ws.title = "Ticket details" if english else "Chi tiết ticket"
    report_title = "FMP IT SERVICE DESK — TICKET REPORT" if english else "BÁO CÁO TICKET — PHẦN MỀM IT SERVICE DESK"
    ws.append([report_title])
    ws.merge_cells("A1:P1")
    title_cell = ws["A1"]
    title_cell.font = Font(name="Aptos Display", size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill("solid", fgColor="17365D")
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 34
    ws.append([])
    headers_vi = ["Mã ticket","Thời gian tạo","Thiết bị","Mã thiết bị","Bộ phận","Vị trí","Loại sự cố","Mô tả sự cố","Mức độ","Người báo","Điện thoại","Trạng thái","Người xử lý","Thời gian nhận xử lý","Thời gian đóng"]
    headers_en = ["Ticket ID","Created At","Device","Device Code","Department","Location","Issue Type","Issue Details","Severity","Reporter","Phone","Status","Assigned To","Received At","Closed At"]
    ws.append(headers_en if english else headers_vi)
    header_fill = PatternFill("solid", fgColor="245B83")
    thin_blue = Side(style="thin", color="D9E7F5")
    for cell in ws[3]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="medium", color="17365D"))
    ws.row_dimensions[3].height = 32
    for r in rows:
        ws.append([r["code"],r["created_at"],device_label(req,r["device"] or ""),r["device_code"],r["department"],r["location"],issue_label(req,r["issue_type"] or ""),r["issue"],sev_label(req,r["severity"] or ""),r["reporter"],r["phone"],status_label(req,r["status"] or ""),r["handler"],r["received_at"],r["closed_at"]])
        for cell in ws[ws.max_row]:
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            cell.border = Border(bottom=thin_blue)
            if ws.max_row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F4F8FC")
    # Freeze the title and column-header rows so headers remain visible while scrolling.
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:O{ws.max_row}"

    summary = wb.create_sheet("KPI summary" if english else "Tổng hợp KPI", 0)
    summary.append(["Technician" if english else "Kỹ thuật viên", "Closed tickets" if english else "Ticket đã đóng", "Rated tickets" if english else "Ticket đã đánh giá", "Average CSAT" if english else "CSAT trung bình", "Satisfaction rate (4–5)" if english else "Tỷ lệ hài lòng (4–5 sao)", "Poor ratings (1–3)" if english else "Đánh giá tệ (1–3 sao)"])
    kpi = query("SELECT COALESCE(NULLIF(t.handler,''), ?) AS technician, COUNT(DISTINCT t.id) AS closed_count, COUNT(f.id) AS rated_count, AVG(f.rating) AS avg_rating, SUM(CASE WHEN f.rating>=4 THEN 1 ELSE 0 END) AS satisfied, SUM(CASE WHEN f.rating BETWEEN 1 AND 3 THEN 1 ELSE 0 END) AS low_count FROM tickets t LEFT JOIN ticket_feedback f ON f.ticket_id=t.id WHERE t.status='Đã đóng' GROUP BY COALESCE(NULLIF(t.handler,''), ?) ORDER BY technician", ("Unassigned" if english else "Chưa phân công", "Unassigned" if english else "Chưa phân công"))
    for r in kpi:
        summary.append([r["technician"],r["closed_count"],r["rated_count"],round(r["avg_rating"],2) if r["avg_rating"] is not None else "—", f'{(r["satisfied"] or 0)/r["rated_count"]:.1%}' if r["rated_count"] else "—",r["low_count"] or 0])

    feedback_ws = wb.create_sheet("Feedback to review" if english else "Phản hồi cần xem")
    feedback_ws.append(["Ticket ID" if english else "Mã ticket", "Technician" if english else "Kỹ thuật viên", "Reviewer" if english else "Người đánh giá", "Rating" if english else "Số sao", "Comment" if english else "Nhận xét", "Submitted at" if english else "Thời gian đánh giá"])
    review_rows = query("SELECT t.code,t.handler,f.reviewer_name,f.rating,f.comment,f.created_at FROM ticket_feedback f JOIN tickets t ON t.id=f.ticket_id WHERE f.rating BETWEEN 1 AND 3 OR TRIM(COALESCE(f.comment,''))<>'' ORDER BY f.created_at DESC")
    for r in review_rows:
        feedback_ws.append([r["code"],r["handler"],r["reviewer_name"],r["rating"],r["comment"],r["created_at"]])

    # Consistent formatting, sensible column widths, and wrapped comments on every sheet.
    for sheet in (summary, feedback_ws):
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = Border(bottom=Side(style="medium", color="17365D"))
        sheet.row_dimensions[1].height = 34
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.border = Border(bottom=thin_blue)
                if cell.row % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="F4F8FC")

    # Auto-fit using content length with per-column limits; wrapped long text gets room without
    # creating excessively wide spreadsheets. Explicitly size the free-text columns.
    from openpyxl.utils import get_column_letter
    detail_min = [14,19,16,15,16,16,22,30,13,18,15,16,18,28,19,19]
    detail_max = [22,22,24,20,24,24,34,48,16,26,18,20,26,48,22,22]
    detail_widths = [18,20,18,16,18,18,24,38,14,20,16,17,20,34,20,20]
    for idx, fixed_width in enumerate(detail_widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = fixed_width
    # Keep long ticket descriptions readable by wrapping and adjusting row heights.
    for row_idx in range(4, ws.max_row + 1):
        max_lines = 1
        for idx in (7, 8):
            value = str(ws.cell(row_idx, idx).value or "")
            width = max(1, int(ws.column_dimensions[get_column_letter(idx)].width) - 2)
            lines = sum(max(1, (len(part) + width - 1) // width) for part in value.splitlines()) if value else 1
            max_lines = max(max_lines, lines)
        ws.row_dimensions[row_idx].height = min(max(18, max_lines * 15), 90)

    widths = {
        summary: [24,18,18,18,27,19],
        feedback_ws: [18,24,24,12,60,22],
    }
    for sheet, suggested in widths.items():
        for idx, fixed_width in enumerate(suggested, 1):
            sheet.column_dimensions[get_column_letter(idx)].width = fixed_width
        for row_idx in range(2, sheet.max_row + 1):
            comment_idx = 5 if sheet is feedback_ws else None
            if comment_idx:
                value = str(sheet.cell(row_idx, comment_idx).value or "")
                width = max(1, int(sheet.column_dimensions[get_column_letter(comment_idx)].width) - 2)
                lines = sum(max(1, (len(part) + width - 1) // width) for part in value.splitlines()) if value else 1
                sheet.row_dimensions[row_idx].height = min(max(18, lines * 15), 120)
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal="left", vertical="top" if cell.row > 1 else "center", wrap_text=True)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition":'attachment; filename="report_ticket_fmp_it_service_desk.xlsx"'})

@app.get("/users", response_class=HTMLResponse)
def users_page(req: Request, msg: str = "", u=Depends(admin_user)):
    rows = query("SELECT * FROM users ORDER BY role,username")
    en = req.session.get("lang") == "en"
    success_messages = {
        "User updated": ("User updated successfully" if en else "Đã cập nhật tài khoản thành công"),
        "User added": ("User added successfully" if en else "Đã tạo tài khoản thành công"),
        "Password reset": ("Password reset successfully" if en else "Đã đặt lại mật khẩu thành công"),
        "User locked": ("Account locked successfully" if en else "Đã khóa tài khoản thành công"),
        "User unlocked": ("Account unlocked successfully" if en else "Đã mở khóa tài khoản thành công"),
        "Passwordless enabled": ("Passwordless login enabled" if en else "Đã bật đăng nhập không cần mật khẩu"),
        "Passwordless disabled": ("Passwordless login disabled" if en else "Đã tắt đăng nhập không cần mật khẩu"),
    }
    toast_message = success_messages.get(msg, "")
    toast_html = (f'<div class="account-toast" role="status" aria-live="polite"><span class="account-toast-icon"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="m5 12 4 4L19 6"/></svg></span><span>{E(toast_message)}</span><button type="button" class="account-toast-close" aria-label="{"Close notification" if en else "Đóng thông báo"}">×</button></div>' if toast_message else "")
    body = f'{toast_html}<div class="box account-create-box"><h2>{tr(req,"add")}</h2><p class="help">{tr(req,"hint6")}</p><form method="post" action="/users"><div class="user-edit-grid"><div><label>{tr(req,"username")}</label><input class="w" name="username" required maxlength="80" autocomplete="off"></div><div><label>{tr(req,"fullname")}</label><input class="w" name="full_name" required maxlength="150"></div><div><label>{tr(req,"dept")}</label><select class="w" name="department" required><option value="">{"Select department" if req.session.get("lang")=="en" else "Chọn bộ phận"}</option>{''.join(f'<option value="{E(d)}">{E(department_label(req,d))}</option>' for d in department_options())}</select></div><div><label>{tr(req,"phone")}</label><input class="w" name="phone" maxlength="50"></div><div><label>{tr(req,"initial")}</label><input class="w" type="password" name="password" minlength="6" title="{tr(req,"hint6")}" required autocomplete="new-password"></div><div><label>{tr(req,"role")}</label><select class="w" name="role"><option value="user">{tr(req,"role_user")}</option><option value="admin">{tr(req,"role_admin")}</option></select></div></div><p><button type="submit">{tr(req,"add")}</button></p></form></div><div class="box account-list-box"><div class="account-list-heading"><h2>{tr(req,"users")}</h2><span class="account-count">{len(rows)}</span></div><div class="account-list">'
    for r in rows:
        active_label = tr(req, "active") if r["active"] else tr(req, "locked")
        role_label = tr(req, "role_admin") if r["role"] == "admin" else tr(req, "role_user")
        lock_title = ("Unlock account" if not r["active"] else "Lock account") if en else ("Mở khóa tài khoản" if not r["active"] else "Khóa tài khoản")
        edit_title = "Edit account" if en else "Sửa tài khoản"
        reset_title = "Reset password" if en else "Đặt lại mật khẩu"
        save_title = "Save changes" if en else "Lưu thay đổi"
        lock_icon = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 10V7a5 5 0 0 1 10 0v3"/><rect x="4" y="10" width="16" height="11" rx="2"/><path d="M12 14v3"/></svg>' if r["active"] else '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 10V7a5 5 0 0 1 9.5-2"/><rect x="4" y="10" width="16" height="11" rx="2"/><path d="M12 14v3"/></svg>'
        edit_icon = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L8 18l-4 1 1-4Z"/></svg>'
        reset_icon = ''
        body += f'<section class="user-card account-row"><div class="user-card-head"><div class="account-identity"><h3 class="account-primary-username">{E(r["username"])}</h3><div class="account-fullname">{E(r["full_name"] or ("—"))}</div><div class="user-meta"><span class="account-role {"role-admin" if r["role"] == "admin" else "role-user"}"><span class="account-role-mark" aria-hidden="true">{"✦" if r["role"] == "admin" else "●"}</span>{role_label}</span><span class="account-status {"is-active" if r["active"] else "is-locked"}"><i></i>{active_label}</span><span class="account-dept">{E(department_label(req,r["department"] or ""))}</span></div></div><div class="account-icon-actions">'
        if r["role"] == "user" and r["id"] != u["id"]:
            body += f'<form method="post" action="/users/{r["id"]}/toggle"><button class="account-icon-btn {"is-unlock" if not r["active"] else "is-lock"}" type="submit" title="{E(lock_title)}" aria-label="{E(lock_title)}">{lock_icon}</button></form>'
        elif r["role"] == "admin":
            admin_protected_title = "Administrator accounts cannot be locked" if en else "Tài khoản quản trị không thể bị khóa"
            body += f'<span class="account-icon-btn is-admin-protected" title="{E(admin_protected_title)}" aria-label="{E(admin_protected_title)}" role="img"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3 5 6v5c0 5 3 8 7 10 4-2 7-5 7-10V6Z"/><path d="m9 12 2 2 4-4"/></svg></span>'
        body += f'<details class="account-action-details"><summary class="account-icon-btn" title="{E(edit_title)}" aria-label="{E(edit_title)}">{edit_icon}</summary><div class="account-popover"><form method="post" action="/users/{r["id"]}/edit"><div class="user-edit-grid"><div><label>{tr(req,"username")}</label><input class="w" name="username" value="{E(r["username"])}" required maxlength="80"></div><div><label>{tr(req,"fullname")}</label><input class="w" name="full_name" value="{E(r["full_name"] or "")}" required maxlength="150"></div><div><label>{tr(req,"dept")}</label><select class="w" name="department" required>{''.join(f'<option value="{E(d)}" {"selected" if (r["department"] or "").strip().casefold()==d.casefold() else ""}>{E(department_label(req,d))}</option>' for d in department_options())}</select></div><div><label>{tr(req,"phone")}</label><input class="w" name="phone" value="{E(r["phone"] or "")}" maxlength="50"></div></div><p><button type="submit">{save_title}</button></p></form></div></details>'
        body += f'<details class="account-action-details"><summary class="account-icon-btn is-reset" title="{E(reset_title)}" aria-label="{E(reset_title)}"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 12a9 9 0 1 0 2.64-6.36"/><path d="M3 4v6h6"/><path d="M12 7v5l3 2"/></svg></summary><div class="account-popover reset-popover"><form method="post" action="/users/{r["id"]}/reset"><div class="reset-inline-row"><label for="reset-password-{r["id"]}">{tr(req,"password")}</label><input id="reset-password-{r["id"]}" class="w" type="password" name="new" minlength="6" title="{tr(req,"hint6")}" placeholder="{tr(req,"newpw")}" required autocomplete="new-password"><button type="submit">{reset_title}</button></div></form></div></details>'
        if r["role"] == "user":
            ql_label = (("Passwordless login: ON" if r["passwordless_login"] else "Passwordless login: OFF") if en else ("Đăng nhập không mật khẩu: BẬT" if r["passwordless_login"] else "Đăng nhập không mật khẩu: TẮT"))
            ql_icon = ('<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3c-4.4 0-8 3.6-8 8v2"/><path d="M8 8.5A4 4 0 0 1 16 10v3a4 4 0 0 1-8 0v-1"/><path d="M12 7v7"/><path d="M4 17c1.5 2.5 4.4 4 8 4s6.5-1.5 8-4"/></svg>' if r["passwordless_login"] else '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3c-4.4 0-8 3.6-8 8v2"/><path d="M8 8.5A4 4 0 0 1 16 10v3a4 4 0 0 1-8 0v-1"/><path d="M12 7v7"/><path d="M4 17c1.5 2.5 4.4 4 8 4s6.5-1.5 8-4"/><path d="m4 4 16 16"/></svg>')
            ql_state_class = "is-passwordless-on" if r["passwordless_login"] else "is-passwordless-off"
            body += f'<form method="post" action="/users/{r["id"]}/passwordless"><button class="account-icon-btn is-passwordless {ql_state_class}" type="submit" title="{E(ql_label)}" aria-label="{E(ql_label)}">{ql_icon}</button></form>'
        if r["role"] == "user" and r["id"] != u["id"]:
            delete_title = tr(req, "delete_user")
            confirm_delete = tr(req, "confirm_delete_user")
            body += f'<form method="post" action="/users/{r["id"]}/delete" onsubmit="return confirm(\'{E(confirm_delete)}\')"><button class="account-icon-btn is-delete-user" type="submit" title="{E(delete_title)}" aria-label="{E(delete_title)}"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6M10 10v6M14 10v6"/></svg></button></form>'
        body += '</div></div></section>'
    body += '</div></div>'
    body += '''<style>
    .account-list-heading{display:flex;align-items:center;gap:10px;margin-bottom:12px}.account-list-heading h2{margin:0;color:#173f68;font-size:19px;font-weight:800;letter-spacing:-.3px}.account-count{display:inline-flex;align-items:center;justify-content:center;min-width:25px;height:25px;padding:0 7px;border-radius:8px;background:#eaf4ff;border:1px solid #d7e9fb;color:#1769aa;font-size:11px;font-weight:800}
    .account-list{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}.account-row.user-card{box-sizing:border-box;padding:12px 14px!important;margin:0!important;border:1px solid #e0eaf4!important;border-radius:12px!important;background:#fff!important;box-shadow:0 2px 8px rgba(30,75,120,.035)!important;overflow:visible!important;min-width:0}.account-row .user-card-head{display:flex!important;align-items:center!important;justify-content:space-between!important;gap:10px!important;padding:0!important;margin:0!important;border:0!important;min-width:0}.account-identity{min-width:0;flex:1}.account-row .user-card-head h3{font-size:13px!important;font-weight:800!important;color:#173f68!important;line-height:1.35!important;margin:0!important;overflow-wrap:anywhere}.account-row .user-card-head h3.account-primary-username{font-size:13px!important;font-weight:850!important;letter-spacing:.05px;color:#173f68!important}.account-fullname{font-size:10px;color:#8497ac;margin-top:2px;line-height:1.35;overflow-wrap:anywhere}.account-row .user-meta{display:flex!important;align-items:center!important;gap:5px!important;flex-wrap:wrap!important;margin-top:6px!important;padding:0!important;border:0!important;background:none!important;font-size:10px!important;line-height:1.3!important;font-weight:600!important;color:#5d7289!important}.account-role,.account-status,.account-dept{display:inline-flex;align-items:center;gap:4px;padding:3px 6px;border-radius:6px;border:1px solid #e2eaf3;background:#f7faff}.account-role{font-size:10px!important;font-weight:800!important;letter-spacing:.05px;white-space:nowrap}.account-role .account-role-mark{font-size:9px;line-height:1}.account-role.role-admin{color:#174ea6;background:linear-gradient(135deg,#edf4ff 0%,#e4edff 100%);border-color:#c7d9ff;box-shadow:inset 0 1px 0 rgba(255,255,255,.8)}.account-role.role-user{color:#087568;background:linear-gradient(135deg,#e9faf6 0%,#e2f6f1 100%);border-color:#bfe8dc}.account-status{font-size:10px!important}.account-status i{display:block;width:5px;height:5px;border-radius:50%;background:#16a36a}.account-status.is-locked{background:#fff4f2;border-color:#ffded8;color:#b42318}.account-status.is-locked i{background:#d04444}.account-dept{color:#3e6c96;background:#f0f7ff;border-color:#dcecfb}.account-icon-actions{display:flex;align-items:center;justify-content:flex-end;gap:5px;flex:0 0 auto}.account-icon-actions form{margin:0!important;display:flex;align-items:center}.account-icon-btn{width:30px;height:30px;min-width:30px!important;min-height:30px!important;display:inline-flex!important;align-items:center!important;justify-content:center!important;padding:0!important;border-radius:8px!important;background:#f0f6fc!important;color:#47739d!important;border:1px solid #dce9f5!important;box-shadow:none!important;cursor:pointer}.account-icon-btn svg{width:15px;height:15px;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}.account-icon-btn:hover{background:#e3f0ff!important;color:#1769aa!important;border-color:#bfdcf6!important}.account-icon-btn.is-lock{background:#fff5f3!important;color:#b5473c!important;border-color:#f6d9d4!important}.account-icon-btn.is-unlock{background:#eaf8f1!important;color:#21805a!important;border-color:#cdebdc!important}.account-icon-btn.is-delete-user{background:#fff0f0!important;color:#c62828!important;border-color:#f4caca!important}.account-icon-btn.is-delete-user:hover{background:#ffe1e1!important;color:#a91616!important;border-color:#e9a6a6!important}.account-icon-btn.is-reset{background:#f2efff!important;color:#6650a8!important;border-color:#e3dcff!important}.account-icon-btn.is-passwordless{background:#eaf8f4!important;color:#16816b!important;border-color:#cdeee4!important}.account-icon-btn.is-passwordless-on{background:#fff7e6!important;color:#a66a08!important;border-color:#f3dfad!important}.account-icon-btn.is-passwordless-off{background:#f0f6fc!important;color:#647b95!important;border-color:#dce9f5!important}.account-icon-btn.is-admin-protected{background:#eef5ff!important;color:#3568a4!important;border-color:#d8e6fa!important;cursor:help!important}.account-icon-btn.is-admin-protected:hover{transform:none!important;background:#eef5ff!important;color:#3568a4!important;border-color:#d8e6fa!important}.account-action-details{position:relative}.account-action-details>summary{list-style:none;cursor:pointer}.account-action-details>summary::-webkit-details-marker{display:none}.account-popover{position:absolute;right:0;top:36px;width:min(420px,75vw);z-index:30;padding:14px;background:#fff;border:1px solid #dce8f4;border-radius:12px;box-shadow:0 12px 35px rgba(25,65,105,.16)}.account-popover .user-edit-grid{display:grid!important;grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:8px!important}.account-popover label{display:block!important;margin:5px 0!important;font-size:10px!important;color:#60758d!important}.account-popover input,.account-popover select{width:100%!important;min-width:0!important;height:34px!important;padding:6px 8px!important;font-size:11px!important;border:1px solid #d7e3ef!important;border-radius:7px!important}.account-popover p{display:flex;justify-content:flex-end;margin:10px 0 0!important;padding-top:9px;border-top:1px solid #e7eef6}.account-popover p button{font-size:11px!important;padding:7px 11px!important;min-height:30px!important}.account-popover .help{font-size:10px!important;margin-top:4px}.reset-popover{width:min(510px,88vw);padding:12px!important;font-family:Inter,Segoe UI,Arial,sans-serif!important}.reset-popover form{margin:0!important}.reset-popover .reset-inline-row{display:grid;grid-template-columns:72px minmax(145px,1fr) auto;align-items:center;gap:9px;min-width:0}.reset-popover label{font-size:11px!important;font-weight:750!important;line-height:1.3!important;margin:0!important;color:#405a75!important;white-space:nowrap}.reset-popover input{box-sizing:border-box!important;width:100%!important;min-width:0!important;height:35px!important;padding:7px 9px!important;font-size:11px!important;line-height:1.35!important;font-family:inherit!important;background:#fbfdff!important;border-radius:7px!important}.reset-popover input::placeholder{color:#8293a7;opacity:1}.reset-popover input:focus{outline:3px solid #dbeafe;border-color:#80b7ee!important}.reset-popover .reset-inline-row button{display:inline-flex!important;align-items:center!important;justify-content:center!important;white-space:nowrap!important;min-height:35px!important;font-size:11px!important;font-weight:750!important;padding:7px 10px!important;border-radius:7px!important}.account-create-box .user-edit-grid{grid-template-columns:repeat(3,minmax(0,1fr))!important}.account-create-box{padding:18px!important}.account-create-box h2{font-size:18px;color:#173f68}.account-create-box .help{font-size:11px}.account-create-box label{font-size:11px}.account-create-box input,.account-create-box select{height:36px;font-size:12px}.account-create-box>form>p button{font-size:12px;font-weight:750}
    
    
    .account-toast{position:fixed;z-index:9999;top:22px;right:24px;display:flex;align-items:center;gap:11px;max-width:min(420px,calc(100vw - 32px));padding:13px 14px;border:1px solid #cdebdc;border-radius:13px;background:rgba(250,255,252,.97);color:#176548;box-shadow:0 12px 34px rgba(20,65,50,.16);font-size:13px;font-weight:700;line-height:1.45;animation:accountToastIn .28s cubic-bezier(.2,.8,.2,1) both;backdrop-filter:blur(10px)}.account-toast-icon{display:inline-flex;align-items:center;justify-content:center;flex:0 0 28px;width:28px;height:28px;border-radius:50%;background:#def7e9;color:#16815b}.account-toast-icon svg{width:16px;height:16px;fill:none;stroke:currentColor;stroke-width:2.2;stroke-linecap:round;stroke-linejoin:round}.account-toast-close{display:inline-flex!important;align-items:center!important;justify-content:center!important;flex:0 0 25px!important;width:25px!important;height:25px!important;min-height:25px!important;padding:0!important;margin-left:4px!important;border:0!important;border-radius:7px!important;background:transparent!important;color:#547767!important;font-size:20px!important;line-height:1!important;box-shadow:none!important}.account-toast-close:hover{background:#e7f5ed!important;color:#155c40!important}.account-toast.is-leaving{animation:accountToastOut .2s ease forwards}@keyframes accountToastIn{from{opacity:0;transform:translate3d(0,-12px,0) scale(.98)}to{opacity:1;transform:translate3d(0,0,0) scale(1)}}@keyframes accountToastOut{to{opacity:0;transform:translate3d(0,-8px,0);pointer-events:none}}.account-icon-btn{transition:background .16s ease,color .16s ease,border-color .16s ease,transform .16s ease!important;position:relative}.account-icon-btn:hover{transform:translateY(-1px)}.account-icon-btn:focus-visible,.account-action-details>summary:focus-visible{outline:3px solid #93c5fd!important;outline-offset:2px!important}.account-icon-btn svg{display:block;flex:0 0 auto;overflow:visible}.account-icon-actions{flex-wrap:nowrap;min-height:30px}.account-icon-btn[title]{cursor:pointer}
    
    .detail-field-note{grid-column:span 3!important;background:#eaf3ff!important;border:1px solid #cfe3fb!important;border-radius:11px!important;min-width:0!important;padding:12px 14px!important}
    .detail-field-note strong{display:flex!important;align-items:center!important;gap:8px!important;color:#245f9e!important;margin:0 0 8px!important}
    .detail-field-note .detail-icon{display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;border-radius:8px;background:#d5e7ff;color:#245f9e;flex:0 0 28px}
    .note-field-content{min-width:0;margin-left:36px}
    .note-value-chip{display:block;white-space:pre-wrap;overflow-wrap:anywhere;color:#193f63;font-weight:650;line-height:1.55}
    .note-empty{color:#71849a;font-size:12px;font-style:italic}
    .public-note-form{display:flex;flex-direction:column;gap:7px;min-width:0}
    .public-note-label{display:block;margin:0;color:#245f9e;font-size:12px;font-weight:650}
    .public-note-form textarea{display:block;width:100%;min-width:0;resize:vertical;border:1px solid #bfd8f4;background:#fff;border-radius:8px;padding:10px 11px;color:#193f63;font:inherit;font-size:13px;line-height:1.5;box-sizing:border-box}
    .public-note-actions{display:flex;justify-content:flex-end}
    .public-note-save{background:linear-gradient(135deg,#2876c7,#245f9e);color:#fff;border:0;border-radius:8px;padding:8px 13px;font-weight:700;cursor:pointer}
    
    
    /* Final unified palette: blue actions, amber in-progress, blue low severity. */
    .btn-take,.btn-close{background:#2563eb!important;color:#fff!important;border:1px solid #2563eb!important}
    .btn-take:hover,.btn-close:hover{background:#1d4ed8!important;border-color:#1d4ed8!important}
    .workflow-panel{background:#f8fafc!important;border-color:#e2e8f0!important}
    .workflow-close-action{background:#f8fafc!important;border-color:#e2e8f0!important}
    .workflow-note-help{color:#64748b!important;line-height:1.45}
    .st-new,.pill.st-new,.ticket-table .st-new{background:#dbeafe!important;color:#1d4ed8!important;border-color:#bfdbfe!important}
    .st-progress,.pill.st-progress,.ticket-table .st-progress{background:#fef3c7!important;color:#92400e!important;border-color:#fcd34d!important}
    .st-closed,.pill.st-closed,.ticket-table .st-closed{background:#dcfce7!important;color:#166534!important;border-color:#bbf7d0!important}
    .sev-low,.pill.sev-low,.priority-pill.sev-low,.priority-label.sev-low,.ticket-table .sev-low{background:#dbeafe!important;color:#1d4ed8!important;border-color:#bfdbfe!important}
    .sev-medium,.pill.sev-medium,.priority-pill.sev-medium,.priority-label.sev-medium,.ticket-table .sev-medium{background:#fef3c7!important;color:#92400e!important;border-color:#fcd34d!important}
    .sev-high,.pill.sev-high,.priority-pill.sev-high,.priority-label.sev-high,.ticket-table .sev-high{background:#fee2e2!important;color:#991b1b!important;border-color:#fecaca!important}
    /* Vạch màu trước STT đồng bộ với bảng Ticket gần đây trên trang chủ: trạng thái mới / đang xử lý / đã đóng. */
    .ticket-table tbody tr.ticket-row-new td:first-child{box-shadow:inset 3px 0 #168be8!important}
    .ticket-table tbody tr.ticket-row-progress td:first-child{box-shadow:inset 3px 0 #f59e0b!important}
    .ticket-table tbody tr.ticket-row-closed td:first-child{box-shadow:inset 3px 0 #20ad70!important}
    .ticket-row-progress td{background:#fffcf2!important}
    .user-ticket-progress{background:linear-gradient(135deg,#fff 0%,#fffcf2 100%)!important;border-left-color:#f59e0b!important}
    .detail-field-note{grid-column:span 3!important;box-sizing:border-box!important;min-height:104px!important;height:auto!important;display:flex!important;flex-direction:column!important;gap:7px!important;padding:10px 12px!important;background:#eff6ff!important;border:1px solid #bfdbfe!important;border-radius:11px!important;color:#1e3a5f!important;overflow:hidden!important}
    .detail-field-note .note-heading{display:flex!important;align-items:center!important;justify-content:space-between!important;gap:8px!important;flex:0 0 auto!important;min-width:0!important}
    .detail-field-note .note-heading strong{display:flex!important;align-items:center!important;gap:8px!important;margin:0!important;color:#1d4ed8!important;min-width:0!important}
    .detail-field-note .detail-icon{display:none!important}.detail-field-note .note-heading strong{font-size:12px!important;line-height:1.35!important}.note-heading-actions{display:flex!important;align-items:center!important;justify-content:flex-end!important;gap:4px!important;flex:0 0 auto!important}
    .note-field-content{margin:0!important;min-width:0!important;flex:1 1 auto!important;min-height:0!important;display:flex!important;flex-direction:column!important}
    .public-note-view{flex:1 1 auto!important;min-height:0!important;overflow:auto!important;white-space:pre-wrap!important;overflow-wrap:anywhere!important;line-height:1.5!important;color:#1e3a5f!important;padding:8px 10px!important;border:1px solid #dbeafe!important;border-radius:8px!important;background:rgba(255,255,255,.78)!important}
    .note-value-chip{display:block!important;white-space:pre-wrap!important;overflow-wrap:anywhere!important;color:#1e3a5f!important;font-weight:600!important;line-height:1.5!important}
    .note-empty{display:block!important;color:#64748b!important;font-size:12px!important;font-style:italic!important;font-weight:400!important}
    .public-note-form{display:flex!important;flex:1 1 auto!important;flex-direction:column!important;gap:7px!important;min-width:0!important;min-height:0!important}
    .public-note-form textarea{width:100%!important;min-width:0!important;max-width:100%!important;min-height:0!important;height:100%!important;flex:1 1 auto!important;box-sizing:border-box!important;resize:none!important;overflow:auto!important;border:1px solid #bfdbfe!important;background:#fff!important;border-radius:8px!important;padding:8px 10px!important;color:#1e3a5f!important;font:inherit!important;font-size:13px!important;line-height:1.5!important}
    .public-note-actions{display:flex!important;justify-content:flex-end!important;align-items:center!important;gap:6px!important;flex:0 0 auto!important}
    .note-icon-btn{display:inline-flex!important;align-items:center!important;justify-content:center!important;width:25px!important;height:25px!important;min-width:25px!important;min-height:25px!important;padding:0!important;border:1px solid #bfdbfe!important;border-radius:7px!important;background:#fff!important;color:#2563eb!important;cursor:pointer!important;box-shadow:none!important;flex:0 0 25px!important;transition:background .15s ease,color .15s ease,border-color .15s ease,opacity .15s ease!important}
    .note-icon-btn svg{width:13px!important;height:13px!important;fill:none!important;stroke:currentColor!important;stroke-width:2!important;stroke-linecap:round!important;stroke-linejoin:round!important}
    .note-icon-btn:hover{background:#dbeafe!important;color:#1d4ed8!important;border-color:#93c5fd!important}
    .note-icon-btn.note-save{background:#2563eb!important;color:#fff!important;border-color:#2563eb!important}.note-icon-btn.note-cancel{color:#64748b!important;border-color:#d7e2ee!important}.note-icon-btn:disabled{opacity:.42!important;cursor:default!important;box-shadow:none!important}.note-icon-btn:disabled:hover{background:#fff!important;color:inherit!important;border-color:#d7e2ee!important}
    .note-icon-btn.note-save:hover{background:#1d4ed8!important;border-color:#1d4ed8!important}
    [hidden]{display:none!important}
    
    
</style>
    <script>(function(){const toast=document.querySelector('.account-toast');if(!toast)return;function dismiss(){toast.classList.add('is-leaving');setTimeout(()=>toast.remove(),220)}const close=toast.querySelector('.account-toast-close');if(close)close.addEventListener('click',dismiss);setTimeout(dismiss,4200);})();</script>'''
    return page(req, tr(req, "users"), body, u)

@app.post("/users")
def users_add(req: Request, username: str = Form(...), full_name: str = Form(...), department: str = Form(...), phone: str = Form(""), password: str = Form(...), role: str = Form("user"), u=Depends(admin_user)):
    username = username.strip().lower()
    department = normalize_department(department)
    if not username or len(password) < 6 or role not in ("user", "admin") or not full_name.strip() or department not in department_options():
        raise HTTPException(400, tr(req, "invalid_data") + " — " + tr(req, "hint6"))
    try:
        query("INSERT INTO users(username,password_hash,full_name,department,phone,role) VALUES(?,?,?,?,?,?)",
              (username,pw_hash(password),full_name.strip(),department,phone.strip(),role), commit=True)
    except sqlite3.IntegrityError:
        raise HTTPException(400, tr(req, "username_exists"))
    return RedirectResponse("/users?msg=User+added", 303)

@app.post("/users/{uid}/edit")
def edit_user(req: Request, uid: int, username: str = Form(...), full_name: str = Form(...), department: str = Form(...), phone: str = Form(""), u=Depends(admin_user)):
    username = username.strip().lower()
    department = normalize_department(department)
    if not username or not full_name.strip() or department not in department_options():
        raise HTTPException(400, tr(req, "invalid_data"))
    if not query("SELECT id FROM users WHERE id=?", (uid,), one=True): raise HTTPException(404, tr(req, "not_found"))
    try:
        query("UPDATE users SET username=?,full_name=?,department=?,phone=? WHERE id=?", (username,full_name.strip(),department,phone.strip(),uid), commit=True)
    except sqlite3.IntegrityError:
        raise HTTPException(400, tr(req, "username_exists"))
    return RedirectResponse("/users?msg=User+updated", 303)

@app.post("/users/{uid}/passwordless")
def toggle_passwordless(req: Request, uid: int, u=Depends(admin_user)):
    target = query("SELECT id,role FROM users WHERE id=?", (uid,), one=True)
    if not target: raise HTTPException(404, tr(req, "not_found"))
    if target["role"] != "user": raise HTTPException(400, tr(req, "invalid_data"))
    next_enabled = not bool(query("SELECT passwordless_login FROM users WHERE id=?", (uid,), one=True)["passwordless_login"])
    query("UPDATE users SET passwordless_login=? WHERE id=? AND role='user'", (1 if next_enabled else 0, uid), commit=True)
    next_msg = "Passwordless enabled" if next_enabled else "Passwordless disabled"
    return RedirectResponse("/users?msg=" + next_msg.replace(" ", "+"), 303)

@app.post("/users/{uid}/reset")
def reset_user(req: Request, uid: int, new: str = Form(...), u=Depends(admin_user)):
    if len(new) < 6: raise HTTPException(400, tr(req, "hint6"))
    if not query("SELECT id FROM users WHERE id=?", (uid,), one=True): raise HTTPException(404, tr(req, "not_found"))
    query("UPDATE users SET password_hash=? WHERE id=?", (pw_hash(new),uid), commit=True)
    return RedirectResponse("/users?msg=Password+reset", 303)

@app.post("/users/{uid}/delete")
def delete_user(req: Request, uid: int, u=Depends(admin_user)):
    if uid == u["id"]:
        raise HTTPException(400, tr(req, "cannot_self_lock"))
    target = query("SELECT id,role FROM users WHERE id=?", (uid,), one=True)
    if not target:
        raise HTTPException(404, tr(req, "not_found"))
    # Protect administrator accounts from accidental removal/lockout.
    if target["role"] != "user":
        message = "Administrator accounts cannot be deleted" if req.session.get("lang") == "en" else "Không thể xóa tài khoản quản trị"
        raise HTTPException(400, message)
    # Keep every ticket and feedback record. Detach account references before deleting the user row.
    query("UPDATE tickets SET user_id=NULL WHERE user_id=?", (uid,), commit=True)
    query("UPDATE tickets SET created_by_user_id=NULL WHERE created_by_user_id=?", (uid,), commit=True)
    feedback_cols = {row["name"] for row in query("PRAGMA table_info(ticket_feedback)")}
    if "reviewer_user_id" in feedback_cols:
        query("UPDATE ticket_feedback SET reviewer_user_id=NULL WHERE reviewer_user_id=?", (uid,), commit=True)
    query("DELETE FROM users WHERE id=? AND role='user'", (uid,), commit=True)
    msg = "User deleted; tickets retained" if req.session.get("lang") == "en" else "Đã xóa tài khoản; ticket được giữ lại"
    from urllib.parse import quote
    return RedirectResponse("/users?msg=" + quote(msg), 303)

@app.post("/users/{uid}/toggle")
def toggle_user(req: Request, uid: int, u=Depends(admin_user)):
    if uid == u["id"]:
        raise HTTPException(400, tr(req, "cannot_self_lock"))
    target = query("SELECT id,role,active FROM users WHERE id=?", (uid,), one=True)
    if not target:
        raise HTTPException(404, tr(req, "not_found"))
    # Only regular User accounts may be locked/unlocked; Admin accounts are protected server-side.
    if target["role"] != "user":
        message = "Administrator accounts cannot be locked" if req.session.get("lang") == "en" else "Không thể khóa tài khoản quản trị"
        raise HTTPException(400, message)
    next_msg = "User unlocked" if not target["active"] else "User locked"
    query("UPDATE users SET active=1-active WHERE id=? AND role='user'", (uid,), commit=True)
    return RedirectResponse("/users?msg=" + next_msg.replace(" ", "+"), 303)
