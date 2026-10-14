"""
KyujinBox (KBX) RPA - Tự động kiểm tra ứng viên mới và ghi vào Google Sheet
- Đọc danh sách công ty từ Sheet マスター管理 (cùng file với AW)
- Chỉ chạy các dòng có cột E = "KBX"
- Chạy mỗi 30 phút qua Windows Task Scheduler (hoặc GitHub Actions)

Cấu trúc cột Master Sheet (KBX):
  A: 会社名
  B: (không dùng)
  C: ID (email đăng nhập)
  D: PW
  E: Media = "KBX"
  F: Mã công ty (VD: 8056-8676)
  G: Sheet ID đích (URL hoặc ID thuần)
  H: Tên tab đích
  I: Link nền tảng KBX (VD: https://saiyo.kyujinbox.com)

Cấu trúc cột Sheet đích (A–N) — ĐÃ CẬP NHẬT 2026-08 theo header thực tế
(thêm カナ ở C và 現在の職業 ở L so với bản cũ, apply_number dời từ L → N):
  A: 応募日時 (VD: 2026/05/12 13:33)
  B: 氏名
  C: カナ (furigana, VD: しんとくりえ)
  D: 性別
  E: 生年月日 (VD: 2004/11/3)
  F: 年齢 (VD: 21歳)
  G: メールアドレス
  H: 電話番号
  I: 住所
  J: 学校名
  K: 応募求人名
  L: 現在の職業
  M: 対応ステータス
  N: 応募ID (dedup key, VD: A2-7491-3914)

Cách chạy:
  python kbx_main.py              → chế độ bình thường (chỉ check mới)
  python kbx_main.py --full       → quét toàn bộ tất cả trang (lần đầu)
  python kbx_main.py --visible    → hiện trình duyệt để debug
  python kbx_main.py --full --visible

  python kbx_main.py --login "email@congty.com"
                                   → đăng nhập TAY 1 lần cho 1 tài khoản KBX
                                     (mở trình duyệt, điền sẵn email/pass,
                                     bạn chỉ cần bấm Login / tick reCAPTCHA
                                     NẾU nó xuất hiện) — dùng khi đăng nhập
                                     tự động thất bại (xem bên dưới).
  python kbx_main.py --login-all  → làm bước trên cho TẤT CẢ tài khoản KBX
                                     trong Master Sheet, lần lượt từng cái.

  (2026-08 cập nhật): kbx_main.py giờ TỰ ĐỘNG đăng nhập (điền form + bấm
  login) khi chưa có session hoặc session đã hết hạn — không cần chạy tay
  --login trước nữa trong đa số trường hợp, vì trang /login KHÔNG luôn
  hiển thị reCAPTCHA. Script chỉ dừng lại và yêu cầu chạy --login tay khi:
    - Phát hiện reCAPTCHA thật sự hiển thị trên trang lúc đăng nhập, hoặc
    - Đăng nhập tự động thất bại vì lý do khác (sai email/mật khẩu, lỗi
      mạng, redirect bất thường...).
  Sau khi đăng nhập tự động thành công, session mới sẽ được lưu lại như
  bình thường để lần chạy sau dùng luôn, không cần đăng nhập lại.

═══════════════════════════════════════════════════════════════════════════
GHI CHÚ CẬP NHẬT (2026-08) — 管理画面 mới ("採用ボード"):
  Sau khi so sánh với HTML thực tế lấy từ trình duyệt, các thay đổi so với
  bản trước:

  1. KHÔNG CÒN bước "chọn công ty" qua /ptr/l-accounts + tab 直接投稿.
     Sau khi login, hệ thống redirect THẲNG vào:
       https://saiyo.kyujinbox.com/company/groups/G{corp_code}-0001/jobs
     → 1 tài khoản = 1 group cố định. Đã BỎ find_and_enter_company() và
       select_direct_tab() khỏi luồng chạy chính.

  2. URL 応募者一覧 vẫn là:
       {origin}/company/groups/G{corp_code}-0001/applications
     (không đổi).

  3. THAY VÌ click vào từng ứng viên rồi chờ JS cập nhật panel (dễ lỗi,
     phải retry nhiều lần), giờ có thể vào THẲNG URL chi tiết:
       {origin}/company/groups/G{corp_code}-0001/applications/{apply_number}
     Trang này trả về đầy đủ HTML với drawer đã mở sẵn (is-drawerOn),
     không cần chờ JS lâu, không cần đóng/mở panel. Đây là cách chính
     dùng để scrape trong bản này.

  4. FIX BUG: has_next_page_kbx() trước đây tìm `li.c-pager__button--next`
     trên toàn trang — selector này trùng với 1 pager KHÁC nằm trong modal
     "求人を選択してレポートを表示" (component Vue, luôn hiển thị is-active
     bất kể danh sách ứng viên có phân trang hay không) → gây false
     positive khiến script tưởng luôn còn trang kế tiếp.
     Bản sửa chỉ chấp nhận pager KHÔNG nằm trong .c-modal / .c-modal__wrap.

  5. Các label trong drawer chi tiết (氏名, 性別, 生年月日, メールアドレス,
     電話番号, 住所, 応募求人, ステータス) không đổi tên → logic scrape theo
     label vẫn dùng lại được gần như nguyên vẹn.

  6. FIX (2026-08 lần 2): Sheet đích thực tế có thêm 2 cột mới (カナ ở C,
     現在の職業 ở L) so với bản trước, khiến apply_number (dedup key) dời
     từ L → N. Code cũ ghi cứng 12 cột A–L theo thứ tự cũ nên bị:
       - Lệch cột: job_name/status/apply_number bị ghi nhầm vào 学校名/
         応募求人名/現在の職業 (vì thứ tự cột cũ không còn khớp).
       - Trùng lặp: dedup đọc nhầm cột L (giờ không còn chứa apply_number)
         nên không bao giờ nhận ra ứng viên cũ → ghi lặp mỗi lần chạy.
       - Thiếu dữ liệu: カナ và 学校名 chưa từng được scrape.
     Đã sửa: scrape thêm カナ (lấy từ span.c-note cạnh tên trong tiêu đề
     drawer, không phải từ li.c-borderList__item) và 学校名/現在の職業
     (từ li.c-borderList__item như các field khác); ghi đúng theo layout
     A–N mới; dedup đọc cột N.
═══════════════════════════════════════════════════════════════════════════
"""

import asyncio
import logging
import re
import sys
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout
from google.oauth2 import service_account
from googleapiclient.discovery import build

def get_proxy():
    """Proxy Nhật (tuỳ chọn) qua biến môi trường PROXY_SERVER/PROXY_USER/PROXY_PASS."""
    import os
    server = os.getenv("PROXY_SERVER")
    if not server:
        return None
    proxy = {"server": server}
    if os.getenv("PROXY_USER"):
        proxy["username"] = os.environ["PROXY_USER"]
        proxy["password"] = os.getenv("PROXY_PASS", "")
    return proxy


# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
CRED_FILE  = BASE_DIR / "credentials" / "credentials.json"
LOG_DIR    = BASE_DIR / "logs"
LOG_FILE   = LOG_DIR / f"kbx_{datetime.now():%Y%m%d}.log"

# Thư mục lưu session đăng nhập (cookie) — tránh phải login lại mỗi lần chạy
# (vì form login có reCAPTCHA v2 checkbox, không thể tự động vượt qua).
# Session được tạo bằng: python kbx_main.py --login "email" (login thủ công 1 lần).
AUTH_STATE_DIR = BASE_DIR / "auth_state"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ── Master Sheet config (cùng file với AW) ────────────────────────────────────
MASTER_SHEET_ID = "1sCYbBWFU0ENhZrmlRMqZMomS--jaZT5YbPb-mu2ORno"
MASTER_TAB      = "マスター管理"

# Chỉ số cột (0-based, khớp với A=0, B=1, ...)
# ĐÃ SỬA (2026-08): đối chiếu ảnh chụp Sheet マスター管理 thực tế —
#   A=No.(không dùng), B=企業名, C=ID, D=PW, E=媒体,
#   F=企業ID(KBXの場合)／通知メール(INDEEDの場合), G=応募者シートID,
#   H=タブ名, I=媒体URL
# → COL_COMPANY trước đây trỏ vào A (No.) thay vì B (企業名), đã sửa lại.
COL_COMPANY    = 1   # B — 企業名 (trước đây sai là cột A = No.)
COL_ID         = 2   # C
COL_PW         = 3   # D
COL_MEDIA      = 4   # E
COL_CORP_CODE  = 5   # F — 企業ID (KBX). Giờ chỉ là DỰ PHÒNG — nếu trống,
                      #     script tự lấy corp_code từ URL sau khi login
                      #     (xem login_kbx). Không còn bắt buộc phải điền tay.
COL_SHEET_URL  = 6   # G
COL_TAB_NAME   = 7   # H
COL_BASE_URL   = 8   # I

# ── Retry / timing config ──────────────────────────────────────────────────────
MAX_DETAIL_RETRIES   = 3
DETAIL_WAIT_MS       = 1000

# ── Logging ───────────────────────────────────────────────────────────────────
LOG_DIR.mkdir(exist_ok=True)
AUTH_STATE_DIR.mkdir(exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
# Parse helpers
# ══════════════════════════════════════════════════════════════════════════════

def parse_birthday_kbx(text: str) -> str:
    """'2004年11月3日（21歳）' → '2004/11/3'"""
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", text)
    return f"{m.group(1)}/{m.group(2)}/{m.group(3)}" if m else ""

def parse_age_kbx(text: str) -> str:
    """'2004年11月3日（21歳）' → '21歳'"""
    m = re.search(r"（(\d+)歳）", text)
    return f"{m.group(1)}歳" if m else ""

def parse_apply_date_kbx(text: str) -> str:
    """
    Field 応募日時・応募No. giờ gộp chung dạng:
      '2026/05/12 13:33\\nA2-7491-3914'
    Chỉ lấy dòng đầu (ngày giờ), không đổi múi giờ.
    """
    line = text.strip().splitlines()[0].strip()
    if re.match(r"\d{4}/\d{2}/\d{2} \d{2}:\d{2}", line):
        return line
    m = re.search(r"(\d{4}/\d{2}/\d{2} \d{2}:\d{2})", text)
    return m.group(1) if m else line

def extract_sheet_id(url: str) -> str:
    if "/d/" in url:
        return url.split("/d/")[1].split("/")[0]
    return url.strip()


def state_file_for_account(username: str) -> Path:
    """
    File lưu session (cookie) riêng cho từng tài khoản KBX, dựa theo email
    đăng nhập.
    """
    safe = re.sub(r"[^a-zA-Z0-9_.-]", "_", username)
    return AUTH_STATE_DIR / f"kbx_{safe}.json"


def meta_file_for_account(username: str) -> Path:
    """
    File lưu kèm origin + corp_code phát hiện được lúc đăng nhập tay
    (--login). Cần thiết vì trang /login của KyujinBox KHÔNG tự redirect
    người đã đăng nhập sang dashboard — nên không thể chỉ dựa vào việc
    "vào lại /login" để biết session còn hợp lệ hay không. Thay vào đó,
    lần sau ta vào THẲNG URL dashboard đã biết (lưu ở đây) — nếu cookie
    hết hạn, trang bảo vệ đó sẽ tự đá về /login, đáng tin cậy hơn nhiều.
    """
    safe = re.sub(r"[^a-zA-Z0-9_.-]", "_", username)
    return AUTH_STATE_DIR / f"kbx_{safe}.meta.json"


def save_session_meta(username: str, origin: str, corp_code: str):
    import json
    meta_file_for_account(username).write_text(
        json.dumps({"origin": origin, "corp_code": corp_code}, ensure_ascii=False),
        encoding="utf-8",
    )


def load_session_meta(username: str) -> dict | None:
    import json
    f = meta_file_for_account(username)
    if not f.exists():
        return None
    try:
        return json.loads(f.read_text(encoding="utf-8"))
    except Exception:
        return None


# ── Giảm dấu hiệu tự động hoá (fingerprint) khi tạo browser context ──────────
# Đây là các thiết lập RPA tiêu chuẩn (viewport/ngôn ngữ/timezone thật, ẩn cờ
# navigator.webdriver) khi tự động hoá TÀI KHOẢN CỦA CHÍNH MÌNH — KHÔNG phải
# kỹ thuật để né/giải reCAPTCHA (captcha vẫn phải giải tay ở bước --login).
_STEALTH_INIT_SCRIPT = """
Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
window.chrome = { runtime: {} };
Object.defineProperty(navigator, 'languages', { get: () => ['ja-JP', 'ja'] });
Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3] });
"""

def new_context_kwargs() -> dict:
    return dict(
        viewport={"width": 1440, "height": 900},
        locale="ja-JP",
        timezone_id="Asia/Tokyo",
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        extra_http_headers={"Accept-Language": "ja-JP,ja;q=0.9"},
    )


# ══════════════════════════════════════════════════════════════════════════════
# Google Sheets
# ══════════════════════════════════════════════════════════════════════════════

def get_sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        str(CRED_FILE), scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)


def read_master_accounts_kbx(service) -> list[dict]:
    """Đọc Master Sheet, chỉ lấy dòng có cột E = 'KBX'."""
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SHEET_ID,
            range=f"{MASTER_TAB}!A:I",
        ).execute()
        rows = result.get("values", [])
    except Exception as e:
        log.error(f"❌ Không đọc được Master Sheet: {e}")
        sys.exit(1)

    accounts = []
    for i, row in enumerate(rows[1:], start=2):
        while len(row) <= COL_BASE_URL:
            row.append("")

        media = row[COL_MEDIA].strip()
        if media != "KBX":
            continue

        company   = row[COL_COMPANY].strip()
        id_val    = row[COL_ID].strip()
        pw_val    = row[COL_PW].strip()
        corp_code = row[COL_CORP_CODE].strip()
        sheet_url = row[COL_SHEET_URL].strip()
        tab_name  = row[COL_TAB_NAME].strip()
        raw_url   = row[COL_BASE_URL].strip().rstrip("/")
        parsed    = urlparse(raw_url)
        base_url  = f"{parsed.scheme}://{parsed.netloc}"
        login_url = raw_url

        # LƯU Ý: corp_code (cột F) KHÔNG còn bắt buộc — nếu trống, script sẽ
        # tự lấy corp_code từ URL sau khi login thành công (xem login_kbx).
        # Chỉ id/pw/sheet_url/base_url là bắt buộc.
        if not all([id_val, pw_val, sheet_url, base_url]):
            log.warning(f"  ⚠️  Dòng {i} [{company}] thiếu thông tin (ID/PW/Sheet/URL) — bỏ qua")
            continue

        if not corp_code:
            log.info(f"  ℹ️  Dòng {i} [{company}]: cột F (企業ID) trống — sẽ tự lấy corp_code từ URL sau khi login")

        accounts.append({
            "company":   company,
            "id":        id_val,
            "pw":        pw_val,
            "corp_code": corp_code,
            "sheet_url": sheet_url,
            "tab_name":  tab_name or "【KBX】応募者リスト",
            "base_url":  base_url,
            "login_url": login_url,
        })
        log.info(f"  ✓ Dòng {i}: [{company}] {id_val}  corp={corp_code}")

    return accounts


def get_existing_ids_kbx(service, sheet_id: str, tab: str) -> set:
    """Đọc cột N (応募ID) làm dedup key (đã dời từ L → N theo layout mới)."""
    existing = set()
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"{tab}!N3:N",
        ).execute()
        for row in result.get("values", []):
            if row and row[0].strip():
                existing.add(row[0].strip())
    except Exception as e:
        log.warning(f"  ⚠️  Không đọc được cột L Sheet: {e}")
    return existing


def get_next_empty_row_kbx(service, sheet_id: str, tab: str) -> int:
    """Tìm dòng trống tiếp theo bằng cột B (氏名)."""
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"{tab}!B3:B",
        ).execute()
        rows = result.get("values", [])
        last_data_row = 2
        for i, row in enumerate(rows):
            if row and row[0].strip():
                last_data_row = i + 3
        return last_data_row + 1
    except Exception as e:
        log.warning(f"  ⚠️  Không đọc được dòng cuối: {e} — ghi vào dòng 3")
        return 3


def append_one_row_kbx(service, sheet_id: str, tab: str, applicant: dict):
    """
    Ghi 1 ứng viên vào dòng trống cuối.
    Layout A–N (cập nhật 2026-08 theo header thực tế của Sheet đích):
      A: apply_date   B: name       C: kana        D: gender
      E: birthday     F: age        G: email       H: tel
      I: address      J: school     K: job_name    L: occupation
      M: status       N: apply_number (dedup key)
    """
    next_row = get_next_empty_row_kbx(service, sheet_id, tab)

    row_data = [[
        applicant.get("apply_date", ""),    # A
        applicant.get("name", ""),          # B
        applicant.get("kana", ""),          # C
        applicant.get("gender", ""),        # D
        applicant.get("birthday", ""),      # E
        applicant.get("age", ""),           # F
        applicant.get("email", ""),         # G
        applicant.get("tel", ""),           # H
        applicant.get("address", ""),       # I
        applicant.get("school", ""),        # J
        applicant.get("job_name", ""),      # K
        applicant.get("occupation", ""),    # L
        applicant.get("status", ""),        # M
        applicant.get("apply_number", ""),  # N ← dedup key
    ]]

    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{tab}!A{next_row}",
        valueInputOption="USER_ENTERED",
        body={"values": row_data},
    ).execute()

    log.info(f"    📝 Ghi dòng {next_row}: {applicant.get('name', '')}  [{applicant.get('apply_number', '')}]")


# ══════════════════════════════════════════════════════════════════════════════
# Login
# ══════════════════════════════════════════════════════════════════════════════
#
# QUAN TRỌNG (2026-08): Form login KBX có reCAPTCHA v2 dạng checkbox
# ("私はロボットではありません") — KHÔNG thể tự động vượt qua bằng script.
# Vì vậy luồng chính giờ là:
#   1. Chạy kbx_main.py --login MỘT LẦN (đăng nhập tay, tự tick captcha) để
#      lưu session/cookie vào auth_state/kbx_{email}.json
#   2. kbx_main.py chỉ REUSE session đó (try_resume_session_kbx) — không tự
#      động điền form + bấm login nữa, vì chắc chắn sẽ bị chặn bởi captcha.
#   3. Khi session hết hạn, script sẽ báo rõ và bỏ qua tài khoản đó — cần
#      chạy lại kbx_main.py --login "email" để làm mới session.
#
# Hàm login_kbx() cũ (điền form + click) không còn dùng trong luồng tự động
# nữa — chỉ còn manual_login_one() (chạy qua --login) dùng logic tương tự
# để điền sẵn email/password cho người dùng đỡ phải gõ tay.
# ══════════════════════════════════════════════════════════════════════════════

async def try_resume_session_kbx(page, username: str) -> tuple[str, str] | None:
    """
    Thử dùng session/cookie đã lưu (từ --login) để vào thẳng dashboard mà
    KHÔNG cần qua form login (né captcha hoàn toàn).

    QUAN TRỌNG: không kiểm tra bằng cách vào lại /login — trang đó KHÔNG
    tự redirect người đã đăng nhập, nên sẽ luôn báo nhầm "hết hạn". Thay
    vào đó, vào THẲNG URL dashboard đã biết từ lần --login trước (lưu
    trong file .meta.json) — nếu cookie hết hạn, trang bảo vệ sẽ tự đá
    về /login, đáng tin cậy hơn nhiều.

    Trả về (actual_origin, corp_code) nếu session còn hợp lệ, None nếu
    session hết hạn / chưa có (cần đăng nhập tay lại).
    """
    meta = load_session_meta(username)
    if not meta or not meta.get("origin") or not meta.get("corp_code"):
        log.warning("  ⚠️  Chưa có thông tin dashboard đã lưu (chưa từng --login) — cần đăng nhập tay")
        return None

    origin    = meta["origin"]
    corp_code = meta["corp_code"]
    probe_url = f"{origin}/company/groups/G{corp_code}-0001/jobs"

    try:
        await page.goto(probe_url, wait_until="domcontentloaded", timeout=30_000)
        await page.wait_for_timeout(1500)
        current = page.url

        if "/login" in current or await page.query_selector("#BtnLogin"):
            log.warning("  ⚠️  Session hết hạn — bị đá về trang login")
            return None

        if "kyujinbox.com" not in current:
            log.warning(f"  ⚠️  URL bất thường sau khi resume session: {current}")
            return None

        parsed = urlparse(current)
        actual_origin = f"{parsed.scheme}://{parsed.netloc}"

        # Corp_code có thể vẫn còn trong URL hiện tại — ưu tiên giá trị mới nhất
        m = re.search(r"/company/groups/G([\d\-]+)-0001", current)
        detected_corp_code = m.group(1) if m else corp_code

        log.info(f"  ✅ Dùng session đã lưu — origin: {actual_origin}, corp={detected_corp_code}")
        return actual_origin, detected_corp_code

    except Exception as e:
        log.error(f"  ❌ Lỗi khi resume session: {e}")
        return None


async def detect_recaptcha(page) -> bool:
    """
    Kiểm tra xem trang hiện tại có đang HIỂN THỊ widget reCAPTCHA hay không
    (iframe recaptcha, div.g-recaptcha, ...). Dùng để quyết định có nên tiếp
    tục đăng nhập tự động hay phải dừng lại nhờ người dùng làm tay.

    Lưu ý: HTML tĩnh của trang /login (class="recaptcha-form") KHÔNG có
    sẵn iframe recaptcha — điều này chỉ xuất hiện lúc runtime nếu server
    quyết định bật captcha (ví dụ sau vài lần đăng nhập sai, IP lạ, v.v.).
    Vì vậy hàm này phải kiểm tra trên trang ĐÃ TẢI XONG (runtime DOM),
    không dựa vào HTML nguồn.
    """
    try:
        selectors = [
            'iframe[src*="recaptcha"]',
            'iframe[title*="recaptcha" i]',
            'div.g-recaptcha',
            '#recaptcha',
            '[data-sitekey]',
        ]
        for sel in selectors:
            el = await page.query_selector(sel)
            if el and await el.is_visible():
                return True
        return False
    except Exception:
        return False


async def login_kbx(page, base_url: str, username: str, password: str) -> tuple[str, str] | None:
    """
    Đăng nhập KyujinBox TỰ ĐỘNG (điền form + bấm login).

    Trả về (actual_origin, detected_corp_code) nếu thành công.
    Trả về None nếu thất bại VÌ BẤT KỲ LÝ DO GÌ, bao gồm:
      - Phát hiện reCAPTCHA đang hiển thị trên trang (không thể tự vượt qua
        được, phải dừng lại và nhờ chạy --login để làm tay).
      - Sai email/mật khẩu, lỗi mạng, redirect bất thường, v.v.
      actual_origin       VD: https://saiyo.kyujinbox.com
      detected_corp_code  Lấy từ chính URL redirect sau login, VD: "9370-8852"
                           (rỗng "" nếu không parse được — khi đó sẽ dùng
                           corp_code từ cột F trong Master Sheet làm dự phòng)

    Sau login thành công, hệ thống mới redirect thẳng vào:
      https://saiyo.kyujinbox.com/company/groups/G{corp_code}-0001/jobs
    (không còn qua bước chọn công ty trung gian nữa) — nhờ vậy corp_code
    có thể lấy tự động từ URL, không cần điền tay vào Sheet nữa.
    """
    login_url = base_url
    try:
        log.info(f"  → Truy cập: {login_url}")
        await page.goto(login_url, wait_until="domcontentloaded", timeout=30_000)
        await page.wait_for_timeout(2000)

        # Kiểm tra reCAPTCHA NGAY khi trang vừa tải xong, trước khi điền gì cả.
        if await detect_recaptcha(page):
            log.warning(
                "  ⚠️  Phát hiện reCAPTCHA đang hiển thị trên trang login — "
                "không thể tự động vượt qua. Dừng lại, cần đăng nhập tay."
            )
            return None

        email_sel = '#login_email, input[name="login[email]"], input[type="email"]'
        await page.fill(email_sel, username)
        # Bắn thêm event input/blur — page.fill() không phải lúc nào cũng
        # kích hoạt đủ JS listener mà form dùng để enable nút #BtnLogin
        # (đặc biệt khi form có validate riêng, VD với email chứa ký tự "+").
        await page.dispatch_event(email_sel, "input")
        await page.dispatch_event(email_sel, "blur")
        log.info("  → Đã điền email")

        pw_sel = '#login_password, input[name="login[password]"], input[type="password"]'
        await page.fill(pw_sel, password)
        await page.dispatch_event(pw_sel, "input")
        await page.dispatch_event(pw_sel, "blur")
        log.info("  → Đã điền password")

        # Chờ chủ động cho tới khi nút hết disabled, thay vì click ngay
        # (nút #BtnLogin chỉ enable sau khi JS validate xong 2 field).
        btn_sel = '#BtnLogin, button[type="submit"]'
        try:
            await page.wait_for_selector(f"{btn_sel}:not([disabled])", timeout=8_000)
        except PlaywrightTimeout:
            # Nút vẫn disabled — có thể form đang báo lỗi validate (VD định
            # dạng email/mật khẩu không hợp lệ theo JS phía client).
            # Thử lại 1 lần bằng cách gõ lại ký tự cuối (đôi khi JS chỉ lắng
            # nghe sự kiện bàn phím thật, không nhận input/blur giả lập).
            log.warning("  ⚠️  Nút login vẫn disabled sau khi điền — thử gõ lại...")
            await page.click(email_sel)
            await page.keyboard.press("End")
            await page.keyboard.type(" ")
            await page.keyboard.press("Backspace")
            await page.click(pw_sel)
            await page.keyboard.press("End")
            await page.keyboard.type(" ")
            await page.keyboard.press("Backspace")
            try:
                await page.wait_for_selector(f"{btn_sel}:not([disabled])", timeout=5_000)
            except PlaywrightTimeout:
                log.error("  ❌ Nút login vẫn disabled — có thể form báo lỗi validate email/password")
                await page.screenshot(path=str(LOG_DIR / "kbx_login_disabled.png"))
                (LOG_DIR / "kbx_login_disabled.html").write_text(
                    await page.content(), encoding="utf-8"
                )
                return None

        # Kiểm tra lại lần nữa — 1 số site chỉ bật recaptcha SAU khi form đã
        # được điền hợp lệ (để giảm tải), ngay trước khi cho phép submit.
        if await detect_recaptcha(page):
            log.warning(
                "  ⚠️  reCAPTCHA xuất hiện sau khi điền form — dừng lại, "
                "cần đăng nhập tay."
            )
            return None

        await page.click(btn_sel, timeout=5_000)
        log.info("  → Đã click login")

        try:
            await page.wait_for_url(
                lambda u: "login" not in u and "kyujinbox.com" in u,
                timeout=20_000,
            )
        except PlaywrightTimeout:
            pass

        await page.wait_for_timeout(2000)
        current = page.url
        log.info(f"  → URL sau login: {current}")

        if "kyujinbox.com" not in current:
            log.warning("  ⚠️  Đăng nhập thất bại")
            await page.screenshot(path=str(LOG_DIR / "kbx_login_fail.png"))
            return None

        parsed = urlparse(current)
        actual_origin = f"{parsed.scheme}://{parsed.netloc}"

        # Tự lấy corp_code từ URL redirect, VD:
        #   /company/groups/G9370-8852-0001/jobs  →  "9370-8852"
        detected_corp_code = ""
        m = re.search(r"/company/groups/G([\d\-]+)-0001", current)
        if m:
            detected_corp_code = m.group(1)
            log.info(f"  🔎 Tự phát hiện corp_code từ URL: {detected_corp_code}")

        log.info(f"  ✅ Đăng nhập thành công — origin: {actual_origin}")
        return actual_origin, detected_corp_code

    except Exception as e:
        log.error(f"  ❌ Lỗi đăng nhập: {e}")
        return None


async def dismiss_message_popup(page):
    """
    Dismiss popup đồng ý メッセージ nếu xuất hiện.
    Nút: 同意せずにもどる (không đồng ý, quay lại) để không trigger side effects.
    """
    try:
        popup_sel = 'button:has-text("同意せずにもどる"), button:has-text("同意して始める")'
        el = await page.wait_for_selector(popup_sel, timeout=3_000)
        if el:
            back_btn = await page.query_selector('button:has-text("同意せずにもどる")')
            if back_btn:
                await back_btn.click()
                log.info("  → Đã đóng popup メッセージ同意")
                await page.wait_for_timeout(800)
    except PlaywrightTimeout:
        pass
    except Exception as e:
        log.warning(f"  ⚠️  Lỗi dismiss popup: {e}")


async def goto_applicant_list(page, base_url: str, corp_code: str) -> bool:
    """
    Vào trang 応募者一覧.
    URL: {base_url}/company/groups/G{corp_code}-0001/applications
    (Không còn bước chọn công ty trung gian — đã xác nhận qua HTML thực tế
    rằng 1 tài khoản = 1 group cố định, vào thẳng được luôn.)
    """
    try:
        group_id  = f"G{corp_code}-0001"
        direct_url = f"{base_url}/company/groups/{group_id}/applications"
        await page.goto(direct_url, wait_until="networkidle", timeout=30_000)
        await page.wait_for_timeout(1500)
        log.info(f"  → 応募者一覧: {page.url}")

        await dismiss_message_popup(page)
        return True

    except Exception as e:
        log.error(f"  ❌ Lỗi vào応募者一覧: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
# Lấy danh sách ứng viên trong trang
# ══════════════════════════════════════════════════════════════════════════════

async def get_applicant_links(page) -> list[dict]:
    """
    Lấy danh sách ứng viên từ các link s-drawer_open.
    Mỗi link có data-application-number và text = tên ứng viên.
    (selector này không đổi so với trước, đã xác nhận qua HTML thực tế)
    """
    links = await page.query_selector_all("a.s-drawer_open[data-application-number]")
    result = []
    for i, link in enumerate(links):
        apply_number = (await link.get_attribute("data-application-number") or "").strip()
        name         = (await link.inner_text()).strip()
        if apply_number:
            result.append({
                "apply_number":    apply_number,
                "name_from_list":  name,
                "element_index":   i,
            })
    log.info(f"  📋 {len(result)} ứng viên trong trang")
    return result


async def has_next_page_kbx(page) -> bool:
    """
    Kiểm tra có trang tiếp theo không.

    FIX: selector 'li.c-pager__button--next' trước đây bị nhầm với 1 pager
    KHÁC nằm trong modal "求人を選択してレポートを表示" (component Vue,
    luôn hiển thị is-active bất kể danh sách ứng viên có phân trang hay
    không) → gây false positive khiến script tưởng luôn còn trang kế tiếp.

    Bản sửa: chỉ chấp nhận phần tử next-page KHÔNG nằm trong bất kỳ
    .c-modal / .c-modal__wrap nào (pager thật của 応募者一覧, nếu có, luôn
    nằm ngay trong nội dung trang, không phải trong modal).
    """
    try:
        candidates = await page.query_selector_all('li.c-pager__button--next')
        for el in candidates:
            in_modal = await el.evaluate(
                "el => !!el.closest('.c-modal, .c-modal__wrap')"
            )
            if in_modal:
                continue
            cls = await el.get_attribute("class") or ""
            if "is-active" in cls:
                log.debug(f"    Next-page hợp lệ tìm thấy: class='{cls}'")
                return True
        return False
    except Exception as e:
        log.debug(f"    has_next_page error: {e}")
    return False


async def goto_applicant_page(page, base_url: str, page_num: int, corp_code: str):
    """
    Chuyển trang trong 応募者一覧.
    URL: {base_url}/company/groups/G{corp_code}-0001/applications?slmsg=d&pg=N
    (giữ nguyên pattern cũ — CHƯA xác nhận được có thật sự tồn tại phân
    trang trong UI mới hay không, vì tài khoản test không có đủ ứng viên
    để phân trang. has_next_page_kbx() đã được sửa để không false-positive,
    nên nếu thực tế không có phân trang, vòng lặp sẽ tự dừng ở trang 1.)
    """
    group_id = f"G{corp_code}-0001"
    if page_num == 1:
        url = f"{base_url}/company/groups/{group_id}/applications"
    else:
        url = f"{base_url}/company/groups/{group_id}/applications?slmsg=d&pg={page_num}"
    log.debug(f"    goto page {page_num}: {url}")
    try:
        await page.goto(url, wait_until="networkidle", timeout=30_000)
    except PlaywrightTimeout:
        await page.goto(url, wait_until="domcontentloaded", timeout=30_000)
    await page.wait_for_timeout(1500)


# ══════════════════════════════════════════════════════════════════════════════
# Chi tiết ứng viên — vào THẲNG URL thay vì click + chờ JS cập nhật panel
# ══════════════════════════════════════════════════════════════════════════════

async def _dump_detail_debug(page, apply_number: str, attempt: int):
    """Dump screenshot + HTML khi không lấy được thông tin."""
    try:
        slug = f"kbx_FAIL_{apply_number}_try{attempt}"
        await page.screenshot(path=str(LOG_DIR / f"{slug}.png"), full_page=False)
        (LOG_DIR / f"{slug}.html").write_text(await page.content(), encoding="utf-8")
        log.warning(f"    🔍 DEBUG dump: {slug}.png + .html")
    except Exception as e:
        log.warning(f"    ⚠️  Không dump được: {e}")


async def scrape_applicant_detail(page, apply_number: str) -> dict:
    """
    Đọc thông tin ứng viên từ drawer sau khi đã điều hướng thẳng tới:
      {base_url}/company/groups/G{corp_code}-0001/applications/{apply_number}

    Cấu trúc HTML thực tế (xác nhận qua DevTools 2026-08):
      <div id="drawerContainer">
        <div class="c-drawer s-drawer is-drawerOn">
          ...
          <main class="c-drawer_body ...">
            <section class="c-panel u-p30">
              <ul class="c-borderList ...">
                <li class="c-borderList__item">
                  <h3 class="c-title ...">氏名</h3>
                  <p class="c-lines">成田 蓮</p>
                </li>
                ...
    Mỗi field nằm trong li.c-borderList__item, có h3 label + p.c-lines value.
    """
    data = {"apply_number": apply_number}

    await dismiss_message_popup(page)

    # ── Chờ drawer render xong (main bên trong #drawerContainer) ─────────
    try:
        await page.wait_for_selector(
            "#drawerContainer main p.c-lines, #drawerContainer main li.c-borderList__item",
            state="visible", timeout=8_000,
        )
    except PlaywrightTimeout:
        log.warning(f"    ⚠️  Drawer không render nội dung cho #{apply_number}")
        return {}

    await page.wait_for_timeout(300)

    # Giới hạn truy vấn trong #drawerContainer để không dính nhầm phần tử
    # cùng class ở nơi khác trên trang.
    # LƯU Ý: <h1> chứa カナ nằm NGOÀI <main> (anh em với <main>, cùng trong
    # .c-drawer__in) nên phải dùng #drawerContainer làm gốc, không chỉ main.
    drawer_root = await page.query_selector("#drawerContainer")
    container   = await page.query_selector("#drawerContainer main")
    if not container:
        container = page  # fallback — không nên xảy ra
    if not drawer_root:
        drawer_root = container

    items = await container.query_selector_all("li.c-borderList__item")

    async def get_field(label_text: str) -> str:
        for li in items:
            h = await li.query_selector("h3.c-title, h4.c-title, [class*='c-title']")
            if h and (await h.inner_text()).strip() == label_text:
                p = await li.query_selector("p.c-lines")
                if p:
                    return (await p.inner_text()).strip()
        return ""

    # ── 氏名 ───────────────────────────────────────────────────────────
    name_raw = await get_field("氏名")
    name_raw = re.sub(r'\s*[\(（][^\)）]{1,30}[\)）]', '', name_raw).strip()
    data["name"] = name_raw

    if not data["name"]:
        log.warning(f"    ⚠️  Không lấy được tên cho #{apply_number}")
        return {}

    # ── カナ ───────────────────────────────────────────────────────────
    # LƯU Ý: カナ KHÔNG nằm trong li.c-borderList__item như các field khác —
    # nó nằm trong tiêu đề drawer, dạng:
    #   <h1 class="c-title ...">成田 蓮<span class="c-note ...">(なりた れん)</span></h1>
    data["kana"] = ""
    try:
        title_el = await drawer_root.query_selector("h1.c-title, .l-content__head h1")
        if title_el:
            note_el = await title_el.query_selector("span.c-note")
            if note_el:
                kana_raw = (await note_el.inner_text()).strip()
                data["kana"] = re.sub(r'^[\(（]|[\)）]$', '', kana_raw).strip()
    except Exception as e:
        log.debug(f"    カナ scrape error: {e}")

    # ── 性別 ───────────────────────────────────────────────────────────
    gender_raw = await get_field("性別")
    data["gender"] = gender_raw if gender_raw in ("男性", "女性", "男", "女") else ""

    # ── 生年月日 + 年齢 ────────────────────────────────────────────────
    bday_raw = await get_field("生年月日")
    if bday_raw and re.search(r"\d{4}年\d{1,2}月\d{1,2}日", bday_raw):
        data["birthday"] = parse_birthday_kbx(bday_raw)
        data["age"]      = parse_age_kbx(bday_raw)
    else:
        data["birthday"] = ""
        data["age"]      = ""

    # ── メールアドレス ────────────────────────────────────────────────
    email_raw = await get_field("メールアドレス")
    data["email"] = email_raw if ("@" in email_raw and " " not in email_raw and len(email_raw) < 100) else ""

    # ── 電話番号 ──────────────────────────────────────────────────────
    tel_raw = await get_field("電話番号")
    clean_tel = re.sub(r"[-\s　]", "", tel_raw)
    data["tel"] = tel_raw if re.match(r"^0\d{9,10}$", clean_tel) else ""

    # ── 住所 ──────────────────────────────────────────────────────────
    data["address"] = await get_field("住所")

    # ── 学校名 ────────────────────────────────────────────────────────
    data["school"] = await get_field("学校名")

    # ── 現在の職業 ────────────────────────────────────────────────────
    data["occupation"] = await get_field("現在の職業")

    # ── 応募求人 ──────────────────────────────────────────────────────
    data["job_name"] = ""
    for li in items:
        h3 = await li.query_selector("h3.c-title, h4.c-title, [class*='c-title']")
        if h3 and (await h3.inner_text()).strip() in ("応募求人", "応募求人名"):
            p = await li.query_selector("p.c-lines")
            if p:
                job_link = await p.query_selector("a.c-link, a[href*='/jobs/edit/']")
                data["job_name"] = (await job_link.inner_text()).strip() if job_link else (await p.inner_text()).strip()
            break

    # ── 応募日時・応募No. (gộp chung 1 field trong UI mới) ─────────────
    data["apply_date"] = ""
    for li in items:
        p = await li.query_selector("p.c-lines")
        if p:
            t = (await p.inner_text()).strip()
            if re.search(r"\d{4}/\d{2}/\d{2} \d{2}:\d{2}", t):
                data["apply_date"] = parse_apply_date_kbx(t)
                break

    # ── 対応ステータス ────────────────────────────────────────────────
    data["status"] = ""
    try:
        sel_el = await page.query_selector(
            '#application_status, select[name="application[status]"]'
        )
        if sel_el:
            selected_text = await sel_el.evaluate(
                "el => el.options[el.selectedIndex] ? el.options[el.selectedIndex].text : ''"
            )
            data["status"] = selected_text.strip()
    except Exception:
        pass

    log.debug(f"    scraped: {data}")
    return data


async def fetch_applicant_detail_kbx(page, base_url: str, corp_code: str, apply_info: dict, index: int) -> dict | None:
    """
    Điều hướng thẳng tới URL chi tiết ứng viên và scrape.
    Thay thế hoàn toàn cách click + chờ JS cập nhật panel + retry cũ —
    ổn định hơn nhiều vì đây là trang server-render đầy đủ, không phải
    trạng thái JS tạm thời.
    """
    apply_number   = apply_info["apply_number"]
    name_from_list = apply_info.get("name_from_list", "").strip()
    group_id       = f"G{corp_code}-0001"
    detail_url     = f"{base_url}/company/groups/{group_id}/applications/{apply_number}"

    for attempt in range(1, MAX_DETAIL_RETRIES + 1):
        try:
            await page.goto(detail_url, wait_until="networkidle", timeout=30_000)
            await page.wait_for_timeout(DETAIL_WAIT_MS)

            detail = await scrape_applicant_detail(page, apply_number)

            if not detail.get("name") and name_from_list:
                log.info(f"    [{index}] 🔄 Dùng tên từ list: {name_from_list}")
                detail["name"] = name_from_list
                detail["apply_number"] = apply_number

            if detail.get("name"):
                if attempt > 1:
                    log.info(f"    [{index}] ✅ Lấy được sau {attempt} lần thử: {detail['name']}")
                return detail

            log.warning(f"    [{index}] Attempt {attempt}/{MAX_DETAIL_RETRIES}: tên rỗng — #{apply_number}")
            await _dump_detail_debug(page, apply_number, attempt)
            await page.wait_for_timeout(800)

        except PlaywrightTimeout:
            log.warning(f"    [{index}] Attempt {attempt}/{MAX_DETAIL_RETRIES}: timeout — #{apply_number}")
            await page.wait_for_timeout(800)
        except Exception as e:
            log.warning(f"    [{index}] Attempt {attempt}/{MAX_DETAIL_RETRIES}: exception — {e}")
            await page.wait_for_timeout(600)

    if name_from_list:
        log.warning(f"    [{index}] ⚠️  Dùng tên từ list (fallback cuối): {name_from_list} #{apply_number}")
        return {"apply_number": apply_number, "name": name_from_list}

    log.error(f"    [{index}] ❌ Hết retry, bỏ qua #{apply_number}")
    return None


# ══════════════════════════════════════════════════════════════════════════════
# Async generator: yield từng ứng viên mới
# ══════════════════════════════════════════════════════════════════════════════

async def iter_new_applicants_kbx(page, base_url: str, corp_code: str, page_num: int, existing_ids: set):
    """Quét HẾT tất cả ứng viên trong trang, bỏ qua những ID đã có."""
    if page_num == 1:
        try:
            await page.screenshot(path=str(LOG_DIR / "kbx_debug_entries.png"), timeout=5_000)
            log.info(f"  📸 kbx_debug_entries.png — URL: {page.url}")
        except Exception:
            pass

    applicant_list = await get_applicant_links(page)

    if not applicant_list:
        log.warning(f"  ⚠️  Không thấy ứng viên ở trang {page_num}")
        (LOG_DIR / f"kbx_empty_p{page_num}.html").write_text(
            await page.content(), encoding="utf-8"
        )
        return

    list_url = page.url  # để quay lại sau khi vào từng trang chi tiết

    for i, apply_info in enumerate(applicant_list):
        apply_number = apply_info["apply_number"]

        if apply_number in existing_ids:
            log.info(f"    [{i+1}] #{apply_number} đã có — bỏ qua")
            continue

        detail = await fetch_applicant_detail_kbx(page, base_url, corp_code, apply_info, i + 1)

        if detail is None:
            continue

        log.info(f"    [{i+1}] ✓ {detail['name']}  {detail.get('apply_date', '')}")
        yield detail

    # Quay lại trang danh sách (vì fetch_applicant_detail_kbx đã điều hướng đi)
    try:
        await page.goto(list_url, wait_until="networkidle", timeout=30_000)
        await page.wait_for_timeout(1000)
    except PlaywrightTimeout:
        await page.goto(list_url, wait_until="domcontentloaded", timeout=30_000)


# ══════════════════════════════════════════════════════════════════════════════
# Process 1 account
# ══════════════════════════════════════════════════════════════════════════════

async def process_account_kbx(browser, account: dict, sheets_service, full_scan: bool):
    company   = account["company"]
    username  = account["id"]
    password  = account["pw"]
    corp_code = account["corp_code"]
    login_url = account["login_url"]  # full:   https://secure.kyujinbox.com/login
    sheet_url = account["sheet_url"]
    tab_name  = account["tab_name"]

    log.info(f"\n{'='*60}")
    log.info(f"🏢 {company}  corp={corp_code}  ({'全件スキャン' if full_scan else '新着チェック'})")
    log.info(f"   ID : {username}")

    sheet_id     = extract_sheet_id(sheet_url)
    existing_ids = get_existing_ids_kbx(sheets_service, sheet_id, tab_name)
    log.info(f"  📊 Sheet đã có {len(existing_ids)} ứng viên")

    state_file = state_file_for_account(username)
    have_session = state_file.exists() and meta_file_for_account(username).exists()

    context_kwargs = new_context_kwargs()
    if have_session:
        context_kwargs["storage_state"] = str(state_file)
    context = await browser.new_context(**context_kwargs)
    await context.add_init_script(_STEALTH_INIT_SCRIPT)
    page = await context.new_page()

    try:
        # 1. Nếu đã có session lưu sẵn, thử dùng nó để vào thẳng dashboard
        #    (nhanh nhất, không cần qua form login).
        login_result = None
        if have_session:
            login_result = await try_resume_session_kbx(page, username)
            if not login_result:
                log.warning(
                    f"  ⚠️  Session cũ hết hạn cho [{company}] ({username}) — "
                    f"thử đăng nhập tự động lại..."
                )

        # 2. Không có session (hoặc đã hết hạn) → thử ĐĂNG NHẬP TỰ ĐỘNG
        #    (điền email/password + bấm login). Chỉ dừng lại nhờ người dùng
        #    làm tay nếu login_kbx() phát hiện reCAPTCHA thật sự xuất hiện
        #    trên trang, hoặc login thất bại vì lý do khác (sai mật khẩu...).
        if not login_result:
            login_result = await login_kbx(page, login_url, username, password)
            if login_result:
                actual_origin, detected_corp_code = login_result
                try:
                    await context.storage_state(path=str(state_file))
                    save_session_meta(username, actual_origin, detected_corp_code or corp_code)
                    log.info(f"  💾 Đã lưu session mới cho [{company}]")
                except Exception as e:
                    log.warning(f"  ⚠️  Không lưu được session mới: {e}")

        if not login_result:
            log.error(
                f"  ❌ Không thể đăng nhập cho [{company}] ({username}) — "
                f"đăng nhập tự động thất bại (có thể do reCAPTCHA xuất hiện, "
                f"sai email/mật khẩu, hoặc lỗi mạng).\n"
                f"     → Chạy tay 1 lần: python kbx_main.py --login \"{username}\"  "
                f"(tick reCAPTCHA nếu có), sau đó chạy lại kbx_main.py."
            )
            return
        actual_origin, detected_corp_code = login_result

        # 2. Xác định corp_code dùng cho phần còn lại:
        #    - Ưu tiên corp_code lấy được từ URL redirect (luôn đúng, real-time)
        #    - Nếu không phát hiện được, dùng corp_code khai trong Master Sheet (cột F)
        if detected_corp_code:
            if corp_code and corp_code != detected_corp_code:
                log.warning(
                    f"  ⚠️  corp_code trong Sheet ({corp_code}) khác với URL thực tế "
                    f"({detected_corp_code}) — dùng giá trị từ URL"
                )
            corp_code = detected_corp_code
        elif not corp_code:
            log.error("  ❌ Không có corp_code (cả URL lẫn Sheet đều trống) — bỏ qua tài khoản")
            return

        # 3. Vào thẳng 応募者一覧 — KHÔNG còn bước chọn công ty trung gian
        #    (đã xác nhận: login redirect thẳng vào group của corp_code này)
        ok = await goto_applicant_list(page, actual_origin, corp_code)
        if not ok:
            return

        page_num      = 1
        total_written = 0

        while True:
            log.info(f"  📄 Trang {page_num}...")

            applicant_list = await get_applicant_links(page)
            total_on_page  = len(applicant_list)

            async for applicant in iter_new_applicants_kbx(page, actual_origin, corp_code, page_num, existing_ids):
                try:
                    append_one_row_kbx(sheets_service, sheet_id, tab_name, applicant)
                    existing_ids.add(applicant["apply_number"])
                    total_written += 1
                    log.info(f"    ✅ Đã ghi: {applicant['name']}")
                except Exception as e:
                    log.error(f"    ❌ Ghi sheet thất bại: {e}")

            if total_on_page == 0:
                log.warning("  ⚠️  Trang trống — dừng")
                break

            if not await has_next_page_kbx(page):
                log.info(f"  ✅ Đã quét hết {page_num} trang (không còn 次へ)")
                break

            page_num += 1
            log.info(f"  ➡️  Chuyển sang trang {page_num}...")
            await goto_applicant_page(page, actual_origin, page_num, corp_code)

        log.info(f"  🆕 Tổng đã ghi: {total_written} ứng viên")

        try:
            shot = LOG_DIR / f"kbx_{company}_{datetime.now():%H%M%S}.png"
            await page.screenshot(path=str(shot), timeout=5_000)
            log.info(f"  📸 {shot.name}")
        except Exception:
            pass

    finally:
        # Lưu lại session mới nhất (cookie có thể được server refresh/xoay
        # vòng trong lúc dùng) để lần chạy sau vẫn còn hợp lệ lâu nhất có thể.
        try:
            await context.storage_state(path=str(state_file))
        except Exception:
            pass
        await context.close()


# ══════════════════════════════════════════════════════════════════════════════
# Đăng nhập tay 1 lần (để lưu session, né reCAPTCHA) — gộp vào cùng 1 file
# ══════════════════════════════════════════════════════════════════════════════
#
# Cách dùng:
#   python kbx_main.py --login "email@congty.com"   → đăng nhập tay 1 tài khoản
#   python kbx_main.py --login-all                  → đăng nhập tay lần lượt
#                                                       tất cả tài khoản KBX
#
# Trình duyệt sẽ mở ra (không headless), email/password điền sẵn — bạn chỉ
# cần tick reCAPTCHA + bấm ログイン. Script tự lưu session lại, từ đó
# `python kbx_main.py` (không cờ) sẽ chạy headless bình thường mà không cần
# đăng nhập lại nữa, cho tới khi session hết hạn.
# ══════════════════════════════════════════════════════════════════════════════

async def manual_login_one(account: dict) -> bool:
    company   = account["company"]
    username  = account["id"]
    password  = account["pw"]
    login_url = account["login_url"]
    state_file = state_file_for_account(username)

    print(f"\n{'='*60}")
    print(f"🏢 {company}  ({username})")
    print(f"{'='*60}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, proxy=get_proxy())
        context = await browser.new_context(**new_context_kwargs())
        await context.add_init_script(_STEALTH_INIT_SCRIPT)
        page = await context.new_page()
        try:
            await page.goto(login_url, wait_until="domcontentloaded", timeout=30_000)
            await page.wait_for_timeout(1500)

            email_sel = '#login_email, input[name="login[email]"], input[type="email"]'
            pw_sel    = '#login_password, input[name="login[password]"], input[type="password"]'
            await page.fill(email_sel, username)
            await page.fill(pw_sel, password)

            print("👉 Đã điền sẵn email/password.")
            print("👉 Vui lòng: tick reCAPTCHA → (giải thử thách ảnh nếu có) → bấm ログイン")
            print("⏳ Đang chờ bạn đăng nhập xong (tối đa 5 phút)...")

            try:
                await page.wait_for_url(
                    lambda u: "login" not in u and "kyujinbox.com" in u,
                    timeout=300_000,
                )
            except PlaywrightTimeout:
                print(f"❌ Hết thời gian chờ (5 phút) — bỏ qua [{company}].")
                return False

            await page.wait_for_timeout(1500)
            current = page.url
            if "login" in current:
                print(f"❌ Vẫn còn ở trang login — đăng nhập chưa thành công. Bỏ qua [{company}].")
                return False

            # Lưu session (cookie) + meta (origin, corp_code) để lần sau
            # kbx_main.py vào thẳng dashboard mà không cần qua /login.
            await context.storage_state(path=str(state_file))
            parsed = urlparse(current)
            actual_origin = f"{parsed.scheme}://{parsed.netloc}"
            m = re.search(r"/company/groups/G([\d\-]+)-0001", current)
            corp_code = m.group(1) if m else ""
            save_session_meta(username, actual_origin, corp_code)

            print(f"✅ Đăng nhập thành công! Đã lưu session: {state_file.name}")
            print(f"   URL sau login: {current}")
            if corp_code:
                print(f"   corp_code phát hiện: {corp_code}")
            else:
                print("   ⚠️  Không phát hiện được corp_code từ URL — kiểm tra lại thủ công nếu cần.")
            return True
        finally:
            await context.close()
            await browser.close()


async def run_manual_login(target: str):
    sheets_service = get_sheets_service()
    accounts = read_master_accounts_kbx(sheets_service)

    if not accounts:
        print("❌ Không tìm thấy tài khoản KBX nào trong Master Sheet.")
        return

    if target == "--login-all":
        targets = accounts
    else:
        targets = [a for a in accounts if a["id"].strip() == target]
        if not targets:
            print(f"❌ Không tìm thấy tài khoản với email: {target}")
            print("   Các email KBX hiện có trong Master Sheet:")
            for a in accounts:
                print(f"   - {a['id']}  ({a['company']})")
            return

    for account in targets:
        await manual_login_one(account)

    print(f"\n{'='*60}")
    print("✅ Hoàn tất. Giờ có thể chạy: python kbx_main.py")
    print(f"{'='*60}")


# ══════════════════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════════════════

async def main(headless: bool = True, full_scan: bool = False):
    log.info(f"\n{'#'*60}")
    log.info(f"# KyujinBox RPA  —  {datetime.now():%Y-%m-%d %H:%M:%S}")
    log.info(f"# Mode: {'headless' if headless else 'visible'} | {'【全件スキャン】' if full_scan else '【新着チェック】'}")
    log.info(f"{'#'*60}")

    sheets_service = get_sheets_service()

    log.info(f"\n📋 Đọc danh sách KBX từ Sheet マスター管理...")
    accounts = read_master_accounts_kbx(sheets_service)
    log.info(f"✅ Tìm thấy {len(accounts)} tài khoản KBX")

    if not accounts:
        log.warning("Không có tài khoản KBX nào!")
        return

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless, proxy=get_proxy())
        for account in accounts:
            await process_account_kbx(browser, account, sheets_service, full_scan)
        await browser.close()

    log.info("\n✅ Hoàn thành tất cả!")


if __name__ == "__main__":
    if "--login-all" in sys.argv:
        asyncio.run(run_manual_login("--login-all"))
    elif "--login" in sys.argv:
        idx = sys.argv.index("--login")
        if idx + 1 >= len(sys.argv):
            print('❌ Thiếu email. Dùng: python kbx_main.py --login "email@congty.com"')
            sys.exit(1)
        asyncio.run(run_manual_login(sys.argv[idx + 1].strip()))
    else:
        headless  = "--visible" not in sys.argv
        full_scan = "--full"    in sys.argv
        asyncio.run(main(headless=headless, full_scan=full_scan))
