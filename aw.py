"""
AirWork RPA - Tự động kiểm tra ứng viên mới và ghi vào Google Sheet
- Đọc danh sách công ty từ Sheet マスター管理
- Chỉ chạy các dòng có cột E = "AW"
- Chạy mỗi 30 phút qua Windows Task Scheduler

Cách chạy:
  python main.py              → chế độ bình thường (chỉ check mới)
  python main.py --full       → quét toàn bộ tất cả trang (lần đầu)
  python main.py --visible    → hiện trình duyệt để debug
  python main.py --full --visible

FIX LOG:
  [FIX-1] Lệch dòng đầu / đổi màu: dùng values.update() + tìm dòng trống thực
  [FIX-2] Trùng ứng viên: check bằng apply_id (unique key từ AirWork)
  [FIX-3] Bỏ sót ứng viên: chế độ thường quét TẤT CẢ row trong trang 1
           thay vì dừng sớm khi gặp người đã có
  [FIX-5] Lẫn 住所 / 応募求人名: dùng selector cụ thể trong modal,
           không dùng vòng lặp heuristic cho địa chỉ
  [FIX-6] Modal load chậm / tên không lấy được:
           - Retry tối đa MAX_MODAL_RETRIES lần mỗi ứng viên
           - Chờ modal ổn định trước khi đọc (wait_for_load_state)
           - Dump HTML + screenshot khi thất bại để debug
           - Thêm nhiều selector dự phòng cho 氏名
  [FIX-7] Bỏ sót ứng viên ở các trang sau (stale DOM):
           - Re-fetch danh sách row_ids SAU mỗi lần đóng modal
             thay vì dùng list đã lấy từ đầu trang
           - Reload trang nếu phát hiện row bị mất
  [FIX-8] apply_date: thử nhiều selector + parse thêm các định dạng
  [FIX-9] apply_date: hỗ trợ format mới span.styles_dateTime__MPfnz (2026/5/26 13:34)
  [FIX-13] Cột D–I (性別/生年月日/年齢/メール/電話/住所) rỗng: AirWork đổi layout +
           hash class CSS-module:
             - hash đổi styles_infoText__UN0fw → __Sznk8 → bỏ hard-code hash,
               dùng [class*="infoText"]
             - icon "human" thêm <div class="styles_humanInfoColumn__.."> bọc giữa
               → '+' (adjacent) không khớp → dùng '~' (general sibling) + fallback
  [FIX-14] Ghi TRÙNG apply_id (cột N): Sheet lưu cột N dạng số, đọc lại có dấu
           phẩy (29,760,116) ≠ apply_id scrape (29760116) → dedup fail. Sửa:
           đọc UNFORMATTED_VALUE + _norm_id (chỉ giữ chữ số) ở cả 2 phía +
           run_seen chống trùng trong cùng 1 lần chạy.
  [FIX-15] Tải CV (PDF レジュメ) về Drive + gắn link vào cột O:
           - Mỗi công ty có folder Drive riêng, link để sẵn ở cột K tab マスター管理
           - Khi modal đang mở: click tab 応募情報 → lấy link <a download> →
             fetch PDF bằng page.request (cookie auth) → lưu tạm vào /resumes
           - Upload PDF lên folder Drive (supportsAllDrives=True cho Shared Drive)
             → ghi webViewLink vào cột O (レジュメ), xoá file tạm sau khi upload
           - Cần thêm scope drive + service Drive (credentials đã bật Drive API)
  [FIX-16] "Modal không xuất hiện (timeout)": scrape_modal cũ chờ cứng
           MODAL_SEL chứa hash styles_container__BMWEr; AirWork đổi hash →
           không khớp → báo timeout giả dù panel đã mở. Sửa: chờ trực tiếp
           tên ứng viên (h1/title), fallback scope = cả page nếu hash đổi.
  [FIX-17] Tên = "応募者" (placeholder) + báo "không có CV" oan: panel load chậm
           (nhất khi có PDF レジュメ) → tên/link tải chưa render. Bản cũ chấp nhận
           placeholder làm tên thật. Sửa:
             - KHÔNG chấp nhận "応募者"/"候補者": chờ networkidle + retry; hết
               lượt vẫn placeholder → trả {} để click_and_scrape() click lại
             - download_resume: chờ link <a download> xuất hiện (tối đa ~12s)
               trước khi kết luận không có CV
"""

import asyncio
import logging
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload  # [FIX-15] upload CV lên Drive

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
CRED_FILE  = BASE_DIR / "credentials" / "credentials.json"
LOG_DIR    = BASE_DIR / "logs"
RESUME_DIR = BASE_DIR / "resumes"   # [FIX-15] nơi lưu tạm CV trước khi upload Drive
LOG_FILE   = LOG_DIR / f"rpa_{datetime.now():%Y%m%d}.log"

# [FIX-15] thêm scope Drive để upload CV (credentials đã bật Drive API)
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ── Master Sheet config ───────────────────────────────────────────────────────
MASTER_SHEET_ID = "1sCYbBWFU0ENhZrmlRMqZMomS--jaZT5YbPb-mu2ORno"
MASTER_TAB      = "マスター管理"

COL_COMPANY      = 1
COL_ID           = 2
COL_PW           = 3
COL_MEDIA        = 4
COL_SHEET_URL    = 6
COL_TAB_NAME     = 7
COL_MEDIA_URL    = 8
COL_DRIVE_FOLDER = 10   # [FIX-15] cột K = link folder Drive chứa CV của công ty

ENTRIES_URL = "https://ats.rct.airwork.net/entries"

ROW_SEL   = 'tr[data-la="entries_detail_transition_click"]'
MODAL_SEL = 'div.styles_container__BMWEr[role="dialog"]'

# ── Retry config ──────────────────────────────────────────────────────────────
MAX_MODAL_RETRIES   = 3   # số lần thử lại khi modal không lấy được tên
MODAL_WAIT_MS       = 150 # ms chờ tối thiểu sau click (modal wait thực tế ở scrape_modal)
MODAL_EXTRA_WAIT_MS = 400 # ms chờ thêm nếu retry

# ── Logging ───────────────────────────────────────────────────────────────────
LOG_DIR.mkdir(exist_ok=True)
RESUME_DIR.mkdir(exist_ok=True)   # [FIX-15]
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

def parse_birthday(text: str) -> str:
    m = re.search(r"(\d+)年(\d+)月(\d+)日生まれ", text)
    return f"{m.group(1)}/{int(m.group(2)):02d}/{int(m.group(3)):02d}" if m else ""

def parse_age(text: str) -> str:
    m = re.search(r"(\d+)歳", text)
    return m.group(1) if m else ""

def parse_gender(text: str) -> str:
    m = re.search(r"[／/]([男女性]+)", text.strip())
    return m.group(1) if m else ""

def parse_kana(text: str) -> str:
    return text.strip().strip("（）()").strip()

def extract_sheet_id(url: str) -> str:
    if "/d/" in url:
        return url.split("/d/")[1].split("/")[0]
    return url.strip()


def extract_drive_folder_id(url: str) -> str:
    """
    [FIX-15] Lấy folder_id từ link Drive ở cột K, ví dụ:
      https://drive.google.com/drive/folders/1LFUHzs8y-3VtbCYOCuuepG9gVmPuTmI8?usp=...
      → 1LFUHzs8y-3VtbCYOCuuepG9gVmPuTmI8
    Nếu người dùng dán thẳng folder_id thì trả về nguyên.
    """
    url = (url or "").strip()
    if not url:
        return ""
    if "/folders/" in url:
        return url.split("/folders/")[1].split("?")[0].split("/")[0].strip()
    if "id=" in url:
        return url.split("id=")[1].split("&")[0].strip()
    return url


def _safe_filename(name: str) -> str:
    """[FIX-15] Bỏ ký tự không hợp lệ cho tên file (giữ chữ Nhật)."""
    name = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", name or "").strip()
    return name[:120] if name else "cv"


def _norm_id(v) -> str:
    """
    [FIX-14] Chuẩn hoá apply_id để so trùng: CHỉ giữ chữ số.
    Tránh lệch khi:
      - Sheet tự format số có dấu phẩy ngăn cách → "29,760,116"
      - Đọc UNFORMATTED_VALUE trả về kiểu int/float thay vì str
    Cả lúc đọc từ Sheet lẫn lúc so sánh đều đi qua hàm này → đảm bảo
    so trùng đồng nhất, không bao giờ ghi trùng vì khác định dạng.
    """
    return re.sub(r"\D", "", str(v))


# ══════════════════════════════════════════════════════════════════════════════
# Google Sheets
# ══════════════════════════════════════════════════════════════════════════════

def get_sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        str(CRED_FILE), scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)


def get_drive_service():
    """[FIX-15] Service Drive để upload CV. Dùng chung credentials với Sheets."""
    creds = service_account.Credentials.from_service_account_file(
        str(CRED_FILE), scopes=SCOPES
    )
    return build("drive", "v3", credentials=creds)


def read_master_accounts(service) -> list[dict]:
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SHEET_ID,
            range=f"{MASTER_TAB}!A:K",   # [FIX-15] mở rộng tới cột K (folder Drive)
        ).execute()
        rows = result.get("values", [])
    except Exception as e:
        log.error(f"❌ Không đọc được Master Sheet: {e}")
        sys.exit(1)

    accounts = []
    for i, row in enumerate(rows[1:], start=2):
        while len(row) <= COL_DRIVE_FOLDER:   # [FIX-15] đảm bảo có đủ cột K
            row.append("")

        media = row[COL_MEDIA].strip()
        if media != "AW":
            continue

        company      = row[COL_COMPANY].strip()
        id_val       = row[COL_ID].strip()
        pw_val       = row[COL_PW].strip()
        sheet_url    = row[COL_SHEET_URL].strip()
        tab_name     = row[COL_TAB_NAME].strip()
        media_url    = row[COL_MEDIA_URL].strip()
        drive_folder = row[COL_DRIVE_FOLDER].strip()   # [FIX-15] cột K

        if not all([id_val, pw_val, media_url, sheet_url]):
            log.warning(f"  ⚠️  Dòng {i} [{company}] thiếu thông tin — bỏ qua")
            continue

        accounts.append({
            "company":      company,
            "id":           id_val,
            "pw":           pw_val,
            "sheet_url":    sheet_url,
            "tab_name":     tab_name or "【AW】応募者リスト",
            "media_url":    media_url,
            "drive_folder": drive_folder,   # [FIX-15]
        })
        if not drive_folder:
            log.warning(f"  ⚠️  Dòng {i} [{company}] thiếu folder Drive (cột K) — sẽ không upload CV")
        log.info(f"  ✓ Dòng {i}: [{company}] {id_val}")

    return accounts


def get_existing_ids(service, sheet_id: str, tab: str) -> set:
    """
    [FIX-2]  Dùng apply_id (cột N = hidden key) làm dedup key thay vì name+date.
    [FIX-14] Đọc UNFORMATTED_VALUE + chuẩn hoá _norm_id để tránh ghi trùng
             do Sheet format số có dấu phẩy (29,760,116 ≠ 29760116).
    """
    existing = set()
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"{tab}!N3:N",
            valueRenderOption="UNFORMATTED_VALUE",  # [FIX-14] không lấy số đã format
        ).execute()
        for row in result.get("values", []):
            if not row:
                continue
            nid = _norm_id(row[0])
            if nid:
                existing.add(nid)
    except Exception as e:
        log.warning(f"  ⚠️  Không đọc được cột ID Sheet: {e}")
    return existing


def get_next_empty_row(service, sheet_id: str, tab: str) -> int:
    """Tìm dòng trống tiếp theo bằng cách đọc cột B (氏名)."""
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


def append_one_row(service, sheet_id: str, tab: str, applicant: dict):
    """
    Ghi 1 ứng viên vào dòng trống cuối, bắt đầu từ cột A.
    Layout cột A–N (xem bên dưới).
    """
    next_row = get_next_empty_row(service, sheet_id, tab)

    row_data = [[
        applicant.get("apply_date", ""),       # A
        applicant.get("name", ""),              # B
        applicant.get("kana", ""),              # C
        applicant.get("gender", ""),            # D
        applicant.get("birthday", ""),          # E
        applicant.get("age", ""),               # F
        applicant.get("email", ""),             # G
        applicant.get("tel", ""),               # H
        applicant.get("address", ""),           # I
        applicant.get("job_name", ""),          # J
        applicant.get("employment", ""),        # K
        applicant.get("work_location", ""),     # L
        applicant.get("status", ""),            # M
        applicant.get("apply_id", ""),          # N ← hidden dedup key
    ]]

    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{tab}!A{next_row}",
        valueInputOption="USER_ENTERED",
        body={"values": row_data},
    ).execute()

    log.info(f"    📝 Ghi vào dòng {next_row}: {applicant.get('name', '')}")


def _to_row(applicant: dict) -> list:
    return [
        applicant.get("apply_date", ""),   # A
        applicant.get("name", ""),          # B
        applicant.get("kana", ""),          # C
        applicant.get("gender", ""),        # D
        applicant.get("birthday", ""),      # E
        applicant.get("age", ""),           # F
        applicant.get("email", ""),         # G
        applicant.get("tel", ""),           # H
        applicant.get("address", ""),       # I
        applicant.get("job_name", ""),      # J
        applicant.get("employment", ""),    # K
        applicant.get("work_location", ""), # L
        applicant.get("status", ""),        # M
        applicant.get("apply_id", ""),      # N ← hidden dedup key
        applicant.get("resume_url", ""),    # O ← [FIX-15] link CV trên Drive (レジュメ)
    ]


def append_batch(service, sheet_id: str, tab: str, applicants: list[dict]) -> int:
    """
    Ghi toàn bộ danh sách ứng viên trong 1 lần gọi API duy nhất.
    Trả về số dòng đã ghi thành công.
    """
    if not applicants:
        return 0

    next_row = get_next_empty_row(service, sheet_id, tab)
    values   = [_to_row(a) for a in applicants]

    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{tab}!A{next_row}",
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()

    log.info(f"    📝 Ghi batch {len(values)} dòng từ dòng {next_row}")
    return len(values)


# ══════════════════════════════════════════════════════════════════════════════
# Google Drive — upload CV  [FIX-15]
# ══════════════════════════════════════════════════════════════════════════════

def upload_resume_to_drive(drive_service, folder_id: str, local_path: str,
                           display_name: str) -> str:
    """
    [FIX-15] Upload 1 file PDF CV lên folder Drive của công ty, trả về link xem.

    LƯU Ý quota: service account KHÔNG có dung lượng Drive riêng. Nếu folder
    (cột K) nằm trong "My Drive" của 1 user, upload có thể lỗi
    "storageQuotaExceeded". Cách chắc ăn: để folder trong 1 Shared Drive
    (ドライブの共有) và share quyền Editor cho email service account. Hàm này
    đã bật supportsAllDrives=True để hỗ trợ Shared Drive.
    """
    if not folder_id:
        return ""
    if not local_path or not os.path.exists(local_path):
        return ""
    try:
        fname = _safe_filename(display_name) or os.path.basename(local_path)
        if not fname.lower().endswith(".pdf"):
            fname += ".pdf"
        metadata = {"name": fname, "parents": [folder_id]}
        media = MediaFileUpload(local_path, mimetype="application/pdf", resumable=False)
        file = drive_service.files().create(
            body=metadata,
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True,   # hỗ trợ Shared Drive
        ).execute()
        link = file.get("webViewLink") or f"https://drive.google.com/file/d/{file.get('id')}/view"
        log.info(f"    ☁️  Upload CV: {fname} → {link}")
        # Xoá file tạm sau khi upload thành công để tránh phình ổ đĩa
        try:
            os.remove(local_path)
        except Exception:
            pass
        return link
    except Exception as e:
        log.warning(f"    ⚠️  Upload CV Drive thất bại ({local_path}): {e}")
        return ""


# ══════════════════════════════════════════════════════════════════════════════
# Login
# ══════════════════════════════════════════════════════════════════════════════

async def login_airwork(page, url: str, username: str, password: str) -> bool:
    try:
        log.info(f"  → Đăng nhập...")
        await page.goto(url, wait_until="domcontentloaded", timeout=30_000)
        await page.wait_for_timeout(2000)

        for sel in ['input[name="username"]', 'input[name="email"]',
                    'input[type="email"]', 'input[id="username"]']:
            if await page.query_selector(sel):
                await page.fill(sel, username)
                log.info(f"  → Username: {sel}")
                break

        for sel in ['input[name="password"]', 'input[type="password"]', 'input[id="password"]']:
            if await page.query_selector(sel):
                await page.fill(sel, password)
                log.info(f"  → Password: {sel}")
                break

        for sel in ['input[type="submit"]', 'button[type="submit"]',
                    'button:has-text("ログイン")', 'button:has-text("サインイン")']:
            try:
                await page.click(sel, timeout=3_000)
                log.info(f"  → Submit: {sel}")
                break
            except Exception:
                continue

        try:
            await page.wait_for_url(
                lambda url: "airwork.net" in url and "login" not in url,
                timeout=30_000,
            )
        except PlaywrightTimeout:
            pass

        await page.wait_for_timeout(2000)
        log.info(f"  → URL sau login: {page.url}")

        if "airwork.net" not in page.url or "login" in page.url:
            log.warning("  ⚠️  Đăng nhập thất bại")
            return False

        log.info("  ✅ Đăng nhập thành công")
        return True

    except PlaywrightTimeout:
        log.error("  ❌ Timeout khi đăng nhập")
        return False
    except Exception as e:
        log.error(f"  ❌ Lỗi: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
# Entries list helpers
# ══════════════════════════════════════════════════════════════════════════════

async def goto_entries_page(page, page_num: int):
    url = ENTRIES_URL if page_num == 1 else f"{ENTRIES_URL}?page={page_num}"
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=30_000)
    except PlaywrightTimeout:
        await page.goto(url, wait_until="commit", timeout=30_000)
    # Chờ row thực sự xuất hiện thay vì fixed timeout
    try:
        await page.wait_for_selector(ROW_SEL, timeout=12_000)
    except PlaywrightTimeout:
        log.warning(f"  ⚠️  Không thấy row ứng viên ở trang {page_num}")


async def get_row_data(page) -> list[dict]:
    """Lấy apply_id + apply_date từ danh sách — TRƯỚC khi click vào modal.
    apply_date lấy từ span.styles_dateTime__MPfnz (thời gian thật, không bị
    lệch múi giờ như thời gian trong modal header).

    Tìm theo thứ tự:
    1. span.styles_dateTime__MPfnz trong row (tr)
    2. span.styles_dateTime__MPfnz trong parent element của row
    3. Regex trực tiếp trên inner_text của row (fallback chắc nhất)
    """
    rows = await page.query_selector_all(ROW_SEL)
    result = []
    for row in rows:
        apply_id = await row.get_attribute("data-la-apply")
        if not apply_id:
            continue
        apply_id = apply_id.strip()

        apply_date = ""

        def _parse_dt(raw: str) -> str:
            m = re.search(r"(\d{4})/(\d{1,2})/(\d{1,2})\s+(\d{1,2}:\d{2})", raw)
            if m:
                return f"{m.group(1)}/{int(m.group(2)):02d}/{int(m.group(3)):02d} {m.group(4)}"
            return ""

        try:
            # 1. Tìm trong row trực tiếp
            dt_el = await row.query_selector("span.styles_dateTime__MPfnz")
            if dt_el:
                apply_date = _parse_dt((await dt_el.inner_text()).strip())

            # 2. Nếu không thấy, tìm trong parent (row có thể là <tr> con của wrapper)
            if not apply_date:
                apply_date = await page.evaluate("""
                    (rowEl) => {
                        // Tìm lên parent tối đa 3 cấp
                        let node = rowEl.parentElement;
                        for (let i = 0; i < 3 && node; i++) {
                            const span = node.querySelector('span.styles_dateTime__MPfnz');
                            if (span) return span.innerText.trim();
                            node = node.parentElement;
                        }
                        return '';
                    }
                """, row)
                if apply_date:
                    apply_date = _parse_dt(apply_date) or apply_date

            # 3. Fallback: regex trên toàn bộ text của row
            if not apply_date:
                row_text = await row.inner_text()
                apply_date = _parse_dt(row_text)

        except Exception as e:
            log.debug(f"    get_row_data date error for {apply_id}: {e}")

        result.append({"apply_id": apply_id, "apply_date": apply_date})
        log.debug(f"    row {apply_id}: apply_date={apply_date!r}")

    log.info(f"  📋 {len(result)} ứng viên trong trang")
    return result


# keep backward compat — used in iter_new_applicants
async def get_row_ids(page) -> list[str]:
    data = await get_row_data(page)
    return [d["apply_id"] for d in data]


async def get_total_pages(page) -> int:
    for sel in ['[class*="pagination"]', '[class*="Pagination"]',
                'nav[aria-label*="ページ"]', '[class*="pager"]']:
        el = await page.query_selector(sel)
        if el:
            text = await el.inner_text()
            nums = re.findall(r"\d+", text)
            if nums:
                return max(int(n) for n in nums)
    for sel in ['a[aria-label="次のページ"]', 'button:has-text("次へ")',
                'a:has-text("次へ")', 'li[class*="next"] a']:
        el = await page.query_selector(sel)
        if el:
            disabled      = await el.get_attribute("disabled")
            aria_disabled = await el.get_attribute("aria-disabled")
            if disabled is None and aria_disabled != "true":
                return 999
    return 1


# ══════════════════════════════════════════════════════════════════════════════
# Modal scraping
# ══════════════════════════════════════════════════════════════════════════════

async def close_modal(page):
    for sel in [
        'button[class*="closeBtn"]',
        'button[class*="close"]',
        'button[aria-label="閉じる"]',
        'button:has-text("閉じる")',
    ]:
        el = await page.query_selector(sel)
        if el:
            try:
                await el.click(timeout=2_000)
                # Chờ modal thực sự biến mất thay vì fixed 500ms
                try:
                    await page.wait_for_selector(MODAL_SEL, state="hidden", timeout=3_000)
                except PlaywrightTimeout:
                    pass
                return
            except Exception:
                continue
    await page.keyboard.press("Escape")
    try:
        await page.wait_for_selector(MODAL_SEL, state="hidden", timeout=3_000)
    except PlaywrightTimeout:
        await page.wait_for_timeout(200)


async def download_resume(page, apply_id: str, name: str) -> str:
    """
    [FIX-15] Tải PDF レジュメ của ứng viên về local KHI MODAL ĐANG MỞ.
    Gọi ngay sau khi scrape xong, TRƯỚC close_modal.

    Luồng theo UI AirWork:
      1. Click tab 応募情報 (đảm bảo đang ở tab có khối レジュメ)
         <button data-la="overlay_entry_detail_tab_application_click">応募情報</button>
      2. Lấy link tải trực tiếp từ thẻ <a download href=".../preview_file_pdf/...">
         rồi fetch bằng page.request (chia sẻ cookie đăng nhập) → nhanh, ổn định.
      3. Fallback: click nút download (img download.svg) và intercept Playwright
         download nếu không tìm thấy thẻ <a download>.

    Trả về đường dẫn file local đã tải, hoặc "" nếu ứng viên không có CV.
    """
    # 1. Đảm bảo đang ở tab 応募情報
    try:
        tab = await page.query_selector(
            'button[data-la="overlay_entry_detail_tab_application_click"]'
        )
        if tab:
            await tab.click()
            await page.wait_for_timeout(400)
    except Exception:
        pass

    safe_name = _safe_filename(name) or apply_id

    RESUME_LINK_SEL = (
        'a[download][href*="preview_file_pdf"], '
        'a[href*="preview_file_pdf"], '
        'a[class*="toolbarLink"][download]'
    )

    # [FIX-17] Khối レジュメ load chậm (nhất là khi có PDF) → CHỜ link tải xuất
    # hiện trước, tránh kết luận "không có CV" oan khi panel chưa render xong.
    try:
        await page.wait_for_selector(RESUME_LINK_SEL, state="attached", timeout=8_000)
    except PlaywrightTimeout:
        # Chờ network ổn định rồi thử lại 1 lần (panel có thể vẫn đang tải)
        try:
            await page.wait_for_load_state("networkidle", timeout=4_000)
        except PlaywrightTimeout:
            pass
        try:
            await page.wait_for_selector(RESUME_LINK_SEL, state="attached", timeout=4_000)
        except PlaywrightTimeout:
            pass

    # 2. Thử lấy href trực tiếp từ thẻ <a download> rồi fetch
    try:
        a_el = await page.query_selector(RESUME_LINK_SEL)
        if a_el:
            href = await a_el.get_attribute("href")
            if href:
                full_url = urljoin(page.url, href)
                resp = await page.request.get(full_url, timeout=30_000)
                if resp.ok:
                    body = await resp.body()
                    if body and body[:4] == b"%PDF":   # đúng là file PDF
                        local_path = str(RESUME_DIR / f"{apply_id}_{safe_name}.pdf")
                        with open(local_path, "wb") as f:
                            f.write(body)
                        log.info(f"    📄 Tải CV: {os.path.basename(local_path)}")
                        return local_path
                    else:
                        log.debug(f"    CV fetch không phải PDF (ID={apply_id})")
                else:
                    log.debug(f"    CV fetch HTTP {resp.status} (ID={apply_id})")
    except Exception as e:
        log.debug(f"    Tải CV trực tiếp lỗi (ID={apply_id}): {e}")

    # 3. Fallback: click nút download và intercept
    try:
        async with page.expect_download(timeout=10_000) as dl_info:
            clicked = False
            for sel in [
                'a[download]',
                'img[src*="download"]',
                'button[data-la*="resume"]',
                '[aria-label*="ダウンロード"]',
                'a:has-text("ダウンロード")',
            ]:
                el = await page.query_selector(sel)
                if el:
                    await el.click()
                    clicked = True
                    break
            if not clicked:
                log.info(f"    ℹ️  Ứng viên ID={apply_id} không có CV để tải")
                return ""
        download = await dl_info.value
        local_path = str(RESUME_DIR / f"{apply_id}_{safe_name}.pdf")
        await download.save_as(local_path)
        log.info(f"    📄 Tải CV (fallback): {os.path.basename(local_path)}")
        return local_path
    except PlaywrightTimeout:
        log.info(f"    ℹ️  Ứng viên ID={apply_id} không có CV để tải")
        return ""
    except Exception as e:
        log.warning(f"    ⚠️  Không tải được CV (ID={apply_id}): {e}")
        return ""


async def _dump_modal_debug(page, apply_id: str, attempt: int):
    """
    [FIX-6] Dump HTML + screenshot khi modal không lấy được tên.
    Giúp debug selector thay đổi mà không cần chạy lại.
    """
    try:
        slug = f"modal_FAIL_{apply_id}_try{attempt}"
        shot_path = LOG_DIR / f"{slug}.png"
        html_path = LOG_DIR / f"{slug}.html"
        await page.screenshot(path=str(shot_path), full_page=False)
        html = await page.content()
        html_path.write_text(html, encoding="utf-8")
        log.warning(f"    🔍 DEBUG dump: {slug}.png + .html")
    except Exception as e:
        log.warning(f"    ⚠️  Không dump được debug: {e}")


async def _get_icon_text(scope, data_type: str) -> str:
    """
    Lấy text gắn với icon data-type (human/mail/call/address).

    [FIX-13] AirWork đổi layout + đổi hash class CSS-module. Element MỚI:

      性別/生年月日/年齢 (human) — CÓ <div> bọc ở giữa:
        <span class="styles_icon__nkRPX" data-type="human"></span>
        <div class="styles_humanInfoColumn__fUcV_">
            <span class="styles_infoText__Sznk8">63歳（1963年2月17日生まれ）／男性</span>
        </div>
        <span class="styles_infoText__Sznk8">63歳（1963年2月17日生まれ）／男性</span>

      メール/電話/住所 (mail/call/address) — span liền kề:
        <span class="styles_icon__nkRPX" data-type="mail"></span>
        <span class="styles_infoText__Sznk8">...@indeedemail.com</span>

    Hai điểm vỡ so với code cũ:
      1) hash class đổi: styles_infoText__UN0fw → styles_infoText__Sznk8
         → KHÔNG hard-code hash nữa, dùng [class*="infoText"]
      2) 'human' có <div> chen giữa → toán tử '+' (adjacent sibling) KHÔNG
         khớp → dùng '~' (general sibling, bắt cả sibling liền kề lẫn cách xa)
         và thêm fallback '~ div span' cho trường hợp text nằm trong <div>.
    """
    selectors = [
        # General sibling: bắt span infoText đứng sau icon — đúng cho cả layout
        # cũ (span liền kề) và mới (human có div chen giữa, vẫn còn span sibling)
        f'span[data-type="{data_type}"] ~ span[class*="infoText"]',
        # Fallback: infoText nằm trong <div> sibling của icon (layout human mới)
        f'span[data-type="{data_type}"] ~ div span[class*="infoText"]',
        # Fallback layout cũ: adjacent sibling trực tiếp
        f'span[data-type="{data_type}"] + span[class*="infoText"]',
    ]
    for sel in selectors:
        try:
            el = await scope.query_selector(sel)
            if el:
                txt = (await el.inner_text()).strip()
                if txt:
                    return txt
        except Exception:
            continue
    return ""


async def scrape_modal(page, apply_id: str) -> dict:
    """
    [FIX-6]  Đọc panel chi tiết với wait ổn định + nhiều selector dự phòng cho 氏名.
    [FIX-16] KHÔNG còn chờ cứng MODAL_SEL ('div.styles_container__BMWEr[role=dialog]').
             AirWork đổi hash CSS-module (styles_container__BMWEr đổi mỗi lần
             deploy) → selector này không khớp → báo "Modal không xuất hiện" GIẢ
             dù panel chi tiết thật ra ĐÃ mở. Sửa: chờ TRỰC TIẾP tên ứng viên
             (h1/title) xuất hiện, không gate qua container hash. Nếu MODAL_SEL
             còn khớp thì dùng làm scope; không thì fallback dùng cả page.
    Nếu tên vẫn không lấy được → trả về {} để caller có thể retry.
    """
    data = {"apply_id": apply_id}

    NAME_SELECTORS = [
        "h1.styles_title__Gs8Yk",
        "h1[class*='title']",
        "h1[class*='name']",
        "h1[class*='applicant']",
        "[class*='applicantName']",
        "[class*='candidateName']",
        "[class*='userName']",
        "h1",
    ]

    # [FIX-16] Chờ BẤT KỲ selector tên nào xuất hiện (panel đã mở) — không gate
    # qua container hash. Mỗi selector thử tối đa 1.5s, tổng tối đa ~12s.
    name_appeared = False
    for sel in NAME_SELECTORS:
        try:
            await page.wait_for_selector(sel, state="visible", timeout=1_500)
            name_appeared = True
            break
        except PlaywrightTimeout:
            continue
    if not name_appeared:
        try:
            await page.wait_for_selector("h1", state="visible", timeout=6_500)
        except PlaywrightTimeout:
            log.warning("    ⚠️  Panel chi tiết không xuất hiện (timeout)")
            return {}

    # [FIX-16] Dùng MODAL_SEL làm scope nếu còn khớp; nếu hash đổi → fallback page
    modal = await page.query_selector(MODAL_SEL)
    if not modal:
        modal = page

    async def mtext(sel: str) -> str:
        try:
            el = await modal.query_selector(sel)
            if el:
                return (await el.inner_text()).strip()
        except Exception:
            pass
        try:
            el = await page.query_selector(sel)
            if el:
                return (await el.inner_text()).strip()
        except Exception:
            pass
        return ""

    async def mtext_modal_only(sel: str) -> str:
        """Chỉ query trong modal — không fallback ra page.
        Dùng cho 応募日時 để tránh lấy nhầm phần tử ngoài modal."""
        try:
            el = await modal.query_selector(sel)
            if el:
                return (await el.inner_text()).strip()
        except Exception:
            pass
        return ""

    # ── 氏名 ──────────────────────────────────────────────────────────────
    # [FIX-17] Panel load chậm (đặc biệt khi có PDF レジュメ) → ban đầu tên hiển
    # thị là placeholder "応募者". Bản cũ CHẤP NHẬN placeholder làm tên thật →
    # vừa ghi sai tên, vừa hụt CV (vì link tải cũng chưa load).
    # Sửa: KHÔNG bao giờ chấp nhận placeholder. Nếu gặp placeholder → chờ
    # network ổn định + thử lại; hết lượt vẫn placeholder → trả {} để
    # click_and_scrape() click lại từ đầu (panel có thêm thời gian load).
    NAME_SELECTORS = [
        "h1.styles_title__Gs8Yk",
        "h1[class*='title']",
        "h1[class*='name']",
        "h1[class*='applicant']",
        "[class*='applicantName']",
        "[class*='candidateName']",
        "[class*='userName']",
        "h1",
    ]
    PLACEHOLDER_NAMES = {"応募者", "応募者様", "候補者"}

    def _valid_name(v: str) -> bool:
        return bool(v) and v not in PLACEHOLDER_NAMES and 2 <= len(v) <= 30 \
            and "http" not in v and "\n" not in v

    async def _read_name_retry(sel: str) -> str:
        for _ in range(3):
            val = await mtext(sel)
            if _valid_name(val):
                return val
            if val in PLACEHOLDER_NAMES:
                # Panel đang load (vd có resume PDF) → chờ network ổn định rồi thử lại
                try:
                    await page.wait_for_load_state("networkidle", timeout=4_000)
                except PlaywrightTimeout:
                    pass
                await page.wait_for_timeout(300)
                continue
            return ""   # giá trị khác (rỗng/không hợp lệ) → thử selector kế
        return ""       # hết 3 lần vẫn placeholder → THẤT BẠI, không chấp nhận

    data["name"] = ""
    for sel in NAME_SELECTORS:
        val = await _read_name_retry(sel)
        if val:
            data["name"] = val
            log.debug(f"    氏名 via [{sel}]: {val}")
            break

    if not data["name"]:
        try:
            for h1 in await modal.query_selector_all("h1, h2"):
                t = (await h1.inner_text()).strip()
                if _valid_name(t):
                    data["name"] = t
                    break
        except Exception:
            pass

    if not data["name"]:
        # Bao gồm cả trường hợp vẫn còn placeholder "応募者" sau retry
        # → trả {} để click_and_scrape() tự click lại từ đầu (fresh attempt)
        return {}

    # ── カナ ──────────────────────────────────────────────────────────────
    data["kana"] = ""
    for sel in ["h3.styles_kana__VRBIm", "h3[class*='kana']", "p[class*='kana']",
                "[class*='furigana']", "[class*='reading']"]:
        val = await mtext(sel)
        if val:
            data["kana"] = parse_kana(val)
            break

    # ── 応募日時 ──────────────────────────────────────────────────────────
    # Lấy từ danh sách (span.styles_dateTime__MPfnz trong row) TRƯỚC khi click
    # → thời gian đúng, không bị lệch múi giờ. Xem get_row_data() + caller.
    # scrape_modal KHÔNG lấy apply_date nữa; caller sẽ ghi đè sau khi return.
    data["apply_date"] = ""

    # ── 年齢・生年月日・性別 ───────────────────────────────────────────────
    # <span data-type="human"></span>
    # <span class="styles_infoText__UN0fw">43歳（1982年8月19日生まれ）／男性</span>
    # E = 生年月日 YYYY/MM/DD,  F = 年齢 (数字のみ),  D = 性別
    bio_text = await _get_icon_text(modal, "human")
    data["gender"]   = parse_gender(bio_text)
    data["birthday"] = parse_birthday(bio_text)
    data["age"]      = parse_age(bio_text)

    # ── Email ─────────────────────────────────────────────────────────────
    # <span data-type="mail"></span>
    # <span class="styles_infoText__UN0fw">xxx@gmail.com</span>
    data["email"] = await _get_icon_text(modal, "mail")

    # ── 電話番号 ──────────────────────────────────────────────────────────
    # <span data-type="call"></span>
    # <span class="styles_infoText__UN0fw">08056915963</span>
    data["tel"] = await _get_icon_text(modal, "call")

    # ── 住所 ─────────────────────────────────────────────────────────────
    # <span data-type="address"></span>
    # <span class="styles_infoText__UN0fw">宮城県 東松島市...</span>
    # ⚠️ Chỉ lấy từ icon address — không fallback heuristic để tránh ghi nhầm số điện thoại
    data["address"] = await _get_icon_text(modal, "address")

    # ── 応募求人名 ────────────────────────────────────────────────────────
    # <a class="styles_jobTitleLink__YzGop ...">フィットネスジムの受付・ご案内スタッフ</a>
    data["job_name"] = ""
    for sel in [
        'a.styles_jobTitleLink__YzGop',
        'a[data-la="entry_detail_job_offer_preview_link_click"]',
        '[class*="jobTitle"]',
        'dt:has-text("応募求人") + dd',
    ]:
        val = await mtext(sel)
        if val:
            # Strip prefix kiểu [11263686] mà AirWork đôi khi thêm vào
            val = re.sub(r"^\[\d+\]\s*", "", val).strip()
            data["job_name"] = val
            break

    # ── 雇用形態 (col K) ──────────────────────────────────────────────────
    # <span class="styles_jobLabel__6SC9n">アルバイト・パート</span>
    data["employment"] = ""
    for sel in [
        'span.styles_jobLabel__6SC9n',
        '[class*="jobLabel"]',
        'dt:has-text("雇用形態") + dd',
    ]:
        val = await mtext(sel)
        if val:
            data["employment"] = val.strip()
            break

    # ── 応募先 (col L) ────────────────────────────────────────────────────
    # <span class="styles_jobCaption__hJsiu">アッティーボジム高見プラザ店</span>
    data["work_location"] = ""
    for sel in [
        'span.styles_jobCaption__hJsiu',
        '[class*="jobCaption"]',
        'dt:has-text("勤務地") + dd',
    ]:
        val = await mtext(sel)
        if val:
            data["work_location"] = val.strip()
            break

    # ── 対応ステータス ────────────────────────────────────────────────────
    data["status"] = ""
    try:
        status_el = await page.query_selector(
            'select[aria-label="selectionInfoStatusBox"], select[class*="status"]'
        )
        if status_el:
            selected_val = await status_el.evaluate("el => el.value")
            option_el    = await status_el.query_selector(f'option[value="{selected_val}"]')
            if option_el:
                data["status"] = (await option_el.inner_text()).strip()
    except Exception:
        pass

    return data


# ══════════════════════════════════════════════════════════════════════════════
# Click + scrape 1 ứng viên — có retry
# ══════════════════════════════════════════════════════════════════════════════

async def click_and_scrape(page, apply_id: str, index: int) -> dict | None:
    """
    Retry tối đa MAX_MODAL_RETRIES lần.
    - Chỉ chờ tối thiểu sau click (scrape_modal tự chờ modal visible)
    - Tăng dần thời gian chờ khi retry
    """
    for attempt in range(1, MAX_MODAL_RETRIES + 1):
        try:
            # Re-query row mỗi lần thử để tránh stale element
            row = await page.query_selector(f'{ROW_SEL}[data-la-apply="{apply_id}"]')
            if not row:
                log.warning(f"    [{index}] Attempt {attempt}: row không còn trong DOM — reload trang")
                await page.reload(wait_until="domcontentloaded", timeout=20_000)
                try:
                    await page.wait_for_selector(ROW_SEL, timeout=8_000)
                except PlaywrightTimeout:
                    pass
                row = await page.query_selector(f'{ROW_SEL}[data-la-apply="{apply_id}"]')
                if not row:
                    log.warning(f"    [{index}] Row vẫn không có sau reload — bỏ qua")
                    return None

            await row.click()
            # Chờ tối thiểu để click register, scrape_modal sẽ wait_for_selector modal
            pre_wait = MODAL_WAIT_MS + (attempt - 1) * MODAL_EXTRA_WAIT_MS
            await page.wait_for_timeout(pre_wait)

            detail = await scrape_modal(page, apply_id)

            if detail.get("name"):
                if attempt > 1:
                    log.info(f"    [{index}] ✅ Lấy được sau {attempt} lần thử: {detail['name']}")
                return detail

            log.warning(f"    [{index}] Attempt {attempt}/{MAX_MODAL_RETRIES}: tên rỗng — ID={apply_id}")
            await _dump_modal_debug(page, apply_id, attempt)
            await close_modal(page)
            await page.wait_for_timeout(300)

        except Exception as e:
            log.warning(f"    [{index}] Attempt {attempt}/{MAX_MODAL_RETRIES}: exception — {e}")
            try:
                await close_modal(page)
            except Exception:
                pass
            await page.wait_for_timeout(300)

    log.error(f"    [{index}] ❌ Hết retry, bỏ qua ID={apply_id} — xem debug dump trong logs/")
    return None


# ══════════════════════════════════════════════════════════════════════════════
# Async generator: yield từng ứng viên mới
# ══════════════════════════════════════════════════════════════════════════════

async def iter_new_applicants(page, page_num: int, existing_ids: set, full_scan: bool):
    """
    [FIX-3] Quét HẾT tất cả row, không dừng sớm.
    [FIX-7] Re-fetch row list sau khi đóng mỗi modal để tránh stale DOM.
    """
    if page_num == 1:
        await page.screenshot(path=str(LOG_DIR / "debug_entries.png"))
        log.info(f"  📸 debug_entries.png — URL: {page.url}")

    # Lấy danh sách apply_id + apply_date từ danh sách (trước khi click modal)
    row_data = await get_row_data(page)

    if not row_data:
        log.warning(f"  ⚠️  Không thấy row ứng viên ở trang {page_num}")
        html = await page.content()
        (LOG_DIR / f"debug_empty_p{page_num}.html").write_text(html, encoding="utf-8")
        return

    for i, rd in enumerate(row_data):
        apply_id   = rd["apply_id"]
        apply_date = rd["apply_date"]
        if _norm_id(apply_id) in existing_ids:   # [FIX-14] so trùng đã chuẩn hoá
            log.info(f"    [{i+1}] ID={apply_id} đã có — bỏ qua")
            continue

        detail = await click_and_scrape(page, apply_id, i + 1)

        if detail is None:
            # Hết retry, bỏ qua nhưng KHÔNG break — tiếp tục ứng viên sau
            continue

        # Ghi đè apply_date bằng thời gian từ danh sách (tránh lệch +2h từ modal)
        if apply_date:
            detail["apply_date"] = apply_date

        log.info(f"    [{i+1}] ✓ {detail['name']}  {detail.get('apply_date', '')}")
        yield detail

        await close_modal(page)
        await page.wait_for_timeout(300)


# ══════════════════════════════════════════════════════════════════════════════
# Process 1 account
# ══════════════════════════════════════════════════════════════════════════════

async def process_account(browser, account: dict, sheets_service, drive_service, full_scan: bool):
    company   = account["company"]
    username  = account["id"]
    password  = account["pw"]
    url       = account["media_url"]
    sheet_url = account["sheet_url"]
    tab_name  = account["tab_name"]

    log.info(f"\n{'='*60}")
    log.info(f"🏢 {company}  ({'全件スキャン' if full_scan else '新着チェック'})")
    log.info(f"   ID : {username}")

    sheet_id        = extract_sheet_id(sheet_url)
    drive_folder_id = extract_drive_folder_id(account.get("drive_folder", ""))   # [FIX-15]
    existing_ids    = get_existing_ids(sheets_service, sheet_id, tab_name)
    log.info(f"  📊 Sheet đã có {len(existing_ids)} ứng viên (theo apply_id)")
    if drive_folder_id:
        log.info(f"  📁 Folder Drive CV: {drive_folder_id}")
    else:
        log.warning("  ⚠️  Không có folder Drive (cột K) — sẽ KHÔNG tải/ghi CV cho công ty này")

    context = await browser.new_context(
        viewport={"width": 1440, "height": 900},
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
    )
    page = await context.new_page()

    try:
        ok = await login_airwork(page, url, username, password)
        if not ok:
            return

        # ── Thu thập tất cả ứng viên mới ─────────────────────────────────
        # AirWork hiển thị mới nhất trước.
        # Dò từng trang cho đến khi gặp apply_id đã có trong sheet → dừng.
        # Nếu full_scan → dò hết tất cả trang, bỏ qua ID đã có.
        all_new: list[dict] = []
        run_seen: set = set()   # [FIX-14] chống trùng trong CÙNG 1 lần chạy
        page_num   = 1
        stop_scan  = False

        while not stop_scan:
            log.info(f"  📄 Trang {page_num}...")
            await goto_entries_page(page, page_num)

            if page_num == 1:
                await page.screenshot(path=str(LOG_DIR / "debug_entries.png"))

            row_data = await get_row_data(page)

            if not row_data:
                log.warning(f"  ⚠️  Không thấy row ứng viên ở trang {page_num} — dừng")
                html = await page.content()
                (LOG_DIR / f"debug_empty_p{page_num}.html").write_text(html, encoding="utf-8")
                break

            for i, rd in enumerate(row_data):
                apply_id   = rd["apply_id"]
                apply_date = rd["apply_date"]
                nid        = _norm_id(apply_id)   # [FIX-14] chuẩn hoá để so trùng

                if nid in existing_ids:
                    if full_scan:
                        log.info(f"    [{i+1}] ID={apply_id} đã có — bỏ qua (full scan)")
                        continue
                    else:
                        # Gặp ID cũ → đã đến vùng đã ghi → dừng hoàn toàn
                        log.info(f"    [{i+1}] Gặp ID đã có ({apply_id}) — dừng quét")
                        stop_scan = True
                        break

                # [FIX-14] Chống trùng trong cùng 1 lần chạy (ID xuất hiện 2 trang
                # do danh sách dịch chuyển khi có entry mới giữa chừng)
                if nid in run_seen:
                    log.info(f"    [{i+1}] ID={apply_id} đã xử lý trong lần chạy này — bỏ qua")
                    continue

                detail = await click_and_scrape(page, apply_id, i + 1)
                if detail is None:
                    continue

                # Ghi đè apply_date bằng thời gian lấy từ danh sách
                # (thời gian trong modal header bị lệch múi giờ +2h)
                if apply_date:
                    detail["apply_date"] = apply_date

                run_seen.add(nid)   # [FIX-14] đánh dấu đã xử lý trong lần chạy này

                # [FIX-15] Tải CV khi modal ĐANG MỞ (trước close_modal)
                detail["resume_local"] = ""
                if drive_folder_id:
                    detail["resume_local"] = await download_resume(
                        page, apply_id, detail.get("name", "")
                    )

                log.info(f"    [{i+1}] ✓ {detail['name']}  {detail.get('apply_date', '')}")
                all_new.append(detail)

                await close_modal(page)
                # close_modal đã chờ modal hidden — không cần wait thêm

            if stop_scan:
                break

            # Kiểm tra còn trang tiếp theo không
            total_pages = await get_total_pages(page)
            if page_num >= total_pages or total_pages == 1:
                log.info(f"  ✅ Đã quét hết {page_num} trang")
                break

            page_num += 1

        # ── Sắp xếp: cũ nhất → mới nhất trước khi ghi ───────────────────
        def _sort_key(a: dict):
            ds = a.get("apply_date", "")
            for fmt in ("%Y/%m/%d %H:%M", "%Y/%m/%d"):
                try:
                    return datetime.strptime(ds, fmt)
                except ValueError:
                    continue
            return datetime.min

        all_new.sort(key=_sort_key)
        log.info(f"  🆕 {len(all_new)} ứng viên mới — ghi theo thứ tự cũ → mới")

        # ── [FIX-15] Upload CV lên Drive + gắn link vào cột O ────────────
        if drive_folder_id and all_new:
            n_up = 0
            for a in all_new:
                lp = a.get("resume_local", "")
                if lp:
                    disp = f"{a.get('name','')}_{a.get('apply_id','')}"
                    a["resume_url"] = upload_resume_to_drive(
                        drive_service, drive_folder_id, lp, disp
                    )
                    if a["resume_url"]:
                        n_up += 1
            log.info(f"  ☁️  Upload {n_up}/{len(all_new)} CV lên Drive")

        # ── Ghi batch 1 lần duy nhất (thay vì N lần API call) ────────────
        total_written = 0
        if all_new:
            try:
                total_written = append_batch(sheets_service, sheet_id, tab_name, all_new)
                for a in all_new:
                    existing_ids.add(_norm_id(a["apply_id"]))   # [FIX-14] chuẩn hoá
                    log.info(f"    ✅ {a['name']}  {a.get('apply_date', '')}")
            except Exception as e:
                log.error(f"    ❌ Ghi batch thất bại: {e}")

        log.info(f"  🆕 Tổng đã ghi: {total_written} ứng viên")

        shot = LOG_DIR / f"{company}_{datetime.now():%H%M%S}.png"
        await page.screenshot(path=str(shot))
        log.info(f"  📸 {shot.name}")

    finally:
        await context.close()


# ══════════════════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════════════════

async def main(headless: bool = True, full_scan: bool = False):
    log.info(f"\n{'#'*60}")
    log.info(f"# AirWork RPA  —  {datetime.now():%Y-%m-%d %H:%M:%S}")
    log.info(f"# Mode: {'headless' if headless else 'visible'} | {'【全件スキャン】' if full_scan else '【新着チェック】'}")
    log.info(f"{'#'*60}")

    sheets_service = get_sheets_service()
    drive_service  = get_drive_service()   # [FIX-15]

    log.info(f"\n📋 Đọc danh sách AW từ Sheet マスター管理...")
    accounts = read_master_accounts(sheets_service)
    log.info(f"✅ Tìm thấy {len(accounts)} tài khoản AW")

    if not accounts:
        log.warning("Không có tài khoản AW nào!")
        return

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        for account in accounts:
            await process_account(browser, account, sheets_service, drive_service, full_scan)
        await browser.close()

    log.info("\n✅ Hoàn thành tất cả!")


if __name__ == "__main__":
    headless  = "--visible" not in sys.argv
    full_scan = "--full"    in sys.argv
    asyncio.run(main(headless=headless, full_scan=full_scan))
