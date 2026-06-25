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
  [FIX-13] Cột D–I (性別/生年月日/年齢/メール/電話/住所) bị rỗng:
           AirWork đổi layout + đổi hash class CSS-module:
             - hash class infoText đổi (styles_infoText__UN0fw → __Sznk8) →
               KHÔNG hard-code hash nữa, dùng [class*="infoText"]
             - icon "human" giờ có thêm <div class="styles_humanInfoColumn__..">
               bọc giữa → toán tử '+' (adjacent) không khớp → phải dùng
               '~' (general sibling) hoặc '~ div span'
  [FIX-14] Ghi TRÙNG apply_id (cột N): Sheet lưu cột N dạng số, đọc lại bằng
           FORMATTED_VALUE có thể trả "29,760,116" (có dấu phẩy) trong khi
           apply_id scrape là "29760116" → dedup fail → ghi trùng. Sửa:
             - Đọc cột N bằng UNFORMATTED_VALUE
             - Chuẩn hoá apply_id qua _norm_id() (chỉ giữ chữ số) ở cả 2 phía
             - Chống trùng ngay trong CÙNG 1 lần chạy (run_seen)
"""

import asyncio
import logging
import re
import sys
from datetime import datetime
from pathlib import Path

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR  = Path(__file__).parent
CRED_FILE = BASE_DIR / "credentials" / "credentials.json"
LOG_DIR   = BASE_DIR / "logs"
LOG_FILE  = LOG_DIR / f"rpa_{datetime.now():%Y%m%d}.log"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ── Master Sheet config ───────────────────────────────────────────────────────
MASTER_SHEET_ID = "1sCYbBWFU0ENhZrmlRMqZMomS--jaZT5YbPb-mu2ORno"
MASTER_TAB      = "マスター管理"

COL_COMPANY   = 1
COL_ID        = 2
COL_PW        = 3
COL_MEDIA     = 4
COL_SHEET_URL = 6
COL_TAB_NAME  = 7
COL_MEDIA_URL = 8

ENTRIES_URL = "https://ats.rct.airwork.net/entries"

ROW_SEL   = 'tr[data-la="entries_detail_transition_click"]'
MODAL_SEL = 'div.styles_container__BMWEr[role="dialog"]'

# ── Retry config ──────────────────────────────────────────────────────────────
MAX_MODAL_RETRIES   = 3   # số lần thử lại khi modal không lấy được tên
MODAL_WAIT_MS       = 150 # ms chờ tối thiểu sau click (modal wait thực tế ở scrape_modal)
MODAL_EXTRA_WAIT_MS = 400 # ms chờ thêm nếu retry

# ── Logging ───────────────────────────────────────────────────────────────────
LOG_DIR.mkdir(exist_ok=True)
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


def _norm_id(v) -> str:
    """
    [FIX-14] Chuẩn hoá apply_id để so trùng: CHỈ giữ chữ số.
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


def read_master_accounts(service) -> list[dict]:
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
        while len(row) <= COL_MEDIA_URL:
            row.append("")

        media = row[COL_MEDIA].strip()
        if media != "AW":
            continue

        company   = row[COL_COMPANY].strip()
        id_val    = row[COL_ID].strip()
        pw_val    = row[COL_PW].strip()
        sheet_url = row[COL_SHEET_URL].strip()
        tab_name  = row[COL_TAB_NAME].strip()
        media_url = row[COL_MEDIA_URL].strip()

        if not all([id_val, pw_val, media_url, sheet_url]):
            log.warning(f"  ⚠️  Dòng {i} [{company}] thiếu thông tin — bỏ qua")
            continue

        accounts.append({
            "company":   company,
            "id":        id_val,
            "pw":        pw_val,
            "sheet_url": sheet_url,
            "tab_name":  tab_name or "【AW】応募者リスト",
            "media_url": media_url,
        })
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
    [FIX-10] AirWork đổi UI -> class CSS-module hash (styles_xxx__HASH) đổi theo
    mỗi lần deploy, nên KHÔNG còn gate cứng theo MODAL_SEL (hash dễ vỡ).
    Thay vào đó: chờ trực tiếp h1 tên ứng viên xuất hiện (đã có fallback
    [class*='title'] / h1 chung) -> không phụ thuộc hash CSS module nào cả.
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

    # Chờ BẤT KỲ selector tên nào xuất hiện, tối đa ~8s -> không gate qua container
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
            log.warning("    ⚠️  Không thấy tên ứng viên xuất hiện (timeout)")
            return {}

    # Cố lấy modal container theo selector cũ -> dùng nếu còn khớp,
    # nếu không khớp (hash đổi) thì fallback dùng page làm scope luôn.
    modal = await page.query_selector(MODAL_SEL)
    if not modal:
        modal = page  # fallback: query trực tiếp trên page

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

    # ── 氏名 ──────────────────────────────────────────────────────────────
    # [FIX-12] FIX-11 từng "chấp nhận" placeholder "応募者" làm tên thật khi
    # retry hết mà vẫn vậy -> SAI: thực tế nó luôn là race condition (load
    # chậm vì có resume PDF đính kèm), không phải tên thật.
    # Sửa: KHÔNG bao giờ chấp nhận placeholder. Nếu vẫn là placeholder sau
    # khi chờ -> trả về {} để click_and_scrape() retry lại bằng cách
    # CLICK LẠI TỪ ĐẦU (fresh DOM, có thời gian load nhiều hơn 1 lần thử
    # đơn lẻ có thể cho được). Dùng wait_for_load_state("networkidle") thay
    # cho sleep cố định -> tự thích nghi theo tốc độ load thật của panel.
    PLACEHOLDER_NAMES = {"応募者", "応募者様"}

    async def _read_name_retry(sel: str) -> str:
        for attempt in range(3):
            val = await mtext(sel)
            if val and val not in PLACEHOLDER_NAMES and 2 <= len(val) <= 30 \
               and "http" not in val and "\n" not in val:
                return val
            if val in PLACEHOLDER_NAMES:
                # Panel có thể đang load nặng hơn (vd có resume PDF) -> chờ
                # network ổn định thay vì sleep cố định cứng nhắc
                try:
                    await page.wait_for_load_state("networkidle", timeout=4_000)
                except PlaywrightTimeout:
                    pass
                await page.wait_for_timeout(300)
                continue
            return ""  # giá trị khác (rỗng / không hợp lệ) -> thử selector kế
        return ""  # hết 3 lần vẫn là placeholder -> THẤT BẠI, không chấp nhận

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
                if t and t not in PLACEHOLDER_NAMES and 2 <= len(t) <= 30 and "http" not in t:
                    data["name"] = t
                    break
        except Exception:
            pass

    if not data["name"]:
        # Bao gồm cả trường hợp vẫn còn placeholder "応募者" sau retry
        # -> trả {} để click_and_scrape() tự click lại từ đầu (fresh attempt)
        return {}

    # ── カナ ──────────────────────────────────────────────────────────────
    data["kana"] = ""
    for sel in ["h3.styles_kana__VRBIm", "h3[class*='kana']", "p[class*='kana']",
                "[class*='furigana']", "[class*='reading']"]:
        val = await mtext(sel)
        if val:
            data["kana"] = parse_kana(val)
            break

    data["apply_date"] = ""

    # ── 年齢・生年月日・性別・Email・Tel・住所 (dùng data-type, bền hơn class hash) ──
    # [FIX-13] _get_icon_text đã sửa để bắt được layout mới (div bọc + hash đổi)
    bio_text = await _get_icon_text(modal, "human")
    if not bio_text and modal is not page:
        bio_text = await _get_icon_text(page, "human")
    data["gender"]   = parse_gender(bio_text)    # D ← 男性
    data["birthday"] = parse_birthday(bio_text)  # E ← 1963/02/17
    data["age"]      = parse_age(bio_text)        # F ← 63

    data["email"]   = await _get_icon_text(modal, "mail") or await _get_icon_text(page, "mail")        # G
    data["tel"]     = await _get_icon_text(modal, "call") or await _get_icon_text(page, "call")        # H
    data["address"] = await _get_icon_text(modal, "address") or await _get_icon_text(page, "address")  # I

    # ── 応募求人名 ────────────────────────────────────────────────────────
    data["job_name"] = ""
    for sel in [
        'a.styles_jobTitleLink__YzGop',
        'a[data-la="entry_detail_job_offer_preview_link_click"]',
        '[class*="jobTitle"]',
        'dt:has-text("応募求人") + dd',
    ]:
        val = await mtext(sel)
        if val:
            val = re.sub(r"^\[\d+\]\s*", "", val).strip()
            data["job_name"] = val
            break

    # ── 雇用形態 ──────────────────────────────────────────────────────────
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

    # ── 応募先 ────────────────────────────────────────────────────────────
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

async def process_account(browser, account: dict, sheets_service, full_scan: bool):
    company   = account["company"]
    username  = account["id"]
    password  = account["pw"]
    url       = account["media_url"]
    sheet_url = account["sheet_url"]
    tab_name  = account["tab_name"]

    log.info(f"\n{'='*60}")
    log.info(f"🏢 {company}  ({'全件スキャン' if full_scan else '新着チェック'})")
    log.info(f"   ID : {username}")

    sheet_id     = extract_sheet_id(sheet_url)
    existing_ids = get_existing_ids(sheets_service, sheet_id, tab_name)
    log.info(f"  📊 Sheet đã có {len(existing_ids)} ứng viên (theo apply_id)")

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

    log.info(f"\n📋 Đọc danh sách AW từ Sheet マスター管理...")
    accounts = read_master_accounts(sheets_service)
    log.info(f"✅ Tìm thấy {len(accounts)} tài khoản AW")

    if not accounts:
        log.warning("Không có tài khoản AW nào!")
        return

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        for account in accounts:
            await process_account(browser, account, sheets_service, full_scan)
        await browser.close()

    log.info("\n✅ Hoàn thành tất cả!")


if __name__ == "__main__":
    headless  = "--visible" not in sys.argv
    full_scan = "--full"    in sys.argv
    asyncio.run(main(headless=headless, full_scan=full_scan))
