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

Cấu trúc cột Sheet đích (A–L):
  A: 応募日時 (VD: 2026/05/12 13:33)
  B: 氏名
  C: (trống — reserved)
  D: 性別
  E: 生年月日 (VD: 2004/11/3)
  F: 年齢 (VD: 21歳)
  G: メールアドレス
  H: 電話番号
  I: 住所
  J: 応募求人名
  K: 対応ステータス
  L: 応募番号 (dedup key, VD: A2-7491-3914)

Cách chạy:
  python kbx_main.py              → chế độ bình thường (chỉ check mới)
  python kbx_main.py --full       → quét toàn bộ tất cả trang (lần đầu)
  python kbx_main.py --visible    → hiện trình duyệt để debug
  python kbx_main.py --full --visible
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

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR  = Path(__file__).parent
CRED_FILE = BASE_DIR / "credentials" / "credentials.json"
LOG_DIR   = BASE_DIR / "logs"
LOG_FILE  = LOG_DIR / f"kbx_{datetime.now():%Y%m%d}.log"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ── Master Sheet config (cùng file với AW) ────────────────────────────────────
MASTER_SHEET_ID = "1sCYbBWFU0ENhZrmlRMqZMomS--jaZT5YbPb-mu2ORno"
MASTER_TAB      = "マスター管理"

# Chỉ số cột (0-based, khớp với A=0, B=1, ...)
COL_COMPANY    = 0   # A
COL_ID         = 2   # C
COL_PW         = 3   # D
COL_MEDIA      = 4   # E
COL_CORP_CODE  = 5   # F — mã công ty KBX (VD: 8056-8676)
COL_SHEET_URL  = 6   # G
COL_TAB_NAME   = 7   # H
COL_BASE_URL   = 8   # I

# ── Retry / timing config ──────────────────────────────────────────────────────
MAX_DRAWER_RETRIES   = 3
DRAWER_WAIT_MS       = 1200
DRAWER_EXTRA_WAIT_MS = 800

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

def parse_birthday_kbx(text: str) -> str:
    """'2004年11月3日（21歳）' → '2004/11/3'"""
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", text)
    return f"{m.group(1)}/{m.group(2)}/{m.group(3)}" if m else ""

def parse_age_kbx(text: str) -> str:
    """'2004年11月3日（21歳）' or '(21歳)' → '21'"""
    m = re.search(r'[\(（](\d+)歳[\)）]', text)
    return m.group(1) if m else ""

def parse_apply_date_kbx(text: str) -> str:
    """
    '2026/05/12 13:33\nA2-7491-3914' → '2026/05/12 13:33'
    Chỉ lấy dòng đầu (ngày giờ), không đổi múi giờ.
    """
    line = text.strip().splitlines()[0].strip()
    # Kiểm tra định dạng YYYY/MM/DD HH:MM
    if re.match(r"\d{4}/\d{2}/\d{2} \d{2}:\d{2}", line):
        return line
    # Fallback: tìm pattern bất kỳ
    m = re.search(r"(\d{4}/\d{2}/\d{2} \d{2}:\d{2})", text)
    return m.group(1) if m else line

def extract_sheet_id(url: str) -> str:
    if "/d/" in url:
        return url.split("/d/")[1].split("/")[0]
    return url.strip()


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
        # Pad đến cột I (index 8)
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
        # Tách origin (scheme + netloc) để ghép path đúng
        # VD: https://secure.kyujinbox.com/login → https://secure.kyujinbox.com
        parsed    = urlparse(raw_url)
        base_url  = f"{parsed.scheme}://{parsed.netloc}"
        login_url = raw_url  # giữ link login đầy đủ để dùng riêng

        if not all([id_val, pw_val, corp_code, sheet_url, base_url]):
            log.warning(f"  ⚠️  Dòng {i} [{company}] thiếu thông tin — bỏ qua")
            continue

        accounts.append({
            "company":   company,
            "id":        id_val,
            "pw":        pw_val,
            "corp_code": corp_code,
            "sheet_url": sheet_url,
            "tab_name":  tab_name or "【KBX】応募者リスト",
            "base_url":  base_url,   # origin: https://secure.kyujinbox.com
            "login_url": login_url,  # full:   https://secure.kyujinbox.com/login
        })
        log.info(f"  ✓ Dòng {i}: [{company}] {id_val}  corp={corp_code}")

    return accounts


def get_existing_ids_kbx(service, sheet_id: str, tab: str) -> set:
    """Đọc cột N (応募番号) làm dedup key."""
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
    Layout A–N:
      A: apply_date   B: name      C: kana(カナ)  D: gender
      E: birthday     F: age       G: email        H: tel
      I: address      J: school    K: job_name     L: current_job
      M: status       N: apply_number (dedup)
    """
    next_row = get_next_empty_row_kbx(service, sheet_id, tab)

    row_data = [[
        applicant.get("apply_date", ""),    # A
        applicant.get("name", ""),          # B
        applicant.get("kana", ""),          # C — カナ
        applicant.get("gender", ""),        # D
        applicant.get("birthday", ""),      # E
        applicant.get("age", ""),           # F — số tuổi, không có 歳
        applicant.get("email", ""),         # G
        applicant.get("tel", ""),           # H
        applicant.get("address", ""),       # I
        applicant.get("school", ""),        # J ← 学校名
        applicant.get("job_name", ""),      # K
        applicant.get("current_job", ""),   # L ← 現在の職業
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


def _to_row_kbx(applicant: dict) -> list:
    return [
        applicant.get("apply_date", ""),    # A
        applicant.get("name", ""),          # B
        applicant.get("kana", ""),          # C — カナ
        applicant.get("gender", ""),        # D
        applicant.get("birthday", ""),      # E
        applicant.get("age", ""),           # F — số tuổi, không có 歳
        applicant.get("email", ""),         # G
        applicant.get("tel", ""),           # H
        applicant.get("address", ""),       # I
        applicant.get("school", ""),        # J ← 学校名
        applicant.get("job_name", ""),      # K
        applicant.get("current_job", ""),   # L ← 現在の職業
        applicant.get("status", ""),        # M
        applicant.get("apply_number", ""),  # N ← dedup key
    ]


def append_batch_kbx(service, sheet_id: str, tab: str, applicants: list[dict]) -> int:
    """Ghi toàn bộ danh sách ứng viên (đã sort) trong 1 lần gọi API."""
    if not applicants:
        return 0
    next_row = get_next_empty_row_kbx(service, sheet_id, tab)
    values   = [_to_row_kbx(a) for a in applicants]
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

async def login_kbx(page, base_url: str, username: str, password: str) -> str | None:
    """
    Đăng nhập KyujinBox.
    Trả về origin thực tế sau khi login (VD: https://saiyo.kyujinbox.com),
    hoặc None nếu thất bại.
    """
    login_url = base_url
    try:
        log.info(f"  → Truy cập: {login_url}")
        await page.goto(login_url, wait_until="domcontentloaded", timeout=30_000)
        await page.wait_for_timeout(2000)

        # Email
        email_sel = '#login_email, input[name="login[email]"], input[type="email"]'
        await page.fill(email_sel, username)
        log.info("  → Đã điền email")

        # Password
        pw_sel = '#login_password, input[name="login[password]"], input[type="password"]'
        await page.fill(pw_sel, password)
        log.info("  → Đã điền password")

        # Submit
        await page.click('#BtnLogin, button[type="submit"]', timeout=5_000)
        log.info("  → Đã click login")

        # Chờ chuyển trang
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

        # Lấy origin thực tế (có thể khác domain với login URL)
        parsed = urlparse(current)
        actual_origin = f"{parsed.scheme}://{parsed.netloc}"
        log.info(f"  ✅ Đăng nhập thành công — origin: {actual_origin}")
        return actual_origin

    except Exception as e:
        log.error(f"  ❌ Lỗi đăng nhập: {e}")
        return None


# ══════════════════════════════════════════════════════════════════════════════
# Chọn tab 直接投稿 và tìm công ty theo corp_code
# ══════════════════════════════════════════════════════════════════════════════

async def select_direct_tab(page) -> bool:
    """
    Chọn tab 直接投稿.
    HTML thực tế: <a href="/ptr/s-accounts" class="c-switch__item">直接投稿</a>
    Khi active:   <a href="/ptr/s-accounts" class="c-switch__item is-active">直接投稿</a>
    """
    try:
        # Selector chính xác theo href thực tế
        sel = 'a[href="/ptr/s-accounts"]'
        el = await page.query_selector(sel)
        if el:
            cls = await el.get_attribute("class") or ""
            if "is-active" not in cls:
                await el.click()
                await page.wait_for_load_state("networkidle", timeout=10_000)
                await page.wait_for_timeout(1000)
                log.info("  → Đã click tab 直接投稿")
            else:
                log.info("  → Tab 直接投稿 đã active")
            return True
        log.warning("  ⚠️  Không tìm thấy tab 直接投稿 (href=/ptr/s-accounts)")
        return False
    except Exception as e:
        log.error(f"  ❌ Lỗi chọn tab: {e}")
        return False


async def find_and_enter_company(page, base_url: str, corp_code: str) -> bool:
    """
    Tìm công ty khớp với corp_code trong danh sách tài khoản,
    sau đó click vào để đăng nhập vào trang công ty đó.

    Selector công ty trong bảng:
      <td class="c-table__data u-center">8056-8676</td>
    Link vào công ty:
      <a href="/ptr/saiyo_login/8056-8676" class="c-link">...</a>
    """
    accounts_url = f"{base_url}/ptr/l-accounts"
    try:
        await page.goto(accounts_url, wait_until="networkidle", timeout=30_000)
        await page.wait_for_timeout(1500)
    except PlaywrightTimeout:
        await page.goto(accounts_url, wait_until="domcontentloaded", timeout=30_000)

    # Chọn tab 直接投稿 trước
    await select_direct_tab(page)
    await page.wait_for_timeout(1000)

    log.info(f"  🔍 Tìm công ty có mã: {corp_code}")

    # Tìm tất cả ô có mã công ty — mã có thể nằm trong text kiểu "会社名（8056-8676）"
    # hoặc là ô riêng biệt chứa đúng mã
    try:
        # Thử 1: tìm td chứa đúng mã (exact match)
        tds = await page.query_selector_all("td")
        matched_td = None
        for td in tds:
            text = (await td.inner_text()).strip()
            if corp_code in text:
                matched_td = td
                log.info(f"  → Tìm thấy ô chứa mã: '{text}'")
                break

        if matched_td:
            # Tìm link trong cùng hàng tr
            row_el = await matched_td.evaluate_handle("el => el.closest('tr')")
            link = await row_el.query_selector(
                f'a[href*="/ptr/saiyo_login/{corp_code}"], '
                f'a[href*="/{corp_code}"], '
                f'a.c-link'
            )
            if link:
                href = await link.get_attribute("href")
                log.info(f"  → Truy cập công ty: {href}")
                target_url = f"{base_url}{href}" if href.startswith("/") else href
                await page.goto(target_url, wait_until="networkidle", timeout=30_000)
                await page.wait_for_timeout(2000)
                log.info(f"  ✅ Đã vào trang công ty: {page.url}")
                return True
            else:
                # Không có link riêng — thử click trực tiếp vào td hoặc row
                log.info("  → Không tìm thấy link, thử click vào row...")
                await matched_td.click()
                await page.wait_for_timeout(2000)
                log.info(f"  → URL sau click: {page.url}")
                return True

        log.warning(f"  ⚠️  Không tìm thấy công ty mã '{corp_code}' trong danh sách")
        await page.screenshot(path=str(LOG_DIR / f"kbx_no_corp_{corp_code}.png"))
        # Dump HTML để debug
        (LOG_DIR / f"kbx_no_corp_{corp_code}.html").write_text(
            await page.content(), encoding="utf-8"
        )
        return False

    except Exception as e:
        log.error(f"  ❌ Lỗi tìm công ty: {e}")
        return False


async def dismiss_message_popup(page):
    """
    Dismiss popup đồng ý メッセージ nếu xuất hiện.
    Nút: 同意せずにもどる (không đồng ý, quay lại) để không trigger side effects.
    """
    try:
        # Chờ tối đa 3s, nếu không có popup thì bỏ qua
        popup_sel = 'button:has-text("同意せずにもどる"), button:has-text("同意して始める")'
        el = await page.wait_for_selector(popup_sel, timeout=3_000)
        if el:
            # Click "同意せずにもどる" để đóng mà không đồng ý
            back_btn = await page.query_selector('button:has-text("同意せずにもどる")')
            if back_btn:
                await back_btn.click()
                log.info("  → Đã đóng popup メッセージ同意")
                await page.wait_for_timeout(800)
    except PlaywrightTimeout:
        pass  # Không có popup, tiếp tục bình thường
    except Exception as e:
        log.warning(f"  ⚠️  Lỗi dismiss popup: {e}")


async def goto_applicant_list(page, base_url: str, corp_code: str) -> bool:
    """
    Vào trang 応募者一覧.
    URL thực tế: /company/groups/G{corp_code}-0001/applications
    """
    try:
        group_id  = f"G{corp_code}-0001"
        direct_url = f"{base_url}/company/groups/{group_id}/applications"
        await page.goto(direct_url, wait_until="networkidle", timeout=30_000)
        await page.wait_for_timeout(1500)
        log.info(f"  → 応募者一覧: {page.url}")

        # Dismiss popup đồng ý メッセージ nếu có
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
    Trả về list[{"apply_number": str, "name_from_list": str, "element_index": int}]
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
    HTML thực tế KBX:
      có trang tiếp: <li class="c-pager__button c-pager__button--next is-active">
      hết trang:     <li class="c-pager__button c-pager__button--next">  (không có is-active)
    """
    try:
        el = await page.query_selector('li.c-pager__button--next')
        if el:
            cls = await el.get_attribute("class") or ""
            result = "is-active" in cls
            log.debug(f"    c-pager__button--next class='{cls}' → has_next={result}")
            return result
    except Exception as e:
        log.debug(f"    has_next_page error: {e}")
    return False


async def get_total_pages_kbx(page) -> int:
    """Lấy tổng số trang từ pagination (dùng làm log, không dùng để quyết định dừng)."""
    for sel in ['[class*="p-paging"]', '[class*="pagination"]', 'nav[class*="paging"]']:
        el = await page.query_selector(sel)
        if el:
            text = await el.inner_text()
            nums = re.findall(r"\d+", text)
            if nums:
                return max(int(n) for n in nums)
    return 1


async def goto_applicant_page(page, base_url: str, page_num: int, corp_code: str):
    """
    Chuyển trang trong 応募者一覧.
    URL thực tế: https://saiyo.kyujinbox.com/company/groups/G{corp_code}-0001/applications?slmsg=d&pg=N
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
# Drawer (chi tiết ứng viên)
# ══════════════════════════════════════════════════════════════════════════════

async def close_drawer(page):
    """
    Đóng panel chi tiết ứng viên KBX.
    Panel bên phải đóng khi click ra ngoài (vào vùng danh sách bên trái).
    """
    try:
        # Thử click vào tiêu đề 応募者一覧 để đóng panel
        title = await page.query_selector('h1.c-title:has-text("応募者一覧")')
        if title:
            await title.click()
            await page.wait_for_timeout(500)
            return
    except Exception:
        pass

    try:
        # Click vào vùng trống bên trái (cột メッセージ của row đầu tiên)
        first_row = await page.query_selector('tr.c-drawer__activeRow td:first-child')
        if first_row:
            # Click vào góc xa panel (x=100 tính từ viewport)
            box = await first_row.bounding_box()
            if box:
                await page.mouse.click(100, box["y"] + box["height"] / 2)
                await page.wait_for_timeout(500)
                return
    except Exception:
        pass

    # Fallback: Escape
    await page.keyboard.press("Escape")
    await page.wait_for_timeout(500)


async def _dump_drawer_debug(page, apply_number: str, attempt: int):
    """Dump screenshot + HTML khi không lấy được thông tin."""
    try:
        slug = f"kbx_FAIL_{apply_number}_try{attempt}"
        await page.screenshot(path=str(LOG_DIR / f"{slug}.png"), full_page=False)
        (LOG_DIR / f"{slug}.html").write_text(await page.content(), encoding="utf-8")
        log.warning(f"    🔍 DEBUG dump: {slug}.png + .html")
    except Exception as e:
        log.warning(f"    ⚠️  Không dump được: {e}")


async def scrape_drawer(page, apply_number: str) -> dict:
    """
    Đọc thông tin ứng viên từ panel bên phải của KBX.

    KBX dùng split-panel layout (không phải overlay drawer):
      - Panel phải luôn hiển thị, cập nhật khi click link bên trái
      - Phải chờ panel cập nhật đúng ứng viên trước khi đọc

    Tất cả dữ liệu nằm trong <p class="c-lines">:
      - apply_date: p.c-lines chứa pattern YYYY/MM/DD HH:MM
      - gender: "男性" hoặc "女性"
      - birthday+age: chứa "年" và "歳"
      - email: chứa "@"
      - tel: số điện thoại 10-11 số
      - address: chứa 県/都/道/府
      job_name: <a href="/jobs/edit/...">
      status: <select id="application_status">
    """
    data = {"apply_number": apply_number}

    # Dismiss popup メッセージ nếu xuất hiện khi mở drawer
    await dismiss_message_popup(page)

    # ── Chờ panel cập nhật đúng ứng viên này ────────────────────────────
    # KBX là split-panel: p.c-lines tồn tại ngay cả khi chưa click đúng người.
    # Cách đáng tin cậy nhất: chờ apply_number xuất hiện trong vùng detail.
    # apply_number thường nằm trong text của panel phải (VD: "A2-3916-0483").
    panel_ready = False
    for _wait_attempt in range(12):   # tối đa ~6 giây
        try:
            # Ưu tiên: tìm apply_number trong text vùng detail bên phải
            detail_text = ""
            for detail_sel in [
                '.s-application-detail',
                '.c-drawer__content',
                '[class*="detail"]',
                'main',
            ]:
                el = await page.query_selector(detail_sel)
                if el:
                    detail_text = await el.inner_text()
                    break
            if not detail_text:
                detail_text = await page.inner_text("body")

            if apply_number in detail_text:
                panel_ready = True
                break
        except Exception:
            pass
        await page.wait_for_timeout(500)

    if not panel_ready:
        # Fallback: chờ ít nhất p.c-lines xuất hiện
        try:
            await page.wait_for_selector("p.c-lines", state="visible", timeout=5_000)
            panel_ready = True
        except PlaywrightTimeout:
            log.warning("    ⚠️  Không thấy p.c-lines sau khi click")
            return {}

    await page.wait_for_timeout(300)

    # Lấy tất cả p.c-lines — ưu tiên từ vùng detail bên phải
    # (tránh lấy nhầm data từ vùng list bên trái)
    all_plines = []
    try:
        # Thử giới hạn trong panel detail trước
        detail_container = None
        for detail_sel in [
            '.s-application-detail',
            '.c-drawer__content',
            '[class*="detail"]',
        ]:
            detail_container = await page.query_selector(detail_sel)
            if detail_container:
                break

        if detail_container:
            els = await detail_container.query_selector_all("p.c-lines")
        else:
            els = await page.query_selector_all("p.c-lines")

        for el in els:
            t = (await el.inner_text()).strip()
            if t:
                all_plines.append(t)
    except Exception:
        pass

    log.debug(f"    p.c-lines ({len(all_plines)}): {all_plines[:8]}")

    # ── 応募日時 (cột A) ─────────────────────────────────────────────────
    data["apply_date"] = ""
    for t in all_plines:
        if re.search(r"\d{4}/\d{2}/\d{2} \d{2}:\d{2}", t):
            data["apply_date"] = parse_apply_date_kbx(t)
            break

    # ── 氏名 + カナ (cột B, C) ────────────────────────────────────────────
    # HTML thực tế KBX (xác nhận từ DevTools):
    #   氏名:  <h3 class="c-title c-title--type4 ...">氏名</h3>
    #          <p class="c-lines">吉川 忠浩</p>
    #   カナ:  <h1 class="c-title c-title--type1 ...">吉川 忠浩
    #              <span class="c-note u-m5--left">(よしかわ ただひろ)</span>
    #          </h1>
    # → Tên lấy từ p.c-lines sau label 氏名
    # → カナ lấy từ span.c-note có parent H1, bỏ dấu ngoặc ()（）
    data["name"] = ""
    data["kana"] = ""

    # ── Lấy カナ: query tất cả span.c-note, lọc cái có hiragana + parent H1 ──
    # Console xác nhận: class="c-note u-m5--left", parent H1.c-title--type1
    try:
        kana_spans = await page.query_selector_all("span.c-note")
        for span_el in kana_spans:
            txt = (await span_el.inner_text()).strip()
            # Phải có hiragana/katakana và parent là H1
            if not re.search(r'[ぁ-んァ-ン]', txt):
                continue
            parent_tag = await span_el.evaluate("el => el.parentElement.tagName")
            if parent_tag.upper() != "H1":
                continue
            # Bỏ dấu ngoặc () hoặc （）
            kana_clean = re.sub(r'^[\(（\s]+|[\)）\s]+$', '', txt).strip()
            if kana_clean:
                data["kana"] = kana_clean
                log.info(f"    カナ: {kana_clean}")
                break
    except Exception as e:
        log.debug(f"    カナ span.c-note method failed: {e}")

    # ── Lấy 氏名 từ h3.c-title--type4 label + p.c-lines ─────────────────
    # Cách 1: Tìm li.c-borderList__item chứa h3 text "氏名"
    try:
        items = await page.query_selector_all("li.c-borderList__item")
        for li in items:
            h = await li.query_selector("h3.c-title, h4.c-title, [class*='c-title']")
            if h and (await h.inner_text()).strip() == "氏名":
                p = await li.query_selector("p.c-lines")
                if p:
                    t = (await p.inner_text()).strip()
                    if t and 2 <= len(t) <= 20:
                        data["name"] = t
                        log.debug(f"    氏名 via li.c-borderList__item: {t}")
                        break
    except Exception as e:
        log.debug(f"    氏名 li method failed: {e}")

    # Cách 2: Tìm h3/h4 bất kỳ có text "氏名" rồi lấy sibling p.c-lines
    if not data["name"]:
        try:
            for sel in ['h3[class*="c-title"]', 'h4[class*="c-title"]', '[class*="c-title--type4"]']:
                labels = await page.query_selector_all(sel)
                for label_el in labels:
                    if (await label_el.inner_text()).strip() == "氏名":
                        sibling = await label_el.evaluate_handle(
                            """el => {
                                let next = el.nextElementSibling;
                                while (next) {
                                    if (next.matches('p.c-lines')) return next;
                                    next = next.nextElementSibling;
                                }
                                return null;
                            }"""
                        )
                        if sibling:
                            t = (await sibling.inner_text()).strip()
                            if t and 2 <= len(t) <= 20:
                                data["name"] = t
                                log.debug(f"    氏名 via h3+sibling: {t}")
                                break
                if data["name"]:
                    break
        except Exception as e:
            log.debug(f"    氏名 sibling method failed: {e}")

    # Cách 3: Fallback — lấy tên thuần từ h1 (bỏ span)
    if not data["name"]:
        try:
            h1_el = await page.query_selector('h1.c-title--type1')
            if h1_el:
                t = await h1_el.evaluate("""el => {
                    const clone = el.cloneNode(true);
                    clone.querySelectorAll('span').forEach(s => s.remove());
                    return clone.innerText.trim();
                }""")
                t = re.sub(r'\s+', ' ', t).strip()
                if t and 2 <= len(t) <= 20:
                    data["name"] = t
                    log.debug(f"    氏名 from h1 (no span): {t}")
        except Exception as e:
            log.debug(f"    氏名 h1 fallback failed: {e}")

    # Cách 4: Fallback cuối — tên từ link danh sách bên trái
    if not data["name"]:
        try:
            link = await page.query_selector(
                f'a.s-drawer_open[data-application-number="{apply_number}"]'
            )
            if link:
                t = (await link.inner_text()).strip()
                if t and 2 <= len(t) <= 20:
                    data["name"] = t
                    log.debug(f"    氏名 from list link: {t}")
        except Exception:
            pass

    if not data["name"]:
        log.warning(f"    ⚠️  Không lấy được tên cho #{apply_number}")
        return {}

    # ── Helper: lấy p.c-lines theo label trong li.c-borderList__item ────
    async def get_field_by_label(label_text: str) -> str:
        """Tìm li có h3 text = label_text, trả về nội dung p.c-lines. Rỗng nếu không có."""
        try:
            lis = await page.query_selector_all("li.c-borderList__item")
            for li in lis:
                h = await li.query_selector("h3.c-title, h4.c-title, [class*='c-title']")
                if h and (await h.inner_text()).strip() == label_text:
                    p = await li.query_selector("p.c-lines")
                    if p:
                        return (await p.inner_text()).strip()
        except Exception as e:
            log.debug(f"    get_field_by_label({label_text}) error: {e}")
        return ""

    # ── 性別 (cột D) ─────────────────────────────────────────────────────
    data["gender"] = ""
    raw = await get_field_by_label("性別")
    if raw in ("男性", "女性", "男", "女"):
        data["gender"] = raw
    else:
        # fallback từ all_plines
        for t in all_plines:
            if t in ("男性", "女性", "男", "女"):
                data["gender"] = t
                break

    # ── 生年月日 + 年齢 (cột E, F) ───────────────────────────────────────
    data["birthday"] = ""
    data["age"]      = ""
    raw = await get_field_by_label("生年月日")
    if raw and re.search(r"\d{4}年\d{1,2}月\d{1,2}日", raw):
        data["birthday"] = parse_birthday_kbx(raw)
        data["age"]      = parse_age_kbx(raw)
    else:
        for t in all_plines:
            if re.search(r"\d{4}年\d{1,2}月\d{1,2}日", t) and "歳" in t:
                data["birthday"] = parse_birthday_kbx(t)
                data["age"]      = parse_age_kbx(t)
                break

    # ── Email (cột G) ────────────────────────────────────────────────────
    data["email"] = ""
    raw = await get_field_by_label("メールアドレス")
    if raw and "@" in raw and " " not in raw and len(raw) < 100:
        data["email"] = raw
    else:
        for t in all_plines:
            if "@" in t and "." in t and " " not in t and len(t) < 100:
                data["email"] = t
                break

    # ── 電話番号 (cột H) ─────────────────────────────────────────────────
    data["tel"] = ""
    raw = await get_field_by_label("電話番号")
    if raw:
        clean = re.sub(r"[-\s　]", "", raw)
        if re.match(r"^0\d{9,10}$", clean):
            data["tel"] = raw
    if not data["tel"]:
        for t in all_plines:
            clean = re.sub(r"[-\s　]", "", t)
            if re.match(r"^0\d{9,10}$", clean):
                data["tel"] = t
                break

    # ── 住所 (cột I) ─────────────────────────────────────────────────────
    # HTML thực tế: <h3 class="c-title ...">住所</h3><p class="c-lines">東京都北区</p>
    # Dùng cùng pattern với 氏名 — tìm label rồi lấy p.c-lines kế tiếp.
    # KHÔNG dùng keyword 県/都/道/府 vì sẽ lấy nhầm nơi làm việc, trường học, v.v.
    data["address"] = ""
    try:
        items = await page.query_selector_all("li.c-borderList__item")
        for li in items:
            h3 = await li.query_selector("h3.c-title, h4.c-title, [class*='c-title']")
            if h3:
                label = (await h3.inner_text()).strip()
                if label == "住所":
                    p = await li.query_selector("p.c-lines")
                    if p:
                        t = (await p.inner_text()).strip()
                        if t:
                            data["address"] = t
                    break
    except Exception as e:
        log.debug(f"    住所 scrape error: {e}")

    # ── 学校名 (cột J) ──────────────────────────────────────────────────
    # <h3 class="c-title ...">学校名</h3>
    # <p class="c-lines">私立安田学園高等学校</p>
    data["school"] = await get_field_by_label("学校名")

    # ── 応募求人名 (cột K) ───────────────────────────────────────────────
    # <h3 class="c-title ...">応募求人</h3>
    # <p class="c-lines"><a href="/jobs/edit/...">求人名</a></p>
    data["job_name"] = ""
    try:
        items = await page.query_selector_all("li.c-borderList__item")
        for li in items:
            h3 = await li.query_selector("h3.c-title, h4.c-title, [class*='c-title']")
            if h3:
                label = (await h3.inner_text()).strip()
                if label in ("応募求人", "応募求人名"):
                    p = await li.query_selector("p.c-lines")
                    if p:
                        # Ưu tiên lấy text của link bên trong
                        job_link = await p.query_selector("a.c-link, a[href*='/jobs/edit/']")
                        if job_link:
                            t = (await job_link.inner_text()).strip()
                        else:
                            t = (await p.inner_text()).strip()
                        if t:
                            data["job_name"] = t
                    break
    except Exception as e:
        log.debug(f"    応募求人名 scrape error: {e}")

    # ── 現在の職業 (cột L) ───────────────────────────────────────────────
    # <h3 class="c-title ...">現在の職業</h3>
    # <p class="c-lines">アルバイト・パート</p>
    data["current_job"] = await get_field_by_label("現在の職業")

    # ── 対応ステータス (cột M) ───────────────────────────────────────────
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


# ══════════════════════════════════════════════════════════════════════════════
# Click + scrape 1 ứng viên — có retry (giống AW)
# ══════════════════════════════════════════════════════════════════════════════

async def click_and_scrape_kbx(page, apply_info: dict, index: int) -> dict | None:
    """
    Click vào link ứng viên, chờ panel bên phải cập nhật, đọc thông tin.
    Retry tối đa MAX_DRAWER_RETRIES lần.

    apply_info phải có: apply_number, name_from_list
    """
    apply_number    = apply_info["apply_number"]
    name_from_list  = apply_info.get("name_from_list", "").strip()

    for attempt in range(1, MAX_DRAWER_RETRIES + 1):
        try:
            # Re-query link mỗi lần (tránh stale element)
            link = await page.query_selector(
                f'a.s-drawer_open[data-application-number="{apply_number}"]'
            )
            if not link:
                log.warning(f"    [{index}] Attempt {attempt}: link không còn trong DOM")
                await page.reload(wait_until="networkidle", timeout=20_000)
                await page.wait_for_timeout(1500)
                link = await page.query_selector(
                    f'a.s-drawer_open[data-application-number="{apply_number}"]'
                )
                if not link:
                    log.warning(f"    [{index}] Link vẫn không có sau reload — bỏ qua")
                    return None

            await link.click()
            # Không dùng fixed wait — scrape_drawer tự chờ panel cập nhật đúng ứng viên
            await page.wait_for_timeout(300)

            detail = await scrape_drawer(page, apply_number)

            # Nếu tên rỗng nhưng panel đã load (có apply_number) → dùng tên từ list
            if not detail.get("name") and name_from_list and detail.get("apply_number"):
                log.info(f"    [{index}] 🔄 Dùng tên từ list: {name_from_list}")
                detail["name"] = name_from_list

            if detail.get("name"):
                if attempt > 1:
                    log.info(f"    [{index}] ✅ Lấy được sau {attempt} lần thử: {detail['name']}")
                return detail

            log.warning(f"    [{index}] Attempt {attempt}/{MAX_DRAWER_RETRIES}: tên rỗng — #{apply_number}")
            await _dump_drawer_debug(page, apply_number, attempt)
            await close_drawer(page)
            await page.wait_for_timeout(800)

        except Exception as e:
            log.warning(f"    [{index}] Attempt {attempt}/{MAX_DRAWER_RETRIES}: exception — {e}")
            try:
                await close_drawer(page)
            except Exception:
                pass
            await page.wait_for_timeout(600)

    # Last resort: nếu có tên từ list — ghi tối thiểu thay vì bỏ qua hoàn toàn
    if name_from_list:
        log.warning(f"    [{index}] ⚠️  Dùng tên từ list (fallback cuối): {name_from_list} #{apply_number}")
        return {"apply_number": apply_number, "name": name_from_list}

    log.error(f"    [{index}] ❌ Hết retry, bỏ qua #{apply_number}")
    return None


# ══════════════════════════════════════════════════════════════════════════════
# Async generator: yield từng ứng viên mới
# ══════════════════════════════════════════════════════════════════════════════

async def iter_new_applicants_kbx(page, page_num: int, existing_ids: set):
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

    for i, apply_info in enumerate(applicant_list):
        apply_number = apply_info["apply_number"]

        if apply_number in existing_ids:
            log.info(f"    [{i+1}] #{apply_number} đã có — bỏ qua")
            continue

        detail = await click_and_scrape_kbx(page, apply_info, i + 1)

        if detail is None:
            continue

        log.info(f"    [{i+1}] ✓ {detail['name']}  {detail.get('apply_date', '')}")
        yield detail

        await close_drawer(page)
        await page.wait_for_timeout(400)


# ══════════════════════════════════════════════════════════════════════════════
# Process 1 account
# ══════════════════════════════════════════════════════════════════════════════

async def process_account_kbx(browser, account: dict, sheets_service, full_scan: bool):
    company   = account["company"]
    username  = account["id"]
    password  = account["pw"]
    corp_code = account["corp_code"]
    base_url  = account["base_url"]   # origin: https://secure.kyujinbox.com
    login_url = account["login_url"]  # full:   https://secure.kyujinbox.com/login
    sheet_url = account["sheet_url"]
    tab_name  = account["tab_name"]

    log.info(f"\n{'='*60}")
    log.info(f"🏢 {company}  corp={corp_code}  ({'全件スキャン' if full_scan else '新着チェック'})")
    log.info(f"   ID : {username}")

    sheet_id     = extract_sheet_id(sheet_url)
    existing_ids = get_existing_ids_kbx(sheets_service, sheet_id, tab_name)
    log.info(f"  📊 Sheet đã có {len(existing_ids)} ứng viên")

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
        # 1. Đăng nhập — trả về origin thực tế sau redirect
        actual_origin = await login_kbx(page, login_url, username, password)
        if not actual_origin:
            return

        # 2. Vào trang danh sách tài khoản — dùng origin thực tế từ browser
        ok = await find_and_enter_company(page, actual_origin, corp_code)
        if not ok:
            return

        # 3. Vào 応募者一覧 — dùng origin thực tế
        ok = await goto_applicant_list(page, actual_origin, corp_code)
        if not ok:
            return

        # ── Thu thập tất cả ứng viên mới từ TẤT CẢ trang ───────────────
        # - Quét hết tất cả trang, không dừng sớm
        # - Bỏ qua ID đã có trong sheet (cột L)
        # - Sau khi quét xong → sort cũ→mới → ghi batch 1 lần
        all_new: list[dict] = []
        page_num = 1

        while True:
            log.info(f"  📄 Trang {page_num}...")

            applicant_list = await get_applicant_links(page)

            if not applicant_list:
                log.warning("  ⚠️  Trang trống — dừng")
                break

            async for applicant in iter_new_applicants_kbx(page, page_num, existing_ids):
                all_new.append(applicant)
                existing_ids.add(applicant["apply_number"])

            if not await has_next_page_kbx(page):
                log.info(f"  ✅ Đã quét hết {page_num} trang (không còn 次へ)")
                break

            page_num += 1
            log.info(f"  ➡️  Chuyển sang trang {page_num}...")
            await goto_applicant_page(page, actual_origin, page_num, corp_code)

        # ── Sắp xếp cũ → mới rồi ghi batch ─────────────────────────────
        def _sort_key(a: dict):
            ds = a.get("apply_date", "")
            try:
                return datetime.strptime(ds, "%Y/%m/%d %H:%M")
            except ValueError:
                return datetime.min

        all_new.sort(key=_sort_key)
        log.info(f"  🆕 {len(all_new)} ứng viên mới — ghi theo thứ tự cũ → mới")

        total_written = 0
        if all_new:
            try:
                total_written = append_batch_kbx(sheets_service, sheet_id, tab_name, all_new)
                for a in all_new:
                    log.info(f"    ✅ {a['name']}  {a.get('apply_date', '')}")
            except Exception as e:
                log.error(f"    ❌ Ghi batch thất bại: {e}")

        log.info(f"  🆕 Tổng đã ghi: {total_written} ứng viên")

        try:
            shot = LOG_DIR / f"kbx_{company}_{datetime.now():%H%M%S}.png"
            await page.screenshot(path=str(shot), timeout=5_000)
            log.info(f"  📸 {shot.name}")
        except Exception:
            pass

    finally:
        await context.close()


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
        browser = await p.chromium.launch(headless=headless)
        for account in accounts:
            await process_account_kbx(browser, account, sheets_service, full_scan)
        await browser.close()

    log.info("\n✅ Hoàn thành tất cả!")


if __name__ == "__main__":
    headless  = "--visible" not in sys.argv
    full_scan = "--full"    in sys.argv
    asyncio.run(main(headless=headless, full_scan=full_scan))
