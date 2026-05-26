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
MAX_MODAL_RETRIES   = 3    # số lần thử lại khi modal không lấy được tên
MODAL_WAIT_MS       = 1200 # ms chờ sau khi click row (tăng từ 800)
MODAL_EXTRA_WAIT_MS = 800  # ms chờ thêm nếu retry

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
    return f"{m.group(1)}/{m.group(2)}/{m.group(3)}" if m else ""

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
    [FIX-2] Dùng apply_id (cột N = hidden key) làm dedup key thay vì name+date.
    """
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
        await page.goto(url, wait_until="networkidle", timeout=30_000)
    except PlaywrightTimeout:
        await page.goto(url, wait_until="domcontentloaded", timeout=30_000)
    await page.wait_for_timeout(1500)
    try:
        await page.wait_for_selector(ROW_SEL, timeout=10_000)
    except PlaywrightTimeout:
        log.warning(f"  ⚠️  Không thấy row ứng viên ở trang {page_num}")


async def get_row_ids(page) -> list[str]:
    """Lấy danh sách apply_id từ tất cả row trong trang."""
    rows = await page.query_selector_all(ROW_SEL)
    result = []
    for row in rows:
        apply_id = await row.get_attribute("data-la-apply")
        if apply_id:
            result.append(apply_id.strip())
    log.info(f"  📋 {len(result)} ứng viên trong trang")
    return result


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
                await page.wait_for_timeout(500)
                return
            except Exception:
                continue
    await page.keyboard.press("Escape")
    await page.wait_for_timeout(500)


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


async def scrape_modal(page, apply_id: str) -> dict:
    """
    [FIX-6] Đọc modal với wait ổn định hơn + nhiều selector dự phòng cho 氏名.
    Nếu tên vẫn không lấy được → trả về {} để caller có thể retry.
    """
    data = {"apply_id": apply_id}

    # Chờ modal xuất hiện VÀ ổn định (không chỉ xuất hiện)
    try:
        await page.wait_for_selector(MODAL_SEL, state="visible", timeout=8_000)
    except PlaywrightTimeout:
        log.warning("    ⚠️  Modal không xuất hiện (timeout)")
        return {}

    # Thêm: chờ network idle sau khi modal mở để dữ liệu load xong
    try:
        await page.wait_for_load_state("networkidle", timeout=5_000)
    except PlaywrightTimeout:
        pass  # Không bắt buộc, tiếp tục

    modal = await page.query_selector(MODAL_SEL)
    if not modal:
        log.warning("    ⚠️  Không query được modal element")
        return {}

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

    async def mtext_all(sel: str) -> list[str]:
        results = []
        try:
            els = await modal.query_selector_all(sel)
            for el in els:
                t = (await el.inner_text()).strip()
                if t:
                    results.append(t)
        except Exception:
            pass
        return results

    # ── 氏名 ──────────────────────────────────────────────────────────────
    # [FIX-6] Thêm nhiều selector dự phòng + log selector nào thành công
    data["name"] = ""
    NAME_SELECTORS = [
        "h1.styles_title__Gs8Yk",      # selector gốc (CSS class cụ thể)
        "h1[class*='title']",
        "h1[class*='name']",
        "h1[class*='applicant']",
        "[class*='applicantName']",
        "[class*='candidateName']",
        "[class*='userName']",
        "h1",                           # fallback: bất kỳ h1 nào trong modal
    ]
    for sel in NAME_SELECTORS:
        val = await mtext(sel)
        # Tên người Nhật: 2–20 ký tự, không chứa URL hay ký tự đặc biệt
        if val and 2 <= len(val) <= 30 and "http" not in val and "\n" not in val:
            data["name"] = val
            log.debug(f"    氏名 via [{sel}]: {val}")
            break

    # Nếu vẫn không có → thử tìm h1 đầu tiên trong toàn modal (rộng hơn)
    if not data["name"]:
        try:
            all_h1 = await modal.query_selector_all("h1, h2")
            for h1 in all_h1:
                t = (await h1.inner_text()).strip()
                if t and 2 <= len(t) <= 30 and "http" not in t:
                    data["name"] = t
                    log.debug(f"    氏名 fallback h1/h2: {t}")
                    break
        except Exception:
            pass

    if not data["name"]:
        # Không lấy được tên — trả {} để caller retry
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
    # [FIX-8] Thêm nhiều selector + parse thêm định dạng ngày
    data["apply_date"] = ""
    APPLY_DATE_SELECTORS = [
        'div[data-type="応募日時"]',
        '[class*="applyDate"]',
        '[class*="apply_date"]',
        '[class*="applyAt"]',
        '[class*="appliedAt"]',
        'dt:has-text("応募日時") + dd',
        'th:has-text("応募日時") + td',
        'label:has-text("応募日時") + span',
        'td:has-text("応募日時") + td',
        # AirWork đôi khi dùng span/p thay vì dd
        'p[class*="date"]',
        'span[class*="date"]',
    ]
    for sel in APPLY_DATE_SELECTORS:
        val = await mtext(sel)
        if val and ("年" in val or "月" in val or re.search(r"\d{4}/\d{2}", val)):
            # "2026年5月13日(水) 15:45" → "2026/05/13 15:45"
            m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日[^0-9]*(\d{1,2}:\d{2})", val)
            if m:
                data["apply_date"] = (
                    f"{m.group(1)}/{int(m.group(2)):02d}/{int(m.group(3)):02d} {m.group(4)}"
                )
            else:
                data["apply_date"] = val.strip()
            log.debug(f"    応募日時 via [{sel}]: {data['apply_date']}")
            break

    # ── 年齢・生年月日・性別 ───────────────────────────────────────────────
    bio_text = ""
    try:
        candidates = await modal.query_selector_all("span, p, dd, td, div, h3")
        for el in candidates:
            t = (await el.inner_text()).strip()
            if "歳" in t and ("生まれ" in t or "男" in t or "女" in t) and len(t) < 80:
                bio_text = t
                break
    except Exception:
        pass

    data["gender"]   = parse_gender(bio_text)
    data["birthday"] = parse_birthday(bio_text)
    data["age"]      = parse_age(bio_text)

    # ── Email ─────────────────────────────────────────────────────────────
    data["email"] = ""
    try:
        candidates = await modal.query_selector_all("span, p, dd, td, a")
        for el in candidates:
            t = (await el.inner_text()).strip()
            if "@" in t and "." in t and len(t) < 100 and " " not in t and "\n" not in t:
                data["email"] = t
                break
    except Exception:
        pass

    # ── 電話番号 ──────────────────────────────────────────────────────────
    data["tel"] = ""
    try:
        candidates = await modal.query_selector_all("span, p, dd, td")
        for el in candidates:
            t = (await el.inner_text()).strip()
            clean = re.sub(r"[-\s　]", "", t)
            if re.match(r"^0\d{9,10}$", clean):
                data["tel"] = t
                break
    except Exception:
        pass

    # ── 住所 ─────────────────────────────────────────────────────────────
    data["address"] = ""
    for sel in [
        '[class*="address"]',
        'dt:has-text("住所") + dd',
        'th:has-text("住所") + td',
        'label:has-text("住所") + span',
    ]:
        val = await mtext(sel)
        if val and len(val) < 120:
            data["address"] = val
            break

    if not data["address"]:
        PREF_KEYWORDS = ["県", "都", "道", "府"]
        JOB_KEYWORDS  = ["求人", "スポーツ", "ジム", "フィットネス", "店舗", "マネジメント",
                         "スタッフ", "受付", "接客", "正社員", "アルバイト", "パート"]
        try:
            candidates = await modal.query_selector_all("span, p, dd, td, li")
            for el in candidates:
                t = (await el.inner_text()).strip()
                has_pref     = any(k in t for k in PREF_KEYWORDS)
                has_job_kw   = any(k in t for k in JOB_KEYWORDS)
                has_postcode = re.search(r"〒?\d{3}[-－]\d{4}", t)
                is_short     = 5 < len(t) < 120
                if has_pref and not has_job_kw and is_short:
                    data["address"] = t
                    break
                if has_postcode and is_short:
                    data["address"] = t
                    break
        except Exception:
            pass

    # ── 応募求人名 ────────────────────────────────────────────────────────
    data["job_name"] = ""
    try:
        p_el = await modal.query_selector("p:has(a.styles_linkDetail__qbl4P)")
        if not p_el:
            p_el = await modal.query_selector("p:has(a[class*='linkDetail'])")
        if p_el:
            full_text = (await p_el.inner_text()).strip()
            a_els = await p_el.query_selector_all("a")
            for a_el in a_els:
                a_text = (await a_el.inner_text()).strip()
                if a_text:
                    full_text = full_text.replace(a_text, "").strip()
            if full_text:
                data["job_name"] = full_text
    except Exception:
        pass

    if not data["job_name"]:
        for sel in [
            "[class*='jobTitle']",
            'dt:has-text("応募求人") + dd',
            'td:has-text("応募求人") + td',
            'dt:has-text("求人") + dd',
        ]:
            val = await mtext(sel)
            if val:
                data["job_name"] = val.replace("求人内容を確認する", "").strip()
                break

    # ── 雇用形態 ──────────────────────────────────────────────────────────
    data["employment"] = ""
    for sel in [
        'div[data-type="応募先（雇用形態）"]',
        'dt:has-text("雇用形態") + dd',
        'td:has-text("雇用形態") + td',
    ]:
        val = await mtext(sel)
        if val:
            data["employment"] = val
            break

    # ── 応募先（勤務地） ──────────────────────────────────────────────────
    data["work_location"] = ""
    for sel in [
        'div[data-type="応募先（勤務地）"]',
        'dt:has-text("勤務地") + dd',
        'td:has-text("勤務地") + td',
    ]:
        val = await mtext(sel)
        if val:
            data["work_location"] = val
            break

    # ── 対応ステータス ────────────────────────────────────────────────────
    data["status"] = ""
    try:
        status_el = await page.query_selector(
            'select[aria-label="selectionInfoStatusBox"], select[class*="status"]'
        )
        if status_el:
            selected_val = await status_el.evaluate("el => el.value")
            option_el    = await page.query_selector(f'select option[value="{selected_val}"]')
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
    [FIX-6] Retry tối đa MAX_MODAL_RETRIES lần.
    Mỗi lần thất bại: đóng modal (nếu có), chờ thêm, click lại.
    Trả None nếu hết retry.
    """
    for attempt in range(1, MAX_MODAL_RETRIES + 1):
        try:
            # [FIX-7] Re-query row mỗi lần thử để tránh stale element
            row = await page.query_selector(f'{ROW_SEL}[data-la-apply="{apply_id}"]')
            if not row:
                log.warning(f"    [{index}] Attempt {attempt}: row không còn trong DOM — reload trang")
                # Row biến mất → reload trang, thử lại
                await page.reload(wait_until="networkidle", timeout=20_000)
                await page.wait_for_timeout(1500)
                row = await page.query_selector(f'{ROW_SEL}[data-la-apply="{apply_id}"]')
                if not row:
                    log.warning(f"    [{index}] Row vẫn không có sau reload — bỏ qua")
                    return None

            await row.click()
            wait_ms = MODAL_WAIT_MS + (attempt - 1) * MODAL_EXTRA_WAIT_MS
            await page.wait_for_timeout(wait_ms)

            detail = await scrape_modal(page, apply_id)

            if detail.get("name"):
                if attempt > 1:
                    log.info(f"    [{index}] ✅ Lấy được sau {attempt} lần thử: {detail['name']}")
                return detail

            # Tên rỗng → dump debug rồi retry
            log.warning(f"    [{index}] Attempt {attempt}/{MAX_MODAL_RETRIES}: tên rỗng — ID={apply_id}")
            await _dump_modal_debug(page, apply_id, attempt)
            await close_modal(page)
            await page.wait_for_timeout(600)

        except Exception as e:
            log.warning(f"    [{index}] Attempt {attempt}/{MAX_MODAL_RETRIES}: exception — {e}")
            try:
                await close_modal(page)
            except Exception:
                pass
            await page.wait_for_timeout(600)

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

    # Lấy danh sách apply_id một lần — chỉ dùng để lặp, không dùng element
    apply_ids = await get_row_ids(page)

    if not apply_ids:
        log.warning(f"  ⚠️  Không thấy row ứng viên ở trang {page_num}")
        html = await page.content()
        (LOG_DIR / f"debug_empty_p{page_num}.html").write_text(html, encoding="utf-8")
        return

    for i, apply_id in enumerate(apply_ids):
        if apply_id in existing_ids:
            log.info(f"    [{i+1}] ID={apply_id} đã có — bỏ qua")
            continue

        detail = await click_and_scrape(page, apply_id, i + 1)

        if detail is None:
            # Hết retry, bỏ qua nhưng KHÔNG break — tiếp tục ứng viên sau
            continue

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

        page_num      = 1
        total_written = 0

        while True:
            log.info(f"  📄 Trang {page_num}...")
            await goto_entries_page(page, page_num)

            row_ids_on_page = await get_row_ids(page)
            total_on_page   = len(row_ids_on_page)

            async for applicant in iter_new_applicants(page, page_num, existing_ids, full_scan):
                try:
                    append_one_row(sheets_service, sheet_id, tab_name, applicant)
                    existing_ids.add(applicant["apply_id"])
                    total_written += 1
                    log.info(f"    ✅ Đã ghi: {applicant['name']}")
                except Exception as e:
                    log.error(f"    ❌ Ghi sheet thất bại: {e}")

            if total_on_page == 0 and page_num > 1:
                log.warning("  ⚠️  Trang trống thật — dừng")
                break

            if not full_scan:
                log.info("  — Chế độ thường: chỉ check trang 1")
                break

            total_pages = await get_total_pages(page)
            if page_num >= total_pages or total_pages == 1:
                log.info(f"  ✅ Đã quét hết {page_num} trang")
                break

            page_num += 1

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
