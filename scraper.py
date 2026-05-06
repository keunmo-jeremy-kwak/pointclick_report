"""
광고 플랫폼 어드민 리포트 자동 수집 → Google Sheets 저장
- 대상: https://admin.pointclick.co.kr/#/report/ads
- 데이터: 전날(yesterday) 광고 리포트
- 저장: Google Sheets 'Report' 시트에 행 추가
"""

import os
import sys
import time
from datetime import datetime, timedelta

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
import gspread
from google.oauth2.service_account import Credentials
import json

# ─── 환경 변수 ────────────────────────────────────────────────────────────────
ADMIN_URL      = "https://admin.pointclick.co.kr"
ADMIN_ID       = os.environ["ADMIN_ID"]
ADMIN_PW       = os.environ["ADMIN_PW"]
SPREADSHEET_ID = "1KeWHUVjAleAf0wz_bbr1y-MeWDDI05P06riC1nCoGzE"
SHEET_NAME     = "Report"
GCP_CREDS_JSON = os.environ["GCP_CREDENTIALS"]  # 서비스 계정 JSON 문자열

# 날짜 (전날, KST 기준)
from datetime import timezone
KST = timezone(timedelta(hours=9))
yesterday = (datetime.now(tz=KST) - timedelta(days=1)).strftime("%Y-%m-%d")
print(f"[INFO] 조회 날짜: {yesterday}")


# ─── Google Sheets 연결 ────────────────────────────────────────────────────────
def get_sheet():
    creds_dict = json.loads(GCP_CREDS_JSON)
    # gspread 최신 방식으로 인증
    gc = gspread.service_account_from_dict(creds_dict)
    sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    print(f"[INFO] 시트 연결 성공: {sheet.title} (행 수: {sheet.row_count})")
    return sheet


def append_rows(sheet, rows: list[list]):
    """시트에 행 추가. 헤더가 없으면 먼저 삽입. 전체 배치로 한 번에 저장."""
    existing = sheet.get_all_values()
    current_row_count = len(existing)
    print(f"[INFO] 현재 시트 데이터: {current_row_count}행")

    if not existing:
        header = ["날짜", "CD", "광고명", "OS", "광고 타입",
                  "광고 단가", "조회수", "클릭수", "전환수", "전환율", "광고비", "ChannelName"]
        sheet.append_row(header, value_input_option="USER_ENTERED")
        current_row_count = 1
        print("[INFO] 헤더 추가 완료")

    # 모든 행의 수식을 미리 계산 후 한 번에 저장 (Race Condition 방지)
    batch = []
    for i, row in enumerate(rows):
        next_row = current_row_count + 1 + i
        channel_formula = f"=iferror(vlookup(B{next_row},'채널정보'!$A$5:$D$1002,4,false),\"\")"
        batch.append(row + [channel_formula])
        print(f"[INFO] {i+1}번 행 준비 → 시트 {next_row}행: {row[:2]}")

    sheet.append_rows(batch, value_input_option="USER_ENTERED")
    print(f"[INFO] {len(rows)}행 일괄 저장 완료")


# ─── Playwright 자동화 ─────────────────────────────────────────────────────────
def scrape(yesterday: str) -> list[list]:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1280, "height": 900})
        page = context.new_page()

        try:
            # 1) 로그인 페이지 이동
            print("[STEP 1] 로그인 페이지 접속")
            page.goto(ADMIN_URL, wait_until="networkidle")
            time.sleep(2)

            # 로그인 폼 대기 (ID 입력 필드)
            id_input = page.locator(
                "input[type='text'], input[type='email'], input[placeholder*='아이디'], "
                "input[placeholder*='ID'], input[placeholder*='id'], input[name='username'], "
                "input[name='userId'], input[name='loginId']"
            ).first
            id_input.wait_for(state="visible", timeout=15000)

            # 2) 로그인
            print("[STEP 2] 로그인")
            id_input.fill(ADMIN_ID)

            pw_input = page.locator(
                "input[type='password']"
            ).first
            pw_input.fill(ADMIN_PW)

            login_btn = page.locator(
                "button[type='submit'], button:has-text('로그인'), button:has-text('Login'), "
                "button:has-text('로 그 인'), input[type='submit']"
            ).first
            login_btn.click()

            # 로그인 완료 대기 (URL 변경 또는 대시보드 요소)
            page.wait_for_url(lambda url: "login" not in url.lower(), timeout=15000)
            print(f"[INFO] 로그인 성공. 현재 URL: {page.url}")
            time.sleep(2)

            # 3) 광고 리포트 메뉴 이동
            print("[STEP 3] 광고 리포트 페이지 이동")
            page.goto(f"{ADMIN_URL}/#/report/ads", wait_until="networkidle")
            time.sleep(3)

            # 4) 날짜 입력 (전날)
            print(f"[STEP 4] 날짜 입력: {yesterday}")
            _set_date(page, yesterday)

            # 5) 조회 버튼 클릭
            print("[STEP 5] 조회")
            search_btn = page.locator(
                "button:has-text('조회'), button:has-text('검색'), button:has-text('Search'), "
                "button[type='submit']:visible"
            ).first
            search_btn.click()

            # AG Grid 데이터 행 대기
            page.wait_for_selector(".ag-center-cols-container .ag-row", timeout=20000)
            time.sleep(2)

            # 6) 테이블 데이터 추출
            print("[STEP 6] 데이터 추출")
            rows = _extract_table(page, yesterday)
            print(f"[INFO] 추출된 행 수: {len(rows)}")

            return rows

        except PWTimeoutError as e:
            # 디버깅용 스크린샷
            page.screenshot(path="error_screenshot.png")
            print(f"[ERROR] Timeout: {e}")
            raise
        except Exception as e:
            page.screenshot(path="error_screenshot.png")
            print(f"[ERROR] {e}")
            raise
        finally:
            browser.close()


def _set_date(page, date_str: str):
    """날짜 입력 필드에 값 설정. antd RangePicker 기준."""

    # 1) 일반 date input
    date_inputs = page.locator("input[type='date']").all()
    if date_inputs:
        for inp in date_inputs:
            inp.fill(date_str)
        return

    # 2) antd RangePicker: 시작일 입력
    start_input = page.locator(".ant-picker-input input").first
    start_input.click()
    time.sleep(0.5)
    start_input.press("Control+a")
    start_input.press("Backspace")
    time.sleep(0.2)
    start_input.type(date_str, delay=50)
    time.sleep(0.3)
    page.keyboard.press("Enter")
    time.sleep(0.5)

    # 3) 종료일 입력
    end_input = page.locator(".ant-picker-input input").last
    end_input.click()
    time.sleep(0.3)
    end_input.press("Control+a")
    end_input.press("Backspace")
    time.sleep(0.2)
    end_input.type(date_str, delay=50)
    time.sleep(0.3)
    page.keyboard.press("Enter")
    time.sleep(0.8)

    # 4) 달력 팝업 닫기
    try:
        page.wait_for_selector(".ant-picker-dropdown", state="hidden", timeout=3000)
    except Exception:
        page.keyboard.press("Escape")
        time.sleep(0.5)
        try:
            page.wait_for_selector(".ant-picker-dropdown", state="hidden", timeout=3000)
        except Exception:
            page.mouse.click(100, 400)
            time.sleep(1)

    print(f"[INFO] 날짜 설정 완료: {date_str} ~ {date_str}")


def _extract_table(page, date_str: str) -> list[list]:
    """AG Grid 테이블에서 행 추출. [날짜, CD, 광고명, OS, 광고타입, 광고단가, 조회수, 클릭수, 전환수, 전환율, 광고비]"""
    js_result = page.evaluate("""
        () => {
            const pinnedRows = document.querySelectorAll('.ag-pinned-left-cols-container .ag-row');
            const centerRows = document.querySelectorAll('.ag-center-cols-container .ag-row');
            const result = [];
            for (let i = 0; i < centerRows.length; i++) {
                const leftCells = pinnedRows[i]
                    ? Array.from(pinnedRows[i].querySelectorAll('.ag-cell')).map(c => c.innerText.trim())
                    : [];
                const centerCells = Array.from(centerRows[i].querySelectorAll('.ag-cell')).map(c => c.innerText.trim());
                const row = [...leftCells, ...centerCells];
                if (row.some(c => c !== '')) result.push(row);
            }
            return result;
        }
    """)

    rows_data = []
    for row in (js_result or []):
        rows_data.append([date_str] + row)

    return rows_data


# ─── 메인 ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 50)
    print("광고 리포트 자동 수집 시작")
    print("=" * 50)

    rows = scrape(yesterday)

    if not rows:
        print("[WARN] 수집된 데이터가 없습니다.")
        sys.exit(0)

    sheet = get_sheet()
    append_rows(sheet, rows)

    print("=" * 50)
    print("완료!")
    print("=" * 50)
