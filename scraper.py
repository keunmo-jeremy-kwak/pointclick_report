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

# 날짜 (전날)
yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
print(f"[INFO] 조회 날짜: {yesterday}")


# ─── Google Sheets 연결 ────────────────────────────────────────────────────────
def get_sheet():
    creds_dict = json.loads(GCP_CREDS_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)


def append_rows(sheet, rows: list[list]):
    """시트에 행 추가. 헤더가 없으면 먼저 삽입."""
    existing = sheet.get_all_values()
    if not existing:
        header = ["날짜", "CD", "광고명", "OS", "광고 타입",
                  "광고 단가", "조회수", "클릭수", "전환수", "전환율", "광고비"]
        sheet.append_row(header)

    for row in rows:
        sheet.append_row(row)
    print(f"[INFO] {len(rows)}행 저장 완료")


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

            # 결과 테이블 대기
            page.wait_for_selector("table, .ant-table, .el-table, [role='table']", timeout=20000)
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
    """날짜 입력 필드에 값 설정. 시작일/종료일 모두 yesterday로 설정."""

    # 1) 일반 date input
    date_inputs = page.locator("input[type='date']").all()
    if date_inputs:
        for inp in date_inputs:
            inp.fill(date_str)
        return

    # 2) antd RangePicker / 일반 DatePicker input
    date_pickers = page.locator(
        ".ant-picker input, .ant-picker-input input, "
        ".el-date-editor input, input[placeholder*='날짜'], "
        "input[placeholder*='YYYY'], input[placeholder*='yyyy']"
    ).all()
    if date_pickers:
        for inp in date_pickers[:2]:  # 시작일, 종료일 최대 2개
            inp.click(click_count=3)
            inp.fill(date_str)
            page.keyboard.press("Tab")
            time.sleep(0.5)
        # 달력 팝업 닫기
        page.keyboard.press("Escape")
        time.sleep(0.5)
        return

    print("[WARN] 날짜 입력 필드를 찾지 못했습니다. 수동 셀렉터 확인 필요.")


def _extract_table(page, date_str: str) -> list[list]:
    """테이블에서 모든 행 추출. [날짜, CD, 광고명, OS, 광고타입, 광고단가, 조회수, 클릭수, 전환수, 전환율, 광고비]"""
    rows_data = []

    # JavaScript로 테이블 데이터 추출 (thead 제외, tbody tr)
    js_result = page.evaluate("""
        () => {
            const tables = document.querySelectorAll('table');
            if (!tables.length) return [];

            // 가장 많은 행을 가진 테이블 선택
            let target = tables[0];
            for (const t of tables) {
                if (t.querySelectorAll('tbody tr').length > target.querySelectorAll('tbody tr').length) {
                    target = t;
                }
            }

            const rows = [];
            target.querySelectorAll('tbody tr').forEach(tr => {
                const cells = [];
                tr.querySelectorAll('td').forEach(td => {
                    cells.push(td.innerText.trim());
                });
                if (cells.length > 0) rows.push(cells);
            });
            return rows;
        }
    """)

    if not js_result:
        # antd / el-table 등 가상 테이블 대응
        js_result = page.evaluate("""
            () => {
                const rows = [];
                const rowEls = document.querySelectorAll(
                    '.ant-table-tbody tr, .el-table__body tr, [role="row"]'
                );
                rowEls.forEach(tr => {
                    const cells = [];
                    tr.querySelectorAll('td, [role="cell"], .cell').forEach(td => {
                        cells.push(td.innerText.trim());
                    });
                    if (cells.length > 0) rows.push(cells);
                });
                return rows;
            }
        """)

    for row in js_result:
        if row and any(cell.strip() for cell in row):
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
