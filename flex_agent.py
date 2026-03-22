import os
import json
import asyncio
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, cast
from dotenv import load_dotenv # type: ignore
from google import genai  # type: ignore
from google.genai import types  # type: ignore
import gspread  # type: ignore

# .env 파일 로드
load_dotenv() # type: ignore
from google.oauth2.service_account import Credentials  # type: ignore
from googleapiclient.discovery import build  # type: ignore
import requests  # type: ignore
from playwright.async_api import async_playwright  # type: ignore
from workflow_handlers import dispatch_workflow  # type: ignore

async def ainput(prompt: str) -> str:
    """비동기 방식으로 사용자 입력을 받는 헬퍼 함수"""
    loop = asyncio.get_event_loop()
    result = await loop.run_in_executor(None, input, prompt)
    return str(result)


GUIDE_FILE = "guide.json"
COURSES_FILE = "courses.json"

def load_courses() -> Dict[str, Any]:
    if os.path.exists(COURSES_FILE):
        with open(COURSES_FILE, "r", encoding="utf-8") as f:
            return cast(Dict[str, Any], json.load(f))
    return {}

def save_courses(courses: Dict[str, Any]) -> None:
    with open(COURSES_FILE, "w", encoding="utf-8") as f:
        json.dump(courses, f, indent=4, ensure_ascii=False)

def load_config() -> Dict[str, Any]:
    config: Dict[str, Any] = {
        "gemini_api_key": os.getenv("GEMINI_API_KEY"),
        "flex_id": os.getenv("FLEX_ID"),
        "flex_pw": os.getenv("FLEX_PW"),
        "service_account_path": os.getenv("SERVICE_ACCOUNT_PATH"),
        "notion_url": os.getenv("NOTION_URL")
    }
    return config

def delete_config():
    if os.path.exists(".env"):
        os.remove(".env")
        print("\n[시스템] 🗑️ 저장된 환경 변수가 모두 삭제되었습니다.")
    else:
        print("\n[시스템] ℹ️ 삭제할 환경 변수가 없습니다.")

def save_config(gemini_key, flex_id, flex_pw, service_account_path, notion_url):
    # 만약 service_account_path가 현재 스크립트 디렉토리 안에 있다면 상대 경로로 변환하여 저장
    base_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.isabs(service_account_path) and service_account_path.startswith(base_dir):
        service_account_path = os.path.relpath(service_account_path, base_dir)

    # .env 파일에 모든 정보 저장
    env_content = f"""GEMINI_API_KEY={gemini_key}
FLEX_ID={flex_id}
FLEX_PW={flex_pw}
SERVICE_ACCOUNT_PATH={service_account_path}
NOTION_URL={notion_url}
"""
    with open(".env", "w", encoding="utf-8") as f:
        f.write(env_content)
        
    # 현재 실행 중인 프로세스의 환경 변수 즉시 업데이트
    os.environ["GEMINI_API_KEY"] = gemini_key
    os.environ["FLEX_ID"] = flex_id
    os.environ["FLEX_PW"] = flex_pw
    os.environ["SERVICE_ACCOUNT_PATH"] = service_account_path
    os.environ["NOTION_URL"] = notion_url
        
    print("\n[시스템] ✅ 환경 변수(.env)에 모든 필수 설정이 안전하게 저장되었습니다.")

def load_guide() -> Optional[Dict[str, Any]]:
    if os.path.exists(GUIDE_FILE):
        try:
            with open(GUIDE_FILE, "r", encoding="utf-8") as f:
                return cast(Dict[str, Any], json.load(f))
        except:
            return None
    return None

def save_guide(guide_data):
    with open(GUIDE_FILE, "w", encoding="utf-8") as f:
        json.dump(guide_data, f, indent=4, ensure_ascii=False)
    print("\n[시스템] 💾 업무 가이드가 로컬에 저장되었습니다.")

async def fetch_notion_guide(notion_url):
    """
    Playwright를 사용하여 노션 페이지의 내용을 실시간으로 추출하고 로컬에 저장합니다.
    """
    print(f"\n[시스템] 🌐 노션 가이드 읽어오는 중: {notion_url}")
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
        page = await context.new_page()
        page.set_default_timeout(60000)
        
        try:
            await page.goto(notion_url, wait_until="domcontentloaded")
            await page.wait_for_selector(".notion-page-content", state="visible", timeout=60000)
            await asyncio.sleep(3)
            
            content = await page.inner_text("body")
            templates = re.findall(r"\[[^\]]+\]", content)
            unique_templates = list(dict.fromkeys(templates))
            
            guide_data = {
                "raw_text": content,
                "templates": unique_templates,
                "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            save_guide(guide_data)
            return guide_data
        except Exception as e:
            print(f"[시스템] ❌ 노션 읽기 오류: {e}")
            return None
        finally:
            await browser.close()

# --- Tools and Helpers ---

def master_sheet_reader_tool(sheet_url: str):
    """
    구글 시트 링크에서 원본 데이터를 읽어 JSON 형태로 반환합니다.
    """
    print(f"\n[시스템] 📊 구글 시트 데이터 추출 시작: {sheet_url}")
    try:
        config = load_config()
        service_account_file = config.get("service_account_path")
        
        if not service_account_file or not os.path.exists(service_account_file):
            return "Error: 서비스 계정 파일 경로 오류"

        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
        if not match: return "Error: 올바른 구글 시트 URL 형식이 아닙니다."
        spreadsheet_id = match.group(1)
        
        scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        creds = Credentials.from_service_account_file(service_account_file, scopes=scopes)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(spreadsheet_id)
        
        all_data = {}
        for sheet in spreadsheet.worksheets():
            all_data[sheet.title] = sheet.get_all_values()
            
        return json.dumps(all_data, ensure_ascii=False)
    except Exception as e:
        return f"Error: {str(e)}"

def fetch_filtered_sheet_data_tool(sheet_url: str, sheet_name: str):
    """
    구글 시트의 특정 탭(sheet_name)에서 데이터를 가져옵니다.
    로직: 2행부터 시작하며, A열(첫 번째 셀)에 데이터가 있는 행만 추출합니다.
    표 형식의 데이터를 리스트의 리스트 형태로 반환합니다.
    """
    print(f"\n[시스템] 📊 구글 시트 필터링 추출 시작: {sheet_url} (탭: {sheet_name})")
    try:
        config = load_config()
        service_account_file = config.get("service_account_path")
        
        if not service_account_file or not os.path.exists(service_account_file):
            return "Error: 서비스 계정 파일 경로 오류"

        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
        if not match: return "Error: 올바른 구글 시트 URL 형식이 아닙니다."
        spreadsheet_id = match.group(1)
        
        scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        creds = Credentials.from_service_account_file(service_account_file, scopes=scopes)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(spreadsheet_id)
        
        worksheet = spreadsheet.worksheet(sheet_name)
        all_values = worksheet.get_all_values()
        
        if len(all_values) < 2:
            return "Error: 데이터가 충분하지 않습니다 (1행만 있거나 비어 있음)."

        # 1행은 헤더로 간주 (필요 시 포함), 2행부터 필터링
        header = all_values[0]
        filtered_rows = [header] # 헤더 포함
        
        for row_raw in all_values[1:]:
            row: Any = row_raw
            if row and len(row) > 0 and row[0].strip(): # type: ignore # A열 데이터가 있는 경우
                filtered_rows.append(row)
                
        return json.dumps(filtered_rows, ensure_ascii=False)
    except Exception as e:
        return f"Error: {str(e)}"

def fetch_rich_sheet_data_tool(sheet_url: str, sheet_name: str):
    """
    구글 시트의 특정 탭(sheet_name)에서 데이터와 메타데이터(병합, 배경색, 글꼴)를 가져옵니다.
    데이터 전처리는 하지 않으며, 원본 격자 데이터를 JSON 형태로 반환합니다.
    """
    print(f"\n[시스템] 🎨 구글 시트 리치 데이터 추출 시작: {sheet_url} (탭: {sheet_name})")
    try:
        config = load_config()
        sa_path = config.get("service_account_path")
        if not sa_path or not os.path.exists(sa_path):
            return "Error: 서비스 계정 파일 경로 오류"

        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
        if not match: return "Error: 올바른 구글 시트 URL 형식이 아닙니다."
        ss_id = match.group(1)

        scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        creds = Credentials.from_service_account_file(sa_path, scopes=scopes)
        service = build('sheets', 'v4', credentials=creds)

        # 시트 메타데이터에서 탭 ID 찾기 및 머지 정보 가져오기
        spreadsheet = service.spreadsheets().get(spreadsheetId=ss_id).execute()
        sheet_id = None
        merges = []
        for s in spreadsheet.get('sheets', []):
            if s.get('properties', {}).get('title') == sheet_name:
                sheet_id = s.get('properties', {}).get('sheetId')
                merges = s.get('merges', [])
                break
        
        if sheet_id is None:
            return f"Error: '{sheet_name}' 탭을 찾을 수 없습니다."

        # 실제 데이터 및 스타일 가져오기 (2행부터 마지막까지 자동 인식하도록 전체 범위 지정 시도)
        # 1-indexed, A:Z 등 전체 범위를 위해 '탭이름'!A1:Z500 정도로 넉넉히 잡거나 get 시 ranges 생략
        range_name = f"'{sheet_name}'!A1:Z500" 
        result = service.spreadsheets().get(
            spreadsheetId=ss_id,
            ranges=[range_name],
            includeGridData=True
        ).execute()

        sheet_data = result['sheets'][0]['data'][0]
        row_data = sheet_data.get('rowData', [])
        column_metadata = sheet_data.get('columnMetadata', [])
        
        # 열 너비 추출 (기본값 100px)
        column_widths = [m.get('pixelSize', 100) for m in column_metadata]

        values: List[List[Any]] = []
        backgrounds: List[List[Any]] = []
        font_weights: List[List[str]] = []

        for row in row_data:
            v_row: List[Any] = []
            b_row: List[Any] = []
            f_row: List[str] = []
            cells = row.get('values', [])
            for cell in cells:
                # 값 (포맷된 값 우선)
                v_row.append(cell.get('formattedValue', ''))
                
                # 배경색
                bg = cell.get('effectiveFormat', {}).get('backgroundColor', {})
                b_row.append(bg) # RGB Dict {red, green, blue}
                
                # 폰트 굵기
                weight = cell.get('effectiveFormat', {}).get('textFormat', {}).get('bold', False)
                f_row.append('bold' if weight else 'normal')
            
            values.append(v_row)
            backgrounds.append(b_row)
            font_weights.append(f_row)

        res_data: Dict[str, Any] = {
            "values": cast(Any, values)[1:],
            "backgrounds": cast(Any, backgrounds)[1:],
            "fontWeights": cast(Any, font_weights)[1:],
            "merges": merges,
            "columnWidths": column_widths,
            "startRow": 2
        }
        return json.dumps(res_data, ensure_ascii=False)

    except Exception as e:
        return f"Error: {str(e)}"

async def update_sheet_cell(config: dict, spreadsheet_url: str, update_info: dict):
    """
    구글 시트의 특정 셀의 상태값을 변경합니다. (예: [희망] -> [완료])
    """
    try:
        service_account_file = config.get("service_account_path")
        if not service_account_file or not os.path.exists(service_account_file):
            print("   ⚠️ 구글 시트 업데이트 오류: 서비스 계정 파일 경로 오류")
            return False

        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match: 
            print("   ⚠️ 구글 시트 업데이트 오류: 올바른 구글 시트 URL 형식이 아닙니다.")
            return False
        spreadsheet_id = match.group(1)
        
        scopes = ['https://www.googleapis.com/auth/spreadsheets']
        creds = Credentials.from_service_account_file(service_account_file, scopes=scopes)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(spreadsheet_id)
        
        sheet_name = update_info.get("sheet_name")
        row_index = update_info.get("row_index")
        col_index = update_info.get("col_index")
        target_value = update_info.get("target_value")

        if sheet_name and row_index and col_index and target_value:
            worksheet = spreadsheet.worksheet(sheet_name)
            worksheet.update_cell(int(row_index), int(col_index), str(target_value))
            print(f"   ✅ 시트 데이터 갱신 성공: '{sheet_name}' 탭의 상태를 '{target_value}'(으)로 변경했습니다.")
            return True
        else:
            print("   ⚠️ 구글 시트 업데이트 오류: 업데이트에 필요한 식별 정보가 부족합니다.")
            return False
    except Exception as e:
        print(f"   ⚠️ 구글 시트 업데이트 중 오류 발생: {e}")
        return False


async def analyze_master_sheet(sheet_url: str):
    """
    구글 시트 데이터를 분석하여 '기업명_과정명'을 추출합니다.
    """
    data_json = master_sheet_reader_tool(sheet_url)
    if data_json.startswith("Error"):
        print(f"\n[시스템] ❌ 분석 실패: {data_json}")
        return None
    
    try:
        all_data = json.loads(data_json)
        company_name = "미확인기업"
        course_name = "미확인과정"
        
        # 모든 시트에서 기업명과 과정명 키워드 검색
        for sheet_title, rows_any in all_data.items():
            rows: Any = rows_any
            for row_any in rows:
                row: Any = row_any
                row_str = " ".join([str(cell) for cell in row])
                if "기업명" in row_str or "고객사" in row_str:
                    # 간단한 매칭 logic (키워드 다음 셀이나 현재 행에서 추출 시도)
                    for i, cell in enumerate(row):
                        if any(k in str(cell) for k in ["기업명", "고객사"]):
                            if i + 1 < len(row) and str(row[i+1]).strip(): # type: ignore
                                company_name = str(row[i+1]).strip() # type: ignore
                                break
                if "과정명" in row_str or "교육명" in row_str:
                    for i, cell in enumerate(row):
                        if any(k in str(cell) for k in ["과정명", "교육명"]):
                            if i + 1 < len(row) and str(row[i+1]).strip(): # type: ignore
                                course_name = str(row[i+1]).strip() # type: ignore
                                break
        
        suggested_name = f"{company_name}_{course_name}"
        print(f"\n[시스템] 🔍 분석 결과 제안: {suggested_name}")
        
        final_name = await ainput(f"등록할 이름을 입력하세요 (엔터 시 '{suggested_name}' 사용): ")
        return final_name.strip() if final_name.strip() else suggested_name
        
    except Exception as e:
        print(f"\n[시스템] ❌ 분석 중 오류 발생: {e}")
        return None

async def manage_courses_menu():
    """
    교육 마스터 시트 링크를 등록, 수정, 삭제하는 메뉴입니다.
    """
    while True:
        courses: Dict[str, Any] = load_courses()
        print("\n" + "="*45)
        print("📊 교육 마스터 시트 관리")
        print("="*45)
        if not courses:
            print("등록된 교육이 없습니다.")
        else:
            for i, (name, info) in enumerate(courses.items(), 1):
                print(f"{i}. {name} \n   🔗 {info['url']}")
        
        print("-" * 45)
        print("1. 신규 교육 등록")
        print("2. 기존 교육 수정")
        print("3. 교육 삭제")
        print("4. 메인 메뉴로 돌아가기")
        print("="*45)
        
        choice = await ainput("입력: ")
        choice = choice.strip().lower()
        
        if choice == "1":
            url = await ainput("구글 시트 링크를 입력하세요: ")
            if not url: continue
            name = await analyze_master_sheet(url)
            if name:
                courses[cast(str, name)] = {"url": url, "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                save_courses(courses)
                print(f"\n[시스템] ✅ '{name}' 교육이 등록되었습니다.")
        
        elif choice == "2":
            if not courses: continue
            idx = await ainput("수정할 교육 번호: ")
            if idx.isdigit() and 1 <= int(idx) <= len(courses):
                old_name = list(courses.keys())[int(idx)-1]
                new_url = await ainput(f"새로운 URL (기존 유지 시 엔터): ")
                new_name = await ainput(f"새로운 이름 (기존 '{old_name}' 유지 시 엔터): ")
                
                info = cast(Dict[str, Any], courses.pop(old_name))
                if new_url.strip(): info['url'] = new_url.strip()
                final_name = str(new_name.strip()) if new_name.strip() else old_name
                courses[cast(str, final_name)] = info
                save_courses(courses)
                print(f"\n[시스템] ✅ 수정 완료되었습니다.")
        
        elif choice == "3":
            if not courses: continue
            idx = await ainput("삭제할 교육 번호: ")
            if idx.isdigit() and 1 <= int(idx) <= len(courses):
                name = list(courses.keys())[int(idx)-1]
                confirm = await ainput(f"'{name}'를 정말 삭제할까요? (y/n): ")
                if confirm.lower() == 'y':
                    courses.pop(name, None)
                    save_courses(courses)
                    print(f"\n[시스템] 🗑️ 삭제되었습니다.")
        
        elif choice == "4":
            break

async def execute_workflow(config: dict, plan: dict):
    """
    Playwright를 사용하여 실제 Flex 워크플로우를 작성하고 상신합니다.
    (현재 사용자 요청으로 제목만 입력하도록 축소됨)
    """
    flex_id = config.get("flex_id")
    flex_pw = config.get("flex_pw")
    template_name = plan.get("template")
    
    print(f"\n[시스템] 🚀 '{template_name}' 업무 자동 수행 엔진 기동...")
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False) 
        context = await browser.new_context()
        page = await context.new_page()
        
        try:
            await page.goto("https://flex.team/auth/login")
            await page.fill('input[name="email"]', flex_id)
            await page.keyboard.press("Enter")
            await page.wait_for_selector('input[name="password"]', timeout=30000)
            await page.fill('input[name="password"]', flex_pw)
            await page.keyboard.press("Enter")
            await page.wait_for_load_state("networkidle")
            await asyncio.sleep(2) 
            print("[시스템] ✅ Flex 로그인 완료 (시크릿 세션)")

            print("[시스템] 📂 워크플로우 메뉴로 이동 중...")
            try:
                nav_selector = 'nav >> text="워크플로우"'
                # 1. '워크플로우' 메뉴가 바로 보이는지 확인 (짧은 대기)
                try:
                    await page.wait_for_selector(nav_selector, timeout=3000)
                except:
                    # 2. 보이지 않으면 사이드 바 펼치기 시도
                    print("[시스템] 🔍 '워크플로우' 메뉴가 보이지 않아 사이드 바 펼치기 시도 중...")
                    # 사용자 제공 클래스 및 SVG 포함 버튼 탐색
                    expand_selectors = [
                        'button.c-MZCiC.c-MZCiC-aVorO-active-false',
                        'button.c-MZCiC:has(svg)',
                        '.c-MZCiC.c-MZCiC-aVorO-active-false',
                        'button:near(nav) >> svg'
                    ]
                    
                    expanded = False
                    for sel in expand_selectors:
                        try:
                            btn = page.locator(sel).first
                            if await btn.is_visible(timeout=1000):
                                await btn.click()
                                await asyncio.sleep(1.5) # 애니메이션 대기
                                expanded = True
                                print(f"   ✅ 사이드 바 펼치기 클릭 성곡 (Selector: {sel})")
                                break
                        except: continue
                    
                    if not expanded:
                        print("[시스템] ⚠️ 사이드 바 펼치기 버튼을 찾지 못했습니다. 직접 이동을 시도합니다.")

                # 3. 다시 '워크플로우' 클릭 시도
                await page.wait_for_selector(nav_selector, timeout=5000)
                await page.click(nav_selector)
            except:
                print("[시스템] ⚠️ 메뉴 클릭 실패, 직접 URL로 이동합니다.")
                await page.goto("https://flex.team/workflow/request")
                
            await page.wait_for_load_state("networkidle")
            await asyncio.sleep(2)
            
            # 이미 템플릿 검색창(양식/워크플로우)이 보이는지 확인
            search_input_template = page.locator('input[placeholder*="양식"], input[placeholder*="워크플로우"]').first
            is_template_search_visible = await search_input_template.is_visible()
            
            if not is_template_search_visible:
                try:
                    # '문서 작성' 계열 버튼 클릭
                    btn_patterns = [r"^문서 작성하기$", r"^문서 작성$", r"^요청하기$"]
                    for pattern in btn_patterns:
                        btn = page.get_by_role("button", name=re.compile(pattern)).first
                        if await btn.is_visible(timeout=2000):
                            await btn.click()
                            await asyncio.sleep(1)
                            break
                except:
                    pass
            
            # 템플릿 검색창 찾기 및 입력
            # [계약서 등 검토 · 승인] 교육 용역 -> [계약서 등 검토 · 승인] 으로 변환하여 검색
            # 단, 사용자 요청: "[계약서 등 검토 · 승인] 계약명을 적어주세요." 텍스트로 검색
            search_name = template_name
            if template_name and "[계약서 등 검토 · 승인]" in template_name:
                search_name = "[계약서 등 검토 · 승인] 계약명을 적어주세요."
            elif template_name and "]" in template_name:
                search_name = template_name.split("]")[0] + "]"
            
            print(f"[시스템] 템플릿 검색 중: {search_name} (원본: {template_name})")
            try:
                # 1. 템플릿 검색창이 나타날 때까지 대기
                search_selector = 'input[placeholder*="양식"], input[placeholder*="워크플로우"], div[role="dialog"] input[placeholder*="검색"]'
                search_input = page.locator(search_selector).first
                await search_input.wait_for(state="visible", timeout=10000)
                
                # 2. 직접 값을 채워넣고 엔터
                await search_input.fill(str(search_name))
                await asyncio.sleep(0.5)
                await search_input.press("Enter")
                print(f"   ✅ 검색어 입력 및 엔터 완료")
                
                # 엔터 후 양식이 자동으로 선택되어 화면이 전환되었는지 확인
                await asyncio.sleep(3) # 전환 대기
                
                # 검색창이 여전히 가시적인지 확인. 사라졌다면 이미 선택된 것임.
                if not await search_input.is_visible():
                    print(f"   ✅ 화면 전환 감지 (엔터로 양식 선택됨)")
                    # 여기서 return 하지 않고 뒤의 선택 블록을 건너뛰게 함
                    search_completed = True
                else:
                    search_completed = False

            except Exception as e:
                print(f"   ⚠️ 검색창 입력 오류: {e}")
                search_completed = False

            if not search_completed:
                await asyncio.sleep(3) 
                
                # 검색 결과 내에서 템플릿 찾기 및 클릭
                try:
                    # 1. '검색 결과' 섹션 내에서 정확한 텍스트 매칭 시도
                    template_locator = page.locator(f'text="{search_name}"').filter(has_not=page.locator('input')).first
                    if await template_locator.is_visible(timeout=2000):
                        await template_locator.click()
                        print(f"   ✅ '{search_name}' 템플릿 선택 완료")
                    else:
                        raise Exception("정확한 매칭 결과 없음")
                except:
                    try:
                        # 2. 폴백: 텍스트 포함 검색
                        clean_name = str(search_name or "").replace("[", "").replace("]", "")
                        if " " in clean_name: clean_name = clean_name.split(" ")[0] # 첫 단어만 추출 (엄격한 매칭 방지)
                        
                        print(f"   ⚠️ 정확한 매칭 실패, '{clean_name}' 포함 요소 탐색 중...")
                        fallback_locator = page.get_by_text(clean_name).first
                        await fallback_locator.wait_for(state="visible", timeout=3000)
                        await fallback_locator.click()
                        print(f"   ✅ '{clean_name}' 포함 템플릿 선택 완료")
                    except:
                        # 3. 최종 폴백: 검색 결과 리스트의 첫 번째 항목 클릭
                        try:
                            print(f"   ⚠️ 텍스트 검색 모두 실패, 검색 결과의 첫 번째 항목 시도 중...")
                            first_item = page.locator('div[role="listbox"] div[role="option"], ul > li div').first
                            if await first_item.is_visible(timeout=3000):
                                await first_item.click()
                                print(f"   ✅ 검색 결과 첫 번째 항목 선택 완료")
                            else:
                                # 이미 화면이 넘어갔을 가능성 최종 확인
                                if not await page.locator(search_selector).first.is_visible():
                                    print(f"   ✅ 이미 양식이 선택되어 화면이 전환되었습니다.")
                                else:
                                    raise Exception("항목을 찾을 수 없음")
                        except Exception as final_err:
                            # 이미 화면이 넘어갔을 수 있음 (검색창이 없으면 성공으로 간주)
                            try:
                                if not await page.locator(search_selector).first.is_visible():
                                    print(f"   ✅ 화면 전환 확인됨 (에러 무시)")
                                else:
                                    raise final_err
                            except:
                                print(f"   ❌ 모든 템플릿 선택 시도 실패: {final_err}")
                                raise Exception(f"'{search_name}' 양식을 선택하지 못했습니다.")
                
            await page.wait_for_load_state("networkidle")
            await asyncio.sleep(2) 
            
            # 4. 데이터 입력 (모듈화된 핸들러 호출)
            print("[시스템] 📝 데이터 입력 중 (양식별 전용 핸들러 호출)...")
            
            # 양식별 전용 핸들러 실행 시도 (계획 데이터 및 설정 전달)
            handled = await dispatch_workflow(template_name, page, plan, config)
            
            if not handled:
                print(f"[시스템] ℹ️ '{template_name}'에 대한 전용 핸들러가 없어 기본 동작만 수행합니다.")
                # 기본 동작: 제목 입력 시도
                title_value = plan.get("title")
                if title_value:
                    try:
                        await page.fill('input[placeholder*="제목"]', title_value)
                        print(f"   ✅ '제목' 기본 입력 완료")
                    except: pass

            print("[시스템] ℹ️ 사용자 요청에 따라 자동화 필드 외 기타 정보는 브라우저에서 직접 확인해 주세요.")
            print("\n[시스템] ✨ 자동 입력 후 브라우저에서 나머지 내용을 확인해 주세요.")
            await ainput("\n[시스템] 내용을 확인하신 후, 브라우저를 닫으려면 엔터를 누르세요...")
            
        except Exception as e:
            print(f"\n[시스템] ❌ 자동 수행 중 오류가 발생했습니다: {e}")
            await ainput("\n[시스템] 오류가 발생했습니다. 확인 후 엔터를 누르면 브라우저를 닫습니다...")
        finally:
            await browser.close()

# --- Planning Flow ---

async def run_planning_flow(config: dict, template_name: str, guide_content: str):
    """
    업무 계획을 수립하고 승인 시 자동 수행을 호출합니다.
    """
    gemini_key = config.get('gemini_api_key')
    if not gemini_key:
        print("\n[오류] Gemini API Key가 설정되지 않았습니다. 메인 메뉴의 3번 메뉴에서 환경 변수를 먼저 설정해 주세요.")
        return
        
    client = genai.Client(api_key=gemini_key)
    model_name = "gemini-2.5-flash"
    
    system_instruction = f"""
# Role
당신은 사용자가 선택한 Flex 워크플로우 양식({template_name})에 맞춰 업무 계획을 세워주는 비서입니다.
제공된 '실시간 노션 업무 가이드' 내용을 철저히 준수하여 구글 시트 데이터를 어떻게 Flex에 매핑할지 계획을 세워야 합니다.

# Real-time Notion Guide Content
{guide_content}

# CRITICAL RULES (MUST FOLLOW)
1. 현재 시스템은 **'제목', '시작일', '종료일', '예상 비용', '예상 매출', '파일 첨부'까지 자동으로 수행**할 수 있습니다. 
2. **동적 가이드 분석 (필수)**: 제공된 `Real-time Notion Guide Content`에서 현재 선택된 양식 `{template_name}`에 해당하는 섹션을 찾아 '데이터 매핑 테이블'과 '본문 작성 규칙'을 분석하세요.
3. 가이드에 정의된 매핑 규칙(예: `개요!B11` 셀 등)에 따라 `fields`를 구성하고, 본문 작성 규칙(예: I. 개요, II. 내용 구성)에 맞춰 `본문 내용`을 생성하세요.
4. 만약 현재 양식에 대한 명시적인 규칙이 가이드에 없다면, 일반적인 상식(Common Sense)에 기반하여 매핑 계획을 세우되, 이 사실을 사용자에게 고지하세요.
5. **표 데이터 추출 (중요)**: 
    - 사용자가 표 데이터를 요청하거나 가이드에 표 추출 규칙이 있다면 `fetch_rich_sheet_data_tool`을 호출하세요. 
    - 이 도구는 병합 정보와 스타일까지 포함된 JSON을 반환합니다.
6. **데이터 전달 방식**:
    - 도구의 결과(JSON 문자열)를 가공하지 말고 **그대로** JSON의 `table_data` 필드에 넣으세요.
    - AI가 직접 표를 그리지 말고, 핸들러가 처리할 수 있도록 원본 데이터를 전달하는 것이 목적입니다.
7. 설명 텍스트에서 자동으로 수행되는 항목(제목, 시작/종료일, 비용/매출, 본문, 파일 첨부)을 안내하고, 자동화되지 않는 부분은 수동 입력을 요청하세요.
8. **매핑 우선순위 규칙 (중요)**: 
    - 가이드에 **(고정값)** 이라고 명시된 필드는 해당 고정값을 사용하세요.
    - 그 외의 모든 필드는 구글 시트의 지정된 위치에서 데이터를 찾아 매핑하세요.
    - 예: '교육 용역'의 '매출 · 매입'은 고정값(매출)이지만, '상태의 유형'이나 '시작일' 등은 시트 데이터나 유도가 필요할 수 있으니 가이드를 꼼꼼히 확인하세요.
9. **본문 작성 규칙 특이사항**: '본문 내용' 필드 작성 시, **표가 들어갈 섹션(예: "II. 예상 비용", "III. 예상 비용" 등)의 제목이나 헤더는 절대로 본문에 포함하지 마세요.** 핸들러가 자동으로 "III. 예상 비용"이라는 제목과 함께 표를 알맞은 위치에 삽입합니다. 본문에는 표 직전의 내용까지만 작성하세요.
9. 터미널에 표를 출력할 때는 `rich` 라이브러리를 쓰는 것처럼 시각적으로 예쁘게 보여주겠다고 언급하세요.
10. **본문 내용 포맷팅 규칙 (중요)**:
    - 모든 섹션 제목(예: **I. 개요**, **II. 상세 내용** 등)은 반드시 `<b>` 태그를 사용하여 **볼드 처리**하세요. (예: `<b>I. 개요</b>`)
    - **중요**: 섹션 제목(`<b>` 태그`)과 그 바로 아래 본문 사이에는 절대로 빈 줄을 넣지 마세요.
    - 대신, 각 섹션이나 문단 사이(내용이 끝나고 다음 제목이 올 때)에는 가독성을 위해 **두 번의 줄바꿈**(`\n\n`)을 넣어 빈 줄을 만드세요.
11. **교육 용역(Education Services) 본문 필수 필드**:
    - **I. 개요**: 교육 명(B4), 교육 일(B10), 계약 담당자(B23)
    - **II. 매출 및 비용**: 총 매출(B14, VAT 별도 기재), 예상 지출(B13)
    - **III. 표준 계약서 기준 변경 사항**: B24 내용

# Output Rules
1. 사용자에게 먼저 친절한 '업무 계획'을 한국어로 설명하세요.
2. 설명 끝에 반드시 자동화를 위한 구조화된 JSON 데이터 **배열**을 ```json ... ``` 블록 안에 포함하세요.
3. 대상 건(기안해야 할 문서)이 여러 개일 경우, 각각을 개별 JSON 객체로 만들고 이들을 묶어 단일 JSON **배열(List)** 형태로 반환해야 합니다. 대상 건이 1개라도 반드시 배열(`[...]`) 형태로 반환하세요.
4. **상태값 업데이트 식별자 제공 (중요)**: 기안이 다 끝나고 난 후, 해당 건의 구글 시트 상태값(`[희망]` 등)을 시스템이 `완료`로 자동 변경할 수 있도록 `sheet_update_info`를 추가하세요. 이 정보 안에는 상태값이 위치한 엑셀 기반 행(row) 순서 번호와 열(col) 순서 번호 정보를 넣어야 합니다.
5. 각 JSON 객체의 형식 (필수 필드 - 양식에 따라 필드명은 달라질 수 있지만 아래 구조는 유지하세요):
    [
        {{
            "template": "{template_name}",
            "title": "워크플로우 제목 (가이드 규칙 적용)",
            "spreadsheet_url": "사용된 시트의 전체 URL",
            "fields": {{
                "본문 내용": "가이드 작성 규칙이 적용된 텍스트",
                "시작일": "YYYY-MM-DD",
                "종료일": "YYYY-MM-DD (가이드에 있는 경우)",
                "예상 비용": "123456 (가이드에 있는 경우)",
                ...
            }},
            "table_data": [["헤더1", "헤더2"], ["데이터1", "데이터2"]],
            "attachments": ["list_of_urls_to_download"],
            "sheet_update_info": {{
                "sheet_name": "상태값을 변경할 탭 이름 (예: [운영] 강사 계약)",
                "row_index": 5,
                "col_index": 2,
                "target_value": "완료"
            }}
        }}
    ]
"""


    
    chat = client.chats.create(
        model=model_name,
        config=types.GenerateContentConfig(
            system_instruction=system_instruction.strip(),
            tools=[master_sheet_reader_tool, fetch_filtered_sheet_data_tool, fetch_rich_sheet_data_tool],
            temperature=0.3,
        )
    )
    
    print(f"\n🤖 AI 비서가 '{template_name}' 양식에 대한 업무 계획을 준비 중입니다...")
    
    courses = load_courses()
    selected_url = None
    
    if courses:
        print("\n--- 등록된 교육 마스터 시트 중 선택하거나 직접 입력해 주세요 ---")
        course_names = list(courses.keys())
        for i, name in enumerate(course_names, 1):
            print(f"{i}. {name}")
        print(f"{len(course_names)+1}. 직접 URL 입력")
        
        c_choice = await ainput("선택: ")
        if c_choice.isdigit() and 1 <= int(c_choice) <= len(course_names):
            selected_url = cast(Dict[str, Any], courses[course_names[int(c_choice)-1]])['url']
            print(f"[시스템] 📂 '{course_names[int(c_choice)-1]}' 시트를 사용합니다.")
        elif c_choice.isdigit() and int(c_choice) == len(course_names) + 1:
            selected_url = await ainput("URL 입력: ")
    else:
        selected_url = await ainput("\n👤 사용자 (시트 링크 입력): ")
    
    if not selected_url:
        return

    # [계약서 등 검토 · 승인] 교육 용역의 경우 시트 분석 전 미리 추가 정보 입력 (사용자 요청으로 제거됨)
    extra_fields = {}

    try:
        user_input = selected_url.strip()
        prompt_with_extra = f"선택한 양식: {template_name}\n시트 링크: {user_input}"
        if extra_fields:
            prompt_with_extra += f"\n추가 입력 정보: {json.dumps(extra_fields, ensure_ascii=False)}"
        
        try:
            response = chat.send_message(f"{prompt_with_extra}\n위 정보를 바탕으로 실시간 노션 가이드를 참고하여 업무 계획을 세워줘.")
            
            if response.candidates and response.candidates[0].content and response.candidates[0].content.parts:
                assistant_text = "".join([part.text for part in response.candidates[0].content.parts if part.text])
            else:
                print("\n[시스템] ⚠️ AI가 업무 계획을 생성하지 못했습니다. (응답이 비어 있거나 안전 정책에 의해 차단되었을 수 있습니다.)")
                print("[시스템] 💡 프로그램을 종료하고 다시 시작해 주세요.")
                return
        except Exception as e:
            print(f"\n[시스템] ❌ AI 응답 처리 중 오류 발생: {e}")
            print("[시스템] 💡 프로그램을 종료하고 다시 시작해 주세요.")
            return
        
        json_match = re.search(r"```json\s*(\[.*?\]|\{.*?\})\s*```", assistant_text, re.DOTALL)
        plan_data_list = []
        if json_match:
            try:
                parsed_json = json.loads(json_match.group(1))
                if isinstance(parsed_json, dict):
                    plan_data_list = [parsed_json]  # 단일 객체인 경우 배열로 감싸기
                elif isinstance(parsed_json, list):
                    plan_data_list = parsed_json

                clean_text = assistant_text.replace(json_match.group(0), "").strip()
                print(f"\n🤖 AI 비서의 업무 계획:\n{clean_text}")
                
                if plan_data_list:
                    # 미리 수집한 추가 정보 주입
                    if extra_fields:
                        for plan_data in plan_data_list:
                            for k, v in extra_fields.items():
                                plan_data.setdefault("fields", {})[k] = v

                    print(f"\n[시스템] 업무 계획({len(plan_data_list)}건) 수립 완료. 즉시 자동 수행을 시작합니다.")
                    for idx, plan_data in enumerate(plan_data_list, 1):
                        print(f"\n==================================================")
                        print(f"[{idx}/{len(plan_data_list)}] 진행 문서: {plan_data.get('title', '제목 없음')}")
                        print(f"==================================================")
                        await execute_workflow(config, plan_data)
                        
                        # 실행 후 스프레드시트 업데이트 연동 로직
                        sheet_update_info = plan_data.get("sheet_update_info")
                        spreadsheet_url = plan_data.get("spreadsheet_url") or selected_url
                        if sheet_update_info and isinstance(sheet_update_info, dict):
                            print(f"\n[시스템] 업무 완료 후 구글 시트 상태 업데이트를 시작합니다...")
                            await update_sheet_cell(config, str(spreadsheet_url), sheet_update_info)

                    print(f"\n[시스템] 총 {len(plan_data_list)}건의 자동 수행이 모두 종료되었습니다.")
            except Exception as e:
                print(f"\n🤖 AI 비서의 업무 계획 (데이터 파싱/실행 실패: {e}):\n{assistant_text}")
        else:
            print(f"\n🤖 AI 비서의 업무 계획:\n{assistant_text}")
    except Exception as e:
        print(f"\n[오류 발생]: {e}")

# --- Main App ---

async def main_menu():
    while True:
        guide_data = load_guide()
        updated_info = f" (최근 업데이트: {guide_data['updated_at']})" if guide_data else " (저장된 가이드 없음)"
        
        print("\n" + "="*45)
        print("🤖 Flex 에이전트 메인 메뉴")
        print("="*45)
        print("1. Flex 에이전트 수행" + updated_info)
        print("2. Flex 로그인 테스트 (Playwright)")
        print("3. 환경 변수 관리 (등록/수정/삭제)")
        print("4. 교육 마스터 시트 관리 (등록/수정/삭제)")
        print("5. 노션 업무 가이드 정적 업데이트")
        print("0. 프로그램 종료")
        print("="*45)
        
        choice = await ainput("입력: ")
        choice = choice.strip().lower()
        
        if choice == '1':
            config = load_config()
            if not guide_data:
                print("\n[알림] 저장된 가이드가 없습니다. 먼저 4번 메뉴로 업데이트해 주세요.")
                continue
            
            # 표시될 템플릿 목록 (사용자 요청에 따른 세분화 및 순서)
            display_templates = [
                "[품의서]",
                "[계약서 등 검토 · 승인] 교육 용역",
                "[계약서 등 검토 · 승인] 강사 용역",
                "[정기-기타/사업소득 자금집행요청서]",
                "[정기-자금집행요청서]",
            ]
            
            print("\n--- 수행할 워크플로우 양식을 선택해 주세요 ---")
            for i, t in enumerate(display_templates, 1):
                print(f"{i}. {t}")
            print(f"{len(display_templates)+1}. 메인 메뉴로 돌아가기")
            
            t_choice = await ainput("번호 선택: ")
            if t_choice.isdigit():
                idx = int(t_choice)
                if 1 <= idx <= len(display_templates):
                    selected_template = str(display_templates[idx-1])
                    # AI에게는 세분화된 이름을 전달하여 성격을 구분하게 함
                    await run_planning_flow(config, selected_template, str(guide_data.get("raw_text", "")))
                elif idx == len(display_templates) + 1:
                    print("\n[시스템] 메인 메뉴로 돌아갑니다.")
                    continue
                else:
                    print("\n[알림] 잘못된 선택입니다.")
            else:
                print("\n[알림] 숫자를 입력해 주세요.")
                
        elif choice == '2':
            config = load_config()
            if config.get("flex_id"):
                await flex_login_test(config)
            else:
                print("\n[경고] 환경 변수 설정이 필요합니다.")
                
        elif choice == '3':
            print("\n--- 환경 변수 관리 ---")
            print("1. 환경 변수 등록 (신규)")
            print("2. 환경 변수 수정 (기존 데이터 기반)")
            print("3. 환경 변수 삭제")
            print("e. 이전 메뉴로")
            sub_choice = await ainput("선택: ")
            sub_choice = sub_choice.strip().lower()
            
            if sub_choice == '1':
                gemini_key = await ainput("1. Gemini API Key: ")
                flex_id = await ainput("2. Flex ID (메일): ")
                flex_pw = await ainput("3. Flex 비밀번호: ")
                sa_path = await ainput("4. 서비스 계정 경로 (.json): ")
                notion_url = await ainput("5. 노션 가이드 URL: ")
                save_config(gemini_key.strip(), flex_id.strip(), flex_pw.strip(), sa_path.strip(), notion_url.strip())
            elif sub_choice == '2':
                config = load_config()
                if not config:
                    print("\n[알림] 저장된 환경 변수가 없습니다. 먼저 등록해 주세요.")
                    continue
                
                print("\n[수정 가이드] 변경할 항목만 입력하고, 유지하려면 그냥 엔터를 누르세요.")
                
                cur_gemini = str(config.get("gemini_api_key", ""))
                gemini_key = await ainput(f"1. Gemini API Key (현재: {cur_gemini}): ")
                gemini_key = gemini_key.strip() or cur_gemini
                
                cur_id = str(config.get("flex_id", ""))
                flex_id = await ainput(f"2. Flex ID (메일) (현재: {cur_id}): ")
                flex_id = flex_id.strip() or cur_id
                
                cur_pw = str(config.get("flex_pw", ""))
                flex_pw = await ainput(f"3. Flex 비밀번호 (현재: {'*'*len(cur_pw)}): ")
                flex_pw = flex_pw.strip() or cur_pw
                
                cur_sa = str(config.get("service_account_path", ""))
                sa_path = await ainput(f"4. 서비스 계정 경로 (.json) (현재: {cur_sa}): ")
                sa_path = sa_path.strip() or cur_sa
                
                cur_notion = str(config.get("notion_url", ""))
                notion_url = await ainput(f"5. 노션 가이드 URL (현재: {cur_notion}): ")
                notion_url = notion_url.strip() or cur_notion
                
                save_config(gemini_key, flex_id, flex_pw, sa_path, notion_url)
                
            elif sub_choice == '3':
                delete_config()
            elif sub_choice == 'e':
                continue
            else:
                print("\n[알림] 잘못된 선택입니다.")
            
        elif choice == '4':
            await manage_courses_menu()
            
        elif choice == '5':
            config = load_config()
            notion_url = config.get("notion_url")
            if not notion_url:
                print("\n[경고] 노션 URL이 설정되지 않았습니다. 3번 메뉴에서 먼저 등록해 주세요.")
                continue
            await fetch_notion_guide(notion_url)

        elif choice == '0':
            print("\n종료합니다.")
            break

async def flex_login_test(config):
    flex_id = config.get("flex_id")
    flex_pw = config.get("flex_pw")
    print("\n[시스템] 🕸️ Flex 로그인 테스트 시작...")
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()
        try:
            await page.goto("https://flex.team/auth/login")
            await page.fill('input[name="email"]', flex_id)
            await page.keyboard.press("Enter")
            await page.wait_for_selector('input[name="password"]')
            await page.fill('input[name="password"]', flex_pw)
            await page.keyboard.press("Enter")
            await page.wait_for_load_state("networkidle")
            
            # 로그인 성공 후 워크플로우 메뉴 진입 테스트 (사이드바 펼치기 로직 포함)
            try:
                nav_selector = 'nav >> text="워크플로우"'
                if not await page.locator(nav_selector).is_visible(timeout=2000):
                    print("[시스템] 🔍 사이드 바 펼치기 시도 중...")
                    expand_btn = page.locator('button.c-MZCiC.c-MZCiC-aVorO-active-false, button.c-MZCiC:has(svg)').first
                    if await expand_btn.is_visible(timeout=1000):
                        await expand_btn.click()
                        await asyncio.sleep(1.5)
                
                await page.wait_for_selector(nav_selector, timeout=5000)
                print("\n[시스템] 🎉 로그인 및 메뉴 접근 성공 확인!")
            except:
                print("\n[시스템] ⚠️ 로그인 성공했으나 메뉴 접근에 실패했습니다. (사이드바 확인 필요)")

            await ainput("\n[시스템] 엔터를 누르면 브라우저를 닫습니다...")
        except Exception as e:
            print(f"\n[시스템] ❌ 오류: {e}")
        finally:
            await browser.close()

if __name__ == "__main__":
    try:
        asyncio.run(main_menu())
    except KeyboardInterrupt:
        print("\n[시스템] 프로그램 종료.")
