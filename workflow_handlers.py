import re
import asyncio
import os
import requests  # type: ignore
import io
import json
from typing import Any, Dict, List, Optional, cast
from google.oauth2.service_account import Credentials  # type: ignore
from google.auth.transport.requests import Request  # type: ignore
from googleapiclient.discovery import build  # type: ignore
from googleapiclient.http import MediaIoBaseDownload  # type: ignore
async def reset_focus(page, template_name):
    """
    UI 포커스를 초기화합니다. 워크플로우를 닫지 않도록 안전한 곳을 클릭합니다.
    """
    await asyncio.sleep(0.5)
    try:
        # Escape는 창을 닫을 위험이 있어 제거하고, 안전한 위치(100, 100) 클릭만 유지
        await page.mouse.click(100, 100) 
    except:
        pass

# --- 공통 자동화 헬퍼 함수 ---
async def scroll_page(page, amount: int):
    """본문 컨테이너([data-part="scrollbar"] 등)를 스크롤합니다."""
    await page.evaluate("""(amount) => {
        const containerSelectors = [
            '[data-part="scrollbar"]',
            '.c-bjNJrx',
            '[data-scope="modal"][data-part="body"]',
            '.scroll-area',
            'main'
        ];
        let scrolled = false;
        for (const sel of containerSelectors) {
            const el = document.querySelector(sel);
            if (el && (el.scrollHeight > el.clientHeight)) {
                el.scrollBy(0, amount);
                scrolled = true;
                break;
            }
        }
        if (!scrolled) window.scrollBy(0, amount);
    }""", amount)
    await asyncio.sleep(0.5)

async def fill_text_field(page, label: str, value: Any):
    if not value: return
    print(f"   [진행] '{label}' 텍스트 입력 시도 중...")
    try:
        clean_val = str(value)
        if "금액" in label:
            clean_val = re.sub(r"[^0-9]", "", str(value))
        
        selectors = [
            f'input[placeholder*="{label}"]',
            f'textarea[placeholder*="{label}"]',
            f'text="{label}" >> xpath=following::*[self::input or self::textarea][not(@type="file")][1]',
            f'input:near(:text("{label}"))'
        ]
        
        found = False
        for i in range(4):
            for sel in selectors:
                for frame in page.frames:
                    try:
                        loc = frame.locator(sel).first
                        if await loc.is_visible(timeout=1000):
                            await loc.scroll_into_view_if_needed()
                            await loc.click(force=True)
                            await asyncio.sleep(0.3)
                            try:
                                await loc.fill(clean_val)
                            except:
                                await loc.focus()
                                await page.keyboard.type(clean_val, delay=30)
                            await loc.press("Enter")
                            print(f"   ✅ '{label}' 입력 완료: {clean_val}")
                            found = True
                            break
                    except: continue
                if found: break
            if found: break
            if i < 3: await scroll_page(page, 300)

        if not found:
            print(f"   ⚠️ '{label}' 필드를 찾지 못했습니다.")
    except Exception as e:
        print(f"   ⚠️ '{label}' 입력 중 오류: {e}")

async def internal_fill_date(page, label: str, value: Any):
    if not value: return
    try:
        date_str = re.sub(r"[^0-9-]", "", str(value))
        print(f"   [진행] '{label}' 처리 중... (값: {date_str})")
        # 1. 달력 버튼 클릭
        trigger_selectors = [
            f"xpath=//div[text()='{label}' or .//span[text()='{label}']]/following-sibling::div//button[contains(., '날짜') or contains(., '선택')]",
            f"//div[contains(., '{label}')]//button[contains(., '날짜') or contains(., '선택')]",
            f"div:has-text('{label}') >> button:has-text('선택')",
            f"div:has-text('{label}') >> button:has-text('날짜')",
            f"text='{label}' >> xpath=following::button[1]",
            f"button:near(:text('{label}'))"
        ]
        
        found = False
        for i in range(4):
            clicked_btn = False
            for ts in trigger_selectors:
                try:
                    btn = page.locator(ts).first
                    if await btn.is_visible(timeout=500):
                        await btn.scroll_into_view_if_needed()
                        await btn.click()
                        clicked_btn = True
                        break
                except:
                    continue

            if clicked_btn:
                await asyncio.sleep(0.5)
                date_input = page.locator('input[placeholder*="날짜 입력"]').first
                if await date_input.is_visible(timeout=3000):
                    await date_input.fill(date_str)
                    await page.keyboard.press("Enter")
                    await asyncio.sleep(0.3)
                    
                    # 팝업 닫기
                    exit_selectors = [f'span:text-is("{label}")', f'div:text-is("{label}")', f'text="{label}"']
                    for es in exit_selectors:
                        try:
                            loc = page.locator(es).first
                            if await loc.is_visible(timeout=1000):
                                await loc.click(force=True)
                                break
                        except: pass
                    found = True
                    break
            if i < 3: await scroll_page(page, 300)
        
        if not found:
            print(f"   ⚠️ '{label}' 달력 버튼을 찾지 못했습니다.")
    except Exception as e:
        print(f"   ⚠️ '{label}' 처리 중 오류: {e}")

async def select_list_field(page, label: str, target_val: str):
    if not target_val: return
    print(f"   [진행] '{label}' 선택 시도 중... (목표: {target_val})")
    try:
        field = None
        found_field = False
        for i in range(4):
            for frame in page.frames:
                try:
                    label_loc = frame.locator(f'text="{label}"').first
                    if await label_loc.is_visible(timeout=500):
                        row = label_loc.locator('xpath=./ancestor::div[contains(@class, "flex") or contains(@style, "display: flex") or contains(@role, "row") or contains(@class, "Row")][1]')
                        if not await row.is_visible(timeout=300):
                            row = label_loc.locator('xpath=./parent::div/parent::div')
                        if await row.is_visible(timeout=500):
                            field = row.locator('div:has-text("옵션"), [role="button"], div:has-text("선택해")').last
                            if await field.is_visible(timeout=500):
                                found_field = True
                                break
                except: continue
            if found_field: break
            if i < 3: await scroll_page(page, 300)

        if not field or not found_field:
            # 폴백 탐색
            field_selectors = [
                f'div:has-text("{label}") >> div:has-text("옵션")',
                f'div:has-text("{label}") >> [role="button"]'
            ]
            for sel in field_selectors:
                for frame in page.frames:
                    try:
                        loc = frame.locator(sel).last
                        if await loc.is_visible(timeout=500):
                            field = loc
                            found_field = True
                            break
                    except: continue
                if found_field: break

        if not field:
            print(f"   ⚠️ '{label}' 필드를 찾지 못했습니다.")
            return

        await field.scroll_into_view_if_needed()
        current_text = await field.inner_text()
        if target_val in current_text:
            print(f"      ℹ️ '{label}'이(가) 이미 '{target_val}'로 설정되어 있습니다.")
            return

        await field.click(force=True)
        await asyncio.sleep(1.0)
        
        option_selectors = [
            f'div[role="option"]:text-is("{target_val}")', 
            f'li:text-is("{target_val}")', 
            f'span:text-is("{target_val}")',
            f'div:text-is("{target_val}")',
            f'button:text-is("{target_val}")'
        ]
        
        option_found = False
        for o_sel in option_selectors:
            for frame in page.frames:
                try:
                    opt_loc = frame.locator(o_sel).last
                    if await opt_loc.is_visible(timeout=1000):
                        await opt_loc.click(force=True)
                        print(f"   ✅ '{label}' 선택 완료: {target_val}")
                        option_found = True
                        break
                except: continue
            if option_found: break

        if option_found:
            await asyncio.sleep(0.5)
            try: await label_loc.click(force=True)
            except: await page.mouse.click(100, 100)
        else:
            print(f"   ⚠️ '{label}'에서 '{target_val}' 옵션을 찾지 못했습니다.")
            await page.keyboard.press("Escape")
    except Exception as e:
        print(f"   ⚠️ '{label}' 선택 중 오류: {e}")

async def download_file(url, download_dir, index, config):
    """
    URL에서 파일을 다운로드하며, 원본 파일명을 최대한 유지합니다.
    구글 드라이브 링크의 경우 서비스 계정을 통한 API를 사용하여 권한 문제를 해결합니다.
    """
    try:
        # 1. 구글 드라이브 링크 판별 및 추출
        file_id = None
        if "drive.google.com" in url or "docs.google.com" in url:
            match = re.search(r"/(?:file|spreadsheets|document|presentation)/d/([a-zA-Z0-9-_]+)", url)
            if match:
                file_id = match.group(1)
            elif "id=" in url:
                id_match = re.search(r"id=([a-zA-Z0-9-_]+)", url)
                if id_match:
                    file_id = id_match.group(1)

        # 2. 구글 드라이브 API를 통한 다운로드 시도
        sa_path = config.get("service_account_path")
        if file_id:
            if sa_path and os.path.exists(sa_path):
                print(f"   [디버그] 드라이브 API 사용 시작 (ID: {file_id})")
                try:
                    scopes = ['https://www.googleapis.com/auth/drive.readonly']
                    creds = Credentials.from_service_account_file(sa_path, scopes=scopes)
                    service = build('drive', 'v3', credentials=creds)
                    
                    # 파일 메타데이터 가져오기 (이름 및 MIME 타입 포함)
                    file_metadata = service.files().get(
                        fileId=file_id, 
                        fields='name, mimeType',
                        supportsAllDrives=True
                    ).execute()
                    
                    filename = file_metadata.get('name', f"attachment_{index}")
                    mime_type = file_metadata.get('mimeType', '')
                    
                    # 구글 문서/시트 등의 경우 내보내기(Export) 필요
                    is_google_type = any(t in mime_type for t in ['document', 'spreadsheet', 'presentation'])
                    
                    if is_google_type:
                        print(f"   [디버그] 구글 전용 문서 탐색됨 ({mime_type}), PDF로 내보내기 시도...")
                        if not filename.endswith(".pdf"):
                            filename += ".pdf"
                        request = service.files().export_media(
                            fileId=file_id,
                            mimeType='application/pdf'
                        )
                    else:
                        request = service.files().get_media(
                            fileId=file_id,
                            supportsAllDrives=True
                        )
                    
                    fh = io.BytesIO()
                    downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    while done is False:
                        status, done = downloader.next_chunk()
                    
                    filename = re.sub(r'[\\/*?:"<>|]', "_", filename)
                    target_path = os.path.join(download_dir, filename)
                    
                    with open(target_path, 'wb') as f:
                        f.write(fh.getvalue())
                        
                    print(f"   [디버그] API 다운로드/내보내기 성공: {filename}")
                    return target_path
                except Exception as api_err:
                    if "cannotExportFile" in str(api_err) or "403" in str(api_err):
                        # 서비스 계정 이메일 추출
                        sa_email = "미확인"
                        try:
                            with open(sa_path, 'r') as f:
                                sa_email = json.load(f).get("client_email", "미확인")
                        except: pass
                        
                        print(f"   ❌ [권한/설정 오류] 파일을 내보낼(Export) 수 없습니다. (ID: {file_id})")
                        print(f"   💡 원인: 파일의 '다운로드/인쇄/복사 제한' 설정이 켜져 있거나 권한이 부족합니다.")
                        print(f"   💡 해결: 구글 문서의 [공유] -> [설정(톱니바퀴)] -> '댓글작성자 및 뷰어의 다운로드, 인쇄, 복사 옵션을 사용 중지합니다' 체크를 해제해 주세요.")
                        print(f"   💡 또는 서비스 계정({sa_email})에 '편집자' 권한을 부여해 주세요.")
                    else:
                        print(f"   ⚠️ 드라이브 API 오류: {api_err}")
            else:
                print("   ⚠️ 서비스 계정 경로가 없거나 파일이 존재하지 않아 일반 다운로드를 시도합니다.")

        # 3. 일반 다운로드 (기타 URL 또는 API 실패 시)
        session = requests.Session()
        actual_url = url
        
        # 구글 드라이브 링크 변환 (export=download fallback용)
        if file_id:
            actual_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        
        # Google URL인 경우 서비스 계정 토큰 추가 (401 Unauthorized 방지)
        if sa_path and os.path.exists(sa_path) and ("google.com" in url):
            try:
                scopes = ['https://www.googleapis.com/auth/drive.readonly']
                creds = Credentials.from_service_account_file(sa_path, scopes=scopes)
                creds.refresh(Request())
                session.headers.update({"Authorization": f"Bearer {creds.token}"})
                print("   [디버그] Authorized session created with Bearer token")
            except Exception as auth_err:
                print(f"   ⚠️ 토큰 획득 실패 (일반 다운로드 시도): {auth_err}")
                
        response = session.get(actual_url, stream=True, timeout=30)
        
        confirm_token = None
        for key, value in response.cookies.items():
            if key.startswith('download_warning'):
                confirm_token = value
                break
        
        if confirm_token:
            actual_url = f"{actual_url}&confirm={confirm_token}"
            response = session.get(actual_url, stream=True, timeout=30)
            
        response.raise_for_status()
        
        content_type = response.headers.get("Content-Type", "unknown")
        print(f"   [디버그] 일반 다운로드 타입: {content_type}")

        filename = None
        content_disposition = response.headers.get("Content-Disposition")
        if content_disposition and "filename=" in content_disposition:
            fname_match = re.findall("filename=\"?(.+?)\"?(;|$)", content_disposition)
            if fname_match:
                filename = fname_match[0][0]
        
        if not filename:
            path_part = url.split("?")[0]
            filename = os.path.basename(path_part)
            
        if not filename or "." not in filename:
            ext = ".pdf"
            if "image/png" in content_type: ext = ".png"
            elif "image/jpeg" in content_type: ext = ".jpg"
            filename = f"attachment_{index}{ext}"
            
        filename = re.sub(r'[\\/*?:"<>|]', "_", filename)
        target_path = os.path.join(download_dir, filename)
        
        with open(target_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk: f.write(chunk)
        
        return target_path
    except Exception as e:
        print(f"   [오류] 최종 다운로드 실패 ({url}): {e}")
        return None

def rgb_to_hex(rgb):
    """
    Sheets API의 RGB(0-1)를 Hex(#RRGGBB)로 변환합니다.
    """
    if not rgb: return "#ffffff"
    r = int(rgb.get('red', 0) * 255)
    g = int(rgb.get('green', 0) * 255)
    b = int(rgb.get('blue', 0) * 255)
    return f"#{r:02x}{g:02x}{b:02x}"

def json_to_html_table(data):
    """
    리스트의 리스트 또는 리치 데이터 객체를 HTML <table> 태그로 변환합니다.
    GAS의 buildHtmlTableAuto 로직을 재현합니다.
    """
    if not data: return ""
    
    # 1. 데이터 형식 판별 (리스트 vs 리치 객체)
    is_rich = False
    start_row = 1
    values: List[List[Any]] = []
    backgrounds: List[List[Any]] = []
    font_weights: List[List[str]] = []
    merges: List[Dict[str, Any]] = []
    column_widths: List[int] = []

    if isinstance(data, dict) and "values" in data:
        is_rich = True
        values = data.get("values", [])
        backgrounds = data.get("backgrounds", [])
        font_weights = data.get("fontWeights", [])
        merges = data.get("merges", [])
        column_widths = data.get("columnWidths", [])
        start_row = data.get("startRow", 1)
    elif isinstance(data, list):
        values = data
    else:
        # JSON 문자열인 경우 파싱 시도
        try:
            parsed = json.loads(data)
            return json_to_html_table(parsed)
        except: pass
        return ""

    if not values: return ""

    merge_map: dict = {}
    row_offset: int = start_row - 1 # API 인덱스 -> values 인덱스 변환용
    
    for m in merges:
        m_dict = cast(Dict[str, Any], m)
        orig_start_r_val = m_dict.get('startRowIndex')
        orig_start_r: int = int(orig_start_r_val) if orig_start_r_val is not None else 0
        
        orig_end_r_val = m_dict.get('endRowIndex')
        orig_end_r: int = int(orig_end_r_val) if orig_end_r_val is not None else (orig_start_r + 1)
        
        # 현재 values 범위 밖(예: 1행)에 대한 머지 정보는 스킵
        if orig_end_r <= row_offset:
            continue
            
        # values 인덱스로 변환
        start_r: int = max(0, orig_start_r - row_offset) # type: ignore
        end_r: int = orig_end_r - row_offset # type: ignore
        
        orig_start_c_val = m_dict.get('startColumnIndex')
        start_c: int = int(orig_start_c_val) if orig_start_c_val is not None else 0
        
        orig_end_c_val = m_dict.get('endColumnIndex')
        end_c: int = int(orig_end_c_val) if orig_end_c_val is not None else (start_c + 1)
        
        rowspan: int = end_r - start_r
        # 만약 머지가 values 시작 위치를 넘어온 경우 rowspan 조정
        if orig_start_r < row_offset:
             # rowspan = orig_end_r - row_offset (이미 계산됨)
             pass
             
        colspan: int = end_c - start_c
        
        merge_map[f"{start_r},{start_c}"] = {
            "rowspan": rowspan,
            "colspan": colspan
        }
        
        for r_m in range(start_r, end_r):
            for c_m in range(start_c, end_c):
                if r_m != start_r or c_m != start_c:
                    merge_map[f"{r_m},{c_m}"] = "SKIP"

    # 3. 테이블 생성 시작
    # table-layout: fixed를 사용하여 열 너비를 강제합니다.
    html_parts: List[str] = ['<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-size: 12px; width: auto; table-layout: fixed;">']
    
    # 열 너비 설정 (<colgroup> 추가)
    if is_rich and column_widths:
        html_parts.append('<colgroup>')
        for w in column_widths:
            html_parts.append(f'<col style="width: {w}px;">')
        html_parts.append('</colgroup>')
    
    for r, row in enumerate(values):
        # [조건] A열(인덱스 0)이 비어있고, 병합의 연장선이 아니면 행 스킵
        row_list = cast(List[Any], row)
        col_a_val = str(row_list[0]).strip() if len(row_list) > 0 else ""
        
        is_part_of_merge = (merge_map.get(f"{r},0") == "SKIP") # type: ignore
        
        if not col_a_val and not is_part_of_merge:
            continue
            
        html_parts.append('<tr>')
        for c, cell in enumerate(row_list):
            key = f"{r},{c}"
            if merge_map.get(key) == "SKIP": # type: ignore
                continue
                
            merge_info = merge_map.get(key, {}) # type: ignore
            m_info: dict = merge_info if isinstance(merge_info, dict) else {}
            
            rowspan = int(m_info.get("rowspan", 1)) # type: ignore
            colspan = int(m_info.get("colspan", 1)) # type: ignore
            
            # 스타일 추출
            bg_color: str = "#ffffff"
            if is_rich and r < len(backgrounds) and c < len(cast(List[Any], backgrounds[r])):
                row_bg = cast(List[Any], backgrounds[r])
                bg_color = rgb_to_hex(row_bg[c])
                
            f_weight: str = "normal"
            if is_rich and r < len(font_weights) and c < len(cast(List[Any], font_weights[r])):
                row_fw = cast(List[str], font_weights[r])
                f_weight = row_fw[c]
            
            # 셀 렌더링
            style = f"background: {bg_color}; font-weight: {f_weight}; text-align: center; vertical-align: middle; padding: 6px 10px; white-space: nowrap;"
            html_parts.append(f'<td rowspan="{rowspan}" colspan="{colspan}" style="{style}">{cell}</td>')
            
        html_parts.append('</tr>')
        
    html_parts.append('</table>')
    return "".join(html_parts)

async def handle_pumiseo(page, plan, template_name, config):
    """
    [품의서] 양식 전용 핸들러 (제목, 시작일, 예상 비용, 파일 첨부 자동 입력)
    """
    print(f"[핸들러] 📝 '{template_name}' 전용 로직 수행 중...")
    
    # 1. 제목 입력
    title_value = plan.get("title")
    if title_value:
        title_selectors = [
            'input[placeholder*="제목"]',
            'textarea[placeholder*="제목"]',
            '[aria-label*="제목"]',
            'input:near(:text("제목"))'
        ]
        for ts in title_selectors:
            try:
                locator = page.locator(ts).first
                if await locator.is_visible(timeout=3000):
                    await locator.click()
                    await locator.fill(title_value)
                    print(f"   ✅ '제목' 입력 완료")
                    break
            except: continue

    # 2. 시작일 / 종료일 입력
    for date_key in ["시작일", "종료일"]:
        date_value = plan.get("fields", {}).get(date_key)
        if date_value:
            clean_date = re.sub(r"[^0-9-]", "", str(date_value))
            # 시작일/종료일 버튼을 정확히 타겟팅하기 위해 구체적인 셀렉터를 상단에 배치
            trigger_selectors = [
                f"div:has-text('{date_key}') >> button:has-text('선택')", # 달력 버튼이 '선택'이라는 텍스트를 가질 때
                f"div:has-text('{date_key}') >> button:has-text('날짜')", # 달력 버튼이 '날짜'라는 텍스트를 가질 때
                f"button:near(:text('{date_key}'))",
                f"text='{date_key}' >> xpath=following::button[1]",
                f"input:near(:text('{date_key}'))",
                f"button:has-text('{date_key}')",
                f"div:has-text('{date_key}') >> xpath=following::button[1]"
            ]
            
            for ts in trigger_selectors:
                try:
                    trigger = page.locator(ts).first
                    if await trigger.is_visible(timeout=3000):
                        await trigger.click(force=True)
                        await asyncio.sleep(1.5)
                        
                        input_selectors = ['input[placeholder*="날짜 입력"]', 'input[placeholder*="YYYY-MM-DD"]']
                        for ins in input_selectors:
                            date_input = page.locator(ins).first
                            if await date_input.is_visible(timeout=2000):
                                await date_input.click()
                                await date_input.fill(clean_date)
                                await asyncio.sleep(0.5)
                                await date_input.press("Enter")
                                await asyncio.sleep(0.5)
                                
                                # 사용자가 '시작일' 또는 '종료일' 글자를 직접 눌러서 캘린더를 닫는 동작을 모방
                                try:
                                    anchor_label = page.locator(f'text="{date_key}"').first
                                    await anchor_label.click(force=True)
                                except:
                                    pass
                                await asyncio.sleep(0.5)
                                
                                print(f"   ✅ '{date_key}' 입력 완료: {clean_date}")
                                break
                        break
                except: continue


    # 3. 예상 비용 / 예상 매출 입력
    cost_mapping = {
        "예상 비용": ["예상 비용", "총 예상 지출 비용", "예상 지출", "지출 비용", "금액"],
        "예상 매출": ["예상 매출", "총 예상 매출", "매출 금액", "매출"]
    }
    
    fields = plan.get("fields", {})
    for label, aliases in cost_mapping.items():
        found_key = next((k for k in aliases if fields.get(k)), label)
        cost_value = fields.get(found_key)
        
        if cost_value:
            clean_cost = re.sub(r"[^0-9]", "", str(cost_value))
            cost_selectors = [
                f'textarea[placeholder*="금액을 입력해"]:near(:text("{found_key}"))',
                f'input[placeholder*="금액을 입력해"]:near(:text("{found_key}"))',
                f'text="{found_key}" >> xpath=following::*[self::input or self::textarea][not(@type="file")][1]',
                f'textarea:near(:text("{found_key}"))',
                f'input:not([type="file"]):near(:text("{found_key}"))'
            ]
            
            found_cost = False
            for cs in cost_selectors:
                try:
                    for frame in page.frames:
                        try:
                            cost_input = frame.locator(cs).first
                            if await cost_input.is_visible(timeout=2000):
                                await cost_input.scroll_into_view_if_needed()
                                await cost_input.click(force=True, timeout=3000)
                                await asyncio.sleep(0.3)
                                try:
                                    await cost_input.fill(clean_cost, timeout=3000)
                                except:
                                    await cost_input.focus()
                                    await frame.keyboard.type(clean_cost, delay=50)
                                await cost_input.press("Enter")
                                print(f"   ✅ '{found_key}' 입력 완료: {clean_cost}")
                                found_cost = True
                                break
                        except: continue
                    if found_cost: break
                except: continue

    # 4. 본문 내용 및 표 입력
    content_value = fields.get("본문 내용", "")
    table_data = plan.get("table_data")
    
    # 줄 바꿈을 기준으로 나누고 빈 줄은 제거한 뒤, 제목과 본문 사이의 간격을 최적화합니다.
    lines = [l.strip() for l in str(content_value).split("\n") if l.strip()]
    formatted_segments = []
    
    for line in lines:
        # 섹션 제목 형식(I., II., 1., (1) 등)인지 확인
        if re.match(r"^(?:[IVX]+\.|\d+[\.\)]|\(\d+\))\s*.*", line):
            # 제목 행이면 볼드 처리
            formatted_segments.append(f"<b>{line}</b>")
        else:
            formatted_segments.append(line)
            
    # 제목 행 앞에만 빈 줄을 추가하여 문단 간격을 확보합니다.
    final_lines = []
    for i, seg in enumerate(formatted_segments):
        if i > 0 and seg.startswith("<b>"):
            final_lines.append("") # 제목 앞에 빈 줄 추가
        final_lines.append(seg)
        
    formatted_text = "<br>".join(final_lines)
    # 품의서도 줄 간격을 촘촘하지 않게 1.8로 설정합니다.
    final_content = f'<div style="line-height: 1.8;">{formatted_text}</div>'
    if table_data:
        table_html = json_to_html_table(table_data)
        # 테이블 섹션 제목인 "III. 예상 비용" 앞에도 빈 줄이 생기도록 조정 (이미 formatted_text 끝에 빈 줄이 없을 수 있으므로)
        final_content += f'<div style="line-height: 1.8;"><br><b>III. 예상 비용</b><br>{table_html}</div>'
        
    if final_content.strip():
        content_selectors = [
            'div[contenteditable="true"]',
            'textarea[placeholder*="내용을 입력해"]',
            '[aria-label*="본문"]',
            'div[class*="editor"]'
        ]
        
        for cs in content_selectors:
            try:
                content_input = page.locator(cs).first
                if await content_input.is_visible(timeout=3000):
                    await content_input.click()
                    # 전체 선택 후 삭제
                    await page.keyboard.press("Meta+a" if os.name == 'posix' else "Control+a")
                    await page.keyboard.press("Backspace")
                    await asyncio.sleep(0.5)
                    
                    # HTML 주입 (Flex 에디터가 rich text를 지원하므로 clipboard나 innerHTML을 시도)
                    # 여기서는 에디터의 innerHTML을 직접 수정하거나, evaluate를 통해 contenteditable 영역에 삽입
                    if await content_input.evaluate("el => el.isContentEditable"):
                        await content_input.evaluate(f"(el) => {{ el.innerHTML = `{final_content}`; }}")
                    else:
                        await content_input.fill(final_content)
                    print(f"   ✅ '본문 내용 및 표' 입력 완료")
                    break
            except: continue

    # 5. 파일 다운로드 및 업로드 통합
    attachments = plan.get("attachments", [])
    if attachments:
        # 다운로드 경로 설정: 현재 파일 위치의 flex_attachments 폴더
        current_dir = os.path.dirname(os.path.abspath(__file__))
        download_dir = os.path.join(current_dir, "flex_attachments")
        if not os.path.exists(download_dir):
            os.makedirs(download_dir)
            
        print(f"[핸들러] 📎 파일 처리 중... (대상: {len(attachments)}개)")
        downloaded_paths = []
        for i, url in enumerate(attachments):
            path = await download_file(url, download_dir, i, config)
            if path:
                downloaded_paths.append(path)
                print(f"   📥 다운로드 완료: {os.path.basename(path)}")

        if downloaded_paths:
            try:
                # 5. Flex 업로드 필드 찾기 (모달/사이드바 로딩 대기 및 프레임 탐색 강화)
                print("   ⏳ 업로드 필드 로딩 대기 중...")
                file_input_selector = 'input[type="file"]'
                
                # 최대 10초간 업로드 필드가 나타날 때까지 대기
                target_input = None
                for _ in range(20): # 0.5초 * 20 = 10초
                    # 모든 프레임에서 파일 입력 필드 검색
                    for frame in page.frames:
                        try:
                            loc = frame.locator(file_input_selector)
                            if await loc.count() > 0:
                                # 가시성 확인 (숨겨진 input일 수 있으므로 존재 여부만 확인하거나 조상 요소 가시성 확인)
                                target_input = loc.first
                                break
                        except: continue
                    if target_input: break
                    await asyncio.sleep(0.5)

                if target_input is not None:
                    await target_input.set_input_files(downloaded_paths) # type: ignore
                    print(f"   ✅ 파일 업로드 완료 ({len(downloaded_paths)}개)")
                    await asyncio.sleep(1)
                else:
                    print("   ⚠️ 업로드 필드를 찾을 수 없습니다. 양식이 완전히 로드되었는지 확인해 주세요.")
            except Exception as e:
                print(f"   ❌ 업로드 중 오류: {e}")
            finally:
                # 사용자의 요청에 따라 파일을 삭제하지 않고 보존합니다.
                pass

async def fetch_sheet_data_range(sheet_url, sheet_name, range_name, config):
    """
    구글 시트의 특정 범위를 가져옵니다.
    """
    sa_path = config.get("service_account_path")
    if not (sa_path and os.path.exists(sa_path)):
        print("   ❌ 서비스 계정 파일이 없어 데이터를 가져올 수 없습니다.")
        return None

    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        creds = Credentials.from_service_account_file(sa_path, scopes=scopes)
        service = build('sheets', 'v4', credentials=creds)

        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
        if not match: return None
        ss_id = match.group(1)
        
        full_range = f"'{sheet_name}'!{range_name}"
        result = service.spreadsheets().values().get(spreadsheetId=ss_id, range=full_range).execute()
        return result.get('values', [])
    except Exception as e:
        print(f"   ❌ 데이터 가져오기 오류 ({range_name}): {e}")
        return None

# 핸들러 등록 대장
HANDLER_REGISTRY = {
    "[품의서]": handle_pumiseo,
}

async def dispatch_workflow(template_name, page, plan, config):
    """
    템플릿 이름에 맞는 핸들러를 찾아 실행합니다.
    """
    if template_name == "[정기-기타/사업소득 자금집행요청서]":
        from workflow_business_income import handle_business_income  # type: ignore
        await handle_business_income(page, plan, template_name, config)
        return True

    if template_name == "[정기-자금집행요청서]":
        from workflow_general_funding import handle_general_funding  # type: ignore
        await handle_general_funding(page, plan, template_name, config)
        return True

    if "[계약서 등 검토 · 승인] 교육 용역" in template_name:
        from workflow_education_services import handle_education_services  # type: ignore
        await handle_education_services(page, plan, template_name, config)
        return True

    if "[계약서 등 검토 · 승인] 강사 용역" in template_name:
        from workflow_contract_instructor import handle_contract_instructor  # type: ignore
        await handle_contract_instructor(page, plan, template_name, config)
        return True

    handler = HANDLER_REGISTRY.get(template_name)
    if handler:
        await handler(page, plan, template_name, config)
        return True
    return False
