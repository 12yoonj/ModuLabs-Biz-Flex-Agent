import re
import asyncio
import os
import io
from datetime import datetime
from typing import List, Any, cast
from google.oauth2.service_account import Credentials  # type: ignore
from google.auth.transport.requests import Request  # type: ignore
from googleapiclient.discovery import build  # type: ignore
from googleapiclient.http import MediaIoBaseDownload  # type: ignore

from workflow_handlers import download_file, json_to_html_table, fetch_sheet_data_range # type: ignore

async def handle_general_funding(page, plan: dict, template_name: str, config: dict):
    """
    [정기-자금집행요청서] 양식 전용 핸들러
    """
    print(f"[핸들러] 📝 '{template_name}' 전용 로직 수행 중...")
    
    # 0. 공통 데이터 추출
    spreadsheet_url = plan.get("spreadsheet_url")
    sheet_name = "[정산] 일반 자금집행요청서"
    
    # 0-1. 시트에서 필수 데이터 추출 (B5:E, D4, 개요!B3, B4, F6)
    real_table_data: List[List[Any]] = [] # 헤더 포함 (B5:E)
    total_amount_val = "-"
    intro_context = ""
    client_name = ""
    course_name = ""
    summary_f6 = ""
    
    if spreadsheet_url:
        print(f"   📊 시트에서 직접 데이터 추출 중...")
        # (1) 본문 표 및 데이터 행 파악을 위한 범위 (B5:E)
        res = await fetch_sheet_data_range(spreadsheet_url, sheet_name, "B5:E", config)
        if isinstance(res, list):
            real_table_data = res
        
        if not real_table_data:
            print("   ⚠️ [디버그] B5:E 데이터를 가져오지 못했습니다. 시트 권한이나 범위를 확인해주세요.")
        else:
            print(f"   [디버그] 가져온 데이터 행 수: {len(real_table_data)}")

        # (2) 총 이체 금액 (D4)
        d4_res = await fetch_sheet_data_range(spreadsheet_url, sheet_name, "D4", config)
        if isinstance(d4_res, list) and len(d4_res) > 0 and len(d4_res[0]) > 0:
            total_amount_val = str(d4_res[0][0])
            print(f"   [디버그] 총 이체 금액(D4): {total_amount_val}")

        # (3) 개요 정보 (B3, B4)
        intro_res = await fetch_sheet_data_range(spreadsheet_url, "개요", "B3:B4", config)
        if isinstance(intro_res, list) and len(intro_res) >= 2:
            client_name = str(intro_res[0][0]) if intro_res[0] else "미확인기업"
            course_name = str(intro_res[1][0]) if intro_res[1] else "미확인과정"
            intro_context = f"({client_name}) {course_name} 강연료 집행 건"
            print(f"   [디버그] 개요: {intro_context}")
            
        # (4) 요약 정보 (F6)
        f6_res = await fetch_sheet_data_range(spreadsheet_url, sheet_name, "F6", config)
        if isinstance(f6_res, list) and len(f6_res) > 0 and len(f6_res[0]) > 0:
            summary_f6 = str(f6_res[0][0])
            print(f"   [디버그] 요약(F6): {summary_f6}")

    # 데이터 행 수 계산 (헤더인 5행 제외, 6행부터)
    data_rows = list(real_table_data)[1:] if len(real_table_data) > 1 else [] # type: ignore
    # 빈 행 필터링 (B열(은행명) 기준 데이터 존재 여부 확인, B열은 인덱스 0)
    valid_data_rows: List[List[Any]] = []
    for r in data_rows:
        if isinstance(r, list) and len(r) > 0:
            val_b = str(r[0]).strip() if r[0] is not None else ""
            if val_b:
                padded_row = list(r)
                while len(padded_row) < 4:
                    padded_row.append("-")
                valid_data_rows.append(padded_row)
    
    row_count = len(valid_data_rows)
    print(f"   [디버그] 유효 데이터 행 수: {row_count}")

    # 1. 제목 입력
    # GAS 기준: [정기-자금집행요청서] {clientName}_{courseName}_강연료 비용 집행
    suggested_title = f"[정기-자금집행요청서] {client_name}_{course_name}_강연료 비용 집행"
    # AI 계획서에 잘못된 말머리가 포함될 수 있으므로 직접 생성한 제목을 우선 사용함
    title_value = suggested_title
    
    if title_value:
        print(f"   [진행] '제목' 입력 시도 중: {title_value}")
        title_selectors = [
            'input[placeholder*="제목"]',
            'textarea[placeholder*="제목"]',
            '[aria-label*="제목"]',
            'div[contenteditable="true"]:near(:text("제목"))', # 제목 근처의 에디터
            'xpath=//div[contains(@class, "Title")]', # 제목 클래스 포함 div
            'text="제목" >> xpath=following::input[1]'
        ]
        
        found_title = False
        for ts in title_selectors:
            try:
                locator = page.locator(ts).first
                if await locator.is_visible(timeout=2000):
                    await locator.click()
                    await asyncio.sleep(0.5)
                    # 전체 선택 후 입력
                    await page.keyboard.press("Meta+a" if os.name == 'posix' else "Control+a")
                    await page.keyboard.press("Backspace")
                    await locator.fill(title_value)
                    await page.keyboard.press("Enter")
                    print(f"   ✅ '제목' 입력 완료")
                    found_title = True
                    break
            except: continue
            
        if not found_title:
            # 최종 수단: 화면 상단에서 '제목' 혹은 '양식 이름' 등 클릭 시도 (Flex 특성상 텍스트 클릭 시 input으로 변함)
            try:
                # 상단 h1 또는 제목 관련 텍스트 클릭 시도
                header_title = page.locator('h1, [class*="Header"] div, [class*="title"]').first
                if await header_title.is_visible(timeout=1000):
                    await header_title.click()
                    await asyncio.sleep(0.5)
                    await page.keyboard.press("Meta+a" if os.name == 'posix' else "Control+a")
                    await page.keyboard.press("Backspace")
                    await page.keyboard.type(title_value)
                    await page.keyboard.press("Enter")
                    print(f"   ✅ '제목' (상단 헤더 방식) 입력 시도 완료")
            except: pass

    # 2. 개별 필드 입력 (GAS 로직에 맞춰 최적화)
    # GAS에서는 B6, C6에서 정보를 가져오고 E열(6~25)에서 유니크한 예금주를 뽑음.
    
    # 모든 예금주 수집 (E열은 index 3)
    names = [str(r[3]).strip() for r in valid_data_rows if len(r) > 3 and str(r[3]).strip()]
    unique_names = list(set(names))
    
    bank_val = "-"
    account_val = "-"
    holder_val = "-"
    
    if len(unique_names) == 1:
        holder_val = unique_names[0]
        # 예금주가 하나라면 은행과 계좌도 첫 번째 유효 행에서 가져옴
        bank_val = str(valid_data_rows[0][0]).strip() if len(valid_data_rows[0]) > 0 else "-"
        account_val = str(valid_data_rows[0][1]).strip() if len(valid_data_rows[0]) > 1 else "-"
    elif len(unique_names) > 1:
        # 여러 명일 경우 본문 기재로 처리
        bank_val = "본문 내 기재"
        account_val = "본문 내 기재"
        holder_val = "본문 내 기재"

    field_configs = [
        {"label": "기타(은행명)", "val": bank_val, "aliases": ["기타(은행명)", "은행", "은행명", "입금 은행", "입금은행"]},
        {"label": "계좌 번호", "val": account_val, "aliases": ["계좌 번호", "계좌번호", "계좌", "입금 계좌", "입금계좌번호"]},
        {"label": "총 이체 금액", "val": total_amount_val, "aliases": ["총 이체 금액", "금액", "이체 금액", "예상 비용"]},
        {"label": "예금주", "val": holder_val, "aliases": ["예금주", "예금주명", "성명", "입금 성명"]}
    ]

    for cfg in field_configs:
        label = cfg["label"]
        val = cfg["val"] or "-"
        aliases = cfg["aliases"]
        clean_val = str(val)
        if label == "총 이체 금액" and val and val != "-":
            clean_val = re.sub(r"[^0-9]", "", str(val))

        all_selectors = []
        for alias in aliases:
            all_selectors.extend([
                f'input[placeholder*="{alias}"]',
                f'textarea[placeholder*="{alias}"]',
                f'text="{alias}" >> xpath=following::*[self::input or self::textarea][not(@type="file")][1]'
            ])
        
        found = False
        for sel in all_selectors:
            try:
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
                                await frame.keyboard.type(clean_val, delay=50)
                            await loc.press("Enter")
                            print(f"   ✅ '{label}' 입력 완료: {clean_val}")
                            found = True
                            break
                    except: continue
                if found: break
            except: continue
        
        if not found:
            print(f"   ⚠️ '{label}' 필드 탐색 실패")

    # 3. [이체요청일] 및 [코스트 센터] (workflow_business_income.py 로직 재사용)
    # (여기서는 편의상 business_income의 코드를 직접 가져옴)
    try:
        from datetime import date, timedelta
        import calendar
        def get_business_day(d):
            holidays = {date(2026, 1, 1), date(2026, 2, 16), date(2026, 2, 17), date(2026, 2, 18)} # 간단히 일부만
            curr = d
            while curr.weekday() >= 5 or curr in holidays:
                curr += timedelta(days=1)
            return curr
        today = date.today()
        candidates = []
        for stage in [0, 1]:
            year, month = today.year, today.month + stage
            if month > 12: month -= 12; year += 1
            last_day = calendar.monthrange(year, month)[1]
            for d in [10, 15, 20, 25, last_day]:
                candidates.append(date(year, month, d))
        valid_biz_days = sorted(list(set(get_business_day(c) for c in candidates if c > today)))
        target_date = valid_biz_days[1] if len(valid_biz_days) > 1 else valid_biz_days[0]
        date_str = target_date.strftime("%Y-%m-%d")
        
        print(f"   📅 '이체요청일' 계산: {date_str}")
        calendar_btn = page.locator("div:has-text('이체요청일') >> button:has-text('날짜')").first
        if await calendar_btn.is_visible(timeout=3000):
            await calendar_btn.click()
            await asyncio.sleep(0.5)
            date_input = page.locator('input[placeholder*="날짜 입력"]').first
            if await date_input.is_visible(timeout=3000):
                await date_input.fill(date_str)
                await page.keyboard.press("Enter")
                await asyncio.sleep(0.5)
                # 닫기 시도
                try: await page.locator('text="이체요청일"').first.click(force=True)
                except: await page.mouse.click(10, 10)
    except Exception as e:
        print(f"   ⚠️ '이체요청일' 오류: {e}")

    # 코스트 센터: 비즈팀
    try:
        cc_selectors = [
            'text="옵션을 선택해 주세요. (여러개 선택 가능)"',
            'div:has-text("코스트 센터") >> :has-text("옵션")'
        ]
        cc_field = None
        for sel in cc_selectors:
            loc = page.locator(sel).first
            if await loc.is_visible(timeout=2000):
                cc_field = loc
                break
        
        if cc_field is not None:
            field_text = await cc_field.inner_text() # type: ignore
            if "비즈팀" not in field_text:
                await cc_field.click() # type: ignore
                await asyncio.sleep(1.0)
                biz_opt = None
                # selector 문법 오류 수정 (text=... 사용)
                for opt_sel in ['div[role="option"]:has-text("비즈팀")', 'li:has-text("비즈팀")', 'text=비즈팀']:
                    try:
                        loc = page.locator(opt_sel).last
                        if await loc.is_visible(timeout=1000):
                            biz_opt = loc
                            break
                    except: continue

                if biz_opt:
                    await biz_opt.click() # type: ignore
                    await asyncio.sleep(0.5)
                    await page.mouse.click(10, 10)
                    print("      ✅ 코스트 센터 '비즈팀' 선택 완료")
    except Exception as e:
        print(f"   ⚠️ '코스트 센터' 오류: {e}")

    # 4. 본문 내용 구성 (GAS 스타일 HTML 테이블 적용)
    table_style = 'border-collapse:collapse; width:100%; border:1px solid #999; font-size:12px; margin-top:10px;'
    header_style = 'background-color:#F2F2F2; font-weight:bold; text-align:center; padding:5px; border:1px solid #999;'
    cell_style = 'padding:5px; border:1px solid #999; vertical-align:middle;'
    center_style = 'text-align:center; padding:5px; border:1px solid #999; vertical-align:middle;'
    label_style = 'background-color:#F2F2F2; padding:5px; border:1px solid #999; vertical-align:middle;'

    final_content = f"""
    <table style="{table_style}">
      <thead>
        <tr>
          <th style="{header_style}">구분</th>
          <th style="{header_style}">필수 여부</th>
          <th style="{header_style}">내용</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td style="{cell_style}">개요</td>
          <td style="{center_style}">O</td>
          <td style="{cell_style}">({client_name}) {course_name} 강연료 집행 건</td>
        </tr>
        <tr>
          <td style="{cell_style}">자금 집행 내용(요약)</td>
          <td style="{center_style}">O</td>
          <td style="{cell_style}">강연료({summary_f6})에 대한 비용 집행입니다.</td>
        </tr>
        <tr>
          <td style="{label_style}"><b>첨부 파일</b></td>
          <td style="{center_style}">O</td>
          <td style="{cell_style}">하단 상세 목록 참조</td>
        </tr>
        <tr>
          <td style="{label_style}">1. 통장 사본</td>
          <td style="{center_style}">O</td>
          <td style="{cell_style}">{holder_val}_통장 사본.pdf</td>
        </tr>
        <tr>
          <td style="{label_style}">2. 사업자등록증</td>
          <td style="{center_style}">O</td>
          <td style="{cell_style}">{holder_val}_사업자등록증.pdf</td>
        </tr>
        <tr>
          <td style="{label_style}">3. 거래명세서</td>
          <td style="{center_style}">O</td>
          <td style="{cell_style}">계약서 파일명 기재</td>
        </tr>
        <tr>
          <td style="{label_style}">4. (세금)계산서(청구)</td>
          <td style="{center_style}">O</td>
          <td style="{cell_style}">확인 후 파일명 기재</td>
        </tr>
      </tbody>
    </table>
    <br>
    <div style="font-size:12px;"><b>상세 내용</b></div>
    """
    
    if valid_data_rows:
        filtered_table_data = [real_table_data[0]] + valid_data_rows
        # GAS의 buildHtmlTable 로직을 재현한 로컬 함수 사용
        def build_detailed_table(data):
            if not data: return ""
            html = '<table border="1" cellpadding="0" cellspacing="0" style="border-collapse:collapse; font-size:12px; border:1px solid #999;">'
            for i, row in enumerate(data):
                html += '<tr>'
                for j, cell in enumerate(row):
                    is_header = (i == 0)
                    bg_color = "#F2F2F2" if is_header else "#FFFFFF"
                    font_weight = "bold" if is_header else "normal"
                    style = f"background:{bg_color}; font-weight:{font_weight}; text-align:center; vertical-align:middle; padding:6px 10px; border:1px solid #999; white-space:nowrap;"
                    tag = "th" if is_header else "td"
                    html += f'<{tag} style="{style}">{cell}</{tag}>'
                html += '</tr>'
            html += '</table>'
            return html
            
        table_html = build_detailed_table(filtered_table_data)
        final_content += table_html
            
    content_selectors = ['div[contenteditable="true"]', 'textarea[placeholder*="내용을 입력해"]']
    for cs in content_selectors:
        try:
            content_input = page.locator(cs).first
            if await content_input.is_visible(timeout=3000):
                await content_input.click()
                await page.keyboard.press("Meta+a" if os.name == 'posix' else "Control+a")
                await page.keyboard.press("Backspace")
                if await content_input.evaluate("el => el.isContentEditable"):
                    await content_input.evaluate(f"(el) => {{ el.innerHTML = `{final_content}`; }}")
                else:
                    await content_input.fill(final_content)
                print(f"   ✅ '본문 내용' 입력 완료")
                break
        except: continue

    # 5. 파일 처리 (XLSX 추출 및 업로드)
    download_dir = os.path.join(os.getcwd(), "flex_attachments")
    if not os.path.exists(download_dir): os.makedirs(download_dir)
    downloaded_paths = []

    if spreadsheet_url:
        from datetime import datetime
        now_str = datetime.now().strftime('%y%m%d')
        # GAS 기반 파일명 생성: {clientName}_{courseName}_기타소득 자금집행요청서_{now}
        xlsx_filename = f"{client_name}_{course_name}_기타소득 자금집행요청서_{now_str}.xlsx"
        
        xlsx_path = await export_sheet_to_xlsx(spreadsheet_url, sheet_name, download_dir, config, custom_filename=xlsx_filename)
        if xlsx_path: downloaded_paths.append(xlsx_path)

    if downloaded_paths:
        try:
            target_input = page.locator('input[type="file"]').first
            if await target_input.is_visible(timeout=5000):
                await target_input.set_input_files(downloaded_paths)
                print(f"   ✅ 파일 업로드 완료 ({len(downloaded_paths)}개)")
        except Exception as e:
            print(f"   ❌ 업로드 중 오류: {e}")

async def export_sheet_to_xlsx(sheet_url, sheet_name, download_dir, config, custom_filename=None):
    """
    구글 시트의 특정 탭을 XLSX 파일로 직접 내보냅니다.
    임시 시트를 생성하지 않고 gid 파라미터를 사용하여 해당 탭만 다운로드합니다.
    """
    sa_path = config.get("service_account_path")
    if not (sa_path and os.path.exists(sa_path)):
        print("   ❌ 서비스 계정 파일이 없어 엑셀 수출을 수행할 수 없습니다.")
        return None

    try:
        from googleapiclient.http import MediaIoBaseDownload  # type: ignore

        scopes = [
            'https://www.googleapis.com/auth/spreadsheets.readonly',
            'https://www.googleapis.com/auth/drive.readonly'
        ]
        creds = Credentials.from_service_account_file(sa_path, scopes=scopes)
        service = build('drive', 'v3', credentials=creds)

        # 1. 원본 시트 ID 추출
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
        if not match: return None
        ss_id = match.group(1)

        # 2. 파일명 및 경로 설정
        from datetime import datetime
        now_str = datetime.now().strftime('%y%m%d')
        filename = custom_filename if custom_filename else f"자금집행요청서_{now_str}.xlsx"
        target_path = os.path.join(download_dir, filename)

        # 3. export_media를 사용한 다운로드 (더 안정적임)
        print(f"   📥 Excel API 수출 시도 중: {filename}")
        request = service.files().export_media(
            fileId=ss_id,
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
        
        with open(target_path, 'wb') as f:
            f.write(fh.getvalue())
            
        print(f"   ✅ Excel API 수출 완료: {filename}")
        return target_path

    except Exception as e:
        print(f"   ❌ 엑셀 수출 중 오류: {e}")
        return None
