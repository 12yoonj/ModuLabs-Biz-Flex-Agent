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

async def handle_business_income(page, plan: dict, template_name: str, config: dict):
    """
    [정기-기타/사업소득 자금집행요청서] 양식 전용 핸들러
    """
    print(f"[핸들러] 📝 '{template_name}' 전용 로직 수행 중...")
    
    # 0. 공통 데이터 추출
    spreadsheet_url = plan.get("spreadsheet_url")
    sheet_name = "[정산] 기타/사업소득 자금집행요청서"
    
    # 0-1. 시트에서 필수 데이터 추출 (M5:P, O4, 개요!B3, B4)
    real_table_data: List[List[Any]] = [] # 헤더 포함 (M5:P)
    total_amount_val = "-"
    intro_context = ""
    client_name = ""
    course_name = ""
    
    if spreadsheet_url:
        print(f"   📊 시트에서 직접 데이터 추출 중...")
        # (1) 본문 표 및 데이터 행 파악을 위한 범위 (M5:P)
        res = await fetch_sheet_data_range(spreadsheet_url, sheet_name, "M5:P", config)
        if isinstance(res, list):
            real_table_data = res
        
        if not real_table_data:
            print("   ⚠️ [디버그] M5:P 데이터를 가져오지 못했습니다. 시트 권한이나 범위를 확인해주세요.")
        else:
            print(f"   [디버그] 가져온 데이터 행 수: {len(real_table_data)}")

        # (2) 총 이체 금액 (O4)
        o4_res = await fetch_sheet_data_range(spreadsheet_url, sheet_name, "O4", config)
        if isinstance(o4_res, list) and len(o4_res) > 0 and len(o4_res[0]) > 0:
            total_amount_val = str(o4_res[0][0])
            print(f"   [디버그] 총 이체 금액(O4): {total_amount_val}")

        # (3) 개요 정보 (B3, B4)
        intro_res = await fetch_sheet_data_range(spreadsheet_url, "개요", "B3:B4", config)
        if isinstance(intro_res, list) and len(intro_res) >= 2:
            client_name = str(intro_res[0][0]) if intro_res[0] else "미확인기업"
            course_name = str(intro_res[1][0]) if intro_res[1] else "미확인과정"
            intro_context = f"{client_name} ({course_name}) 에 따른 비용 집행"
            print(f"   [디버그] 개요: {intro_context}")

    # 데이터 행 수 계산 (헤더인 5행 제외, 6행부터)
    real_data_list = cast(List[List[Any]], real_table_data)
    data_rows = real_data_list[1:] if len(real_data_list) > 1 else [] # type: ignore
    # 빈 행 필터링 (M열(은행명) 기준 데이터 존재 여부 확인, M열은 인덱스 0)
    valid_data_rows: List[List[Any]] = []
    for r in data_rows:
        if isinstance(r, list) and len(r) > 0:
            # M열(index 0) 값이 존재하는지 확인 (None 처리 포함)
            val_m = str(r[0]).strip() if r[0] is not None else ""
            if val_m:
                # 테이블 구조 유지를 위해 부족한 열은 "-"로 채움 (최소 4열: M, N, O, P)
                padded_row = list(r)
                while len(padded_row) < 4:
                    padded_row.append("-")
                valid_data_rows.append(padded_row)
    
    row_count = len(valid_data_rows)
    print(f"   [디버그] 유효 데이터 행 수: {row_count}")

    # 1. 제목 입력
    title_value = plan.get("title")
    if title_value:
        title_selectors = ['input[placeholder*="제목"]', 'textarea[placeholder*="제목"]', '[aria-label*="제목"]']
        for ts in title_selectors:
            try:
                locator = page.locator(ts).first
                if await locator.is_visible(timeout=3000):
                    await locator.click()
                    await locator.fill(title_value)
                    print(f"   ✅ '제목' 입력 완료")
                    break
            except: continue

    # 2. 개별 필드 입력 (기타(은행명) -> 계좌 번호 -> 총 이체 금액 -> 예금주 순서)
    # 규칙: 1줄(해당 행 값), 2줄+(본문 내 기재), 0줄(-)
    
    def get_row_value(col_idx: int) -> str:
        if row_count == 1 and valid_data_rows:
            row = valid_data_rows[0]
            if len(row) > col_idx:
                return str(row[col_idx]).strip()
        return ""

    bank_val = "-"
    account_val = "-"
    holder_val = "-"
    
    if row_count == 1:
        bank_val = get_row_value(0) # M열
        account_val = get_row_value(1) # N열
        holder_val = get_row_value(3) # P열
    elif row_count >= 2:
        bank_val = "본문 내 기재"
        account_val = "본문 내 기재"
        holder_val = "본문 내 기재"

    # 필드 설정 (레이블, 값, 에디터 타입, 별칭들)
    # UI에서 "계좌 번호"가 아닌 "계좌번호"일 수 있으므로 별칭(aliases)을 적극 활용
    field_configs = [
        {
            "label": "기타(은행명)",
            "val": bank_val,
            "aliases": ["기타(은행명)", "은행", "은행명", "입금 은행", "입금은행"]
        },
        {
            "label": "계좌 번호",
            "val": account_val,
            "aliases": ["계좌 번호", "계좌번호", "계좌", "입금 계좌", "입금계좌번호"]
        },
        {
            "label": "총 이체 금액",
            "val": total_amount_val,
            "aliases": ["총 이체 금액", "금액", "이체 금액", "예상 비용"]
        },
        {
            "label": "예금주",
            "val": holder_val,
            "aliases": ["예금주", "예금주명", "성명", "입금 성명"]
        }
    ]

    for cfg in field_configs:
        label = cfg["label"]
        val = cfg["val"] or "-"
        aliases = cfg["aliases"]
        
        clean_val = str(val)
        if label == "총 이체 금액":
            if val and val != "-":
                clean_val = re.sub(r"[^0-9]", "", str(val))
            else:
                clean_val = "-"

        # 모든 별칭에 대해 셀렉터 생성
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
                        if await loc.is_visible(timeout=1000): # 빠른 시도를 위해 타임아웃 단축
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
            print(f"   ⚠️ '{label}' 필드를 찾을 수 없습니다. (시도한 별칭: {', '.join(aliases)})")

    # 3. [이체요청일] 입력 로직
    try:
        from datetime import date, timedelta
        import calendar
        def get_business_day(d):
            holidays = {
                date(2026, 1, 1), date(2026, 2, 16), date(2026, 2, 17), date(2026, 2, 18),
                date(2026, 3, 1), date(2026, 3, 2), date(2026, 5, 5), date(2026, 5, 24),
                date(2026, 5, 25), date(2026, 6, 6), date(2026, 8, 15), date(2026, 9, 23),
                date(2026, 9, 24), date(2026, 9, 25), date(2026, 10, 3), date(2026, 10, 9),
                date(2026, 12, 25)
            }
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
        
        print(f"   📅 '이체요청일' 계산 완료: {date_str}")
        
        # 품의서 시작일/종료일과 동일한 로직 적용
        date_label = "이체요청일"
        # 1. 달력 버튼 클릭
        calendar_btn_sel = f"div:has-text('{date_label}') >> button:has-text('날짜')"
        calendar_btn = page.locator(calendar_btn_sel).first
        
        if await calendar_btn.is_visible(timeout=3000):
            await calendar_btn.click()
            await asyncio.sleep(0.5)
            
            # 2. 캘린더 상단 입력창에 날짜 입력
            date_input_sel = 'input[placeholder*="날짜 입력"]'
            date_input = page.locator(date_input_sel).first
            
            if await date_input.is_visible(timeout=3000):
                await date_input.fill(date_str)
                await page.keyboard.press("Enter")
                await asyncio.sleep(0.3)
                
                # 3. 이체요청일 레이블을 다시 클릭하여 빠져나오기
                # 단순히 레이블 클릭이 안 될 경우를 대비해 좌표 클릭(0,0) 폴백 추가 및 셀렉터 확장
                print(f"   [디버그] '{date_label}' 팝업 닫기 시도 중...")
                await asyncio.sleep(0.5)
                
                success_close = False
                # 정확히 "이체요청일" 텍스트만 가진 요소를 찾음 (상단 레이블 등)
                exit_selectors = [
                    f'span:text-is("{date_label}")',
                    f'div:text-is("{date_label}")',
                    f'label:text-is("{date_label}")',
                    f'text="{date_label}"'
                ]
                
                for es in exit_selectors:
                    try:
                        loc = page.locator(es).first
                        if await loc.is_visible(timeout=1000):
                            await loc.click(force=True)
                            success_close = True
                            print(f"   ✅ '{date_label}' 팝업 닫힘")
                            break
                    except:
                        pass
                
                if not success_close:
                    await page.mouse.click(0, 0)
    except Exception as e:
        print(f"   ⚠️ '이체요청일' 처리 중 오류: {e}")

    try:
        # 1. 코스트 센터 필드(인풋/옵션란) 찾기
        print("      🔍 '코스트 센터' 필드 탐색 중 (Placeholder 기반)...")
        cc_field = None
        
        # 방법 A: 사용자가 말한 대로 placeholder 텍스트를 직접 찾기
        placeholder = "옵션을 선택해 주세요. (여러개 선택 가능)"
        cc_field = page.locator(f'text="{placeholder}"').first
        
        if await cc_field.is_visible(timeout=3000):
            print(f"      ✅ placeholder '{placeholder}' 로 필드 특정 성공")
        else:
            print("      ⚠️ Placeholder로 못 찾음. '코스트 센터' 레이블 기반 탐색 시도...")
            # 방법 B: "코스트 센터" 레이블을 찾고 그 옆의 div/버튼 찾기
            label_loc = page.locator('text="코스트 센터"').first
            if await label_loc.is_visible(timeout=1000):
                # 레이블의 부모 행(div) 안에서 옵션/클릭 영역 찾기
                # 보통 레이블과 입력란은 같은 flex 행에 있습니다.
                row = label_loc.locator('xpath=./ancestor::div[contains(@class, "flex") or contains(@style, "display: flex") or contains(@class, "row")][1]')
                if await row.is_visible(timeout=500):
                    cc_field = row.locator(':has-text("옵션")').last
            
            if not cc_field or not await cc_field.is_visible(timeout=500):
                 # 방법 C: 폴백 - 코스트 센터 텍스트를 포함한 클릭 가능한 영역 (사이드바 제외)
                 cc_field = page.locator('div:not(aside):not([class*="sidebar"]):has-text("코스트 센터")').locator(':has-text("옵션")').last

        if cc_field is not None:
            await cc_field.scroll_into_view_if_needed()
            is_already_selected = False
            field_text = await cc_field.inner_text()
            print(f"      📄 현재 필드 텍스트: '{field_text}'")
            
            # 메인 폼의 필드 텍스트에 이미 '비즈팀'이 들어가 있으면 선택된 것
            if "비즈팀" in field_text:
                is_already_selected = True
                print("      ℹ️ '코스트 센터'가 이미 '비즈팀'으로 설정되어 있습니다.")
            
            if not is_already_selected:
                # 3. 클릭하여 옵션 열기
                print("      🔘 필드 클릭하여 옵션 목록 열기...")
                await cc_field.click()
                await asyncio.sleep(1.0) # 로딩 대기
                
                # 4. "[비즈팀]" 또는 "비즈팀" 옵션 찾아서 클릭
                print("      🔍 옵션 목록에서 '비즈팀' 탐색 중...")
                option_found = False
                
                # 목록이 떴는지 확인하기 위해 모든 옵션을 가져와봅니다.
                try:
                    all_options = page.locator('div[role="option"], li[role="option"], [role="listbox"] >> text=비즈팀')
                    options_count = await all_options.count()
                    print(f"      📊 발견된 매칭 옵션 수: {options_count}")
                except: pass

                for target_opt in ["[비즈팀]", "비즈팀"]:
                    opt_sel = f'div[role="option"]:has-text("{target_opt}"), li:has-text("{target_opt}"), [role="listbox"] >> text="{target_opt}"'
                    biz_option = page.locator(opt_sel).last
                    
                    if await biz_option.is_visible(timeout=2000):
                        print(f"      ✅ '{target_opt}' 옵션 클릭 시도...")
                        await biz_option.click()
                        print(f"      ✨ '{target_opt}' 선택 완료")
                        option_found = True
                        break
                
                if not option_found:
                    print("      ⚠️ '비즈팀' 옵션을 시각적으로 찾지 못했습니다. 강제 텍스트 클릭 시도...")
                    try:
                        await page.locator('text="비즈팀"').last.click(timeout=2000)
                        print("      ✅ 텍스트 기반 '비즈팀' 클릭 성공")
                        option_found = True
                    except:
                        print("      ❌ 모든 옵션 탐색 실패")
                
                await asyncio.sleep(0.5)
                
                # 5. 빈 공간을 클릭하여 드롭다운 닫기 (사용자 피드백 반영)
                if option_found:
                    print("      🔄 빈 공간(좌측 상단) 클릭하여 드롭다운 닫기 시도...")
                    try:
                        # 0,0은 가끔 메뉴바 등에 가릴 수 있으므로 10,10 정도를 클릭
                        await page.mouse.click(10, 10)
                        print("      ✅ 빈 공간 클릭 성공")
                    except Exception as e:
                        print(f"      ⚠️ 빈 공간 클릭 실패: {e}")
            else:
                print("      이미 선택되어 있어 건너뜁니다.")
        else:
            print("   ❌ '코스트 센터' 필드(인풋란)를 찾을 수 없습니다. (구조가 변경되었을 수 있음)")
            
    except Exception as e:
        print(f"   ⚠️ '코스트 센터' 처리 중 오류: {e}")

    # 본문 내용 구성 (사용자 요청 양식 적용)
    final_content = f"""
    <p><strong>1. 개요</strong></p>
    <p>- ({client_name}) {course_name}에 따른 비용 집행</p>
    <p>- 강연료 등 기타/사업소득 지급 건</p>
    <br>
    <p><strong>2. 상세 내용</strong></p>
    """
    
    if valid_data_rows:
        # 헤더(index 0) + 유효 데이터 행만 합쳐서 표 생성
        filtered_table_data = [real_table_data[0]] + valid_data_rows
        table_html = json_to_html_table(filtered_table_data)
        final_content += table_html
    
    final_content += f"""
    <br>
    <p><strong>3. 첨부 파일</strong></p>
    <p>- 강사진 필수 첨부 서류 (신분증 사본 / 통장 사본)</p>
    <p>- 자금집행요청서 엑셀 파일 첨부 완료</p>
    """
            
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
                print(f"   ✅ '본문 내용 및 표' 입력 완료")
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
