import asyncio
import re
from playwright.async_api import Page  # type: ignore
from typing import Dict, Any

async def handle_contract_instructor(page: Page, plan: Dict[str, Any], template_name: str, config: Dict[str, Any]):
    """
    [계약서 등 검토 · 승인] 강사 용역 양식 전용 핸들러.
    """
    from workflow_handlers import download_file, reset_focus, scroll_page, fill_text_field, internal_fill_date, select_list_field # type: ignore
    import os
    print(f"[핸들러] 📝 '{template_name}' 전용 로직 수행 중...")
    
    # 헤더(공유 버튼 등)를 제외한 메인 폼 영역으로 탐색 범위 제한
    main_area = page.locator('[data-scope="modal"][data-part="body"], .c-bjNJrx, main, #root').first

    # 1. 전달받은 데이터 로깅
    fields = plan.get("fields", {})
    print(f"   [디버그] 전달된 필드 데이터 키: {list(fields.keys())}")

    # 2. 제목 입력 섹션
    title_value = plan.get("title")
    if title_value:
        try:
            print(f"   [진행] 1. 제목 입력 시도 중... (목표: {title_value})")
            title_selectors = [
                'input[placeholder*="문서 제목을 입력해 주세요"]',
                'textarea[placeholder*="문서 제목을 입력해 주세요"]',
                'input[placeholder*="제목"]',
                'input[placeholder*="기안"]',
                '.flex-doc-title input',
                'input[name="title"]'
            ]
            
            title_found = False
            for sel in title_selectors:
                try:
                    title_locator = page.locator(sel).first
                    if await title_locator.is_visible(timeout=1500):

                        await title_locator.click()
                        await asyncio.sleep(0.3)
                        
                        # 텍스트 삭제
                        is_mac = os.name == 'posix'
                        cmd_key = "Meta" if is_mac else "Control"
                        await page.keyboard.press(f"{cmd_key}+a")
                        await page.keyboard.press("Backspace")
                        
                        # 입력
                        await page.keyboard.type(str(title_value), delay=30)
                        print(f"   ✅ '제목' 입력 성공")
                        
                        # 스크롤 
                        await scroll_page(page, 250)
                        title_found = True
                        break
                except Exception as e:
                    continue
        except Exception as e:
            print(f"   ⚠️ '제목' 섹션 처리 중 오류: {e}")

    # 3. 필드 매핑 수행 (순서 조정됨)
    await fill_text_field(page, "계약명", "전문가 출강 계약")
    await fill_text_field(page, "체결 상대자", fields.get("체결 상대자"))

    # 매출 · 매입: 선택 후 스크롤 동작 추가
    await select_list_field(page, "매출 · 매입", fields.get("매출 · 매입"))

    # 상대의 유형: 시트 값에 따른 정확한 매핑
    party_type_raw = str(fields.get("상대의 유형", "")).strip()
    if party_type_raw == "개인":
        party_type = "개인"
    elif "개인 사업자" in party_type_raw or party_type_raw == "개인사업자":
        party_type = "개인사업자"
    else:
        party_type = party_type_raw
    
    await select_list_field(page, "상대의 유형", party_type)

    await internal_fill_date(page, "체결(예정)일", fields.get("체결(예정)일"))
    await internal_fill_date(page, "시작일", fields.get("시작일"))
    await internal_fill_date(page, "종료일", fields.get("종료일"))
    
    amount = fields.get("계약금액(부가세 포함)")
    if amount:
        clean_amount = re.sub(r"[^0-9]", "", str(amount))
        await fill_text_field(page, "계약금액(부가세 포함)", clean_amount)

    # 날인 방법: 고정된 '모두싸인' 선택
    await select_list_field(page, "날인 방법", "모두싸인")


    # 4. 본문 내용 작성
    content_value = fields.get("본문 내용")
    if content_value:
        # 번호 붙은 제목(예: 1. 제목) 볼드 처리
        lines = str(content_value).split("\n")
        bolded_lines = []
        for line in lines:
            line = line.strip()
            # "1. ", "1) " 또는 "(1) " 로 시작하는 제목 스타일 매칭
            if i_match := re.match(r"^(\d+[\.\)]|\(\d+\))\s*(.*)", line):
                prefix = i_match.group(1)
                rest = i_match.group(2)
                bolded_lines.append(f"<b>{prefix} {rest}</b>")
            else:
                bolded_lines.append(line)
        
        formatted_text = "<br>".join(bolded_lines)
        final_content = f"<div>{formatted_text}</div>"
        content_selectors = ['div[contenteditable="true"]', 'textarea[placeholder*="내용을 입력해"]']
        for cs in content_selectors:
            try:
                content_input = page.locator(cs).first
                if await content_input.is_visible(timeout=3000):
                    await content_input.click()
                    is_mac = os.name == 'posix'
                    await page.keyboard.press(f"{('Meta' if is_mac else 'Control')}+a")
                    await page.keyboard.press("Backspace")
                    await content_input.evaluate(f"(el) => {{ el.innerHTML = `{final_content}`; }}")
                    print(f"   ✅ '본문 내용' 입력 완료")
                    break
            except: continue

    # 5. 파일 업로드
    attachments = plan.get("attachments", [])
    if attachments:
        download_dir = "flex_attachments"
        if not os.path.exists(download_dir): os.makedirs(download_dir)
        
        downloaded_paths = []
        for i, url in enumerate(attachments):
            path = await download_file(url, download_dir, i, config)
            if path: downloaded_paths.append(path)

        if downloaded_paths:
            try:
                file_input = page.locator('input[type="file"]').first
                await file_input.set_input_files(downloaded_paths)
                print(f"   ✅ 파일 업로드 완료 ({len(downloaded_paths)}개)")
            except: pass
