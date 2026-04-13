import re
import asyncio
import os
from typing import Any, Dict, List, Optional
from playwright.async_api import Page  # type: ignore
from workflow_handlers import download_file, reset_focus, scroll_page, fill_text_field, internal_fill_date, select_list_field  # type: ignore

async def handle_education_services(page: Page, plan: Dict[str, Any], template_name: str, config: Dict[str, Any]):
    """
    [계약서 등 검토 · 승인] 교육 용역 전용 핸들러
    """
    print(f"[핸들러] 📝 '{template_name}' 전용 로직 수행 중...")
    
    main_area = page.locator('main').first
    
    fields = plan.get("fields", {})
    
    # 1. 제목 입력 (강사 용역 로직 동기화)
    title_value = plan.get("title")
    if title_value:
        try:
            print(f"   [진행] 1. 제목 입력 시도 중... (목표: {title_value})")
            title_selectors = [
                'input[placeholder*="문서 제목을 입력해 주세요"]',
                'textarea[placeholder*="문서 제목을 입력해 주세요"]',
                'input[placeholder*="제목"]'
            ]
            for sel in title_selectors:
                try:
                    title_locator = page.locator(sel).first
                    if await title_locator.is_visible(timeout=2000):
                        await title_locator.click()
                        await asyncio.sleep(0.3)
                        is_mac = os.name == 'posix'
                        cmd_key = "Meta" if is_mac else "Control"
                        await page.keyboard.press(f"{cmd_key}+a")
                        await page.keyboard.press("Backspace")
                        await page.keyboard.type(str(title_value), delay=30)
                        await page.keyboard.press("Enter")
                        print(f"   ✅ '제목' 입력 완료")
                        await scroll_page(page, 250)
                        break
                except: continue
        except Exception as e:
            print(f"   ⚠️ '제목' 입력 중 오류: {e}")

    # 2. 필드 입력 (교육 용역 특화 매핑)
    await fill_text_field(page, "계약명", fields.get("계약명"))
    await fill_text_field(page, "체결 상대자", fields.get("체결 상대자"))
    
    # 매출/매입: 고정값 '매출' (가이드 준수)
    await select_list_field(page, "매출 · 매입", "매출")
    await scroll_page(page, 300)

    # 상대의 유형: 고정값 '법인사업자' (가이드 준수)
    await select_list_field(page, "상대의 유형", "법인사업자")

    # 체결(예정)일은 공란으로 둡니다. (가이드 준수)
    await scroll_page(page, 150)
    await internal_fill_date(page, "시작일", fields.get("시작일"))
    await internal_fill_date(page, "종료일", fields.get("종료일"))
    
    amount = fields.get("계약금액(부가세 포함)")
    if amount:
        await fill_text_field(page, "계약금액(부가세 포함)", amount)

    # 3. 본문 내용 작성
    content_value = fields.get("본문 내용")
    if content_value:
        # 줄 바꿈을 기준으로 나누되, 빈 줄은 제거하여 본문 내용을 촘촘하게 만듭니다.
        # 단, 제목(문단) 사이에는 줄바꿈을 하여 띄어쓰도록 조정합니다.
        lines = [l.strip() for l in str(content_value).split("\n") if l.strip()]
        formatted_segments = []
        
        for line in lines:
            # 숫자로 시작하는 제목 행인지 확인 (예: 1., 2), (1) 등)
            if re.match(r"^(\d+[\.\)]|\(\d+\))\s*.*", line):
                # 제목 행이면 볼드 처리
                formatted_segments.append(f"<b>{line}</b>")
            else:
                # 일반 본문 행
                formatted_segments.append(line)
        
        # 제목 행 앞에만 빈 줄을 추가하여 문단 간격을 확보합니다.
        # 첫 번째 행은 제외합니다.
        final_lines = []
        for i, seg in enumerate(formatted_segments):
            if i > 0 and seg.startswith("<b>"):
                final_lines.append("") # 제목 앞에 빈 줄 추가
            final_lines.append(seg)
            
        formatted_text = "<br>".join(final_lines)
        # 줄 간격은 기존의 1.8로 복구합니다.
        final_content = f'<div style="line-height: 1.8;">{formatted_text}</div>'
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

    # 4. 파일 업로드 (품의서 로직 적용)
    attachments = plan.get("attachments", [])
    if attachments:
        download_dir = "flex_attachments"
        if not os.path.exists(download_dir): os.makedirs(download_dir)
        
        print(f"[핸들러] 📎 파일 처리 중... (대상: {len(attachments)}개)")
        downloaded_paths = []
        for i, url in enumerate(attachments):
            path = await download_file(url, download_dir, i, config)
            if path:
                downloaded_paths.append(path)
                print(f"   📥 다운로드 완료: {os.path.basename(path)}")
 
        if downloaded_paths:
            try:
                print("   ⏳ 업로드 필드 로딩 대기 중...")
                file_input_selector = 'input[type="file"]'
                target_input = None
                for _ in range(20): # 0.5초 * 20 = 10초
                    for frame in page.frames:
                        try:
                            loc = frame.locator(file_input_selector)
                            if await loc.count() > 0:
                                target_input = loc.first
                                break
                        except: continue
                    if target_input: break
                    await asyncio.sleep(0.5)

                if target_input is not None:
                    await target_input.set_input_files(downloaded_paths)  # type: ignore
                    print(f"   ✅ 파일 업로드 완료 ({len(downloaded_paths)}개)")
                    await asyncio.sleep(1)
                else:
                    print("   ⚠️ 업로드 필드를 찾을 수 없습니다.")
            except Exception as e:
                print(f"   ❌ 업로드 중 오류: {e}")