# ModuLabs Biz-Flex-Agent 아키텍처 및 동작 개요

이 문서는 AI(Gemini)에게 현재 프로젝트의 맥락과 구조를 설명하여, **향후 개발 방향성(고도화, 리팩토링, 기능 확장 등)을 논의하기 위한 기준 문서(Context Document)**입니다.

---

## 1. 프로젝트 개요 (Overview)
*   **프로젝트명:** ModuLabs Biz-Flex-Agent
*   **목적:** 사내 전자결재 시스템(Flex)의 기안 작성을 자동화하여, Biz팀의 반복적인 운영 업무(계약서, 품의서, 자금집행 등)를 획기적으로 줄이는 AI 기반 자동화 에이전트.
*   **핵심 가치:** 
    1.  **Rule-based + AI의 결합:** 노션(Notion)에 작성된 실시간 업무 가이드라인을 AI(Gemini)가 해석하고, 구글 시트의 원본 데이터에 알맞게 매핑하여 실행 계획을 수립합니다.
    2.  **UI 자동화:** 수립된 계획을 바탕으로 Playwright가 실제 브라우저를 조작하여 Flex 기안을 자동으로 작성합니다.
    3.  **데이터 동기화:** 처리가 완료된 건은 원본 구글 시트의 상태값을 자동으로 업데이트합니다.

## 2. 주요 기술 스택 (Tech Stack)
*   **언어:** Python 3.10+
*   **LLM API:** Google Gemini API (gemini-2.5-flash) - 데이터 매핑 및 본문 텍스트 생성, Tool Calling 활용
*   **브라우저 자동화:** Playwright (Chromium 기반, Async)
*   **데이터 연동:** 
    *   `gspread`, `google-api-python-client`: 구글 시트 데이터 및 서식(병합, 색상 등) 읽기, 구글 드라이브 파일 다운로드 및 상태 업데이트
    *   `requests`, `Playwright`: 노션 페이지 크롤링 및 웹 파일 다운로드
*   **기타:** `dotenv` (로컬 환경변수 관리), `asyncio` (비동기 처리)

## 3. 핵심 아키텍처 (Architecture & Modules)

프로젝트는 기능별로 모듈화되어 있으며, 각 스크립트는 다음과 같은 역할을 수행합니다.

### 3-1. `flex_agent.py` (메인 엔진 및 오케스트레이션)
*   **환경/설정 관리:** API 키, 계정 정보, 서비스 계정(JSON), 노션 가이드 URL 등을 로컬 `.env`로 관리.
*   **AI 플래닝 엔진 (`run_planning_flow`):**
    *   Gemini 모델에 `system_instruction`을 주입하여 에이전트의 역할을 정의.
    *   **Tool Calling:** `master_sheet_reader_tool`, `fetch_filtered_sheet_data_tool`, `fetch_rich_sheet_data_tool` 도구를 통해 AI가 스스로 구글 시트 데이터를 분석하고 가져오도록 설계.
    *   가져온 데이터와 노션 가이드를 바탕으로 **구조화된 JSON 배열(실행 계획서)**을 출력.
*   **상태 업데이트 (`update_sheet_cell`):** 자동화가 끝난 후, 구글 시트의 특정 행(Row) 상태값을 `[완료]` 등으로 자동 변경. (최근 업데이트를 통해 필터링 과정에서 소실될 수 있는 원본 엑셀 행 번호(`_original_row_index`)를 추적하여 정확도를 개선함)

### 3-2. `workflow_handlers.py` (공통 자동화 로직 및 라우터)
*   **공통 UI 조작 함수:** Flex 폼에서 반복적으로 사용되는 액션들을 추상화.
    *   `fill_text_field`: 텍스트 및 금액 입력
    *   `internal_fill_date`: 달력 UI 조작 및 날짜 입력 (모달창 닫힘 방지 등 예외 처리 포함)
    *   `select_list_field`: 드롭다운 목록 선택
    *   `download_file`: 구글 드라이브(문서, PDF 변환 지원) 및 일반 웹 링크를 통한 첨부파일 로컬 임시 다운로드
    *   `json_to_html_table`: 구글 시트의 리치 데이터(병합, 배경색 등)를 HTML Table 태그로 변환하여 본문에 주입.
*   **라우터 (`dispatch_workflow`):** 실행할 양식 이름(Template Name)에 따라 아래의 전용 핸들러로 작업을 분배.

### 3-3. 전용 양식 핸들러 (Form-Specific Handlers)
각 양식별로 Flex의 DOM 구조나 요구하는 필드가 다르기 때문에, 세부적인 입력 로직을 분리하여 관리합니다.
*   `workflow_contract_instructor.py`: [계약서 등 검토 · 승인] 강사 용역 (상대의 유형, 체결일/시작일 분기 처리 등)
*   `workflow_education_services.py`: [계약서 등 검토 · 승인] 교육 용역 (매출/매입 처리 등)
*   `workflow_business_income.py`: [정기-기타/사업소득 자금집행요청서] (다수 인원에 대한 복잡한 테이블 생성 및 합계 처리)
*   `workflow_general_funding.py`: [정기-자금집행요청서] (일반 정산 건 엑셀 내보내기 로직 등)

## 4. 전체 자동화 워크플로우 (Execution Flow)

1.  **실행 및 타겟 선택:** 사용자가 CLI 메뉴에서 실행할 양식(예: 강사 용역 계약)과 마스터 시트(구글 시트)를 선택합니다.
2.  **데이터 로드 및 AI 플래닝:** 
    *   `flex_agent.py`가 노션 가이드를 읽어오고, 구글 시트 데이터를 Tool로 가져와 Gemini에게 전달합니다.
    *   Gemini는 가이드 규칙에 맞게 데이터를 매핑하여 JSON 포맷의 실행 계획(Plan)을 반환합니다.
3.  **Flex 자동 기안:**
    *   Playwright가 Flex에 시크릿 세션으로 로그인합니다.
    *   워크플로우 양식을 검색하여 선택합니다.
    *   `workflow_handlers`를 통해 텍스트, 날짜, 금액, 드롭다운 필드를 자동으로 채웁니다.
    *   `attachments`에 URL이 존재하면 백그라운드에서 다운로드 후 Flex에 업로드합니다.
4.  **수동 검토 및 완료:**
    *   기안 전송 직전 상태에서 멈추며, 사용자가 브라우저를 통해 눈으로 최종 확인 후 '요청하기'를 누르도록 유도합니다.
5.  **동기화:** 브라우저 종료 시, 구글 시트의 해당 건 상태를 '완료'로 업데이트하여 중복 기안을 방지합니다.

