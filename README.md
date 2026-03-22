# ModuLabs Biz팀 전용 Flex Agent 🤖

이 프로젝트는 ModuLabs의 Flex 전자결재 기안 작성을 자동화하고, Notion 가이드라인 및 Google Spreadsheet 데이터를 연동하여 업무를 보조하는 AI 에이전트입니다. Python 기반으로 작성되었으며 Playwright를 통해 웹 브라우저 자동화를 수행하고, Gemini API를 통해 문맥을 이해하여 필요한 결재 양식을 초안 단계까지 자동으로 완성합니다.

---

## 🛠 준비물 (Prerequisites)

이 프로그램을 실행하기 위해 사전에 아래 항목들이 준비되어 있어야 합니다.

1. **Python 3.10 이상**
   - https://www.python.org/downloads/
   - Windows: `winget install Python.Python.3`
   - MacOS: `brew install python`
2. **Google Gemini API Key**
   - 아래 링크를 통해 api key를 복사하여 환경 변수로 저장해주세요.
   - https://drive.google.com/file/d/17ZuW7ekx9RCQQABhO5MMkMe30ONVqSRw/view?usp=sharing
3. **Google Cloud 서비스 계정 (Service Account) 발급 방법**
   - 구글 시트 연동을 위해 서비스 계정이 필요합니다. 
   - 아래 링크를 통해 json 파일을 다운 받은 후, flex_agent.py와 같은 폴더에 저장해주세요.
   - https://drive.google.com/file/d/1iIJxDkMOxhyMXoFz7KjQMwVZGfAgcKDF/view?usp=sharing
4. **Flex 계정 정보**
   - Flex 포털 로그인을 위한 이메일 ID 및 비밀번호
5. **Notion 업무 가이드 URL**
   - 사내 업무 결재 가이드라인이 작성된 Notion 페이지 주소 
   - 아래 링크를 통해 url을 복사하여 환경 변수로 저장해주세요.
   - https://drive.google.com/file/d/1cxNuSpyT1uXZaG1UZs_bhboOtOav_0wi/view?usp=sharing
6. **[비즈] 고객사명_과정명(YYMMDD)_마스터 시트**
   - 본 에이전트는 교육 과정별 마스터 시트를 기반으로 작동합니다. 
   - 아래 템플릿의 사본을 만들어, 담당 과정에 대한 시트 url을 `4. 교육 마스터 시트 관리` 메뉴에서 등록해 주세요.
   - https://docs.google.com/spreadsheets/d/18Ydul7kg7j7zlp9ZmRb-q5aLnvq57yow-ThqCr5x51s/edit?gid=137597507#gid=137597507

---

## 🚀 설치 및 세팅 방법

### 1. 소스 코드 다운로드
프로젝트 폴더를 로컬 컴퓨터의 원하는 위치에 다운로드(또는 Git Clone)합니다.

### 2. 필수 라이브러리 설치
터미널(또는 명령 프롬프트)을 열고 프로젝트 폴더로 이동한 뒤, 아래 명령어를 실행하여 필수 패키지들을 설치합니다.

```bash
# 프로젝트 폴더로 이동
cd 경로/선택/포함된/에이전트

# 라이브러리 설치
pip install -r requirements.txt
```

### 3. 브라우저 자동화 도구(Playwright) 초기화
Flex 자동 로그인을 수행하기 위해 내장 브라우저 모듈을 설치해야 합니다.

```bash
playwright install chromium
```

*(참고: 에러가 발생하면 `python3 -m playwright install chromium` 으로 시도해 보세요.)*

---

## ⚙️ 초기 환경 변수 세팅

설치가 완료되었으면, 스크립트를 최초 실행하여 환경 변수 및 설정값을 등록해야 합니다. **소스 코드에 직접 민감한 정보를 입력하실 필요가 없습니다.**

1. 터미널에서 아래 명령어로 에이전트를 실행합니다.
   ```bash
   python3 flex_agent.py
   ```

2. 메인 메뉴가 나타나면 **`3. 환경 변수 관리 (등록/수정/삭제)`** 를 선택합니다.
3. 이어서 **`1. 환경 변수 등록 (신규)`** 를 선택하고, 화면의 지시에 따라 준비해 둔 정보들을 입력합니다:
   - `Gemini API Key`
   - `Flex ID (메일)` *본인 사내 계정 정보
   - `Flex 비밀번호` *본인 사내 계정 정보
   - `서비스 계정 경로` *.json을 포함한 파일명만 입력하시면 됩니다.
   - `노션 가이드 URL`

> **보안 안내**: 입력하신 민감한 정보(비밀번호 등)는 프로그램이 존재하는 폴더 내 숨김 파일인 `.env`에 안전하게 저장되며, 팀원들과 공유하는 버전 관리(Git 등)에는 포함되지 않습니다.

---

## 🕹 사용 방법

세팅이 완료된 후, 메인 메뉴의 기능을 자유롭게 사용할 수 있습니다.

### 메인 메뉴 구성
```text
=============================================
🤖 Flex 에이전트 메인 메뉴
=============================================
1. Flex 에이전트 수행
2. Flex 로그인 테스트 (Playwright)
3. 환경 변수 관리 (등록/수정/삭제)
4. 교육 마스터 시트 관리 (등록/수정/삭제)
5. 노션 업무 가이드 정적 업데이트
0. 프로그램 종료
=============================================
```

- **`1. Flex 에이전트 수행`**: 핵심 자동화 기능입니다. 최신 노션 가이드를 기반으로 결재 양식을 선택하고, 필요한 데이터를 Google Sheet 등에서 가져와 Flex 기안창을 자동으로 띄워 작성해 줍니다.
- **`2. Flex 로그인 테스트`**: Playwright 브라우저가 정상 구동되며 내 계정으로 플렉스 로그인이 잘 되는지 눈으로 직접 테스트해 볼 수 있습니다.
- **`3. 환경 변수 관리 (등록/수정/삭제)`**: 환경 변수 및 설정값을 등록, 수정, 삭제할 수 있습니다.
- **`4. 교육 마스터 시트 관리 (등록/수정/삭제)`**: 교육 과정별 마스터 시트를 등록, 수정, 삭제할 수 있습니다.
- **`5. 노션 업무 가이드 정적 업데이트`**: 사내 규칙이 변경되었을 때, 노션 데이터를 수동으로 다시 긁어와 로컬 지식으로 업데이트합니다.

---

## ⚠️ 주의 사항 / 트러블슈팅

1. **구글 시트 접근 에러 권한 부족 (`403 Forbidden`)**
   - 팀 드라이브에 서비스 계정 ID(`~.iam.gserviceaccount.com`)가 편집자 권한으로 등록되어 있는지 꼭 확인하세요.
