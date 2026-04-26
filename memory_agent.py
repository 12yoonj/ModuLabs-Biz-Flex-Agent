import os
import requests
import json
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

NOTION_API_KEY = os.getenv("NOTION_API_KEY")
NOTION_MEMORY_PAGE_ID = os.getenv("NOTION_MEMORY_PAGE_ID")
NOTION_VERSION = "2022-06-28"

def save_memory_to_notion(content: str, tag: str = "WORK"):
    """
    Notion 페이지에 기억(텍스트 블록)을 추가합니다.
    """
    if not NOTION_API_KEY or not NOTION_MEMORY_PAGE_ID:
        return {"status": "error", "message": "Notion API Key or Page ID is missing."}

    headers = {
        "Authorization": f"Bearer {NOTION_API_KEY}",
        "Content-Type": "application/json",
        "Notion-Version": NOTION_VERSION
    }

    url = f"https://api.notion.com/v1/blocks/{NOTION_MEMORY_PAGE_ID}/children"

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # 간단한 포맷으로 저장 (Callout 블록 활용)
    # 아이콘 설정
    icon_emoji = "💡"
    if tag.upper() == "WORK": icon_emoji = "💼"
    elif tag.upper() == "PERSON": icon_emoji = "👤"
    elif tag.upper() == "PREF": icon_emoji = "⭐"

    data = {
        "children": [
            {
                "object": "block",
                "type": "callout",
                "callout": {
                    "rich_text": [
                        {
                            "type": "text",
                            "text": {
                                "content": f"[{tag}] {current_time}\n"
                            },
                            "annotations": {
                                "bold": True,
                                "color": "blue"
                            }
                        },
                        {
                            "type": "text",
                            "text": {
                                "content": content
                            }
                        }
                    ],
                    "icon": {
                        "emoji": icon_emoji
                    },
                    "color": "gray_background"
                }
            }
        ]
    }

    try:
        response = requests.patch(url, headers=headers, data=json.dumps(data))
        if response.status_code == 200:
            return {"status": "success"}
        else:
            return {"status": "error", "message": f"Notion API Error: {response.text}"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

def fetch_recent_memories():
    """
    페이지 내의 자식 블록들을 읽어와서 최근 메모리를 반환합니다.
    (간단한 읽기 기능)
    """
    if not NOTION_API_KEY or not NOTION_MEMORY_PAGE_ID:
        return []

    headers = {
        "Authorization": f"Bearer {NOTION_API_KEY}",
        "Notion-Version": NOTION_VERSION
    }

    url = f"https://api.notion.com/v1/blocks/{NOTION_MEMORY_PAGE_ID}/children?page_size=50"
    
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            blocks = response.json().get("results", [])
            memories = []
            for idx, block in enumerate(blocks):
                if block["type"] == "callout":
                    rich_text = block["callout"]["rich_text"]
                    text_content = "".join([rt["plain_text"] for rt in rich_text])
                    # 간단한 파싱
                    title = text_content
                    tag = "work"
                    if "[PERSON]" in text_content: tag = "person"
                    elif "[PREF]" in text_content: tag = "preference"
                    
                    # 날짜 추출 (기본적으로 오늘로 처리하거나 파싱)
                    date_str = datetime.now().strftime("%Y.%m.%d")
                    import re
                    match = re.search(r"\[.*?\] (\d{4}-\d{2}-\d{2})", text_content)
                    if match:
                        date_str = match.group(1).replace("-", ".")
                        
                    # 제목 정제
                    clean_title = text_content.split("\n", 1)[-1] if "\n" in text_content else text_content
                    
                    memories.append({
                        "id": block["id"],
                        "title": clean_title,
                        "tag": tag,
                        "date": date_str
                    })
            return list(reversed(memories))  # 최신순 정렬 (아래에 추가되므로)
        else:
            print(f"Error fetching memories: {response.text}")
            return []
    except Exception as e:
        print(f"Error fetching memories: {e}")
        return []

def analyze_memory_from_chat(chat_history: list):
    from google import genai
    from google.genai import types
    
    gemini_key = os.getenv("GEMINI_API_KEY")
    if not gemini_key:
        return "Gemini API Key is missing.", "WORK"
        
    client = genai.Client(api_key=gemini_key)
    
    history_text = ""
    for msg in chat_history:
        sender = msg.get("sender", "unknown")
        text = msg.get("text", "")
        history_text += f"[{sender}] {text}\n"
        
    system_prompt = """
주어진 EVE(AI 인턴)와 USER의 대화 내용을 분석하여 앞으로 계속 기억해야 할 중요 정보(새로운 업무 규칙, 강사 특이사항, 사용자의 선호도 등)를 추출하세요.
단 1~2문장으로 명확하게 요약하고, 해당 내용의 카테고리 태그를 다음 중 하나로 결정하세요:
- WORK: 업무 진행 현황, 새로운 업무 프로세스 등
- PERSON: 강사, 담당자 등 사람과 관련된 특이사항
- PREF: 사용자의 개인적인 선호도나 패턴

응답은 반드시 아래 JSON 형태로만 출력하세요:
{"summary": "요약된 핵심 기억 내용", "tag": "WORK"}
"""
    
    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=history_text,
            config=types.GenerateContentConfig(
                system_instruction=system_prompt,
                response_mime_type="application/json"
            ),
        )
        data = json.loads(response.text)
        return data.get("summary", "중요 정보를 찾지 못했습니다."), data.get("tag", "WORK")
    except Exception as e:
        print("Gemini Analysis Error:", e)
        return "대화 요약 중 오류가 발생했습니다.", "WORK"
