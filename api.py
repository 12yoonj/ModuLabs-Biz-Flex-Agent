import asyncio
import builtins
import flex_agent
import memory_agent
from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

class MemoryRequest(BaseModel):
    content: str = ""
    tag: str = "WORK"
    chat_history: list = []

@app.post("/api/memory")
async def save_memory(req: MemoryRequest):
    content = req.content
    tag = req.tag
    
    if req.chat_history:
        content, tag = memory_agent.analyze_memory_from_chat(req.chat_history)
        
    result = memory_agent.save_memory_to_notion(content, tag)
    # 함께 요약된 내용도 응답에 포함
    result["extracted_content"] = content
    result["extracted_tag"] = tag
    return result

@app.get("/api/memory")
async def get_memories():
    memories = memory_agent.fetch_recent_memories()
    return {"memories": memories}

class WebConsole:
    def __init__(self, websocket: WebSocket):
        self.ws = websocket

    async def print(self, *args, **kwargs):
        text = " ".join(map(str, args))
        # 터미널에 뜨는 것처럼 한 줄씩 전송
        try:
            await self.ws.send_json({"type": "stdout", "text": text})
        except:
            pass

    async def ainput(self, prompt=""):
        try:
            await self.ws.send_json({"type": "prompt", "text": prompt})
            # 클라이언트로부터 텍스트 수신 대기
            data = await self.ws.receive_text()
            return data
        except Exception as e:
            print("WebSocket Receive Error:", e)
            return ""

@app.websocket("/ws/flex")
async def websocket_endpoint(websocket: WebSocket):
    await websocket.accept()
    console = WebConsole(websocket)

    original_print = builtins.print
    original_ainput = flex_agent.ainput

    def mock_print(*args, **kwargs):
        # 터미널용 원래 print도 실행 (디버깅용)
        original_print(*args, **kwargs)
        # 웹소켓으로도 전송
        asyncio.create_task(console.print(*args, **kwargs))

    # 몽키 패칭
    builtins.print = mock_print
    flex_agent.ainput = console.ainput

    try:
        # 사용자가 "Flex 기안 올리기"를 누르면, 
        # 메인 메뉴로 들어가지 않고 바로 1번(수행) 동작으로 진입하도록 하거나 
        # 처음부터 메인 메뉴를 보여줄 수 있습니다.
        # 여기서는 메인 메뉴를 띄워 터미널과 동일한 경험을 제공합니다.
        await flex_agent.main_menu()
    except WebSocketDisconnect:
        pass
    except Exception as e:
        await console.print(f"\n[오류] {e}")
    finally:
        # 복구
        builtins.print = original_print
        flex_agent.ainput = original_ainput
        try:
            await websocket.close()
        except:
            pass

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
