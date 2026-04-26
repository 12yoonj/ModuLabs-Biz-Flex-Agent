import { useState, useRef, useEffect } from 'react'
import './App.css'

function App() {
  const [isWorking, setIsWorking] = useState(false);
  const [messages, setMessages] = useState([
    { id: 1, sender: 'eve', text: '시스템 대기 중. 지시를 기다립니다.', isInitial: true }
  ]);
  const [inputValue, setInputValue] = useState('');
  const [memories, setMemories] = useState([]);
  const wsRef = useRef(null);
  const messagesEndRef = useRef(null);

  const fetchMemories = async () => {
    try {
      const response = await fetch('http://localhost:8000/api/memory');
      const data = await response.json();
      if (data.memories) {
        setMemories(data.memories);
      }
    } catch (error) {
      console.error('Error fetching memories:', error);
    }
  };

  useEffect(() => {
    fetchMemories();
  }, []);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  const connectWebSocket = () => {
    if (wsRef.current) return;
    const ws = new WebSocket('ws://localhost:8000/ws/flex');
    
    ws.onopen = () => {
      setMessages(prev => [...prev, {
        id: Date.now(),
        sender: 'eve',
        text: '[시스템 연결] 터미널 인터페이스를 활성화합니다...'
      }]);
    };
    
    ws.onmessage = (event) => {
      const data = JSON.parse(event.data);
      if (data.type === 'stdout' || data.type === 'prompt') {
        setMessages(prev => {
          const lastMsg = prev[prev.length - 1];
          if (lastMsg && lastMsg.isTerminal) {
            const updated = [...prev];
            updated[updated.length - 1] = {
              ...lastMsg,
              text: lastMsg.text + (lastMsg.text ? '\n' : '') + data.text
            };
            return updated;
          } else {
            return [...prev, {
              id: Date.now(),
              sender: 'eve',
              isTerminal: true,
              text: data.text
            }];
          }
        });
      }
    };
    
    ws.onclose = () => {
      wsRef.current = null;
      setMessages(prev => [...prev, {
        id: Date.now(),
        sender: 'eve',
        text: '[시스템 연결 종료] 터미널 인터페이스가 닫혔습니다.'
      }]);
    };
    
    wsRef.current = ws;
  };

  const handleToggleWork = async () => {
    if (!isWorking) {
      setIsWorking(true);
      setMessages(prev => [...prev, {
        id: Date.now(),
        sender: 'eve',
        text: '출근 처리 되었습니다. 시스템이 활성화되었습니다.'
      }]);
    } else {
      setIsWorking(false);
      setMessages(prev => [...prev, {
        id: Date.now(),
        sender: 'eve',
        text: '퇴근 처리 되었습니다. 시스템을 대기 모드로 전환합니다.'
      }]);
    }
  };

  const handleDirectCommand = (cmdText) => {
    const userMsg = {
      id: Date.now(),
      sender: 'user',
      text: cmdText
    };
    
    setMessages(prev => [...prev, userMsg]);

    if (wsRef.current && wsRef.current.readyState === WebSocket.OPEN) {
      wsRef.current.send(cmdText);
      return;
    }

    if (cmdText.includes('기안') || cmdText.includes('flex')) {
      connectWebSocket();
    } else {
      setTimeout(() => {
        handleEveCommand(cmdText, [...messages, userMsg]);
      }, 1000);
    }
  };

  const handleSendMessage = async (e) => {
    e.preventDefault();
    if (!inputValue.trim()) return;

    const userMsg = {
      id: Date.now(),
      sender: 'user',
      text: inputValue
    };

    setMessages(prev => [...prev, userMsg]);
    setInputValue('');

    if (wsRef.current && wsRef.current.readyState === WebSocket.OPEN) {
      wsRef.current.send(userMsg.text);
      return;
    }

    if (userMsg.text.includes('기안') || userMsg.text.includes('flex')) {
      connectWebSocket();
    } else {
      setTimeout(() => {
        handleEveCommand(userMsg.text, [...messages, userMsg]);
      }, 1000);
    }
  };

  const handleEveCommand = async (text, chatHistory = []) => {
    const id = Date.now();
    if (text.includes('기억해')) {
      setMessages(prev => [...prev, {
        id: Date.now() + 1,
        sender: 'eve',
        text: '대화 맥락을 분석하여 중요 정보를 Notion 메모리 DB에 기록 중입니다...'
      }]);

      try {
        const res = await fetch('http://localhost:8000/api/memory', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ 
            content: '', 
            tag: 'WORK',
            chat_history: chatHistory 
          })
        });
        const data = await res.json();
        
        if (data.status === 'success') {
          setMessages(prev => [...prev, {
            id: Date.now() + 2,
            sender: 'eve',
            text: `명령 확인. 대화를 분석하여 다음 내용을 안전하게 기록했습니다:\n- 요약: ${data.extracted_content}\n- 태그: [${data.extracted_tag}]`
          }]);
          fetchMemories(); // Refresh memory list
        } else {
          setMessages(prev => [...prev, {
            id: Date.now() + 2,
            sender: 'eve',
            text: '기억 저장에 실패했습니다: ' + data.message
          }]);
        }
      } catch (e) {
        setMessages(prev => [...prev, {
          id: Date.now() + 2,
          sender: 'eve',
          text: 'API 서버와 통신할 수 없습니다.'
        }]);
      }
    } else {
      setMessages(prev => [...prev, {
        id,
        sender: 'eve',
        text: '메시지를 수신했습니다. 업무 지시가 필요하시면 말씀해 주세요.'
      }]);
    }
  };

  return (
    <div className="app-container">
      {/* Sidebar Panel */}
      <aside className="sidebar">
        <div className="eve-profile">
          <div className="eve-orb">
            <div className="eve-eye"></div>
          </div>
          <div className="eve-name">E·V·E</div>
          <div className={`status-badge ${isWorking ? 'working' : 'off'}`}>
            <div className="status-dot"></div>
            {isWorking ? 'ONLINE' : 'STANDBY'}
          </div>
        </div>

        <div className="action-buttons">
          <button className="btn" onClick={handleToggleWork}>
            {isWorking ? '⏹ 퇴근하기' : '▶ 출근하기'}
          </button>
          <button className="btn" onClick={() => handleDirectCommand('Flex 기안 올리기 시작')}>
            📄 Flex 기안 올리기
          </button>
          <button className="btn" onClick={() => handleDirectCommand('지금 내용 기억해줘')}>
            💾 현재 대화 기억하기
          </button>
        </div>
        
        <div style={{marginTop: 'auto', fontSize: '9px', fontFamily: 'var(--font-display)', color: 'var(--color-space-gray)', textAlign: 'center'}}>
          SYSTEM v0.1.0<br/>
          RENDER HOSTED
        </div>
      </aside>

      {/* Chat Panel */}
      <main className="chat-panel">
        <header className="chat-header">
          <div className="chat-title">EVE COMM LINK</div>
          <div className="chat-time">{new Date().toLocaleTimeString('ko-KR')}</div>
        </header>

        <div className="message-stream">
          {messages.map((msg) => (
            <div key={msg.id} className={`message ${msg.sender}`} style={msg.isTerminal ? { whiteSpace: 'pre-wrap', fontFamily: 'monospace' } : {}}>
              {msg.text}
              {msg.toolResult && (
                <div className="tool-card">
                  <div className="tool-header">
                    <span>{msg.toolResult.title}</span>
                    <span className={`tool-status ${msg.toolResult.status === 'SUCCESS' ? 'success' : ''}`}>
                      {msg.toolResult.status}
                    </span>
                  </div>
                  <div className="tool-body">
                    {msg.toolResult.details.map((detail, idx) => (
                      <div key={idx} className="tool-row">
                        <span style={{color: 'var(--color-dim-blue)'}}>{detail.label}</span>
                        <span>{detail.value}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          ))}
          <div ref={messagesEndRef} />
        </div>

        <form className="input-area" onSubmit={handleSendMessage}>
          <input
            type="text"
            className="chat-input"
            value={inputValue}
            onChange={(e) => setInputValue(e.target.value)}
            placeholder="EVE에게 업무를 지시하세요..."
            disabled={!isWorking && messages.length > 1}
          />
          <button type="submit" className="send-btn" disabled={!isWorking && messages.length > 1 && !inputValue}>
            TRANSMIT
          </button>
        </form>
      </main>

      {/* Memory Viewer Panel */}
      <aside className="memory-viewer">
        <div className="memory-header">
          <span>MEMORY DB</span>
          <span style={{fontSize: '10px', color: 'var(--color-dim-blue)'}}>{memories.length}</span>
        </div>
        
        <div className="memory-list">
          {memories.map((mem) => (
            <div key={mem.id} className="memory-card">
              <div className="memory-title">{mem.title}</div>
              <div className="memory-meta">
                <span className={`memory-tag ${mem.tag}`}>
                  {mem.tag === 'work' ? 'WORK' : mem.tag === 'person' ? 'PERSON' : 'PREF'}
                </span>
                <span className="memory-date">{mem.date}</span>
              </div>
            </div>
          ))}
        </div>
      </aside>
    </div>
  )
}

export default App
