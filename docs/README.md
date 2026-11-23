# 📊 자연어 기반 사내 데이터·문서 통합 분석 AI Agent (`data_analysis_agent`)

사내 생산·품질·검증 데이터(**CSV, 다중 시트 Excel**)와 일반 업무 문서(**Word, TXT**)를 업로드하면  
① **확장자별 자동 분류 및 핵심 시트·셀 배경색 분리를 통해 로컬 SQLite DB를 자동 구축**하고,  
② **데이터 양식 자동 검증, 미기입 TC ID 전수 추출, 사내 품질 기준값(`%`) 실시간 판정과 연동된 다차원 시각화**를 수행하며,  
③ **사내 보안용 Local LLM(Ollama)과 외부 Cloud AI(Groq, Gemini)를 병렬 비교하여 마스터 엑셀 리포트(`.xlsx`) 및 쪽나누기 워드(`.docx`)를 생성하는 실무형 AI 에이전트**입니다.

> **💡 대용량 토큰 최적화 & RAG 지식 연동:** 수백 건의 데이터 처리 시 발생하는 API 토큰 한도 초과 에러를 '데이터 초강력 다이어트(Chunking)'로 방지하며, 사내 규정 문서(Word/TXT)를 업로드하면 이를 읽고 규정에 맞게 답변하는 RAG(검색 증강 생성) 기술이 적용되어 있습니다.

---

## 🔄 1. 한눈에 보는 실행 흐름도 (Architecture Flow)

```mermaid
flowchart TB
    User(["👤 사용자"])

    subgraph UI["🖥️ 1. 메인 대시보드 계층 (app.py)"]
        Upload["📂 분석 데이터 업로드<br/>(Excel / CSV)"]
        RAGUp["📚 사내 지식 문서 업로드<br/>(RAG용 규정서/회의록)"]
        Chat["💬 자유 스크롤 플로팅 채팅방"]
    end

    subgraph Loader["🗄️ 2. 데이터·지식 파싱 계층 (data_loader.py)"]
        Parse["📄 포맷 파서 & 배경색 감지"]
        SQLite[("💾 SQLite DB 적재")]
        Knowledge["🧠 사내 지식 텍스트 추출"]
    end

    subgraph Graph["🧠 3. LangGraph 워크플로우 엔진 (graph.py)"]
        Analyzer["⚙️ [1단계] 파이썬 전수 집계 노드<br/>미기입 명단·기준 초과(%) 100% 산출"]
        Reporter["📝 [2단계] 무환각 리포트 생성 노드<br/>🚨/✅ 판정 배너 + RAG 지식 주입"]
    end

    subgraph LLMs["⚡ 4. 멀티 LLM 라우터 (llm.py)"]
        Ollama["🔒 Local: Ollama (사내보안)"]
        Groq["⚡ Cloud: Groq (토큰 최적화)"]
        Gemini["☁️ Cloud: Gemini (Google API)"]
    end

    subgraph Exporter["📥 5. 리포트 패키징 계층 (excel_exporter.py)"]
        OutFiles["📚 마스터 엑셀(.xlsx) & 📄 쪽나누기 워드(.docx)"]
    end

    User --> Upload & RAGUp
    Upload --> Parse --> SQLite
    RAGUp --> Knowledge
    SQLite --> Analyzer
    Knowledge -.-> Reporter
    Analyzer --> Reporter
    Reporter --> Ollama & Groq & Gemini
    Ollama & Groq & Gemini --> OutFiles
    Chat --> Ollama & Groq
    OutFiles --> User