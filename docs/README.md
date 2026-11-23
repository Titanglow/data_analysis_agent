# 📊 자연어 기반 사내 데이터·문서 통합 분석 AI Agent (`data_analysis_agent`)

사내 생산·품질·검증 데이터(**CSV, 다중 시트 Excel**)와 일반 업무 문서(**Word `.docx`, `.txt`**)를 업로드하면  
① **확장자별 자동 분류 및 핵심 시트·보조 시트(Summary/Environment)·셀 배경색을 스마트 분리하여 로컬 SQLite DB를 자동 구축**하고,  
② **데이터 양식 자동 검증(Validation), 컬럼별 미기입 TC ID 전수 명단 추출, 사내 품질 기준값(`%`) 실시간 판정과 연동된 자연어 핀셋 SQL 조회·다차원 시각화**를 수행하며,  
③ **사내 보안용 Local LLM(Ollama 정밀 모드)과 외부 Cloud AI(Groq, Gemini)를 병렬 비교하여 마스터 엑셀(`.xlsx`), 페이지 분리형 워드(`.docx`), 1·2·3번 자유 스크롤 플로팅 팝업 채팅방(Chatflow)**을 제공하는 실무형 AI 에이전트입니다.

> **💡 대용량 데이터 토큰 최적화 완료:** 600건 이상의 엑셀 처리 시 발생하는 API 토큰 한도 초과(413, 429) 에러를 방지하기 위해 데이터 다이어트 로직(핵심 열 10개, 최대 3~5건 샘플, 최대 20자 컷팅)이 강력하게 적용되어 있습니다.

---

## 🔄 1. 한눈에 보는 파이썬(`.py`) 파일 간 실행 흐름도 (최종 아키텍처)

```mermaid
flowchart TB
    User(["👤 사용자 (웹 브라우저)"])

    subgraph UI["🖥️ 1. 메인 대시보드 & 워크플로우 계층 (app.py)"]
        Upload["📂 다중 파일 업로드 & 확장자별 분류<br/>(📊Excel / 📈CSV / 📄Word / 📝TXT)"]
        NL2SQL["🤖 build_smart_sql() & _detect_mentioned_columns()<br/>(전체조회 · 컬럼별 핀셋 추출 · 부분키워드)"]
        ExtHub["🌟 4️⃣ 플로팅 팝업 채팅방(A) & 3회 검증(C)<br/>💬 1·2·3번 자유 스크롤 · 데이터 다이어트 기반 팝업 채팅"]
    end

    subgraph ConfigState["⚙️ 2. 설정 & 상태 스키마 계층"]
        Cfg["📌 data_agent/config.py<br/>(에메랄드 테마 #059669 · 사내 기본 기준 1.5%)"]
        St["🗂️ data_agent/state.py<br/>(LangGraph AgentState 스키마)"]
    end

    subgraph Loader["🗄️ 3. 데이터·문서 파싱 및 DB 계층 (data_agent/tools/data_loader.py)"]
        Parse["📄 다중 포맷 & 배경색 파서<br/>(메인/보조 분리 · 실제 RGB 감지 · 워드 파싱)"]
        SQLite[("💾 internal_analysis.db<br/>(SQLite 실시간 동기화)")]
    end

    subgraph Vis["📊 4. 양식 검증 & 다차원 시각화 계층 (data_agent/tools/visualizer.py)"]
        ValBanner["🛡️ render_validation_banner()<br/>(작성률% · 단일값 배제 · 중복ID)"]
        Charts["📈 render_smart_charts()<br/>(4대 KPI · 가로 TOP 순위 · 유사 시나리오)"]
    end

    subgraph Graph["🧠 5. LangGraph 분석 엔진 (data_agent/graph.py)"]
        Start(("START"))
        Analyzer["⚙️ nodes/analyzer_node.py<br/>[1단계: 파이썬 전수 집계]<br/>미기입 TC ID 명단·기준 초과(%) 100% 산출"]
        Reporter["📝 nodes/reporter_node.py<br/>[2단계: 무환각 리포트 생성]<br/>🚨/✅ 판정 배너 + 샘플 덤프 차단"]
        EndNode(("END"))
    end

    subgraph LLMs["⚡ 6. 멀티 LLM 라우터 (data_agent/llm.py)"]
        Ollama["🔒 Local: Ollama<br/>(사내보안)"]
        Groq["⚡ Cloud: Groq<br/>(토큰 한도 초과 방지 최적화)"]
        Gemini["☁️ Cloud: Gemini<br/>(Google API)"]
    end

    subgraph Exporter["📥 7. 다중 포맷 패키징 계층 (data_agent/excel_exporter.py)"]
        StatTables["📋 build_korean_stat_tables()<br/>(미기입 ID 전수 명단 · 통계표)"]
        OutFiles["📚 마스터 엑셀(.xlsx) & 📄 워드(.docx)<br/>(워드 2p~ AI별 강제 쪽나누기)"]
    end

    User --> Upload
    Cfg -.-> UI & Vis
    St -.-> Graph
    Upload --> Parse
    Parse --> SQLite
    SQLite --> ValBanner & NL2SQL
    NL2SQL --> Charts
    Charts --> StatTables
    StatTables --> Start & ExtHub
    Start --> Analyzer
    Analyzer --> Reporter
    Reporter --> Ollama & Groq & Gemini
    Ollama & Groq & Gemini --> EndNode
    EndNode --> OutFiles
    ExtHub --> Ollama & Groq
    OutFiles --> User