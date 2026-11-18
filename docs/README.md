# 📊 자연어 기반 사내 데이터·문서 통합 분석 AI Agent (`data_analysis_agent`)

사내 생산·품질·검증 데이터(**CSV, 다중 시트 Excel**)와 일반 업무 문서(**Word `.docx`, `.txt`**)를 업로드하면 **① 확장자별 자동 분류 및 핵심 시트·보조 시트·셀 배경색을 스마트 분리하여 로컬 SQLite DB를 자동 구축**하고, **② 양식 자동 검증(Validation) 및 사내 품질 기준값(`%`)에 연동된 자연어 SQL 조회·가로형 TOP 순위 다차원 시각화**를 수행하며, **③ 사내 보안용 Local LLM(Ollama 정밀 모드)과 외부 Cloud AI(Groq, Gemini)를 동시에 병렬 비교하여 마스터 엑셀(`.xlsx`), 페이지 분리형 워드(`.docx`), SQL 덤프(`.sql`)를 자동 생성**하는 실무형 AI 에이전트입니다.

---

## 🔄 1. 한눈에 보는 파이썬(`.py`) 파일 간 실행 흐름도 (W47 최신 아키텍처)

> 💡 GitHub에서 열면 아래 Mermaid 코드가 **자동으로 화살표 플로우차트**로 렌더링됩니다.

```mermaid
flowchart TB
    User(["👤 사용자 (웹 브라우저)"])

    subgraph UI["🖥️ 1. 메인 대시보드 & 워크플로우 계층 (app.py)"]
        Upload["📂 다중 파일 업로드 & 확장자별 분류<br/>(📊Excel / 📈CSV / 📄Word / 📝TXT + 원클릭 초기화)"]
        NL2SQL["🤖 build_smart_sql()<br/>(전체조회 · 부분키워드 · 미기입 · 정렬 SQL 변환)"]
        ExtHub["🌟 4️⃣ 심화 확장 센터 (A·B·C 동시 가동)<br/>💬 로컬 채팅방(Chatflow) · 📚 기준서 대조(RAG) · 🔁 3회 편차 검증"]
    end

    subgraph ConfigState["⚙️ 2. 설정 & 상태 스키마 계층"]
        Cfg["📌 data_agent/config.py<br/>(테마색상 · 사내 기본 기준값 1.5% · 모델 라벨)"]
        St["🗂️ data_agent/state.py<br/>(LangGraph AgentState 스키마 정의)"]
    end

    subgraph Loader["🗄️ 3. 데이터·문서 파싱 및 DB 계층 (data_agent/tools/data_loader.py)"]
        Parse["📄 다중 포맷 & 배경색 파서<br/>(메인/보조시트 분리 · RGB 강조 감지 · 워드 표 추출)"]
        SQLite[("💾 internal_analysis.db<br/>(SQLite 실시간 동기화 & .sql 덤프)")]
    end

    subgraph Vis["📊 4. 양식 검증 & 다차원 시각화 계층 (data_agent/tools/visualizer.py)"]
        ValBanner["🛡️ render_validation_banner()<br/>(작성률% · 단일값 컬럼 배제 · 중복ID 검증)"]
        Charts["📈 render_smart_charts()<br/>(4대 KPI 카드 · 가로 TOP 순위 · 시나리오 그룹핑)"]
    end

    subgraph Graph["🧠 5. LangGraph 분석 엔진 (data_agent/graph.py)"]
        Start(("START"))
        Analyzer["⚙️ nodes/analyzer_node.py<br/>[1단계: 파이썬 고속 정밀 집계]<br/>기준값(%) 초과 항목 · 이상치 0.01초 추출"]
        Reporter["📝 nodes/reporter_node.py<br/>[2단계: 무환각 보고서 생성]<br/>🚨확인필요/✅이상없음 배너 + 4단락 리포트"]
        EndNode(("END"))
    end

    subgraph LLMs["⚡ 6. 멀티 LLM 병렬 라우터 (data_agent/llm.py)"]
        Ollama["🔒 Local: Ollama<br/>(사내보안 정밀 모드 · 채팅 기본 엔진)"]
        Groq["⚡ Cloud: Groq<br/>(초고속 무료 API)"]
        Gemini["☁️ Cloud: Gemini<br/>(Google 무료 API)"]
    end

    subgraph Exporter["📥 7. 다중 포맷 패키징 계층 (data_agent/excel_exporter.py)"]
        StatTables["📋 build_korean_stat_tables()<br/>(컬럼별 미기입 판정 · 시나리오 군집화 · 한글 통계표)"]
        OutFiles["📚 마스터 엑셀(.xlsx) & 📄 워드(.docx) 추출<br/>(워드 1p 종합요약표 + 2p~ AI별 쪽나누기 적용)"]
    end

    User --> Upload
    Cfg -.-> UI & Vis
    St -.-> Graph
    Upload --> Parse
    Parse --> SQLite
    SQLite --> ValBanner & NL2SQL
    NL2SQL --> Charts
    Charts --> StatTables
    NL2SQL --> Start
    Start --> Analyzer
    Analyzer --> Reporter
    Reporter --> Ollama & Groq & Gemini
    Ollama & Groq & Gemini --> EndNode
    EndNode --> OutFiles
    NL2SQL --> ExtHub
    ExtHub --> Ollama
    OutFiles --> User