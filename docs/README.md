# 📊 자연어 기반 사내 데이터·문서 통합 분석 AI Agent (`data_analysis_agent`)

사내 생산·품질·검증 데이터(**CSV, 다중 시트 Excel**)와 일반 업무 문서(**Word `.docx`, `.txt`**)를 업로드하면  
① **확장자별 자동 분류 및 핵심 시트·보조 시트(Summary/Environment)·셀 배경색을 스마트 분리하여 로컬 SQLite DB를 자동 구축**하고,  
② **데이터 양식 자동 검증(Validation), 컬럼별 미기입 TC ID 전수 명단 추출, 사내 품질 기준값(`%`) 실시간 판정과 연동된 자연어 핀셋 SQL 조회·가로형 TOP 순위 다차원 시각화**를 수행하며,  
③ **사내 보안용 Local LLM(Ollama 정밀 모드)과 외부 Cloud AI(Groq, Gemini)를 동시에 병렬 비교하여 마스터 엑셀(`.xlsx`), 페이지 분리형 워드(`.docx`), 1·2·3번 자유 스크롤 플로팅 팝업 채팅방(Chatflow) 및 3회 반복 일관성 검증 리포트**를 제공하는 실무형 AI 에이전트입니다.

---

## 🎯 1. 무엇을 만들었나? (핵심 한 줄 요약)
- 사내 엑셀/CSV/워드/텍스트 파일을 올리면 자동으로 양식을 검증해 DB화하고, 자연어 질문과 사내 기준값(`%`)에 맞춰 **다차원 차트 · 3대 AI 비교 보고서(엑셀 `.xlsx` / 워드 `.docx`) · 실시간 팝업 질의응답(Chatflow)**을 자동 수행하는 에이전트입니다.

---

## 🔄 2. 한눈에 보는 파이썬(`.py`) 파일 간 실행 흐름도 (W47 최종 아키텍처)

> 💡 GitHub에서 열면 아래 Mermaid 코드가 **자동으로 화살표 플로우차트**로 렌더링됩니다.

```mermaid
flowchart TB
    User(["👤 사용자 (웹 브라우저)"])

    subgraph UI["🖥️ 1. 메인 대시보드 & 워크플로우 계층 (app.py)"]
        Upload["📂 다중 파일 업로드 & 확장자별 분류<br/>(📊Excel / 📈CSV / 📄Word / 📝TXT + 원클릭 초기화)"]
        NL2SQL["🤖 build_smart_sql() & _detect_mentioned_columns()<br/>(전체조회 · 컬럼별 미기입 핀셋 추출 · 부분키워드 · 정렬 SQL 변환)"]
        ExtHub["🌟 4️⃣ 플로팅 팝업 채팅방(A) & 3회 반복 상세 검증(C)<br/>💬 1·2·3번 자유 스크롤 팝업 채팅 · 🔁 1~3회차 차이점 비교 및 회의용 워드(.docx) 추출"]
    end

    subgraph ConfigState["⚙️ 2. 설정 & 상태 스키마 계층"]
        Cfg["📌 data_agent/config.py<br/>(테마색상 #059669 · 사내 기본 기준값 1.5% · 모델 라벨)"]
        St["🗂️ data_agent/state.py<br/>(LangGraph AgentState 스키마 정의)"]
    end

    subgraph Loader["🗄️ 3. 데이터·문서 파싱 및 DB 계층 (data_agent/tools/data_loader.py)"]
        Parse["📄 다중 포맷 & 배경색 파서<br/>(메인/보조시트 분리 · 실제 RGB 강조 감지 · 워드 문단/표 추출)"]
        SQLite[("💾 internal_analysis.db<br/>(SQLite 실시간 동기화 & .sql 덤프)")]
    end

    subgraph Vis["📊 4. 양식 검증 & 다차원 시각화 계층 (data_agent/tools/visualizer.py)"]
        ValBanner["🛡️ render_validation_banner()<br/>(문서유형 · 작성률% · 단일값 컬럼 배제 · 중복ID 검증)"]
        Charts["📈 render_smart_charts()<br/>(4대 KPI 카드 · 가로 TOP 순위 · 유사 제목 시나리오 그룹핑)"]
    end

    subgraph Graph["🧠 5. LangGraph 분석 엔진 (data_agent/graph.py)"]
        Start(("START"))
        Analyzer["⚙️ nodes/analyzer_node.py<br/>[1단계: 파이썬 고속 정밀 집계]<br/>기준값(%) 초과 항목 · 이상치 0.01초 추출"]
        Reporter["📝 nodes/reporter_node.py<br/>[2단계: 무환각 보고서 생성 & 정제]<br/>🚨확인필요/✅이상없음 배너 + 샘플나열 차단 + 4단락 리포트"]
        EndNode(("END"))
    end

    subgraph LLMs["⚡ 6. 멀티 LLM 병렬 라우터 (data_agent/llm.py)"]
        Ollama["🔒 Local: Ollama<br/>(사내보안 정밀 모드 · 채팅 기본 엔진)"]
        Groq["⚡ Cloud: Groq<br/>(초고속 무료 API · 2,200 토큰)"]
        Gemini["☁️ Cloud: Gemini<br/>(Google 무료 API)"]
    end

    subgraph Exporter["📥 7. 전수 통계 집계 & 다중 포맷 패키징 계층 (data_agent/excel_exporter.py)"]
        StatTables["📋 build_korean_stat_tables()<br/>(컬럼별 미기입 실제 TC ID 전수 명단 · 시나리오 군집화 · 기준 판정표)"]
        OutFiles["📚 마스터 엑셀(.xlsx) & 📄 워드(.docx) 추출<br/>(워드 1p 종합요약표 + 2p~ AI별 독립 페이지 쪽나누기 · 빈 페이지 0건)"]
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
    ExtHub --> Ollama
    OutFiles --> User