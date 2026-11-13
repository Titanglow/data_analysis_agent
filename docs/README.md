# 📊 자연어 기반 사내 데이터·문서 통합 분석 AI Agent (`data_analysis_agent`)

사내 생산·품질·검증 데이터(**CSV, 다중 시트 Excel**)뿐만 아니라 일반 업무 문서(**Word `.docx`, `.txt`, `.md`**)를 업로드하면 **① 핵심 데이터와 보조 시트/내장 표를 스마트 분리하여 로컬 SQLite DB를 자동 구축**하고, **② 한국어 자연어로 SQL 조회 및 한눈에 보는 TOP 순위 시각화**를 수행하며, **③ 사내 보안용 Local LLM(Ollama 정밀 모드)과 외부 Cloud AI(Groq, Gemini)를 동시에 병렬 비교하여 마스터 엑셀 리포트(`.xlsx`) 및 SQL 덤프(`.sql`)를 자동 생성**하는 실무형 AI 에이전트입니다.

---

## 🔄 1. 한눈에 보는 파이썬(`.py`) 파일 간 실행 흐름도 (Architecture Flow)

> 💡 GitHub에서 이 문서를 열면 아래 다이어그램이 **자동으로 화살표 그래프**로 표시됩니다.

```mermaid
flowchart TB
    User(["👤 사용자 (웹 브라우저)"])

    subgraph UI["🖥️ 1. 메인 화면 계층 (app.py)"]
        Upload["📂 파일 업로드 & 자연어 질문 입력"]
        NL2SQL["🤖 build_smart_sql()<br/>(자연어 → SQL 자동 변환)"]
        Charts["📊 render_smart_charts()<br/>(4대 KPI 카드 & 가로 TOP 순위 차트)"]
        Scope["🔄 resolve_analysis_scope()<br/>(1번 조회 + 3번 질문 스마트 연동)"]
    end

    subgraph Loader["🗄️ 2. 데이터·문서 추출 및 DB 계층 (data_agent/tools/data_loader.py)"]
        Parse["📄 다중 포맷 파서<br/>(CSV / 다중시트 .xlsx / Word .docx / .txt)"]
        SQLite[("💾 internal_analysis.db<br/>(SQLite 자동 동기화 & .sql 덤프)")]
    end

    subgraph Graph["🧠 3. LangGraph 워크플로우 엔진 (data_agent/graph.py)"]
        Start(("START"))
        Analyzer["⚙️ analyzer_node.py<br/>[1단계: 파이썬 고속 집계 노드]<br/>시트요약 · 색상문맥 · 이상치 0.01초 추출"]
        Reporter["📝 reporter_node.py<br/>[2단계: 도메인 맞춤 보고서 노드]<br/>IF/ELSE 판정 배너 + 4단락 리포트 완성"]
        EndNode(("END"))
    end

    subgraph LLMs["⚡ 4. 멀티 LLM 병렬 라우터 (data_agent/llm.py)"]
        Ollama["🔒 Local: Ollama<br/>(사내보안 정밀 모드)"]
        Groq["⚡ Cloud: Groq<br/>(초고속 무료 API)"]
        Gemini["☁️ Cloud: Gemini<br/>(Google 무료 API)"]
    end

    subgraph Exporter["📥 5. 리포트·차트 패키징 계층 (data_agent/excel_exporter.py)"]
        StatTables["📋 build_korean_stat_tables()<br/>(유사 제목 클러스터링 & 한글 통계표)"]
        ExcelOut["📚 마스터 엑셀 추출 (.xlsx)<br/>(모델별 종합비교 + 시트통합요약 + BarChart)"]
    end

    User --> Upload
    Upload --> Parse
    Parse --> SQLite
    SQLite --> NL2SQL
    NL2SQL --> Charts
    Charts --> StatTables
    NL2SQL --> Scope
    Scope --> Start
    Start --> Analyzer
    Analyzer --> Reporter
    Reporter --> Ollama & Groq & Gemini
    Ollama & Groq & Gemini --> EndNode
    EndNode --> ExcelOut
    ExcelOut --> User