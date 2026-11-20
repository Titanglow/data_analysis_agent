| 내부 개발 역량 강화 과제  \|  주간 진행 보고서 (W47)  자연어 기반 사내 데이터·문서 분석 AI Agent 개발 현황  수행 기간: W47 (2026.11.16 ~ 2026.11.20)  \|  작성자: 김광태 (임베디드 Test팀 PEG2)  \|  최종 반영: GitHub 커밋 최신화   |
| 구분   | 상세 내용   |
| 프로젝트명   | 자연어 기반 사내 데이터·문서 통합 분석 AI Agent (data_analysis_agent)   |
| 핵심 한 줄 요약   | 사내 엑셀/CSV/워드 파일을 올려 DB화하고, 핀셋 SQL 및 사내 기준(%)에 맞춰 다차원 차트·3대 AI 비교 보고서·실시간 플로팅 채팅을 수행하는 실무 AI 에이전트   |
| 금주 핵심 목표   | ① 보조 모듈 3종(config, state, visualizer) 리팩토링 ② 자동 양식 검증 및 기준값(%) 연동 ③ 워드(.docx) 추출, 플로팅 팝업 채팅, 3회 반복 검증 및 토큰 최적화   |
| 진행률 및 UI 상태   | W47 계획 100% 완료 (+ W48 대용량 토큰 최적화 및 플로팅 채팅 선제 달성) / 에메랄드 그린(#059669) UI 최종 확정   |
1. W46 차주 계획 대비 W47 달성 현황
| 구분   | 계획 항목   | 달성 상태   | 주요 구현 내역   |
| W47 계획 ①   | 보조 모듈 3종 분리 리팩토링(config.py, state.py, visualizer.py)   | ✅ 100% 완료   | 0바이트 상태였던 3개 파일에 설정·상태·시각화 코드를 완전 분리하고 app.py·graph.py와 실연동 완료   |
| W47 계획 ②   | 부서별 엑셀 양식 자동 검증 및사내 기준값(%) 실시간 조절   | ✅ 100% 완료   | 업로드 즉시 문서유형·작성률(%)·중복ID 진단 및 기준값(0.5%~50.0%) 기반 🚨확인필요/✅이상없음 자동 판정   |
| W47 계획 ③   | 동일 데이터 3회 반복 실행일관성(편차) 검증   | ✅ 100% 완료   | 1~3회차 단락별 실제 출력 문장 비교표, 공통 유지 수치 vs 차이점 자동 진단 및 파트회의 검토용 워드(.docx) 추출   |
| 선제 달성 ④   | 플로팅 팝업 채팅방 및 대용량 데이터 토큰 초과 한계 극복   | 🌟 선제 완료   | 1·2·3번 자유 스크롤 플로팅 팝업 채팅 탑재 / 수백 건 데이터(592건 등) 처리 시 발생하는 Groq 413, 429 에러(토큰 초과) 원천 해결   |
2. W47 주요 개발 실적 및 트러블슈팅 내역
① 클린 아키텍처 3대 보조 모듈 분리 리팩토링 (config.py / state.py / visualizer.py)
• data_agent/config.py: 에메랄드 테마 색상(#059669), 고대비 차트 팔레트, 사내 기본 품질 기준값(1.5%), 멀티 LLM 라벨 상수를 중앙 관리하도록 분리
• data_agent/state.py: LangGraph 노드 간 공유 작업 저장소(AgentState TypedDict) 스키마를 분리 정의하여 graph.py와 연결
• data_agent/tools/visualizer.py: 문서 양식 자동 검증 배너, 4대 핵심 KPI 카드, 단일값(100% 동일) 컬럼 자동 배제, 가로형 TOP 순위 차트 전담 모듈화
② 다중 파일 확장자 자동 분류 및 파싱 안정화 (app.py / data_loader.py)
• 다중 파일 분류 및 원클릭 초기화: Excel, CSV, Word, TXT 다중 첨부 시 확장자별 묶음 표시, 2개 이상 동시 병합 분석 및 원클릭 일괄 초기화 버튼 구현
• Pandas 3.0 NaN 충돌 해결: 빈 셀(float NaN) 정규식 에러를 _safe_str() 방어 함수로 차단하고, 엑셀 기본 테마 색상이 강조 색상으로 오탐지되던 문제 수정
③ 미기입 TC ID 전수 추출 및 컬럼별 핀셋 SQL (excel_exporter.py / app.py)
• 전체 건수 조회 & 핀셋 SQL 변환: 100건 제한을 해제해 전체 건수를 조회하도록 복원하고, 'Expected Result 미기입', 부분 키워드(배달) 등을 정확히 SQL로 변환
• 미기입 TC ID 전수 명단 자동 기록: 파이썬이 전체 행을 전수 스캔하여 컬럼별 미기입된 실제 TC ID(Title) 전수 명단을 기초통계표에 기록함으로써, AI가 누락 없이 정확한 TC ID를 식별하도록 개선
④ [트러블슈팅] 대용량 데이터 토큰 한도 초과(413, 429 에러) 원천 해결 (llm.py / app.py)
• 데이터 초강력 다이어트 로직 적용: 500건 이상의 엑셀 업로드 시 Groq 무료 티어의 분당 입력(ITPM)/출력(OTPM) 제한에 걸리는 현상 발견. AI 전송 시 샘플 데이터를 '최대 3~5건, 각 셀당 20자 이내, 필수 10개 열'로 강제 압축(Chunking)하여 프롬프트 토큰 크기를 15,000자 이내로 획기적으로 감축.
• 안정적 AI 모델 고정: 약관 동의가 필요하거나(400 에러) 한도가 작은 모델을 제외하고, 가장 한도가 큰 'llama-3.3-70b-versatile' 등 검증된 모델로 고정하여 끊김 현상 해결
⑤ 마스터 엑셀(.xlsx) 정제 & 빈 페이지 없는 워드(.docx) 추출 (excel_exporter.py)
• 엑셀 리포트 정제: 모든 시트에 분석 일시 및 사내 기준값(threshold%) 판정 요약을 기록하고, 시인성이 떨어지던 불필요 차트 제거
• 경영진 보고용 워드(.docx) 쪽 나누기: 1페이지(종합 요약표)와 2페이지 이후(AI별 독립 페이지 시작)로 분리하고 <w:pageBreakBefore/> 적용으로 빈 페이지(Blank Page) 0건 보장
⑥ 1·2·3번 자유 스크롤 플로팅 팝업 채팅(A) & 3회 반복 상세 검증기(C) 
• 플로팅 팝업 채팅방(Chatflow): 로컬 Ollama 구동 지원. 우측 플로팅 팝업 실행 시 배경 화면 잠금 없이 마우스 휠로 1번(표)·2번(차트)·3번(리포트)을 자유롭게 오가며 대화 가능.
• 3회 반복 '일관성 vs 차이점' 상세 비교기: 1~3회차 출력 비교, 공통 유지 수치 vs 표현 차이점을 자동 진단하고 파트회의 공유용 워드(.docx)로 원클릭 다운로드 제공.
3. 프로젝트 폴더 및 핵심 파일 현황 (W47 최신 기준)
| 경로 및 파일명   | 핵심 역할 및 구현 내용   | 상태   |
| app.py   | 메인 웹 대시보드 (파이프라인, 핀셋 SQL, 플로팅 팝업 채팅, 3회 검증, 토큰 최적화)   | ✅ 고도화   |
| data_agent/config.py   | 에메랄드 테마(#059669), 고대비 팔레트, 기준값(1.5%), 모델 라벨 상수 정의   | ✅ 분리완료   |
| data_agent/state.py   | LangGraph 공유 작업 저장소(AgentState) 스키마 정의   | ✅ 분리완료   |
| data_agent/tools/visualizer.py   | 데이터 양식 자동 검증 배너, 4대 KPI 카드 및 가로형 TOP 순위 차트 모듈   | ✅ 분리완료   |
| data_agent/tools/data_loader.py   | 다중 시트·RGB 배경색 추출(Pandas 3.0), 워드 파싱, SQLite DB 동기화   | ✅ 안정화   |
| data_agent/excel_exporter.py   | 컬럼별 미기입 TC ID 명단 집계, 마스터 엑셀(.xlsx) 및 쪽나누기 워드(.docx) 생성   | ✅ 고도화   |
| data_agent/graph.py   | LangGraph 워크플로우 연결 (START → analyzer → reporter → END)   | ✅ 완료   |
| data_agent/llm.py   | 멀티 LLM 라우터 (Groq 413/429 토큰 초과 방지 및 Llama 3.3 모델 고정 적용)   | ✅ 안정화   |
| data_agent/nodes/analyzer_node.py   | [1단계] 파이썬 기반 통계·기준 초과·이상치 고속 추출 노드   | ✅ 완료   |
| data_agent/nodes/reporter_node.py   | [2단계] 실제 컬럼 기반 4단락 무환각 리포트 생성 및 샘플 행 덤프 자동 차단   | ✅ 고도화   |
| docs/ (5종 문서)   | README.md, W46/W47_progress.md, 파일간 호출 순서.txt, 노드구성도_뷰어.html   | ✅ 최신화   |
4. 전체 5단계 실시간 실행 흐름 요약
| 단계   | 파이프라인 명칭   | 담당 모듈 (.py)   | 핵심 동작 내용   |
| 1단계   | 문서·시트·양식 스캔   | app.py, data_loader.py,visualizer.py   | 확장자별 자동 분류, 메인/보조 시트 분리, 작성률(%)·중복ID 양식 자동 검증   |
| 2단계   | SQLite DB·전수 집계   | data_loader.py,excel_exporter.py   | SQLite DB 자동 적재, 핀셋 SQL 변환, 미기입 TC ID 전수 명단 및 기준(%) 판정 집계   |
| 3단계   | 다차원 시각화 & 3대 AI 추론   | visualizer.py, graph.py,analyzer_node.py, llm.py   | 4대 KPI 카드·TOP 순위 차트 생성, Ollama·Groq·Gemini 병렬 동시 추론 실행 (토큰 초과 방지 적용)   |
| 4단계   | 리포트 품질 검증·보정   | reporter_node.py   | 사내 기준 자동 판정 배너(🚨/✅) 부착, 숫자 환각 및 샘플 덤프 절단 정제   |
| 5단계   | 산출물 패키징 & 후속 대화   | excel_exporter.py,app.py   | 엑셀/워드(.docx) 패키징, 1·2·3번 자유 스크롤 플로팅 채팅 및 3회 일관성 검증   |
5. 차주(W48) 계획 및 파트회의 공유 안건
• 금주 파트 주간회의 시연 및 피드백 수렴: W47 주간 진행 보고서(.docx) 및 공유용 노드구성도 뷰어(.html)를 통해 아키텍처·플로팅 채팅·3회 반복 일관성 검증 결과 시연 및 의견 수렴
• 실무 데이터셋 확대 및 스트레스 테스트: 대용량 데이터(600건 이상) 토큰 최적화 완료를 바탕으로, 병합 셀 엑셀·결함 리포트 등 다양한 실무 양식에 대한 예외 케이스 점검
• 저장소 문서화 최종 정비: 프로젝트 루트 대문 README.md 파일 복사·동기화 및 공유용 실행 흐름도 뷰어 지속 유지관리
```eof

### 📄 1. `docs/README.md` (루트 폴더에도 복사해주세요!)
```markdown:docs/README.md
# 📊 자연어 기반 사내 데이터·문서 통합 분석 AI Agent (`data_analysis_agent`)

사내 생산·품질·검증 데이터(**CSV, 다중 시트 Excel**)와 일반 업무 문서(**Word `.docx`, `.txt`**)를 업로드하면  
① **확장자별 자동 분류 및 핵심 시트·보조 시트(Summary/Environment)·셀 배경색을 스마트 분리하여 로컬 SQLite DB를 자동 구축**하고,  
② **데이터 양식 자동 검증(Validation), 컬럼별 미기입 TC ID 전수 명단 추출, 사내 품질 기준값(`%`) 실시간 판정과 연동된 자연어 핀셋 SQL 조회·다차원 시각화**를 수행하며,  
③ **사내 보안용 Local LLM(Ollama 정밀 모드)과 외부 Cloud AI(Groq, Gemini)를 병렬 비교하여 마스터 엑셀(`.xlsx`), 페이지 분리형 워드(`.docx`), 1·2·3번 자유 스크롤 플로팅 팝업 채팅방(Chatflow)**을 제공하는 실무형 AI 에이전트입니다.

> **💡 대용량 데이터 토큰 최적화 완료:** 600건 이상의 엑셀 처리 시 발생하는 API 토큰 한도 초과(413, 429) 에러를 방지하기 위해 데이터 다이어트 로직(핵심 3~5건, 최대 20자 컷팅)이 적용되어 있습니다.

---

## 🔄 1. 한눈에 보는 파이썬(`.py`) 파일 간 실행 흐름도 (W47 최종 아키텍처)

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