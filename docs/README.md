# 📊 자연어 기반 사내 데이터 분석 AI Agent (`data_analysis_agent`)

사내 생산·품질·검증 데이터(CSV, 다중 시트 Excel)를 업로드하면 
① 핵심 데이터 시트와 요약/환경 보조 시트를 스마트 분리하여 로컬 SQLite DB를 자동 구축하고, 
② 한국어 자연어 질문으로 특정 컬럼/부분 키워드 SQL 조회 및 유사 주제 클러스터링 다차원 시각화를 수행하며, 
③ 사내 보안용 Local LLM(Ollama 정밀 모드)과 외부 Cloud AI(Groq, Gemini)를 동시에 병렬 비교하여 차트가 포함된 마스터 엑셀 리포트(.xlsx) 및 SQL 덤프(.sql)를 자동 생성하는 실무형 AI 에이전트입니다.

---

## 📁 프로젝트 폴더 및 파일 구조 (Project Structure)

```text
data_analysis_agent/
├── .streamlit/
│   └── config.toml              # 에메랄드 그린(#059669) UI 테마 자동 설정
├── data_agent/                  # AI 에이전트 핵심 코어 패키지
│   ├── __init__.py
│   ├── config.py                # 사내 품질 기준 및 공통 설정 관리 (W47 분리 예정)
│   ├── state.py                 # LangGraph 상태(State) 스키마 정의 (W47 분리 예정)
│   ├── llm.py                   # 멀티 LLM 라우터 (Ollama 정밀 3,500토큰, Groq/Gemini 자동탐색)
│   ├── graph.py                 # LangGraph 워크플로우 (START -> analyzer -> reporter -> END)
│   ├── excel_exporter.py        # 유사 제목 클러스터링, 색상 문맥 해석, BarChart 내장 엑셀 추출기
│   ├── nodes/                   # LangGraph 작업자 노드 모듈
│   │   ├── analyzer_node.py     # [1단계] 파이썬 고속 정밀 집계 및 이상 징후·색상 문맥 추출 노드
│   │   └── reporter_node.py     # [2단계] 도메인 맞춤(공정/TC) 4단락 보고서 생성 및 마크다운 보정 노드
│   └── tools/                   # 데이터 처리 및 시각화 도구 모듈
│       ├── data_loader.py       # 다중 시트 분리, 셀 배경색 추출, SQLite 동기화 및 .sql 덤프 생성기
│       └── visualizer.py        # 다차원 시각화 보조 모듈 (W47 분리 예정)
├── docs/                        # 주차별 개발 진행 보고서 및 문서 보관소
│   ├── README.md                # 프로젝트 구조 및 핵심 기능 안내 문서
│   └── W46_progress.md          # W46 주차 개발 실적 및 기술 고도화 리포트
├── app.py                       # Streamlit 메인 대시보드 (자연어 SQL, 주제 묶음 차트, AI 비서, 뉴스 연동)
├── main.py                      # CLI 터미널 단독 테스트용 실행 파일
├── run.bat                      # 윈도우 원클릭 가상환경 활성화 및 대시보드 실행 배치 파일
├── sample_data.csv              # 다차원 공정 샘플 데이터 (일자·주차·시간대·라인·불량유형 포함)
├── requirements.txt             # 필수 Python 패키지 목록
├── .env.example                 # 환경변수 설정 템플릿 (실제 키는 .env에 보관 및 Git 제외)
└── .gitignore                   # 보안 파일(.env, .venv, *.db) GitHub 업로드 차단 규칙