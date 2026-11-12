# 📊 자연어 기반 사내 데이터 분석 AI Agent (`data_analysis_agent`)

사내 생산·품질·업무 데이터(CSV, Excel)를 업로드하면 **① 자동으로 로컬 SQLite DB를 구축**하고, **② 한국어 자연어 질문으로 SQL 조회 및 다차원 시각화**를 수행하며, **③ 사내 보안용 Local LLM(Ollama)과 외부 고성능 Cloud AI(Groq, Gemini)를 동시에 병렬 비교하여 차트가 포함된 마스터 엑셀 리포트(.xlsx)를 자동 생성**하는 실무형 AI 에이전트입니다.

---

## 📁 프로젝트 폴더 및 파일 구조 (Project Structure)

```text
data_analysis_agent/
├── .streamlit/
│   └── config.toml              # 에메랄드 그린(#059669) UI 테마 자동 설정
├── data_agent/                  # AI 에이전트 핵심 코어 패키지
│   ├── __init__.py
│   ├── config.py                # 사내 품질 기준(1.5%) 및 공통 설정 관리
│   ├── state.py                 # LangGraph 상태(State) 스키마 정의
│   ├── llm.py                   # 멀티 LLM 라우터 (Ollama 고속/정밀, Groq/Gemini 자동탐색)
│   ├── graph.py                 # LangGraph 워크플로우 (START -> analyzer -> reporter -> END)
│   ├── excel_exporter.py        # 100% 한글 통계·집계표 및 BarChart 내장 엑셀 추출기
│   ├── nodes/                   # LangGraph 작업자 노드 모듈
│   │   ├── analyzer_node.py     # [1단계] 데이터 패턴·이상 징후(Anomaly)·트렌드 추출 노드
│   │   └── reporter_node.py     # [2단계] IF/ELSE 기준 판정 배너 + 4단락(5M1E·R&R) 보고서 노드
│   └── tools/                   # 데이터 처리 및 시각화 도구 모듈
│       ├── data_loader.py       # CSV/Excel -> SQLite DB(internal_analysis.db) 자동 변환기
│       └── visualizer.py        # 다차원 시각화 보조 모듈
├── docs/                        # 주차별 개발 진행 보고서 및 문서 보관소
│   ├── README.md                # 문서 폴더 안내 및 프로젝트 구조 요약
│   └── W46_progress.md          # W46 주차 개발 실적 및 트러블슈팅 리포트
├── app.py                       # Streamlit 메인 웹 대시보드 (AI 비서 질문조합, 뉴스연동, 팝업 UI)
├── main.py                      # CLI 터미널 단독 테스트용 실행 파일
├── run.bat                      # 윈도우 원클릭 가상환경 활성화 및 대시보드 실행 배치 파일
├── sample_data.csv              # 다차원 공정 샘플 데이터 (일자·주차·시간대·라인·불량유형 포함)
├── requirements.txt             # 필수 Python 패키지 목록
├── .env.example                 # 환경변수 설정 템플릿 (실제 키는 .env에 보관 및 Git 제외)
├── .gitignore                   # 보안 파일(.env, .venv, *.db) GitHub 업로드 차단 규칙
└── README.md                    # 프로젝트 메인 소개 문서