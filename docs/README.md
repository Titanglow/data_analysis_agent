에이 에이전트 핵심

인식 > 계획 > 실행 > 평가의 반복 루프와 도구 사용 


 - 주제 : 사내 환경에서 구동되는 자연어 기반 데이터 분석 Agent
 - 핵심 방향: 사내 파일 및 DB에 직접 연결되어, 사용자가 코드를 몰라도 질문만 하면 SQL/분석 코드 생성, 시각화 차트, 인사이트 리포트까지 자동으로 도출되는 구조
 - 보안 및 비교: 평소에는 사내 유출이 없는 로컬 LLM(Ollama 등)을 기본 모드로 사용하되,(파이썬언어로) PoC 검증용 비교 모드를 두어 로컬 모델과 외부 클라우드 API 간의 속도·정확도 성능 비교 결과로 예정입니다.


## 프로젝트 구조

data_analysis_agent/              # 프로젝트 루트 폴더
│
├── run.bat                       # [실행] 더블클릭 원클릭 실행 배치 파일 (Streamlit 구동)
├── app.py                        # [화면] Streamlit 웹 UI (파일 업로드, 질문 입력, 차트/리포트 출력)
├── requirements.txt              # [환경] 파이썬 필수 패키지 목록
├── .env                          # [보안] 외부 클라우드 API 키 보관 (공유·배포 시 제외)
│
├── data_agent/                   # [핵심] LangGraph 에이전트 본체 패키지
│   ├── config.py                 # 모델 설정 (Ollama 로컬 vs 클라우드 API 스위치), 경로 설정
│   ├── llm.py                    # 로컬/클라우드 LLM 호출 공통 계층 (성능·속도 측정 포함)
│   ├── state.py                  # 그래프 전체가 공유하는 단일 작업 저장소 (AgentState) 정의
│   ├── prompts.py                # Analyzer(분석/코드생성) 및 Reporter(요약보고서) 프롬프트 고정
│   ├── graph.py                  # StateGraph 구성 (노드 배치, 순환 엣지, 자기교정 분기 연결)
│   │
│   ├── nodes/                    # 파이썬 함수 단위의 작업자(Node) 모듈
│   │   ├── analyzer_node.py      # 질문 의도 파악, SQL/Pandas 분석 및 이상치·트렌드 도출
│   │   └── reporter_node.py      # 분석 결과를 바탕으로 경영진/실무자용 마크다운 리포트 작성
│   │
│   └── tools/                    # 결정론적 파이썬 도구(Tool) 모듈
│       ├── data_loader.py        # 사내 CSV, Excel 파일 및 SQLite DB 연결/스키마 추출
│       ├── visualizer.py         # Plotly 기반 시각화 차트 생성
│       └── excel_exporter.py     # 분석 데이터 + 리포트 다중 시트 엑셀(.xlsx) 저장
│
├── input_data/                   # 분석할 사내 파일(CSV/Excel) 및 샘플 DB 보관 폴더
├── output/                       # 실행마다 날짜/시간별로 생성되는 결과 엑셀 및 비교 로그 폴더
└── docs/                         # 매주 회의 공유용 진행 상황 메모 (W46_progress.md 등)

-보안 및 설정 분리 (.env, config.py): 외부 클라우드 API 키는 .env 파일에 분리하여 배포 시 유출을 막고, config.py와 llm.py를 통해 코드 수정 없이 로컬 LLM과 클라우드 API를 스위칭합니다.  
-역할 기반 멀티 노드 (nodes/): 만능 에이전트 하나에 모든 일을 시키지 않고, 데이터를 분석해 JSON 구조로 패턴과 이상치를 뽑아내는 Analyzer(분석가)와 이를 받아 읽기 쉬운 보고서로 쓰는 Reporter(보고자)로 작업자 노드를 분리합니다.   
-도구 분리 (tools/): 파일 로드나 최종 엑셀 파일 저장(export_to_excel)처럼 정확해야 하는 작업은 LLM이 아닌 결정론적인 파이썬 코드가 처리하도록 도구를 분리합니다.   