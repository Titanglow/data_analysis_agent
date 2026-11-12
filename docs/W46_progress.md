# 📊 [W46 주간 진행 보고] 사내 데이터 분석 AI Agent 구축

## 1. 프로젝트 개요
-   과제 목적:   사내 기밀 데이터 외부 유출 방지(Local LLM) 및 외부 고성능 Cloud AI 비교(PoC)를 통한 실무형 데이터 분석 자동화 에이전트 자체 개발 역량 확보
-   핵심 설계 원칙:   
  1.   역할 분리:   합계·건수·불량률·전주 대비 증감(%p) 계산은 Python(Pandas/SQLite) 코드가 100% 정확히 수행하고, LLM은 추세 해석·원인 가설(5M1E)·조치 제안 작성에 집중
  2.   멀티 모델 병렬 비교:   Local Ollama(`qwen2.5-coder:7b`) + Cloud API(`Groq`, `Google Gemini`) 동시 병렬 실행

## 2. 프로젝트 폴더 및 파일 구조 현황 (W46)
| 경로 및 파일명 | 핵심 역할 및 기능 | 상태 |
| :--- | :--- | :---: |
| `app.py` | Streamlit 메인 대시보드 (자연어→SQL, 다차원 차트, 병렬 비교, 시각화 팝업) | ✅ 완료 |
| `run.bat` / `main.py` | 원클릭 대시보드 실행 배치 파일 및 CLI 테스트 실행 파일 | ✅ 완료 |
| `sample_data.csv` | 다차원 공정 샘플 데이터 (일자, 주차, 시간대, 라인, 생산량, 불량수, 불량유형 등) | ✅ 완료 |
| `data_agent/llm.py` | 멀티 LLM 호출기 (Ollama 정밀/빠른 모드, Groq/Gemini 활성 모델 자동 탐색 및 폴백) | ✅ 완료 |
| `data_agent/graph.py` | LangGraph 워크플로우 (`START` → `analyzer` → `reporter` → `END`) | ✅ 완료 |
| `data_agent/excel_exporter.py` | 100% 한글 통계·라인·불량유형 집계표 및 엑셀 네이티브 차트(`BarChart`) 추출기 | ✅ 완료 |
| `data_agent/nodes/analyzer_node.py` | 1단계 분석가 노드 (JSON 패턴 추출 및 빠른 모드 파이썬 고속 처리) | ✅ 완료 |
| `data_agent/nodes/reporter_node.py` | 2단계 보고자 노드 (사내 기준 1.5% IF/ELSE 경고 배너 + 4단락 5M1E 리포트) | ✅ 완료 |
| `data_agent/tools/data_loader.py` | CSV/Excel 업로드 즉시 SQLite DB(`internal_analysis.db`) 자동 구축 및 SQL 실행 | ✅ 완료 |
| `data_agent/config.py` | 공통 설정값(불량률 임계치, 모델 기본값 등) 분리용 파일 | ⚪ W47 분리 예정 |
| `data_agent/state.py` | LangGraph State 정의 분리용 파일 (현재 `graph.py` 내 포함) | ⚪ W47 분리 예정 |
| `data_agent/tools/visualizer.py` | 시각화 차트 함수 모듈 분리용 파일 (현재 `app.py` 내 포함) | ⚪ W47 분리 예정 |

## 3. W46 주요 개발 및 문제 해결 실적
1.   Git/GitHub 연동 및 보안 설정:   PyCharm `.venv` 환경 구축 및 `.gitignore`를 통한 `.env` API 키 외부 노출 원천 차단
2.   SQLite 자동 DB화 및 자연어→SQL 변환기:   파일 업로드 즉시 DB 테이블(`internal_data`) 생성 및 한국어 자연어 조건(`"A라인만"`, `"불량수 20 넘는 행"`) 자동 SQL 변환
3.   API 오류 해결 및 병렬 실행:   OpenAI 크레딧 소진 이슈를 무료 초고속 `Groq API`로 대체하고, 서버 활성 모델 자동 탐색(`models.list`)으로 404 에러 해결. `ThreadPoolExecutor`로 다중 모델 동시 병렬 실행 구현
4.   UI/UX 및 다차원 시각화 고도화:   초록색(`#059669`) 테마 통일, `st.session_state`로 엑셀 추출 시 화면 증발 현상 해결, 모델 제목 클릭 시 `@st.dialog` 상세 시각화 팝업 구현, 사전 집계(Groupby)로 합쳐진 막대그래프 가독성 해결
5.   한글 집계 및 차트 내장 마스터 엑셀 추출:   영문 `describe()` 제거 후 100% 한글 기초통계·불량유형·전주 대비 증감(%p) 표 구현 및 엑셀 시트 내 비교 막대차트 자동 삽입

## 4. 향후 계획 (W47 ~ W48)
-   W47:   빈 모듈 3개(`config.py`, `state.py`, `visualizer.py`) 코드 분리 리팩토링, 사내 불량률 기준값(1.5%) UI 조절 기능, 타 부서 CSV 업로드 시 양식/결측치 자동 검증 추가
-   W48:   사내 규정·기준서·과거 보고서 지식(RAG) 연동 및 Word(`.docx`) 주간보고서 자동 다운로드 기능 추가