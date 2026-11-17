# 📊 [W47 주간 진행 보고] 모듈 리팩토링 완성 및 양식 자동 검증·시각화 고도화

## 1. W47 핵심 추진 목표 (W46 차주 계획 이행)
1. **클린 아키텍처 모듈 분리 리팩토링:** `config.py`, `state.py`, `visualizer.py` 3대 보조 모듈 코드 분리 및 `app.py` / `graph.py` 실연동 완료
2. **부서별 다양한 엑셀 양식 자동 검증(Validation) 및 기준값 조절:** 파일 업로드 시 데이터 유형·결측률·단일값 컬럼·중복 ID 자동 검증 배너 탑재 및 사이드바 품질 기준값(기본 1.5%) 실시간 슬라이더 구현
3. **다차원 시각화 및 리포트 정밀도 개선:** 전체 건수 무제한 조회(`592건 중 592건`), 단일값(100% 동일) 컬럼 차트 배제, 가로형 TOP 순위 차트 및 시나리오 클러스터링 적용

---

## 2. 프로젝트 폴더 및 핵심 파이썬(.py) 구조 (W47 완성본)
```text
📁 data_analysis_agent/
 ├─ app.py                        : 메인 웹 대시보드 (1~5단계 실시간 파이프라인, 자연어 SQL, 병렬 비교 UI)
 ├─ run.bat / main.py             : 원클릭 실행 배치 파일 및 CLI 실행 진입점
 ├─ docs/                         : 주차별 진행 보고서(W46, W47) 및 아키텍처 흐름도(README.md, HTML 뷰어)
 └─ data_agent/                   : AI 에이전트 핵심 엔진 패키지
     ├─ config.py                 : [W47 분리 완료] 테마 색상, 사내 기본 품질 기준값(1.5%), 모델 라벨 설정
     ├─ state.py                  : [W47 분리 완료] LangGraph 상태(AgentState) 스키마 정의
     ├─ llm.py                    : 멀티 AI 라우터 (로컬 Ollama / 클라우드 Groq·Gemini 자동 탐색)
     ├─ graph.py                  : LangGraph 워크플로우 (START → analyzer → reporter → END)
     ├─ excel_exporter.py         : 한글 기초통계·시나리오 군집화·색상 문맥 집계 및 차트 내장 엑셀 생성
     ├─ tools/
     │   ├─ data_loader.py        : 다중 시트·배경색 추출(Pandas 3.0 호환), 워드(.docx) 파싱, SQLite 동기화
     │   └─ visualizer.py         : [W47 분리 완료] 양식 자동 검증 배너, 4대 KPI 카드, 가로형 TOP 차트 엔진
     └─ nodes/
         ├─ analyzer_node.py      : [1단계] 파이썬 기반 통계·색상문맥·이상치 고속 계산 노드
         └─ reporter_node.py      : [2단계] TC/제조/일반 문서 맞춤형 4단락 리포트 생성 및 수치 붙음 방지 정제