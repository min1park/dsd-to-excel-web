📊 DART DSD to Excel 변환기 (Web App)

DART(전자공시시스템)에서 다운로드한 감사보고서(.dsd) 파일을 읽기 편한 Excel(.xlsx) 워크북으로 변환해 주는 Streamlit 기반 웹 애플리케이션입니다.

🚀 주요 기능

.dsd 파일 업로드 및 구조 분석

재무상태표, 포괄손익계산서, 자본변동표, 현금흐름표 자동 분리 및 Excel 시트화

원 단위 재무제표를 천원 단위로 자동 변환한 시트 추가 제공

개별 주석 및 감사 의견 등 문서 전체를 Excel로 깔끔하게 포맷팅

🛠️ 설치 및 실행 방법

저장소 클론 및 폴더 이동

git clone <이 저장소의 URL>
cd <생성된 폴더명>


필요한 라이브러리 설치

pip install -r requirements.txt


웹 애플리케이션 실행

streamlit run app.py


명령어 실행 후 브라우저에서 http://localhost:8501로 접속하여 사용할 수 있습니다.

📂 파일 구성

app.py: Streamlit 웹 인터페이스 구동 스크립트

dsd_to_excel.py: DSD 파일을 파싱하고 Excel로 변환하는 핵심 로직

requirements.txt: 구동에 필요한 Python 패키지 목록