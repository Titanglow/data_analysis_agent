# main.py (실행 진입점 코드 예시)
import sys
from streamlit.web import cli as stcli

if __name__ == '__main__':
    # main.py를 실행하면 자동으로 'streamlit run app.py' 명령을 대신 쳐줍니다.
    sys.argv = ["streamlit", "run", "app.py", "--global.developmentMode=false"]
    sys.exit(stcli.main())