import os
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials

# 1. .env 파일 로딩 (경로 주의!)
load_dotenv(dotenv_path="./config/.env")

# 2. 환경변수에서 값 읽기
SERVICE_ACCOUNT_FILE = os.getenv("GSHEET_SERVICE_ACCOUNT")
SHEET_URL = os.getenv("GSHEET_URL")
WORKSHEET_NAMES = os.getenv("GSHEET_WORKSHEETS").split(",")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    # "https://www.googleapis.com/auth/drive.readonly",
]

# 3. 인증 및 gspread 연결
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
gc = gspread.authorize(creds)

# 4. 구글 시트 열기 및 워크시트 지정
sh = gc.open_by_url(SHEET_URL)

# 확인: 구글 시트가 열렸는지 확인
if not sh:
    raise ValueError("구글 시트를 열 수 없습니다. URL을 확인하세요.")

# 5. 워크시트 순회하며 데이터 읽기
for ws_name in WORKSHEET_NAMES:
    ws_name = ws_name.strip()  # 공백 제거
    worksheet = sh.worksheet(ws_name)
    data = worksheet.get_all_values()
    print(f"\n===== 워크시트: {ws_name} =====")
    for row in data:
        print(row)
