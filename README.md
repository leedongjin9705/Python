# Python
Python을 사용하여 전자공시시스템 DART에서 기업의 정보 가져오기


https://opendart.fss.or.kr/uat/uia/egovLoginUsr.do 에서 기업정보를 가져온 후, 그 기업정보를 Python 스크립트를 사용해서
dart에 보내 기업개황을 불러와보자


1. 파이썬 설치하기
 https://www.python.org/downloads/ 사이트에 들어가서 파이썬 설치.
설치할 때 Add Python to PATH만 체크해주면 됩니다.

2. 작업폴더 만들기
 본인은 python이라고 만들어줬습니다.
 
3.  파이썬 라이브러리 설치하기
  cmd창에 pip install requests 입력

4. 회사고유번호 (CORPCODE.xml) 다운로드
  python폴더에 dart_corpcode.py 파일 생성.
  이하 해당 파일의 코드 및 주석 설명
 
#필요한 모듈 불러오기
from urllib.request import urlopen
from io import BytesIO
from zipfile import ZipFile
import xml.etree.ElementTree as ET

# 인증키 입력 - 꼭 본인 키로 바꾸기
my_api_key = '여기에_발급받은_API_키를_입력하세요'

# 회사 고유번호 파일 다운로드
url = f'https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={my_api_key}'

# 다운로드해서 압축풀기
with urlopen(url) as zip_response:
    with ZipFile(BytesIO(zip_response.read())) as zip_file:
        zip_file.extractall()  # 현재 폴더에 CORPCODE.xml 생김

print("CORPCODE.xml 다운로드 및 압축 해제 완료")

# XML 파일 불러오기
tree = ET.parse('CORPCODE.xml')  # 현재 폴더에 CORPCODE.xml 있어야 함
root = tree.getroot()

print("XML 파일 불러오기 완료")

# 회사 이름으로 고유번호 찾는 함수 만들기
def find_corp_code(company_name):
    for item in root.iter("list"):
        name = item.findtext("corp_name")
        code = item.findtext("corp_code")
        if name == company_name:
            return code
    return None
 
5. 실행하기
  cmd에서 해당 폴더 경로로 이동 
  cd C:\python\
  python dart_corpcode.py

생성된 파일 확인.

6. 기업개황 정보(company.json) 가져오기
   파일 수정.
   이하 수정된 전체 코드

`from urllib.request import urlopen
from urllib.parse import quote
from io import BytesIO
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import requests
import pandas as pd
import time

# API 인증키 입력
my_api_key = '여기에_너의_API_KEY_입력'

# CORPCODE.xml 다운로드 및 압축 해제
url = 'https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key=' + quote(my_api_key)
with urlopen(url) as zip_response:
    with ZipFile(BytesIO(zip_response.read())) as zip_file:
        zip_file.extractall()
print("CORPCODE.xml 다운로드 완료")

# XML 파싱
tree = ET.parse('CORPCODE.xml')
root = tree.getroot()
print("CORPCODE.xml 파싱 완료")

# 기업개황 불러오는 함수
def get_company_info(corp_code):
    url = f"https://opendart.fss.or.kr/api/company.json?crtfc_key={my_api_key}&corp_code={corp_code}"
    response = requests.get(url)
    return response.json()

# 10,000개 기업 정보 수집
company_list = []

for idx, item in enumerate(root.iter("list")):
    if idx >= 10000:  # 최대 10,000개로 제한
        break

    corp_code = item.findtext("corp_code")
    corp_name = item.findtext("corp_name")

    try:
        info = get_company_info(corp_code)
        if info.get("status") == "013":
            continue  # 없는 회사 코드일 경우 건너뜀

        company_list.append({
            "회사명": info.get("corp_name"),
            "대표자": info.get("ceo_nm"),
            "사업자등록번호": info.get("bizr_no"),
            "업종코드": info.get("industry_code"),
            "설립일": info.get("est_dt"),
            "법인구분": info.get("corp_cls"),
            "주소": info.get("adres"),
        })

        if idx % 100 == 0:
            print(f"▶ {idx} / 10000 기업 처리 중...")

        # API 과부하 방지를 위한 약간의 쉬는 시간 (0.1초)
        time.sleep(0.1)

    except Exception as e:
        print(f"오류 발생: {corp_name} → {e}")
        continue

print("모든 기업 개황 수집 완료!")`

# 엑셀 저장
df = pd.DataFrame(company_list)
df.to_excel("기업개황_전체_10000개.xlsx", index=False)
print("저장 완료: 기업개황_전체_10000개.xlsx")



API는 하루 최대 10000건이고, 가져 올 전체 건수는 8만건 정도 됩니다.
그래서 8일을 돌려야 합니다.

 7. 의문 해결
 내일 다시 이 스크립트를 실행한다면, 어떻게 내가 몇번째 줄까지 받은지 알려주지?

 처리된 corp_code 목록을 processed.txt에 저장 -> 처리된 기업은 건너뜀

8. 엑셀로 저장
   다운로드 받은 데이터를 엑셀에 저장하고 싶다.
    
   cmd에서 pip install pandas openpyxl 실행

9. 코드 수정
   이하 수정된 코드

from urllib.request import urlopen
from urllib.parse import quote
from io import BytesIO
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import requests
import pandas as pd
import time
import os

# API 인증키
my_api_key = '여기에_너의_API_KEY_입력'

# CORPCODE.xml 파일이 없으면 다운로드
if not os.path.exists('CORPCODE.xml'):
    url = 'https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key=' + quote(my_api_key)
    with urlopen(url) as zip_response:
        with ZipFile(BytesIO(zip_response.read())) as zip_file:
            zip_file.extractall()
    print("CORPCODE.xml 다운로드 완료")
else:
    print("CORPCODE.xml 이미 존재 (재다운로드 생략)")

# CORPCODE.xml 읽기
tree = ET.parse('CORPCODE.xml')
root = tree.getroot()

# 이전에 수집한 기업코드 불러오기
processed_file = "processed.txt"
if os.path.exists(processed_file):
    with open(processed_file, "r") as f:
        already_processed = set(f.read().splitlines())
else:
    already_processed = set()

# 기업개황 요청 함수
def get_company_info(corp_code):
    url = f"https://opendart.fss.or.kr/api/company.json?crtfc_key={my_api_key}&corp_code={corp_code}"
    response = requests.get(url)
    return response.json()

# 기업정보 수집 시작 (최대 10,000개)
company_list = []
count = 0
for item in root.iter("list"):
    corp_code = item.findtext("corp_code")
    corp_name = item.findtext("corp_name")

    if corp_code in already_processed:
        continue  # 이미 수집한 기업은 건너뜀

    try:
        info = get_company_info(corp_code)
        if info.get("status") == "013":
            continue  # 존재하지 않는 기업은 스킵

        company_list.append({
            "회사명": info.get("corp_name"),
            "대표자": info.get("ceo_nm"),
            "사업자등록번호": info.get("bizr_no"),
            "업종코드": info.get("industry_code"),
            "설립일": info.get("est_dt"),
            "법인구분": info.get("corp_cls"),
            "주소": info.get("adres"),
        })

        # 처리된 기업코드 저장
        with open(processed_file, "a") as f:
            f.write(corp_code + "\n")

        count += 1
        if count % 100 == 0:
            print(f"▶ {count} / 10000 수집 중...")

        if count >= 10000:
            break

        time.sleep(0.1)  # 과부하 방지

    except Exception as e:
        print(f"오류: {corp_name} → {e}")
        continue

print(f"오늘 수집 완료: 총 {count}개")

# 날짜 기반 파일명으로 저장
from datetime import datetime
today = datetime.today().strftime("%Y%m%d")
file_name = f"기업개황_{today}.xlsx"

df = pd.DataFrame(company_list)
df.to_excel(file_name, index=False)
print(f"저장 완료: {file_name}")


  10. 결과물
  CORPCODE.xml
![Image](https://github.com/user-attachments/assets/608b667a-8743-4d21-b5b5-d6c5d148ff7f)

 processed.txt

![Image](https://github.com/user-attachments/assets/ebe3bc36-8b9f-4140-a598-893146d2488f)

기업개황_20250430.xlsx

![Image](https://github.com/user-attachments/assets/a332f338-10e8-4313-b780-3cde4c4d45a9)
