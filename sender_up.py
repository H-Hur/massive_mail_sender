# -*- coding: cp949 -*- 
# Original file: https://underflow101.tistory.com/33
# 공대생의 차고
# 아래 라이브러리들은 파이썬 기본 내장 라이브러리이므로 별도의 설치가 필요 없습니다.

### Partly modified by Carrotday
### 네이버 메일용

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
### 메일 본문을 한글파일에서 읽어오려면 아래 라이브러리 설치가 필요합니다. 
import olefile


SMTP_SERVER = "smtp.naver.com"
SMTP_PORT = 465
SMTP_USER = "Naver_ID@naver.com"
SMTP_PASSWORD = "NAVER_PW"
sender = SMTP_USER 

# smtp로 접속할 서버 정보를 가진 클래스변수 생성
smtp = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
#smtp.ehlo()
if SMTP_PORT == 587: smtp.starttls()
# 해당 서버로 로그인
smtp.login(SMTP_USER, SMTP_PASSWORD)

# 만약 아래 메일 유효성 검사 함수에서 False가 나오면 메일을 보내지 않습니다.
def is_valid(addr):
    import re
    if re.match('(^[a-zA-Z-0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)', addr):
        return True
    else:
        return False
# 이메일 보내기 함수
def send_mail(sender,addr,subj,cont,attach1=None,attach2=None):
    if not is_valid(addr):
        print("Wrong email: " + addr)
        return
    
    # 텍스트 파일
    msg = MIMEMultipart("alternative")
    # 첨부파일이 있는 경우 mixed로 multipart 생성
    if attach1:
        msg = MIMEMultipart('mixed')
    msg["From"] = sender
    msg["To"] = addr
    msg["Subject"] = subj
    contents = cont
    text = MIMEText(_text = contents, _charset = "utf-8")
    msg.attach(text)

    if attach1:
        from email.mime.base import MIMEBase
        from email import encoders
        file_data = MIMEBase("application", "octect-stream")
        file_data.set_payload(open(attach1, "rb").read())
        encoders.encode_base64(file_data)
        import os
        filename = os.path.basename(attach1)
        file_data.add_header("Content-Disposition", 'attachment', filename=('UTF-8', '', filename))
        msg.attach(file_data)
    ### 두 번째 첨부파일이 있는 경우
    if attach2:
        file_data = MIMEBase("application", "octect-stream")
        file_data.set_payload(open(attach2, "rb").read())
        encoders.encode_base64(file_data)
        filename = os.path.basename(attach2)
        file_data.add_header("Content-Disposition", 'attachment', filename=('UTF-8', '', filename))
        msg.attach(file_data)    
    # 메일 발송
    smtp.sendmail(sender, addr, msg.as_string())

# 엑셀 파일에 정리된 명단으로 한꺼번에 보낼 때
# 아래 openpyxl 라이브러리는 외부 라이브러리이므로 pip3를 통해 설치 후 사용하시기 바랍니다.
from openpyxl import load_workbook
#ws = load_workbook('send_test.xlsx').active
ws = load_workbook('send.xlsx').active

for row in ws.iter_rows():
    addr = row[0].value
    subj = row[1].value

### 액셀에 본문을 그대로 넣으혀면 아래 두 줄을 쓰세요.
#    cont = row[2].value
#    cont_file= row[2].value
### 액셀에 한글파일 이름을 넣고 한글파일에 메일 본문을 넣으려면 위 두 줄 대신 아래  세 줄을 쓰세요
    file=olefile.OleFileIO( row[2].value)
    encoded_text = file.openstream('PrvText').read()
    cont = encoded_text.decode('UTF-16')

### 기본 첨부파일은 1개(D열), 최대 2개(E열)
    attach1 = row[3].value
### 액셀 E 열에 스페이스가 있으면 오류 발생 
    if  row[4].value != '': attach2 = row[4].value 
    if attach2 != '': send_mail(sender, addr, subj, cont, attach1, attach2) 
    else: send_mail(sender, addr, subj, cont, attach1)

# 닫기
smtp.close()

