import os
import smtplib
import time
import openpyxl as pyxl
from email import encoders
from email.utils import formataddr
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart


# ----------------------------------------------
#               메일 보내는 함수
# ----------------------------------------------
def SendMail_Function(row, email, pw, dir_path, file_type, Company):
    if file_type == "":
        #return -1     # <- 여기는 구분 라디오버튼 (급여, 상여, 연말정산) 선택 안하면 발생하는 오류 이벤트에 사용했던 것입니다.
        file_type = '.pdf'
    ToNum = str(row[1])             # 받는 사람 사번
    SendTo = row[5]                 # 받는 사람 이메일
    ToName = row[2]                 # 받는 사람 이름
    print(ToName, ToNum, SendTo)    # 테스트 용도 출력문 (이름, 사번, 메일주소)

    # from_addr : 보내는 사람 정보 (이름, 이메일)
    # to_addr : 받는 사람 정보 (이름, 이메일)
    from_addr = formataddr((Company, email)) 
    to_addr = formataddr((ToName, SendTo))

    session = None
    try:
        # SMTP 세션 생성 (포트번호 587)
        session = smtplib.SMTP('smtp.gmail.com', 587)
        #session.set_debuglevel(True)
        session.set_debuglevel(2)

        '''
        set_debuglevel(1 or True) : 서버로부터 만들어진 모든 연결과 메시지 전송-수신 디버깅 메시지 확인
        2 : 시간 값을 포함함 (timestamp)
        '''

        # SMTP 계정 인증 설정
        session.ehlo()      # ESMTP 서버에 식별
        session.starttls()  # 전송계층보안 TLS 적용
        session.ehlo()      # 적용 후 재호출
        session.login(email, pw)   # email: gmail 계정 아이디,   pw: 해당 계정의 "앱 비밀번호"
    
        # 메일 콘텐츠 설정
        '''
        multipart/mixed : 다른 content-type 헤더를 갖는 파일 전송에 사용
        일반 그림이나 파일이 아닌 "첨부 파일"에 해당
        '''

        message = MIMEMultipart("mixed")  
        
        # 메일 송/수신 옵션 설정
        message.set_charset('utf-8')
        message['From'] = from_addr
        message['To'] = to_addr

        # 선택 타입에 따라 전달되는 메일 제목
        # 연말정산 고정되게 변경한 상태
        message['Subject'] = file_type.split('.')[0] + "근로소득 명세서 발송"      
    

        '''
        MIMEText(Text, SubType='plain', charset=None)
        MiMEMultipart.attach() : 메시지에 서브 파트 첨부
        
        MIMEBase(MainType, SubType)
        maintype은 content-type의 주 유형(text, image 등)
        subtype은 content-type의 부 유형(plain, gif 등)

        MIME type 목록 : https://developer.mozilla.org/ko/docs/Web/HTTP/Basics_of_HTTP/MIME_types/Common_types
        application/octet-stream : 모든 종류의 이진 데이터


        'Content-Disposition', 'attachment', filename
        헤더 정보에 들어가는 항목이며 Content-Disposition/attachment는 전달되는 값을 다운로드 받을 수 있다는 의미
        filename에 들어가는 형식은 (CHARSET=ASCII, LANGUAGE=NONE, VALUE)
        '''
        
        # 메일 콘텐츠 - 내용
        #body = "<h2>" + file_type + "관련 첨부파일을 전송합니다.</h2>"     # <- 이건 메일 전송 구분 (급여, 상여, 연말정산) 있을 때
        body = "<h2> 연말정산 근로소득 관련 첨부파일을 전송합니다.</h2>"    # <- 구분 안하고 연말정산 먼저 테스트 할 거니까 이거!
        
        bodyPart = MIMEText(body, 'html', 'utf-8')   
        message.attach(bodyPart)     # body 부분은 받는 메일의 텍스트로 들어갑니다.
    
        try:
            ImportFile = Company + '_' + ToNum + '_' + ToName + file_type  # 업체 규칙에 맞게 첨부 파일 이름 지정 요망 (현재 규칙 [업체명]_[사번]_[이름].pdf)
            print(ImportFile)
            file_path = dir_path
            File_Name = os.path.join(file_path, ImportFile)   # 첨부 파일 위치 절대 경로 (설정한 폴더 경로를 바탕으로 세팅)
            print(f"File_Name = {File_Name}")
            if os.path.isfile(File_Name):
                Binary_File = MIMEBase("application", "octect-stream")
            
                # Binary Format 변환 및 헤더에 정보 담기
                with open(File_Name, "rb") as f:                # 첨부 파일 절대 경로에 있는 pdf를 처리한다.        
                    Binary_File.set_payload(f.read())           # 전송 데이터(페이로드)를 인코딩 형식으로 변환
                    encoders.encode_base64(Binary_File)         # Base64 포맷으로 인코딩
                    ImportFile = os.path.basename(File_Name)    # 폴더 경로와 파일명을 분리합니다. (위의 ImportFile을 그대로 가져와도 되지만, 혹시 모르니...)
                    print(f"Import File = {ImportFile}")
                    Binary_File.add_header('Content-Disposition', 'attachment', filename=('utf-8', '', ImportFile))
                    message.attach(Binary_File)
            else:
                return -2
        except Exception as e:
            print(e)
            return -3

        # --------------------------------------------
        #     session 연결된 상태에서 메일을 전송!
        # --------------------------------------------
        session.sendmail(from_addr, to_addr, message.as_string())
        #print(message.as_string())     # 첨부파일 전송되는지 확인을 위한 출력     
        print(f'메일 전송 완료 : {to_addr}')
        return 1
            
    except Exception as e:
        print(e)
        return -4
    finally:
        if session is not None:
            session.quit()




# -----------------------------------------------
#      여기는 천보기의 테스트 공간입니다.....
#    크게 상관하실 필요가 없는 낙서장 같은거죠..ㅎ
# -----------------------------------------------
if __name__ == '__main__':
    start = time.time()
    
    # 이메일 보낼 대상 명단 가져오기
    wb = pyxl.load_workbook('./SendMail.xlsx', data_only=True)
    ws = wb.active

    '''
    시트 명시해서 가져오려면 wb.get_sheet_by_name("시트명(sheet1)")
    '''

    # 행 1개씩 데이터 묶음(리스트)으로 가져오기
    for idx, row in enumerate(ws.iter_rows(min_row=0)):
        if row[idx].value == None:
            # 작업 처리 시간 계산 (업무 참고용)
            print(f"처리 완료. 작업 시간= ", round(time.time()-start,2), "초")
            os.system('pause')
        else:
            SendMail_Function(row)


# 501KB pdf 파일 3명 전송 = 56.96초 