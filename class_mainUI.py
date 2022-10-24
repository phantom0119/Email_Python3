import tkinter as tk   # UI 구성
from tkinter.constants import ANCHOR   # UI 구성 도구
import tkinter.font as tkFont          # UI 구성 폰트
from tkinter import Image, image_names, ttk, filedialog   # UI 구성 세부
from datetime import datetime          # 날짜 처리
import socket  # DB 소켓 통신
from types import new_class                          
import pymssql                         # MSSQL 객체 (연결용)
from pymssql import _mssql  # pyinstaller 적용
import uuid                 # pyinstaller 적용
import 급여명세xl           # 명세 파일 작성 (테스트용)
import sendmail             # 메일 보낼 때 쓰는 Py
import os                   # 시스템 관리

"""
 [연말정산] 대상 메일 시스템 개발 과정   -2021년 12월 31일 박천복-
  급여/상여/연말정산 대상으로 구현하려 했지만, 지금은 연말정산만 서비스 할 수 있도록 배치

 exe 실행 프로그램으로 만드려면 pyinstaller 명령을 사용합니다.
 -w : 콘솔 윈도우를 표시하지 않습니다.
 --onefile : exe 파일 하나로 모든 기능 적용할 수 있게 합니다. (다만 이미지 파일은 처리 못해서 같이 두어야 합니다.)
 --icon : exe 프로그램 아이콘을 생성합니다. 같은 경로의 favicon.ico 파일을 사용합니다. (ico 파일만 세팅 가능합니다.)
 -n : exe 프로그램으로 만들 때 이름을 지정합니다. (이거 세팅 안하고 exe 파일 만들고서 이름 변경해도 됩니다.)
.\class_mainUI.py : exe 파일 만들 python 코드 파일

 ->  pyinstaller -w --onefile --icon ./favicon.ico -n SendMail .\class_mainUI.py


명령 처리 결과는 동일 경로 내의 dist 폴더에 생성됩니다.
dist 폴더에 있는 exe 프로그램을 그대로 실행하면 오류가 발생할 것입니다.
프로그램에 사용되는 이미지 파일을 모두 exe 프로그램과 동일한 경로에 배치해 주어야 합니다.

사용한 이미지 파일 : connect, constart, link_folder, notconnect, search, sendmail
"""


# -----------------------------------------
#     Mail Program 본체를 구성하는 내용
# -----------------------------------------
class MailProgram(tk.Tk):

    # -----------------------------------
    #   생성자 (Mail_Interface 호출)
    # -----------------------------------
    def __init__(self, master):
        tk.Tk.__init__(self, master)
        self.master = master
        self.Mail_Interface()



    # ----------------------------------------------
    #           실행 프로그램 뼈대입니다.
    #   Delphi Main Form과 같다고 볼 수 있겠네요
    # ----------------------------------------------
    def Mail_Interface(self):
        self.grid()
        self.select_cnt = 0  # 선택 인원 수
        self.file_type = ""  # 첨부파일 초기 값
        self.program_log = os.getcwd() + '/SendMail_Info.txt'   # 메일 로그인 기록 저장파일 경로
        self.DB_log = os.getcwd() + '/DB_Info.txt'              # DB 로그인 정보 저장 파일 경로
    

        # -----------------------------------------------------------------
        #                              폰트 정의
        #          family로 폰트 종류 정의, size로 글자 크기 조절
        #    한글 글자 설정하는 것처럼 다양한 속성 있으니 기호에 맞게...!
        # -----------------------------------------------------------------
        self.label_font = tkFont.Font(family="Arial", size=13, weight="bold")     # 제목 라벨
        self.radio_font = tkFont.Font(family="MS Serif", size=13, weight="bold")  # 라디오 버튼
        self.combo_font = tkFont.Font(family="Arial", size=11, weight="bold")     # 콤보박스
        self.table_font = tkFont.Font(family="Arial", size=10, weight="bold")     # 테이블 구성 항목
        self.login_font = tkFont.Font(family="MS Serif", size=12)                 # 로그인 항목


        # ----------------------------------------------------------------------------------------------------
        #                                   배치 공간 설정 (LabelFrame)
        #       별도의 프레임 공간을 만드는 녀석입니다. 반드시 필요하진 않지만 사용해볼 필요는 있었습니다.
        #    년월에 콤보박스 주변으로 실선 그어져 있는 거 보이시죠? 그 공간이 여기서 정의한 LabelFrame 입니다.
        # ----------------------------------------------------------------------------------------------------
        self.radio_label_frame = tk.LabelFrame(self, relief="solid", bd=1)
        self.radio_label_frame.place(x=70, y=10)    # 첨부파일 종류 설정 라디오버튼  (현재는 폴더 경로 공간으로 사용)
        self.combo_label_frame = tk.LabelFrame(self, relief="solid", bd=1, padx=3)
        self.combo_label_frame.place(x=70, y=60)    # 연월 콤보박스 설정 공간
    

        # ------------------------------------------------------------------
        #                           Label, Entry List
        #           text에 적혀있는 문구 보이시죠? 그거 구성하는 녀석들입니다.
        #     relief는 선명함/두드러짐의 의미를 가지고 있습니다.  --> 실선 테두리 디자인 옵션
        #                   bd = borderwidth    --> 테두리 두께
        # ------------------------------------------------------------------
        #self.label1 = tk.Label(self, text="구분", width=5, height=2, relief="solid", bg="#dcdcdc", bd=0.5)
        self.label1 = tk.Label(self, text="경로", width=5, height=2, relief="solid", bg="#dcdcdc", bd=0.5)
        self.label2 = tk.Label(self, text="년월", width=5, height=2, relief="solid", bg="#dcdcdc", bd=0.5)
        self.label3 = tk.Label(self, text="서버 연결 설정", width=12, height=2, relief="solid", bg="#dcdcdc", bd=0.5)
        self.label4 = tk.Label(self.combo_label_frame, text="년", width=2, height=2)
        self.label5 = tk.Label(self.combo_label_frame, text="월", width=2, height=2)
        self.label6 = tk.Label(self, text="선택인원", width=10, relief="solid", bg="#d2d2ff", bd=0.5)
        self.label7 = tk.Label(self, text="명", width=1)
        self.label8 = tk.Label(self, text="정렬옵션", width=10, relief="solid", bg="#d2d2ff", bd=0.5)
        self.label9 = tk.Label(self, text="메일 전송 계정(gmail)", width=20)
        self.label10 = tk.Label(self, text="ID :", width=6)
        self.label11 = tk.Label(self, text="PW :", width=6)

        self.entry1 = tk.Entry(self, width=15)              # 메일 전송 계정 ID 입력 공간
        self.entry2 = tk.Entry(self, show="*", width=15)    # 메일 전송 계정 PW 입력 공간
    


        # ---------------------------------
        #            Text List
        #       텍스트를 다루는 공간
        # ---------------------------------
        self.text1 = tk.Text(self, width=4, height=1, relief="solid", bd=0.5)   # 선택 인원 출력 (명)
        self.text2 = tk.Text(self, width=15, height=1, relief="solid", bd=0.5, wrap="none", spacing1=9, spacing3=9) # 서버 연결 상태 출력
        self.text3 = tk.Text(self, width=50, height=8, relief="solid", bd=0.5)  # 상태 표시 텍스트
        self.text5 = tk.Text(self.radio_label_frame, width=25, height=1, relief="solid", bd=0.5, wrap="none", spacing1=9, spacing3=9) # 설정된 폴더 경로 출력




        # ---------------------------------
        #       Radio Button List
        #        라 디 오 버 튼 
        # ---------------------------------
        '''
        self.radioset1 = tk.IntVar()
        self.radio1 = tk.Radiobutton(self.radio_label_frame, text="급여",
                        variable=self.radioset1, value=1, pady=2, command=lambda: self.Radio_Call1())
        self.radio2 = tk.Radiobutton(self.radio_label_frame, text="상여",
                        variable=self.radioset1, value=2, pady=2, command=lambda: self.Radio_Call1())
        self.radio3 = tk.Radiobutton(self.radio_label_frame, text="연말정산",
                        variable=self.radioset1, value=3, pady=2, command=lambda: self.Radio_Call1())
        '''
        self.radioset2 = tk.IntVar()
        self.radio4 = tk.Radiobutton(self, text="사번", variable=self.radioset2, value=1,
                        command=lambda: self.Radio_Call2())
        self.radio5= tk.Radiobutton(self, text="사원명", variable=self.radioset2, value=2,
                        command=lambda: self.Radio_Call2())
        self.radio6 = tk.Radiobutton(self, text="부서명", variable=self.radioset2, value=3,
                        command=lambda: self.Radio_Call2())




        # --------------------------------------
        #              ComboBox List
        #      년/월에 사용하는 콤 보 박 스
        # --------------------------------------
        Date = datetime.now()       # 현재 시간
        Year = Date.strftime("%Y")  # 올해
        self.years = [int(Year)-1, int(Year), int(Year)+1]      # [년] 콤보박스 구성
        self.months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]   # [월] 콤보박스 구성
        self.year_combo = ttk.Combobox(self.combo_label_frame, width=5, values=self.years)
        self.year_combo.current(self.years.index(datetime.today().year))
        self.month_combo = ttk.Combobox(self.combo_label_frame, width=2, values=self.months)
        self.month_combo.current(self.months.index(datetime.today().month))


        # -----------------------------------
        #         CheckButton List
        #            체 크 박 스
        # -----------------------------------
        self.chkvalue = tk.BooleanVar() # 전체 선택
        self.chkvalue.set(False)
        self.chkvalue2 = tk.BooleanVar() # 전송 정보 저장 유무
        self.checkbox1 = ttk.Checkbutton(self, text="전체선택", variable=self.chkvalue,
                            onvalue=True, offvalue=False, command=lambda: self.Check_Call1())
        self.checkbox2 = ttk.Checkbutton(self, text="전송 정보 저장", variable=self.chkvalue2,
                            onvalue=True, offvalue=False)

        # -----------------------------------------------------------------------------------
        #                                    TreeView
        #        데이터 테이블 - sql 처리 결과를 표현하기 위한 공간을 정의하고 구성합니다.
        # -----------------------------------------------------------------------------------
        self.view_table = ["선택", "사번", "사원명", "부서명", "직급", "이메일"]
        self.treeview = ttk.Treeview(self, columns=self.view_table, displaycolumns=self.view_table, height=20)
        self.treeview['show'] = "headings"

        # 각 Column 크기 조절
        self.treeview.column("선택", width=50)
        self.treeview.column("사번", width=100)
        self.treeview.column("사원명", width=100)
        self.treeview.column("부서명", width=350)
        self.treeview.column("직급", width=100)
        self.treeview.column("이메일", width=200)

        # 각 Column에 해당하는 탭 이름 지정
        # anchor = 공간에서 텍스트 배치 정렬 (좌 left, 중 center, 우 right)
        self.treeview.heading("선택", text="선택", anchor="center")
        self.treeview.heading("사번", text="사번", anchor="center")
        self.treeview.heading("사원명", text="사원명", anchor="center")
        self.treeview.heading("부서명", text="부서명", anchor="center")
        self.treeview.heading("직급", text="직급", anchor="center")
        self.treeview.heading("이메일", text="이메일", anchor="center")
    
        scroll1 = ttk.Scrollbar(self,command=self.treeview.yview)   # treeview 스크롤바
        self.treeview.configure(yscrollcommand=scroll1.set)         # y 스크롤 적용 (위아래 스크롤)
        # 태그가 더블클릭되면 이벤트 발생
        self.treeview.tag_bind("select_tag", sequence='<Double-1>', callback=self.treeview_select)


        # -----------------------------------
        #      Image, Button, Canvas List
        # -----------------------------------
        self.Search_Image = tk.PhotoImage(file="search.png")            # 검색(돋보기)
        self.Mail_Image = tk.PhotoImage(file="sendmail.png")            # 메일 전송
        self.Connect_Image = tk.PhotoImage(file="connect.png")          # 정상 연결(파랑)
        self.Not_Connect_Image = tk.PhotoImage(file="notconnect.png")   # 연결 실패 (초기)
        self.Con_Image = tk.PhotoImage(file="constart.png")             # 연결 버튼 (플러그)
        self.Link_Fold_Image = tk.PhotoImage(file="link_folder.png")    # 폴더 경로

        Search_Button = tk.Button(self, image=self.Search_Image, width=38, height=38,
                    overrelief="groove", command=lambda: self.Get_Table())
        Send_Mail_Button = tk.Button(self, text="전송", image=self.Mail_Image, compound="top", width=45, height=45,
                    overrelief="groove", command=lambda: self.Test_Sending())
        Connection_Button = tk.Button(self, text="연결", image=self.Con_Image, compound="top", width=45, height=45,
                    overrelief="groove", command=lambda: self.IP_Connection())
        LinkF_Button = tk.Button(self, text="폴더\n경로", image=self.Link_Fold_Image, compound="top", width=50, height=50,
                    overrelief="groove", command=lambda: self.Find_Folder())

        self.Canvas1 = tk.Canvas(self, width=28, height=28)             # 시스템 상태 표시 공간
        self.Canvas1.create_image(12, 12, image=self.Not_Connect_Image) # 비연결 상태 이미지


        # -------------------------------------------------------
        #                 Settings & Place
        #          configure는 구성요소 정의하는 부분 
        #     (여기에는 font 사용에 대한 것만 정의했습니다.)
        # -------------------------------------------------------
        self.label1.configure(font=self.label_font)
        self.label2.configure(font=self.label_font)
        self.label3.configure(font=self.label_font)
        self.label4.configure(font=self.combo_font)
        self.label5.configure(font=self.combo_font)
        self.label6.configure(font=self.table_font)
        self.label7.configure(font=self.table_font)
        self.label8.configure(font=self.table_font)
        self.label9.configure(font=self.label_font)
        #self.radio1.configure(font=self.radio_font)
        #self.radio2.configure(font=self.radio_font)
        #self.radio3.configure(font=self.radio_font)
        self.year_combo.configure(font=self.combo_font)
        self.month_combo.configure(font=self.combo_font)
        self.text1.configure(font=self.table_font, state="disabled")  # 선택인원 수
        self.text2.configure(font=self.label_font)
        self.text5.configure(font=self.login_font)
        self.label10.configure(font=self.login_font)
        self.label11.configure(font=self.login_font)
        self.entry2.configure(font=self.login_font)
        self.entry1.configure(font=self.login_font)
    


        # ------------------------------------
        #            place는 배치 옵션
        #      x, y 좌표에 해당 구성요소 배치
        # ------------------------------------
        self.label1.place(x=10, y=10)    # 구분
        self.label2.place(x=10, y=60)    # 년월
        self.label3.place(x=10, y=110)   # 연결 IP 입력
        self.label4.pack(side="left")    # 년 레이블, Frame 안에 있기 때문에 배치 옵션 다름.
        self.label5.pack(side="left")    # 월 레이블, Frame 안에 있기 때문에 배치 옵션 다름.
        self.label6.place(x=80, y=170)   # 선택인원
        self.label7.place(x=205, y=170)  # 명
        self.label8.place(x=350, y=170)  # 정렬옵션
        self.label9.place(x=830, y=20)    # 메일 전송 계정
        self.label10.place(x=825, y=60)   # ID
        self.label11.place(x=825, y=85)   # PW
        

        self.text1.place(x=170, y=170)   # 선택인원 출력
        self.text2.place(x=143, y=111)   # 입력 IP 주소
        self.text3.place(x=400, y=10)    # 상태 출력 공간
        self.text5.pack(side = "left")   # 폴더 경로 텍스트, Frame 안에 있기 때문에 배치 옵션 다름.
        
        #self.radio1.pack(side="left")        # 급여 라디오 버튼
        #self.radio2.pack(side="left")        # 상여 라디오 버튼
        #self.radio3.pack(side="left")        # 연말정산 라디오 버튼

        self.entry1.place(x=880, y=60)    # ID 입력 공간
        self.entry2.place(x=880, y=85)    # PW 입력 공간
        self.checkbox1.place(x=240, y=170)   # 전체선택 체크박스
        self.checkbox2.place(x=840,y=120) # 전송 정보 저장


        self.radio4.place(x=450, y=170)      # 정렬옵션 라디오 버튼1
        self.radio5.place(x=500, y=170)      # 정렬옵션 라디오 버튼2
        self.radio6.place(x=562, y=170)      # 정렬옵션 라디오 버튼3
        
        self.year_combo.pack(side="left")    # 년 콤보박스
        self.month_combo.pack(side="left")   # 월 콤보박스

    
        Search_Button.place(x=230, y=60)        # 검색 버튼
        Send_Mail_Button.place(x=1010, y=57)    # 메일 전송 버튼
        LinkF_Button.place(x=280, y=10)         # 폴더 경로 버튼
        Connection_Button.place(x=290, y=105)   # IP 연결 버튼
        
        self.Canvas1.place(x=30, y=170)          # 연결 상태 이미지
        self.treeview.place(x=80, y=200)         # 데이터 테이블
        scroll1.place(x=990, y=200, height=430)  # 테이블 스크롤바



        if os.path.isfile(self.DB_log):
            try:
                self.Load_DB_log()
                print("DB Load Success")
            except:
                pass

        # -----------------------------------------------------------------------------------------------------
        #                          여기는 Gmail 로그인 정보 가져오기 위한 부분
        #               프로그램 실행 경로에 "SendMail_Info.txt" 파일 있으면 조회합니다.           
        #     체크박스 상태 (True,False), 첨부파일이 저장되어있던 경로, 접속한 IP, Gmail 계정, Gmail PW 순서
        # -----------------------------------------------------------------------------------------------------
        if os.path.isfile(self.program_log):
            # 첨부파일 폴더 경로, IP 주소, 메일 ID, 메일 PW 순서
            try:
                with open(self.program_log, 'r', encoding='utf-8') as f:
                    logs = f.readlines()
                    if len(logs) == 0:  # 파일 내용 비어있으면 사용 안 함.
                        pass
                    else:
                        mail_boolean = self.Text_En_Decryption(logs[0].split('\n')[0], 2) # Boolean 값 가져오기
                        #print(boolean)
                        self.chkvalue2.set(mail_boolean)  # Boolean 결과를 프로그램에 반영
                        
                        # 마지막으로 사용한 상태에서 "전송 정보 저장" 체크박스를 선택했다면,
                        # 실행 프로그램과 같은 경로에 있는 SendMail_Info.txt 파일에서 정보 읽음
                        if mail_boolean == "True":
                            self.dir_path = self.Text_En_Decryption(logs[1].split('\n')[0], 2)
                            self.text2.insert(1.0, self.Text_En_Decryption(logs[2].split('\n')[0], 2))
                            self.entry1.insert(0, self.Text_En_Decryption(logs[3].split('\n')[0], 2))
                            self.entry2.insert(0, self.Text_En_Decryption(logs[4].split('\n')[0], 2))

                            # 이전 정보를 가져왔다는 안내 Text를 표시해줌
                            self.text3.configure(state="normal")
                            self.text3.insert("end", '\n' + "이전 Gmail 로그인 정보를 가져왔습니다.")
                            self.text3.configure(state="disabled")
                        f.close()
            except:
                print("Load Error")
    
                


    # ----------------------------------------------------------------
    #      전송 정보 저장 로깅 (시저 암호 방식으로 텍스트를 저장)
    # ----------------------------------------------------------------
    def Save_Con_Log(self):
        self.txtfile = open(self.program_log, 'w+', encoding='utf-8')
        self.txtfile.seek(0) # 첫 줄부터 탐색
        # 여기는 gmail 로그인 정보 저장 구문
        self.txtfile.write(self.Text_En_Decryption(str(self.chkvalue2.get()) , 1) + '\n')       # 체크박스 선택 유무
        self.txtfile.write(self.Text_En_Decryption(self.dir_path , 1) + '\n')                   # 마지막 폴더 경로
        self.txtfile.write(self.Text_En_Decryption(self.text2.get(1.0, "end-1c") , 1) + '\n')   # 마지막 접속 IP
        self.txtfile.write(self.Text_En_Decryption(self.entry1.get(), 1) + '\n')                # Gmail ID
        self.txtfile.write(self.Text_En_Decryption(self.entry2.get() , 1) + '\n')               # Gmail PW
        self.txtfile.close()    # 기록 후 파일 닫기


    # --------------------------------------------------------
    #         DB 로그인 기록 저장 (시저 암호 방식)
    # --------------------------------------------------------
    def Save_DB_Log(self):
        self.txtfile = open(self.DB_log, 'w+', encoding='utf-8')
        self.txtfile.seek(0)
        # 여기는 DB 연결 정보 저장 구문
        self.txtfile.write(str(self.chkvalue3.get()) + '\n')  # 체크박스 선택 유무
        self.txtfile.write(str(self.entry3.get()) + '\n')     # 서버 IP
        self.txtfile.write(str(self.entry4.get()) + '\n')     # 서버 Port
        self.txtfile.write(str(self.entry5.get()) + '\n')     # DB ID
        self.txtfile.write(str(self.entry6.get()) + '\n')     # DB PW
        self.txtfile.close()  # 기록 후 파일 닫기




    # -----------------------------------------------------------------------------------------
    #                           이전 DB 서버 로그인 기록 가져오기
    #               같은 경로의 "DB_Info.txt" 파일이 있을 때 조회됩니다.
    #    원격 DB 백업할 때 사용하는 conf 파일처럼 수정해서 작업할 수 있도록 만들어봤습니다.
    # ------------------------------------------------------------------------------------------
    def Load_DB_log(self):
        try:
            with open(self.DB_log, 'r', encoding='utf-8') as f:
                logs = f.readlines()
                if len(logs) == 0:
                    pass  
                else:
                    DB_boolean = logs[0].split('\n')[0]
                    print(f"0. {DB_boolean}")

                    if DB_boolean:
                        ip_value = logs[1].split('\n')[0]
                        port_value = logs[2].split('\n')[0]
                        id_value = logs[3].split('\n')[0]
                        pw_value = logs[4].split('\n')[0]

                        #print(self.entry3.get())
                        print("Strat DB Connecting")
                        self.Connection = pymssql.connect(host=ip_value, port=port_value, user=id_value, password=pw_value, charset='EUC-KR')  # DB 연결 시도
                        self.cursor = self.Connection.cursor()

                        # 연결 성공 시 출력 이미지 교체
                        self.Canvas1.delete("all")
                        self.Canvas1.create_image(16, 14, image=self.Connect_Image)

                        # 이전 DB 연결 정보를 성공적으로 가져왔다는 Text를 출력합니다.
                        self.text3.configure(state="normal")
                        self.text3.insert("end",'\n' + "이전 DB 연결 정보를 가져왔습니다.")
                        self.text3.configure(state="disabled")
                    f.close()
        
        except:
            # 이전 DB 접속 기록을 가져오지 못했다는 Text를 출력합니다.
            print("DB 로드 중 오류 발생")
            self.text3.configure(state="normal")
            self.text3.bind("<Key>", lambda a: "break")
            self.text3.delete(1.0, "end")
            self.text3.insert(1.0, "DB 연결 정보를 가져오지 못했습니다.")
            self.text3.configure(state="disabled")

            self.Canvas1.delete("all")
            self.Canvas1.create_image(12, 12, image=self.Not_Connect_Image)


    
    # -------------------------------------------------------------------------------------
    #                 txt에 저장되는 텍스트 시저 암호 (암/복호화)
    #    key는 시저 암/복호 열쇠입니다. 값을 높이면 OverFlow로 실행이 안 될 수 있습니다.
    #        매개변수 type은 암/복호 옵션이며 1은 암호화, 2는 복호화로 했습니다.
    # -------------------------------------------------------------------------------------
    def Text_En_Decryption(self, Text, type:int):
        if type == 1:
            key = 88
        elif type == 2:
            key = -88

        encryp_str = ""
        for s in Text:
            tmp = ord(s) + key
            encryp_str += chr(tmp)
        return encryp_str
        


    # -----------------------------
    #       파일 폴더 경로 설정
    # -----------------------------
    def Find_Folder(self):
        self.dir_path = filedialog.askdirectory(parent=self, initialdir='/', title='문서 파일 경로 지정')
        
        self.text3.configure(state="normal")
        self.text3.bind("<Key>", lambda a: "break")
        self.text3.delete(1.0, "end")
        self.text3.insert(1.0, "폴더 경로가 설정되었습니다.\n경로=" + self.dir_path)
        self.text3.configure(state="disabled")

        self.text5.configure(state="normal")
        self.text5.bind("<Key>", lambda a: "break")
        self.text5.delete(1.0, "end")
        self.text5.insert(1.0, self.dir_path)
        self.text5.configure(state="disabled")

        #print(self.dir_path)



    # -------------------------
    #      메일 보내기 처리
    # -------------------------
    def Test_Sending(self):
        selections = self.Get_Selection()   # 테이블 행 정보 리스트
        email = self.entry1.get()           # ID에 적힌 email 정보
        pw = self.entry2.get()              # PW에 적힌 PW 정보

        # 사업장 정보 가져옴
        Company = ''        
        sql = "select c002 from HPUBLIC00.dbo.T0104 where C009 ='1';"
        try:
            Company = self.sql_query(sql, 1)  # sql 실행
            Company = Company[0].strip()
        except:
            Company = ""



        if email != "" and pw != "":
            # 선택, 사번, 이름, 부서명, 직급, 이메일
            # selections는 sql 조회 결과인 모든 행 테이블, select는 1개의 행으로 이해하면 됩니다.
            for select in selections:  
                #print(f"select = {select}")
                select_value = self.treeview.item(select).get('values')
                cnt = sendmail.SendMail_Function(select_value, email, pw, self.dir_path, self.file_type, Company)  # 메일 전송 함수


                # --------------------------------------------------------
                #      메일 전송 문제 생겼을 때 출력하는 Text 안내문
                # --------------------------------------------------------
                if cnt == -1:   # 구분 라디오버튼 미선택
                    self.text3.configure(state="normal")
                    self.text3.bind("<Key>", lambda a: "break")
                    self.text3.delete(1.0, "end")
                    self.text3.insert(1.0, "전송할 파일 형식을 지정해 주세요.")
                    self.text3.configure(state="disabled")
                elif cnt == -2: # 첨부파일이 있는 폴더 경로 미설정(경로 불일치)
                    self.text3.configure(state="normal")
                    self.text3.insert("end", "\n" +  f" {select_value[2]}님의 파일이 경로에 존재하지 않습니다. (전송 실패)")
                    self.text3.configure(state="disabled")
                elif cnt == -3: # 메일 전송 시스템 관련 오류
                    self.text3.configure(state="normal")
                    self.text3.bind("<Key>", lambda a: "break")
                    self.text3.delete(1.0, "end")
                    self.text3.insert(1.0, "메일 서버에 문제가 발생했습니다.")
                    self.text3.configure(state="disabled")
                elif cnt == -4: # 메일 전송용 계정 정보 불일치
                    self.text3.configure(state="normal")
                    self.text3.bind("<Key>", lambda a: "break")
                    self.text3.delete(1.0, "end")
                    self.text3.insert(1.0, "메일 전송에 실패했습니다.\n계정 정보 확인 부탁드립니다.")
                    self.text3.configure(state="disabled")                                    
                else:   # 정상 처리
                    self.text3.configure(state="normal")
                    self.text3.insert("end", "\n" + select_value[2] + "님(" + select_value[5] + ") 전송 완료.")
                    self.text3.configure(state="disabled")
                    self.Save_Con_Log()
                                
        else:   # ID나 PW에 값이 없는 경우
            #print("Please Insert ID and PW!!")
            self.text3.configure(state="normal")
            self.text3.bind("<Key>", lambda a: "break")
            self.text3.delete(1.0, "end")
            self.text3.insert(1.0, "메일 전송을 위한 ID와 PW 입력해주세요.")
            self.text3.configure(state="disabled")



    # -------------------------------------
    #   테이블에서 직원 더블클릭 이벤트
    # -------------------------------------
    def treeview_select(self, event):
        select_row = self.treeview.selection()[0]  # 더블클릭된 행 선택
        value = self.treeview.item(select_row).get('values')[0]   # values = 행 데이터 저장

        # 선택되어 있지 않다면 ★, 선택되어 있다면 ""
        if value == "":
            self.treeview.set(select_row, "선택", "★")
            self.select_cnt += 1
        else:
            self.treeview.set(select_row, "선택", "")
            self.select_cnt -= 1
            
        # 선택 시에 선택인원에 명수 반영
        self.text1.configure(state="normal")
        self.text1.bind("<Key>", lambda a: "break")
        self.text1.delete(1.0, "end")
        self.text1.insert(1.0, self.select_cnt)
        self.text1.configure(state="disabled")


    # -----------------------------------
    #       전체선택 체크박스 이벤트
    # -----------------------------------
    def Check_Call1(self):
        if self.chkvalue.get() == False:
            select_rows = self.treeview.get_children()
            for row in select_rows:
                self.treeview.set(row, "선택", "")
            self.select_cnt = 0   
        else:
            select_rows = self.treeview.get_children()
            for row in select_rows:
                self.treeview.set(row, "선택", "★")
            self.select_cnt = len(select_rows)

        self.text1.configure(state="normal")
        self.text1.bind("<Key>", lambda a: "break")
        self.text1.delete(1.0, "end")
        self.text1.insert(1.0, self.select_cnt)
        self.text1.configure(state="disabled")



    # -------------------------------------------------------------------
    #              급여, 상여, 연말정산 라디오 버튼 이벤트
    #     현재는 해당 라디오 버튼을 비활성화 했기 때문에 사용 안 함.
    # -------------------------------------------------------------------
    def Radio_Call1(self):
        value = self.radioset1.get()

        if value == 1:  # 급여
            self.file_type = "급여.pdf"
        elif value == 2: # 상여
            self.file_type = "상여.pdf"
        elif value == 3: # 연말정산
            self.file_type = "연말정산.pdf"
        else:
            self.file_type = ".pdf"
            
        #print(f"get file_type = {self.file_type}")


    # -----------------------------------------------------
    #     정렬(사번, 사원명, 부서명) 라디오 버튼 이벤트
    # -----------------------------------------------------
    def Radio_Call2(self):
        value = self.radioset2.get()
        if value == 1:  # 사번 On
            sorting_list = [(self.treeview.set(table, "사번"), table) for table in self.treeview.get_children()]
        elif value == 2:  # 사원명 On
            sorting_list = [(self.treeview.set(table, "사원명"), table) for table in self.treeview.get_children()]
        elif value == 3:  # 부서명 On
            sorting_list = [(self.treeview.set(table, "부서명"), table) for table in self.treeview.get_children()]
        
        sorting_list.sort()

        for idx, (selection, table) in enumerate(sorting_list):  # 위에 정렬 마치고 테이블 재구성하는 부분
            self.treeview.move(table, '', idx)



    # --------------------------------
    #      DB 연결 설정 페이지 호출
    # --------------------------------
    def IP_Connection(self):
        self.DB_Address = ''  # 연결 성공한 데이터 저장
        self.DB_Port = ''
        self.DB_ID = ''
        self.DB_PW = ''


        self.New_Window = tk.Toplevel(window)       # New_Window = DB 연결 설정 창 (새 창 열림)
        self.New_Window.title("DB 서버 연결 설정")  # 새 창의 이름
        self.New_Window.geometry("360x360")         # 창 크기 명시
        self.New_Window.resizable(False, True)      # 창 조절(가로 불가, 세로 허용)

        self.login_font = tkFont.Font(family="MS Serif", size=15)
        self.label12 = tk.Label(self.New_Window, text="서버 IP", width=20, height=2, relief="solid", bg="#dcdcdc", bd=0.5)
        self.label13 = tk.Label(self.New_Window, text="접속 Port", width=20, height=2, relief="solid", bg="#dcdcdc", bd=0.5)
        self.label14 = tk.Label(self.New_Window, text="DB 서버 ID", width=20, height=2, relief="solid", bg="#dcdcdc", bd=0.5)
        self.label15 = tk.Label(self.New_Window, text="DB 서버 PW", width=20, height=2, relief="solid", bg="#dcdcdc", bd=0.5)
        self.text4 = tk.Text(self.New_Window, width=30, height=2, relief="solid", bd=0.5)  # 상태 표시 텍스트

        self.entry3 = tk.Entry(self.New_Window, width=12)              # 서버 IP 입력 공간
        self.entry4 = tk.Entry(self.New_Window, width=12)              # 서버 Port 입력 공간
        self.entry5 = tk.Entry(self.New_Window, width=12)              # DB ID 입력 공간
        self.entry6 = tk.Entry(self.New_Window, width=12, show="*")    # DB PW 입력 공간

        Login_Button2 = tk.Button(self.New_Window, text="DB 접속", width=10, height=1, overrelief="groove",
                                command=lambda: self.Module_DB_Connectiong())

        self.chkvalue3 = tk.BooleanVar() # DB 연결 정보 저장 유무
        self.chkvalue3.set(True)  # 기본값 True (체크)
        self.checkbox3 = ttk.Checkbutton(self.New_Window, text="연결 정보 저장", variable=self.chkvalue3, onvalue=True, offvalue=False)



        self.label12.place(x=25, y=25)    # 서버 IP
        self.label13.place(x=25, y=70)    # 접속 Port
        self.label14.place(x=25, y=115)   # DB 서버 ID
        self.label15.place(x=25, y=160)   # DB 서버 PW
        self.entry3.place(x=190, y=30)    # 서버 IP 입력 공간
        self.entry4.place(x=190, y=75)    # 접속 Port 입력 공간
        self.entry5.place(x=190, y=120)   # DB 서버 ID 입력 공간
        self.entry6.place(x=190, y=165)   # DB 서버 PW 입력 공간
        Login_Button2.place(x=210, y=220)  # 로그인 버튼
        self.checkbox3.place(x=45, y=220) # 체크박스
        self.text4.place(x=65, y=280)     # 상태 출력 공간. DB 연결 설정하는 공간에서 씀(큰 역할은 아직).


        self.entry3.configure(font=self.login_font)
        self.entry4.configure(font=self.login_font)
        self.entry5.configure(font=self.login_font)
        self.entry6.configure(font=self.login_font)
        
        # ---------------------------------------------------
        #      DB_Info.txt 파일 있을 때 값 가져올지 판단
        # ---------------------------------------------------
        if os.path.isfile(self.DB_log):
            try:
                with open(self.DB_log, 'r', encoding='utf-8') as f:
                    logs = f.readlines()
                    if len(logs) == 0:
                        pass  
                    else:
                        DB_boolean = logs[0].split('\n')[0]

                        if DB_boolean:
                            self.entry3.insert(0, logs[1].split('\n')[0])
                            self.entry4.insert(0, logs[2].split('\n')[0])
                            self.entry5.insert(0, logs[3].split('\n')[0])
                            self.entry6.insert(0, logs[4].split('\n')[0])
            except:
                pass

        self.New_Window.mainloop()
        
        
        
    # --------------------------------------------------
    #       DB 연결 입력 값에 대한 연결 시도
    # --------------------------------------------------
    def Module_DB_Connectiong(self):
        address = self.entry3.get()   # DB 서버 연결 설정 화면의 값들 가져오기
        port = self.entry4.get()
        id = self.entry5.get()
        pw = self.entry6.get()

        print("printing " + address + " " + port + " " + id + " " + pw)

        try:
            self.Connection = pymssql.connect(host=address, port=port, user=id, password=pw, charset='EUC-KR')  # DB 연결 시도
            self.cursor = self.Connection.cursor()  # 연결 완료 되면 SQL 실행결과 반영해주는 커서

            self.text3.configure(state="normal")          # text 입력 허용
            self.text3.bind("<Key>", lambda a: "break")   # 특정 키 입력하면 Break (허용 거부)
            self.text3.delete(1.0, "end")                 # text의 첫 줄(1.0)부터 가장 마지막 줄(end)까지 삭제
            self.text3.insert(1.0, address + " 연결에 성공했습니다!!")  # 첫 줄에 Text 입력
            self.text3.configure(state="disabled")                     # Text 임의 수정 불가 상태

            self.text2.configure(state="normal")
            self.text2.bind("<Key>", lambda a: "break")
            self.text2.delete(1.0, "end")
            self.text2.insert(1.0, "서버 연결 완료")
            self.text2.configure(state="disabled")
            
            # 연결 성공 시 출력 이미지 교체
            self.Canvas1.delete("all")
            self.Canvas1.create_image(16, 14, image=self.Connect_Image)

            # 로그인 성공한 정보는 저장한다.
            self.DB_Address = address  
            self.DB_Port = port
            self.DB_ID = id
            self.DB_PW = pw
            self.Save_DB_Log()
            self.New_Window.destroy()  # 연결 성공하면 창 닫음.
            #print("IP Server is Connected")
        except:
            #print("IP Server is Not Connected")

            self.text3.configure(state="normal")
            self.text3.bind("<Key>", lambda a: "break")
            self.text3.delete(1.0, "end")
            self.text3.insert(1.0, address + " 연결에 실패했습니다.")
            self.text3.configure(state="disabled")


            self.text2.configure(state="normal")
            self.text2.bind("<Key>", lambda a: "break")
            self.text2.delete(1.0, "end")
            self.text2.insert(1.0, "서버 연결 필요!!")
            self.text2.configure(state="disabled")
            
            # 연결 실패 시 출력 이미지 교체
            self.Canvas1.delete("all")
            self.Canvas1.create_image(12, 12, image=self.Not_Connect_Image)

        #print("printing" + address + " " + port + " " + id + " " + pw)
        
        '''    ↓ 요건 9월의 천보기가 업체 DB 서버 IP와 Port를 제대로 몰랐던 시기에 개발했던
               ↓ 일반화의 오류 흔적. 이것도 곧 본인의 역사이기에...남겨두겠습니다.
        DB_Server = ""
        if self.text2.get("1.0", "end-1c") == "local":
            DB_Server = socket.gethostbyname(socket.getfqdn()) # 현재 접속 IP
            print(DB_Server)
        else:
            DB_Server = self.text2.get("1.0", "end-1c") # 개행문자 없이 받아오기
            IP_Part1 = DB_Server.split('.')[0]
            if IP_Part1 == '119':
                DB_Server = DB_Server + ':1455'
        '''
        
        


    # -----------------------------------
    #         SQL 쿼리 실행용 함수
    # -----------------------------------
    def sql_query(self, sql, type:int):
        try:
            self.cursor.execute(sql)

            if type == 1:   # fetchone
                return self.cursor.fetchone()
            elif type == 2: # fetchall
                return self.cursor.fetchall()
            else:           # sql commit
                self.Connection.commit()
        except:
            return -1
    


    # ----------------------------
    #     검색 시 테이블 초기화
    # ----------------------------
    def Table_Clear(self):
        self.treeview.delete(*self.treeview.get_children())
        
        self.text1.configure(state="normal")
        self.text1.bind("<Key>", lambda a: "break")
        self.text1.delete(1.0, "end")
        self.text1.insert(1.0, 0)
        self.text1.configure(state="disabled")

        self.chkvalue.set(False)




    # ------------------------------
    #      테이블에 데이터 표시
    # ------------------------------
    def Get_Table(self):
        # c001:사번, c002:이름, c004:부서코드, c008:직급코드, c010:입사일자, c013:퇴직일자
        select_year = self.year_combo.get()     # 콤보박스에 선택한 연
        select_month = self.month_combo.get()   # 콤보박스에 선택한 월
        self.Table_Clear()
        return_table = []

        # 퇴직일자가 검색한 년보다 낮은 경우 (퇴직처리 된 인원) 추출 대상에서 제외.
        # 퇴직일자가 NULL 또는 등록되지 않은 경우 (현재 근무 중인 인원) 추출.
        sql = 'select [c001], [c002], [c004], [c008], [c010], [c013] from [hpublic00].[dbo].[T0201] ' +\
                f"where [C013] in('', 'None', NULL) or not (select SUBSTRING([c013],1,4)) < '{select_year}';" 
        DataSet = self.sql_query(sql, 2)

        if DataSet == -1:   # 서버 IP와의 연결을 하지 않고 검색 진행
            self.text3.configure(state="normal")
            self.text3.bind("<Key>", lambda a: "break")
            self.text3.delete(1.0, "end")
            self.text3.insert(1.0, "서버 연결을 먼저 진행해주세요.")
            self.text3.configure(state="disabled")

            self.text2.configure(state="normal")
            self.text2.bind("<Key>", lambda a: "break")
            self.text2.delete(1.0, "end")
            self.text2.insert(1.0, "서버 연결 필요!!")
            self.text2.configure(state="disabled")  
        else:
            for Data in DataSet:
                #print(f"테이블 정보 = {Data[0]},{Data[1]},{Data[2]},{Data[3]},{Data[4]}, {Data[5]}")
                # Data[0]:사번, Data[1]:이름, Data[2]:부서코드, Data[3]:직급코드, Data[4]:입사일자, Data[5]:퇴직일자
                if Data[5] == None:
                    return_table.append(self.Search_User(Data))
                else:
                    end_work = Data[5].strip()   # end_work = 퇴직일자
                    if end_work == "" or end_work == None:
                        return_table.append(self.Search_User(Data))
                    else:
                        if int(end_work.split(".")[1]) < int(select_month)-1:
                            continue
                        else:
                            return_table.append(self.Search_User(Data))

            # 테이블 데이터로 구현
            for row in return_table:
                if row == None:
                    continue
                else:
                    self.treeview.insert('', 'end', values=row, tags="select_tag")
            
            self.text3.configure(state="normal")
            self.text3.insert("end", "\n 테이블 조회가 완료되었습니다!!")
            self.text3.configure(state="disabled")




    # ---------------------------------------
    #      조회(돋보기아이콘) 클릭 이벤트
    # ---------------------------------------
    def Search_User(self, Data):
        # 사원 코드 획득
        if Data[0] == None:  # 사원코드
            return None
        else:
            user_num = Data[0].strip()  


        # 사원 이름과 이메일 획득
        if Data[1] == None: # 사원 이름
            return None
        else:
            name = Data[1].strip()
            sql = f"select [c002] from [hpublic00].[dbo].[T0201B] where [c001]='{user_num}';"
            user_email = self.sql_query(sql, 1) # 사원 이메일
            if user_email == None: # 이메일 정보가 없다면
                return None
            else:
                user_email = user_email[0].strip()

        # 부서명 획득
        if Data[2] == None:
            part = ""
        else:
            part = Data[2].strip()        
            sql = f"select [c002] from [hpublic00].[dbo].[T0102] where [c001]='{part}';"
            try:
                user_part = self.sql_query(sql, 1)
                user_part = user_part[0].strip()
            except:
                user_part = ""

        # 직급 획득
        if Data[3] == None:
            user_position = ""
        else:
            position = Data[3].strip()
            sql = f"select [c002] from [hpublic00].[dbo].[T0107] where [c001]='{position}';"
            try:
                user_position = self.sql_query(sql, 1)
                user_position = user_position[0].strip()
            except:
                user_position = ""

        # 이메일 정보가 없는 인원은 테이블에서 제외
        print(user_num, name, user_part, user_position, user_email)
        if user_email == "":
            return None
        else:
            return ["", user_num, name, user_part, user_position, user_email]



    # ----------------------------------------------------------
    #   테이블에서 선택된 인원 리스트 생성 (메일 전송에 사용)
    # ----------------------------------------------------------
    def Get_Selection(self):
        selection_list = []
        select_rows = self.treeview.get_children()
        #print(select_rows)

        for value in select_rows:
            if self.treeview.item(value).get('values')[0] == "":
                pass
            else:
                selection_list.append(value)
        
        return selection_list



    # 여긴 금여 명세 Excel 자동화 해보겠다고 끄적거린 흔적입니다...
    # 지우긴 아쉬워서 이곳에 봉안합니다...
    '''
    def Make_Excel(self):
        self.Get_Excel_Values(self.Get_Selection())


    def Get_Excel_Values(self, selection_list):
        # 필요한 항목
        """
        성명, 사번, 부서명, 직책, 은행, 계좌번호
        기본급, 직무급, 능률급, 직책수당, 자격수당, 환경수당,
        안전수당, 가족수당, 식대수당, 연장수당, 근태공제, 반납공제,
        기타수당, 소급, 연차수당, 증감사항, 상여수당, 돌발수당
        소득세, 주민세, 국민연금, 건강보험, 고용보험, 상조회비,
        기타공제, 정산소득세, 정산주민세, 건강보험정산, 식권공제, 체권압류
        지급합계액, 공제합계액, 치안지급액
        평일잔업시간, 안전준비시간, 심야근무시간, 조퇴시간, 연차일수, 무급일수
        """

        Excel_Data = []
        for selection in selection_list:
            values = self.treeview.item(selection).get('values')
            Excel_Data.append(values[1]) # 사번
            Excel_Data.append(values[2]) # 사원명
            Excel_Data.append(values[3]) # 부서명
            
            # 직책 정보
            n = str(values[1])
            sql = f"select [c002] from [hpublic00].[dbo].[T0103] where [hpublic00].[dbo].[T0103].[c001] IN " + \
                f"(select [c005] from [hpublic00].[dbo].[T0201] where [hpublic00].[dbo].[T0201].[c001] = '{n}');"

            print(sql)
            Data = self.sql_query(sql, 1)
            Data = Data[0].strip()
            Excel_Data.append(Data) # 직책


            # 은행 정보
            sql = f"select [c002] from [hsalary].[dbo].[T0110] where [hsalary].[dbo].[T0110].[c001] IN " + \
                f"(select [c035] from [hpublic00].[dbo].[T0201] where [hpublic00].[dbo].[T0201].[c001] = '{n}');"
            Data = self.sql_query(sql, 1)
            Data = Data[0].strip()
            Excel_Data.append(Data) # 은행

            # 계좌, 기본급, 정보
            sql = f"select [c031], [c024] from [hpublic00].[dbo].[T0201] where [hpublic00].[dbo].[T0201].[c001] = '{n}';"
            Data = self.sql_query(sql, 1)
            Data = Data[0].strip()
            Excel_Data.append(Data) # 계좌 정보

        print(Excel_Data)
    '''


# 실제 EXE 프로그램을 실행하면 수행되는 부분.
# Window 객체는 MailProgram 본체가 됩니다.
if __name__ == "__main__":
    window = MailProgram(None)
    window.title("e-mail 전송 시스템")
    window.geometry("1080x700")    # 창 크기 명시
    window.resizable(False, True)  # 창 조절(가로, 세로)
    window.mainloop()
