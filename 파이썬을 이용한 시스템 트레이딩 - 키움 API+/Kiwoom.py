'''
Kiwoom.py
@author : yeojin

- Description : kiwoom class is developed to do system trading with Open API provided by Kiwoom Securities.
- Updated on : 2017.03.22

'''


from PyQt4.QtCore import * # SIGNAL, QEventLoop 포함
from PyQt4.QAxContainer import * # QAxWidget 포함


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self.create_kiwoom_instance()
        self.connect_event_handler()

    ################################################################################################
    ################################## Open API+ 컨트롤 연결 #######################################
    def create_kiwoom_instance(self):
        # Contrl CLSID(Class Identifier): {A1574A0D-6BFA-4BD7-9020-DED88711818D}
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")  # @arg : ProgID (Program ID)

    ################################################################################################
    ################################## 이벤트 핸들러 설정 ##########################################
    # 이벤트 함수, 이벤트 처리(핸들러) 함수 연결
    def connect_event_handler(self):
        self.connect(self, SIGNAL(
            "OnReceiveTrData(QString, QString, QString, QString, QString, int, QString, QString, QString)"),
                     self._OnReceiveTrData)
        self.connect(self, SIGNAL("OnEventConnect(int)"), self._OnEventConnect)
        # self.connect(self, SIGNAL("OnReceiveRealData(QString, QString, QString)"), self.OnReceiveRealData)
        self.connect(self, SIGNAL("OnReceiveConditionVer(int, QString)"), self._OnReceiveConditionVer)
        self.connect(self, SIGNAL("OnReceiveChejanData(QString, int, QString)"), self._OnReceiveChejanData)
        self.connect(self, SIGNAL("OnReceiveMsg(QString, QString, QString, QString)"), self._OnReceiveMsg)
        self.connect(self, SIGNAL("OnReceiveTrCondition(QString, QString, QString, int, int)"),
                     self._OnReceiveTrCondition)

    ################################################################################################# 
    ################################## 로그인 #######################################################
    def login(self):
        """
        [GetConnectState() 함수] : 통신 접속 상태 반환
        - 0 : 미연결, 1 : 연결 완료
        [CommConnect() 함수] : 로그인 윈도우 실행
        - OnEventConnect() 이벤트 발생, 로그인 성공 여부가 이벤트의 인자값으로 들어감
        """
        if self.dynamicCall("GetConnectState()") == 0:
            self.dynamicCall("CommConnect()")  # OnEventConnect() 이벤트 발생 (통신 연결 상태 변경 시 발생)
            self.login_event_loop = QEventLoop()
            self.login_event_loop.exec()
        elif self.dynamicCall("GetConnectState()") == 1:
            print("이미 로그인 되었습니다")

    def _OnEventConnect(self, ErrorCode):
        """ [OnEventConnect() 함수] : 통신 연결 상태 변경 시 이벤트 핸들러 함수
        - CommConnect() 로그인 함수 실행 후 로그인 성공 여부를 알 수 있다
        """
        if ErrorCode == 0:  # 성공
            print("< 로그인 성공 >")
            print("----- 사용자 정보 -----")
            print("아이디 : " + self.dynamicCall('getLoginInfo("USER_ID")'))
            print("이름 : " + self.dynamicCall('getLoginInfo("USER_NAME")'))
            account = self.dynamicCall('getLoginInfo("ACCNO")')
            # print(account) # 0000000000;0000000000;0000000000; 형식

            # 계좌 정보 String 처리
            account = account.split(";")[:]
            account = account[:len(account) - 1]
            # print(str(len(account)))
            print("계좌 번호")
            for item in account:
                print(item)
            print("----------------------")

        else:  # 실패
            self.info.append("< 로그인 실패 >")
            print(ErrorCode)
            # 100 : 사용자 정보교환 실패
            # 101 : 서버접속 실패
            # 102 : 버전처리 실패
        self.login_event_loop.exit()

    ################################################################################################
    ######################################### 기타 함수 ############################################

    def SetInputValue(self, sID, sValue):
        """ [SetInPutValue() 함수] : Tran 입력 값을 서버 통신 전에 입력
        sIdD : 아이템명
        sValue : 입력 값
        ex) kiwoom.SetInputValue("종목코드", "000660");
        """
        self.dynamicCall("SetInputValue(QString, QString)", sID, sValue)

    def InitOHLCVRawData(self): # Open High Low Close Volume
        """
        TRAN OPT10081 (주식일봉차트조회요청) 요청 시 데이터 저장을 위한 변수 생성
        """
        self.ohlc = {'date': [], 'open': [], 'high': [], 'low': [], 'close': [], 'volume' : []}

    def GetRepeatCnt(self, sTrCode, sRecordName):
        """ [GetRepeatCnt() 함수] : 조회수신한 멀티데이터의 갯수를 얻을 수 있다.
        - OnReceiveTrData() 이벤트 함수 내에서 사용
        """
        ret = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRecordName)
        return ret

    @staticmethod
    def change_format(data):
        """ [ChangeFormat() 함수 : 천의 자리마다 콤마 표시]
        """
        strip_data = data.lstrip('-0')  # 처음 -,0 모두 지우기
        if strip_data == '':
            strip_data = '0'

        format_data = format(int(strip_data), ',d')  # 1000의 자리마다 , 찍기
        if data.startswith('-'):  # - 붙이기
            format_data = '-' + format_data

        return format_data

    @staticmethod
    def change_format2(data):
        """ [Change_format2() 함수 : 수익률 변환 함수]
        """ 
        strip_data = data.lstrip('-0')
        if strip_data == '':
            strip_data = '0'

        if strip_data.startswith('.'):
            strip_data = '0' + strip_data

        if data.startswith('-'):
            strip_data = '-' + strip_data

        return strip_data

    ################################################################################################
    ################################## 조회 요청하기 #################################################

    def CommRqData(self, sRQName, sTRCode, nPrevNext, sScreenNo):
        """ [CommRqData() 함수] : 통신 데이터를 서버로 송신하는 함수
        sRQName : 사용자 구분 명
        sTrCode : Tran명 입력
        sPrevNext : 0 조회, 2 연속
        sScreenNo : 4자리의 화면번호(KOAStudio 참조)
        - nPrevNext = 0 : 처음 조회 시 혹은 연속 데이터가 없을 때 / nPrevNext = 2 : 연속 조회 시
        - 조회 요청이 성공하면 관련 실시간 데이터를 서버에서 자동으로 OnReceiveRealData()이벤트 함수로 전달
        """
        self.dynamicCall("CommRqData(QString, QString, int, QString)", sRQName, sTRCode, nPrevNext, sScreenNo)
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec()

    def _OnReceiveRealData(self, code, realType, realData):
        """[OnReceiveRealData() 함수] : 실시간 시세 이벤트 핸들러 함수
        code : 종목코드
        realType : 리얼타입 (KOA Studio의 '실시간목록' 참조)
        realData : 실시간 데이터 전문
        - 실시간 데이터를 받은 시점을 알려준다
        - CommRqData() 함수를 통해 조회 요청이 성공하면 관련 실시간 데이터를 OnReceiveRealData() 이벤트 함수로 전달
        """
        print("OnReceiveRealData() 이벤트 발생")
        print("< 실시간 데이타 >")
        # data = self.GetCommRealData(code, realType, realData)
        # detail = data.detail
        print("종목코드 : " + code)
        print("리얼타입 : " + realType)
        # print(realData)
        print("누적거래량 : " + GetCommRealData(code, 13))
        print("체결강도 : " + GetCommRealData(code, 228))
        print("체결 시간 : " + GetCommRealData(code, 20))

        self.real_event_loop = QEventLoop()

        try:
            self.tr_event_loop.exit()
        except AttributeError:
            pass

    def GetCommRealData(self, strCode, nFid):
        """ [GetCommRealData() 함수]
        - OnReceiveRealData() 이벤트 함수 호출 시 실시간 데이터 얻어오는 함수
        """
        data = self.dynamicCall("GetCommRealData(QString, QString)", strCode, nFid)

        return data.strip()

    def _OnReceiveTrData(self, ScrNo, RQName, TrCode, RecordName, PrevNext, DataLength, ErrorCode, Message, SplmMsg):
        """ [onReceiveTrData() 이벤트 핸들러 함수]
        - 조회 요청 응답을 받거나 조회데이터를 수신했을 때 호출
        - 조회 데이터는 함수 내부에서 GetCommData()함수를 이용해 얻어올 수 있다
        ScrNo : 화면번호
        RQName : 사용자 구분 명
        sTrCode : Tran 명
        RecordName : Record 명
        PreNext : 연속조회 유무
        DataLength, ErrorCode, Message, SplmMsg : 1.0.0.1 버전 이후 사용 안함 
        """

        print("<< OnReceiveTrData 이벤트 발생 >>")
        self.prev_next = PrevNext # 연속조회 유무 인자 값
        data_len = self.GetRepeatCnt(TrCode, RQName)  # 데이터 길이 Get
        print("TR 조회명 : " + RQName) 
        print("레코드 반복 횟수 : " + str(data_len))
        print("")

        if RQName == "주식일봉차트조회요청":
            for i in range(data_len):
                date = self.CommGetData(TrCode, "", RQName, i, "일자")
                open = self.CommGetData(TrCode, "", RQName, i, "시가")
                high = self.CommGetData(TrCode, "", RQName, i, "고가")
                low = self.CommGetData(TrCode, "", RQName, i, "저가")
                close = self.CommGetData(TrCode, "", RQName, i, "현재가")
                volume = self.CommGetData(TrCode, "", RQName, i, "거래량")


                self.ohlc['date'].append(date)
                self.ohlc['open'].append(int(open))
                self.ohlc['high'].append(int(high))
                self.ohlc['low'].append(int(low))
                self.ohlc['close'].append(int(close))
                self.ohlc['volume'].append(int(volume))

        elif RQName == "주식기본정보요청":
            name = self.CommGetData(TrCode, "", RQName, 0, "종목명")
            volume = self.CommGetData(TrCode, "", RQName, 0, "거래량")
            close = self.CommGetData(TrCode, "", RQName, 0, "현재가")
            open = self.CommGetData(TrCode, "", RQName, 0, "시가")
            high = self.CommGetData(TrCode, "", RQName, 0, "고가")
            low = self.CommGetData(TrCode, "", RQName, 0, "저가")
            
            # upperLimit = self.CommGetData(TrCode, "", RQName, 0, "상한가")
            # lowerLimit = self.CommGetData(TrCode, "", RQName, 0, "하한가")
            net_change = self.CommGetData(TrCode, "", RQName, 0, "전일대비")
            fluctuation_rate = self.CommGetData(TrCode, "", RQName, 0, "등락율")

            print("종목명 : " + name)
            print("거래량 : " + volume)
            print("현재가 : " + close)
            print("시가 : " + open)
            print("고가 : " + high)
            print("저가 : " + low)
            # print("상한가 : " + upperLimit)
            # print("하한가 : " + lowerLimit)
            print("전일대비 : " + net_change)
            print("등락률 : " + fluctuation_rate + "%")

        elif RQName == "계좌평가잔고내역요청":
            self.temp = self.CommGetData(TrCode, "", RQName, 0, "총수익률(%)")
            print("총매입 : " + Kiwoom.change_format(self.CommGetData(TrCode, "", RQName, 0, "총매입금액")))
            print("총평가 : " + Kiwoom.change_format(self.CommGetData(TrCode, "", RQName, 0, "총평가금액")))
            print("총평가손익 : " + Kiwoom.change_format(self.CommGetData(TrCode, "", RQName, 0, "총평가손익금액")))
            print("총수익률 (%) : " + Kiwoom.change_format2(self.CommGetData(TrCode, "", RQName, 0, "총수익률(%)")))
            print("추정자산 : " + Kiwoom.change_format(self.CommGetData(TrCode, "", RQName, 0, "추정예탁자산")))

            # 보유 종목별 잔고 데이터 출력
            for i in range(0, data_len):
                print("<보유 종목>")
                print("(" + str(i) + ")")
                print("종목명 : " + self.CommGetData(TrCode, "", RQName, i, "종목명"))
                print("9 : " + Kiwoom.change_format(self.CommGetData(TrCode, "", RQName, i, "보유수량")))
                print("매입가 : " + Kiwoom.change_format(self.CommGetData(TrCode, "", RQName, i, "매입가")))
                print("총매입금액 : " + Kiwoom.change_format(self.CommGetData(TrCode, "", RQName, i, "매입금액")))
                print("현재가 : " + Kiwoom.change_format(self.CommGetData(TrCode, "", RQName, i, "현재가")))
                print("평가손익 : " + Kiwoom.change_format(self.CommGetData(TrCode, "", RQName, i, "평가손익")))
                print("수익률 (%) : " + Kiwoom.change_format2(self.CommGetData(TrCode, "", RQName, i, "수익률(%)")))

        elif RQName == "예수금상세현황":
            print("d+2추정예수금 : " + Kiwoom.change_format(self.CommGetData(TrCode, "", "RQName", 0, "d+2추정예수금")))
            print("주식증거금현금 : " + Kiwoom.change_format(self.CommGetData(TrCode, "", "RQName", 0, "주식증거금현금")))
            print("주문가능금액 : " + Kiwoom.change_format(self.CommGetData(TrCode, "", "RQName", 0, "주문가능금액")))

        try:
            self.tr_event_loop.exit()
        except AttributeError:
            pass

    def CommGetData(self, sJongmokCode, sRealType, sFieldName, nIndex, sInnerFiledName):
        """ [CommGetData() 함수]
        - OnReceiveTrData() 함수가 호출될 때 조회된 데이터를 얻어오는 함수
        - 반드시 OnReceiveTrData() 함수가 호출될 때 그 안에서 사용 (이벤트 핸들러 함수 사용)
        1) Tran 데이터
        sJongmokCode : Tran명, sRealType : 사용 안함, sFieldName : 레코드명, nIndex : 반복인덱스, sInnerFieldName : 아이템명
        ex) kiwoom.CommGetData(“OPT00001”, “”, “주식기본정보”, 0, “현재가”);
        2) 실시간 데이터
        sJongmokCode : key Code, sRealType : Real Type, sFieldName : Item Index, nIndex : 사용 안함, sInnerFieldName : 사용 안함
        ex) kiwoom.CommGetData(“000660”, “A”, 0);
        3) 체결 데이터
        sJongmokCode : 체결 구분, sRealType : "-1", sFieldName : 사용 안함, nIndex : ItemIndex, sInnerFieldName : 사용 안함
        ex) kiwoom.CommGetData(“000660”, “-1”, 1);
        """
        data = self.dynamicCall("CommGetData(QString, QString, QString, int, QString)", sJongmokCode, sRealType,
                                sFieldName, nIndex, sInnerFiledName)

        return data.strip()

    ################################################################################################
    ################################## 종목 정보 관련 함수 ############################################

    def GetCodeListByMarket(self, market):
        """ [GetCodeListByMarket()함수] : 시장 구분에 따른 종목 코드 반환
        sMarket - 0 : 장내, 8 : ETF, 10 : 코스닥 등
        """
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market)
        code_list = code_list.split(';')
        return code_list[:-1]

    def GetMasterCodeName(self, strCode):
        """ [GetMasterCodeName()함수] : 종목코드에 해당하는 한글명을 반환한다
        strCode : 종목코드
        """
        cmd = 'GetMasterCodeName("%s")' % strCode
        ret = self.dynamicCall(cmd)
        return ret

    ################################################################################################
    ################################## 주문과 잔고 처리 ############################################

    def SendOrder(self, rQName, screenNo, accNo, orderType, code, qty, price, hogaGb, orgOrderNo):

        """ [SendOrder() 함수]
         BSTR sRQName, // 사용자 구분명
         BSTR sScreenNo, // 화면번호
         BSTR sAccNo,  // 계좌번호 10자리
         LONG nOrderType,  // 주문유형 1:신규매수, 2:신규매도 3:매수취소, 4:매도취소, 5:매수정정, 6:매도정정
         BSTR sCode, // 종목코드
         LONG nQty,  // 주문수량
         LONG nPrice, // 주문가격
         BSTR sHogaGb,   // 거래구분(혹은 호가구분) (00 : 지정가)
         BSTR sOrgOrderNo  // 원주문번호입니다. 신규주문에는 공백, 정정(취소)주문할 원주문번호를 입력합니다.
        """
        self.order_event_loop = QEventLoop()

        # return 값 0 이면 성공
        ret = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                               [rQName, screenNo, accNo, orderType, code, qty, price, hogaGb, orgOrderNo])
        self.order_event_loop.exec()
        return ret

    def GetChejanData(self, nFid):
        """ [GetChejanData() 함수] : 체결잔고 데이터를 반환한다
        nFid : 체결잔고 아이템 (키움 API+ 개발 가이드 참조)
        """
        cmd = 'GetChejanData("%s")' % nFid
        ret = self.dynamicCall(cmd)
        return ret

    def _OnReceiveChejanData(self, sGubun, nItemCnt, sFidList):
        """ [OnReceiveChejanData() 이벤트 핸들러 함수] : 체결데이터를 받은 시점을 알려준다
        sGubun : 체결구분 (0 : 주문체결통보, 1 : 잔고통보, 3: 특이신호)
        nItemCnt : 아이템갯수
        sFidList : 데이터리스트 (데이터 구분 ;)
        """
        print("<< OnReceiveChejanData() 이벤트 발생 >> ")
        if sGubun == "0" :
            print("체결 구분 : ", sGubun, "(0 : 주문체결통보, 1: 잔고통보)")
            time = self.GetChejanData(908)
            print("주문/체결 시간 : ", time[0:2], "시", time[2:4], "분", time[4:6], "초")  # 주문/체결시간
            print("매도/수 구분 : ", self.GetChejanData(907), "(1 : 매도, 2  : 매수)")
            print("주문번호 : ", self.GetChejanData(9203))
            print("종목명 : ", self.GetChejanData(302))
            print("주문수량 : ", self.GetChejanData(900))
            print("주문가격 : ", self.GetChejanData(901))
            print("")
        elif sGubun == "1": # 잔고 정보
            print("체결 구분 : ", sGubun, "(0 : 주문체결통보, 1: 잔고통보)")
            pass
        
        # print(self.order_event_loop.isRunning())
        
        try:
            self.order_event_loop.exit()
        except AttributeError:
            pass

    def _OnReceiveMsg(self, sScrNo, QRName, TrCode, Msg):
        print("<< OnReceiveMsg() 이벤트 발생 >>")
        print(QRName + Msg)
        print("")

    ################################################################################################
    ###################################### 조건 검색 #################################################

    def GetConditionLoad(self):
        """ [GetConditionLoad() 함수]
        - 영웅문 HTS에서 작성한 사용자 조건검색 목록을 서버에 요청
        - 성공 시 1, 아니면 0 리턴
        """
        ret = self.dynamicCall("GetConditionLoad()")  # 성공 : 1, 실패 : 0 return
        self.condition_event_loop = QEventLoop()
        self.condition_event_loop.exec()
        # return ret

    def _OnReceiveConditionVer(self, ret, msg):
        """ [OnReceiveConditionVer() 이벤트 핸들러 함수]
        - 사용자 조건검색 목록 요청에 대한 응답을 서버에서 수신하면 호출되는 이벤트 함수
        """
        print("<< OnReceiveConditionVer() 이벤트 발생 >>")
        if (ret != 0):  # 조건검색식 요청 성공
            print("조건검색식 목록")
            condition = self.GetConditionNameList()
            print(condition)
            # 임시로 로컬 파일에 조건검색식 저장
            text_file = open("condition.txt", "w")
            text_file.write(condition)
            text_file.close()
        else:
            print("No Data")
        self.condition_event_loop.exit()

    def GetConditionNameList(self):
        """ [GetConditionNameList() 함수]
        - 서버에서 수신한 사용자 조건식을 조건명 인덱스와 조건식 이름을 한쌍으로 하는 문자열들로 전달됨
        """
        ret = self.dynamicCall("GetConditionNameList()")
        return ret

    def SendCondition(self, scrNo, conditionName, index, search):
        """ [SendCondition() 함수]
        scrNo : 화면번호
        conditionName : 조건식 이름
        index : 조건명 인덱스
        search : 조회 구분 (0 : 조건 검색, 1 :실시간 조건 검색)
        - 서버에 조건 검색 요청
        ex) GetConditionNameList()함수로 얻은 조건식 목록이
        "0^조건식1;3^조건식1;8^조건식3;23^조건식5"일때 조건식3을 검색
        -----> long lRet = SendCondition("0156", "조건식3", 8, 1);
        """
        ret = self.dynamicCall("SendCondition(QString, QString, int, int)", scrNo, conditionName, index, search)
        self.condition_event_loop = QEventLoop()
        self.condition_event_loop.exec()
        return ret

    def _OnReceiveTrCondition(self, sScrNo, strCodeList, strConditionName, nIndex, nNext):
        """ [OnReceiveTrCondition() 이벤트 핸들러 함수]
        - 조건검색 요청으로 검색된 종목 코드 리스트를 전달하는 이벤트 함수
        """
        print("<< OnReceiveTrCondition() 이벤트 발생 >>")
        print("< 사용자 조건 검색 종목 코드 >")
        if (len(strCodeList) != 0):
            print(strCodeList)
        else:
            print("No Data")

        try:
            self.condition_event_loop.exit()
        except AttributeError:
            pass