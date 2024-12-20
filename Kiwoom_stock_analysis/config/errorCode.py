def errors(err_code):
    err_dict = {0:('OP_ERR_NONE', '정상처리'),
                -10: ('OPP_ERR_FAIL','실패'),
               -100: ('OPP_ERR_LOGIN','사용자정보교환실패'),
               -101: ('OPP_ERR_CONNECT','서버접속실패'),
               -102: ('OPP_ERR_VERSION','버전처리실패'),
               -103: ('OPP_ERR_FIREWALL','개인방화벽실패'),
               -104: ('OPP_ERR_MEMORY','메모리보호실패'),
               -105: ('OPP_ERR_INPUT','함수입력값오류'),
               -106: ('OPP_ERR_SOCKET_CLOSED','통신연결종료'),
               -200: ('OPP_ERR_SISE_OVERFLOW','시세조회과부하'),
               -201: ('OPP_ERR_RQ_STRUCT_FAIL','전문작성초기화실패'),
               -202: ('OPP_ERR_RQ_STRING_FAIL','전문작성입력값오류'),
               -203: ('OPP_ERR_NO_DATA','데이터없음'),
               -204: ('OPP_ERR_OVER_MAX_DATA','조회가능한종목수초과'),
               -205: ('OPP_ERR_DATA_RCV_FAIL','데이터수신실패'),
               -206: ('OPP_ERR_OVER_MAX_FID','조회가능한FID수초과'),
               -207: ('OPP_ERR_REAL_CANCEL','실시간해제오류'),
               -300: ('OPP_ERR_ORD_WRONG_INPUT','입력값오류'),
               -301: ('OPP_ERR_ORD_WRONG_ACCTNO','계좌비밀번호없음'),
               -302: ('OPP_ERR_OTHER_ACC_USE','타인계좌사용오류'),
               -303: ('OPP_ERR_MIS_2BILL_EXC','주문가격이20억원을초과'),
               -304: ('OPP_ERR_MIS_5BILL_EXC','주문가격이50억원을초과'),
               -305: ('OPP_ERR_MIS_1PER_EXC','주문수량이총발행주수의1%초과오류'),
               -306: ('OPP_ERR_MIS_3PER_EXC','주문수량은총발행주수의3%초과오류'),
               -307: ('OPP_ERR_SEND_FAIL','주문전송실패'),
               -308: ('OPP_ERR_ORD_OVERFLOW','주문전송과부하'),
               -309: ('OPP_ERR_MIS_300CNT_EXC','주문수량300계약초과'),
               -310: ('OPP_ERR_MIS_500CNT_EXC','주문수량500계약초과'),
               -340: ('OPP_ERR_ORD_WRONG_ACCTINFO','계좌정보없음'),
               -500: ('OPP_ERR_SYMCODE_EMPTY','종목코드없음')
               }
    
    result = err_dict[err_code]
    
    return result