import win32com.client
 
# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()
 
# 현재가 객체 구하기
objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
objStockMst.SetInputValue(0, 'A005930')   #종목 코드 - 삼성전자
objStockMst.BlockRequest()
 
# 현재가 통신 및 통신 에러 처리 
rqStatus = objStockMst.GetDibStatus()
rqRet = objStockMst.GetDibMsg1()
print("통신상태", rqStatus, rqRet)
if rqStatus != 0:
    exit()
 
# 현재가 정보 조회 주식을 바로 매수할 수 있는 시장 가격인 매도호가만 남기고 나머지는 다 지운다. 
offer = objStockMst.GetHeaderValue(16)  #매도호가

import requests

class MyMsg():
    def send_msg(self, msg=""):
        response = requests.post(
            'https://slack.com/api/chat.postMessage',
            headers={
                'Authorization': 'Bearer '+"slack_code"
            },
            data={
                'channel':'#stock',
                'text':msg
            }
        )
        print(response)

MyMsg().send_msg(msg="삼성전자 현재가 :"+ str(offer) )

