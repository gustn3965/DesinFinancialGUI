import win32com.client
import time as Time

g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')

class CpFutureOrder :
    def __init__(self):


        ret = g_objCpTrade .TradeInit(0)
        print(ret)
        # 자신의 계좌
        self.acc = g_objCpTrade .AccountNumber[0]
        # 주식 상품 구분
        self.accFalg = g_objCpTrade .GoodsList(self.acc, 1)
        print(self.acc, self.accFalg[0])
        self.objOrder = win32com.client.Dispatch("CpTrade.CpTd0311")

    def Order(self, price, amount):

        # 1- 매도, 2 - 매수
        self.objOrder.SetInputValue(0, '2')
        self.objOrder.SetInputValue(1, self.acc)
        self.objOrder.SetInputValue(2,self.accFalg[0])
        self.objOrder.SetInputValue(3,"A102280")
        self.objOrder.SetInputValue(4,amount)
        # 주문 가격은 현재가격으로 되야한다.
        self.objOrder.SetInputValue(5,price)
        # self.objOrder.SetInputValue(5,'2')
        # self.objOrder.SetInputValue(6,'1')
        self.objOrder.SetInputValue(7,'0')
        # 01 - 보통, 02 - 임의, 03 - 시장가,
        self.objOrder.SetInputValue(8, '01')

        ret = self.objOrder.BlockRequest()

        if ret == 4:
            remainTime = g_objCpStatus.LimitRequestRemainTime
            print('연속조회 제한 오류, 남은 시간', remainTime)
            self._wait()
            ret = self.objOrder.BlockRequest()


        print("주문수량 ", self.objOrder.GetHeaderValue(4))
        print("주문수량 : ", self.objOrder.GetDataValue(4,0))
        print("주문가격 ",self.objOrder.Getheadervalue(5))
        print("주문번호 ",self.objOrder.GetHeaderValue(8))



    def buyOrder(self, price, amount):
        return self.Order(price, amount)



    def _wait(self):
        time_remained = g_objCpStatus.LimitRequestRemainTime
        cnt_remained = g_objCpStatus.GetLimitRemainCount(1)
        print("남은 제한 횟수 : " + str(cnt_remained))

        if cnt_remained <= 0 :
            while cnt_remained <= 0 :
                Time.sleep(time_remained/1000)
                time_remained = g_objCpStatus.LimitRequestRemainTime
                cnt_remained = g_objCpStatus.GetLimitRemainCount(1)


class StockPricedData:
    def __init__(self):
        self.dicEx = {ord('0'): "동시호가/장중 아님", ord('1'): "동시호가", ord('2'): "장중"}
        self.code = ""
        self.name = ""
        self.cur = 0  # 현재가
        self.open = self.high = self.low = 0  # 시/고/저
        self.diff = 0
        self.diffp = 0
        self.objCur = None
        self.objBid = None
        self.vol = 0  # 거래량
        self.offer = [0 for _ in range(10)]  # 매도호가
        self.bid = [0 for _ in range(10)]  # 매수호가
        self.offervol = [0 for _ in range(10)]  # 매도호가 잔량
        self.bidvol = [0 for _ in range(10)]  # 매수호가 잔량




class CpRPCurrentPrice :
    def __init__(self):
        if (g_objCpStatus.IsConnect == 0) :
            print("정상적으로 연결안됨")
            return
        self.objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        return
    def Request(self,code, rtMst):
        rqtime = Time.time()

        self.objStockMst.SetInputValue(0, code)
        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print("통신상태", self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False
        rtMst.code = code
        rtMst.name = g_objCodeMgr.CodeToName(code)
        rtMst.cur = self.objStockMst.GetHeaderValue(11)  # 종가
        rtMst.diff = self.objStockMst.GetHeaderValue(12)  # 전일대비
        rtMst.baseprice = self.objStockMst.GetHeaderValue(27)  # 기준가
        rtMst.exFlag = self.objStockMst.GetHeaderValue(58)  # 예상플래그
        if rtMst.baseprice:
            rtMst.diffp = (rtMst.diff / rtMst.baseprice) * 100

        # 10차호가
        for i in range(10):
            rtMst.offer[i] = (self.objStockMst.GetDataValue(0, i))  # 매도호가
            rtMst.bid[i] = (self.objStockMst.GetDataValue(1, i))  # 매수호가
            rtMst.offervol[i] = (self.objStockMst.GetDataValue(2, i))  # 매도호가 잔량
            rtMst.bidvol[i] = (self.objStockMst.GetDataValue(3, i))  # 매수호가 잔량
        return True




if __name__ == "__main__" :
    # 현재 계좌번호 출력
    aa = CpFutureOrder()
    print(aa.acc)


     # A102280 주식의 현재가 구함.
    rtMst = StockPricedData()
    current = CpRPCurrentPrice()
    current.Request("A102280",rtMst)

    print(rtMst.cur)

    # 해당주식의 현재가로 1주만큼 삼.
    order = CpFutureOrder()
    order.buyOrder(rtMst.cur, 1)