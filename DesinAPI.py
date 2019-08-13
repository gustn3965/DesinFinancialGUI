import win32com.client
import time as Time
import pandas as pd
import re
from PyQt5.QtWidgets import *


#  시장구분.
# 종목코드 총 개수 : 3163

#  KOSPI = 거래소. 1번. 1527개
#  코스닥 = 코스닥  2번    1349개
#  K-OTC 중소/중견  3번   137개
#  KONEX           5번   150개



class DesinAPI :
    a = 1

    # Check conenction
    def __init__(self):
        self.result = ""
        self.instCpStockCode = win32com.client.Dispatch('CpUtil.CpCybos')
        if self.instCpStockCode.IsConnect == 1 :
            print("연결되었습니다.")
            self.result = "연결되었습니다."
        else :
            print("연결 되지 않았습니다. " )
            self.result = "연결되지 않았습니다."



    # Get StockCode about enterprise I want
    def GetStockCode(self,name) :
        instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
        maxCodeNum = instCpStockCode.GetCount()
        print("종목코드 총 개수 : ", maxCodeNum)
        for i in range(0,maxCodeNum) :
            # 0 코드명 , 1 종목명, 2 FullCode
            # print(instCpStockCode.GetData(1,i))
            if instCpStockCode.GetData(1, i ) == name :
                nameCode = instCpStockCode.GetData(0, i)
                print(name + " 의 코드명 : " + nameCode)
                print(name + " 의 인덱스번호 : " + str(i) )
                return nameCode



    def SearchNameListByName(self, name ):
        nameList = []
        codeList = []

        instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
        maxCodeNum = instCpStockCode.GetCount()
        for i in range(0,maxCodeNum) :
            nameList.append(instCpStockCode.GetData(1,i))

        name = name.upper()
        regex = re.compile(name)
        matches = [string for string in nameList if re.match(regex, string)]
        print(matches)

        dataDict = { }
        nameList = []

        for i in matches :
            nameList.append(i)
            codeList.append(instCpStockCode.NameToCode(i))

        self.dataDict = {'name' : nameList , 'code' : codeList}


    def SearchNameListByCode(self, code ):
        nameList = []
        codeList = []

        instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
        maxCodeNum = instCpStockCode.GetCount()

        for i in range(0,maxCodeNum) :
            codeList.append(instCpStockCode.GetData(0,i))



        name = code.upper()
        regex = re.compile(name)


        matches = [string for string in codeList if re.match(regex, re.split("\D",string)[1])]
        print(matches)

        dataDict = { }
        codeList = []

        for i in matches :
            codeList.append(i)
            nameList.append(instCpStockCode.CodeToName(i))

        self.dataDict = {'name' : nameList , 'code' : codeList}








    #  시장구분.
    #  KOSPI = 거래소. 1번.    1527 개
    #  코스닥 = 코스닥  2번     1349 개
    #  K-OTC 중소/중견  3번     137 개
    #  KONEX           5번     150 개



    # 가장 첫번째 화면에서 데이터불러오기.
    # codeName - 종목코드
    # count  - 갯수
    # progressBar -
    # mT - [ m = 분봉 , T - 틱봉 ]
    def GetRecentData(self,codeName, count,tick_range,  progressBar, mT):

        columns = ['날짜', '시간', '시가', '고가', '저가', '종가', '거래량']
        self.dict = {'날짜':[], '시간':[],'시가':[],'고가':[],'저가':[],'종가':[],'거래량':[]}


# CpSysDib.StockChart를 사용

        instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")


# SetInputValue
# 0 - 종목코드
# 1 - 요청구분 ( '1' 기간, '2' 개수 )
# 2 - 요청종료일 ( 기간으로 했을 경우 )
# 3 - 요청싲가일 ( 기간으로 했을 경우 )
# 4 - 요청개수
# 5 - 필드 배열 ( 가져올 수 있는 항목 )
# 6 - 차트 구분 ( 'D' 일, 'W' 주 , 'M' 월, 'm' 분 , 'T', 틱)
# 7 - 주기 ( default 1 )
# 9 - 수정주가 ( '0' 무수정주가 , '1' 수정주가 )



        instStockChart.SetInputValue(0, codeName)
        instStockChart.SetInputValue(1, ord('2'))
        instStockChart.SetInputValue(4,count)
        instStockChart.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])
        instStockChart.SetInputValue(6, ord(mT))
        instStockChart.SetInputValue(7, tick_range)
        instStockChart.SetInputValue(9, ord('1'))

#


        # 요청하는 값에 비해, 한번에 받을 수 있는 개수는 6665개.
        # 그러므로, 요청을 반복해야한다. + time.sleep을 걸어준다.
        # 요청값이 총 받은 개수보다 크면, ( 같아도 False) stop.
        # 참고 https://github.com/gyusu/Creon-Datareader/blob/master/creonAPI.py
        rcv_count = 0

        progressBar.setMinimum(rcv_count)
        progressBar.setMaximum(count)

        # KospiList = self.GetMarketCode(1)
        # KosdaqList = self.GetMarketCode(2)


        duplicatedCount = 0

        while count > rcv_count :
            progressBar.setValue(rcv_count)

            instStockChart.BlockRequest()
            Time.sleep(0.25)
            self.numData = instStockChart.GetHeaderValue(3)
            self.numData = min(self.numData, count - rcv_count)


            print("받은 데이타 : ",self.numData)

            if self.numData != 2856 :
                duplicatedCount += 1
            if duplicatedCount > 1 :
                break

            numField = instStockChart.GetHeaderValue(1)

            for i in range(0, self.numData):
                for j in range(0, numField):
                    self.dict[columns[j]].append(instStockChart.GetDataValue(j,i))
            rcv_count += self.numData


            # 2년치의 데이터가 넘어가면, 최신데이터도 가져오는데,
            # 중복의 최신데이터는 필요없기 때문에,
            # 만약 최신데이터가 나타나면 자르고, break한다.

            if self.numData == 0 :
                break
        progressBar.setValue(count)

        self.df = pd.DataFrame(self.dict).sort_index(ascending=False).reset_index(drop=True)
        # print(self.df)


        self._wait()



####   기존 데이터와 엎어치기 위한 메소드  ( 일봉일때, )
    ## 날짜별로
    ## 일별데이터를 얻는다.
    ##  노터치
    def GetUpdatePeriodDay(self, codeName, today,recendDay ):

        columns = ['날짜', '시간', '시가', '고가', '저가', '종가', '전일대비', '거래량', '거래대금', '누적체결매도수량', '누적체결매수수량', '상장주식수', '시가총액',
                   '외국인주문한도수량', '외국인주문가능수량', '외국인현보유수량', '외국인현보유비율', '수정주가일자', '수정주가비율', '기관순매수', '기관누적순매수']

        self.dict = {'날짜': [], '시간': [], '시가': [], '고가': [], '저가': [], '종가': [], '전일대비': [], '거래량': [], '거래대금': [],
                     '누적체결매도수량': [], '누적체결매수수량': [], '상장주식수': [], '시가총액': [], '외국인주문한도수량': [], '외국인주문가능수량': [],
                     '외국인현보유수량': [], '외국인현보유비율': [], '수정주가일자': [], '수정주가비율': [], '기관순매수': [], '기관누적순매수': []}


        # 4 - 개수, 5 - 데이터타입, 6 - 일,주,월,분, 7 - 분봉 주기 ,  9 - 1수정주가
        instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        instStockChart.SetInputValue(0, codeName)
        instStockChart.SetInputValue(1, ord('1'))
        instStockChart.SetInputValue(2, today)
        instStockChart.SetInputValue(3, recendDay)
        # instStockChart.SetInputValue(4, count)
        instStockChart.SetInputValue(5,
                                     [0, 1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21])

        instStockChart.SetInputValue(6, ord('D'))
        instStockChart.SetInputValue(7, 2)
        instStockChart.SetInputValue(9, ord('1'))
        instStockChart.SetInputValue(10, ord('3'))

        # 요청하는 값에 비해, 한번에 받을 수 있는 개수는 6665개.
        # 그러므로, 요청을 반복해야한다. + time.sleep을 걸어준다.
        # 요청값이 총 받은 개수보다 크면, ( 같아도 False) stop.
        # 참고 https://github.com/gyusu/Creon-Datareader/blob/master/creonAPI.py
        rcv_count = 0



        duplicatedCount = 0

        while True :
            # progressBar.setValue(rcv_count)
            #
            instStockChart.BlockRequest()
            # Time.sleep(0.25)
            self.numData = instStockChart.GetHeaderValue(3)
            # self.numData = min(self.numData, count - rcv_count)

            print("받은 데이타 : ", self.numData)

            if self.numData != 1999:
                duplicatedCount += 1
            if duplicatedCount > 1:
                break

            numField = instStockChart.GetHeaderValue(1)

            for i in range(0, self.numData):
                for j in range(0, numField):
                    self.dict[columns[j]].append(instStockChart.GetDataValue(j, i))
            rcv_count += self.numData

            # 2년치의 데이터가 넘어가면, 최신데이터도 가져오는데,
            # 중복의 최신데이터는 필요없기 때문에,
            # 만약 최신데이터가 나타나면 자르고, break한다.

            if self.numData == 0:
                break
            self._wait()


        self.df = pd.DataFrame(self.dict).sort_index(ascending=False).reset_index(drop=True)
        print(self.df)
        # self.df.to_csv(codeName + "미포함.csv", mode='a', index=False,
        #                encoding="euc-kr")
        # print(self.df)

        self._wait()






####   기존 데이터와 엎어치기 위한 메소드  ( 분봉/틱봉일때, )
## 날짜별로 일별데이터를 얻는다.
##  노터치

    def GetUpdatePeriodMinutes(self, codeName, today,recendDay ,mT):

        columns = ['날짜', '시간', '시가', '고가', '저가', '종가', '거래량', '거래대금', '누적체결매도수량', '누적체결매수수량']
        self.dict = {'날짜': [], '시간': [], '시가': [], '고가': [], '저가': [], '종가': [], '거래량': [], '거래대금': [],
                     '누적체결매도수량': [], '누적체결매수수량': []}

        # 4 - 개수, 5 - 데이터타입, 6 - 일,주,월,분, 7 - 분봉 주기 ,  9 - 1수정주가
        instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        instStockChart.SetInputValue(0, codeName)
        instStockChart.SetInputValue(1, ord('1'))
        instStockChart.SetInputValue(2, today)
        instStockChart.SetInputValue(3, recendDay)
        # instStockChart.SetInputValue(4, count)
        instStockChart.SetInputValue(5,
                                     [0, 1, 2, 3, 4, 5, 8, 9, 10, 11])

        instStockChart.SetInputValue(6, ord(mT))
        instStockChart.SetInputValue(7, 1)
        instStockChart.SetInputValue(9, ord('1'))
        instStockChart.SetInputValue(10, ord('3'))

        # 요청하는 값에 비해, 한번에 받을 수 있는 개수는 6665개.
        # 그러므로, 요청을 반복해야한다. + time.sleep을 걸어준다.
        # 요청값이 총 받은 개수보다 크면, ( 같아도 False) stop.
        # 참고 https://github.com/gyusu/Creon-Datareader/blob/master/creonAPI.py
        rcv_count = 0



        duplicatedCount = 0

        while True :
            # progressBar.setValue(rcv_count)
            #
            instStockChart.BlockRequest()
            # Time.sleep(0.25)
            self.numData = instStockChart.GetHeaderValue(3)
            # self.numData = min(self.numData, count - rcv_count)

            print("받은 데이타 : ", self.numData)

            if self.numData != 1999:
                duplicatedCount += 1
            if duplicatedCount > 1:
                break

            numField = instStockChart.GetHeaderValue(1)

            for i in range(0, self.numData):
                for j in range(0, numField):
                    self.dict[columns[j]].append(instStockChart.GetDataValue(j, i))
            rcv_count += self.numData

            # 2년치의 데이터가 넘어가면, 최신데이터도 가져오는데,
            # 중복의 최신데이터는 필요없기 때문에,
            # 만약 최신데이터가 나타나면 자르고, break한다.

            if self.numData == 0:
                break
            self._wait()


        self.df = pd.DataFrame(self.dict).sort_index(ascending=False).reset_index(drop=True)
        print(self.df)
        # self.df.to_csv(codeName + "미포함.csv", mode='a', index=False,
        #                encoding="euc-kr")
        # print(self.df)

        self._wait()





### 일봉 데이터얻기 , 개수로 얻기.
# 노터치.



    def GetDayData(self, codeName, count, tick_range, progressBar):

        columns = ['날짜', '시간', '시가', '고가', '저가', '종가','전일대비', '거래량', '거래대금', '누적체결매도수량', '누적체결매수수량', '상장주식수', '시가총액',
                   '외국인주문한도수량', '외국인주문가능수량', '외국인현보유수량', '외국인현보유비율', '수정주가일자', '수정주가비율', '기관순매수', '기관누적순매수']

        self.dict = {'날짜': [], '시간': [], '시가': [], '고가': [], '저가': [], '종가': [],'전일대비':[], '거래량': [], '거래대금': [],
                     '누적체결매도수량': [], '누적체결매수수량': [], '상장주식수': [], '시가총액': [], '외국인주문한도수량': [], '외국인주문가능수량': [],
                     '외국인현보유수량': [], '외국인현보유비율': [], '수정주가일자': [], '수정주가비율': [], '기관순매수': [], '기관누적순매수': []}

        # 4 - 개수, 5 - 데이터타입, 6 - 일,주,월,분, 7 - 분봉 주기 ,  9 - 1수정주가
        instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        instStockChart.SetInputValue(0, codeName)
        instStockChart.SetInputValue(1, ord('2'))

        instStockChart.SetInputValue(4, count)
        instStockChart.SetInputValue(5,
                                     [0, 1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21])

        instStockChart.SetInputValue(6, ord("D"))
        instStockChart.SetInputValue(7, tick_range)
        instStockChart.SetInputValue(9, ord('0'))
        # 1 시간외모두포함
        # 2 장종료시간외 거래만 포함
        # 3 시간외거래량 모두 제외
        # 4 장전시간외 거래만 포함
        instStockChart.SetInputValue(10,ord('1'))

        # 요청하는 값에 비해, 한번에 받을 수 있는 개수는 6665개.
        # 그러므로, 요청을 반복해야한다. + time.sleep을 걸어준다.
        # 요청값이 총 받은 개수보다 크면, ( 같아도 False) stop.
        # 참고 https://github.com/gyusu/Creon-Datareader/blob/master/creonAPI.py
        rcv_count = 0

        progressBar.setMinimum(rcv_count)
        progressBar.setMaximum(count)

        duplicatedCount = 0

        while count > rcv_count:
            progressBar.setValue(rcv_count)

            instStockChart.BlockRequest()
            Time.sleep(0.25)
            self.numData = instStockChart.GetHeaderValue(3)
            self.numData = min(self.numData, count - rcv_count)

            print("받은 데이타 : ", self.numData)

            if self.numData != 951:
                duplicatedCount += 1
            if duplicatedCount > 1:
                break

            numField = instStockChart.GetHeaderValue(1)

            for i in range(0, self.numData):
                for j in range(0, numField):
                    self.dict[columns[j]].append(instStockChart.GetDataValue(j, i))
            rcv_count += self.numData

            # 2년치의 데이터가 넘어가면, 최신데이터도 가져오는데,
            # 중복의 최신데이터는 필요없기 때문에,
            # 만약 최신데이터가 나타나면 자르고, break한다.

            if self.numData == 0:
                break
            self._wait()
        progressBar.setValue(count)

        self.df = pd.DataFrame(self.dict).sort_index(ascending=False)

        # print(self.df.sort_index(ascending=False))

        # self.df.to_csv(codeName + "장전후미포함.csv", mode='w', index=False,
        #                encoding="euc-kr")
        # print(self.df)

        self._wait()




########   분봉데이터 얻기. or 틱데이터 얻기.
##            노터치.

    def GetMinuteOrTickData(self, codeName, count, tick_range, progressBar, mT):

        columns = ['날짜', '시간', '시가', '고가', '저가', '종가', '거래량', '거래대금', '누적체결매도수량', '누적체결매수수량']
        self.dict = {'날짜': [], '시간': [], '시가': [], '고가': [], '저가': [], '종가': [], '거래량': [], '거래대금': [],
                     '누적체결매도수량': [], '누적체결매수수량': []}

        # 4 - 개수, 5 - 데이터타입, 6 - 일,주,월,분, 7 - 분봉 주기 ,  9 - 1수정주가
        instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        instStockChart.SetInputValue(0, codeName)
        instStockChart.SetInputValue(1, ord('2'))

        instStockChart.SetInputValue(4, count)
        instStockChart.SetInputValue(5,
                                     [0, 1, 2, 3, 4, 5, 8, 9, 10, 11])

        instStockChart.SetInputValue(6, ord(mT))
        instStockChart.SetInputValue(7, tick_range)
        instStockChart.SetInputValue(9, ord('1'))
        instStockChart.SetInputValue(10, ord('3'))

        # 요청하는 값에 비해, 한번에 받을 수 있는 개수는 6665개.
        # 그러므로, 요청을 반복해야한다. + time.sleep을 걸어준다.
        # 요청값이 총 받은 개수보다 크면, ( 같아도 False) stop.
        # 참고 https://github.com/gyusu/Creon-Datareader/blob/master/creonAPI.py
        rcv_count = 0

        progressBar.setMinimum(rcv_count)
        progressBar.setMaximum(count)

        duplicatedCount = 0

        while count > rcv_count:
            progressBar.setValue(rcv_count)

            instStockChart.BlockRequest()
            Time.sleep(0.25)
            self.numData = instStockChart.GetHeaderValue(3)
            self.numData = min(self.numData, count - rcv_count)

            print("받은 데이타 : ", self.numData)

            if self.numData != 1999:
                duplicatedCount += 1
            if duplicatedCount > 1:
                break

            numField = instStockChart.GetHeaderValue(1)

            for i in range(0, self.numData):
                for j in range(0, numField):
                    self.dict[columns[j]].append(instStockChart.GetDataValue(j, i))
            rcv_count += self.numData

            # 2년치의 데이터가 넘어가면, 최신데이터도 가져오는데,
            # 중복의 최신데이터는 필요없기 때문에,
            # 만약 최신데이터가 나타나면 자르고, break한다.

            if self.numData == 0:
                break
            self._wait()
        progressBar.setValue(count)

        self.df = pd.DataFrame(self.dict).sort_index(ascending=False)
        # self.df.to_csv(codeName + "미포함.csv", mode='a', index=False,
        #                encoding="euc-kr")
        # print(self.df)

        self._wait()




    def _wait(self):
        time_remained = self.instCpStockCode.LimitRequestRemainTime
        cnt_remained = self.instCpStockCode.GetLimitRemainCount(1)
        print("남은 제한 횟수 : " + str(cnt_remained))

        if cnt_remained <= 0 :
            while cnt_remained <= 0 :
                Time.sleep(time_remained/1000)
                time_remained = self.instCpStockCode.LimitRequestRemainTime
                cnt_remained = self.instCpStockCode.GetLimitRemainCount(1)

















