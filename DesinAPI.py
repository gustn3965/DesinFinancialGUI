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



    def SearchNameList(self, name ):
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
        print(self.dataDict)

        print(instCpStockCode.NameToCode(name))








    #  시장구분.
    #  KOSPI = 거래소. 1번.    1527 개
    #  코스닥 = 코스닥  2번     1349 개
    #  K-OTC 중소/중견  3번     137 개
    #  KONEX           5번     150 개


    # 시장구분에 따른 (코스피,코스닥) 종목리스트 얻기.
    # 노터치.
    def GetCodeList(self, sort):
        instCpCode = win32com.client.Dispatch("CpUtil.CpCodeMgr")
        codeList = instCpCode.GetStockListbyMarket(sort)
        print(len(codeList))
        return codeList[:4]

    def GetMarketCode(self, code2) :
        instCpCode = win32com.client.Dispatch("CpUtil.CpCodeMgr")
        chart = win32com.client.Dispatch("CpSysDib.StockChart")

        # 1 - 거래소주식 ( 코스피 ), 2 - 코스닥주식 , 3 - 중소기업(상장) , 5 - KONEX
        # print(instCpCode.GetIndustryList)
        # kospi = {'code':[], 'name':[]}
        allList = []
        for i in range(0,6) :

            self.codeList = instCpCode.GetStockListbyMarket(i)


            for code in self.codeList :
                allList.append(code)



                # # 1 주권, 10 ETF,  17 ETN
                # secondCode = instCpCode.GetStockSectionKind(code)
                # if secondCode == 10 or secondCode == 12  :
                #     name = instCpCode.CodeToName(code)
                #
                #     kospi['code'].append(code)
                #     kospi['name'].append(name)


                    # print(name, " " , code )

        for code in allList :
            chart.SetInputValue(0,code)
            chart.SetInputValue(1,ord('2'))
            chart.BlockRequest()

            if chart.GetHeaderValue(17) == '3' :
                print(code , "거래정지")
            # print(chart.GetHeaderValue(17))

        #
        #
        # print("총 개수 : ", " : ", len(kospi))
        # print("모든 종목 개수 : ", len(allList))
        #
        # etfList = pd.DataFrame(kospi)
        #
        # etfList.to_csv("etfList.csv", mode='a', index=False,
        #                                   encoding="euc-kr")
        #
        #
        #
        #
        # # # print(self.codeList)
        # # f = open("C:\\Users\\Administrator\\PycharmProjects\\PracticeDesinApi\\data\\kospi.csv",'w')
        # # for key, value in kospi.items() :
        # #     f.write("%s,%s\n" % (key,value))
        # # f.close()
        #
        #
        #
        # # 해당 종목의 업종 코드 반환.
        # print(instCpCode.GetStockIndustryCode(code))
        # # 해당 종목의 부구분코드 반환
        # print(instCpCode.GetStockSectionKind(code))

        # print()







        #############  한국지수  업종코드리스트 반환 .
        # print(instCpCode.GetIndustryList())
        # indusList = instCpCode.GetIndustryList()
        # dic = {'code' : [], 'name':[]}
        #
        # for indus in indusList :
        #     # 업종 이름 반환.
        #     print(indus, " ", instCpCode.GetIndustryName(indus))
        #     dic['code'].append(indus)
        #     dic['name'].append(instCpCode.GetIndustryName(indus))
        #
        # data = pd.DataFrame(dic)
        # data.to_csv("koreaIndexList.csv",index=False, mode='a',encoding ='euc-kr')






        #
        #
        # print()
        # print(instCpCode.GetGroupCodeList("24"))


        # return self.codeList



    # 해당종목의 날짜별로  주가를 알 수 있다.
    # https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=102&page=4&searchString=CpSysDib&p=8839&v=8642&m=9508
    def GetRecentDataFromNumber(self) :
        # 4 - 개수, 5 - 데이터타입, 6 - 일,주,월,분, 9 - 1수정주가
        instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        instStockChart.SetInputValue(0, "A035420")
        instStockChart.SetInputValue(1,ord('1'))
        instStockChart.SetInputValue(2,'20190703')
        instStockChart.SetInputValue(3,'20190701')

        instStockChart.SetInputValue(5, [0,1,2])
        instStockChart.SetInputValue(6,ord('m'))
        instStockChart.SetInputValue(9,ord('1'))
        instStockChart.BlockRequest()
        numData = instStockChart.GetHeaderValue(3)

        numField = instStockChart.GetHeaderValue(1)

        for i in range(0, numData) :
            for j in range(0,numField) :
                print(instStockChart.GetDataValue(j, i), end=" ")
            print()

        self._wait()








    # codeName - 종목코드
    # count  - 갯수
    # progressBar -
    # mT - [ m = 분봉 , T - 틱봉 ]
    def GetRecentDataFromNumber2(self,codeName, count,tick_range,  progressBar, mT):

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


    def GetRecentAllDataFromNumber(self,codeName, count,tick_range,  progressBar, mT):

        columns = ['날짜','시간','시가','고가','저가','종가','전일대비','거래량','거래대금','누적체결매도수량','누적체결매수수량','상장주식수','시가총액','외국인주문한도수량','외국인주문가능수량','외국인현보유수량','외국인현보유비율','수정주가일자','수정주가비율','기관순매수','기관누적순매수','등락주선','등락비율','예탁금','주식회전율','거래성립률','대비부호']
        print(len(columns))
        print(len([0,1,2,3,4,5,6,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,37]))
        self.dict = {'날짜':[],'시간':[],'시가':[],'고가':[],'저가':[],'종가':[],'전일대비':[],'거래량':[],'거래대금':[],'누적체결매도수량':[],'누적체결매수수량':[],'상장주식수':[],'시가총액':[],'외국인주문한도수량':[],'외국인주문가능수량':[],'외국인현보유수량':[],'외국인현보유비율':[],'수정주가일자':[],'수정주가비율':[],'기관순매수':[],'기관누적순매수':[],'등락주선':[],'등락비율':[],'예탁금':[],'주식회전율':[],'거래성립률':[],'대비부호':[]}

        # 4 - 개수, 5 - 데이터타입, 6 - 일,주,월,분, 7 - 분봉 주기 ,  9 - 1수정주가
        instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        instStockChart.SetInputValue(0, codeName)
        instStockChart.SetInputValue(1, ord('2'))

        instStockChart.SetInputValue(4,count)
        instStockChart.SetInputValue(5, [0,1,2,3,4,5,6,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,37])
        instStockChart.SetInputValue(6, ord(mT))
        instStockChart.SetInputValue(7, 1)
        instStockChart.SetInputValue(9, ord('1'))




        # 요청하는 값에 비해, 한번에 받을 수 있는 개수는 6665개.
        # 그러므로, 요청을 반복해야한다. + time.sleep을 걸어준다.
        # 요청값이 총 받은 개수보다 크면, ( 같아도 False) stop.
        # 참고 https://github.com/gyusu/Creon-Datareader/blob/master/creonAPI.py
        rcv_count = 0

        progressBar.setMinimum(rcv_count)
        progressBar.setMaximum(count)




        duplicatedCount = 0

        while count > rcv_count :
            progressBar.setValue(rcv_count)

            instStockChart.BlockRequest()
            Time.sleep(0.25)
            self.numData = instStockChart.GetHeaderValue(3)
            self.numData = min(self.numData, count - rcv_count)


            print("받은 데이타 : ",self.numData)

            if self.numData != 739 :
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

        self.df = pd.DataFrame(self.dict)
        self.df.to_csv(codeName+".csv", mode='a', index=False,
                                          encoding="euc-kr")
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



    def GetRecentDataFromDate(self):
        instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        instStockChart.SetInputValue(0,"A035420")
        instStockChart.SetInputValue(1,ord('1'))
        instStockChart.SetInputValue(2,20190705)
        instStockChart.SetInputValue(3,20190625)
        instStockChart.SetInputValue(5,[0,5])
        instStockChart.SetInputValue(6,ord('D'))
        instStockChart.SetInputValue(9,ord('1'))
        instStockChart.BlockRequest()
        numData = instStockChart.GetHeaderValue(3)
        numField = instStockChart.GetHeaderValue(1)
        for i in range(0, numData) :
            for j in range(0, numField) :
                print(instStockChart.GetDataValue(j,i), end = " " )
            print()


    # 해당 종목코드의 PER,EPS, 분기년월... 100몇가지를 알 수 있다.
    # https://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=131&page=1&searchString=MarketEye&p=8839&v=8642&m=9508
    def GetPER(self):
        instMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")
        instMarketEye.SetInputValue(0,[4,67,70,111])
        instMarketEye.SetInputValue(1,"A035420")

        instMarketEye.BlockRequest()

        print("현재가: ", instMarketEye.GetDataValue(0, 0))
        print("PER: ", instMarketEye.GetDataValue(1, 0))
        print("EPS: ", instMarketEye.GetDataValue(2, 0))
        print("최근분기년월: ", instMarketEye.GetDataValue(3, 0))



    def CatchGoodStock(self, instStockChart, code):
        # 4 - 개수, 5 - 데이터타입, 6 - 일,주,월,분, 9 - 1수정주가
        # instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        instStockChart.SetInputValue(0, code)
        instStockChart.SetInputValue(1, ord('2'))
        instStockChart.SetInputValue(4, 40)
        instStockChart.SetInputValue(5, [8])
        instStockChart.SetInputValue(6, ord('D'))
        instStockChart.SetInputValue(9, ord('1'))
        instStockChart.BlockRequest()

        volumes = []
        numData = instStockChart.GetHeaderValue(3)
        for i in range(0, numData) :
            volumne = instStockChart.GetDataValue(0,i)
            volumes.append(volumne)

        averageVolume = ((sum(volumes) - volumes[0]) / len(volumes) -1)

        if (volumes[0] > averageVolume * 10 ) :
            return 1
        else :
            return 0



    def GetSeveralCode(self):

        df = pd.read_csv('C:/Users/Administrator/PycharmProjects/DesinFinancialGUI/data/Code_List_1.csv', encoding='euc-kr',index_col=False)
        self.codeList = list(df['Code'][:200])
        # codeName = list(df['종목명'])
        print(self.codeList)
        instMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")
        instMarketEye.SetInputValue(0, [1,4,5,6,7,10,22])
        instMarketEye.SetInputValue(1,self.codeList)

        instMarketEye.BlockRequest()

        header = instMarketEye.GetHeaderValue(0)


        # print("현재가: ", instMarketEye.GetDataValue(0, 0))
        # print("PER: ", instMarketEye.GetDataValue(1, 0))
        # print("EPS: ", instMarketEye.GetDataValue(2, 0))
        # print("최근분기년월: ", instMarketEye.GetDataValue(3, 0))

        for i in range(0,len(self.codeList)) :
            print(self.codeList[i] )
            for j in range(0,header) :
                print(str(instMarketEye.GetDataValue(j,i)))


    def OnReceived(self):
        objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        win32com.client.WithEvents()
        objStockCur.SetInputValue(0,"A005930")
        objStockCur.Subscribe()

        print(objStockCur.GetHeaderValue(18))
        print(objStockCur.GetHeaderValue(13))
        print(objStockCur.GetHeaderValue(19))
        print(objStockCur.GetHeaderValue(20))
        if objStockCur.GetHeaderValue(19) == ord('2') :
            print(objStockCur.GetHeaderValue(13))


    def _wait(self):
        time_remained = self.instCpStockCode.LimitRequestRemainTime
        cnt_remained = self.instCpStockCode.GetLimitRemainCount(1)
        print("남은 제한 횟수 : " + str(cnt_remained))

        if cnt_remained <= 0 :
            while cnt_remained <= 0 :
                Time.sleep(time_remained/1000)
                time_remained = self.instCpStockCode.LimitRequestRemainTime
                cnt_remained = self.instCpStockCode.GetLimitRemainCount(1)

    def getWorld(self) :
        object =  win32com.client.Dispatch("CpUtil.CpUsCode")

        dic = {'code':[] , 'name': []}
        va = object.GetUsCodeList(3)
        # print(va)
        for i in va :
            dic['code'].append(i)
            dic['name'].append(object.GetNameByUsCode(i))
            print(i," ", object.GetNameByUsCode(i))

        data = pd.DataFrame(dic)
        # data.to_csv("worldList2.csv", mode='a', index=False,
        #                                   encoding="euc-kr")




    # http://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=89&page=1&searchString=CpUsCode&p=&v=&m=
    def getWorldData(self,world, count):
        self.dict = {"date" :[],"open":[],"high":[],"low":[],"close":[],"volume":[]}
        object = win32com.client.Dispatch("Dscbo1.CpSvr8300")
        columns = ['date','open','high','low','close','volume']
        object.SetInputValue(0,world)
        object.SetInputValue(1, ord('D'))
        object.SetInputValue(3,count)


        rcv_count = 0
        duplicatedCount = 0

        while count > rcv_count :
            object.BlockRequest()
            Time.sleep(0.25)
            self.numData = object.GetHeaderValue(3)

            self.numData = min(self.numData, count - rcv_count)

            print("받은 데이타 : ", self.numData)


            if self.numData != 1820 :
                duplicatedCount += 1
            if duplicatedCount > 1 :
                break


            for i in range(self.numData):
                for j in range(0, 6):
                    # print(object.GetDataValue(j, i))
                    self.dict[columns[j]].append(object.GetDataValue(j, i))

            rcv_count += self.numData


            # 2년치의 데이터가 넘어가면, 최신데이터도 가져오는데,
            # 중복의 최신데이터는 필요없기 때문에,
            # 만약 최신데이터가 나타나면 자르고, break한다.

            if self.numData == 0 :
                break


        self.df = pd.DataFrame(self.dict)
        # self.df.to_csv("worldData/"+world+"이상해씨.csv", mode='a', index=False,
        #                                   encoding="euc-kr")

        print(self.df)
        # print(self.df)

        self._wait()



    def getWorldData8312(self,world,count):
        object = win32com.client.Dispatch("Dscbo1.CpFore8312")
        object.SetInputValue(0,world)
        object.SetInputValue(1, 1)
        object.SetInputValue(2,count)
        object.BlockRequest()

        print()
        print()
        print("받은 개수 : ",object.GetHeaderValue(2))
        print(object.GetHeaderValue(0))
        print(object.GetHeaderValue(3))

        for i in range(object.GetHeaderValue(2)) :
            for j in range(0,3) :
                print(object.GetDataValue(0,i))


    def getMargin(self,code,price):
        ret = g_objCpTrade.TradeInit(0)
        object = win32com.client.Dispatch("CpTrade.CpTdNew5331A")
        object.SetInputValue(0,"335022427")
        object.SetInputValue(1,"10")
        object.SetInputValue(2,"A005930")
        object.SetInputValue(3,"01")
        object.SetInputValue(4,price)
        object.SetInputValue(5,"N")
        object.SetInputValue(6,'2')
        object.BlockRequest()


        print(object.GetHeaderValue(0))
        print(object.GetHeaderValue(4))
        print(object.GetHeaderValue(12))




















class StockUniCurEvent :
    def set_params(self, client ):
        self.client = client
    def OnReceived(self):
        code = self.client.GetHeaderValue(0)
        name = self.client.GetHeaderValue(1)
        time = self.client.GetHeaderValue(2)  # 시간
        open = self.client.GetHeaderValue(3)
        # high = self.client.GetHeaderValue(4)
        low = self.client.GetHeaderValue(5)
        #
        # if self.cprice > 43550 :
        #     self.aa = CpFutureOrder()
        #         #
        #         #     rtMst = stockPricedData()
        #         #     current = CpRPCurrentPrice()
        #         #     current.Request("A005930",rtMst)
        #         #
        #         #     order = CpFutureOrder()
        #         #     order.buyOrder(rtMst.cur, 1)

        print(code , name)
        print(code, name, time, open , low )

# 해외국가지수 실시간 수신
class StockUniCur :
    def Subscribe(self, code):

        self.objStockCur = win32com.client.Dispatch("CpSysDib.StockUniCur")
        hadler = win32com.client.WithEvents(self.objStockCur, StockUniCurEvent)
        self.objStockCur.SetInputValue(0, code)
        hadler.set_params(self.objStockCur)
        self.objStockCur.Subscribe()

    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()






# 해외국가지수 이벤트
class WorldCurEvent :
    def set_params(self, client ):
        self.client = client
    def OnReceived(self):
        code = self.client.GetHeaderValue(0)
        name = self.client.GetHeaderValue(1)
        time = self.client.GetHeaderValue(2)  # 시간
        open = self.client.GetHeaderValue(3)
        high = self.client.GetHeaderValue(4)
        low = self.client.GetHeaderValue(5)
        #
        # if self.cprice > 43550 :
        #     self.aa = CpFutureOrder()
        #         #
        #         #     rtMst = stockPricedData()
        #         #     current = CpRPCurrentPrice()
        #         #     current.Request("A005930",rtMst)
        #         #
        #         #     order = CpFutureOrder()
        #         #     order.buyOrder(rtMst.cur, 1)

        print(code , name)
        print(code, name, time, open, high , low )

# 해외국가지수 실시간 수신
class WorldCur :
    def Subscribe(self, code):

        self.objStockCur = win32com.client.Dispatch("CpSysDib.WorldCur")
        hadler = win32com.client.WithEvents(self.objStockCur, WorldCurEvent)
        self.objStockCur.SetInputValue(0, code)
        hadler.set_params(self.objStockCur)
        self.objStockCur.Subscribe()

    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()








## 주식종목 하나 실시간 조회 .
# http://money2.daishin.com/e5/mboard/ptype_basic/plusPDS/DW_Basic_Read.aspx?boardseq=299&seq=45&page=2&searchString=%ec%8b%a4%ec%8b%9c%ea%b0%84&prd=&lang=&p=8831&v=8638&m=9508
class CpStockCurEvent:


    def set_params(self, client , cprice):
        self.client = client
        self.cprice = cprice

    def OnReceived(self):
        code = self.client.GetHeaderValue(0)
        name = self.client.GetHeaderValue(1)
        time = self.client.GetHeaderValue(3)  # 시간
        open = self.client.GetHeaderValue(4)
        high = self.client.GetHeaderValue(5)
        low = self.client.GetHeaderValue(6)
        timess = self.client.GetHeaderValue(18)  # 초
        exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
        self.cprice = self.client.GetHeaderValue(13)  # 현재가
        diff = self.client.GetHeaderValue(2)  # 대비
        cVol = self.client.GetHeaderValue(17)  # 순간체결수량
        vol = self.client.GetHeaderValue(9)  # 거래량
        print(self.cprice)



        #
        # if self.cprice > 43550 :
        #     self.aa = CpFutureOrder()
        #         #
        #         #     rtMst = stockPricedData()
        #         #     current = CpRPCurrentPrice()
        #         #     current.Request("A005930",rtMst)
        #         #
        #         #     order = CpFutureOrder()
        #         #     order.buyOrder(rtMst.cur, 1)
        # if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
        #     print("실시간(예상체결)", timess, "*", self.cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        # elif (exFlag == ord('2')):  # 장중(체결)
        #     print(code + "  " +name)
        #     print("실시간(장중 체결)", time, timess, self.cprice, " open : " , open , "high : ", high , "low : " ,low, "체결량", cVol, "거래량", vol)


class CpStockCur:
    def Subscribe(self, code, cprice):

        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        hadler = win32com.client.WithEvents(self.objStockCur, CpStockCurEvent)
        self.objStockCur.SetInputValue(0, code)
        hadler.set_params(self.objStockCur, cprice)
        self.objStockCur.Subscribe()

    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()
###################################################












g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
#  https://money2.daishin.com/e5/mboard/ptype_basic/plusPDS/DW_Basic_Read.aspx?boardseq=299&seq=79&page=1&searchString=WithEvents&prd=&lang=&p=8831&v=8638&m=9508

################################################
# PLUS 실행 기본 체크 함수
def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False

    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False

    # # 주문 관련 초기화
    # if (g_objCpTrade.TradeInit(0) != 0):
    #     print("주문 초기화 실패")
    #     return False

    return True


################################################
# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관

    def OnReceived(self):
        # 실시간 처리 - 현재가 주문 체결
        if self.name == 'stockcur':
            code = self.client.GetHeaderValue(0)  # 초
            name = self.client.GetHeaderValue(1)  # 초
            timess = self.client.GetHeaderValue(18)  # 초
            exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
            cprice = self.client.GetHeaderValue(13)  # 현재가
            diff = self.client.GetHeaderValue(2)  # 대비
            cVol = self.client.GetHeaderValue(17)  # 순간체결수량
            vol = self.client.GetHeaderValue(9)  # 거래량
            high = self.client.GetHeaderValue(5)
            low = self.client.GetHeaderValue(6)

            print("-----",timess, code,vol, cVol)

            if exFlag != ord('2'):
                return

            item = {}
            item['code'] = code
            item['time'] = timess
            item['diff'] = diff
            item['cur'] = cprice
            item['vol'] = vol

            # print(item)

            # 현재가 업데이트
            self.caller.updateCurData(item)

            return


################################################
# plus 실시간 수신 base 클래스
class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False

    def Subscribe(self, var, caller):
        if self.bIsSB:
            self.Unsubscribe()

        if (len(var) > 0):
            self.obj.SetInputValue(0, var)

        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, caller)
        self.obj.Subscribe()
        self.bIsSB = True

    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
        self.bIsSB = False


################################################
# CpPBStockCur: 실시간 현재가 요청 클래스
class CpPBStockCur(CpPublish):
    def __init__(self):
        super().__init__('stockcur', 'DsCbo1.StockCur')


class CMinchartData:
    def __init__(self):
        self.minDatas = {}
        self.objCur = {}

    def stop(self):
        for k, v in self.objCur.items():
            v.Unsubscribe()

    def addCode(self, code):
        if (code in self.minDatas):
            return

        self.minDatas[code] = []
        self.objCur[code] = CpPBStockCur()
        self.objCur[code].Subscribe(code, self)
        self._wait()

    def updateCurData(self, item):
        code = item['code']
        time = item['time']
        cur = item['cur']
        vol = item['vol']
        self.makeMinchart(code, time, cur,vol)

    def makeMinchart(self, code, time, cur, vol ):
        hh, mm = divmod(time, 10000)
        # print("받아온 : ", hh,mm )
        mm, tt = divmod(mm, 100)
        mm += 1
        if (mm == 60):
            hh += 1
            mm = 0

        hhmm = hh * 100 + mm
        # print(hhmm)
        if hhmm > 1530:
            hhmm = 1530
        bFind = False
        minlen = len(self.minDatas[code])
        # print(minlen)
        if (minlen > 0):
            # 0 : 시간 1 : 시가 2: 고가 3: 저가 4: 종가
            if (self.minDatas[code][-1][0] == hhmm):
                # print("일떄 : " , hhmm)
                item = self.minDatas[code][-1]
                print(item)
                bFind = True
                item[4] = cur
                if (item[2] < cur):
                    item[2] = cur
                if (item[3] > cur):
                    item[3] = cur

        if bFind == False:
            self.minDatas[code].append([hhmm, cur, cur, cur, cur,vol])

            print(code, self.minDatas[code])



        return

    def print(self, code):
        print('====================================================-')
        print('분데이터 print', code, g_objCodeMgr.CodeToName(code))
        print('시간,시가,고가,저가,종가')
        for item in self.minDatas[code]:
            hh, mm = divmod(item[0], 100)
            print("%02d:%02d,%d,%d,%d,%d" % (hh, mm, item[1], item[2], item[3], item[4]))

    def _wait(self):
        time_remained = g_objCpStatus.LimitRequestRemainTime
        cnt_remained = g_objCpStatus.GetLimitRemainCount(2)
        print("남은 제한 횟수 : " + str(cnt_remained) + " 남은 시간 : " + str(time_remained))

        if cnt_remained <= 0:
            while cnt_remained <= 0:
                Time.sleep(time_remained / 1000)
                time_remained = g_objCpStatus.LimitRequestRemainTime
                cnt_remained = g_objCpStatus.GetLimitRemainCount(2)












######## 주식현재가 조회 및 현금주문하기.
# http://money2.daishin.com/e5/mboard/ptype_basic/plusPDS/DW_Basic_List.aspx?boardseq=299&m=9508&p=8831&v=8638
class stockPricedData:
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

    # 전일 대비 계산
    def makediffp(self, baseprice):
        lastday = 0
        if baseprice:
            lastday = baseprice
        else:
            lastday = self.cur - self.diff
        if lastday:
            self.diffp = (self.diff / lastday) * 100
        else:
            self.diffp = 0

    def debugPrint(self, type):
        if type == 0:
            print("%s, %s %s, 현재가 %d 대비 %d, (%.2f), 1차매도 %d(%d) 1차매수 %d(%d)"
                  % (self.dicEx.get(self.exFlag), self.code,
                     self.name, self.cur, self.diff, self.diffp,
                     self.offer[0], self.offervol[0], self.bid[0], self.bidvol[0]))
        else:
            print("%s %s, 현재가 %.2f 대비 %.2f, (%.2f), 1차매도 %.2f(%d) 1차매수 %.2f(%d)"
                  % (self.code,
                     self.name, self.cur, self.diff, self.diffp,
                     self.offer[0], self.offervol[0], self.bid[0], self.bidvol[0]))

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



# 주식 현금 주문
class CpFutureOrder :
    def __init__(self):
        ret = g_objCpTrade.TradeInit(0)
        print(ret)
        self.acc = g_objCpTrade.AccountNumber[0]
        # 주식 상품 구분
        self.accFalg = g_objCpTrade.GoodsList(self.acc, 1)
        print(self.acc, self.accFalg[0])
        self.objOrder = win32com.client.Dispatch("CpTrade.CpTd0311")

    def Order(self, price, amount):

        # 1- 매도, 2 - 매수
        self.objOrder.SetInputValue(0, '1')
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



######## 주문취소하기
class CpCancel :
    def __init__(self):
        ret = g_objCpTrade.TradeInit(0)
        print(ret)
        # 계좌번호
        self.acc = g_objCpTrade.AccountNumber[0]
        # 주식 상품 구분
        self.accFalg = g_objCpTrade.GoodsList(self.acc, 1)
        print(self.acc, self.accFalg[0])
        self.objOrder = win32com.client.Dispatch("CpTrade.CpTd0314")

    def Cancel(self,number,code):

        # 1- 매도, 2 - 매수

        self.objOrder.SetInputValue(1,number )
        self.objOrder.SetInputValue(2,self.acc)
        self.objOrder.SetInputValue(3,self.accFalg[0])
        self.objOrder.SetInputValue(4,code)
        # 주문 가격은 현재가격으로 되야한다.
        self.objOrder.SetInputValue(5,1)

        ret = self.objOrder.BlockRequest()

        if ret == 4:
            remainTime = g_objCpStatus.LimitRequestRemainTime
            print('연속조회 제한 오류, 남은 시간', remainTime)
            self._wait()
            ret = self.objOrder.BlockRequest()


        print("원주문번호 ", self.objOrder.GetHeaderValue(1))
        print("종목코드 : ", self.objOrder.GetDataValue(4))
        print("취소수량 ",self.objOrder.Getheadervalue(5))
        print("주문번호 ",self.objOrder.GetHeaderValue(6))

    def buyOrder(self, price, amount):
        return self.Order(price, amount)






# 주식 잔고 조회

class Cp6033:
    def __init__(self):
        acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        accFlag = g_objCpTrade.GoodsList(acc, 1)  # 주식상품 구분
        print(acc, accFlag[0])

        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 50)  # 요청 건수(최대 50)

    # 실제적인 6033 통신 처리
    def requestJango(self, caller):
        while True:
            ret = self.objRq.BlockRequest()
            if ret == 4:
                remainTime = g_objCpStatus.LimitRequestRemainTime
                print('연속조회 제한 오류, 남은 시간', remainTime)
                return False
            # 통신 및 통신 에러 처리
            rqStatus = self.objRq.GetDibStatus()
            rqRet = self.objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            if rqStatus != 0:
                return False

            cnt = self.objRq.GetHeaderValue(7)
            print(cnt)

            for i in range(cnt):
                item = {}
                code = self.objRq.GetDataValue(12, i)  # 종목코드
                item['종목코드'] = code
                item['종목명'] = self.objRq.GetDataValue(0, i)  # 종목명
                item['현금신용'] = self.objRq.GetDataValue(1, i)  # 신용구분
                print(code, '현금신용', item['현금신용'])
                item['대출일'] = self.objRq.GetDataValue(2, i)  # 대출일
                item['잔고수량'] = self.objRq.GetDataValue(7, i)  # 체결잔고수량
                item['매도가능'] = self.objRq.GetDataValue(15, i)
                item['장부가'] = self.objRq.GetDataValue(17, i)  # 체결장부단가
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']

                # 잔고 추가
                caller.jangoData[code] = item

                if len(caller.jangoData) >= 200:  # 최대 200 종목만,
                    break

            if len(caller.jangoData) >= 200:
                break
            if (self.objRq.Continue == False):
                break
        return True
