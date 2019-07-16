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

    def GetMarketCode(self, number) :
        instCpCode = win32com.client.Dispatch("CpUtil.CpCodeMgr")
        # 1 - 거래소주식 ( 코스피 ), 2 - 코스닥주식 , 3 - 중소기업(상장) , 5 - KONEX
        self.codeList = instCpCode.GetStockListbyMarket(number)
        kospi = { }
        for code in self.codeList :

            # 1 주권, 10 ETF,  17 ETN
            secondCode = instCpCode.GetStockSectionKind(code)
            if secondCode == 1 :
                name = instCpCode.CodeToName(code)
                kospi[code] = name

        print("총 개수 : ",number , " : ", len(self.codeList))
        # print(self.codeList)
        f = open("C:\\Users\\Administrator\\PycharmProjects\\PracticeDesinApi\\data\\kospi.csv",'w')
        for key, value in kospi.items() :
            f.write("%s,%s\n" % (key,value))
        f.close()

        return self.codeList




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

        print("머지이이이ㅣㅇ/")
        for i in range(0, numData) :
            for j in range(0,numField) :
                print(instStockChart.GetDataValue(j, i), end=" ")
            print()

        self._wait()

     # 지금 해야할건, 기간으로 데이터를 가져오는 것.
    # 기간으로 분봉데이터 가져오기.
    # - 하지만 기간으로는 주,월,분,틱 은 안된다고 하는데 , 해봐야함.
    # 분은 가능함!

    def GetPeriodMinute(self,code,  lastDate, recentDate):
        columns = ['date', 'time', 'open', 'high', 'low', 'close', 'volume']
        self.dict = {'date': [], 'time': [], 'open': [], 'high': [], 'low': [], 'close': [], 'volume': []}

        # 4 - 개수, 5 - 데이터타입, 6 - 일,주,월,분, 9 - 1수정주가
        instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        instStockChart.SetInputValue(0, code)
        instStockChart.SetInputValue(1, ord('1'))
        instStockChart.SetInputValue(2, recentDate)
        instStockChart.SetInputValue(3, lastDate)

        instStockChart.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])
        instStockChart.SetInputValue(6, ord('m'))
        instStockChart.SetInputValue(9, ord('1'))

        duplicatedCount = 0


        while True :
            instStockChart.BlockRequest()
            numData = instStockChart.GetHeaderValue(3)

            if numData != 2856 :
                duplicatedCount += 1
            if duplicatedCount > 1 :
                break
            numField = instStockChart.GetHeaderValue(1)
            for i in range(0, numData):
                for j in range(0, numField):
                    # print(instStockChart.GetDataValue(j, i), end=" ")
                    self.dict[columns[j]].append(instStockChart.GetDataValue(j, i))
                print()


            print("데이터 크기 : ", numData)

        self.df = pd.DataFrame(self.dict)


        self._wait()



    # codeName - 종목코드
    # count  - 갯수
    # progressBar -
    # mT - [ m = 분봉 , T - 틱봉 ]
    def GetRecentDataFromNumber2(self,codeName, count, progressBar, mT):

        columns = ['date','time','open','high','low','close','volume']
        self.dict = {'date':[], 'time':[],'open':[],'high':[],'low':[],'close':[],'volume':[]}

        # 4 - 개수, 5 - 데이터타입, 6 - 일,주,월,분, 9 - 1수정주가
        instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        instStockChart.SetInputValue(0, codeName)
        instStockChart.SetInputValue(1, ord('2'))

        instStockChart.SetInputValue(4,count)
        instStockChart.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])
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

        KospiList = self.GetMarketCode(1)
        KosdaqList = self.GetMarketCode(2)


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

        self.df = pd.DataFrame(self.dict)
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



    def _wait(self):
        time_remained = self.instCpStockCode.LimitRequestRemainTime
        cnt_remained = self.instCpStockCode.GetLimitRemainCount(1)
        print("남은 제한 횟수 : " + str(cnt_remained))

        if cnt_remained <= 0 :
            while cnt_remained <= 0 :
                Time.sleep(time_remained/1000)
                time_remained = self.instCpStockCode.LimitRequestRemainTime
                cnt_remained = self.instCpStockCode.GetLimitRemainCount(1)






