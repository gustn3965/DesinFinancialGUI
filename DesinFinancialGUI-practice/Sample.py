import win32com.client
import pandas as pd
import time as Time
import datetime


class Sample :


    ## 우선적으로 CREON 연결 확인.
    def __init__(self):
        self.result = ""
        self.instCpStockCode = win32com.client.Dispatch('CpUtil.CpCybos')
        print(self.instCpStockCode.IsConnect)
        if self.instCpStockCode.IsConnect == 1:
            print("연결되었습니다.")
            self.result = "연결되었습니다."
        else:
            print("연결 되지 않았습니다. ")
            self.result = "연결되지 않았습니다."



    # 일봉데이터 얻기, - 삼성전자  - A005930
    def GetDayData(self):

        # 도움말
        # https://money2.creontrade.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=284&seq=102&page=1&searchString=StockChart&p=8841&v=8643&m=9505

        columns = ['날짜', '시간', '시가', '고가', '저가', '종가', '전일대비', '거래량', '거래대금', '누적체결매도수량', '누적체결매수수량', '상장주식수',
                   '시가총액',
                   '외국인주문한도수량', '외국인주문가능수량', '외국인현보유수량', '외국인현보유비율', '수정주가일자', '수정주가비율', '기관순매수', '기관누적순매수']


        self.dict = {'날짜': [], '시간': [], '시가': [], '고가': [], '저가': [], '종가': [], '전일대비': [], '거래량': [], '거래대금': [],
                     '누적체결매도수량': [], '누적체결매수수량': [], '상장주식수': [], '시가총액': [], '외국인주문한도수량': [], '외국인주문가능수량': [],
                     '외국인현보유수량': [], '외국인현보유비율': [], '수정주가일자': [], '수정주가비율': [], '기관순매수': [], '기관누적순매수': []}


        # SetInputValue 으로 초기값 설정해줌.
        # SetInputValue 1 - "2" 데이터 개수에따라 요청,
        # SetInputValue 4 - 데이터 개수
        # SetInputValue 5 - 데이터 타입 - columns와 동일
        # SetInputValue 6 - "D" =  Day = 일봉
        # SetInputValue 7 - 주기 * 일봉은 주기 바꿔도 달라지지 않음.
        instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        instStockChart.SetInputValue(0, "A005930")
        instStockChart.SetInputValue(1, ord('2'))
        instStockChart.SetInputValue(4, 3000)
        instStockChart.SetInputValue(5, [0, 1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21])
        instStockChart.SetInputValue(6, ord("D"))
        instStockChart.SetInputValue(7, 2)
        instStockChart.SetInputValue(9, ord('0'))
        instStockChart.SetInputValue(10, ord('1'))

        rcv_count = 0
        duplicatedCount = 0

        ## while 문으로 데이터 반복 요청
        while 3000 > rcv_count:

            instStockChart.BlockRequest()
            Time.sleep(0.25)


            # GetHeaderValue 1   필드 개수   ( SetInputValue 5 에 설정한 받아올 컬럼 개수 )
            # GetHeaderValue 3   데이터 수신 개수

            self.numData = instStockChart.GetHeaderValue(3)
            self.numData = min(self.numData, 3000 - rcv_count)


            print("받은 데이타 : ", self.numData)

            # 데이터 타입 ( columns의 개수 ) 에 따라  한 번에 받을 수 있는 데이터개수가 달라짐.
            # 현재의 columns의 개수에는 한 번에 받을 수 있는 데이터개수는 951개.
            # ex ) 3000개를 요청할 경우 951 951 951 받고 마지막 147개 받음.
            #      951개가 아닌 경우 while 문 종료.
            #      그렇지 않으면 최신데이터를 받아옴.

            if self.numData != 951:
                duplicatedCount += 1
            if duplicatedCount > 2:
                break

            numField = instStockChart.GetHeaderValue(1)


            for i in range(0, self.numData):
                for j in range(0, numField):
                    self.dict[columns[j]].append(instStockChart.GetDataValue(j, i))
            rcv_count += self.numData


            if self.numData == 0:
                break
            self._wait()

        self.df = pd.DataFrame(self.dict).sort_index(ascending=False)
        print(self.df)

        self._wait()



    # 분봉 틱봉 데이터 가져오기  - 삼성전자  A005930
    def GetMinuteOrTickData(self):

        # 분봉 or 틱봉 과 일봉이 가져올 수 있는 컬럼개수가 다름.

        columns = ['날짜', '시간', '시가', '고가', '저가', '종가', '거래량', '거래대금', '누적체결매도수량', '누적체결매수수량']
        self.dict = {'날짜': [], '시간': [], '시가': [], '고가': [], '저가': [], '종가': [], '거래량': [], '거래대금': [],
                     '누적체결매도수량': [], '누적체결매수수량': []}

        # SetInputValue 으로 초기값 설정해줌.
        # SetInputValue 1 - "2" 데이터 개수에따라 요청,
        # SetInputValue 4 - 데이터 개수
        # SetInputValue 5 - 데이터 타입 - columns와 동일
        # SetInputValue 6 - "m" = 분봉 , "T" = 틱봉
        # SetInputValue 7 - 주기 * 1 = 1분 or 1틱 마다, 2 = 2분 or 2틱 마다 ....



        instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        instStockChart.SetInputValue(0, "A000720")
        instStockChart.SetInputValue(1, ord('2'))
        instStockChart.SetInputValue(4, 120000)
        instStockChart.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8, 9, 10, 11])
        instStockChart.SetInputValue(6, ord('m'))
        instStockChart.SetInputValue(7, 4)
        instStockChart.SetInputValue(9, ord('1'))
        # instStockChart.SetInputValue(10, ord('3'))

        rcv_count = 0
        duplicatedCount = 0

        ## while 문으로 데이터 반복 요청
        while 200000 > rcv_count:
            instStockChart.BlockRequest()
            Time.sleep(0.25)
            self.numData = instStockChart.GetHeaderValue(3)
            self.numData = min(self.numData, 200000 - rcv_count)

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

            if self.numData == 0:
                break
            self._wait()

        self.df = pd.DataFrame(self.dict).sort_index(ascending=False)
        print(self.df)

        self._wait()


## _wait()
    def _wait(self):
        time_remained = self.instCpStockCode.LimitRequestRemainTime
        cnt_remained = self.instCpStockCode.GetLimitRemainCount(1)
        print("남은 제한 횟수 : " + str(cnt_remained))

        if cnt_remained <= 0 :
            while cnt_remained <= 0 :
                Time.sleep(time_remained/1000)
                time_remained = self.instCpStockCode.LimitRequestRemainTime
                cnt_remained = self.instCpStockCode.GetLimitRemainCount(1)






    def UpdateIndexList(self):
        # 1. 코스피 지수 리스트
        # 2. 코스닥 지수 리스트 1
        # 3. 코스닥 지수 리스트 2
        # 4. 해외지수 리스트
        # 5. 해외지수 리스트 2
        # 6. ETF 지수 리스트

        now = datetime.datetime.now()
        nowDate = now.strftime('%y%m%d')

        koreaClient = win32com.client.Dispatch("CpUtil.CpCodeMgr")
        worldClient = win32com.client.Dispatch("CpUtil.CpUsCode")




        # 코스닥 산업별 =           GetKosdaqIndustry1List
        # 코스닥 지수업종 코드 =     GetKosdaqIndustry2List
        # 증권전산업종코드  =        GetIndustryList

        #  해외 국가 대표  = GetUsCodeList(2)
        #  해외 업종  =  GetUsCodeList(3 )

        kospiCodeList = koreaClient.GetIndustryList()
        kosdaqCodeList1 = koreaClient.GetKosdaqIndustry1List()
        kosdaqCodeList2 = koreaClient.GetKosdaqIndustry2List()
        worldCodeList1 = worldClient.GetUsCodeList(2)
        worldCodeList2 = worldClient.GetUsCodeList(3)

        codeListString = ["KOSPI","KOSDAQ1","KOSDAQ2","WORLD1","WORLD2"]
        codeLists = []
        codeLists.append(kospiCodeList)
        codeLists.append(kosdaqCodeList1)
        codeLists.append(kosdaqCodeList2)
        codeLists.append(worldCodeList1)
        codeLists.append(worldCodeList2)

        a = 0
        for list in codeLists :

            index = {'code': [], 'name': []}
            # 한국 주식일 때,
            if a < 3 :
                for code in list :
                    index['code'].append(code)
                    index['name'].append(koreaClient.GetIndustryName(code))
                indexList = pd.DataFrame(index)
                indexList_name = codeListString[a] + nowDate + ".csv"
                print("-"*50)
                print("\n"*3)
                print(codeListString[a])
                print(indexList)
                # indexList.to_csv("C:/Users/Administrator/PycharmProjects/PracticeDesinApi/지수리스트/" + codeListString[a] + "/" + indexList_name, index=False, mode='a', encoding='euc-kr')


            # 외국 주식 일 때,
            else :
                for code in list :
                    index['code'].append(code)
                    index['name'].append(worldClient.GetNameByUsCode(code))
                indexList = pd.DataFrame(index)
                indexList_name = codeListString[a] + nowDate + ".csv"
                print("-"*50)
                print("\n"*3)
                print(codeListString[a])
                print(indexList)
                # indexList.to_csv("C:/Users/Administrator/PycharmProjects/PracticeDesinApi/지수리스트/" + codeListString[a] + "/" + indexList_name, index=False, mode='a', encoding='euc-kr')
            a += 1



        #  ETF 일 때,
        etf = {'code': [], 'name': []}
        allList = []
        for i in range(0, 6):

            self.codeList = koreaClient.GetStockListbyMarket(i)

            for code in self.codeList:
                allList.append(code)

                # 1 주권, 10 ETF,  17 ETN
                secondCode = koreaClient.GetStockSectionKind(code)
                if secondCode == 10 or secondCode == 12:
                    name = koreaClient.CodeToName(code)

                    etf['code'].append(code)
                    etf['name'].append(name)

                    # print(name, " " , code )

        etfList = pd.DataFrame(etf)
        print(etfList)
        etf_name = "LIST_ETF_" + nowDate + ".csv"
        # etfList.to_csv("C:/Users/Administrator/PycharmProjects/PracticeDesinApi/지수리스트" + "/ETF/" + etf_name, mode='a', index=False,
        #                encoding="euc-kr")






if __name__ == "__main__" :


    sample = Sample()
    sample.GetDayData()
    # print("-"*50)
    # print("\n"*3)
    # sample.GetMinuteOrTickData()
  #   sample.UpdateIndexList()

    # for i in range(0,20000000) :
    #     print(i)
    #


