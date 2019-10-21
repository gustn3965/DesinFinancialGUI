import win32com.client
import pandas as pd
import datetime


## 주식종목 하나 실시간 조회 .
# http://money2.daishin.com/e5/mboard/ptype_basic/plusPDS/DW_Basic_Read.aspx?boardseq=299&seq=45&page=2&searchString=%ec%8b%a4%ec%8b%9c%ea%b0%84&prd=&lang=&p=8831&v=8638&m=9508

class CpStockCurEvent:



    def set_params(self, client, caller ):
        self.client = client
        self.caller = caller


    # 이벤트가 발생 할 때, 실행되는 함수
    def OnReceived(self):
        code = self.client.GetHeaderValue(0)
        name = self.client.GetHeaderValue(1)
        diff = self.client.GetHeaderValue(2)
        time = self.client.GetHeaderValue(3)  # 시간
        open = self.client.GetHeaderValue(4)
        high = self.client.GetHeaderValue(5)
        low = self.client.GetHeaderValue(6)
        sellPrice = self.client.GetHeaderValue(7)
        buyPrice =  self.client.GetHeaderValue(8)


        vol = self.client.GetHeaderValue(9)  # 거래량
        cprice = self.client.GetHeaderValue(13)  # 현재가


        timess = self.client.GetHeaderValue(18)  # 초
        exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그

        diff = self.client.GetHeaderValue(2)  # 대비
        cVol = self.client.GetHeaderValue(17)  # 순간체결수량

        expect = self.client.GetHeaderValue(20)  # 장구분플래그


        item = {}
        item['code'] = code
        item['diff'] = diff
        item['time'] = time
        item['open'] = open
        item['high'] = high
        item['low'] = low
        item['sellPrice'] = sellPrice
        item['buyPrice'] = buyPrice
        item['cprice'] = cprice
        item['vol'] = vol
        item['cvol'] = cVol
        item['timess'] = timess
        item['expect'] = expect


        print(item)

        # if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
        #     print("실시간(예상체결)", timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        # elif (exFlag == ord('2')):  # 장중(체결)
        #     print(code + "  " +name)
        #     print("실시간(장중 체결)", time, timess, cprice, " open : " , open , "high : ", high , "low : " ,low, "체결량", cVol, "거래량", vol)



        self.caller.updateCurData(item)



class CpStockCur:

    def __init__(self, data ):
        self.result = ""


    # 특정 업종을 Subscribe 한다. - handler를 통해 이벤트를 받아온다.
    def Subscribe(self, code ):

        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        hadler = win32com.client.WithEvents(self.objStockCur, CpStockCurEvent)
        self.objStockCur.SetInputValue(0, code)
        hadler.set_params(self.objStockCur, self )
        self.objStockCur.Subscribe()

    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()

    def updateCurData(self, items) :

        self.data[items['code']][0].append(items['time'])
        self.data[items['code']][1].append(items['open'])
        self.data[items['code']][2].append(items['high'])
        self.data[items['code']][3].append(items['low'])
        self.data[items['code']][4].append(items['cprice'])
        self.data[items['code']][5].append(items['vol'])
        self.data[items['code']][6].append(items['cvol'])
        self.data[items['code']][7].append(items['timess'])
        self.data[items['code']][8].append(items['sellPrice'])
        self.data[items['code']][9].append(items['buyPrice'])
        self.data[items['code']][10].append(items['diff'])
        self.data[items['code']][11].append(items['expect'])


        ## 15시 19분일 때 실시간으로 발생했던 데이터들을 업종별로 csv 로 저장
        targetTime = '2019-08-22 15:19:00'
        if targetTime == datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'):

            for code in self.data.keys():
                data = pd.DataFrame({'time': self.data[code][0], 'open': self.data[code][1], 'high': self.data[code][2],
                                     'low': self.data[code][3], 'cprice': self.data[code][4],
                                     'vol': self.data[code][5], 'cvol': self.data[code][6],
                                     'timess': self.data[code][7], 'sellPrice': self.data[code][8],
                                     'buyPrice': self.data[code][9], 'diff': self.data[code][10],
                                     'expect': self.data[code][11]})

                data.to_csv('시간내/' + code + ".csv", encoding='euc-kr', index=False)


if __name__ == "__main__" :


    allVol = pd.read_csv('allVol.csv', encoding='euc-kr', index_col=False)
    allVol = allVol.sort_values(['vol'], ascending=False)
    top400 = allVol['code'].to_list()[:400]
    # top400 = 거래량이 많은 순으로 업종 400개


    # csv로 저장하기 위한 빈 데이터
    outData = {}
    data = {}
    uniData = {}
    for code in top400:
        data[code] = [[], [], [], [], [], [], [], [], [], [], [], []]
        outData[code] = [[], [], [], [], []]
        uniData[code] = [[], [], [], [], [], [], [], [], []]

    print(data)

    objCur = []

    # CpStockCur에 데이터를 넣고,
    # handler 를 통해 이벤트를 받아온다,
    for i in range(0, len(top400)):
        objCur.append(CpStockCur(data))
        objCur[i].Subscribe(top400[i])

