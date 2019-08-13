import sys , csv
from PyQt5.QtWidgets import  *
from PyQt5 import uic
from PyQt5.QtGui import *
import DesinAPI
import pandas as pd
import re
import math
import os

import datetime
from os import listdir
from os.path import isfile, join

form_class = uic.loadUiType("DesinGUI.ui")[0]
form_class_2 = uic.loadUiType("AllDialog.ui")[0]

mT = "m"
class AllDialog(QDialog,  form_class_2 ) :
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.pushButton.setEnabled(False)
        self.pushButton_4.setEnabled(False)
        self.lineEdit.setEnabled(False)
        self.lineEdit_2.setEnabled(False)
        self.lineEdit_4.setText("1")

        self.radioButton.clicked.connect(self.radioButtonClicked1)
        self.radioButton_2.clicked.connect(self.radioButtonClicked2)
        self.radioButton_3.clicked.connect(self.radioButtonClicked3)
        self.radioButton_4.clicked.connect(self.radioButtonClicked4)
        self.radioButton_5.clicked.connect(self.radioButtonClicked5)
        self.radioButton_6.clicked.connect(self.radioButtonClicked6)
        self.radioButton_7.clicked.connect(self.radioButtonClicked7)
        self.radioButton_8.clicked.connect(self.radioButtonClicked8)
        self.radioButton_9.clicked.connect(self.radioButtonClicked9)
        self.someCodeList = False
        self.boolDate = False
        self.mT = 'm'
        self.cospi_cosdaq = 1

        self.pushButton.clicked.connect(self.btn_clicked)
        self.pushButton_2.clicked.connect(self.btn_clicked_2)

        self.pushButton_3.clicked.connect(self.btn_clicked_3)

        self.pushButton_4.clicked.connect(self.btn_clicked_4)



    # MyWindow클래스에서 DesinAPI 객체 가져오기 .
    def getConn(self,conn):
        self.conn = conn



    # 특정 종목 리스트가 담겨있는 csv파일 불러오기.
    def btn_clicked(self):
        try :
            QMessageBox.about(self, "주의", "하나의 csv파일안에 종목코드들이 1열로 있는 경우만 가능합니다. ")
            filter = "csv(*.csv)"
            filePath = QFileDialog.getOpenFileName(self, filter=filter)
            print(len(filePath), filePath[0])
            dataList = pd.read_csv(str(filePath[0]), encoding= 'euc-kr')
            self.codeList = dataList.iloc[:, 0]
            print(len(self.codeList))
        except Exception as e :
            print( e )





    # 폴더내의 종목코드명으로 가져오기.
    def btn_clicked_4(self):
        QMessageBox.about(self, "주의","폴더내의 파일이름이 종목코드인 경우만 가능합니다. \n즉 폴더내의 파일이름들만 가져옵니다. ")

        try :
            filePath = QFileDialog.getExistingDirectory(self, '폴더를 선택해주세요')
            print(filePath)
            file_list_csv = os.listdir(filePath)

            self.codeList = []
            for file in file_list_csv :
                self.codeList.append(re.split('\.csv',file)[0])

            print(self.codeList)
        except Exception as ex :
            print(ex)



    # 새로 저장하기
    def btn_clicked_2(self):
        pb = QProgressBar(self)

        filePath = QFileDialog.getExistingDirectory(self, '폴더를 선택해주세요')
        lastDate = ""
        recentDate = ""
        count = 0


        # 날짜이며, 모든 종목일 때,
        if self.boolDate and self.someCodeList == False :
            codeList = self.conn.GetCodeList(self.cospi_cosdaq)

            lastDate = self.lineEdit.text()
            recentDate = self.lineEdit_2.text()



            print(len(codeList), "개 남음.")

            try :

                for code in codeList :
                    print(code , "시작")
                    if self.mT == 'D':
                        print()
                        self.conn.GetUpdatePeriodDay(code, lastDate, recentDate)
                        self.conn.df.to_csv(str(filePath) + "/" + code + ".csv", mode='a', index=False,
                                            encoding="euc-kr")
                    else:
                        self.conn.GetUpdatePeriodMinutes(code, lastDate, recentDate, self.mT)
                        self.conn.df.to_csv(str(filePath) + "/" + code + ".csv", mode='a', index=False,
                                            encoding="euc-kr")

                # buttonReplpy = QMessageBox.question(self, '안내', '하나로 합치시겠습니까? \n\nresult.csv 로 저장됩니다.', QMessageBox.Yes | QMessageBox.No,  QMessageBox.No)
                # if buttonReplpy == QMessageBox.Yes:


            except Exception  as ex:
                print(ex)
                QMessageBox.about(self, "주의", "날짜를 입력해주세요.")



        # 날짜이며, 특정 종목일 때,
        elif self.boolDate and self.someCodeList :

            lastDate = self.lineEdit.text()
            recentDate = self.lineEdit_2.text()

            print(lastDate,recentDate)
            print(len(self.codeList))
            try :
                for code in self.codeList :
                    print(code)

                    if self.mT == 'D':
                        print()
                        self.conn.GetUpdatePeriodDay(code, lastDate, recentDate)
                        self.conn.df.to_csv(str(filePath) + "/" + code + ".csv", mode='a', index=False,
                                            encoding="euc-kr")
                    else:
                        self.conn.GetUpdatePeriodMinutes(code, lastDate, recentDate,self.mT)
                        self.conn.df.to_csv(str(filePath) + "/" + code + ".csv", mode='a', index=False,
                                            encoding="euc-kr")
                # buttonReplpy = QMessageBox.question(self,'안내', '하나로 합치시겠습니까? \n\nresult.csv 로 저장됩니다.', QMessageBox.Yes | QMessageBox.No,QMessageBox.No)
                # if buttonReplpy == QMessageBox.Yes:


            except Exception  as ex:
                print(ex)
                QMessageBox.about(self, "주의", "날짜를 입력해주세요.")


        # 개수이며, 모든 종목일 때,
        elif self.boolDate == False and self.someCodeList == False :
            codeList = self.conn.GetCodeList(self.cospi_cosdaq)
            tick_range = int(self.lineEdit_4.text())

            try :
                count = int(self.lineEdit_3.text())
                for code in codeList:
                    if self.mT == 'D':
                        print()
                        self.conn.GetDayData(code, count, tick_range, pb)
                        self.conn.df.to_csv(str(filePath) + "/" + code + ".csv", mode='a', index=False,
                                            encoding="euc-kr")
                    else:
                        self.conn.GetMinuteOrTickData(code, count, tick_range, pb, self.mT)
                        self.conn.df.to_csv(str(filePath) + "/" + code + ".csv", mode='a', index=False,
                                            encoding="euc-kr")
                    print(code)
            except Exception  as ex:
                print(ex)
                QMessageBox.about(self, "주의", "개수를 입력해주세요.")



        # 개수이며, 특정 종목일 때,
        elif self.boolDate == False and self.someCodeList :
            tick_range = int(self.lineEdit_4.text())

            try :
                count = int(self.lineEdit_3.text())
                for code in self.codeList :
                    if self.mT == 'D' :
                        print()
                        self.conn.GetDayData(code, count,tick_range, pb )
                        self.conn.df.to_csv(str(filePath) + "/" + code + ".csv", mode='a', index=False,
                                            encoding="euc-kr")
                    else :
                        self.conn.GetMinuteOrTickData(code, count, tick_range, pb,self.mT)
                        self.conn.df.to_csv(str(filePath) + "/" + code + ".csv", mode='a', index=False, encoding="euc-kr")
                    print(code)
            except Exception  as ex:
                print(ex)
                QMessageBox.about(self, "주의", "개수를 입력해주세요.")





    # 새로운 데이터를 기존의 데이터와 결합하기.
    def btn_clicked_3(self):
        QMessageBox.about(self,"주의","날짜 선택하는 건 추천하지 않음. \n모든 종목과 개수로 \n여유를 두어 덮어씌우는 걸 추천. ")
        filePath = QFileDialog.getExistingDirectory(self, '폴더를 선택해주세요')
        pb = QProgressBar(self)


        lastDate = ""
        recentDate = ""
        count = 0

        # 날짜이며, 모든 종목일 때,
        if self.boolDate and self.someCodeList == False:
            dict = {'date': [], 'time': [], 'open': [], 'high': [], 'low': [], 'close': [], 'volume': []}
            result = pd.DataFrame(dict)
            lastDate = self.lineEdit.text()
            recentDate = self.lineEdit_2.text()
            codeList = self.conn.GetMarketCode(self.cospi_cosdaq)

            try:
                for code in codeList:
                    self.conn.GetPeriodMinute(code, lastDate, recentDate)
                    addNewFileOnOldFile(self.conn.df, code, str(filePath))
                    print(code)


            except Exception  as ex:
                print(ex)
                QMessageBox.about(self, "주의", "날짜를 입력해주세요.")




        # 날짜이며, 특정 종목일 때,
        elif self.boolDate and self.someCodeList:
            dict = {'date': [], 'time': [], 'open': [], 'high': [], 'low': [], 'close': [], 'volume': []}
            result = pd.DataFrame(dict)
            lastDate = self.lineEdit.text()
            recentDate = self.lineEdit_2.text()

            print(lastDate, recentDate)

            a = 0
            try :
                for code in self.codeList:
                    a += 1
                    print(a)
                    print(code)
                    self.conn.GetPeriodMinute(code, lastDate, recentDate)
                    print(self.conn.df)

                    addNewFileOnOldFile(self.conn.df, code,str(filePath))

                    print(code)
            except Exception  as ex:
                print(ex)
                QMessageBox.about(self, "주의", "날짜를 입력해주세요.")



        # 개수이며, 모든 종목일 때,
        elif self.boolDate == False and self.someCodeList == False:

            codeList = self.conn.GetCodeList(self.cospi_cosdaq)
            tick_range = int(self.lineEdit_4.text())

            try :

                for code in codeList:
                    if self.mT == 'D':
                        print()
                        addNewFileOnOldFileDay(self.conn,code,pb,str(filePath))
                    else:
                        print("주기 : ",self.mT)
                        addNewFileOnOldFileMinute(self.conn, code, pb, str(filePath), self.mT)

            except Exception  as ex:
                print(ex)
                QMessageBox.about(self, "주의", "개수를 입력해주세요. \nex( 200000 )")


        # 개수이며, 특정 종목일 때,
        elif self.boolDate == False and self.someCodeList:

            tick_range = int(self.lineEdit_4.text())
            try :

                for code in self.codeList:
                    if self.mT == 'D':
                        print()
                        addNewFileOnOldFileDay(self.conn, code, pb, str(filePath))
                    else:
                        print("주기 : ", self.mT)
                        addNewFileOnOldFileMinute(self.conn, code, pb, str(filePath), self.mT)

                    print(code)
            except Exception  as ex:
                print(ex)
                QMessageBox.about(self, "주의", "개수를 입력해주세요. \nex( 200000 )")











    def radioButtonClicked1(self):
        self.someCodeList = False
        if self.radioButton.isChecked():
            self.radioButton_2.setChecked(False)
            self.radioButton_7.setEnabled(False)
            self.radioButton_8.setEnabled(False)
            self.pushButton.setEnabled(True)
            self.pushButton_4.setEnabled(True)
            self.someCodeList = True

    def radioButtonClicked2(self):
        if self.radioButton_2.isChecked():
            self.radioButton.setChecked(False)
            self.radioButton_7.setEnabled(True)
            self.radioButton_8.setEnabled(True)
            self.pushButton.setEnabled(False)
            self.pushButton_4.setEnabled(False)
            self.someCodeList = False

    def radioButtonClicked3(self):
        self.boolDate = False
        if self.radioButton_3.isChecked():
            self.radioButton_4.setChecked(False)
            self.radioButton_5.setEnabled(True)
            self.radioButton_6.setEnabled(True)
            self.radioButton_5.setChecked(True)
            self.radioButton_6.setChecked(False)

            self.radioButton_9.setEnabled(True)
            self.lineEdit_3.setEnabled(False)
            self.lineEdit.setEnabled(True)
            self.lineEdit_2.setEnabled(True)
            self.lineEdit_4.setEnabled(False)
            self.pushButton_3.setEnabled(False)

            self.boolDate = True

    def radioButtonClicked4(self):
        if self.radioButton_4.isChecked():
            self.radioButton_3.setChecked(False)
            self.radioButton_5.setChecked(True)
            self.radioButton_5.setEnabled(True)
            self.radioButton_6.setEnabled(True)
            self.radioButton_9.setEnabled(True)
            self.radioButton_9.setChecked(False)
            self.lineEdit_3.setEnabled(True)
            self.lineEdit.setEnabled(False)
            self.lineEdit_2.setEnabled(False)
            self.lineEdit_4.setEnabled(True)
            self.boolDate = False
            self.pushButton_3.setEnabled(True)

    def radioButtonClicked5(self):
        self.mT = 'm'
        if self.radioButton_5.isChecked():
            self.radioButton_6.setChecked(False)
            self.radioButton_9.setChecked(False)
            self.lineEdit_4.setEnabled(True)
            self.mT = 'm'

    def radioButtonClicked6(self):
        if self.radioButton_6.isChecked():
            self.radioButton_5.setChecked(False)
            self.radioButton_9.setChecked(False)
            self.lineEdit_4.setEnabled(True)
            self.mT = 'T'

    def radioButtonClicked7(self):
        self.cospi_cosdaq = 1
        if self.radioButton_7.isChecked():
            self.radioButton_8.setChecked(False)
            self.cospi_cosdaq = 1


    def radioButtonClicked8(self):
        if self.radioButton_8.isChecked():
            self.radioButton_7.setChecked(False)
            self.cospi_cosdaq = 2

    def radioButtonClicked9(self):
        self.mT = 'D'
        if self.radioButton_9.isChecked():
            self.radioButton_5.setChecked(False)
            self.radioButton_6.setChecked(False)
            self.lineEdit_4.setEnabled(False)
            self.mT = 'D'








class MyWindow(QMainWindow, form_class) :




    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.lineEdit.returnPressed.connect(self.btn_clicked)
        self.lineEdit_3.returnPressed.connect(self.btn_clicked)
        self.lineEdit_4.returnPressed.connect(self.btn_clicked)
        self.lineEdit_4.setText('1')
        self.pushButton.clicked.connect(self.btn_clicked)
        self.pushButton_2.clicked.connect(self.btn_clicked2)
        self.tableWidget.setRowCount(6)
        self.tableWidget.setColumnCount(7)
        self.conn = DesinAPI.DesinAPI()
        # self.conn.GetRecentAllDataFromNumber("A005930",130,120,self.progressBar, 'm')
        # self.conn.GetRecentAllDataFromNumber("A005930", 350, 1, self.progressBar, 'D')
        # self.conn.OnReceived()
        # self.conn.GetSeveralCode()
        # self.conn.GetPeriodMinute('A051900','20180504','20180504')





#############        리스트중에 다운받은 데이터가있는지 확인,
#                    다운받은 데이터중에 가장 최근 날짜를 가져옴.
#                    그리고 현재날짜부터 가장 최근 날짜까지 데이터를 가져온다 .
#
#         self.conn.GetUpdatePeriod("A005930","20190802","20190701" ,'1', 'D')
#         self.conn.df.to_csv('0812/A005930.csv', mode='w', index=False, encoding='euc-kr')

        # currData = pd.read_csv("0812/A005930.csv", index_col=False, encoding="euc-kr")
        # print("가장 빠른 날짜")
        # recentDay = currData.iloc[-1,][0]
        # currLastDateIndex = currData[currData['날짜']==recentDay].index.values[0]
        #
        #
        # #오늘 날짜,
        # today = datetime.date.today().strftime("%Y%m%d")
        # # 오늘날짜부터 최근날짜까지 데이터 불러옴.
        # self.conn.GetUpdatePeriod("A005930",today,str(recentDay), "1",'D')
        #
        #
        # # 최근데이터와 오늘데이터를 합침.
        # updateData = pd.concat([currData.iloc[:currLastDateIndex,],self.conn.df])
        # updateData.to_csv("0812/미포함.csv", mode='a', index=False,
        #                encoding="euc-kr")

        # print(updateData)




        ### 시간 맞춰서 실행.
        # targetTime = '2019-08-06 14:11:00'
        # thisTime = datetime.datetime.strptime(targetTime,  '%Y-%m-%d %H:%M:%S')
        # print(thisTime)
        # while True :
        #     # print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        #     if targetTime == datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'):
        #         # print(datetime.datetime.now())
        #         print("싲가!!!!")
        #
        #         self.cprice = 0
        #         objCur = []
        #         objCur.append(DesinAPI.CpStockCur())
        #         objCur[0].Subscribe("A005930", self.cprice)
        #         break
        #
        # print(datetime.datetime.now())




        ###### 주식종목 실시간조회 1개.
        # self.cprice = 0
        # objCur = []
        # objCur.append(DesinAPI.CpStockCur())
        # objCur[0].Subscribe("A005930", self.cprice)

        ###### 주식종목 실시간조회 1개.





        ########## 코스피,코스닥 지수 가져오기
        # koreaList = ["U180","U201"]
        # for code in koreaList :
        #     self.conn.GetRecentDataFromNumber2(code, 200000,1,self.progressBar,'m')




#########국가지수 실시간 조회
        # self.conn.getWorld()
        # worldList  = [".DJI","COMP","SPX","JP#NI225","HK#HS","SHANG"]
        # for i in worldList :
        #     self.conn.getWorldData(i,4000)
        #
        # objCur = []
        # for i in range(len(worldList)) :
        #     objCur.append(DesinAPI.WorldCur())
        #     objCur[i].Subscribe(worldList[i])


        # self.conn.getWorldData("COMP", 3660)

################## 실시간 시간외 수신
        # self.cprice = 0
        # objCur = []
        # objCur.append(DesinAPI.StockUniCur())
        # objCur[0].Subscribe("A005930")



########################### 주문하기
        # self.aa = DesinAPI.CpFutureOrder()
        # print(self.aa.acc)
        #
        #
        # rtMst = DesinAPI.stockPricedData()
        # current = DesinAPI.CpRPCurrentPrice()
        # current.Request("A102280",rtMst)
        #
        # print(rtMst.cur)
        #
        # order = DesinAPI.CpFutureOrder()
        # order.buyOrder(1195, 1)

        ###########################

########### 취소하기
        # cancel = DesinAPI.CpCancel()
        # cancel.Cancel('1066','A102280')

################## 증거금포함 가능 수량 확인
        #
        # rtMst = DesinAPI.stockPricedData()
        # current = DesinAPI.CpRPCurrentPrice()
        #
        # current.Request("A005930", rtMst)
        # print(rtMst.cur)
        # self.conn.getMargin("A005930",rtMst.cur)



########## ETF 시장종류가 무엇인지,
        # self.conn.GetMarketCode("A114800")




##########  ETF 일봉, 분봉 다운
        # dataList = pd.read_csv("etfList2.csv", encoding='euc-kr')
        # self.codeList = dataList.iloc[:, 0]


        #for code in self.codeList:
         #   self.conn.GetRecentAllDataFromNumber3(code, 200000, 1, self.progressBar, "m")

        # for code in self.codeList:
        #     self.conn.GetRecentAllDataFromNumber4(code, 200000, 1, self.progressBar, "D")


        # self.conn.GetRecentAllDataFromNumber3("A005930",200000,1,self.progressBar,"m")








######## 삼성전자 장전후 미포함 거래량만 뽑아내기.
        sam1 = pd.read_csv("A005930yes.csv", encoding='euc-kr')
        sam2 = pd.read_csv("A005930no.csv",encoding='euc-kr')
        #
        # # print(sam1)
        # samsung1 = sam1[(sam1['시간']==1530) | (sam1['시간']==901)]
        # # print(sam1901)
        # samsung2 = sam2[(sam2['시간']==1530) | (sam2['시간']==901)]
        # samsung1.to_csv("samsung1.csv", index=False, encoding = "euc-kr", mode = 'w')
        # samsung2.to_csv("samsung2.csv", index=False, encoding="euc-kr", mode='w')

        # sam1 = pd.read_csv("samsung1.csv", encoding='euc-kr')
        # sam2 = pd.read_csv("samsung2.csv", encoding='euc-kr')
        #
        # sam1['거래량'] = sam1['거래량'] - sam2['거래량']
        # # print(sam1)
        # sam1['거래대금'] = sam1['거래대금'] - sam2['거래대금']
        #
        # sam1.to_csv("samsungOriginal.csv", mode = 'w' , encoding= 'euc-kr', index= False)
        # print(sam1)
        #
        # date15 = sam1[sam1['시간']==1530]['날짜']
        # date9 = sam1[sam1['시간']==901]['날짜']
        #
        # lastVol = sam1[sam1['시간']==1530]['거래량']
        # firstVol = sam1[sam1['시간']==901]['거래량']
        # lastPay = sam1[sam1['시간']==1530]['거래대금']
        # firstPay = sam1[sam1['시간']==901]['거래대금']
        #
        # samsung = pd.DataFrame({'날짜':date9, '거래량':firstVol, '거래대금': firstPay})
        # samsung.to_csv('samsungBefore.csv', mode='w', encoding='euc-kr', index=False)
        # samsung = pd.DataFrame({'날짜': date15, '거래량': lastVol, '거래대금': lastPay})
        # samsung.to_csv('samsungAfter.csv', mode='w', encoding='euc-kr', index=False)

        # samsung = pd.DataFrame({'날짜':date, '장전거래량':firstVol, '장종료거래량':lastVol, '장전거래대금':firstPay, '장종료거래대금':lastPay})
        # samsung.to_csv('samsungChange.csv',mode = 'w', encoding = 'euc-kr', index= False )




########### 거래정지가 된 종목이 잇는지 ㅗ학인
        # self.conn.GetMarketCode('df')














        # self.conn.GetSeveralCode()
        # objCur = []
        #
        # for i in range(len(self.conn.codeList)) :
        #
        #     objCur.append(DesinAPI.CpStockCur())
        #     objCur[i].Subscribe(self.conn.codeList[i])





        self.label_4.setText(self.conn.result)
        if self.label_4.text() == "연결되지 않았습니다." :
            self.pushButton.setEnabled(False)
            self.pushButton_2.setEnabled(False)
            self.pushButton_3.setEnabled(False)
            self.pushButton_4.setEnabled(False)

        elif self.label_4.text() == "연결되었습니다." :
            self.pushButton.setEnabled(True)
            self.pushButton_2.setEnabled(True)
            self.pushButton_3.setEnabled(True)
            self.pushButton_4.setEnabled(True)

        else :
            self.pushButton.setEnabled(True)
            self.pushButton_2.setEnabled(True)
            self.pushButton_3.setEnabled(True)
            self.pushButton_4.setEnabled(True)



        self.progressBar.setValue(0)
        self.setTableWidgetData()
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.mT = "m"
        self.radioButton.setChecked(True)
        self.radioButton.clicked.connect(self.radioButtonClicked)
        self.radioButton_2.clicked.connect(self.radioButtonClicked)
        self.radioButton_3.clicked.connect(self.radioButtonClicked)
        # self.radioButton.setChecked(True)


        self.lineEdit_2.returnPressed.connect(self.btn_clicked3)
        self.pushButton_3.clicked.connect(self.btn_clicked3)
        self.tableWidget_2.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.pushButton_4.clicked.connect(self.btn_clicked4)






    def setTableWidgetData(self):
        data_column_header = ["날짜","시간","시가","고가","저가","종가","거래량"]
        self.tableWidget.setHorizontalHeaderLabels(data_column_header)

        nameList_column_header = ["종목명","코드명"]
        self.tableWidget_2.setHorizontalHeaderLabels(nameList_column_header)



    def btn_clicked(self):

        column_idx_lookup = {'날짜': 0, '시간': 1, '시가': 2,'고가':3,'저가':4,'종가':5,'거래량':6}

        self.codeName = self.lineEdit.text()
        rqCount = self.lineEdit_3.text()

        tick_range = self.lineEdit_4.text()
        try :

            self.conn.GetRecentDataFromNumber2(self.codeName, int(rqCount),int(tick_range), self.progressBar, self.mT)
            # print(self.conn.numData)
            self.data = self.conn.dict

            rowCnt = len(self.data['날짜'])
            self.tableWidget.setRowCount(rowCnt)

            for k,v in self.data.items() :
                col = column_idx_lookup[k]
                for row,val in enumerate(v) :

                    item = QTableWidgetItem(val)
                    self.tableWidget.setItem(row,col,QTableWidgetItem(str(val)))
        except Exception  as ex :
            print(ex)
            QMessageBox.about(self,"주의", "종목코드를 입력하고, \n요청개수를 입력하고, \n주기를 입력해주세요.  \n누르세요.")




    def btn_clicked2(self):

        try :


            filePath = QFileDialog.getExistingDirectory(self, '폴더를 선택해주세요')
            self.conn.df.to_csv(str(filePath)+"/"+self.codeName+".csv",mode='a', index=False, encoding='euc-kr')
        except Exception  as ex :
            print(ex)
            QMessageBox.about(self,"주의", "종목코드를 입력하고, \n요청개수를 입력하여, \n찾기 버튼 클릭 후 누르세요.")


    def btn_clicked3(self):
        column_idx_lookup = {'name' : 0 , 'code' : 1}

        name = self.lineEdit_2.text()

        self.conn.SearchNameList(name)
        rowCnt = len(self.conn.dataDict['name'])
        self.tableWidget_2.setRowCount(rowCnt)
        for k,v in self.conn.dataDict.items() :
            col = column_idx_lookup[k]
            for row,val in enumerate(v) :
                self.tableWidget_2.setItem(row,col,QTableWidgetItem(str(val)))

    def radioButtonClicked(self):
        self.mT = "m"
        if self.radioButton.isChecked():
            print('첫번째선택됨')
            self.mT = "m"
        elif self.radioButton_2.isChecked():
            print("틱선택됨")
            self.mT = "T"
        elif self.radioButton_3.isChecked():
            print("일봉선택됨")
            self.mT = "D"

    def btn_clicked4(self):

        dlg = AllDialog()

        dlg.getConn(self.conn)
        dlg.exec_()



    def btn_clicked6(self):

        addNewFileOnOldFile(self.conn.df, self.codeName)



#--------------------------



def addNewFileOnOldFileDay(conn,code,pb,filePath) :
    print(code, " 시작 ")
    path = filePath + "/" + code + ".csv"
    try:
        currData = pd.read_csv(path, index_col=False, encoding="euc-kr")
    except Exception  as ex:

        conn.GetDayData(code, 1000000,1, pb)
        conn.df.to_csv(path, mode='w', index=False,
                           encoding="euc-kr")

        return True

    # currData = pd.read_csv("0812/A005930.csv", index_col=False, encoding="euc-kr")
    print("가장 빠른 날짜")
    recentDay = currData.iloc[-1,][0]
    currLastDateIndex = currData[currData['날짜'] == recentDay].index.values[0]
    print(recentDay)

    # 오늘 날짜,
    today = datetime.date.today().strftime("%Y%m%d")
    print("오늘 날짜 : ", today)
    # 오늘날짜부터 최근날짜까지 데이터 불러옴.
    print("오늘데이터 ---->>>>")
    conn.GetUpdatePeriodDay(code, today, str(recentDay))



    # 최근데이터와 오늘데이터를 합침.

    updateData = pd.concat([currData.iloc[:currLastDateIndex, ], conn.df])
    updateData.to_csv(path, mode='w', index=False,
                      encoding="euc-kr")

    print(updateData)





def addNewFileOnOldFileMinute(conn,code,pb,filePath,mT ) :
    print(code, " 시작 ")
    path = filePath + "/" + code + ".csv"
    try:
        currData = pd.read_csv(path, index_col=False, encoding="euc-kr")
    except Exception  as ex:
        conn.GetMinuteOrTickData(code, 200000,1, pb, mT)
        conn.df.to_csv(path, mode='w', index=False,
                           encoding="euc-kr")
        return True

    # currData = pd.read_csv("0812/A005930.csv", index_col=False, encoding="euc-kr")
    print("가장 빠른 날짜")
    recentDay = currData.iloc[-1,][0]
    currLastDateIndex = currData[currData['날짜'] == recentDay].index.values[0]
    print(recentDay)


    # 오늘 날짜,
    today = datetime.date.today().strftime("%Y%m%d")
    print("오늘 날짜 : ", today)
    # 오늘날짜부터 최근날짜까지 데이터 불러옴.
    print("오늘데이터 ---->>>>")
    conn.GetUpdatePeriodMinutes(code, today, str(recentDay), mT)


    # 최근데이터와 오늘데이터를 합침.
    updateData = pd.concat([currData.iloc[:currLastDateIndex, ], conn.df])
    updateData.to_csv(path, mode='w', index=False,
                      encoding="euc-kr")

    print(updateData)




def addNewFileOnOldFile(conn,code,tick_range,pb, newFile,codeName,filePath,mT ):
    new_df = newFile.sort_index(ascending=False).reset_index(drop=True)

    path = filePath+"/" + codeName + ".csv"

    try :
        old_df = pd.read_csv(path,  index_col=False, encoding = 'euc-kr')
    except Exception  as ex:

        if mT == 'D' :
            conn.GetDayData(code, 100, tick_range, pb)
            conn.df.to_csv(str(filePath) + "/" + code + ".csv", mode='a', index=False,
                                encoding="euc-kr")
        else :
            conn.GetMinuteOrTickData(code, 20000, tick_range, pb, mT)
            conn.df.to_csv(str(filePath) + "/" + code + ".csv", mode='a', index=False,
                                encoding="euc-kr")
        return True


    # if old_df.columns.size == 8 :
    #     old_df = old_df.iloc[:,1:]

    print(old_df)
    print('완료')
    old_first = old_df.iloc[-1,][0]
    print("가장 빠른 날짜 : ",old_first)
    try :
        old_index = new_df[new_df['날짜'] == old_first].index.values[0]

        print(old_index)
        new_df2 = new_df.iloc[old_index+1:, ]
        print("new_df : ",new_df)
        print(new_df2)

        real_new_df = pd.concat([old_df,new_df2])
        real_new_df.to_csv(path, mode='w', index=False, encoding='euc-kr')
        return True
    except Exception  as ex:
        print(ex)
        explain =  "겹치는 구간이 없습니다. \n\n 가장 최근 데이터는 \n\n" + str(math.trunc(old_first))
        QMessageBox.about(QWidget(), "주의",explain )
        return True







if __name__ == "__main__" :



    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()



