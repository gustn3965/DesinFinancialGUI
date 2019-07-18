import sys , csv
from PyQt5.QtWidgets import  *
from PyQt5 import uic
from PyQt5.QtGui import *
import DesinAPI
import pandas as pd
import re

import os

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

        self.radioButton.clicked.connect(self.radioButtonClicked1)
        self.radioButton_2.clicked.connect(self.radioButtonClicked2)
        self.radioButton_3.clicked.connect(self.radioButtonClicked3)
        self.radioButton_4.clicked.connect(self.radioButtonClicked4)
        self.radioButton_5.clicked.connect(self.radioButtonClicked5)
        self.radioButton_6.clicked.connect(self.radioButtonClicked6)
        self.radioButton_7.clicked.connect(self.radioButtonClicked7)
        self.radioButton_8.clicked.connect(self.radioButtonClicked8)
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
            dataList = pd.read_csv(str(filePath[0]))
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
            dict = {'date': [], 'time': [], 'open': [], 'high': [], 'low': [], 'close': [], 'volume': []}
            result = pd.DataFrame(dict)
            lastDate = self.lineEdit.text()
            recentDate = self.lineEdit_2.text()
            codeList = self.conn.GetMarketCode(self.cospi_cosdaq)

            print(len(codeList), "개 남음.")
            a = 0
            try :
                a +=1
                print(a,"번째 완료")
                for code in codeList :
                    self.conn.GetPeriodMinute(code, lastDate,recentDate)
                    self.conn.df.to_csv(str(filePath) + "/" + code  + ".csv", mode='a', index=False,
                                          encoding="euc-kr")
                    result = pd.concat([result,self.conn.df])
                    print(code)

                buttonReplpy = QMessageBox.question(self, '안내', '하나로 합치시겠습니까? \n\nresult.csv 로 저장됩니다.', QMessageBox.Yes | QMessageBox.No,  QMessageBox.No)
                if buttonReplpy == QMessageBox.Yes:
                    result.to_csv(str(filePath) + "/" + "result" + ".csv", mode='a', index=False,
                                  encoding="euc-kr")

            except Exception  as ex:
                print(ex)
                QMessageBox.about(self, "주의", "날짜를 입력해주세요.")





        # 날짜이며, 특정 종목일 때,
        elif self.boolDate and self.someCodeList :
            dict = {'date': [], 'time': [], 'open': [], 'high': [], 'low': [], 'close': [], 'volume': []}
            result = pd.DataFrame(dict)
            lastDate = self.lineEdit.text()
            recentDate = self.lineEdit_2.text()

            print(lastDate,recentDate)
            print(len(self.codeList))
            a = 0
            try :
                for code in self.codeList :
                    a += 1
                    print(a)
                    print(code)
                    self.conn.GetPeriodMinute(code, lastDate,recentDate)
                    self.conn.df.to_csv(str(filePath) + "/" + code + ".csv", mode='a', index=False,
                                        encoding="euc-kr")
                    print(code)
                    result = pd.concat([result, self.conn.df])
                    print(code)
                buttonReplpy = QMessageBox.question(self,'안내', '하나로 합치시겠습니까? \n\nresult.csv 로 저장됩니다.', QMessageBox.Yes | QMessageBox.No,QMessageBox.No)
                if buttonReplpy == QMessageBox.Yes:
                    result.to_csv(str(filePath) + "/" + "result" + ".csv", mode='a', index=False,
                                  encoding="euc-kr")

            except Exception  as ex:
                print(ex)
                QMessageBox.about(self, "주의", "날짜를 입력해주세요.")


        # 개수이며, 모든 종목일 때,
        elif self.boolDate == False and self.someCodeList == False :
            codeList = self.conn.GetMarketCode(self.cospi_cosdaq)

            try :
                count = int(self.lineEdit_3.text())
                for code in codeList:
                    self.conn.GetRecentDataFromNumber2(code, count,1, pb ,self.mT)
                    self.conn.df.to_csv(str(filePath) + "/" + code + ".csv", mode='a', index=False, encoding="euc-kr")
                    print(code)
            except Exception  as ex:
                print(ex)
                QMessageBox.about(self, "주의", "개수를 입력해주세요.")



        # 개수이며, 특정 종목일 때,
        elif self.boolDate == False and self.someCodeList :

            try :
                count = int(self.lineEdit_3.text())
                for code in self.codeList :
                    self.conn.GetRecentDataFromNumber2(code, count,1, pb ,self.mT)
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

            codeList = self.conn.GetMarketCode(self.cospi_cosdaq)


            try :
                count = int(self.lineEdit_3.text())
                for code in codeList:
                    self.conn.GetRecentDataFromNumber2(code, count,1, pb ,self.mT)
                    addNewFileOnOldFile(self.conn.df, code, str(filePath))
                    print(code)
            except Exception  as ex:
                print(ex)
                QMessageBox.about(self, "주의", "개수를 입력해주세요. \nex( 200000 )")


        # 개수이며, 특정 종목일 때,
        elif self.boolDate == False and self.someCodeList:


            try :
                count = int(self.lineEdit_3.text())
                for code in self.codeList:
                    self.conn.GetRecentDataFromNumber2(code, count,1, pb ,self.mT)
                    addNewFileOnOldFile(self.conn.df, code, str(filePath))
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
            self.radioButton_6.setEnabled(False)
            self.lineEdit_3.setEnabled(False)
            self.lineEdit.setEnabled(True)
            self.lineEdit_2.setEnabled(True)
            self.boolDate = True

    def radioButtonClicked4(self):
        if self.radioButton_4.isChecked():
            self.radioButton_3.setChecked(False)
            self.radioButton_6.setEnabled(True)
            self.lineEdit_3.setEnabled(True)
            self.lineEdit.setEnabled(False)
            self.lineEdit_2.setEnabled(False)
            self.boolDate = False

    def radioButtonClicked5(self):
        self.mT = 'm'
        if self.radioButton_5.isChecked():
            self.radioButton_6.setChecked(False)
            self.mT = 'm'

    def radioButtonClicked6(self):
        if self.radioButton_6.isChecked():
            self.radioButton_5.setChecked(False)
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
        # self.conn.GetPeriodMinute('A051900','20180504','20180504')




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

        column_idx_lookup = {'date': 0, 'time': 1, 'open': 2,'high':3,'low':4,'close':5,'volume':6}

        self.codeName = self.lineEdit.text()
        rqCount = self.lineEdit_3.text()

        tick_range = self.lineEdit_4.text()
        try :

            self.conn.GetRecentDataFromNumber2(self.codeName, int(rqCount),int(tick_range), self.progressBar, self.mT)
            # print(self.conn.numData)
            self.data = self.conn.dict

            rowCnt = len(self.data['date'])
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
            data = pd.DataFrame(self.data)

            filePath = QFileDialog.getExistingDirectory(self, '폴더를 선택해주세요')
            data.to_csv(str(filePath)+"/"+self.codeName+".csv",mode='a', header=False)
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


        else :
            print("틱선택됨")
            self.mT = "T"

    def btn_clicked4(self):

        dlg = AllDialog()

        dlg.getConn(self.conn)
        dlg.exec_()



    def btn_clicked6(self):

        addNewFileOnOldFile(self.conn.df, self.codeName)



#--------------------------








def addNewFileOnOldFile(newFile,codeName,filePath):



    new_df = newFile

    path = filePath+"/" + codeName + ".csv"

    old_df = pd.read_csv(path,  index_col=False)
    if old_df.columns.size == 8 :
        old_df = old_df.iloc[:,1:]

    print(old_df)
    print('완료')
    old_first = old_df.iloc[0,][0]

    old_index = new_df[new_df['date'] == old_first].index.values[0]
    new_df2 = new_df.iloc[:old_index, ]
    real_new_df = pd.concat([new_df2, old_df])
    real_new_df.to_csv(path, mode='w', index=False)





if __name__ == "__main__" :



    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()



