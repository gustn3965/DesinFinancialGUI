import sys , csv
from PyQt5.QtWidgets import  *
from PyQt5 import uic
from PyQt5.QtGui import *
import DesinAPI
import pandas as pd

import os

from os import listdir
from os.path import isfile, join



form_class = uic.loadUiType("GetButton.ui")[0]

columnList = []

mT = "m"

class MyWindow(QMainWindow, form_class) :




    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.lineEdit.returnPressed.connect(self.btn_clicked)
        self.lineEdit_3.returnPressed.connect(self.btn_clicked)
        self.pushButton.clicked.connect(self.btn_clicked)
        self.pushButton_2.clicked.connect(self.btn_clicked2)
        self.tableWidget.setRowCount(6)
        self.tableWidget.setColumnCount(7)
        self.conn = DesinAPI.DesinAPI()
        self.label_4.setText(self.conn.result)
        if self.label_4.text() == "연결되지 않았습니다." :
            self.pushButton.setEnabled(False)
            self.pushButton_2.setEnabled(False)
            self.pushButton_3.setEnabled(False)
            self.pushButton_4.setEnabled(False)
            self.pushButton_5.setEnabled(False)
            self.pushButton_6.setEnabled(False)
        elif self.label_4.text() == "연결되었습니다." :
            self.pushButton.setEnabled(True)
            self.pushButton_2.setEnabled(True)
            self.pushButton_3.setEnabled(True)
            self.pushButton_4.setEnabled(True)
            self.pushButton_5.setEnabled(True)
            self.pushButton_6.setEnabled(True)
        else :
            self.pushButton.setEnabled(True)
            self.pushButton_2.setEnabled(True)
            self.pushButton_3.setEnabled(True)
            self.pushButton_4.setEnabled(True)
            self.pushButton_5.setEnabled(True)
            self.pushButton_6.setEnabled(True)


        self.progressBar.setValue(0)
        self.setTableWidgetData()
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.radioButton.clicked.connect(self.radioButtonClicked)
        self.radioButton_2.clicked.connect(self.radioButtonClicked)
        # self.radioButton.setChecked(True)


        self.lineEdit_2.returnPressed.connect(self.btn_clicked3)
        self.pushButton_3.clicked.connect(self.btn_clicked3)
        self.tableWidget_2.setEditTriggers(QAbstractItemView.NoEditTriggers)

        self.pushButton_4.clicked.connect(self.btn_clicked4)

        self.pushButton_6.clicked.connect(self.btn_clicked6)




    def setTableWidgetData(self):
        data_column_header = ["날짜","시간","시가","고가","저가","종가","거래량"]
        self.tableWidget.setHorizontalHeaderLabels(data_column_header)

        nameList_column_header = ["종목명","코드명"]
        self.tableWidget_2.setHorizontalHeaderLabels(nameList_column_header)



    def btn_clicked(self):
        column_idx_lookup = {'date': 0, 'time': 1, 'open': 2,'high':3,'low':4,'close':5,'volume':6}

        self.codeName = self.lineEdit.text()
        rqCount = self.lineEdit_3.text()


        self.conn.GetRecentDataFromNumber2(self.codeName, int(rqCount), self.progressBar, self.mT)
        # print(self.conn.numData)
        self.data = self.conn.dict

        rowCnt = len(self.data['date'])
        self.tableWidget.setRowCount(rowCnt)

        for k,v in self.data.items() :
            col = column_idx_lookup[k]
            for row,val in enumerate(v) :

                item = QTableWidgetItem(val)
                self.tableWidget.setItem(row,col,QTableWidgetItem(str(val)))

    def btn_clicked2(self):

        self.conn.df.to_csv(self.codeName+".csv",mode='a', header=False)



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
        kospiList = ['A001630', 'A001680', 'A001685', 'A001720', 'A001725', 'A001740', 'A001745', 'A001750', 'A001755', 'A001770', 'A001780', 'A001790', 'A001795', 'A001799', 'A001800', 'A001820', 'A001880', 'A001940', 'A002020', 'A002025', 'A002030', 'A002070', 'A002100', 'A002140', 'A002150', 'A002170', 'A002200', 'A002210', 'A002220', 'A002240', 'A002270', 'A002300', 'A002310', 'A002320', 'A002350', 'A002355', 'A002360', 'A002820', 'A002840', 'A002870', 'A002880', 'A002900', 'A002920', 'A002960', 'A002990', 'A002995', 'Q500043', 'Q500044', 'Q500046', 'Q500047', 'Q500048', 'Q500049', 'Q550052', 'Q570034']


        # for code in kospiList :
        #     self.conn.GetRecentDataFromNumber2(code,200000,self.progressBar,"m")
        #     path = "C:\\Users\\Administrator\\PycharmProjects\\PracticeDesinApi\\data\\some\\"+code+ ".csv"
        #     print(path)
        #     # self.conn.df.to_csv(path, mode='a', header=False)
        #     self.conn.df.to_csv(path, mode='a')
        #
        #


        # kosdaqList = self.conn.GetMarketCode(2)
        #
        #
        #
        # for code in kosdaqList:
        #     self.conn.GetRecentDataFromNumber2(code, 200000, self.progressBar, "m")
        #     path = "C:\\Users\\Administrator\\PycharmProjects\\PracticeDesinApi\\data\\KOSDAQ\\" + code + ".csv"
        #     print(path)
        #     # self.conn.df.to_csv(path, mode='a', header=False)
        #     self.conn.df.to_csv(path, mode='a')
        #
        #

    def btn_clicked6(self):

        addNewFileOnOldFile(self.conn.df, self.codeName)


        # new_df = self.conn.df
        #
        # path = 'C:\\Users\\Administrator\\PycharmProjects\\PracticeDesinApi\\'+self.codeName+".csv"
        #
        # old_df = pd.read_csv(path,usecols=range(1,8),index_col=False)
        # old_first = old_df.iloc[1,][0]
        #
        # old_index = new_df[new_df['date'] == old_first].index.values[0]
        # new_df2 = new_df.iloc[:old_index,]
        # real_new_df = pd.concat([new_df2,old_df])
        # real_new_df.to_csv(path,mode='w', index=False)
#--------------------------








def addNewFileOnOldFile(newFile,codeName):
    # new_df = pd.read_csv("C:\\Users\\Administrator\\PycharmProjects\\PracticeDesinApi\\new.csv", usecols=range(1, 8),
    #                      index_col=False)
    # path = 'C:\\Users\\Administrator\\PycharmProjects\\PracticeDesinApi\\old.csv'
    #
    # old_df = pd.read_csv(path, usecols=range(1, 8), index_col=False)
    # old_first = old_df.iloc[0,][0]
    #
    #
    # old_index = new_df[new_df['date'] == old_first].index.values[0]
    # new_df2 = new_df.iloc[:old_index, ]
    #
    # real_new_df = pd.concat([new_df2, old_df])
    #
    #
    #
    # path = 'C:\\Users\\Administrator\\PycharmProjects\\PracticeDesinApi\\new4.csv'
    # real_new_df.to_csv(path, mode='w', index=False)


    new_df = newFile

    path = 'C:\\Users\\Administrator\\PycharmProjects\\PracticeDesinApi\\' + codeName + ".csv"

    old_df = pd.read_csv(path, usecols=range(1, 8), index_col=False)
    old_first = old_df.iloc[1,][0]

    old_index = new_df[new_df['date'] == old_first].index.values[0]
    new_df2 = new_df.iloc[:old_index, ]
    real_new_df = pd.concat([new_df2, old_df])
    real_new_df.to_csv(path, mode='w', index=False)





if __name__ == "__main__" :

    filepath = "C:\\Users\\Administrator\\PycharmProjects\\PracticeDesinApi\\data\\some"

    files = [f for f in listdir(filepath) if isfile(join(filepath,f))]
    files = [x[:7] for x in files]
    print(files)



    conn = DesinAPI.DesinAPI()

    # conn.SearchNameList("lg")
    # conn.GetStockCode('NAVER')

    li = conn.GetMarketCode(3)
    print(li)
    # conn.GetRecentDataFromNumber()




    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()



