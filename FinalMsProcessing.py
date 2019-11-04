import xlwt
import xlrd
import pymzml
from pyimzml.ImzMLParser import ImzMLParser
from pyimzml.ImzMLWriter import ImzMLWriter
from progressbar import *
import numpy as np
from PyQt5 import QtCore, QtGui, QtWidgets
import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib import pyplot as plt


class Error_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(400, 200)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/Mydata/panda.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(50, 30, 341, 51))
        self.label.setWordWrap(True)
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(10)
        self.label.setFont(font)
        self.label.setObjectName("label")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "运行错误"))
        self.label.setText(_translate("Form", "您的输入有误，请重新输入！"))

class My_Error_Form(QtWidgets.QWidget,Error_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setMinimumSize(400, 200)
        self.setMaximumSize(400, 200)

class Main_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(449, 295)
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(120, 20, 241, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(16)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.line = QtWidgets.QFrame(Form)
        self.line.setGeometry(QtCore.QRect(10, 50, 20, 231))
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.line_2 = QtWidgets.QFrame(Form)
        self.line_2.setGeometry(QtCore.QRect(420, 50, 20, 231))
        self.line_2.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.line_3 = QtWidgets.QFrame(Form)
        self.line_3.setGeometry(QtCore.QRect(20, 40, 81, 16))
        self.line_3.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.line_4 = QtWidgets.QFrame(Form)
        self.line_4.setGeometry(QtCore.QRect(340, 40, 91, 20))
        self.line_4.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_4.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_4.setObjectName("line_4")
        self.line_5 = QtWidgets.QFrame(Form)
        self.line_5.setGeometry(QtCore.QRect(20, 270, 411, 16))
        self.line_5.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_5.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_5.setObjectName("line_5")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(30, 60, 61, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(50, 110, 101, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        self.pushButton_2.setGeometry(QtCore.QRect(180, 110, 101, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.line_6 = QtWidgets.QFrame(Form)
        self.line_6.setGeometry(QtCore.QRect(30, 170, 391, 16))
        self.line_6.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_6.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_6.setObjectName("line_6")
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(30, 200, 61, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.pushButton_3 = QtWidgets.QPushButton(Form)
        self.pushButton_3.setGeometry(QtCore.QRect(110, 200, 101, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_4 = QtWidgets.QPushButton(Form)
        self.pushButton_4.setGeometry(QtCore.QRect(310, 110, 101, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.pushButton_4.setFont(font)
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_5 = QtWidgets.QPushButton(Form)
        self.pushButton_5.setGeometry(QtCore.QRect(250, 200, 101, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.pushButton_5.setFont(font)
        self.pushButton_5.setObjectName("pushButton_5")

        self.retranslateUi(Form)
        self.pushButton.clicked.connect(Form.Imzml_Biomap)
        self.pushButton_2.clicked.connect(Form.Imzml_xls)
        self.pushButton_3.clicked.connect(Form.Mzml_Translate)
        self.pushButton_4.clicked.connect(Form.IMzml_Image)
        self.pushButton_5.clicked.connect(Form.Mzml_Image)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "综合质谱数据处理程序"))
        self.label.setText(_translate("Form", "综合质谱数据处理程序"))
        self.label_2.setText(_translate("Form", "Imzml"))
        self.pushButton.setText(_translate("Form", "Biomap处理"))
        self.pushButton_2.setText(_translate("Form", "导出成xls"))
        self.label_3.setText(_translate("Form", "Mzml"))
        self.pushButton_3.setText(_translate("Form", "导出到xls"))
        self.pushButton_4.setText(_translate("Form", "绘制图像"))
        self.pushButton_5.setText(_translate("Form", "绘制图像"))

class MyPyQT_Main_Form(QtWidgets.QWidget,Main_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setMinimumSize(449,295)
        self.setMaximumSize(449,295)

    def Imzml_Biomap(self):
        self.NewTranslate2 = My_Imzml_1_Form()
        self.NewTranslate2.show()

    def Imzml_xls(self):
        self.NewTranslate3 = My_Imzml_2_Form()
        self.NewTranslate3.show()

    def Mzml_Translate(self):
        self.NewTranslate = Mzml_1_Form()
        self.NewTranslate.show()

    def IMzml_Image(self):
        self.NewImzmlImage = My_Imzml_3_Form()
        self.NewImzmlImage.show()

    def Mzml_Image(self):
        self.NewImzmlImage2 = Mzml_2_Form()
        self.NewImzmlImage2.show()

class Imzml_1_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(613, 465)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("Mydata/panda.ico"))
        Form.setWindowIcon(icon)
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(20, 30, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(20, 90, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.line = QtWidgets.QFrame(Form)
        self.line.setGeometry(QtCore.QRect(20, 140, 571, 20))
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setLineWidth(1)
        self.line.setMidLineWidth(0)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setObjectName("line")
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(40, 170, 91, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.line_2 = QtWidgets.QFrame(Form)
        self.line_2.setGeometry(QtCore.QRect(280, 160, 51, 291))
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setLineWidth(1)
        self.line_2.setMidLineWidth(0)
        self.line_2.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_2.setObjectName("line_2")
        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setGeometry(QtCore.QRect(40, 220, 141, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(Form)
        self.label_5.setGeometry(QtCore.QRect(70, 260, 51, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(Form)
        self.label_6.setGeometry(QtCore.QRect(70, 310, 51, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(80, 410, 101, 41))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.label_7 = QtWidgets.QLabel(Form)
        self.label_7.setGeometry(QtCore.QRect(340, 170, 91, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.label_8 = QtWidgets.QLabel(Form)
        self.label_8.setGeometry(QtCore.QRect(340, 220, 141, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_8.setFont(font)
        self.label_8.setObjectName("label_8")
        self.label_9 = QtWidgets.QLabel(Form)
        self.label_9.setGeometry(QtCore.QRect(370, 260, 51, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_9.setFont(font)
        self.label_9.setObjectName("label_9")
        self.label_10 = QtWidgets.QLabel(Form)
        self.label_10.setGeometry(QtCore.QRect(370, 310, 51, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_10.setFont(font)
        self.label_10.setObjectName("label_10")
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        self.pushButton_2.setGeometry(QtCore.QRect(400, 410, 101, 41))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.progressBar = QtWidgets.QProgressBar(Form)
        self.progressBar.setGeometry(QtCore.QRect(30, 370, 251, 23))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.progressBar.setFont(font)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.progressBar_2 = QtWidgets.QProgressBar(Form)
        self.progressBar_2.setGeometry(QtCore.QRect(340, 370, 251, 23))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.progressBar_2.setFont(font)
        self.progressBar_2.setProperty("value", 0)
        self.progressBar_2.setObjectName("progressBar_2")
        self.lineEdit = QtWidgets.QLineEdit(Form)
        self.lineEdit.setGeometry(QtCore.QRect(200, 40, 391, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit.setFont(font)
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(Form)
        self.lineEdit_2.setGeometry(QtCore.QRect(220, 100, 371, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.lineEdit_3 = QtWidgets.QLineEdit(Form)
        self.lineEdit_3.setGeometry(QtCore.QRect(130, 180, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_3.setFont(font)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.lineEdit_4 = QtWidgets.QLineEdit(Form)
        self.lineEdit_4.setGeometry(QtCore.QRect(130, 270, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_4.setFont(font)
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.lineEdit_5 = QtWidgets.QLineEdit(Form)
        self.lineEdit_5.setGeometry(QtCore.QRect(130, 320, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_5.setFont(font)
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.lineEdit_6 = QtWidgets.QLineEdit(Form)
        self.lineEdit_6.setGeometry(QtCore.QRect(440, 180, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_6.setFont(font)
        self.lineEdit_6.setObjectName("lineEdit_6")
        self.lineEdit_7 = QtWidgets.QLineEdit(Form)
        self.lineEdit_7.setGeometry(QtCore.QRect(440, 270, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_7.setFont(font)
        self.lineEdit_7.setObjectName("lineEdit_7")
        self.lineEdit_8 = QtWidgets.QLineEdit(Form)
        self.lineEdit_8.setGeometry(QtCore.QRect(440, 320, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_8.setFont(font)
        self.lineEdit_8.setObjectName("lineEdit_8")

        self.retranslateUi(Form)
        self.pushButton.clicked.connect(Form.TargetMass_Extract)
        self.pushButton_2.clicked.connect(Form.InnerStandard_Solve)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "imzml数据提取到Biomap窗口"))
        self.label.setText(_translate("Form", "待处理的imzml文件路径:"))
        self.label_2.setText(_translate("Form", "处理后imzml文件输出路径:"))
        self.label_3.setText(_translate("Form", "目标质荷比： "))
        self.label_4.setText(_translate("Form", "目标质荷比范围： "))
        self.label_5.setText(_translate("Form", "起始： "))
        self.label_6.setText(_translate("Form", "结束： "))
        self.pushButton.setText(_translate("Form", "开始转换"))
        self.label_7.setText(_translate("Form", "内标质荷比： "))
        self.label_8.setText(_translate("Form", "内标质荷比范围： "))
        self.label_9.setText(_translate("Form", "起始： "))
        self.label_10.setText(_translate("Form", "结束： "))
        self.pushButton_2.setText(_translate("Form", "开始转换"))

def PeakIntensitySum(spec,diffl,diffr):
    s=spec[0][np.where(spec[0]<=diffr)]
    ss=spec[1][np.where(s>=diffl)]
    return(ss.sum())

def MaxIntensitySeeker(p,diffl,diffr):
    Coor = p.coordinates
    total = len(Coor)
    Max=0
    for indecount in range(0, total):
        m = p.getspectrum(indecount)
        Tmep = PeakIntensitySum(m, diffl, diffr)
        if Tmep>Max:
            Max = Tmep
    return Max

class My_Imzml_1_Form(QtWidgets.QWidget,Imzml_1_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

    def TargetMass_Extract(self):
        try:
            self.mbt = Imzml_1_thread_1(self.lineEdit.text(), self.lineEdit_2.text(), eval(self.lineEdit_3.text()), self.lineEdit_4.text(),self.lineEdit_5.text())
            self.mbt.trigger.connect(self.slot_thread)
            self.mbt.trigger3.connect(self.error)
            self.mbt.start()
        except Exception as e:
            m='运行错误，错误信息：'+str(e)
            self.error(m)

    def error(self,m):
        self.eW=My_Error_Form()
        self.eW.label.setText(m)
        self.eW.show()

    def slot_thread(self, msg_1):
        self.progressBar.setValue(msg_1)

    def InnerStandard_Solve(self):
        try:
            self.mbt2 = Imzml_1_thread_2(self.lineEdit.text(), self.lineEdit_2.text(), self.lineEdit_6.text(),self.lineEdit_7.text(),self.lineEdit_8.text())
            self.mbt2.trigger.connect(self.slot_thread2)
            self.mbt2.trigger3.connect(self.error)
            self.mbt2.start()
        except Exception as e :
            m='运行错误，错误信息：'+str(e)
            self.error(m)

    def slot_thread2(self, msg_2):
        self.progressBar_2.setValue(msg_2)

class Imzml_1_thread_1(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int)
    trigger3 = QtCore.pyqtSignal(str)

    def __init__(self,a,b,c,d,e):
        super().__init__()
        self.mzmlIntPath = a
        self.mzmlOutPath = b
        self.TargetMass = c
        self.left = d.split(',')
        self.right = e.split(',')

    def run(self):
        try:
            self.trigger.emit(20)
            p = ImzMLParser(self.mzmlIntPath)
            out = ImzMLWriter(self.mzmlOutPath)
            QtWidgets.QApplication.processEvents()
            Max=[]
            Coor = p.coordinates
            total = len(Coor)
            for i in range(0,len(self.left)):
                v1 = int((i / (len(self.left) - 1)) * 50)
                if v1>20:
                    self.trigger.emit(v1)
                Max.append(MaxIntensitySeeker(p,eval(self.left[i]),eval(self.right[i])))

            for indecount in range(0, total):
                v1 = int((indecount / (total - 1)) * 50)+50
                self.trigger.emit(v1)
                m = p.getspectrum(indecount)
                inte = m[1]
                inte = inte[np.where(inte!=0)[0]]
                I = 0
                for l in range(0,len(self.left)):
                    Tmep = PeakIntensitySum(m, eval(self.left[l]), eval(self.right[l]))
                    Tmep = Tmep*100 / Max[l]
                    I+=Tmep
                if len(inte)!=0 :
                    inte=sorted(inte)
                    key = inte[int(len(inte)/2)]
                    I = I / key
                mzs = np.array([self.TargetMass])
                intensity = np.array([I])
                out.addSpectrum(mzs, intensity, Coor[indecount])
            out.close()
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.trigger.emit(0)
            self.trigger3.emit(m)

class Imzml_1_thread_2(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int)
    trigger3 = QtCore.pyqtSignal(str)

    def __init__(self,a,b,c,d,e):
        super().__init__()
        self.ImzmlIntPath = a
        self.ImzmlOutPath = b
        self.TargetIS = c
        self.left = d.split(',')
        self.right = e.split(',')

    def run(self):
        try:
            ImzmlOutPath2 = self.ImzmlOutPath + 'after' + 'standardization'+ self.TargetIS
            self.trigger.emit(5)
            p = ImzMLParser(self.ImzmlIntPath)
            p2 = ImzMLParser(self.ImzmlOutPath + '.imzml')
            out2 = ImzMLWriter(ImzmlOutPath2)

            Coor = p2.coordinates
            total = len(Coor)
            for indecount in range(0, total):
                v1 = int((indecount / (total - 1)) * 100)
                if v1 > 5:
                    self.trigger.emit(v1)
                m = p2.getspectrum(indecount)
                N = p.getspectrum(indecount)
                I=0
                for l in range(0,len(self.left)):
                    Tmep = PeakIntensitySum(N, eval(self.left[l]), eval(self.right[l]))
                    I+=Tmep
                if int(I) == 0:
                    intensity = m[1] * 100
                else:
                    intensity = m[1] * 100 / I
                out2.addSpectrum(m[0], intensity, Coor[indecount])
            out2.close()
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.trigger.emit(0)
            self.trigger3.emit(m)

class Imzml_2_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(608, 399)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/Mydata/panda.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(20, 30, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(20, 90, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.line = QtWidgets.QFrame(Form)
        self.line.setGeometry(QtCore.QRect(20, 140, 571, 20))
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setLineWidth(1)
        self.line.setMidLineWidth(0)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setObjectName("line")
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(40, 170, 91, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.line_2 = QtWidgets.QFrame(Form)
        self.line_2.setGeometry(QtCore.QRect(290, 160, 51, 231))
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setLineWidth(1)
        self.line_2.setMidLineWidth(0)
        self.line_2.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_2.setObjectName("line_2")
        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setGeometry(QtCore.QRect(40, 220, 141, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(Form)
        self.label_5.setGeometry(QtCore.QRect(70, 260, 51, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(Form)
        self.label_6.setGeometry(QtCore.QRect(70, 310, 51, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(470, 260, 101, 41))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        self.pushButton_2.setGeometry(QtCore.QRect(470, 340, 101, 41))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.progressBar = QtWidgets.QProgressBar(Form)
        self.progressBar.setGeometry(QtCore.QRect(350, 180, 251, 23))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.progressBar.setFont(font)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.lineEdit = QtWidgets.QLineEdit(Form)
        self.lineEdit.setGeometry(QtCore.QRect(200, 40, 391, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit.setFont(font)
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(Form)
        self.lineEdit_2.setGeometry(QtCore.QRect(220, 100, 371, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.lineEdit_3 = QtWidgets.QLineEdit(Form)
        self.lineEdit_3.setGeometry(QtCore.QRect(130, 180, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_3.setFont(font)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.lineEdit_4 = QtWidgets.QLineEdit(Form)
        self.lineEdit_4.setGeometry(QtCore.QRect(130, 270, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_4.setFont(font)
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.lineEdit_5 = QtWidgets.QLineEdit(Form)
        self.lineEdit_5.setGeometry(QtCore.QRect(130, 320, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_5.setFont(font)
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.label_7 = QtWidgets.QLabel(Form)
        self.label_7.setGeometry(QtCore.QRect(350, 220, 121, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.label_8 = QtWidgets.QLabel(Form)
        self.label_8.setGeometry(QtCore.QRect(350, 300, 121, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_8.setFont(font)
        self.label_8.setObjectName("label_8")

        self.retranslateUi(Form)
        self.pushButton.clicked.connect(Form.Grid)
        self.pushButton_2.clicked.connect(Form.Coordinates)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "imzml数据导出窗口"))
        self.label.setText(_translate("Form", "待处理的imzml文件路径:"))
        self.label_2.setText(_translate("Form", "处理后xls文件输出路径:"))
        self.label_3.setText(_translate("Form", "目标质荷比： "))
        self.label_4.setText(_translate("Form", "目标质荷比范围： "))
        self.label_5.setText(_translate("Form", "起始： "))
        self.label_6.setText(_translate("Form", "结束： "))
        self.pushButton.setText(_translate("Form", "开始转换"))
        self.pushButton_2.setText(_translate("Form", "开始转换"))
        self.label_7.setText(_translate("Form", "网格形式转换："))
        self.label_8.setText(_translate("Form", "坐标形式转换："))

class My_Imzml_2_Form(QtWidgets.QWidget,Imzml_2_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

    def Grid(self):
        try:
            self.mbt = Imzml_2_thread_1(self.lineEdit.text(), self.lineEdit_2.text(), eval(self.lineEdit_4.text()),eval(self.lineEdit_5.text()))
            self.mbt.trigger.connect(self.slot_thread)
            self.mbt.trigger3.connect(self.error)
            self.mbt.start()
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.error(m)

    def error(self,m):
        self.eW=My_Error_Form()
        self.eW.label.setText(m)
        self.eW.show()

    def slot_thread(self, msg_1):
        self.progressBar.setValue(msg_1)

    def Coordinates(self):
        try:
            self.mbt2 = Imzml_2_thread_2(self.lineEdit.text(), self.lineEdit_2.text(), self.lineEdit_3.text(),eval(self.lineEdit_4.text()), eval(self.lineEdit_5.text()))
            self.mbt2.trigger.connect(self.slot_thread)
            self.mbt2.trigger3.connect(self.error)
            self.mbt2.start()
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.error(m)

class Imzml_2_thread_1(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int)
    trigger3 = QtCore.pyqtSignal(str)

    def __init__(self,a,b,c,d):
        super().__init__()
        self.imzmlIntPath = a
        self.imzmlOutPath = b
        self.left = c
        self.right = d

    def run(self):
        try:
            self.trigger.emit(5)
            p = ImzMLParser(self.imzmlIntPath)
            book = xlwt.Workbook(encoding='utf-8', style_compression=0)
            sheet = book.add_sheet('test', cell_overwrite_ok=True)

            Coor = p.coordinates
            total = len(Coor)
            for indecount in range(0, total):
                v1 = int((indecount / (total - 1)) * 100)
                if v1 > 5:
                    self.trigger.emit(v1)
                m = p.getspectrum(indecount)
                x = Coor[indecount][0]
                y = Coor[indecount][1]
                I = PeakIntensitySum(m, self.left, self.right)
                sheet.write(y, x, float(I))
            book.save(self.imzmlOutPath)
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.trigger.emit(0)
            self.trigger3.emit(m)

class Imzml_2_thread_2(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int)
    trigger3 = QtCore.pyqtSignal(str)

    def __init__(self,a,b,c,d,e):
        super().__init__()
        self.ImzmlIntPath = a
        self.ImzmlOutPath = b
        self.TargetIS = c
        self.left = d
        self.right = e

    def run(self):
        try:
            self.trigger.emit(5)
            p = ImzMLParser(self.ImzmlIntPath)
            book = xlwt.Workbook(encoding='utf-8', style_compression=0)
            sheet = book.add_sheet('test', cell_overwrite_ok=True)

            Coor = p.coordinates
            num = 1
            total = len(Coor)
            sheet.write(num, 1, 'x')
            sheet.write(num, 2, 'y')
            sheet.write(num, 3, 'z')
            for indecount in range(0, total):
                num += 1
                v1 = int((indecount / (total - 1)) * 100)
                if v1 > 5:
                    self.trigger.emit(v1)
                m = p.getspectrum(indecount)
                x = Coor[indecount][0]
                y = Coor[indecount][1]
                I = PeakIntensitySum(m, self.left, self.right)
                sheet.write(num, 1, x)
                sheet.write(num, 2, y)
                sheet.write(num, 3, float(I))
            book.save(self.ImzmlOutPath)
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.trigger.emit(0)
            self.trigger3.emit(m)

class Imzml_3_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(613, 284)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/Mydata/panda.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(20, 30, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.line = QtWidgets.QFrame(Form)
        self.line.setGeometry(QtCore.QRect(20, 90, 571, 20))
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setLineWidth(1)
        self.line.setMidLineWidth(0)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setObjectName("line")
        self.line_2 = QtWidgets.QFrame(Form)
        self.line_2.setGeometry(QtCore.QRect(280, 110, 51, 161))
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setLineWidth(1)
        self.line_2.setMidLineWidth(0)
        self.line_2.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_2.setObjectName("line_2")
        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setGeometry(QtCore.QRect(20, 100, 141, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(Form)
        self.label_5.setGeometry(QtCore.QRect(70, 140, 51, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(Form)
        self.label_6.setGeometry(QtCore.QRect(70, 200, 51, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(340, 200, 101, 41))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.progressBar = QtWidgets.QProgressBar(Form)
        self.progressBar.setGeometry(QtCore.QRect(350, 150, 251, 23))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.progressBar.setFont(font)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.lineEdit = QtWidgets.QLineEdit(Form)
        self.lineEdit.setGeometry(QtCore.QRect(200, 40, 391, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit.setFont(font)
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_4 = QtWidgets.QLineEdit(Form)
        self.lineEdit_4.setGeometry(QtCore.QRect(100, 150, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_4.setFont(font)
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.lineEdit_5 = QtWidgets.QLineEdit(Form)
        self.lineEdit_5.setGeometry(QtCore.QRect(100, 210, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_5.setFont(font)
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        self.pushButton_2.setGeometry(QtCore.QRect(470, 200, 101, 41))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")

        self.retranslateUi(Form)
        self.pushButton.clicked.connect(Form.Image)
        self.pushButton_2.clicked.connect(Form.ExtractToXls)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Imzml绘图窗口"))
        self.label.setText(_translate("Form", "待处理的Imzml文件路径:"))
        self.label_4.setText(_translate("Form", "坐标："))
        self.label_5.setText(_translate("Form", "X： "))
        self.label_6.setText(_translate("Form", "Y： "))
        self.pushButton.setText(_translate("Form", "开始绘制"))
        self.pushButton_2.setText(_translate("Form", "转存到xls"))

class My_Imzml_3_Form(QtWidgets.QWidget,Imzml_3_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

    def Image(self):
        try:
            self.mbt = Imzml_3_thread_1(self.lineEdit.text(),eval(self.lineEdit_4.text()),eval(self.lineEdit_5.text()))
            self.mbt.trigger.connect(self.slot_thread)
            self.mbt.trigger2.connect(self.slot_thread_2)
            self.mbt.trigger3.connect(self.error)
            self.mbt.start()
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.error(m)

    def error(self,m):
        self.eW=My_Error_Form()
        self.eW.label.setText(m)
        self.eW.show()

    def slot_thread(self, msg_1):
        self.progressBar.setValue(msg_1)

    def slot_thread_2(self,x,y):
        try:
            plt.figure()
            plt.plot(x, y, color='k')
            plt.show()
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.error(m)

    def ExtractToXls(self):
        try:
            self.mbt2 = Imzml_3_thread_2(self.lineEdit.text(),eval(self.lineEdit_4.text()),eval(self.lineEdit_5.text()))
            self.mbt2.trigger.connect(self.slot_thread)
            self.mbt2.trigger3.connect(self.error)
            self.mbt2.start()
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.error(m)

class Imzml_3_thread_1(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int)
    trigger2 = QtCore.pyqtSignal(np.ndarray,np.ndarray)
    trigger3 = QtCore.pyqtSignal(str)

    def __init__(self,a,b,c):
        super().__init__()
        self.ImzmlIntPath = a
        self.x = b
        self.y = c

    def run(self):
        try:
            self.trigger.emit(25)
            p = ImzMLParser(self.ImzmlIntPath)
            self.trigger.emit(50)
            Coor = p.coordinates
            zuobiao = (self.x,self.y,1)
            for i in range(0,len(Coor)):
                if Coor[i]==zuobiao:
                    zuobiao=i
                    break
            m = p.getspectrum(zuobiao)
            x1 = m[0][np.where(m[1] != 0)[0]]
            y1 = m[1][np.where(m[1] != 0)[0]]
            self.trigger.emit(75)
            self.trigger2.emit(x1,y1)
            self.trigger.emit(100)
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.trigger.emit(0)
            self.trigger3.emit(m)

class Imzml_3_thread_2(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int)
    trigger3 = QtCore.pyqtSignal(str)

    def __init__(self,a,b,c):
        super().__init__()
        self.ImzmlIntPath = a
        self.x = b
        self.y = c

    def run(self):
        try:
            self.trigger.emit(5)
            p = ImzMLParser(self.ImzmlIntPath)
            book = xlwt.Workbook(encoding='utf-8', style_compression=0)
            sheet = book.add_sheet('test', cell_overwrite_ok=True)
            Coor = p.coordinates
            zuobiao = (self.x,self.y,1)
            for i in range(0,len(Coor)):
                if Coor[i]==zuobiao:
                    zuobiao=i
                    break
            m = p.getspectrum(zuobiao)
            x1 = m[0][np.where(m[1] != 0)[0]]
            y1 = m[1][np.where(m[1] != 0)[0]]
            for inte in range(0,len(x1)):
                g = len(x1)
                info = int((inte/g)*100)
                if info > 5:
                    self.trigger.emit(info)
                sheet.write(inte,0,float(x1[inte]))
                sheet.write(inte,1,float(y1[inte]))
            path = self.ImzmlIntPath[:-6]+str(self.x)+','+str(self.y)+'.xls'
            book.save(path)
            self.trigger.emit(100)
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.trigger.emit(0)
            self.trigger3.emit(m)

class Mzml_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(613, 364)
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(20, 30, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(20, 90, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.line = QtWidgets.QFrame(Form)
        self.line.setGeometry(QtCore.QRect(20, 140, 571, 20))
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setLineWidth(1)
        self.line.setMidLineWidth(0)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setObjectName("line")
        self.line_2 = QtWidgets.QFrame(Form)
        self.line_2.setGeometry(QtCore.QRect(280, 160, 51, 191))
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setLineWidth(1)
        self.line_2.setMidLineWidth(0)
        self.line_2.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_2.setObjectName("line_2")
        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setGeometry(QtCore.QRect(20, 170, 141, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(Form)
        self.label_5.setGeometry(QtCore.QRect(70, 220, 51, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(Form)
        self.label_6.setGeometry(QtCore.QRect(70, 270, 51, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(400, 260, 101, 41))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.progressBar = QtWidgets.QProgressBar(Form)
        self.progressBar.setGeometry(QtCore.QRect(350, 200, 251, 23))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.progressBar.setFont(font)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.lineEdit = QtWidgets.QLineEdit(Form)
        self.lineEdit.setGeometry(QtCore.QRect(200, 40, 391, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit.setFont(font)
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(Form)
        self.lineEdit_2.setGeometry(QtCore.QRect(220, 100, 371, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.lineEdit_4 = QtWidgets.QLineEdit(Form)
        self.lineEdit_4.setGeometry(QtCore.QRect(130, 230, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_4.setFont(font)
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.lineEdit_5 = QtWidgets.QLineEdit(Form)
        self.lineEdit_5.setGeometry(QtCore.QRect(130, 280, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit_5.setFont(font)
        self.lineEdit_5.setObjectName("lineEdit_5")

        self.retranslateUi(Form)
        self.pushButton.clicked.connect(Form.TargetMass_Extract)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "mzml数据提取窗口"))
        self.label.setText(_translate("Form", "待处理的mzml文件路径:"))
        self.label_2.setText(_translate("Form", "处理后mzml文件输出路径:"))
        self.label_4.setText(_translate("Form", "ID号范围： "))
        self.label_5.setText(_translate("Form", "起始： "))
        self.label_6.setText(_translate("Form", "结束： "))
        self.pushButton.setText(_translate("Form", "开始转换"))

class Mzml_1_Form(QtWidgets.QWidget,Mzml_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

    def TargetMass_Extract(self):
        try:
            self.mbt = Mzml_1_1_Thread(self.lineEdit.text(),self.lineEdit_2.text(),eval(self.lineEdit_4.text()),eval(self.lineEdit_5.text()))
            self.mbt.trigger.connect(self.slot_thread)
            self.mbt.trigger3.connect(self.error)
            self.mbt.start()
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.error(m)

    def error(self,m):
        self.eW=My_Error_Form()
        self.eW.label.setText(m)
        self.eW.show()

    def slot_thread(self, msg_1):
        self.progressBar.setValue(msg_1)

class Mzml_1_1_Thread(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int)
    trigger3 = QtCore.pyqtSignal(str)

    def __init__(self,a,b,c,d):
        super().__init__()
        self.mzmlIntPath = a
        self.mzmlOutPath = b
        self.left = c
        self.right = d

    def run(self):
        try:
            raw_data = pymzml.run.Reader(self.mzmlIntPath)
            index = 1
            sheet = []
            book = xlwt.Workbook(encoding='utf-8', style_compression=0)
            sheet.append(book.add_sheet(str(index), cell_overwrite_ok=True))
            c = -1
            sheet[0].write(0,0,'kkk')
            self.trigger.emit(0)
            for spec in raw_data:
                if spec.ID >= self.left and spec.ID <= self.right:
                    info = int(((spec.ID - self.left) / (self.right - self.left)) * 100)
                    self.trigger.emit(info)
                    if c > 200:
                        index += 1
                        sheet.append(book.add_sheet(str(index), cell_overwrite_ok=True))
                        sheet[index - 1].write(0, 0, 'kkk')
                        c = -1
                    c += 2
                    smass = 'Mass' + str(spec.ID)
                    sintensity = 'Intensity' + str(spec.ID)
                    sheet[index - 1].write(0, c, smass)
                    sheet[index - 1].write(0, (c + 1), sintensity)

                    for i in range(0, len(spec.mz)):
                        sheet[index - 1].write(i + 1, c, spec.mz[i])
                        sheet[index - 1].write(i + 1, (c + 1), spec.i[i])
                if spec.ID > self.right:
                    break
            book.save(self.mzmlOutPath)
            self.trigger.emit(100)
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.trigger.emit(0)
            self.trigger3.emit(m)

class Mzml_Form_2(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(613, 244)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/newPrefix/Mydata/panda.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(20, 30, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(10)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.line = QtWidgets.QFrame(Form)
        self.line.setGeometry(QtCore.QRect(20, 90, 571, 20))
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setLineWidth(1)
        self.line.setMidLineWidth(0)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setObjectName("line")
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(230, 170, 131, 51))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.progressBar = QtWidgets.QProgressBar(Form)
        self.progressBar.setGeometry(QtCore.QRect(40, 120, 561, 23))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.progressBar.setFont(font)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.lineEdit = QtWidgets.QLineEdit(Form)
        self.lineEdit.setGeometry(QtCore.QRect(200, 40, 391, 31))
        font = QtGui.QFont()
        font.setFamily("Adobe 黑体 Std R")
        font.setPointSize(12)
        self.lineEdit.setFont(font)
        self.lineEdit.setObjectName("lineEdit")

        self.retranslateUi(Form)
        self.pushButton.clicked.connect(Form.Image)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "mzml绘图窗口"))
        self.label.setText(_translate("Form", "待处理的xls or xlsx文件路径:"))
        self.pushButton.setText(_translate("Form", "开始绘制"))

class Mzml_2_Form(QtWidgets.QWidget,Mzml_Form_2):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

    def Image(self):
        try:
            self.mbt = Mzml_2_thread_1(self.lineEdit.text())
            self.mbt.trigger.connect(self.slot_thread)
            self.mbt.trigger2.connect(self.slot_thread_2)
            self.mbt.trigger3.connect(self.error)
            self.mbt.start()
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.error(m)

    def error(self,m):
        self.eW=My_Error_Form()
        self.eW.label.setText(m)
        self.eW.show()

    def slot_thread(self, msg_1):
        self.progressBar.setValue(msg_1)

    def slot_thread_2(self,x,y):
        try:
            plt.figure()
            plt.bar(x,y,width=1.5)
            plt.show()
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.error(m)

class Mzml_2_thread_1(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int)
    trigger2 = QtCore.pyqtSignal(np.ndarray,np.ndarray)
    trigger3 = QtCore.pyqtSignal(str)

    def __init__(self,a):
        super().__init__()
        self.xlsIntPath = a

    def run(self):
        try:
            self.trigger.emit(25)
            data = xlrd.open_workbook(self.xlsIntPath)
            self.trigger.emit(50)
            table1 = data.sheets()[0]
            x1=np.array(table1.col_values(0))
            y1=np.array(table1.col_values(1))
            self.trigger.emit(75)
            self.trigger2.emit(x1,y1)
            self.trigger.emit(100)
        except Exception as e:
            m = '运行错误，错误信息：' + str(e)
            self.trigger.emit(0)
            self.trigger3.emit(m)



if __name__ == '__main__':
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QtWidgets.QApplication(sys.argv)
    my_pyqt_form = MyPyQT_Main_Form()
    my_pyqt_form.show()
    sys.exit(app.exec_())
