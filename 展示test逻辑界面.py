import sys
from PyQt5.QtWidgets import QMainWindow, QApplication
from test import Ui_MainWindow
class mywindow(QMainWindow,Ui_MainWindow):#继承QMainWindow和界面UI
    def  __init__(self,parent=None):#类的初始化函数
        super(mywindow,self).__init__(parent)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.slot1)
    def slot1(self):
        print("点击确定")
if __name__ == "__main__":

    app = QApplication(sys.argv)
    mywin = mywindow()
    mywin.show()
    sys.exit(app.exec_())

