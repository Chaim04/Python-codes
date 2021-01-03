import sys
from PyQt5 import QtWidgets
from TextReplacement_UI import Ui_Form
import os

class MyPyQT_Form(QtWidgets.QWidget,Ui_Form):
    def __init__(self):
        super(MyPyQT_Form,self).__init__()
        self.setupUi(self)
        self.PB_Replace.clicked.connect(self.Click)

    #实现pushButton_click()函数，textEdit是我们放上去的文本框的id
    def Click(self):

        path = self.LE_FilePath.text()  # 目标路径

        filename_list = os.listdir(path)  # 扫描目标路径的文件,将文件名存入列表
        Original_Text = self.LE_OriginalText.text()
        Replacement_Text = self.LE_ReplacementText.text()

        for i in range(0, len(filename_list)):
            used_name = path + "\\" + filename_list[i]
            new_name = path + "\\" + filename_list[i].replace(Original_Text, Replacement_Text)
            os.rename(used_name, new_name)
        self.L_Status.setText("Status:Completed")

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    my_pyqt_form = MyPyQT_Form()
    my_pyqt_form.show()
    sys.exit(app.exec_())