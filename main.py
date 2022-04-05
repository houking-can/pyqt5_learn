import sys
import os
import time
import threading

from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QWidget
from PyQt5 import QtCore

import excel
from utils import get_html, fun_switch, close_excel_file, timer

COLORS = ["#8A2BE2","#A52A2A","#DEB887","#5F9EA0","#7FFF00",
          "#D2691E","#FF7F50","#6495ED","#FF00FF","#DC143C",
          "#00FFFF","#00008B","#008B8B","#B8860B","#A9A9A9",
          "#006400","#BDB76B","#8B008B","#556B2F","#FF8C00"
          ]

class Excel(QWidget):
    def __init__(self):
        super().__init__()
        self.MainWindow = QMainWindow()
        self.ui = excel.Ui_MainWindow()
        self.ui.setupUi(self.MainWindow)
        self.ui.actionOpen.triggered.connect(self.openFile)
        self.ui.OpenButton.clicked.connect(self.openRes)
        self.ui.ProcessButton.clicked.connect(self.process)
        self.ui.LikeButton.clicked.connect(self.like)
        self.ui.DonateButton.clicked.connect(self.donate)
        self.file = ""
        self.success = False
        self.savePath = './'
        if not os.path.exists(self.savePath):
            os.makedirs(self.savePath)

        self.ui.OpenButton.setEnabled(False)
        self.ui.ProcessButton.setEnabled(False)


    def openFile(self):
        self.ui.LogText.clear()
        openfile_name = QFileDialog.getOpenFileName(self, "宝贝儿，选择一个你想偷懒的表格", './')
        self.file = openfile_name[0]
        start_time = time.time()
        while (True):
            QtCore.QCoreApplication.processEvents()
            if self.file != '':
                timer(self)
                self.ui.InputText.setPlainText(self.file)
                self.ui.ProcessButton.setEnabled(True)
                filename = os.path.basename(self.file)
                name, extend_name = os.path.splitext(filename)
                fileRoot = os.path.dirname(self.file)
                self.savePath = f"{fileRoot}/{name}_res.xls"
                self.ui.LogText.appendPlainText(f"结果文件：{self.savePath}")
                break
            elif time.time() - start_time > 5:
                timer(self)
                self.ui.LogText.setPlainText("宝贝，你还没有选择文件！")
                return

    def monitor(self):
        start_time = time.time()
        begin_time = start_time
        while (True):
            QtCore.QCoreApplication.processEvents()
            if self.success==True:
                self.ui.LogText.appendPlainText(f"哈哈哈，已经完成了！")
                self.ui.OpenButton.setEnabled(True)
                self.success = False
                break
            if time.time()-start_time>=2:
                run_time = time.time()-begin_time
                self.ui.LogText.appendPlainText(f"已运行{run_time:.1f}秒...")
                start_time = time.time()
            if self.success=="error":
                self.success = False
                break

    def process(self):
        timer(self)
        self.fun_name = self.ui.FuncomboBox.currentText()
        if self.fun_name=="选择":
            self.ui.LogText.appendPlainText("Warning: 需要指定功能！")
            return
        self.ui.LogText.appendPlainText(f"选择的功能：{self.fun_name}，处理中...")
        t = threading.Thread(target=self.monitor)
        t.start()
        fun_switch(self)


    def openRes(self):
        close_excel_file(self.savePath)
        os.system(f"start {self.savePath}")

    def like(self):
        start = time.mktime(time.strptime('2020-08-13', '%Y-%m-%d'))
        now = time.time()
        days = int((now - start) // (24 * 3600))
        seconds = int((now - start) % (24 * 3600))
        hours = seconds//3600
        seconds = seconds % 3600
        minutes = seconds//60
        seconds = seconds%60
        color = COLORS[seconds%20]
        content=f"爱你的第{days}天:{hours}小时:{minutes}分钟:{seconds}秒"
        html = get_html(content,color)
        self.ui.LogText.appendHtml(html)

    def donate(self):
        now = int(time.time())
        color = COLORS[now%20]
        content=f"打工不易，求打赏！WeChatId: houkingyu"
        html = get_html(content,color)
        self.ui.LogText.appendHtml(html)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    Label = Excel()
    Label.MainWindow.show()
    sys.exit(app.exec_())
