import threading
import time
from datetime import datetime
import pandas as pd
import csv_process
import Histogram
import Boxplot
import Sweep_1 as Sweep
import os
import xlwings as xw

import main_window_rc as mw
from PyQt5.QtWidgets import *
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtCore import *
import thread_run_pic
import run_report


class mainwindow_action(mw.Ui_MainWindow):
    def setupUi(self, MainWindow_Risk):
        mw.Ui_MainWindow.setupUi(self, MainWindow_Risk)
        self.btn_ini.clicked.connect(self.openfile_ini)
        self.btn_CSV.clicked.connect(self.openfile_csv)
        self.btn_pic.clicked.connect(self.opendir_pic)
        self.btn_run.clicked.connect(self.logon)
        self.btn_report.clicked.connect(self.run_report)

    def openfile_ini(self):
        file_name, file_type = QFileDialog.getOpenFileName(caption='配置文件', filter="xlsx Files (*.xlsx);;All Files (*)")
        self.text_ini.setText(file_name)

    def openfile_csv(self):
        file_name, file_type = QFileDialog.getOpenFileName(caption='数据', filter="csv Files (*.csv);;All Files (*)")
        self.text_CSV.setText(file_name)

    def opendir_pic(self):
        file_path = QFileDialog.getExistingDirectory(caption='图片目录')
        self.text_pic.setText(file_path)

    # def run_pic(self):
    #     run_pic.run_pic(self)

    def run_report(self):
        run_report.run_report(self)

    def logon(self):
        self.t1 = thread_run_pic.Thread(self.text_CSV.text(), self.text_ini.text(), self.entry_config.text(), self.entry_qty.text())
        self.t1.show_signal.connect(self.send)
        self.t1.start()

    def send(self, print_info):
        self.text_info.append(print_info)



