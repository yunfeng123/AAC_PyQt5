import main_window_rc as mw
from PyQt5.QtWidgets import *
import run_pic


class mainwindow_action(mw.Ui_MainWindow):
    def setupUi(self, MainWindow_Risk):
        mw.Ui_MainWindow.setupUi(self, MainWindow_Risk)
        self.btn_ini.clicked.connect(self.openfile_ini)
        self.btn_CSV.clicked.connect(self.openfile_csv)
        self.btn_pic.clicked.connect(self.opendir_pic)
        self.btn_run.clicked.connect(self.run)

    def openfile_ini(self):
        file_name, file_type = QFileDialog.getOpenFileName(caption='配置文件', filter="xlsx Files (*.xlsx);;All Files (*)")
        self.text_ini.setText(file_name)

    def openfile_csv(self):
        file_name, file_type = QFileDialog.getOpenFileName(caption='数据', filter="csv Files (*.csv);;All Files (*)")
        self.text_CSV.setText(file_name)

    def opendir_pic(self):
        file_path = QFileDialog.getExistingDirectory(caption='图片目录')
        self.text_pic.setText(file_path)

    def run(self):
        run_pic.run(self)
