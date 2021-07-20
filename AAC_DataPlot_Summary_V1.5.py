import tkinter.messagebox
from datetime import datetime
import main_window_action as mw
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow

dt = datetime.now()
now_date = dt.strftime('%Y%m%d')
overdue_date = '20230704'

if now_date > overdue_date:
    tkinter.messagebox.askokcancel(title='Error', message='License Expired !')
    exit()

if __name__ == '__main__':
    # 创建QApplication类的实例
    app = QApplication(sys.argv)
    # 创建一个主窗口
    mainWindow = QMainWindow()
    # 创建Ui_MainWindow的实例
    ui = mw.mainwindow_action()
    # 调用setupUi在指定窗口(主窗口)中添加控件
    ui.setupUi(mainWindow)
    # 显示窗口
    mainWindow.show()
    # 进入程序的主循环，并通过exit函数确保主循环安全结束
    sys.exit(app.exec_())
