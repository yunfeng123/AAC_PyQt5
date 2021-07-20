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
import run_pic
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
        self.t1 = Thread(self.text_CSV.text(), self.text_ini.text(), self.entry_config.text(), self.entry_qty.text())
        self.t1.show_signal.connect(self.send)
        self.t1.start()

    def send(self, print_info):
        self.text_info.append(print_info)


class Thread(QThread):
    # 自定义信号
    show_signal = pyqtSignal(str)

    # 构造函数，接受参数
    def __init__(self, filepath_csv, filepath_ini, config, qty):
        QThread.__init__(self)
        self.filepath_csv = filepath_csv
        self.filepath_ini = filepath_ini
        self.config = config
        self.qty = qty

    # 重写run()方法
    def run(self):
        dt = datetime.now()
        now_date = dt.strftime('%Y%m%d')
        time_t0 = time.time()
        print_info = ('Start -> Reading CSV File')
        self.show_signal.emit(print_info)
        data_csv, USL, LSL, Overlay, Station, project = csv_process.csv_process(self.filepath_csv)
        config_name = self.config
        data_ini = pd.read_excel(self.filepath_ini, sheet_name=Station)

        graph_flag = ''  # 画什么图
        graph_class = ''  # 图的分类标签

        boxplot_data_tem = pd.DataFrame()
        sweep_data_tem = pd.DataFrame()
        USL_tem = []
        LSL_tem = []
        boxplot_X_name = []

        # 创建保存图片的文件夹

        pic_dic = os.path.split(self.filepath_csv)[0] + '/' + 'Saved Photo' + '/'
        # pic_dic_0 = os.path.split(filepath_csv)[0] + '/' + 'Saved Photo'
        # isExists = os.path.exists(pic_dic)
        # if not isExists:
        #     os.makedirs(pic_dic)
        # main_frame.text_pic.setText(pic_dic_0)

        out_data = pd.DataFrame(index=['Item', 'USL', 'LSL', 'Mean', 'STDEV', 'CPK'])  # 输出数据表格

        ini_pic_order = 0
        for i in range(len(data_ini['Data Item'])):

            # 画图种类定义+标签
            if data_ini.loc[i, 'Plot Type'] == data_ini.loc[i, 'Plot Type']:
                graph_flag = data_ini.loc[i, 'Plot Type']
                graph_class = data_ini.loc[i, 'Plot Title']
                ini_pic_order = ini_pic_order + 1

            item_name = data_ini.loc[i, 'Data Item']

            find_flag = 0
            for item_ext in LSL.index:
                if item_ext.find(item_name) >= 0:
                    find_flag = 1
                    break

            if find_flag:

                if graph_flag == 'Histogram' and data_ini.loc[i, 'Monitor'] == data_ini.loc[
                    i, 'Monitor']:  # Monitor列判定是否画图

                    # USL 配置文件有，则用配置文件，无则用数据USL
                    if data_ini.loc[i, 'Upper Limit'] == data_ini.loc[i, 'Upper Limit']:
                        USL_histogram = data_ini.loc[i, 'Upper Limit']
                    else:
                        USL_histogram = USL[item_name]

                    # LSL 配置文件有，则用配置文件，无则用数据LSL
                    if data_ini.loc[i, 'Lower Limit'] == data_ini.loc[i, 'Lower Limit']:
                        LSL_histogram = data_ini.loc[i, 'Lower Limit']
                    else:
                        LSL_histogram = LSL[item_name]

                    print_info, out_his = Histogram.histogram(Station, pic_dic, config_name, graph_class, item_name,
                                                              USL_histogram, LSL_histogram, data_csv[item_name],
                                                              data_ini.loc[i, 'Axis Upper Limit'],
                                                              data_ini.loc[i, 'Axis Lower Limit'], ini_pic_order,
                                                              self.qty)
                    out_data = pd.concat([out_data, out_his], axis=1)
                    self.show_signal.emit(print_info)
                    # main_frame.text_info.append(print_info)
                    # main_frame.cursot = main_frame.text_info.textCursor()
                    # main_frame.text_info.moveCursor(main_frame.cursot.End)
                    # QApplication.processEvents()

                elif graph_flag == 'Boxplot' and data_ini.loc[i, 'Monitor'] == data_ini.loc[
                    i, 'Monitor']:  # Monitor列判定是否画图

                    boxplot_data_tem = pd.concat([boxplot_data_tem, data_csv[item_name]], axis=1)

                    if data_ini.loc[i, 'X-Axis Mark'] == data_ini.loc[i, 'X-Axis Mark']:
                        boxplot_X_name.append(data_ini.loc[i, 'X-Axis Mark'])

                    # USL 配置文件有，则用配置文件，无则用数据USL
                    if data_ini.loc[i, 'Upper Limit'] == data_ini.loc[i, 'Upper Limit']:
                        USL_tem.append(data_ini.loc[i, 'Upper Limit'])
                    else:
                        USL_tem.append(USL[item_name])

                    # LSL 配置文件有，则用配置文件，无则用数据LSL
                    if data_ini.loc[i, 'Lower Limit'] == data_ini.loc[i, 'Lower Limit']:
                        LSL_tem.append(data_ini.loc[i, 'Lower Limit'])
                    else:
                        LSL_tem.append(LSL[item_name])

                    if data_ini.loc[i + 1, 'Plot Type'] == data_ini.loc[i + 1, 'Plot Type']:
                        print_info, out_box = Boxplot.Boxplot(Station, pic_dic, config_name, graph_class, USL_tem,
                                                              LSL_tem,
                                                              boxplot_data_tem, data_ini.loc[i, 'Axis Upper Limit'],
                                                              data_ini.loc[i, 'Axis Lower Limit'], boxplot_X_name,
                                                              ini_pic_order, self.qty)
                        boxplot_X_name = []
                        out_data = pd.concat([out_data, out_box], axis=1)
                        self.show_signal.emit(print_info)
                        # main_frame.text_info.append(print_info)
                        # main_frame.cursot = main_frame.text_info.textCursor()
                        # main_frame.text_info.moveCursor(main_frame.cursot.End)
                        #QApplication.processEvents()
                        boxplot_data_tem = pd.DataFrame()
                        USL_tem = []
                        LSL_tem = []

                elif graph_flag == 'Sweep' and data_ini.loc[i, 'Monitor'] == data_ini.loc[
                    i, 'Monitor']:  # Monitor列判定是否画图
                    for j in data_csv.columns:
                        if j.find(item_name) >= 0:
                            sweep_data_tem = pd.concat([sweep_data_tem, data_csv[j]], axis=1)
                            USL_tem.append(USL[j])
                            LSL_tem.append(LSL[j])
                    print_info = Sweep.Sweep(Station, pic_dic, config_name, graph_class, USL_tem, LSL_tem,
                                             sweep_data_tem,
                                             data_ini.loc[i, 'Axis Upper Limit'], data_ini.loc[i, 'Axis Lower Limit'],
                                             ini_pic_order, self.qty)
                    self.show_signal.emit(print_info)
                    # main_frame.text_info.append(print_info)
                    # main_frame.cursot = main_frame.text_info.textCursor()
                    # main_frame.text_info.moveCursor(main_frame.cursot.End)
                    QApplication.processEvents()
                    sweep_data_tem = pd.DataFrame()
                    USL_tem = []
                    LSL_tem = []

            else:
                print_info = '<span style="background:red;">' + f'Error -> {item_name} IS not Found on Data !' + '</span>'
                self.show_signal.emit(print_info)

        # 保存数据到Excel Table
        save_table_name = project + '_' + Station + f'_Summary Table_{now_date}.xlsx'
        out_data_path = os.path.split(self.filepath_csv)[0] + '/' + save_table_name
        app = xw.App(visible=False, add_book=False)
        if os.path.exists(out_data_path):
            wb = app.books.open(out_data_path)
            for i in wb.sheets:
                if i.name == self.config:
                    wb.sheets[self.config].delete()
                    break
            if self.config == '':
                ws = wb.sheets.add()
            else:
                ws = wb.sheets.add(self.config)
        else:
            if self.config == '':
                wb = app.books.add()
                ws = wb.sheets[0]
            else:
                wb = app.books.add()
                ws = wb.sheets[0]
                ws.name = self.config

        ws.range('A1').expand('table').value = out_data
        ws.range('A1').api.EntireRow.Delete()
        wb.save(out_data_path)
        wb.close()
        app.quit()

        time_delta = time.time() - time_t0

        print_info = '<span style="background:green;">' + 'Finished All in ' + str(
            round(time_delta, 1)) + ' Seconds' + '</span>'
        self.show_signal.emit(print_info)
