import pandas as pd

import xlwings as xw
import csv_process
import Histogram
import Boxplot
import Sweep_1 as Sweep
import time
import os
import datetime


def run(self):
    dt = datetime.datetime.now()
    now_date = dt.strftime('%Y%m%d')
    overdue_date = '20230704'

    time_t0 = time.time()
    text_info_line = 2  # 监控text_info已经输入到第几行
    self.text_info.clear()
    self.text_info.append('Start -> Reading CSV File' + '\n')
    self.text_info.update()
    data_csv, USL, LSL, Overlay, Station, project = csv_process.csv_process(self.text_CSV.text())

    config_name = self.entry_config.text()
    data_ini = pd.read_excel(self.text_ini.text(), sheet_name=Station)

    graph_flag = ''  # 画什么图
    graph_class = ''  # 图的分类标签

    boxplot_data_tem = pd.DataFrame()
    sweep_data_tem = pd.DataFrame()
    USL_tem = []
    LSL_tem = []
    boxplot_X_name = []

    print_info = ''  # 显示打印信息

    # 创建保存图片的文件夹
    filepath_csv = self.text_CSV.text()
    pic_dic = os.path.split(filepath_csv)[0] + '/' + 'Saved Photo' + '/'
    pic_dic_0 = os.path.split(filepath_csv)[0] + '/' + 'Saved Photo'
    isExists = os.path.exists(pic_dic)
    if not isExists:
        os.makedirs(pic_dic)
    self.text_pic.setText(pic_dic_0)

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

            if graph_flag == 'Histogram' and data_ini.loc[i, 'Monitor'] == data_ini.loc[i, 'Monitor']:  # Monitor列判定是否画图

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
                                                          self.entry_qty.text())
                out_data = pd.concat([out_data, out_his], axis=1)
                self.text_info.append(print_info + '\n')
                self.cursot = self.text_info.textCursor()
                self.text_info.moveCursor(self.cursot.End)
                # self.text_info.see(END)
                text_info_line = text_info_line + 1
                self.text_info.update()

            elif graph_flag == 'Boxplot' and data_ini.loc[i, 'Monitor'] == data_ini.loc[i, 'Monitor']:  # Monitor列判定是否画图

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
                    print_info, out_box = Boxplot.Boxplot(Station, pic_dic, config_name, graph_class, USL_tem, LSL_tem,
                                                          boxplot_data_tem, data_ini.loc[i, 'Axis Upper Limit'],
                                                          data_ini.loc[i, 'Axis Lower Limit'], boxplot_X_name,
                                                          ini_pic_order, self.entry_qty.text())
                    boxplot_X_name = []
                    out_data = pd.concat([out_data, out_box], axis=1)
                    self.text_info.append(print_info + '\n')
                    text_info_line = text_info_line + 1
                    self.cursot = self.text_info.textCursor()
                    self.text_info.moveCursor(self.cursot.End)
                    # self.text_info.see(END)
                    self.text_info.update()
                    boxplot_data_tem = pd.DataFrame()
                    USL_tem = []
                    LSL_tem = []

            elif graph_flag == 'Sweep' and data_ini.loc[i, 'Monitor'] == data_ini.loc[i, 'Monitor']:  # Monitor列判定是否画图
                for j in data_csv.columns:
                    if j.find(item_name) >= 0:
                        sweep_data_tem = pd.concat([sweep_data_tem, data_csv[j]], axis=1)
                        USL_tem.append(USL[j])
                        LSL_tem.append(LSL[j])
                print_info = Sweep.Sweep(Station, pic_dic, config_name, graph_class, USL_tem, LSL_tem, sweep_data_tem,
                                         data_ini.loc[i, 'Axis Upper Limit'], data_ini.loc[i, 'Axis Lower Limit'],
                                         ini_pic_order, self.entry_qty.text())
                self.text_info.append(print_info + '\n')
                text_info_line = text_info_line + 1
                # self.text_info.see(END)
                self.cursot = self.text_info.textCursor()
                self.text_info.moveCursor(self.cursot.End)
                self.text_info.update()
                sweep_data_tem = pd.DataFrame()
                USL_tem = []
                LSL_tem = []

        else:
            print_info = f'Error -> {item_name} IS not Found on Data !'
            self.text_info.append(print_info + '\n')
            self.text_info.tag_add('tag0', str(text_info_line) + '.9',
                                   str(text_info_line) + '.' + str(9 + len(item_name)))  # Error 打印突出显示
            self.text_info.tag_config('tag0', background='red', font=('Times'))
            text_info_line = text_info_line + 1
            # self.text_info.see(END)
            self.text_info.update()

    # 保存数据到Excel Table
    save_table_name = project + '_' + Station + f'_Summary Table_{now_date}.xlsx'
    out_data_path = os.path.split(filepath_csv)[0] + '/' + save_table_name
    app = xw.App(visible=False, add_book=False)
    if os.path.exists(out_data_path):
        wb = app.books.open(out_data_path)
        for i in wb.sheets:
            if i.name == self.entry_config.get():
                wb.sheets[self.entry_config.get()].delete()
        ws = wb.sheets.add(self.entry_config.get())
        ws.range('A1').expand('table').value = out_data
        ws.range('A1').api.EntireRow.Delete()
    else:
        wb = app.books.add()
        wb.sheets[0].name = self.entry_config.text()
        ws = wb.sheets[0]
        ws.range('A1').expand('table').value = out_data
        ws.range('A1').api.EntireRow.Delete()
    wb.save(out_data_path)
    wb.close()
    app.quit()

    time_delta = time.time() - time_t0
    self.text_info.append('Finished All in ' + str(round(time_delta, 1)) + ' Seconds' + '\n')
    self.text_info.tag_add('tag', str(text_info_line) + '.0', str(text_info_line + 1) + '.0')  # Finish 打印突出显示
    self.text_info.tag_config('tag', background='green', font=('Times', 15))
    # self.text_info.see(END)
    self.text_info.update()
    text_info_line = text_info_line + 1
