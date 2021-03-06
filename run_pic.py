import pandas as pd
import xlwings as xw
import csv_process
import Histogram
import Boxplot
import Sweep_1 as Sweep
import time
import os
from datetime import datetime
from PyQt5.QtWidgets import *


def run_pic(main_frame):
    dt = datetime.now()
    now_date = dt.strftime('%Y%m%d')
    overdue_date = '20230704'
    time_t0 = time.time()
    main_frame.text_info.clear()
    main_frame.text_info.append('Start -> Reading CSV File')
    data_csv, USL, LSL, Overlay, Station, project = csv_process.csv_process(main_frame.text_CSV.text())
    config_name = main_frame.entry_config.text()
    data_ini = pd.read_excel(main_frame.text_ini.text(), sheet_name=Station)

    graph_flag = ''  # 画什么图
    graph_class = ''  # 图的分类标签

    boxplot_data_tem = pd.DataFrame()
    sweep_data_tem = pd.DataFrame()
    USL_tem = []
    LSL_tem = []
    boxplot_X_name = []

    # 创建保存图片的文件夹
    filepath_csv = main_frame.text_CSV.text()
    pic_dic = os.path.split(filepath_csv)[0] + '/' + 'Saved Photo' + '/'
    pic_dic_0 = os.path.split(filepath_csv)[0] + '/' + 'Saved Photo'
    isExists = os.path.exists(pic_dic)
    if not isExists:
        os.makedirs(pic_dic)
    main_frame.text_pic.setText(pic_dic_0)

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
                                                          main_frame.entry_qty.text())
                out_data = pd.concat([out_data, out_his], axis=1)
                main_frame.text_info.append(print_info)
                main_frame.cursot = main_frame.text_info.textCursor()
                main_frame.text_info.moveCursor(main_frame.cursot.End)
                QApplication.processEvents()

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
                                                          ini_pic_order, main_frame.entry_qty.text())
                    boxplot_X_name = []
                    out_data = pd.concat([out_data, out_box], axis=1)
                    main_frame.text_info.append(print_info)
                    main_frame.cursot = main_frame.text_info.textCursor()
                    main_frame.text_info.moveCursor(main_frame.cursot.End)
                    QApplication.processEvents()
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
                                         ini_pic_order, main_frame.entry_qty.text())
                main_frame.text_info.append(print_info)
                main_frame.cursot = main_frame.text_info.textCursor()
                main_frame.text_info.moveCursor(main_frame.cursot.End)
                QApplication.processEvents()
                sweep_data_tem = pd.DataFrame()
                USL_tem = []
                LSL_tem = []

        else:
            print_info = '<span style="background:red;">' + f'Error -> {item_name} IS not Found on Data !' + '</span>'
            main_frame.text_info.append(print_info)

    # 保存数据到Excel Table
    save_table_name = project + '_' + Station + f'_Summary Table_{now_date}.xlsx'
    out_data_path = os.path.split(filepath_csv)[0] + '/' + save_table_name
    app = xw.App(visible=False, add_book=False)
    if os.path.exists(out_data_path):
        wb = app.books.open(out_data_path)
        for i in wb.sheets:
            if i.name == main_frame.entry_config.text():
                wb.sheets[main_frame.entry_config.text()].delete()
                break
        if main_frame.entry_config.text() == '':
            ws = wb.sheets.add()
        else:
            ws = wb.sheets.add(main_frame.entry_config.text())
    else:
        if main_frame.entry_config.text() == '':
            wb = app.books.add()
            ws = wb.sheets[0]
        else:
            wb = app.books.add()
            ws = wb.sheets[0]
            ws.name = main_frame.entry_config.text()

    ws.range('A1').expand('table').value = out_data
    ws.range('A1').api.EntireRow.Delete()
    wb.save(out_data_path)
    wb.close()
    app.quit()

    time_delta = time.time() - time_t0

    print_info = '<span style="background:green;">' + 'Finished All in ' + str(
        round(time_delta, 1)) + ' Seconds' + '</span>'
    main_frame.text_info.append(print_info)
    main_frame.text_info.repaint()
