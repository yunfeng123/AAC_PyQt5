import pandas as pd
from multiprocessing import cpu_count, Pool
import numpy as np

cores = cpu_count()
partitions = cores


def parallelize(df):
    data_split = np.array_split(df, partitions)
    pool = Pool(cores)
    data = pool.map(func, data_split)
    pool.close()
    pool.join()
    return data


def func(li):
    for i in li.index:
        if li.loc[i, 'Test Pass/Fail Status'] == 'FAIL':
            li.drop(i, axis=0, inplace=True)
    return li


def csv_process(path):
    data_head = pd.read_csv(path, header=1, delimiter=',', nrows=5)
    USL = data_head.iloc[2]
    LSL = data_head.iloc[3]

    data = pd.read_csv(path, delimiter=',', header=None, skiprows=7)
    data.columns = data_head.columns

    data = parallelize(data)
    data_all = pd.concat(data, axis=0)

    Overlay = data_all['Version'].tolist()[0]
    Station = Overlay.split('_')[2]
    project = Overlay.split('_')[0]
    return data_all, USL, LSL, Overlay, Station, project

