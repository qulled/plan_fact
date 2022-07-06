import json
import pandas as pd
import datetime as dt
import os
import shutil
from pathlib import Path

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
cred_file = os.path.join(BASE_DIR, 'credentials.json')
day = (dt.datetime.now() - dt.timedelta(days=1)).strftime('%d')
month = (dt.datetime.now() - dt.timedelta(days=1)).strftime('%m')
year = (dt.datetime.now() - dt.timedelta(days=1)).strftime("%Y")



def copy_excel(path,path_new):
    with open(cred_file, 'r', encoding="utf-8") as f:
        cred = json.load(f)
        for name in cred:
            if name != 'Савельева':
                shutil.copyfile(f'{path}/{name}-{year}-{month}-{day}.xlsx', f'{path_new}/excel_docs/{name}-{year}-{month}-{day}.xlsx')
    return


def common_excel(paths):
    with open(cred_file, 'r', encoding="utf-8") as f:
        cred = json.load(f)
        for name in cred:
            if name != 'Савельева':
                df = pd.read_excel(f'{paths}/{name}-{year}-{month}-{day}.xlsx')
                df = df.drop(index=[0])
                df.to_excel(f'{paths}/{name}-{year}-{month}-{day}.xlsx')
    path = Path(f'{paths}/')
    df = pd.concat([pd.read_excel(f) for f in path.glob("*.xlsx")],
                   ignore_index=True)
    return df.to_excel(f'final_excel/ALL-{year}-{month}-{day}.xlsx')


def del_start_excel(path):
    with open(cred_file, 'r', encoding="utf-8") as f:
        cred = json.load(f)
        for name in cred:
            if name != 'Савельева':
                file = os.path.join(f'{path}/{name}-{year}-{month}-{day}.xlsx')
                os.remove(file)
    return
