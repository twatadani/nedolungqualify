''' step 1のメイン実行ファイル
全データの入っているExcelファイルを読み込み、thin slice CTが撮影されている
肺結節のみに絞り込み、結果ファイルに書き出す。
'''

import json
import os.path

import openpyxl
from pprint import pprint

__author__ = 'Takeyuki Watadani<twatadani@g.ecc.u-tokyo.ac.jp>'
__version__ = '0.1'
__date__ = '2021/10/25'
__status__ = 'dev'

print('step 1を開始します。')

# configの読み込み
with open('config.json', mode='rt') as fp:
    cfdict = json.load(fp)

# 生データの読み込み
datadir = os.path.expanduser(os.path.expandvars(cfdict['datadir']))
rawfile = os.path.join(datadir, cfdict['rawdata'])

workbook = openpyxl.load_workbook(rawfile)
sheet = workbook['シート1']

# データの確認
rows = sheet.iter_rows()
for row in rows:
    pprint(row)

print('step 1が終了しました。')