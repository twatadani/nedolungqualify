''' check_ct_variation.py - step 1で作成した中間ファイルからCT機種表記のバリエーションを調べて出力する '''

import os.path
import json

import openpyxl

__author__ = 'Takeyuki Watadani<twatadani@g.ecc.u-tokyo.ac.jp>'
__version__ = '0.1'
__date__ = '2021/10/25'
__status__ = 'dev'

print('check_ct_variationを開始します。')

# configの読み込み
with open('config.json', mode='rt') as fp:
    cfdict = json.load(fp)

# 中間ファイルを読み込む

datadir = os.path.expanduser(os.path.expandvars(cfdict['datadir']))
srcfile = os.path.join(datadir, cfdict['step1resultfile'])

workbook = openpyxl.load_workbook(srcfile)
sheet = workbook['Sheet1']

# 機種を調べる
ct_variation = set()

for i, row in enumerate(sheet.iter_rows()):
    if i != 0:
        ct = row[9].value
        # print(ct)
        if not ct in ct_variation:
            ct_variation.add(ct)
            print(ct, 'を加えました。')


print('check_ct_variationを終了します。')