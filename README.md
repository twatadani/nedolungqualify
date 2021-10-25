# nedolungqualify
NEDO肺結節予備調査Excelの仕分け

2021年8-9月に坂本、谷島、仲谷、金丸先生に調べて貰った2017-2021年の肺結節の調査結果を仕分けして新しいExcelファイルに整理するスクリプト。  

---

## 生成するファイル

* step 1: 全データ → 「thin slice CTが撮影してある肺結節」へ絞り込む
* step 2: 「thin slice CTが撮影してある肺結節」→ 機種ごと、良悪性に分けた肺結節へ整理する。

---

## 必要なライブラリ

本スクリプト群はPython, openpyxlを用いて作成する。

```shell-session
$ conda install openpyxl
```

---

## version 0.1へ向けたTODO

* config.jsonを作成してコンフィグ事項を記載する
* config.jsonを読み込みdictに保持する
* 全データのファイルをopenpyxlで読み込み、読み込みができていることを確認する。
