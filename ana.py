#!/usr/bin/python
# -*- coding: utf-8 -*-

import pandas as pd
import jieba
from tqdm import tqdm
import csv

excel_path = "58318_軍機繞台資料(四組)_190304.xlsx"
compare_path = "58279_NTUSD_negative_utf8.txt"
compare_dict = dict()
output=[["斷詞內容", "負面詞頻"]]


def get_word_frequency(row):
    content = row['內容']
    count = 0
    jieba_list = []
    if content and type(content) == str:
        seg_list = jieba.cut(content, cut_all=False)
        for word in seg_list:
            jieba_list.append(word)
            if compare_dict.get(word):
                count += 1
        content = '/'.join(jieba_list)

    output.append([content, count])

with open(compare_path, 'r', encoding="utf8") as f:
    compare_list = f.readlines()
    for word in compare_list:
        word = word.replace('\n', '')
        compare_dict[word] = True

# print(compare_dict.keys())

# df = pd.read_excel(excel_path)
xl = pd.ExcelFile(excel_path)
sheet_list = xl.sheet_names  # see all sheet names

# target is 1 and 3
target_sheet_index = [1, 3]
for sheet_index in target_sheet_index:
    sheet_name = sheet_list[sheet_index]
    df = xl.parse(sheet_name)

    # print(df.columns)
    # row = df.iloc[1]

    for idx, row in tqdm(df.iterrows()):
        get_word_frequency(row)

    with open('word_freq_{}.csv'.format(sheet_name), 'w') as f:
        w = csv.writer(f)
        w.writerows(output)


