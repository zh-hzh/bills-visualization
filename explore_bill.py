# -*- coding: utf-8 -*-
"""探查微信/支付宝账单真实文件结构，确定解析用的表头位置和列顺序。"""

import pandas as pd
import os

DATA_DIR = r"D:\每日任务\2026-08-25\data"

# 遍历目录找文件
for fname in os.listdir(DATA_DIR):
    fpath = os.path.join(DATA_DIR, fname)
    print("=" * 60)
    print("文件:", fname)

    if fname.lower().endswith((".xlsx", ".xls")):
        # 微信 xlsx：列出所有 sheet
        xl = pd.ExcelFile(fpath)
        print("  sheets:", xl.sheet_names)
        for sheet in xl.sheet_names:
            print(f"  --- sheet: {sheet} ---")
            # 读前 25 行，看表头位置
            df_head = pd.read_excel(fpath, sheet_name=sheet, header=None, nrows=25)
            print("  行数可见:", len(df_head))
            print((df_head.head(25).to_string()))  # 直接显示原始行
            break  # 只看第一个 sheet
    elif fname.lower().endswith(".csv"):
        # 支付宝 csv：探查编码和结构
        for enc in ["utf-8", "gbk", "gb18030"]:
            try:
                with open(fpath, "r", encoding=enc) as f:
                    lines = f.readlines()
                print(f"  encoding 成功: {enc}, 总行数: {len(lines)}")
                print("  --- 前 10 行原始文本 ---")
                for i, line in enumerate(lines[:10]):
                    print(f"  [{i}] {line.rstrip()}")
                print("  --- 最后 10 行原始文本 ---")
                for i, line in enumerate(lines[-10:]):
                    print(f"  [-{len(lines)-i}] {line.rstrip()}")
                break
            except Exception as e:
                print(f"  {enc} 失败: {e}")
    else:
        print("  未知类型")