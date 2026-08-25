# -*- coding: utf-8 -*-
"""验证脚本：用真实账单跑一遍解析 + 分类，打印结果供人工核对。"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from main import read_wechat, read_alipay, classify  # noqa: E402

DATA_DIR = r"D:\每日任务\2026-08-25\data"

wechat_file = None
alipay_file = None
for fname in os.listdir(DATA_DIR):
    if "微信" in fname:
        wechat_file = os.path.join(DATA_DIR, fname)
    elif "支付宝" in fname:
        alipay_file = os.path.join(DATA_DIR, fname)

print("=" * 70)
print("微信账单解析:")
if wechat_file:
    df_wx = read_wechat(wechat_file)
    print("  文件:", os.path.basename(wechat_file))
    print("  解析笔数:", len(df_wx), "| 收入:", (df_wx["收/支"] == "收入").sum(),
          "| 支出:", (df_wx["收/支"] == "支出").sum())
    print("  列名:", list(df_wx.columns))
    print(df_wx[["交易时间", "来源", "收/支", "金额", "交易对方", "商品说明"]].head(12).to_string())
    print("  ...")
    print("  微信表头统计（合计金额）: 收入 %.2f | 支出 %.2f" % (
        df_wx[df_wx["收/支"] == "收入"]["金额"].sum(),
        df_wx[df_wx["收/支"] == "支出"]["金额"].sum()))
else:
    print("  未找到微信文件")

print("=" * 70)
print("支付宝账单解析:")
if alipay_file:
    df_zfb = read_alipay(alipay_file)
    print("  文件:", os.path.basename(alipay_file))
    print("  解析笔数:", len(df_zfb), "| 收入:", (df_zfb["收/支"] == "收入").sum(),
          "| 支出:", (df_zfb["收/支"] == "支出").sum())
    print("  列名:", list(df_zfb.columns))
    print(df_zfb[["交易时间", "来源", "收/支", "金额", "交易对方", "商品说明"]].head(12).to_string())
    print("  支付宝表头统计（合计金额）: 收入 %.2f | 支出 %.2f" % (
        df_zfb[df_zfb["收/支"] == "收入"]["金额"].sum(),
        df_zfb[df_zfb["收/支"] == "支出"]["金额"].sum()))
else:
    print("  未找到支付宝文件")

print("=" * 70)
print("合并 + 分类结果:")
if wechat_file or alipay_file:
    import pandas as pd
    from main import apply_transfer_switch, build_report
    parts = []
    if wechat_file:
        parts.append(df_wx)
    if alipay_file:
        parts.append(df_zfb)
    df_all = pd.concat(parts, ignore_index=True)
    df_all = df_all.dropna(subset=["交易时间"])
    df_all = classify(df_all)
    print("  总笔数:", len(df_all))
    print("  支出分类合计（全量，含往来转账）:")
    print(df_all[df_all["收/支"] == "支出"].groupby("类别")["金额"].sum()
          .sort_values(ascending=False).to_string())
    print("  收入分类合计（全量，含家庭往来）:")
    print(df_all[df_all["收/支"] == "收入"].groupby("类别")["金额"].sum()
          .sort_values(ascending=False).to_string())
    print("=" * 70)
    df_stat = apply_transfer_switch(df_all)
    inc = df_stat[df_stat["收/支"] == "收入"]
    exp = df_stat[df_stat["收/支"] == "支出"]
    tsf_out = df_all[(df_all["收/支"] == "支出") & (df_all["类别"] == "往来转账")]
    print("统计口径（剔除消费口径的往来转出）:")
    print("  总收入 %.2f | 消费支出 %.2f | 往来转出 %.2f | 净结余 %.2f" % (
        inc["金额"].sum(), exp["金额"].sum(), tsf_out["金额"].sum(),
        inc["金额"].sum() - exp["金额"].sum() - tsf_out["金额"].sum()))
    print("  收入分类（含家庭往来，收入为主）:")
    print(inc.groupby("类别")["金额"].sum().sort_values(ascending=False).to_string())
    print("  消费支出分类（不含往来转出）:")
    print(exp.groupby("类别")["金额"].sum().sort_values(ascending=False).to_string())

    print("=" * 70)
    print("端到端：生成 HTML 报告…")
    stats = {
        "总收入": inc["金额"].sum(),
        "消费支出": exp["金额"].sum(),
        "往来转出": tsf_out["金额"].sum(),
        "净结余": inc["金额"].sum() - exp["金额"].sum() - tsf_out["金额"].sum(),
        "收入笔数": len(inc),
        "支出笔数": len(exp),
        "收入分类金额": inc.groupby("类别")["金额"].sum().sort_values(ascending=False).to_dict(),
        "支出分类金额": exp.groupby("类别")["金额"].sum().sort_values(ascending=False).to_dict(),
        "每日收支": df_stat.pivot_table(index=df_stat["交易时间"].dt.normalize(),
                                        columns="收/支", values="金额",
                                        aggfunc="sum", fill_value=0),
    }
    out = build_report(df_all, stats)
    ok_img = "<img " in open(out, "r", encoding="utf-8").read()
    print("  HTML 已生成:", out)
    print("  文件大小: %.1f KB | 含图片(base64): %s" % (os.path.getsize(out) / 1024, ok_img))