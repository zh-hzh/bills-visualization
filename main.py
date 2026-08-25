# -*- coding: utf-8 -*-
"""
账单分析工具 main.py
====================
功能：读取微信/支付宝官方导出的账单，自动分类，生成可视化 HTML 报告。

使用：
    python main.py                 # 弹窗选择账单文件
    python main.py -n               # 无窗口模式（直接扫描 data 目录）

生成：账单分析报告.html（浏览器打开即可看，含饼图/柱状图/折线图，每图带金额）
"""

import os
import sys
import base64
import datetime
import webbrowser
import tkinter as tk
from tkinter import filedialog

import pandas as pd
import matplotlib

matplotlib.use("Agg")  # 后端：不弹窗口，直接存图片
import matplotlib.pyplot as plt

# 引入分类规则（单独文件，可自行修改）
from rules import (INCOME_RULES, EXPENSE_RULES, CURRENCY_PREFIXES,
                   TRANSFER_KEYWORDS, EXPENSE_TRANSFER_AS_CONSUMPTION)

# 中文字体（Windows 用微软雅黑，避免乱码）
plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

CUR_DIR = os.path.dirname(os.path.abspath(__file__))


# ---------------------------------------------------------------------------
# 一、带说明的账单文件解析（微信 xlsx / 支付宝 gbk csv）
# ---------------------------------------------------------------------------

def find_header_row(raw_df, header_kw="交易时间"):
    """在原始数据框里找到表头所在行索引（包含'交易时间'的行）。"""
    for i in range(len(raw_df)):
        row_text = " ".join(str(x) for x in raw_df.iloc[i].tolist() if pd.notna(x))
        if header_kw in row_text:
            return i
    raise ValueError("未找到表头行（缺少关键词: %s）" % header_kw)


def clean_amount(value):
    """去除金额里的货币符号和空格，转成 float。"""
    s = str(value).strip()
    for p in CURRENCY_PREFIXES:
        s = s.replace(p, "")
    s = s.strip()
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def read_wechat(path):
    """解析微信账单 xlsx：动态定位表头，统一列名。"""
    # 先整表读入（不带表头），找出表头行
    raw = pd.read_excel(path, header=None, dtype=str)
    header_idx = find_header_row(raw)
    # 从表头行重新读，得到正常 DataFrame
    df = pd.read_excel(path, header=header_idx, dtype=str)
    df = df.rename(columns=lambda c: str(c).strip())

    # 只保留关键列（用列名找，防止位置变化）
    keep = {
        "交易时间": "交易时间",
        "交易类型": "交易类型",
        "交易对方": "交易对方",
        "商品": "商品说明",
        "收/支": "收/支",
        "金额(元)": "金额",
        "当前状态": "状态",
    }
    for col in list(keep):
        if col not in df.columns:
            df[col] = ""  # 缺列补空
    df = df[list(keep)].rename(columns=keep)

    df["来源"] = "微信"
    # 中性交易（收/支 为 '/' 或空）是零钱通存取/还款等，不计入统计
    df["收/支"] = df["收/支"].astype(str).str.strip()
    df = df[df["收/支"].isin(["收入", "支出"])]
    try:
        df["交易时间"] = pd.to_datetime(df["交易时间"], errors="coerce")
    except Exception:
        pass
    df["金额"] = df["金额"].apply(clean_amount)
    return df


def find_header_row_in_file(path, encoding, header_kw="交易时间"):
    """逐行读文本，找到包含表头关键词的行号（0 起）。"""
    with open(path, "r", encoding=encoding) as f:
        for i, line in enumerate(f):
            if header_kw in line:
                return i
    # 找不到再改用 utf-8 兜底
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        for i, line in enumerate(f):
            if header_kw in line:
                return i
    raise ValueError("未找到表头行（缺少关键词: %s）" % header_kw)


def read_alipay(path):
    """解析支付宝账单 csv（GBK 编码）：动态定位表头。"""
    import io
    # 逐行读文本找到表头行，然后把【表头行 + 全部数据行】拼成字符串喂给 pandas。
    # 不用 pandas 的 header= 参数：脏文件里说明行/空行会导致 pandas 行号错位。
    header_idx = find_header_row_in_file(path, "gbk")
    with open(path, "r", encoding="gbk") as f:
        lines = f.readlines()
    content = "".join(lines[header_idx:])  # 从表头行开始，保证物理行号可控
    df = pd.read_csv(io.StringIO(content), dtype=str)
    df = df.rename(columns=lambda c: str(c).strip().strip("﻿").strip())

    keep = {
        "交易时间": "交易时间",
        "交易分类": "交易类型",
        "交易对方": "交易对方",
        "商品说明": "商品说明",
        "收/支": "收/支",
        "金额": "金额",
        "交易状态": "状态",
    }
    for col in list(keep):
        if col not in df.columns:
            df[col] = ""
    df = df[list(keep)].rename(columns=keep)

    df["来源"] = "支付宝"
    # 中性交易（不计收支 / 空）不计入统计
    df["收/支"] = df["收/支"].astype(str).str.strip()
    df = df[df["收/支"].isin(["收入", "支出"])]
    try:
        df["交易时间"] = pd.to_datetime(df["交易时间"], errors="coerce")
    except Exception:
        pass
    df["金额"] = df["金额"].apply(clean_amount)
    return df


# ---------------------------------------------------------------------------
# 二、分类
# ---------------------------------------------------------------------------

def is_transfer(text):
    """判断一笔交易是否属于家庭/账户往来转账（命中关键词）。"""
    for kw in TRANSFER_KEYWORDS:
        if kw and kw in text:
            return True
    return False


def classify(df):
    """按规则给每一笔交易打类别标签（转账按收支方向区分）。"""
    def apply_rules(text, rules):
        for rule in rules:
            for kw in rule["关键词"]:
                if kw and kw in text:
                    return rule["类别"]
        return rules[-1]["类别"]  # 兜底类别在列表最后

    cats = []
    for _, row in df.iterrows():
        # 拼接匹配文本：交易对方 + 商品说明 + 交易类型
        text = "%s %s %s" % (row.get("交易对方", ""), row.get("商品说明", ""), row.get("交易类型", ""))
        # 命中「往来转账」关键词 → 按收支方向归成不同类别
        if is_transfer(text):
            if row.get("收/支") == "收入":
                cats.append("家庭往来")   # 钱进来（父母给生活费等）→ 计入收入统计
            else:
                cats.append("往来转账")   # 钱出去（转给家人等）→ 不参与消费统计
            continue
        if row.get("收/支") == "收入":
            cat = apply_rules(text, INCOME_RULES)
        else:
            cat = apply_rules(text, EXPENSE_RULES)
        cats.append(cat)
    df["类别"] = cats
    return df


def apply_transfer_switch(df):
    """按 EXPENSE_TRANSFER_AS_CONSUMPTION 开关决定支出侧「往来转账」是否计入消费统计。

    设计（收入为主原则）：
      收入侧的「家庭往来」转账 永远计入收入 —— 父母给的生活费就是收入大头，
      必须展示出来（否则收入统计为空，报告没意义）。
      支出侧的「往来转账」（转给家人/账户划转）默认不算消费：
        False = 剔除，不参与消费分类占比（推荐，避免"转给母亲5000"占98%的假象）
        True  = 计入，当普通支出看。
    """
    if EXPENSE_TRANSFER_AS_CONSUMPTION:
        return df
    # 默认：只剔除支出侧的往来转账；收入侧的家庭往来保留
    return df[~((df["类别"] == "往来转账") & (df["收/支"] == "支出"))]


# ---------------------------------------------------------------------------
# 三、出图（matplotlib → PNG，再内嵌进 HTML）
# ---------------------------------------------------------------------------

def _fig_to_base64(fig):
    """把 matplotlib 图形转成 <img> 标签（base64 内嵌，离线可显示）。"""
    import io
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=130, bbox_inches="tight")
    plt.close(fig)
    b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
    return '<img src="data:image/png;base64,%s" alt="图表"/>' % b64


def fig_pie(amounts_by_cat, title):
    """饼图：类别占比，图上标【类别 金额 占比】。"""
    cats = list(amounts_by_cat.keys())
    vals = list(amounts_by_cat.values())
    total = sum(vals) or 1

    def autopct(p):
        v = p / 100 * total
        return "%.1f%%\n%.2f元" % (p, v)

    fig, ax = plt.subplots(figsize=(7, 5))
    colors = plt.cm.Set3(range(len(cats)))
    ax.pie(vals, labels=cats, autopct=autopct, colors=colors, startangle=90,
           textprops={"fontsize": 11})
    ax.set_title(title, fontsize=14, fontweight="bold")
    ax.axis("equal")
    return _fig_to_base64(fig)


def fig_bar(amounts_by_cat, title, color="#4C9AFF"):
    """柱状图：每类金额，柱顶标金额数字。"""
    cats = list(amounts_by_cat.keys())
    vals = list(amounts_by_cat.values())
    fig, ax = plt.subplots(figsize=(8, 5))
    bars = ax.bar(cats, vals, color=color, edgecolor="white", width=0.6)
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + total_y_offset(vals),
                "%.2f" % v, ha="center", va="bottom", fontsize=10)
    ax.set_ylabel("金额（元）")
    ax.set_title(title, fontsize=14, fontweight="bold")
    ax.set_ylim(0, max(vals) * 1.15 if vals else 1)
    plt.xticks(rotation=30, ha="right")
    return _fig_to_base64(fig)


def fig_line(daily, title):
    """折线图：每日收支趋势，每点标金额。"""
    # 防御单边数据：某月只有收入/只有支出时，pivot_table 会缺列，补零列再画线
    daily = daily.copy()
    for col in ("收入", "支出"):
        if col not in daily.columns:
            daily[col] = 0.0
    dates = list(daily.index)
    exp_vals = daily["支出"].tolist()
    inc_vals = daily["收入"].tolist()
    fig, ax = plt.subplots(figsize=(9, 4.5))
    ax.plot(dates, exp_vals, marker="o", color="#F25F5C", label="支出", linewidth=2)
    ax.plot(dates, inc_vals, marker="s", color="#8AC926", label="收入", linewidth=2)
    for i, v in enumerate(exp_vals):
        if v:
            ax.text(i, v + 5, "%.0f" % v, ha="center", fontsize=8, color="#F25F5C")
    for i, v in enumerate(inc_vals):
        if v:
            ax.text(i, v + 5, "%.0f" % v, ha="center", fontsize=8, color="#8AC926")
    ax.set_ylabel("金额（元）")
    ax.set_title(title, fontsize=14, fontweight="bold")
    ax.legend()
    dates_fmt = [d.strftime("%m-%d") for d in dates]
    ax.set_xticks(dates)
    ax.set_xticklabels(dates_fmt, rotation=45, ha="right", fontsize=8)
    return _fig_to_base64(fig)


def total_y_offset(vals):
    """柱状图文字上浮偏移，避免贴边。"""
    m = max(vals) if vals else 1
    return m * 0.02


# ---------------------------------------------------------------------------
# 四、生成 HTML 报告
# ---------------------------------------------------------------------------

def build_report(df, stats):
    """生成单个 HTML 文件（图片 base64 内嵌，离线可看）。"""
    # ---- 图 1：收入分类 饼图 ----
    inc_cat = stats["收入分类金额"]
    pie_inc = fig_pie(inc_cat, "收入分类占比（饼图）") if inc_cat else ""
    # ---- 图 2：收入分类 柱状图 ----
    bar_inc = fig_bar(inc_cat, "各类收入金额（柱状图）", "#8AC926") if inc_cat else ""
    # ---- 图 3：支出分类 饼图 ----
    exp_cat = stats["支出分类金额"]
    pie_exp = fig_pie(exp_cat, "消费支出分类占比（饼图）") if exp_cat else ""
    # ---- 图 4：支出分类 柱状图 ----
    bar_exp = fig_bar(exp_cat, "各类消费支出金额（柱状图）", "#4C9AFF") if exp_cat else ""
    # ---- 图 5：每日收支趋势 折线图 ----
    line_trend = fig_line(stats["每日收支"], "每日收支趋势（折线图）") if not stats["每日收支"].empty else ""

    # 汇总卡片（消费支出与往来转出分开，避免转账污染真实消费）
    cards = (
        '<div class="card"><div class="num">%.2f</div><div class="lbl">总收入（元）</div></div>'
        '<div class="card exp"><div class="num">%.2f</div><div class="lbl">消费支出（元）</div></div>'
        '<div class="card tsf"><div class="num">%.2f</div><div class="lbl">往来转出（元）</div></div>'
        '<div class="card bal"><div class="num">%.2f</div><div class="lbl">净结余（元）</div></div>'
        '<div class="card cnt"><div class="num">%d / %d</div><div class="lbl">收入笔数 / 消费支出笔数</div></div>'
    ) % (stats["总收入"], stats["消费支出"], stats["往来转出"], stats["净结余"],
         stats["收入笔数"], stats["支出笔数"])

    # 明细表
    rows = ""
    for _, r in df.iterrows():
        rows += (
            "<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td>"
            "<td class='%s'>%.2f</td></tr>"
        ) % (
            r["交易时间"].strftime("%Y-%m-%d %H:%M") if pd.notna(r["交易时间"]) else "",
            r["来源"], r["类别"], r["交易对方"], r["商品说明"],
            "inc" if r["收/支"] == "收入" else "exp",
            r["金额"])

    html = f"""<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>账单分析报告</title>
<style>
  body {{ font-family: "Microsoft YaHei", sans-serif; margin: 0; background: #f5f7fa; color: #333; }}
  h1 {{ background: #2f5bff; color: #fff; padding: 24px 40px; margin: 0; font-size: 24px; }}
  sub {{ display:block; font-size:13px; font-weight:normal; opacity:.85; margin-top:4px;}}
  .wrap {{ max-width: 1100px; margin: 20px auto; padding: 0 20px; }}
  .cards {{ display: flex; gap: 16px; flex-wrap: wrap; }}
  .card {{ flex:1; min-width:180px; background:#fff; border-radius:12px; padding:18px 22px;
           box-shadow: 0 2px 8px rgba(0,0,0,.06); border-top:4px solid #8AC926; }}
  .card.exp {{ border-top-color:#F25F5C; }} .card.tsf {{ border-top-color:#b074e0; }}
  .card.bal {{ border-top-color:#2f5bff; }} .card.cnt {{ border-top-color:#F4A31A; }}
  .num {{ font-size:26px; font-weight:bold; }} .lbl {{ color:#888; font-size:13px; margin-top:4px; }}
  .chart {{ background:#fff; border-radius:12px; padding:18px; margin:20px 0;
            box-shadow:0 2px 8px rgba(0,0,0,.06); }}
  .chart h2 {{ margin:0 0 10px; font-size:17px; color:#2f5bff; }}
  img {{ max-width:100%; }}
  table {{ width:100%; border-collapse:collapse; background:#fff; border-radius:12px; overflow:hidden;
          box-shadow:0 2px 8px rgba(0,0,0,.06); font-size:13px; }}
  th {{ background:#2f5bff; color:#fff; padding:10px 8px; }}
  td {{ border-bottom:1px solid #eee; padding:8px; text-align:left; }}
  tr:hover td {{ background:#f0f4ff; }}
  .inc {{ color:#3a9d23; font-weight:bold; }} .exp {{ color:#d63338; font-weight:bold; }}
</style>
</head>
<body>
<h1>💰 账单分析报告<sub>数据范围由你导入的账单决定 · 程序自动分类（规则见 rules.py 可自行修改）</sub></h1>
<div class="wrap">
  <div class="cards">{cards}</div>
  <div class="chart"><h2>💰 收入分类占比（钱从哪来）</h2>{pie_inc}</div>
  <div class="chart"><h2>💰 收入分类金额对比</h2>{bar_inc}</div>
  <div class="chart"><h2>🛒 消费支出分类占比（钱花在哪）</h2>{pie_exp}</div>
  <div class="chart"><h2>🛒 消费支出分类金额对比</h2>{bar_exp}</div>
  <div class="chart"><h2>📈 每日收支趋势</h2>{line_trend}</div>
  <div class="chart"><h2>交易明细（共 {len(df)} 笔）</h2>
    <table><thead><tr><th>时间</th><th>来源</th><th>类别</th><th>交易对方</th><th>商品说明</th><th>金额</th></tr></thead>
    <tbody>{rows}</tbody></table>
  </div>
</div>
</body>
</html>"""

    out = os.path.join(OUTPUT_DIR, "账单分析报告.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    return out


# ---------------------------------------------------------------------------
# 五、主流程
# ---------------------------------------------------------------------------

def detect_source(path):
    """自动判断账单来源。三层：文件名优先 → 扩展名 → 内容特征。"""
    fname = os.path.basename(path)
    # 第一层：文件名里带「微信」/「支付宝」
    if "微信" in fname:
        return "微信"
    if "支付宝" in fname:
        return "支付宝"
    # 第二层：扩展名兜底（微信官方导出 xlsx，支付宝官方导出 csv）
    low = fname.lower()
    if low.endswith((".xlsx", ".xls")):
        return "微信"
    if low.endswith(".csv"):
        return "支付宝"
    # 第三层：内容特征（读开头几行找来源标记）
    try:
        with open(path, "r", encoding="gbk", errors="replace") as f:
            head = f.read(3000)
        if "支付宝" in head:
            return "支付宝"
        if "微信" in head:
            return "微信"
    except Exception:
        pass
    return "未知"


def main():
    """主流程：多选文件 → 逐个自动识别来源 → 解析 → 分类 → 出图 → 生成 HTML。"""
    # 注意：print 只用 GBK 兼容字符（中文+数字），exe 打包后 stdout 按系统编码(GBK)，
    #       任何 emoji/特殊符号都会导致 UnicodeEncodeError 崩溃。
    root = tk.Tk()
    root.withdraw()

    print("请选择账单文件（可多选：按 Ctrl/Shift 一次选多个，或拖选）")
    paths = filedialog.askopenfilenames(
        title="选择账单文件（可多选）",
        filetypes=[("账单文件", "*.xlsx *.csv"), ("所有文件", "*.*")])

    if not paths:
        print("未选择任何账单，退出。")
        sys.exit(0)

    df_list = []
    n_ok = 0
    for path in paths:
        src = detect_source(path)
        try:
            if src == "微信":
                df = read_wechat(path)
            elif src == "支付宝":
                df = read_alipay(path)
            else:
                print("无法识别来源，跳过：%s" % os.path.basename(path))
                continue
            print("导入 %d 笔（%s，%s）" % (len(df), src, os.path.basename(path)))
            df_list.append(df)
            n_ok += 1
        except Exception as e:
            msg = str(e).replace("\n", " ")[:100]
            print("解析失败，跳过：%s（%s）" % (os.path.basename(path), msg))

    if not df_list:
        print("没有成功导入任何账单，退出。")
        sys.exit(1)

    df = pd.concat(df_list, ignore_index=True)

    df = pd.concat(df_list, ignore_index=True)
    df = df.dropna(subset=["交易时间"])
    df = df.sort_values("交易时间").reset_index(drop=True)
    df = classify(df)

    # 区分「统计口径」与「明细展示」
    #   df_full ：全部交易（含往来转账），用于明细表展示
    #   df_stat ：参与统计的口径（按 INCLUDE_TRANSFERS 开关剔除往来转账）
    df_full = df.copy()
    df_stat = apply_transfer_switch(df)

    # ---- 统计（口径：收入全计入；支出剔除「往来转账」，转出单独统计）----
    inc = df_stat[df_stat["收/支"] == "收入"]
    exp = df_stat[df_stat["收/支"] == "支出"]
    transfer_out = df_full[(df_full["收/支"] == "支出") & (df_full["类别"] == "往来转账")]
    total_income = inc["金额"].sum()
    consume_exp = exp["金额"].sum()
    transfer_out_sum = transfer_out["金额"].sum()
    stats = {
        "总收入": total_income,
        "消费支出": consume_exp,
        "往来转出": transfer_out_sum,
        "净结余": total_income - consume_exp - transfer_out_sum,
        "收入笔数": len(inc),
        "支出笔数": len(exp),
        "收入分类金额": inc.groupby("类别")["金额"].sum().sort_values(ascending=False).to_dict(),
        "支出分类金额": exp.groupby("类别")["金额"].sum().sort_values(ascending=False).to_dict(),
        "每日收支": df_stat.pivot_table(index=df_stat["交易时间"].dt.normalize(),
                                        columns="收/支", values="金额",
                                        aggfunc="sum", fill_value=0),
    }

    out = build_report(df_full, stats)
    print("报告已生成：%s" % out)
    webbrowser.open("file://" + out.replace("\\", "/"))
    print("已自动打开浏览器。若未打开，请手动双击该文件。")


if __name__ == "__main__":
    main()