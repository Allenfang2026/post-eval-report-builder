#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_charts.py —— 出报告要用的图（matplotlib）+ 防伪二维码占位（qrcode）

投后评价报告里常见的图：
    1) 投资控制对比图（工程类命门）：概算/预算/合同/结算 等各环节金额柱状对比，
       直观展示"钱一路花下来控制得怎么样"。
    2) 财务趋势图（股权类常见）：营收/净利/现金流等逐年折线。
    3) 财务指标柱状图：几个关键比率横向对比。
    4) 防伪二维码占位：机构报告页脚"扫码验真"的二维码。真码由机构系统每份生成，
       这里出的是占位/示例码，模板里预留位置。

全部输出 PNG，交给 build_report.py 用 docxtpl 的 InlineImage 插进报告。

依赖（本机已装）：matplotlib、qrcode[pil]

用法：
    python3 make_charts.py --demo                 # 出一套示例图到 ./charts/ 自检
    作为库：
        from make_charts import investment_control_bar, trend_line, indicator_bar, qrcode_png
"""

import os
import sys

import matplotlib
matplotlib.use("Agg")  # 无界面后端，服务器/headless 也能出图
import matplotlib.pyplot as plt
from matplotlib import font_manager


# ---------- 中文字体 ----------
def _setup_cn_font():
    """
    matplotlib 默认不显示中文（豆腐块）。这里挑一个系统里存在的中文字体。
    macOS 常见：Songti SC / PingFang SC / STHeiti；找不到就用 sans-serif（图里中文可能变框，
    但不影响脚本跑通，评估师本机装了中文字体即可正常）。
    """
    candidates = ["Songti SC", "PingFang SC", "STHeiti", "Heiti SC",
                  "Arial Unicode MS", "SimHei", "Microsoft YaHei"]
    available = {f.name for f in font_manager.fontManager.ttflist}
    for name in candidates:
        if name in available:
            plt.rcParams["font.sans-serif"] = [name]
            break
    plt.rcParams["axes.unicode_minus"] = False  # 负号正常显示


_setup_cn_font()

# 机构风格：偏正式的红/灰配色（机构报告是红头风格）
BAR_COLOR = "#B22222"
BAR_COLOR2 = "#8C8C8C"


def investment_control_bar(labels, values, out_path, title="投资控制各环节金额对比",
                           unit="万元"):
    """
    投资控制对比柱状图（工程类核心）。
    labels: ["概算", "预算", "招标控制价", "合同", "结算送审", "结算审定"]
    values: 对应金额（数字），单位自定（默认万元）。
    """
    fig, ax = plt.subplots(figsize=(8, 4.5))
    bars = ax.bar(labels, values, color=BAR_COLOR, width=0.6)
    ax.set_ylabel(f"金额（{unit}）")
    ax.set_title(title, fontsize=13)
    ax.bar_label(bars, fmt="%.2f", padding=3, fontsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.xticks(rotation=15)
    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    return out_path


def trend_line(years, series, out_path, title="主要财务指标趋势", ylabel="金额（万元）"):
    """
    财务趋势折线图（股权类常见）。
    years:  ["2022", "2023", "2024"]
    series: {"营业收入": [..], "净利润": [..]}  多条线
    """
    fig, ax = plt.subplots(figsize=(8, 4.5))
    for name, vals in series.items():
        ax.plot(years, vals, marker="o", linewidth=2, label=name)
    ax.set_ylabel(ylabel)
    ax.set_title(title, fontsize=13)
    ax.legend()
    ax.grid(True, linestyle="--", alpha=0.4)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    return out_path


def indicator_bar(labels, values, out_path, title="关键财务指标", ylabel="数值"):
    """财务指标横向柱状图（如各类比率对比）。"""
    fig, ax = plt.subplots(figsize=(8, 4.5))
    bars = ax.barh(labels, values, color=BAR_COLOR2)
    ax.set_xlabel(ylabel)
    ax.set_title(title, fontsize=13)
    ax.bar_label(bars, fmt="%.2f", padding=3, fontsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    return out_path


def qrcode_png(data, out_path, box_size=10, border=2):
    """
    出一张二维码 PNG（报告页脚防伪验真码占位）。
    data: 编码内容（真码是机构验真链接，这里可放示例文本或占位 URL）。
    """
    try:
        import qrcode
    except ImportError:
        print("[warn] 未装 qrcode：pip3 install qrcode[pil]", file=sys.stderr)
        return None
    qr = qrcode.QRCode(box_size=box_size, border=border)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(out_path)
    return out_path


def _demo():
    """出一套示例图，验证脚本能实际跑通、出图。"""
    outdir = "charts"
    os.makedirs(outdir, exist_ok=True)

    # 全部用脱敏假数据做示例（非任何真实项目数字，只为验证脚本能出图）。
    # 真实案例数据出正式报告时由 data.json 传入，别把某一单的真实金额焊进这个通用示例函数。
    investment_control_bar(
        ["概算", "预算", "招标控制价", "合同", "结算送审", "结算审定"],
        [1000.00, 1010.00, 980.00, 950.00, 1005.00, 920.00],  # 示例假数据
        os.path.join(outdir, "投资控制对比.png"),
    )
    trend_line(
        ["2022", "2023", "2024"],
        {"营业收入": [3000.0, 3600.0, 4200.0], "净利润": [60.0, 120.0, 200.0]},  # 示例假数据
        os.path.join(outdir, "财务趋势.png"),
    )
    indicator_bar(
        ["资产负债率", "流动比率", "毛利率", "净利率"],
        [50.0, 1.5, 20.0, 2.0],  # 示例假数据
        os.path.join(outdir, "财务指标.png"),
    )
    qrcode_png(
        "https://example.com/verify?report=DEMO-2026",
        os.path.join(outdir, "防伪二维码占位.png"),
    )
    print(f"[done] 示例图已出到 ./{outdir}/：")
    for f in sorted(os.listdir(outdir)):
        print("   -", f)


if __name__ == "__main__":
    if "--demo" in sys.argv or len(sys.argv) == 1:
        _demo()
    else:
        print(__doc__)
