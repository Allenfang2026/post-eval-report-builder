#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ocr_pdf.py —— 扫描件 / 图片型 PDF 转中文文字

用途：
    投后评价的甲方资料里常有扫描件 PDF（党委会纪要、结算审核报告、许可证等），
    这类 PDF 里是图片不是文字，直接 pdftotext 抽不出东西。本脚本按两级策略处理：
      1) 先用 pdftotext 试着直接抽文字（很多"电子生成再打印扫描"的 PDF 其实带文字层）；
         如果抽出的字数够多，说明是文字型，直接返回，快且准。
      2) 抽不到（或字数太少）→ 判定为扫描件，用 pdftoppm 把每页转成 PNG 图片，
         再用 tesseract（中文简体 chi_sim）逐页 OCR，拼成全文。

扫描件已用这套方法验证可行（tesseract 已装 chi_sim / chi_sim_vert）。

依赖（本机已装）：
    - pdftotext, pdftoppm  (poppler，brew install poppler)
    - tesseract + chi_sim  (brew install tesseract tesseract-lang)

用法：
    python3 ocr_pdf.py 某扫描件.pdf                 # 输出到 某扫描件.txt
    python3 ocr_pdf.py 某扫描件.pdf out.txt          # 指定输出
    python3 ocr_pdf.py 某扫描件.pdf --force-ocr      # 强制走 OCR（跳过 pdftotext 直抽）

也可作为库调用：
    from ocr_pdf import pdf_to_text
    text = pdf_to_text("某扫描件.pdf")
"""

import os
import sys
import shutil
import subprocess
import tempfile

# pdftotext 直抽后，中文字符少于这个数就判定为"扫描件"，转走 OCR
MIN_TEXT_CHARS = 50


def _check_tool(name):
    """确认外部命令存在，不存在给出清晰的安装提示。"""
    if shutil.which(name) is None:
        raise RuntimeError(
            f"缺少命令 `{name}`。请先安装：\n"
            f"  pdftotext/pdftoppm → brew install poppler\n"
            f"  tesseract          → brew install tesseract tesseract-lang"
        )


def try_pdftotext(pdf_path):
    """用 pdftotext 直接抽文字层。抽不到返回空串。"""
    _check_tool("pdftotext")
    try:
        # -layout 尽量保留版面（表格类友好）；- 表示输出到 stdout
        out = subprocess.run(
            ["pdftotext", "-layout", pdf_path, "-"],
            capture_output=True, timeout=120,
        )
        return out.stdout.decode("utf-8", errors="ignore")
    except Exception as e:
        print(f"[warn] pdftotext 直抽失败：{e}", file=sys.stderr)
        return ""


def ocr_via_tesseract(pdf_path, lang="chi_sim", dpi=300):
    """把 PDF 每页转 PNG，逐页 tesseract OCR，拼成全文。"""
    _check_tool("pdftoppm")
    _check_tool("tesseract")

    texts = []
    with tempfile.TemporaryDirectory() as tmp:
        prefix = os.path.join(tmp, "page")
        # PDF → 一系列 page-1.png / page-2.png ...
        subprocess.run(
            ["pdftoppm", "-r", str(dpi), "-png", pdf_path, prefix],
            check=True, timeout=600,
        )
        pngs = sorted(
            f for f in os.listdir(tmp) if f.startswith("page") and f.endswith(".png")
        )
        if not pngs:
            raise RuntimeError("pdftoppm 没有生成任何图片，PDF 可能损坏或为空。")

        for i, png in enumerate(pngs, 1):
            img = os.path.join(tmp, png)
            print(f"[ocr] 第 {i}/{len(pngs)} 页 …", file=sys.stderr)
            # tesseract 输出到 stdout（最后参数 stdout）
            res = subprocess.run(
                ["tesseract", img, "stdout", "-l", lang, "--psm", "6"],
                capture_output=True, timeout=300,
            )
            texts.append(res.stdout.decode("utf-8", errors="ignore"))

    return "\n\n".join(texts)


def pdf_to_text(pdf_path, force_ocr=False, lang="chi_sim"):
    """
    主入口：返回 PDF 全文。
    先试 pdftotext 直抽；抽到的中文字数够多且未强制 OCR 则直接返回，
    否则转扫描件 OCR 流程。
    """
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(pdf_path)

    if not force_ocr:
        direct = try_pdftotext(pdf_path)
        # 统计中日韩文字数量，判断是否为有效文字层
        cjk = sum(1 for ch in direct if "一" <= ch <= "鿿")
        if cjk >= MIN_TEXT_CHARS:
            print(f"[info] pdftotext 直抽成功（中文 {cjk} 字），判定为文字型 PDF。",
                  file=sys.stderr)
            return direct
        print(f"[info] 直抽中文仅 {cjk} 字，判定为扫描件，转 OCR。", file=sys.stderr)

    return ocr_via_tesseract(pdf_path, lang=lang)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = [a for a in sys.argv[1:] if a.startswith("--")]

    if not args:
        print(__doc__)
        sys.exit(1)

    pdf_path = args[0]
    out_path = args[1] if len(args) > 1 else os.path.splitext(pdf_path)[0] + ".txt"
    force = "--force-ocr" in flags

    text = pdf_to_text(pdf_path, force_ocr=force)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"[done] 已写出：{out_path}（{len(text)} 字符）")


if __name__ == "__main__":
    main()
