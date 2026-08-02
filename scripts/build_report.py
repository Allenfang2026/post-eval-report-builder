#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_report.py —— docxtpl 渲染：模板 docx + data.json + 插图 → 成品报告 docx

这是最后一步：把前面读好的数据（read_inputs.py 产的 data.json，再经 AI 填充确定项、
留空判断项后的完整 context）灌进 Word 模板（templates/ 下的 .docx），渲染出成品报告。

设计原则（对应 SKILL.md 的核心原则）：
    - 排版全在模板里做死，本脚本只负责"灌数据"，不碰任何字体/样式。
    - 图（投资控制对比图/财务趋势图/防伪二维码）用 docxtpl 的 InlineImage 插入。
    - 表格用模板里预先画好的一行 + {% for %} 循环由 context 的列表喂数据。

★★★ TODO：模板文件还没做 ★★★
    templates/ 目前只有 README.md，实际的 股权类模板.docx / 工程类模板.docx 需要下一步
    拿真实报告在 Word 里改成带 {{变量}} 和 {%for%} 的模板。做法见 templates/README.md。
    模板做好后，把下面 TEMPLATE_PATHS 里的路径指向真实模板即可跑通。

依赖（本机已装）：docxtpl（含 jinja2）。插图用到 make_charts.py（同目录）。

用法（模板就绪后）：
    python3 build_report.py context.json 成品报告.docx
      context.json = data.json 经 AI 填充后的完整渲染上下文
      （里含所有确定项字段、留空判断项的占位提示、要插的图片路径列表）

    也可指定报告类型强制选模板：
    python3 build_report.py context.json 成品.docx --type engineering
"""

import os
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# ★ TODO：模板做好后把这里指向真实模板文件
_SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATHS = {
    "equity":      os.path.join(_SKILL_DIR, "templates", "股权类模板.docx"),
    "engineering": os.path.join(_SKILL_DIR, "templates", "工程类模板.docx"),
}


def load_context(context_json):
    """读 AI 填充好的渲染上下文。"""
    with open(context_json, "r", encoding="utf-8") as f:
        return json.load(f)


def pick_template(context, force_type=None):
    """
    选模板：优先命令行 --type，其次 context 里的 report_type。
    report_type: "equity"(股权) / "engineering"(工程)
    """
    rtype = force_type or context.get("report_type")
    if rtype not in TEMPLATE_PATHS:
        raise ValueError(
            f"report_type 必须是 equity 或 engineering，当前为 {rtype!r}。"
            f"请在 context.json 里填好 report_type，或用 --type 指定。"
        )
    tpl = TEMPLATE_PATHS[rtype]
    if not os.path.isfile(tpl):
        raise FileNotFoundError(
            f"模板文件不存在：{tpl}\n"
            f"★ 模板还没做。请按 templates/README.md 用真实报告制作 {os.path.basename(tpl)}。"
        )
    return tpl, rtype


def build(context_json, out_docx, force_type=None):
    from docxtpl import DocxTemplate, InlineImage
    from docx.shared import Mm

    context = load_context(context_json)
    tpl_path, rtype = pick_template(context, force_type)

    doc = DocxTemplate(tpl_path)

    # ---- 处理图片：context["images"] 里每项 {"var":"投资控制图","path":"...","width_mm":150} ----
    # 模板里用 {{ 投资控制图 }} 引用。图片要包成 InlineImage 才能插入。
    # 先取模板声明的变量集，用来检测哪些图的占位符模板里根本没有（否则会被静默丢弃）。
    try:
        declared_vars = doc.get_undeclared_template_variables()
    except Exception:
        declared_vars = None  # 拿不到就退化为不检测，不影响渲染
    images = context.pop("images", []) or []
    for img in images:
        path = img["path"]
        var = img["var"]
        if not os.path.isfile(path):
            print(f"[warn] 图片不存在，跳过：{path}", file=sys.stderr)
            continue
        if declared_vars is not None and var not in declared_vars:
            print(f"[warn] 图片 {os.path.basename(path)} 的占位符 {{{{ {var} }}}} 不在模板里，未插入",
                  file=sys.stderr)
            continue
        width = Mm(img.get("width_mm", 150))
        context[var] = InlineImage(doc, path, width=width)

    # ---- 渲染 ----
    # context 里应包含：所有确定项字段、逐字固定话术、表格循环用的列表、
    # 以及"待评估师判断"的占位提示字符串（AI 已按 reference 话术填好占位文案）。
    doc.render(context)
    doc.save(out_docx)

    print(f"[done] 已生成（{rtype}）：{out_docx}")
    print("       提醒：")
    print("       1) Word 打开后按 F9 更新目录域（页码才正确）；")
    print("       2) 检查所有【待评估师判断：…】占位处，补齐风险定档/成功度打分/评价结论；")
    print("       3) 数据缺失处应显示为「待补充」，确认没有被编造的数字。")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    force_type = None
    if "--type" in sys.argv:
        force_type = sys.argv[sys.argv.index("--type") + 1]

    if len(args) < 2:
        print(__doc__)
        sys.exit(1)

    build(args[0], args[1], force_type=force_type)


if __name__ == "__main__":
    main()
