# Word (.docx) 生成方案调研 —— 投后评价报告自动生产 skill 选型

> 调研日期：2026-07-07
> 场景：资产评估机构「投后评价报告」批量生产。输入 = 财务数据(Excel) + 文字材料；输出 = 中文正式报告，排版讲究（多级中文标题「一、（一）、1.」、页眉页脚、目录、表格、图示、封面、落款签章页、可能二维码）。骨架和口径固定，内容随项目变化，反复批量生产同类报告。

---

## 一、结论先行（可直接定架构）

**主引擎用 `docxtpl`（模板填充范式），辅以 python-docx 做细节兜底、matplotlib/qrcode 出图。不走 pandoc，不走官方 docx skill 的 docx-js 代码生成路子作主力。**

一句话理由：本场景是典型的「骨架/口径固定、内容变化、批量同类」——这正是模板填充范式的主场。让评估师用真实 Word 报告排好版存成 `.docx` 模板（页眉页脚、封面、签章页、多级标题样式、目录域全在模板里一次性做对），Jinja 变量/循环占位，代码只灌数据。排版保真度 = 100%（就是那份真报告本身），远胜任何「代码逐段构建」或「Markdown 转换」的近似还原。

---

## 二、本机环境实查结果

| 组件 | 状态 |
|------|------|
| python3 | ✅ `/usr/bin/python3`，版本 **3.9.6**（偏老，docxtpl/python-docx 均兼容，无碍） |
| python-docx | ✅ 已装 **1.2.0** |
| docxtpl | ❌ 未装（需 `pip3 install docxtpl`，会带上 jinja2 依赖） |
| docxcompose | ❌ 未装 |
| markdown | ❌ 未装 |
| qrcode | ❌ 未装（二维码用它，`pip3 install qrcode[pil]`） |
| matplotlib | ❌ 未装（图表出图用它） |
| **pandoc** | ❌ **未装**（`which pandoc` 无结果）—— 直接影响 pandoc 转换路线和官方 docx skill 的部分能力 |

> 官方 docx skill 已存在于 `~/.claude/skills/docx`：其「创建新文档」走的是 **docx-js（Node，`npm install -g docx`）** 代码生成；「编辑已有」走 unpack XML→改→repack；「读取/转图」依赖 **pandoc + LibreOffice(soffice)**。本机 pandoc 缺失，该 skill 的读取/转换链需先补装。

---

## 三、各方案逐一评估

### 方案 A —— docxtpl（python-docx-template）★ 推荐主引擎

- **定位**：把一份 `.docx` 当 Jinja2 模板用。在 Word 里排好版、插 `{{变量}}`/`{% for %}` 标签，代码填数据生成成品。
- 仓库：https://github.com/elapouya/python-docx-template ｜ PyPI：https://pypi.org/project/docxtpl/
- **star / 维护**：**2.7k star**，最新版 **0.20.2（2025-11-13 发布）**，周下载 ~12.5 万，60+ 贡献者，**活跃维护**。LGPL-2.1。
- **能力边界**：变量替换、条件、循环、循环生成表格行、RichText 富文本（分段设色/加粗）、`InlineImage` 插图片（二维码、matplotlib 图表都走它）、页眉页脚（在模板里画好即可）、subdoc 合并子文档。
- **中文支持**：无中文专属坑——所有中文排版（字体、字号、行距、缩进、多级编号、宋体/仿宋）都在 Word 模板里设好，docxtpl 只负责往占位符里灌文字，不碰样式，所以中文表现 = 你在 Word 里看到的样子。这是它相对「代码生成」最大的优势。
- **适配本场景优劣**：
  - 优：排版保真度天花板（模板即成品）；评估师能自己维护模板不需改代码；批量生产极快；封面/签章页/页眉页脚这类「代码难写、Word 里几分钟搞定」的东西全部外包给模板。
  - 劣：动态结构（章节数量随项目大幅增减、深度不定的嵌套）在 Jinja 里写循环会略绕；纯代码的灵活度不如 python-docx。但投后评价报告骨架固定，这个劣势基本不触发。

### 方案 B —— python-docx（纯代码生成）☆ 辅助/兜底

- 仓库：https://github.com/python-openxml/python-docx ｜ 已装 1.2.0。
- **定位**：逐段/逐表用代码构建 docx。docxtpl 底层就是它。
- star ~5k，维护中（1.2.0 较新）。中文可用但**样式要逐个手写**（字体得同时设 `w:eastAsia` 东亚字体名，否则中文回退默认字体，这是最常见的中文坑）。
- **关键短板**：**没有原生 TOC/目录 API**（GitHub issue #36、#1207 长期未内建），要靠 lxml 手插 TOC 域 XML；也没有图表 API。
- 适配：不适合当主力（排版全靠手写、还原真报告成本高），但适合做 docxtpl 覆盖不到的**局部细节兜底**——比如动态往表格补行、事后处理 XML。**建议角色：配角。**

### 方案 C —— pandoc + reference-doc（Markdown/HTML → docx）☆ 不推荐作主力

- **定位**：LLM 出结构化 Markdown → `pandoc --reference-doc=公司样式.docx` 转成套好样式的 docx。
- **本机 pandoc 未装**，需先 `brew install pandoc`。
- reference-doc 机制：`pandoc -o ref.docx --print-default-data-file reference.docx` 拿到默认样式文档，改里面的样式（**只能改 pandoc 认识的既有样式名，新建样式会被忽略**），之后所有转换套这套样式。
- **中文/排版坑**：
  - 多级中文标题「一、（一）、1.」pandoc 的 heading 只给 Heading1/2/3，中文编号得靠 Word 列表样式在 reference-doc 里配，pandoc 不管编号格式，容易对不齐口径。
  - 封面、签章页、复杂表格合并单元格、精确页眉页脚分节——pandoc 表达力有限，还原不了真报告的讲究排版。
  - 中文字体要在 reference-doc 里设东亚字体，且 pandoc 对 CJK 断行/标点挤压不处理。
- 适配：**保真度不够**。骨架口径要贴合真实报告的硬要求下，pandoc 的「近似还原」满足不了。仅适合「内容为主、排版宽松」的文档，不适合本场景。可作为「读取已有 docx 转 md 喂给 LLM」的**输入侧工具**（但那也需要装 pandoc）。

### 方案 D —— 官方 docx skill（docx-js 代码生成 / XML 编辑）☆ 借鉴架构，不作主力生成

- 位置：`~/.claude/skills/docx`。创建新文档用 **docx-js（Node）** 代码生成；编辑已有文档用 unpack→改 XML→repack；能做 TOC/页眉页脚/letterhead/表格。
- **本质仍是「代码生成」范式**（换成 JS），相对 docxtpl 的「模板填充」，在「还原一份既定真实报告」上同样吃亏——所有排版都得代码复刻。且引入 Node 技术栈，和 Python 数据处理链割裂。
- **可借鉴**：它的 SKILL.md 结构、unpack/repack XML 的编辑手法（做「填充后再微调域/样式」时有用）、accept_changes（LibreOffice 处理修订）等脚本。**当工具箱，不当主生成器。**

### 方案 E —— 中文公文/报告专用现成项目（借模板与规范，不直接接管流程）

- `Drenches/gov-doc-formatter`：基于 LLM agent 的**党政机关公文自动排版**工具（GB/T 9704-2012），上传 Word 自动识别结构套格式。—— 思路可借鉴（「结构识别 + 规范套用」），但它是**公文**规范，非资产评估报告口径，直接用会跑偏。
- `louguanglong/DoOfficialDocument`：提供符合 GB/T 9704-2012 的公文模板 `.dotx` + Word COM 加载项。—— 模板可参考中文正式排版的**字体/行距/编号硬参数**（标题2号方正小标宋、正文3号仿宋_GB2312、固定行距28磅、首行缩进2字、每页22行×28字），移植到我们自己的模板。
- 结论：**这两个不接管我们的生产流程**，只当「中文正式排版参数」的现成参照，抄它们的字体/间距硬值进我们的 docxtpl 模板即可。

### 方案 F —— docxcompose（多文档合并）☆ 按需

- 仓库：https://github.com/4teamwork/docxcompose，维护中。把多个 docx 拼成一个并统一编号/样式。
- 适配：若报告要「封面 + 正文 + 若干附件独立生成后合并」，用它拼装比在单模板里硬塞更干净。**备选工具，非必需。**

---

## 四、图表与表格怎么做

- **财务表 / 评分表**：docxtpl 的**循环生成表格行**（模板里画好一行带 `{% for %}`，代码喂数据列表）。合并单元格、表头样式在模板里预先做好。这是最省力、最保真的做法。
- **图表（柱状 / 趋势）**：docx 原生 chart 无 Python API、极难控。**正解 = matplotlib（或 plotly）出 PNG → docxtpl `InlineImage` 插入**。图的样式完全可控，且和真报告里「贴图表」的实际做法一致。
- **二维码**：`qrcode` 库生成 PNG → `InlineImage` 插入。docxtpl 官方就是这个用法（`width=Mm(...)`）。

---

## 五、每条路必踩的坑（预警）

1. **中文字体回退**（python-docx 纯代码路专属）：设字体必须同时设东亚字体 `rPr/rFonts w:eastAsia`，否则中文变默认字体。→ **走 docxtpl 模板路可完全绕开**（字体在 Word 模板里设死）。
2. **目录（TOC）不自动更新**：docx 里 TOC 是「域(field)」，程序生成后页码/条目是空的，要在 Word 打开按 F9 更新域才显示。
   - 缓解：模板里放好 TOC 域并设「打开时更新域」；或生成后用 LibreOffice headless 跑一遍触发域更新（`soffice` 转 PDF 时会更新）；或用官方 skill 的 XML 手法插 `<w:updateFields>`。**这是所有 Python 方案的共性坑，交付前务必验证目录页码正确。**
3. **页眉页脚 / 分节**：封面无页眉、正文有页眉、章节页码分段——这些「分节符(section break)」用代码写极痛苦。→ **模板里用 Word 分节做好**，docxtpl 保留，一劳永逸。
4. **二维码/内联图 decode 竞态**：这是 HTML 截图路（html-to-image）的坑，docxtpl 走 `InlineImage` 是直接嵌 PNG 到 docx，**无此问题**，放心用。
5. **pandoc 未装**：C 路和官方 skill 读取链要 `brew install pandoc` 才能用；主力若定 docxtpl 则不必装。
6. **python 3.9.6 偏老**：docxtpl/python-docx/qrcode/matplotlib 均支持，但装新库时留意个别库要求 3.10+ 的情况，必要时用 pyenv 升到 3.11。

---

## 六、推荐架构（可直接落地）

```
输入: Excel 财务数据 + 文字材料
  │
  ├─ pandas/openpyxl 读 Excel → 结构化数据 dict
  ├─ LLM 生成分析文字（分析结论、评语）→ 填入 dict 的文本字段
  ├─ matplotlib 出图(柱状/趋势) → PNG；qrcode 出二维码 → PNG
  │
  ▼
docxtpl 渲染:  报告模板.docx (评估师在 Word 里排好的真实报告骨架，
              含封面/多级标题样式/页眉页脚/分节/签章页/TOC域/占位表格)
              + context(dict) + InlineImage(图表/二维码)
  │
  ├─ (可选) python-docx / XML 微调：补动态表格行、触发 updateFields
  ├─ (可选) docxcompose：合并附件
  ├─ LibreOffice headless 跑一遍：更新目录域 + 生成 PDF 交付
  ▼
成品: 投后评价报告.docx / .pdf
```

**分工定论**：
- **docxtpl = 主引擎**（排版保真、批量、可维护）
- **python-docx = 兜底配角**（动态结构微调）
- **matplotlib + qrcode = 出图**（InlineImage 嵌入）
- **LibreOffice/pandoc = 收尾工具**（更新域、转 PDF；按需装）
- **官方 docx skill / 公文项目 = 参考素材库**（借 XML 手法和中文排版硬参数，不接管流程）

**第一步动作**：`pip3 install docxtpl qrcode[pil] matplotlib`，然后拿一份真实投后评价报告改成带 Jinja 占位的 `.docx` 模板——模板质量直接决定成品质量，这是整个 skill 的核心资产。

---

### 来源
- docxtpl: https://github.com/elapouya/python-docx-template ｜ https://pypi.org/project/docxtpl/ ｜ https://docxtpl.readthedocs.io/
- python-docx: https://github.com/python-openxml/python-docx （TOC issue #36 / #1207）
- pandoc reference-doc: https://pandoc-templates.org/ ｜ https://quarto.org/docs/output-formats/ms-word-templates.html
- 中文公文: https://github.com/Drenches/gov-doc-formatter ｜ https://github.com/louguanglong/DoOfficialDocument
- Claude/LLM 报告 skills: https://github.com/anthropics/financial-services-plugins ｜ https://github.com/tfriedel/claude-office-skills ｜ https://github.com/ComposioHQ/awesome-claude-skills
- docxcompose: https://github.com/4teamwork/docxcompose

---

注:本文档技术方案仅覆盖"首次生成报告"阶段;"送审后修订/批注处理"阶段不复用 docxtpl 渲染链,走 python-docx+lxml 手写 XML 的成品 docx 外科手术路线,见 reference/送审后批注处理与交付自检.md。
