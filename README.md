# lex-hj公文排版

国企/体制内公文排版格式化技能，内置股权投资议案模板的视觉格式标准。

## 适用场景

- 股权投资议案、请示、报告、会议材料、上会材料等公文的排版
- 公文格式合规性检查
- 标题层级、表格、图片架构图样式规范化

## 安装

```bash
npx skills add Deva0401/lex-gongwen-format
```

或直接将本目录复制到你的 skills 目录下。

## 前置条件

- [ ] 已安装 Node.js 与 npx：`node --version && npx --version`
- [ ] 已安装 Python 3：`python3 --version`
- [ ] 处理 .docx 文件时已安装 python-docx：`pip install python-docx`
- [ ] 非 Windows 平台已安装对应开源字体（见下「跨平台字体」）

## 你可以直接这样说

- "把这份议案排版成规范的公文格式"
- "检查一下这份文档的字体字号行距是否符合公文要求"
- "按议案模板的格式调整这个文档"
- "公文里的表格和图注格式帮我统一一下"
- "这份请示的标题层级格式对不对"

## 格式标准摘要

| 元素 | 格式 |
|---|---|
| 页边距 | 上 2.54cm、下 2.54cm、左 3.18cm、右 3.18cm |
| 正文 | 仿宋_GB2312 三号，数字英文 Times New Roman，固定行距30磅，首行缩进2字符 |
| 一级标题 | 华文中宋 二号 加粗 居中 |
| 二级标题 | 黑体 三号 |
| 三级标题 | 楷体_GB2312 三号 加粗 |
| 四级标题 | 仿宋_GB2312 三号 |
| 表题 | 仿宋_GB2312 小四 加粗（表上） |
| 表格正文 | 仿宋_GB2312 小四，行距20磅 |
| 图题 | 仿宋_GB2312 小四 加粗（图下） |
| 架构图 | 黑色边框/文字，微软雅黑，可编辑对象（禁止截图） |

完整规范见 `references/format-spec.md`；跨平台字体映射见 `references/font-mapping.md`。

## 跨平台字体

Windows 使用中易字库（仿宋_GB2312 / 华文中宋 / 楷体_GB2312），随 Office 预装无需额外安装。macOS / Linux 默认无这些字体，需改用开源 OFL 字体，否则会缺字、字体回退导致版式错乱：

| 用途 | Windows | macOS / Linux |
|---|---|---|
| 正文仿宋 | 仿宋_GB2312 | 朱雀仿宋（Zhuque Fangsong） |
| 一级标题宋体 | 华文中宋 | 思源宋体（Noto Serif CJK SC） |
| 二级标题黑体 | 黑体 | 思源黑体（Noto Sans CJK SC） |
| 三级标题楷体 | 楷体_GB2312 | 霞鹜文楷（LXGW WenKai） |

完整映射与下载源见 `references/font-mapping.md`。内置脚本 `scripts/docx_format.py` 会自动按平台选择字体。

## 快速开始

```bash
pip install python-docx
python scripts/docx_format.py 你的公文.docx --platform auto
```

## 验证

```bash
python3 scripts/validate_skill.py .
```

## Troubleshooting

- **字体缺失（macOS/Linux）**：仿宋_GB2312、华文中宋、楷体_GB2312 为 Windows 中易字库，非 Windows 平台需按 `references/font-mapping.md` 安装开源替代字体（朱雀仿宋 / 思源宋体 / 思源黑体 / 霞鹜文楷），或使用 `scripts/docx_format.py --platform` 自动切换。直接硬套 Windows 字体名会导致缺字。
- **python-docx 未安装**：运行 `scripts/docx_format.py` 前需 `pip install python-docx`；脚本会在依赖缺失时给出明确提示，不会静默失败。
- **格式未生效**：检查文档是否有手动设置的直接格式覆盖了样式；建议清除手动格式后重新应用样式。
- **表格行距异常**：表格正文行距为20磅，与正文30磅不同，需单独设置表格样式。
- **页边距不对**：模板要求上 2.54cm、下 2.54cm、左 3.18cm、右 3.18cm；脚本会自动应用，人工排版在 Word「布局 → 页边距 → 自定义页边距」设置。

## 来源

格式标准提取自《议案模板（股权投资-存续经营类项目）-征求意见》。

## License

MIT

Copyright (c) 2026 Lex (Deva0401)
