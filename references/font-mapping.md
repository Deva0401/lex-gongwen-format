# 跨平台公文字体映射表

本技能排版标准源自 Windows 中易字库（仿宋_GB2312、华文中宋、楷体_GB2312 等）。这些字体为 Windows 独占商业字体，macOS/Linux 上默认不存在。

为避免排版结果在非 Windows 平台出现缺字、字体回退导致的版式错乱，本表给出三套平台的字体映射。**排版脚本与人工核对均按本表选择字体**，不得在缺失平台硬套 Windows 字体名。

## 字体映射总表

| 公文元素 | Windows（中易字库） | macOS（系统+开源） | Linux（开源 OFL） |
|---|---|---|---|
| 正文中文 | 仿宋_GB2312 | 朱雀仿宋 | 朱雀仿宋 |
| 正文数字/英文 | Times New Roman | Times New Roman | Liberation Serif / Times New Roman |
| 题目（居中大标题） | 华文中宋 | 思源宋体 Bold | 思源宋体 Bold |
| 一级标题 | 黑体 | 思源黑体 | 思源黑体 |
| 二级标题 | 楷体_GB2312 | 霞鹜文楷 | 霞鹜文楷 |
| 三级标题 | 仿宋_GB2312 | 朱雀仿宋 | 朱雀仿宋 |
| 表题/图题 | 仿宋_GB2312（小四加粗） | 朱雀仿宋（小四加粗） | 朱雀仿宋（小四加粗） |
| 架构图字体 | 微软雅黑 | 苹方 / 思源黑体 | 思源黑体 |

## 各平台字体说明

### Windows

中易字库随 Windows 与 Office 预装，无需额外安装。缺失时（精简版系统）可通过安装完整版 Office 或中易字库补丁恢复。

### macOS

| 用途 | 推荐字体 | 字体族名（Font Family） | 许可 |
|---|---|---|---|
| 正文仿宋 | 朱雀仿宋 | Zhuque Fangsong / 朱雀仿宋 | SIL OFL 1.1 |
| 题目宋体 | 思源宋体 | Noto Serif CJK SC | SIL OFL 1.1 |
| 一级标题黑体 | 思源黑体 | Noto Sans CJK SC | SIL OFL 1.1 |
| 二级标题楷体 | 霞鹜文楷 | LXGW WenKai | SIL OFL 1.1 |
| 架构图 | 思源黑体 / 苹方 | Noto Sans CJK SC / PingFang SC | OFL / 系统自带 |

### Linux

| 用途 | 推荐字体 | 字体族名 | 许可 |
|---|---|---|---|
| 正文仿宋 | 朱雀仿宋 | Zhuque Fangsong | SIL OFL 1.1 |
| 题目宋体 | 思源宋体 | Noto Serif CJK SC | SIL OFL 1.1 |
| 一级标题黑体 | 思源黑体 | Noto Sans CJK SC | SIL OFL 1.1 |
| 二级标题楷体 | 霞鹜文楷 | LXGW WenKai | SIL OFL 1.1 |
| 数字/英文 | Liberation Serif | Liberation Serif | SIL OFL 1.1 |

## 开源字体下载源（OFL，可免费商用）

- **朱雀仿宋**（仿宋正文）：https://github.com/TrionesType/zhuque
- **霞鹜文楷**（楷体）：https://github.com/lxgw/LxgwWenKai
- **思源宋体 / 思源黑体**（标题宋体/黑体）：https://github.com/adobe-fonts/source-han-serif 与 https://github.com/adobe-fonts/source-han-sans
- **Liberation Serif**（Times New Roman 替代）：https://github.com/liberationfonts/liberation-fonts

> 说明：上述均为业界广泛使用的开源字体，许可均为 SIL OFL 1.1 或等价开源协议，可免费商用。安装方式：下载 .ttf/.otf 后，macOS 双击「安装字体」入字体册；Linux 复制到 `~/.fonts/` 或 `~/.local/share/fonts/` 后运行 `fc-cache -f`。

## 使用约定

1. 排版脚本（`scripts/docx_format.py`）运行时自动检测平台，按本表选择字体族名。
2. 人工排版时，先确认目标平台，再按对应列选择字体，不得跨平台硬套 Windows 字体名。
3. 若目标平台已安装中易字库（如通过 Wine/CrossOver 运行 Word），则优先使用中易字库以保证与模板完全一致。
4. 最终交付若需跨平台流转，建议在文档中嵌入字体或导出 PDF，避免接收方缺字。
