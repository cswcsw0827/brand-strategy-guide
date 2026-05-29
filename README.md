# 品牌全案策略生成器 (Brand Strategy Guide)

基于《品牌全案策略指南 2.0》的 Claude Code skill，通过深度对话式信息收集，自动生成包含市场调研、战略定位、战略执行的完整品牌全案方案。每个关键策略判断必有真实商业案例佐证。

## 安装

```bash
mkdir -p ~/.claude/skills
git clone https://github.com/DevinKuang/brand-strategy-guide.git ~/.claude/skills/brand-strategy-guide
```

克隆后无需额外配置，Claude Code 会自动加载。

## 三种工作模式

| 模式 | 适合场景 | 输出格式 |
|------|---------|---------|
| **标准对话** | 尚未想清楚，需要顾问引导梳理 | HTML / PPTX |
| **brief 文档** | 已想清楚、有完整素材 | HTML / PPTX（80+ 页深度方案） |
| **MD 输出** | 大型方案，策略与排版解耦 | 结构化 MD brief → 再独立排版 |

### 触发词

「做品牌全案」「品牌策略」「品牌方案」「新品牌怎么做」「品牌战略规划」等。

English: "brand strategy", "brand positioning", "brand plan", "brand blueprint".

## 依赖

### 输出 Skill（至少需要一个）

| Skill | 用途 | 输出 |
|-------|------|------|
| [`guizang-ppt-skill`](https://github.com/) | HTML 网页 PPT（横向翻页，WebGL 背景） | HTML |
| [`ppt-master`](https://github.com/) | AI 多角色协作生成 SVG → PPTX | PPTX |

> 如果以上两个 skill 都未安装，只能走 **MD 输出模式**——生成结构化 Markdown brief 后，单独交给排版工具处理。

### Python 工具（仅 brief 问卷转换）

`scripts/convert_brief_to_docx.py` 用于将 Markdown 格式的 brief 问卷转为 Word 文档。需要：

```bash
pip install python-docx
```

## 目录结构

```
brand-strategy-guide/
├── SKILL.md                    # 主 skill 定义（三种模式 + 完整对话流程）
├── instructions/               # Agent 指令
│   ├── agent_market.md         #   市场调研 agent
│   ├── agent_positioning.md    #   战略定位 agent
│   ├── agent_execution.md      #   战略执行 agent
│   ├── brief_template.md       #   brief 问卷模板
│   └── brief_sample.md         #   brief 问卷示例
├── references/                 # 方法论参考
│   ├── system-overview.md      #   系统全景图
│   ├── output-spec.md          #   输出质量规范
│   ├── methodology.md          #   方法论详解
│   └── theory-library.md       #   理论库
└── scripts/
    └── convert_brief_to_docx.py  # MD → DOCX 转换工具
```

## 许可

MIT License. 本 skill 引用的《品牌全案策略指南 2.0》为作者原创知识体系。
