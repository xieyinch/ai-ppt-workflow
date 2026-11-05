# AI PPT 工作流（模板驱动）

不要用代码自由画版式。流程是：**找模板 → 按模板填内容 → 校验 → 交付**。

## 目录
```
ai-ppt-workflow/
├── deck.config.json          # 一次运行的全部参数（主题/模板/输出/质量门槛）
├── templates/                # 预制模板包（MIT，下载即可用）+ 你自己的候选模板
│   ├── manifest.json         # 模板清单
│   ├── inner-chapter/        # 59 套版式的原生模板 + starter.pptx 示例稿
│   ├── inner-chapter-navy/   # 同母版深蓝换色版
│   └── inner-chapter-ink/    # 同母版暖纸墨金版
├── decks/<名字>/             # new_deck.py 生成的稿件工作区
│   ├── content.json          # 模板原文 → 新内容替换表
│   └── deck.config.json      # 指向本稿件的模板/输出/渲染目录
├── input/*-content-text.json # 手工维护的替换表（不走 new_deck 时）
├── research/sources.md       # 来源清单（先查资料，别想当然）
├── output/                   # 成品 pptx + 讲稿
├── scripts/
│   ├── new_deck.py           # 从模板脚手架出一份稿子
│   └── render_previews.ps1   # 逐页导出 PNG
└── src/
    ├── run_workflow.py       # 稳定的一键入口
    ├── build.py              # PowerPoint COM 填充、重开校验、PNG 导出
    └── qa.py                 # 结构校验（页数/字数/关键词/媒体/文件完整性）
```

## 运行
```bash
cd <project-root>
python src/run_workflow.py --config deck.config.json
# 可选：同时导出浏览器预览
python src/run_workflow.py --config deck.config.json --html-preview
```
退出码 0 = 生成、结构 QA、渲染计数和资料检查全部通过；非 0 会打印具体失败项。

## 预制模板包（默认起点）

仓库自带三套同母版、可直接用的模板，通常不必再去外面找：

```bash
python scripts/new_deck.py 我的汇报 --template inner-chapter-navy
```

| slug | 风格 | 适用 |
|---|---|---|
| `inner-chapter` | 黑白极简 + 蓝橙强调 | 咨询汇报、评审、复盘 |
| `inner-chapter-navy` | 深蓝 + 微软雅黑 | 技术方案、课程汇报、企业对外 |
| `inner-chapter-ink` | 暖纸墨金 + 思源黑体 CN | 材料/建筑/人文类学术汇报 |

每套都含 `template.pptx`（空母版，59 套版式）、`starter.pptx`（9 页示例稿，改字即用）、
`previews/`（逐页 PNG）和 `layouts.json`（版式语义表，仅 inner-chapter）。授权见 `templates/ATTRIBUTION.md`。

模板不够用时才去外部找（GitHub、Microsoft Create、Slidesgo、SlidesCarnival、Canva、学校模板），
**先看授权**：允许署名使用 ≠ 允许在仓库里再分发模板文件。

## 统一后端选择

工作流不再把所有 PPT 当成同一种文件，而是先选择产物路线：

| 用户目标 | 路由 | 产物 |
|---|---|---|
| 答辩、汇报、交付、需要编辑 | `native` | 原生可编辑 PPTX |
| 杂志风、演讲模式、网页展示 | `guizang` | HTML 演示稿 |
| 高保真 HTML 且愿意遵守转换约束 | `huashu` | HTML，或部分可编辑 PPTX |
| 只追求视觉保真 | `image` | 图片型 PPTX |

统一调度入口：

```bash
python src/make_ppt.py --route auto --config deck.config.json
```

所有路由都写入 `output/last-route.json`，记录实际后端、产物类型和成功状态。禁止把 HTML 或图片型 PPTX 宣称为逐元素可编辑文件。

## 六步流程
1. **定场景**：课程分享 / 毕设答辩 / 汇报 / 路演，确定时长与页数。
2. **定模板**：先用 `templates/` 里的三套预制模板；不够再找外部模板，并记录授权来源。
3. **定大纲**：写 `input/blueprint.json`，先有叙事再填字。
4. **语义版式规划**：为每页确定 page job 和 visual type，再从 `layouts.json` 中选候选版式；结果保存为 `*-layout-plan.json`。
5. **填内容**：把模板原文与新内容的对应关系写进 `content.json`，改 `deck.config.json` 指向选定模板，运行 `run_workflow.py`。
6. **校验交付**：统一入口会生成 PPTX、讲稿、来源、渲染 PNG 和 contact sheet；`--html-preview` 只生成浏览器审阅版，不替代原生 PPTX。

## Agent 设计路由

`src/design_system.json` 提供四套可复用的视觉系统：

- `editorial`：杂志叙事、低密度、图片主导；
- `swiss`：网格、直角、单一锚点色、事实表达；
- `cyberpunk`：深色底、霓虹锚点、短标签、有限装饰；
- `corporate`：证据、图表、路线图和稳定的商务层级。

内容先形成页面蓝图，再选择模板或风格系统，最后通过 PowerPoint COM 填充。这样把 PPT Master 的原生对象路线、归藏/花叔的设计系统约束、宝玉的视觉素材流程、Frontend Slides/Lewis 的浏览器预览思路拆成可验证的独立阶段，而不是把 HTML 截图冒充成可编辑 PPTX。

## 换模板怎么做
`python scripts/new_deck.py <名字> --template <slug>` 会自动复制模板、扫出每页可编辑文字槽，
生成 `content.json` 与 `deck.config.json`。要手工换模板就复制这两份文件，
只改 `template`、`content`、`output`、`render_dir` 四项，并重建内容映射（键 = 新模板里的原文）。
同一占位符在多个卡片中重复时，把值写成数组，程序会按模板顺序逐个填入。程序入口和 QA 不用动。

## 图片与图标替换
在页面条目中增加 `images` 数组即可替换模板图片：
```json
{
  "src": 26,
  "map": {},
  "images": [
    {"occurrence": 2, "file": "assets/nylon_leg.jpg"},
    {"kind": "group", "occurrence": 1, "file": "assets/icon.png"}
  ]
}
```
默认按页面内图片出现顺序替换；`kind` 可用 `picture`、`group` 或 `icon`。也可以使用 PowerPoint 形状名：`{"name": "Picture 114", "file": "..."}`。程序会保留原位置和尺寸，并在重新打开 PPT 后导出 PNG 验证。

**图片必须服务于该页内容**：每张图都要能说清它证明了什么、说明了哪一步。
为了填满版式而放无关图、装饰图或库存图，等于把版面做假，宁可留白或换成表格。

## 原生图表数据替换
页面条目可增加 `charts`，按原生 PowerPoint 图表的名称或出现顺序更新分类和系列：
```json
{
  "src": 24,
  "charts": [
    {
      "occurrence": 1,
      "categories": ["A", "B", "C"],
      "series": {"PA6": [8, 16, 24], "PA6,6": [10, 20, 30]}
    }
  ]
}
```
图表会先在模板副本中更新，再交给 PowerPoint 生成最终文件，避免直接改 PPTX 包。

## 质量门槛（deck.config.json → quality）
- `min_slides`：最少页数
- `min_text_chars`：最少正文汉字数（防“漂亮但空”）
- `required_terms`：必须出现的关键知识点（防跑题、防缺章节）

## 当前样例
预制模板包自带 9 页示例稿：`templates/inner-chapter-navy/starter.pptx`
（封面 / 章节页 / 图文页 / 三栏 / 双栏对比 / 2×2 图文 / 三层结构 / 联系页 / 深色结尾），
逐页效果见同目录 `previews/`。

## 已知限制
- 当前主流程依赖 Windows PowerPoint COM；本机已具备并已实跑逐页 PNG 导出。
- 模板中的示意图、图表数据不会被文字映射自动换成主题相关内容，需要选用匹配模板，或在第二遍流程中单独替换。
- 映射键必须与模板中的原文逐字一致；未命中的键需要重新检查模板文本。
- 换色版只替换主题色板与字体方案，母版里写死的 `srgbClr` 装饰色不会跟着变；要彻底改色需改母版。
- 旧的代码画版式方案（`src/generate.mjs` + `blueprint*.json` + `npm run build`）已停用，仅作历史记录保留。
