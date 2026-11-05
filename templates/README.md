# 模板包 / Template Pack

开箱即用的 PPTX 模板，下载仓库后直接可用：**全部保留原生母版与版式，文字仍在占位符里，可逐字编辑**。

```
templates/<slug>/
├── template.pptx   空母版：59 套版式，从零挑选版式（作者模式）
├── starter.pptx    9 页示例稿：内容齐全、颜色统一，改文字即可用（推荐入口）
├── previews/       示例稿逐页 PNG（slide-01..09）
├── preview.png     封面预览
├── LICENSE         MIT（含上游版权声明）
└── layouts.json    版式语义表：visual type → layout index（仅 inner-chapter）
```

## 目录

| slug | 风格 | 适用 | 版式 |
|---|---|---|---|
| `inner-chapter` | 黑白极简 + 蓝橙强调色 | 咨询汇报、评审、产品复盘 | 59 |
| `inner-chapter-navy` | 深蓝版（微软雅黑） | 技术方案、课程汇报、企业对外 | 59 |
| `inner-chapter-ink` | 暖纸墨金（思源黑体 CN） | 人文/建筑/材料类学术汇报、编辑风 | 59 |

三个模板同一套母版，区别在配色与中文字体；换主题不必重排版式。

## 用法

最快路径（推荐）：

```bash
# 1) 从模板脚手架出一份自己的稿子
python scripts/new_deck.py 我的汇报 --template inner-chapter-navy

# 2) 改 decks/我的汇报/content.json（map：模板原文 → 你的文字）
# 3) 构建 + QA + 逐页 PNG
python src/run_workflow.py --config decks/我的汇报/deck.config.json
```

只想看看效果：直接用 PowerPoint 打开 `templates/*/starter.pptx`，或看 `previews/`。

> `new_deck.py --blank` 会从 `template.pptx` 起步，适合自己按版式写内容；
> 默认用 `starter.pptx`，因为“从已有 9 页改文字”比“从空母版造 9 页”快得多。

## 版式怎么挑

`inner-chapter/layouts.json` 把 59 套版式按用途编了号（`title_cover`、`column_3`、`grid_2x2_2body`、`contact` 等）。
按语义选版式，别按页码选：

- 开场/结尾：`title_cover` → `title_centered` → `contact_black`
- 并列 2/3/4/5 项：`column_2` / `column_3` / `column_4` / `column_5`
- 图文混排：`content_with_image`、`story_card`、`grid_2x2_2body`
- 纯论述：`content_centered`

## 三套模板怎么轮换（别整册只用一套）

同一份 PPT 里不要连续 3 页用同一版式，也不要让一套模板撑完所有页 ——
`inner-chapter` / `-navy` / `-ink` 共享 59 套版式，正好用来换气口：

- 开场用 `title-cover`，第一章结束用 `title-centered`，不要每章都重复封面构图；
- 并列内容在 `column-3` 与 `column-3-centered-a`、`content-image-right-a` 与 `-b` 之间轮换；
- 全册至少出现 4 种不同版式，单一版式不超过 40%；
- 换主题（navy ↔ ink）能明显拉开视觉差异，且标题层级、页眉页脚、插图语言仍然一致。

构建后跑 `python src/check_layout_diversity.py <deck.pptx>` 自检，`REVIEW` 就回去改版式，不是改文字。

## 授权

- `inner-chapter` 来自 [tristan-mcinnis/pptx-from-layouts-skill](https://github.com/tristan-mcinnis/pptx-from-layouts-skill)，MIT，版权归原作者（见 `LICENSE`）。
- `inner-chapter-navy` / `inner-chapter-ink` 是同一母版的换色派生版，同样按 MIT 分发，需保留上游版权声明。
- `starter.pptx` 的示例文字与配图为本项目自制（MIT）；示例配图是抽象色块，不涉及第三方素材。

MIT 允许商用、修改、再分发，**保留版权与许可声明即可**。
若需要更多风格，可自行从 Microsoft Create / Slidesgo / SlidesCarnival / Canva 取免费模板，
但请先确认其授权允许你要的用途（多数要求署名，且不允转售模板本身），本仓库不再分发它们。
