# AI PPT Workflow

模板驱动、原生可编辑、可验证的 PowerPoint 生成工作流。

它不是“从零画一套卡片”的 PPT 生成器，而是：

```text
选择真实模板 → 研究内容 → 映射文字 → 替换图片/图标/图表 → PowerPoint 重写 → 重新打开验证 → PNG 复核
```

## 特点

- Windows PowerPoint COM 写入，保留模板母版和可编辑性
- 重复占位符可用数组逐个填充
- 支持图片、图标和原生图表数据替换
- 自动重新打开生成的 PPTX
- 自动导出每页 PNG
- 结构 QA：页数、占位符、关键词、讲稿、来源和媒体
- 不绑定任何私有模板；用户可以提供自己的 `.pptx` / `.potx`
- 新增语义版式规划、统一路由、contact sheet 和可选 HTML 预览
- 每次运行记录实际后端和产物可编辑性，避免把 HTML 或图片型 PPTX 冒充为原生可编辑

## 环境

- Windows 10+
- Microsoft PowerPoint（用于生成和 PNG 导出）
- Python 3.10+
- Python 依赖：`pip install -r requirements.txt`

## 快速开始

1. 准备一个可编辑的 PowerPoint 模板，放入 `templates/`。
2. 复制 `examples/deck.config.example.json` 为 `deck.config.json`。
3. 按模板中的真实占位符编写内容映射 JSON。
4. 运行：

```bash
python src/run_workflow.py --config deck.config.json
python src/qa.py deck.config.json
```

生成结果由配置文件中的 `output` 和 `render_dir` 指定。

## 内容映射

普通文字替换：

```json
{"map": {"模板原文": "新的内容"}}
```

同一占位符重复出现在多个卡片中时使用数组：

```json
{"map": {"这里是小标题": ["第一栏", "第二栏", "第三栏"]}}
```

图片或图标替换：

```json
{"images": [{"kind": "picture", "occurrence": 1, "file": "assets/example.png"}]}
```

原生图表数据替换：

```json
{"charts": [{"occurrence": 1, "categories": ["A", "B"], "series": {"系列1": [10, 20]}}]}
```

## 公开发布范围

本仓库只发布工作流代码、文档、示例和自制资源，不包含任何用户私有模板、夸克网盘文件、个人资料、生成成品、聊天记录、浏览器数据或运行缓存。模板版权和第三方素材授权由使用者自行确认。

## 许可证

代码采用 MIT License。模板、示例素材和第三方文件遵循各自许可证。
