# Portfolio 项目工作规则

本项目是自动化作品集整理流程。后续在 `/Users/Sue/Documents/Portfolio` 中开启新对话时，优先按这里的规则协助用户；日常操作说明以 `README.md` 为准。

## 固定目标

- `portfolio_registry.xlsx` 是唯一作品台账。
- 新媒体作品通过发布链接入库；纸刊作品先放入 `input/inbox/` 再入库。
- 输出包括静态网站 `output/site/` 和 PDF 作品集 `output/pdf/`。
- 网站只读取 `output/summaries/` 中已确认的 keyline 与章节精选，不公开 `output/content/` 全文。

## 新增作品处理

- 如果用户只发链接、文件名、文章标题或说“帮我入库”，先检查台账是否已有完整信息；缺失时提醒用户补齐，不要主观猜测。
- 新媒体文章至少需要：`来源类型=新媒体`、`链接/文件名`、`撰稿类型`、`主题标签`、`发布媒体`、`发布日期`、`是否公开展示`、`是否为精选作品`、`处理状态=待处理`；`Keyline` 选填。
- 纸刊文章需要先把文件放入 `input/inbox/`，并在台账填写 `来源类型=纸刊`、`链接/文件名=<文件名>`，其余字段同新媒体文章。
- `Keyline` 不阻塞入库；如果台账已填写，summary 阶段沿用人工版本，不改写；如果台账为空，summary 阶段自动生成并写入 summary YAML 的 `keyline` 字段。
- 文章内容提取失败或网页正文不足时，标记为 `需手动补充`，不猜文章内容。

## 入库与摘要确认

- 新增文章录入时先运行 `./portfolio.sh ingest`，不要用 `./portfolio.sh all` 跳过确认流程。
- 入库阶段只生成 `output/content/` 全文归档并更新台账状态，不静默写入 `output/summaries/`。
- 入库后读取新生成的正文 Markdown，优先使用 `$portfolio-summary-style`，参考现有已确认 summaries 的手工摘取风格，在对话中展示 keyline 与章节精选候选；人工已填时标明“沿用台账 Keyline”，未填时标明“AI 生成 Keyline”。
- summary 正文只摘取原文章内容，不加入或重复 keyline；keyline 只保留在 YAML `keyline` 字段中。
- 如果原文已有章节结构，只保留原章节结构，不自动添加 `## 精选节选`，也不创造新分段；没有章节结构时才可添加 `## 精选节选`。
- 用户确认或修改后，才创建或更新 `output/summaries/`；未明确确认状态时，`summary_status` 保持 `待审核`。

## 网站与 PDF 规则

- 网站首页标题固定为 `Su FAN 作品集`，并保持当前杂志档案风、个人简介、合作媒体汇总、`input/assets/profile.jpg` 个人照片和“精选作品 + 作品总览”结构。
- 精选作品严格以台账 `是否为精选作品` 列为准；不要用标签或备注中的“精选”字样判断，也不要在没有勾选时自动取最新三篇。
- 网站单篇页只展示 `summary_status: 已确认` 的 keyline 与章节精选；未确认时只显示基础信息、阅读原文和“摘要待确认”，不得输出 `output/content/` 全文。
- 默认 `./portfolio.sh pdf "..."` 生成摘要版；只有 `./portfolio.sh pdf-full "..."` 或查询中明确包含 `全文版` 时才读取 `output/content/` 生成全文版。
- 生成某个岗位方向的 PDF 时，先确认或使用用户给出的岗位/主题关键词作为筛选条件。

## 部署与安全边界

- 只发布 `output/site/`，不发布 `portfolio_registry.xlsx`、`.env.tencent-server`、`.vercel/` 或 `output/content/`。
- 摘要 PDF 是公开产物；全文 PDF 不默认复制到网站产物。
- `.env.tencent-server` 只保存在本机，真实服务器地址、私钥和备案信息不要写入示例文件。
- 腾讯云部署路径是保留并加固的支持路径，不要默认移除。
- Vercel 生产地址默认按 `https://sufan-portfolio.vercel.app` 处理，除非用户说明已变化。

## 常用命令

用户明确要求执行时，在项目根目录运行对应命令，并反馈结果：

```bash
./portfolio.sh ingest
./portfolio.sh summaries
./portfolio.sh site
./portfolio.sh pdf "文旅类撰稿人岗位"
./portfolio.sh pdf "全部作品集，分门别类"
./portfolio.sh pdf-full "全部作品集，分门别类"
./portfolio.sh all "文旅类撰稿人岗位"
./portfolio.sh publish "更新作品集网站"
./portfolio.sh deploy-vercel
./portfolio.sh deploy-tencent-server
```

## 执行后检查

- 入库后检查台账状态是否变成 `已入库`、`需手动补充` 或 `错误`。
- 入库后检查 `output/content/` 是否生成对应 `.md`。
- 新入库作品必须等待用户确认 summary，不能静默写入。
- 更新网站后检查 `output/site/index.html` 是否存在。
- 生成 PDF 后检查 `output/pdf/` 是否存在对应文件。
- 只改和当前任务直接相关的内容；发现其他问题可以文字提醒，不要擅自修改。
