# 自动化作品集整理流程

这套流程以 `portfolio_registry.xlsx` 为唯一录入入口，配合本地文件夹完成作品入库、静态网站更新和岗位定制 PDF 输出。

后续在本项目里开启新对话时，可让助手先读取 `AGENTS.md`，它记录了新增作品时应提醒填写的字段和执行规则。

## 文件夹说明

- `portfolio_registry.xlsx`：作品总台账，填写链接、文件名、撰稿类型、发布媒体、图片路径等信息。
- `input/assets/`：放首页照片、作品配图和联系方式图标；首页个人照片固定命名为 `profile.jpg`，作品配图可在台账填写文件名。
- `input/inbox/`：放纸刊原始文档，支持 `.docx`、`.pdf`、`.txt`、`.md`。
- `input/contact_info_template.md`：联系方式页面内容。
- `output/content/`：自动整理后的标准化作品正文。
- `output/summaries/`：用于网站和默认 PDF 的 keyline 与章节精选。
- `output/site/`：生成的静态网站文件。
- `output/pdf/`：生成的 PDF 作品集。
- `work/tmp/`：临时文件、预览图和中间草稿。
- `scripts/portfolio.py`：主操作入口。

## 台账字段

台账固定使用这些列：

`作品ID`、`标题`、`Keyline`、`来源类型`、`链接/文件名`、`撰稿类型`、`主题标签`、`发布媒体`、`发布日期`、`是否公开展示`、`是否为精选作品`、`处理状态`、`备注`、`作品图片`

填写建议：

- `作品ID`：可留空，入库时会自动生成。
- `Keyline`：必填，一句话总结文章核心内容、报道方法和体现的个人能力。
- `来源类型`：填写 `新媒体` 或 `纸刊`。
- `链接/文件名`：新媒体填文章链接；纸刊填 `input/inbox/` 内的文件名。
- `主题标签`：多个标签用中文顿号、逗号或分号分隔，例如 `文旅、人物`。
- `是否公开展示`：填写 `是` 或 `否`。
- `是否为精选作品`：填写 `是` 或 `否`，首页精选作品严格读取这一列；没有勾选时不自动取最新三篇。
- `处理状态`：新作品填 `待处理`，入库后会更新为 `已入库`、`需手动补充` 或 `错误`。
- `作品图片`：可选，填写本地图片文件名、本地路径或图片 URL；如果只填文件名，系统会优先到 `input/assets/` 中查找，并自动兼容 `.jpg`、`.jpeg`、`.png`、`.webp`。

台账格式：

- `作品台账` 第一行冻结，便于滚动查看字段名。
- `来源类型`、`撰稿类型`、`是否公开展示`、`是否为精选作品`、`处理状态` 使用下拉列表。
- 表格边框只使用横线：正文行使用灰色横线，第一行使用上下黑色边框，不设置竖线和表格尾边框。

## 常用操作

所有命令都在当前文件夹运行。

初始化或重建台账模板：

```bash
./portfolio.sh init
```

整理入库：

```bash
./portfolio.sh ingest
```

更新网站：

```bash
./portfolio.sh site
```

生成或检查摘要草稿（新增文章录入后不要直接运行；先在对话中确认 AI 生成的 keyline 与章节精选）：

```bash
./portfolio.sh summaries
```

更新网站并备份到 GitHub：

```bash
./portfolio.sh publish "更新作品集网站"
```

`publish` 会先重新生成 `output/site/`，再把本次变更提交到 Git，并在配置了 GitHub 远程仓库后推送到 GitHub。GitHub 只作为代码和网站产物备份，不负责自动上线。

发布 Vercel 全球版：

```bash
./portfolio.sh deploy-vercel
```

首次运行时会进入 Vercel 登录和项目确认流程；如果不想使用命令行，也可以进入 Vercel Drop，把 `output/site/` 文件夹拖拽上传。

发布腾讯云服务器国内版：

```bash
./portfolio.sh deploy-tencent-server
```

首次使用前需要复制 `.env.tencent-server.example` 为 `.env.tencent-server`，并填写服务器公网 IP、SSH 用户名、端口和远程目录。

安全约束：

- `.env.tencent-server` 只保存在本机，已被 Git 忽略；不要把真实服务器地址、私钥或备案号写入示例文件。
- SSH 私钥不要放在项目文件夹内，建议放在 `~/.ssh/`，并在 `.env.tencent-server` 中只填写路径。
- 自动部署默认禁用密码 SSH 和交互式登录，只允许公钥认证；如果失败，先检查服务器公钥和安全组，不要改成密码部署。

生成摘要版 PDF，例如文旅类岗位：

```bash
./portfolio.sh pdf "文旅类撰稿人岗位"
```

生成摘要版全部作品集，并按撰稿类型分门别类：

```bash
./portfolio.sh pdf "全部作品集，分门别类"
```

生成全文版 PDF：

```bash
./portfolio.sh pdf-full "全部作品集，分门别类"
```

一键完成入库、网站更新和 PDF（适用于已有 summary 已确认的日常更新；新增文章录入时不要用它跳过 AI 确认流程）：

```bash
./portfolio.sh all "文旅类撰稿人岗位"
```

一键完成入库、网站更新、PDF 后，如果需要备份并同步全球版网站，再运行：

```bash
./portfolio.sh publish "更新作品集和网站"
./portfolio.sh deploy-vercel
./portfolio.sh deploy-tencent-server
```

## 网站上线

本项目采用两条线上线：

- GitHub：手动备份代码和 `output/site/` 静态网站产物。
- Vercel：今天先上线全球版。
- 腾讯云：域名注册、实名认证和 ICP 备案并行推进；国内版通过腾讯云 CVM + Nginx 上线。

### GitHub 备份

首次备份需要：

1. 在 GitHub 创建一个公开仓库。
2. 在本地配置远程仓库地址：

```bash
git remote add origin <你的 GitHub 仓库地址>
git push -u origin main
```

之后日常更新可以运行：

```bash
./portfolio.sh publish "更新作品集网站"
```

### Vercel 全球版

今天最快上线方式：

```bash
./portfolio.sh site
npx vercel --prod output/site
```

也可以直接运行：

```bash
./portfolio.sh deploy-vercel
```

首次运行时，按 Vercel 提示登录、确认项目名称和发布设置。发布完成后保存 Vercel 返回的生产地址，例如：

```text
https://your-project.vercel.app
```

如果要使用自定义域名，建议在备案期间先使用 `global.<你的域名>` 指向 Vercel；备案主域名先保留给腾讯云国内版。

### 腾讯云国内版

国内版使用腾讯云服务器承载 `output/site/` 静态网站。默认部署目录为：

```text
/var/www/sufan-portfolio
```

首次部署前，在本机创建服务器配置：

```bash
cp .env.tencent-server.example .env.tencent-server
```

然后填写：

```text
TENCENT_SERVER_HOST=<服务器公网 IP>
TENCENT_SERVER_USER=root
TENCENT_SERVER_PORT=22
TENCENT_SERVER_PATH=/var/www/sufan-portfolio
TENCENT_SERVER_SSH_KEY=<可选，本机 SSH 私钥路径>
ICP_BEIAN_TEXT=
```

服务器需要提前准备：

- 安装 Nginx。
- 安全组放行 `80`、`443`；`22` 端口只允许自己的固定 IP 或可信办公网络访问。
- SSH 建议只允许密钥登录，禁用 root 密码登录；如必须用 root，至少关闭密码登录并限制来源 IP。
- 将 `deploy/nginx/sufan-portfolio.conf` 放到 Nginx 站点配置中，并让 Nginx 加载。
- 运行 `sudo nginx -t` 确认配置无误。

备案中可以先用公网 IP 测试：

```bash
./portfolio.sh deploy-tencent-server
```

备案通过后：

1. 在腾讯云 DNS 中把 `www.sufan-freelancewriter.com.cn` 解析到服务器公网 IP。
2. 可选：把根域名 `sufan-freelancewriter.com.cn` 也解析到同一 IP，并在 Nginx 中跳转到 `www`。
3. 申请并配置 HTTPS 证书，并把 HTTP 自动跳转到 HTTPS；备案完成后不要长期把公网 IP 当作正式入口。
4. 在 `.env.tencent-server` 填写 `ICP_BEIAN_TEXT`。
5. 重新运行 `./portfolio.sh deploy-tencent-server`。
6. 验证首页、作品总览、单篇文章页、CSS 和图片资源都能访问。

公开产物边界：

- 网站只发布 `output/site/` 中的静态文件，不发布 `portfolio_registry.xlsx`、`.env.tencent-server`、`.vercel/` 或 `output/content/`。
- 联系方式页的 PDF 下载只使用最新 `SuFAN全部作品集-摘要版.pdf`；全文版 PDF 不默认复制到网站产物。

日常更新流程：

```
./portfolio.sh ingest
./portfolio.sh publish "更新作品集网站"
./portfolio.sh deploy-vercel
./portfolio.sh deploy-tencent-server
```

## 处理规则

- 新媒体链接会尝试自动提取标题和正文；如果网页限制导致正文不足，会标记为 `需手动补充`，不会猜测内容。
- 从微信公众号提取正文时，会优先读取 `js_content` / `rich_media_content` 正文区域，并过滤图注、服装/品牌露出、作品名+材质、图片来源等非正文内容。
- 从微信公众号提取正文时，`撰稿`、`撰文`、`作者`、`编辑`、`责编` 等编辑信息需要自动读取并保留，例如 `撰文：Susu`、`编辑：杨叶`；如果这些信息在网页正文中较早出现，入库时也会统一移动到文章最后。
- 纸刊文件会从 `input/inbox/` 中读取并抽取正文。
- 入库阶段只生成 `output/content/` 全文归档，并更新台账状态；不会静默写入 `output/summaries/`。
- 入库完成后，需要先由 AI 在对话中生成 keyline 与章节精选候选，展示给用户确认或修改；确认后才创建或更新 summary 文件。
- summary 文件需要人工确认；把 summary 文件中的 `summary_status` 改为 `已确认` 后，网站和摘要版 PDF 才展示该作品的 keyline 与章节精选。
- PDF 生成阶段会根据岗位关键词筛选作品，默认生成摘要版；只有 `pdf-full` 命令或查询中明确包含 `全文版` 时才生成全文版。
- 网站只展示 `是否公开展示` 为 `是` 且 `处理状态` 为 `已入库` 的作品。
- 入库后的 Markdown 文件命名格式为 `发表日期-刊物-文章名.md`，例如 `20260629-男人风尚-在场者廿四载他在南浔筑起一座数万平方米的精神原乡.md`；文章名会去掉标点符号，便于后续手动查找和编辑。

文章录入后的 AI 确认流程：

1. 先运行 `./portfolio.sh ingest`，完成正文提取、清洗、全文归档和台账状态更新。
2. 新文章入库前，台账必须已填写 `Keyline`；缺失时停止入库并标记错误。
3. 入库完成后，AI 读取新生成的正文 Markdown，并参考现有已确认 summaries 的 keyline 与章节精选风格。
4. AI 在对话中展示台账 `Keyline` 和章节精选候选；如果文章没有章节划分，可以生成 `## 精选节选`；如果文章已有章节划分，只保留原文章章节结构，不自动添加 `## 精选节选`，也不创造新的节选分段。
5. 用户确认或修改后，才写入对应 `output/summaries/` 文件；未明确确认状态时，`summary_status` 保持 `待审核`。
6. summary 写入后，再按需运行 `./portfolio.sh site` 或 PDF 命令。

## 网站展示规则

- 首页标题固定为 `Su FAN 作品集`。
- 首页个人简介固定展示：
  - 现为自由撰稿人，base 北京。
  - 中英文流利，能从容应对双语采访与撰稿任务。
  - 教育背景：经济学博士，北京大学（2025）；联合培养博士研究生，巴黎第一大学（2023）；经济学、社会学学士，北京大学（2020）。
- 首页会从已入库作品的 `发布媒体` 字段自动汇总合作媒体名称。
- 如果 `input/assets/profile.jpg` 存在，首页会在合作媒体右侧展示个人照片。
- 首页个人照片尺寸应接近精选作品头图，并固定对齐到首页资料区最右栏。
- 如果台账填写了 `作品图片`，精选作品右侧、作品总览左侧、单篇文章页正文栏左边会展示对应图片。
- 如果作品的 `链接/文件名` 是网址，网页端会在作品标题下方显示 `阅读原文` 超链接，指向原始发布地址。
- 网站图片保持彩色显示，不使用黑白滤镜；图片主体完整显示，背景使用同图模糊铺底来保持版面整齐，不裁切关键信息。
- 图片主体边缘使用淡出过渡，模糊背景也要有边缘柔化，使主图自然衔接到铺底背景。
- 作品分类区标题为 `写作领域`，分类按钮每行最多三个。
- 作品总览中日期和媒体合并显示，例如 `2026-06-29 男人风尚`；没有作品图片的条目也保留图片占位，使标题与有图条目对齐。
- 网站风格为杂志档案风：纯白背景、黑字、现代无衬线字体、少量细线，首页保留“精选作品 + 作品总览”。
- 网页使用统一 type scale：正文 `16px`，`h1` 为 `42px`，`h2` 为 `28px`，标题字号不得超过 `44px`；标题与正文的视觉断层要尽量小。
- 网站只读取 `output/summaries/` 中 `summary_status: 已确认` 的内容，不直接展示 `output/content/` 全文。
- summary 文件使用 `keyline` 字段保存一句话概述，用于快速说明文章核心内容；内容应简洁、有信息密度，不保留“概述”“文章简介”等前缀。
- 章节精选需要保留原文章章节结构；只有文章没有任何章节划分时，才可以添加 `## 精选节选`。如果文章本身已有章节划分，不得自动添加 `## 精选节选`，也不得自行创造新的节选分段。
- 精选作品预览优先使用已确认 summary 的 `keyline`；未确认时只显示“摘要待确认”和阅读原文入口。
- 单篇文章页标题上方保留小字 `Published Work`；文章页右侧顺序为：返回上页、`Published Work`、标题、媒体/日期/写作领域/标签、keyline、分割线、`完整内容请点击 阅读原文，以下为精选节选`、章节精选。
- 单篇文章页使用窄栏长文阅读版，正文行距和段距保持宽松。
- 单篇文章页会识别 Markdown 中的章节标题，例如 `##`、`###`，并按标题级别设置字号。

## PDF 输出规则

- PDF 文件名区分摘要版和全文版，例如 `20260711-SuFAN文旅类作品集-摘要版.pdf`、`20260711-SuFAN全部作品集-全文版.pdf`。
- 当指令包含 `全部`、`全量`、`所有` 或 `完整` 时，导出全部已入库作品。
- PDF 首页标题会标注 `摘要版` 或 `全文版`。
- 当指令包含 `分门别类` 时，目录和正文按撰稿类型分组。
- PDF 首页包含与网页一致的个人简介和合作媒体信息；个人简介分三段展示。
- PDF 首页包含可点击目录，格式为 `媒体名｜标题    页码`；目录页码右对齐，条目链接到对应正文。
- 首页页码使用罗马数字；正文从下一页开始，页码使用阿拉伯数字。
- 摘要版 PDF 每篇作品单独开页，结构为：标题、媒体/日期/写作领域/标签、keyline、`完整内容请点击 阅读原文，以下为精选节选`、章节精选。
- 全文版 PDF 每篇作品单独开页，结构为：标题、媒体/日期/类型、阅读原文链接、正文全文。
- PDF 正文字号关系：文章标题三号字，一级标题小四号字，二级标题和正文五号字。
- 如果作品的 `链接/文件名` 是网址，PDF 正文页会在媒体/日期/类型下方显示可点击的 `阅读原文` 超链接。
