# Portfolio 安全审计记录

审计日期：2026-07-13

## 结论

- 未在已跟踪文本文件中发现明显 API token、私钥、密码或 Vercel/GitHub/Tencent 密钥。
- `.env.tencent-server` 和 `.vercel/` 保持本地私密状态，不进入 Git。
- Vercel 项目当前无环境变量，生产部署状态为 Ready，正式别名包含 `https://sufan-portfolio.vercel.app`。
- Git 当前提交身份已使用 `sue.fan1998@gmail.com`；历史提交中仍存在旧学校邮箱元数据，不建议为此重写历史。
- `output/site/assets/20260712-SuFAN全部作品集-摘要版.pdf` 是预期公开产物，继续保留。

## 命中项分类

| 类别 | 结果 | 处理 |
| --- | --- | --- |
| 真实密钥 / token / 私钥 | 未发现 | 无需轮换 |
| Vercel 本地项目元数据 | `.vercel/project.json` 仅本地存在，已被忽略 | 不提交 `.vercel/` |
| 腾讯云环境文件 | `.env.tencent-server` 仅本地存在，已被忽略 | 不提交真实服务器配置 |
| 账号标识 | 当前 Git email 为 `sue.fan1998@gmail.com`；历史提交含旧学校邮箱 | 不重写历史，未来提交继续使用 Gmail |
| 服务器配置 | 腾讯云 Nginx 配置包含域名和临时公网 IP 入口 | 公网 IP 仅用于备案前测试，备案后使用域名 + HTTPS |
| 公开内容 | 网站产物包含摘要版 PDF | 视为预期公开作品集，不公开全文版 PDF |

## 已加固项

- `.gitignore` 已覆盖真实环境文件、常见 SSH 私钥和 pem/key 文件，并保留 `.env*.example` 可提交。
- 腾讯云部署脚本默认禁用密码 SSH 和交互式登录，只允许公钥认证。
- 腾讯云部署脚本会拒绝项目目录内的 SSH 私钥，避免误提交。
- README 已补充腾讯云部署安全基线：SSH key、限制 22 端口来源、备案后 HTTPS、公开产物边界。
- Nginx 配置已标注公网 IP 只用于备案前测试。

## 后续上线检查

- 每次提交前运行敏感信息扫描，确认没有 `.env`、`.vercel/`、私钥、token、台账文件进入 Git。
- 腾讯云正式上线前，在安全组中限制 `22` 端口来源，并确认服务器禁用密码登录。
- 备案和证书完成后，把 HTTP 跳转到 HTTPS，不长期使用公网 IP 作为正式入口。
- 若任何真实密钥曾进入 Git 或公开产物，先轮换密钥，再清理仓库和线上产物。
