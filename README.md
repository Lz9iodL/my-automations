# my-automations · 个人定时推送自动化仓库

用 **GitHub Actions** 在云端按计划跑脚本，把「抓取 / 生成」的内容**自动推送到你指定的渠道**（邮箱等）。
**免费、无需服务器、你电脑关机也照常执行。**

> 本仓库的定位是**通用容器**：一个任务 = 一个脚本 + 一个 workflow，互不干扰，可随时扩展。
> 下面「任务一」是当前已实现的；以后的每日/每周推送按同样套路加即可。

| 任务 | 脚本 | 工作流 | 频率 | 推送渠道 |
| --- | --- | --- | --- | --- |
| ① CET-4 每日精读 | `main.py` | `.github/workflows/cet4-daily.yml` | 每天 07:00（北京） | 邮件 → `2026200703@ruc.edu.cn` |
| ② （待添加） | — | — | — | — |

---

## 一、首次配置（3 步）

### 1. 配置密钥（Secrets）

仓库 → **Settings → Secrets and variables → Actions → New repository secret**：

| Name | 用途 | 当前任务是否需要 |
| --- | --- | --- |
| `LLM_API_KEY` | 大模型 API Key（DeepSeek） | ✅ 必需 |
| `SMTP_USER` | 发信邮箱，如 `1553895355@qq.com` | ✅ 必需 |
| `SMTP_PASS` | 邮箱 SMTP 授权码（16 位） | ✅ 必需 |
| `MAIL_TO` | 收件邮箱（可选，默认 RUC 邮箱） | 可选 |
| `LLM_MODEL` | 模型名（可选，默认 `deepseek-chat`） | 可选 |
| `SMTP_HOST` | SMTP 服务器（可选，默认 `smtp.qq.com`） | 可选 |

### 2. 启用 Actions

仓库 → **Actions** 标签 → 若出现提示，点 **I understand my workflows, go ahead and enable them**。

### 3. 手动测试一次

Actions → 左侧 **CET-4 每日精读** → **Run workflow** → 等 1~3 分钟，看是否绿色 ✓，然后去邮箱查收（含垃圾箱）。

测试通过后无需再操作，**每天 07:00 自动执行**。

---

## 二、任务一：CET-4 每日精读（`main.py`）

**做什么**：抓取 China Daily 当天英文文章（优先**中华文化 / 科技前沿**，自动排除时政、涉台、军事、灾祸等题材）→ 调用大模型生成精读 Markdown → 生成 Word（.docx）→ 通过 SMTP 发到指定邮箱。

**文档结构**（五部分）：① 文章信息 ② 原文（生词 ①—㉚ 加粗标注、语法短语加下划线并标【1】—【6】）③ 30 个高频生词（音标/词性/释义/例句/固定搭配/词形变化）④ 重点语法与长难句（注明「第几段第几句」）⑤ 学习建议。

**发送前自检**：自动核对生词数是否为 30、①—㉚ 与【1】—【6】是否齐全、标题是否完整；不通过会重新生成一次。

**可选环境变量**：
- `LLM_BASE_URL`：默认 `https://api.deepseek.com/v1`
- `LLM_MODEL`：默认 `deepseek-chat`
- `SMTP_HOST` / `SMTP_PORT`：默认 `smtp.qq.com` / `465`
- `MAIL_TO`：默认 `2026200703@ruc.edu.cn`

**本地调试**：
```bash
set LLM_API_KEY=sk-xxx
set SMTP_USER=xxx@qq.com
set SMTP_PASS=xxxxxxxxxxxxxxxx
python main.py
```
（纯 Python 标准库，无需安装依赖。）

---

## 三、以后怎么加新任务（每天 / 每周 / 每月）

**三步**：

1. **写脚本**，例如 `tasks/weekly_example.py`（读环境变量、跑完退出即可）；
2. **加 workflow**，例如 `.github/workflows/weekly-example.yml`，复制 `cet4-daily.yml` 改两处：
   - `cron:` 改成你要的时间（**UTC = 北京时间 − 8 小时**）
   - `run:` 改成 `python tasks/weekly_example.py`
3. 提交，去 Actions 手动跑一次验证。

**Secrets 全仓库共用**，不用重复配置。

常用 cron 对照：

| 需求 | 北京时间 | cron（UTC） |
| --- | --- | --- |
| 每天 07:00 | 07:00 | `0 23 * * *` |
| 每天 08:00 | 08:00 | `0 0 * * *` |
| 每周一 09:00 | 周一 09:00 | `0 1 * * 1` |
| 每月 1 日 09:00 | 1 日 09:00 | `0 1 1 * *` |

---

## 四、说明与注意事项

- **免费额度**：私有仓库每月 2000 分钟 Actions，本任务每天约 2–3 分钟，绰绰有余；公开仓库不限时。
- **时间延迟**：GitHub 定时任务在高峰期可能延迟几分钟~几十分钟，属正常。
- **60 天保活**：仓库连续 60 天无提交，定时任务会被自动停用。`cet4-daily.yml` 已内置每天一次「保活提交」，可避免。
- **产物备份**：每次生成的 Word 会作为 Artifact 保留 30 天，可在 Actions 运行页下载。
- **国内访问**：github.com 时通时断，但**任务执行不依赖你访问**，只在偶尔维护时需要打开网页。
- **安全**：所有密钥只存在 GitHub Secrets，代码中不硬编码；泄露时可随时重置。
