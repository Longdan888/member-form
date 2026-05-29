# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 目录概述

本目录主要包含业务数据文件和在线信息收集工具。

### 业务数据文件

| 文件 | 说明 |
|------|------|
| `88vip会员名单.xlsx` | 淘宝 88VIP 会员名单 |
| `京东plus会员名单.xlsx` | 京东 Plus 会员名单 |
| `舒舒克京东plus会员名单.xlsx` | 舒舒克渠道京东 Plus 会员 |
| `i茅台(1).xlsx` | i茅台相关名单 |
| `龙蛋业务人头名单.xlsx` | 龙蛋业务人员名单 |
| `龙蛋-吴浩民业务人头名单.xlsx` | 吴浩民负责的龙蛋业务 |
| `AD钙名单-人员统计.xlsx` | AD钙业务人员统计 |
| `群发文字稿.docx` | 群发消息话术模板 |

---

## 在线信息收集工具 (app.py)

Web 表单应用，会员在线填写信息，数据存入 SQLite 数据库，支持导出 Excel。

### 本地运行

```bash
pip install -r requirements.txt
python app.py                    # 开发模式 (Flask debug, 端口 5000)

# 临时公网分享（另开终端）
cloudflared tunnel --url http://localhost:5000
```

### 页面路由

| 路由 | 用途 | 认证 |
|------|------|------|
| `/form` | 会员登记表单（手机端适配） | 无 |
| `/admin` | 管理后台：查看数据、统计、导出 | 网页密码登录 |
| `/admin/download` | 下载全部数据为 Excel | 需登录 |
| `/admin/export/<类型>` | 按会员类型导出 Excel | 需登录 |

### 环境变量

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `SECRET_KEY` | `change-me-in-production-2026` | Flask session 密钥 |
| `ADMIN_PASSWORD` | `admin123` | 管理后台登录密码 |
| `PORT` | `5000` | 服务端口（云端自动注入） |

### 技术架构

- **Web 框架**：Flask
- **数据库**：SQLite (`data/submissions.db`)，WAL 模式，支持并发读写
- **生产服务器**：Waitress（`python app.py --prod`）
- **Excel 导出**：openpyxl（按需生成 .xlsx，不存储中间文件）
- **认证**：Session + Cookie 会话管理
- **去重**：手机号 UNIQUE 约束，重复提交自动拒绝

### 数据库表结构

```sql
submissions (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at      TEXT    NOT NULL,
    name            TEXT    NOT NULL,
    phone           TEXT    NOT NULL UNIQUE,
    membership_type TEXT    NOT NULL,
    account_id      TEXT    DEFAULT '',
    remarks         TEXT    DEFAULT ''
)
```

---

## 云端部署指南

推荐使用 Render 免费方案（或 Railway / Fly.io）。

### 步骤 1：推送到 GitHub

```bash
git init
git add app.py requirements.txt Procfile runtime.txt templates/ data/
echo "data/*.db" >> .gitignore       # 数据库不提交
git commit -m "v2.0 SQLite + waitress"
git remote add origin <你的仓库地址>
git push -u origin main
```

### 步骤 2：在 Render 创建 Web Service

1. 登录 [render.com](https://render.com)，New → Web Service
2. 连接 GitHub 仓库
3. 配置：
   - **Build Command**：`pip install -r requirements.txt`
   - **Start Command**：`python app.py --prod`
   - **Runtime**：Python 3
4. 添加环境变量：`SECRET_KEY`、`ADMIN_PASSWORD`
5. 点击 Deploy

### 步骤 3：绑定域名（可选）

Render 会分配 `xxx.onrender.com` 免费域名，也可以在 Settings 中绑定自定义域名。

### 注意事项

- 免费 Render 实例 15 分钟无访问会自动休眠，唤醒需 30-60 秒
- SQLite 数据库在实例休眠/重启时会保留（持久化磁盘）
- 建议定期从管理后台下载 Excel 备份数据
- 生产环境务必修改默认 `SECRET_KEY` 和 `ADMIN_PASSWORD`
