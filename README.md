# 📄 Word to Markdown Converter Online

> 一款完全免费、开源的在线工具，轻松将 Word 文档转换为 Markdown 格式，图片自动上传至图床并获取永久外链。

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.11+-green.svg)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.115+-teal.svg)](https://fastapi.tiangolo.com/)
[![Render](https://img.shields.io/badge/Deploy-Render-purple.svg)](https://render.com)

---

## ✨ 特性

| 特性 | 说明 |
|------|------|
| 🚀 **一键转换** | 上传 `.docx` 文件，自动转换为 Markdown 格式 |
| 🖼️ **图片自动上传** | 文档中的图片自动上传至 GitHub 图床 |
| 🌐 **CDN 加速** | 使用 jsDelivr 全球 CDN 加速图片访问 |
| 📊 **表格支持** | 完整保留 Word 表格结构 |
| 🎨 **格式保留** | 支持标题、粗体、斜体等文本格式 |
| 🔒 **隐私保护** | 文件即时转换，不在服务器永久保存 |
| 💰 **完全免费** | 无需注册、无需付费、无需安装任何软件 |
| 🔓 **开源透明** | 代码完全开源，可自行部署 |

---

## 🎯 快速开始

### 在线使用（推荐）

直接访问：**[https://word2md-online.onrender.com](https://word2md-online.onrender.com)**

1. 点击或拖拽上传 `.docx` 文件
2. 点击「转换为 Markdown」
3. 自动下载转换后的 `.md` 文件

### 本地部署

```bash
# 克隆仓库
git clone https://github.com/Herbariaa/word2md-online.git
cd word2md-online

# 安装依赖
pip install -r requirements.txt

# 配置环境变量
export GITHUB_TOKEN="your_github_token"
export GITHUB_REPO="your_username/word2md-images"

# 启动服务
uvicorn app:app --host 0.0.0.0 --port 8000
```

---

## 🏗️ 技术架构

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│   Frontend  │────▶│   FastAPI   │────▶│  python-docx│
│   (HTML/JS) │     │   Backend   │     │   Parser    │
└─────────────┘     └─────────────┘     └─────────────┘
                           │                    │
                           ▼                    ▼
                    ┌─────────────┐     ┌─────────────┐
                    │   GitHub    │     │   jsDelivr  │
                    │   Image CDN │     │     CDN     │
                    └─────────────┘     └─────────────┘
```

| 组件 | 技术 | 说明 |
|------|------|------|
| 后端框架 | FastAPI | 高性能异步 Web 框架 |
| 文档解析 | python-docx | Word 文档解析库 |
| 图片存储 | GitHub API | 自动上传至指定仓库 |
| CDN 加速 | jsDelivr | 全球加速图片访问 |
| 部署平台 | Render | 免费容器托管服务 |

---

## 📁 项目结构

```
word2md-online/
├── app.py                 # FastAPI 后端主程序
├── requirements.txt       # Python 依赖
├── Dockerfile            # Docker 镜像配置
├── static/
│   └── index.html        # 前端页面
└── README.md             # 项目文档
```

---

## 🔧 配置说明

### 环境变量

| 变量 | 必填 | 说明 |
|------|------|------|
| `GITHUB_TOKEN` | ✅ | GitHub Personal Access Token |
| `GITHUB_REPO` | ✅ | 图片存储仓库，格式：`用户名/仓库名` |

### GitHub Token 权限

创建 Token 时需要勾选：
- ✅ `repo` (Full control of private repositories)

---

## 🎨 支持的格式

| 类型 | 支持情况 |
|------|----------|
| 标题 (H1-H5) | ✅ 完全支持 |
| 粗体 / 斜体 | ✅ 完全支持 |
| 列表 | ✅ 完全支持 |
| 表格 | ✅ 完全支持 |
| 图片 | ✅ 自动上传至图床 |
| 段落 | ✅ 完全支持 |

---

## 🚀 在线演示

访问 **[https://word2md-online.onrender.com](https://word2md-online.onrender.com)** 立即体验！

---

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

---

## 📄 开源协议

本项目采用 MIT 协议开源，详见 [LICENSE](LICENSE) 文件。

---

## 📧 联系方式

- GitHub Issues: [https://github.com/Herbariaa/word2md-online/issues](https://github.com/Herbariaa/word2md-online/issues)
- 在线工具: [https://word2md-online.onrender.com](https://word2md-online.onrender.com)

---

## 🌟 支持项目

如果这个工具对你有帮助，欢迎给项目点个 Star ⭐️！
