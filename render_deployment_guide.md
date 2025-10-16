# Render 部署详细指南

本指南提供了在 Render 平台上部署压实度文档处理工具的详细步骤。

## 前提条件

- GitHub 账号，且项目已上传到 GitHub
- Render 账号（可免费注册）
- 基本的 Git 操作知识

## 步骤 1：准备项目文件

确保您的项目包含以下关键文件：

1. `web_app.py` - Flask 应用主文件
2. `requirements.txt` - Python 依赖列表
3. `excel_to_word.py` - 核心功能模块
4. `Procfile` - 进程配置文件

## 步骤 2：注册/登录 Render

1. 访问 [Render 官网](https://render.com/)
2. 点击 "Sign Up" 创建账号或 "Log In" 登录现有账号

## 步骤 3：连接 GitHub 仓库

1. 登录后，点击顶部导航栏中的 "New +" 按钮
2. 选择 "Web Service" 选项
3. 在 "Connect a Repository" 部分，选择您的 GitHub 账号
4. 在仓库列表中找到您的项目仓库（如 `yashidu`）
5. 点击 "Connect" 按钮

## 步骤 4：配置部署设置

在配置页面填写以下信息：

### 基本设置

- **Name**: 输入应用名称（如 `yashidu-yashidu`）
- **Region**: 选择最靠近您用户的区域（推荐选择 `Oregon (US West)`）
- **Branch**: 选择 `master`（或您的主分支名称）
- **Runtime**: 选择 `Python`

### 构建和启动命令

- **Build Command**: `pip install -r requirements.txt`
- **Start Command**: `gunicorn web_app:application`

### 环境变量

点击 "Add Environment Variable" 添加以下变量：

| 变量名 | 值 | 说明 |
|--------|-----|------|
| `FLASK_ENV` | `production` | 生产环境模式 |
| `FLASK_SECRET_KEY` | 任意随机字符串 | 用于加密会话数据 |
| `PYTHONUNBUFFERED` | `1` | 确保日志实时输出 |

## 步骤 5：配置资源

- **Plan**: 对于简单测试，可以选择 "Free"
- **Instance Type**: 选择适合的实例类型
- 对于生产环境，建议选择至少 2GB 内存的付费计划

## 步骤 6：启动部署

1. 确认所有设置无误后，点击 "Create Web Service" 按钮
2. Render 将开始构建和部署您的应用
3. 部署过程中可以查看实时日志

## 步骤 7：验证部署

1. 部署完成后，Render 会显示绿色的 "Live" 状态
2. 点击提供的 URL（如 `https://your-app-name.onrender.com`）访问应用
3. 测试文件上传和文档转换功能

## 自动部署设置

Render 默认已配置自动部署，每次推送到主分支时都会触发重新部署。

## 常见问题及解决方案

### 1. 构建失败

- 检查构建日志中的错误信息
- 确保 `requirements.txt` 中的依赖版本正确且兼容
- 尝试在本地运行构建命令以复现问题

### 2. 应用启动失败

- 检查启动日志
- 确保 `web_app.py` 中正确定义了 `application` 函数
- 验证 `gunicorn` 已包含在依赖中

### 3. 文件上传失败

- 检查应用是否有写入临时目录的权限
- 验证 Flask 的文件大小限制设置
- 查看应用日志中的错误信息

### 4. 内存不足错误

- 如果处理大文件时遇到内存错误，升级到内存更大的实例类型
- 优化代码以减少内存使用

## 资源监控

- 在 Render 控制台可以查看应用的 CPU 和内存使用情况
- 设置告警以监控应用状态

## 扩展指南

### 配置自定义域名

1. 在 Render 控制台中选择您的应用
2. 点击 "Settings" > "Custom Domains"
3. 按照提示添加您的自定义域名

### 配置持久存储

对于需要保存上传文件的场景，可以：

1. 使用 Render 的 "Disks" 功能（需要付费计划）
2. 配置云存储服务（如 AWS S3、Google Cloud Storage）

### 配置环境分离

可以创建多个环境（开发、测试、生产）：

1. 创建不同的分支对应不同环境
2. 在 Render 中为每个分支创建独立的服务
3. 配置相应的环境变量

## 联系支持

如有部署问题，可以：

1. 查看 [Render 官方文档](https://render.com/docs)
2. 在 [Render 社区论坛](https://community.render.com/) 寻求帮助
3. 联系 Render 客户支持

---

祝您部署顺利！如有任何问题，请参考 Render 官方文档或联系技术支持。