# Excel到Word自动化工具

这是一个用于将Excel数据自动转换为格式化Word文档的工具。该项目可以作为桌面应用程序运行，也可以部署为Web服务，通过网页界面进行操作。

## 功能特点

- 从Excel文件读取数据并填充到Word文档模板
- 支持多工作表处理
- 自动生成格式化的表格
- 支持批量处理多个数据组
- 提供桌面GUI和Web界面两种使用方式

## 部署为Web应用

### 方法1：使用GitHub Pages + 后端服务

1. 将项目推送到GitHub仓库
2. 使用Heroku、Render或其他平台部署后端服务
3. 配置CI/CD工作流自动部署（项目已包含基础配置）

### 方法2：本地开发运行

1. 确保已安装Python 3.7或更高版本
2. 安装依赖：
   ```
   pip install -r requirements.txt
   ```
3. 运行Web应用：
   ```
   python web_app.py
   ```
4. 在浏览器中访问 http://localhost:5000

## 部署到云平台

### Heroku部署

项目已包含基础的Heroku部署配置（Procfile和ci-cd.yml）。需要在GitHub Secrets中设置：
- HEROKU_API_KEY：您的Heroku API密钥
- HEROKU_APP_NAME：您的Heroku应用名称

### Render部署

#### 部署步骤

1. **创建Render账号**
   - 访问 [Render官网](https://render.com/) 注册账号

2. **创建新的Web Service**
   - 登录后点击 "New +" 按钮
   - 选择 "Web Service"
   - 连接你的GitHub仓库
   - 选择你要部署的仓库

3. **配置部署设置**
   - **Name**: 输入应用名称
   - **Region**: 选择最近的区域（如 "Oregon (US West)"）
   - **Branch**: 选择主分支（如 "master" 或 "main"）
   - **Runtime**: 选择 "Python"
   - **Build Command**: 输入 `pip install -r requirements.txt`
   - **Start Command**: 输入 `gunicorn web_app:application`
   - **Environment Variables**:
     - 点击 "Add Environment Variable"
     - 添加 `FLASK_ENV` = `production`
     - 添加 `FLASK_SECRET_KEY` = 任意安全的随机字符串

4. **配置资源**
   - 选择 "Free" 或适合的付费计划
   - 确认配置后点击 "Create Web Service"

5. **等待部署完成**
   - Render会自动构建和部署应用
   - 部署完成后，可以通过提供的URL访问应用

#### 注意事项

- Render的免费计划会在一段时间不活跃后自动休眠
- 对于文件处理任务，建议使用至少2GB内存的付费计划
- 应用会自动连接GitHub，每次推送到master/main分支会触发重新部署

#### 故障排除

如果部署失败，请检查：

1. 构建日志中的错误信息
2. 确保requirements.txt中的依赖版本兼容
3. 检查Flask应用是否正确配置了WSGI入口点（application函数）

#### 所需环境变量

在GitHub Secrets中设置：
- RENDER_SERVICE_ID：您的Render服务ID
- RENDER_API_KEY：您的Render API密钥

## 项目结构

- `excel_to_word.py`：核心转换功能
- `web_app.py`：Web应用接口
- `main_app.py`：桌面GUI应用
- `templates/index.html`：Web界面模板
- `.github/workflows/ci-cd.yml`：CI/CD配置

## 依赖项

- pandas 1.3.5
- python-docx 0.8.11
- pyinstaller 5.13.2
- lxml 4.6.3
- pywin32 301
- Flask 2.0.1
- Werkzeug 2.0.1

## 使用说明

### Web界面使用

1. 访问Web应用URL
2. 上传Excel文件（.xlsx格式）
3. 上传Word模板文件（.docx格式）
4. 设置复制次数
5. 点击"处理文件"按钮
6. 等待处理完成后，下载生成的Word文件

## 注意事项

- 确保Excel文件格式符合预期（特别是数据位置和格式）
- Web应用处理大文件可能需要较长时间
- 某些高级功能可能需要Microsoft Excel安装（通过pywin32）

## License

[MIT](https://opensource.org/licenses/MIT)