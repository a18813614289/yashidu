@echo off
REM 运行压实度文档处理工具
REM 创建于 %date% %time%

REM 设置中文显示
chcp 65001 >nul

REM 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到Python安装。请先安装Python。
    pause
    exit /b 1
)

REM 检查是否存在必要的Python库
python -c "import openpyxl, pandas, docx, win32com.client, lxml" >nul 2>&1
if %errorlevel% neq 0 (
    echo 检测到缺少必要的Python库，正在尝试安装...
    pip install openpyxl pandas python-docx pywin32 lxml
    if %errorlevel% neq 0 (
        echo 错误: 安装必要的Python库失败。请手动安装以下库：openpyxl, pandas, python-docx, pywin32, lxml
        pause
        exit /b 1
    )
)

REM 运行主程序
echo 正在启动压实度文档处理工具...
python "%~dp0压实度文档处理工具.py"

REM 检查程序是否正常退出
if %errorlevel% neq 0 (
    echo 程序异常退出，错误代码: %errorlevel%
    pause
    exit /b %errorlevel%
)

echo 程序已正常退出
REM pause >nul  # 取消注释此行以在程序退出后保持窗口打开
exit /b 0