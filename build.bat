@echo off
chcp 65001 >nul
echo ========================================
echo 邮件发送工具 - 打包脚本 (增强版)
echo ========================================
echo.

echo [1/4] 检查Python环境...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到Python，请先安装Python 3.7或更高版本
    pause
    exit /b 1
)
echo Python环境检查通过
echo.

echo [2/4] 安装依赖...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo 错误: 依赖安装失败
    pause
    exit /b 1
)
echo 依赖安装完成
echo.

echo [3/4] 清理旧的打包文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo 清理完成
echo.

echo [4/4] 开始打包...
pyinstaller --onefile --windowed --name="邮件发送工具" --clean ^
    --icon=icon.png ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.ttk ^
    --hidden-import=tkinter.scrolledtext ^
    --hidden-import=tkinter.filedialog ^
    --hidden-import=tkinter.messagebox ^
    --hidden-import=openpyxl ^
    main.py
if %errorlevel% neq 0 (
    echo 错误: 打包失败
    pause
    exit /b 1
)
echo 打包完成
echo.

echo ========================================
echo 打包成功！
echo ========================================
echo.
echo 可执行文件位置: dist\邮件发送工具.exe
echo.
echo 使用说明:
echo 1. 将 dist\邮件发送工具.exe 复制到任意位置
echo 2. 双击运行即可使用
echo 3. 数据库文件 email_data.db 会自动生成
echo.
echo 注意事项:
echo - 请确保网络连接正常
echo - 首次使用需要配置邮箱信息
echo - 建议定期备份数据库文件
echo - Excel导入功能需要openpyxl库支持
echo ========================================
echo.

pause
