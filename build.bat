@echo off
chcp 65001 >nul
echo ==========================================
echo    选择题答题软件 - 打包脚本
echo ==========================================
echo.

:: 检查Python环境
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python，请安装Python 3.8或更高版本
    pause
    exit /b 1
)

echo [1/5] 正在安装依赖...
pip install -r requirements.txt
if errorlevel 1 (
    echo [错误] 安装依赖失败
    pause
    exit /b 1
)
echo [完成] 依赖安装完成
echo.

echo [2/5] 正在清理旧构建...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "*.spec" del /q "*.spec"
echo [完成] 清理完成
echo.

echo [3/5] 正在打包应用程序...
pyinstaller --clean --onefile --windowed ^
    --name "选择题答题软件" ^
    --add-data "database.py;." ^
    --add-data "excel_handler.py;." ^
    --icon "NONE" ^
    quiz_app.py

if errorlevel 1 (
    echo [错误] 打包失败
    pause
    exit /b 1
)
echo [完成] 打包完成
echo.

echo [4/5] 正在复制必要文件...
if not exist "dist\saves" mkdir "dist\saves"
copy "sample_questions.xlsx" "dist\" >nul 2>&1
copy "README.md" "dist\" >nul 2>&1
echo [完成] 文件复制完成
echo.

echo [5/5] 正在创建快捷方式...
powershell -Command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('dist\选择题答题软件.lnk'); $Shortcut.TargetPath = '%CD%\dist\选择题答题软件.exe'; $Shortcut.WorkingDirectory = '%CD%\dist'; $Shortcut.Save()" >nul 2>&1
echo [完成] 快捷方式创建完成
echo.

echo ==========================================
echo    打包完成！
echo ==========================================
echo.
echo 输出文件位置：
echo   - 可执行文件：dist\选择题答题软件.exe
echo   - 快捷方式：dist\选择题答题软件.lnk
echo   - 示例文件：dist\sample_questions.xlsx
echo.
echo 首次运行前请确保：
echo   1. 安装目录有读写权限
echo   2. 安装目录下会自动创建saves文件夹存储存档
echo.
pause
