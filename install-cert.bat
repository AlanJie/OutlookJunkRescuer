@echo off
chcp 65001 >nul
echo 正在将 ClickOnce 证书导入受信任的证书存储区...

cd /d "%~dp0"
if not exist "OutlookJunkRescuer.cer" (
    echo [错误] 未找到 OutlookJunkRescuer.cer 证书文件！
    pause
    exit /b 1
)

certutil -addstore -user Root OutlookJunkRescuer.cer >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 导入“受信任的根证书颁发机构”失败，错误代码: %errorlevel%
    echo 请确认是否已使用管理员权限运行此脚本。
    pause
    exit /b %errorlevel%
)

certutil -addstore -user TrustedPublisher OutlookJunkRescuer.cer >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 导入“受信任的发布者”失败，错误代码: %errorlevel%
    echo 请确认是否已使用管理员权限运行此脚本。
    pause
    exit /b %errorlevel%
)

echo ========================================================
echo  [成功] 证书已成功导入“受信任的根证书颁发机构”与“受信任的发布者”！
echo  现在你可以直接双击 OutlookJunkRescuer.vsto 执行标准 ClickOnce 安装。
echo ========================================================

pause
