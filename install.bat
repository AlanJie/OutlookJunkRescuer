@echo off
chcp 65001 >nul
echo 正在安装 OutlookJunkRescuer 插件...

cd /d "%~dp0"
set ADDIN_PATH=%~dp0OutlookJunkRescuer.vsto

reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookJunkRescuer" /v "FriendlyName" /d "OutlookJunkRescuer" /f >nul
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookJunkRescuer" /v "Description" /d "OutlookJunkRescuer" /f >nul
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookJunkRescuer" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookJunkRescuer" /v "Manifest" /d "file:///%ADDIN_PATH%|vstolocal" /f >nul

if %ERRORLEVEL% equ 0 (
    echo ========================================================
    echo  [成功] 插件已成功注册到当前用户的 Outlook 加载项！
    echo  请完全重启经典版 Outlook 生效。
    echo ========================================================
) else (
    echo [错误] 注册表写入失败，请以普通用户或管理员权限重新运行。
)

pause
