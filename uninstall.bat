@echo off
chcp 65001 >nul
echo 正在卸载 OutlookJunkRescuer 插件...

reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookJunkRescuer" /f >nul 2>&1

echo ========================================================
echo  [成功] 插件已从当前用户的 Outlook 加载项中注销。
echo ========================================================

pause
