@echo off
chcp 65001 >nul
cd /d %~dp0
echo Dang do ket noi CSDL MISA (chi doc, khong sua du lieu)...
python -m ketoan_ai.cli misa kham-pha
echo.
echo Gui file misa_kham_pha.txt trong thu muc nay cho nguoi cai dat.
pause
