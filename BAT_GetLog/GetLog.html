@echo on
setlocal enabledelayedexpansion

rem デスクトップパス
set "desktop=%USERPROFILE%\Desktop"

rem 日時取得
for /f "tokens=1-3 delims=/ " %%a in ("%date%") do (
    set "YYYY=%%a"
    set "MM=%%b"
    set "DD=%%c"
)
for /f "tokens=1-3 delims=:." %%a in ("%time%") do (
    set "HH=%%a"
    set "Min=%%b"
    set "Sec=%%c"
)

rem 0埋め
if 1%!HH!  LSS 110 set "HH=0!HH!"
if 1%!Min! LSS 110 set "Min=0!Min!"
if 1%!Sec! LSS 110 set "Sec=0!Sec!"

rem タイムスタンプ
set "timestamp=%YYYY%%MM%%DD%%HH%%Min%%Sec%"

rem コピー元 / コピー先
set "source=%desktop%\SQL.sql"
set "destination=%desktop%\SQL_bak\SQL_%timestamp%.sql"

rem コピー
copy "%source%" "%destination%" >nul

pause
exit
