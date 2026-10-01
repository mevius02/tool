@echo off
cd /d "%~dp0"
if exist merge.sql del merge.sql
set COUNT=0
for %%f in (*.sql) do if /i not "%%~nxf"=="merge.sql" call :append "%%f"
if %COUNT%==0 (
    echo No .sql files found in: %~dp0
) else (
    echo Done: merged %COUNT% files into merge.sql
)
exit /b

:append
type %1 >> merge.sql
echo.>> merge.sql
echo.>> merge.sql
echo.>> merge.sql
set /a COUNT+=1
exit /b