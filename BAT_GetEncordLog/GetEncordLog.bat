@echo off
setlocal

REM === コピー元ファイルパス ==========
set SRC_PATH=C:\abc\abc.log

REM === ファイル名 切り抜き ==========
set SRC_FILE_NM=%~nx1
for %%A in ("%SRC_PATH%") do set SRC_NAME=%%~nxA

REM === コピー先パス ==========
REM このバッチと同じ場所
set DST_PATH=%~dp0

"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" ^
  -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ts = Get-Date -Format 'yyyyMMddHHmmss'; " ^
  "Get-Content '%SRC_PATH%' -Encoding UTF8 | " ^
  "ForEach-Object { $_ -replace [char]0,'' } | " ^
  "Out-File -Encoding UTF8 ('%DST_PATH%' + $ts + '_%SRC_FILE_NM%')"

endlocal
