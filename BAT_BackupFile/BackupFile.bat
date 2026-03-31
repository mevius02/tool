@echo off
echo :::::::::: Backupバッチ 開始 ::::::::::::::::::::::::::::::::::::::::
powershell -ExecutionPolicy Bypass -File ".\BackupFileScript1.ps1"
powershell -ExecutionPolicy Bypass -File ".\BackupFileScript2.ps1"
echo :::::::::: Backupバッチ 終了 ::::::::::::::::::::::::::::::::::::::::
pause
