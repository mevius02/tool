@echo off
echo :::::::::: 削除バッチ 開始 ::::::::::::::::::::::::::::::::::::::::
powershell -ExecutionPolicy Bypass -File ".\DeleteScript1.ps1"
powershell -ExecutionPolicy Bypass -File ".\DeleteScript2.ps1"
echo :::::::::: 削除バッチ 終了 ::::::::::::::::::::::::::::::::::::::::
pause
