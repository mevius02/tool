@echo off

rem ========================================
rem  バッチと同じ場所に[befre.txt]を作成し、変換したい文字列を書いて、バッチ実行
rem ========================================

cd /d "%~dp0"
set "befFile=%~dp0before.txt"
set "aftFile=%~dp0after.txt"

powershell -Command ^
  "$lines = Get-Content -LiteralPath '%befFile%';" ^
  "$converted = foreach ($line in $lines) {" ^
  "    $line = $line.Trim();" ^
  "    if ($line -cmatch '^[A-Z0-9_]+$') {" ^
  "        $parts = $line -split '_';" ^
  "        $camel = ($parts | ForEach-Object { $_.Substring(0,1).ToUpper() + $_.Substring(1).ToLower() }) -join '';" ^
  "        $camel.Substring(0,1).ToLower() + $camel.Substring(1);" ^
  "    } elseif ($line -cmatch '^[a-z0-9]+([A-Z][a-z0-9]*)*$') {" ^
  "        ($line -creplace '([a-z0-9])([A-Z])', '$1_$2' -creplace '([A-Za-z])([0-9])', '$1_$2').ToUpper();" ^
  "    } else {" ^
  "        $line;" ^
  "    }" ^
  "};" ^
  "$converted | Set-Content -LiteralPath '%aftFile%'"

exit