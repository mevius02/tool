<#
.SYNOPSIS
  ログレベルEnum
#>
enum LogLvl {
    INFO
    WARN
    ERROR
}

$LogColors = @{
    [LogLvl]::INFO  = "Blue"
    [LogLvl]::WARN  = "Yellow"
    [LogLvl]::ERROR = "Red"
}

$LogPadding = @{
    ([LogLvl]::INFO)  = "  "
    ([LogLvl]::WARN)  = "  "
    ([LogLvl]::ERROR) = " "
}

<#
.SYNOPSIS
  ログ出力(WARN用)
.DESCRIPTION
  出力例
  [WARN]  メッセージ内容
  (WARN文字だけ黄色)
.PARAMETER Msg
  出力したいメッセージ内容
.EXAMPLE
  LogInfo -Msg "処理を開始します"
#>
function WriteLog {
    param(
        [LogLvl]$Lvl,
        [string]$Msg,
        [switch]$OutputLogFile
    )
    # タイムスタンプ生成
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff") + " "
    # コンソール出力
    $color = $LogColors[$Lvl]
    $pad = $LogPadding[$Lvl]
    Write-Host "[" -NoNewline
    Write-Host $Lvl -ForegroundColor $color -NoNewline
    Write-Host "]$pad$Msg"
    # ファイル出力(ファイル存在しない場合、作成。ファイル存在する場合、追記)
    if ($OutputLogFile) {
        Add-Content -Path "script.log" -Value "[$Lvl]$pad$timestamp$Msg"
    }
}
<#
.SYNOPSIS
  ログ出力(INFO用)
.DESCRIPTION
  出力例
  [INFO]  メッセージ内容
  (INFO文字だけ青色)
.PARAMETER Msg
  出力したいメッセージ内容
.RETURNS
  -
.EXAMPLE
  LogInfo -Msg "処理を開始します"
#>
function LogInfo {
    param(
        [string]$Msg
    )
    Write-Host "[" -NoNewline
    Write-Host "INFO" -ForegroundColor Blue -NoNewline
    Write-Host "] " $Msg
}
<#
.SYNOPSIS
  ログ出力(WARN用)
.DESCRIPTION
  出力例
  [WARN]  メッセージ内容
  (WARN文字だけ黄色)
.PARAMETER Msg
  出力したいメッセージ内容
.RETURNS
  -
.EXAMPLE
  LogInfo -Msg "処理を開始します"
#>
function LogWarn {
    param(
        [string]$Msg
    )
    Write-Host "[" -NoNewline
    Write-Host "WARN" -ForegroundColor Yellow -NoNewline
    Write-Host "] " $Msg
}
<#
.SYNOPSIS
  ログ出力(ERROR用)
.DESCRIPTION
  出力例
  [ERROR] メッセージ内容
  (ERROR文字だけ赤色)
.PARAMETER Msg
  出力したいメッセージ内容
.RETURNS
  -
.EXAMPLE
  LogInfo -Msg "処理を開始します"
#>
function LogError {
    param(
        [string]$Msg
    )
    Write-Host "[" -NoNewline
    Write-Host "ERROR" -ForegroundColor Red -NoNewline
    Write-Host "]" $Msg
}
<#
.SYNOPSIS
  パス存在チェック
.DESCRIPTION
  引数パスが存在するか判定、結果を返す
.PARAMETER Path
  出力したいメッセージ内容
.RETURNS
  TRUE:存在する / FALSE:存在しない
.EXAMPLE
  Exists -Path "C:\Users\Desctop"
.EXAMPLE
  Exists -Path "C:\Users\Desctop\MEMO.txt"
#>
function Exists {
    param(
        [string]$Path
    )
    return Test-Path -LiteralPath $Path
}
<#
.SYNOPSIS
  指定パスのファイル、またはフォルダを削除する
.DESCRIPTION
  Remove-Item をラップし、例外処理・ログ出力を統一する
  Recurse 指定時はフォルダ配下を再帰的に削除する
.PARAMETER Path
  削除対象ファイル、またはフォルダの絶対パス
.PARAMETER Recurse
  配下も削除する場合、記載する
.RETURNS
  -
.EXAMPLE
  RemoveItem -Path "C:\Temp\新しいフォルダー\MEMO.txt"
.EXAMPLE
  RemoveItem -Path "C:\Temp\新しいフォルダー" -Recurse
#>
function RemoveItem {
    param(
        [string]$Path,
        [switch]$Recurse
    )
    ### === 削除 ==========
    try {
        # 最下層まで削除する場合
        if ($Recurse) {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        } else {
            Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
        }
    } catch {
        if ($ErrCheckFlg) {
            LogError -Msg "削除に失敗しました"
            throw
        } else {
            LogWarn -Msg "削除に失敗しましたが、続行します"
        }
    }
}
<#
.SYNOPSIS
  配下一覧 取得
.DESCRIPTION
  指定のフォルダ配下一覧を取得、返す
.PARAMETER Path
  取得対象フォルダの絶対パス
.PARAMETER Recurse
  配下も削除する場合、記載する
.RETURNS
  ファイル&フォルダ一覧
.EXAMPLE
  RemoveItem -Path "C:\Temp\新しいフォルダー"
.EXAMPLE
  RemoveItem -Path "C:\Temp\新しいフォルダー" -Recurse
#>
function GetChildItems {
    param(
        [string]$Path,
        [switch]$Recurse
    )
    try {
        if ($Recurse) {
            return Get-ChildItem -LiteralPath $Path -Recurse -ErrorAction Stop
        } else {
            return Get-ChildItem -LiteralPath $Path -ErrorAction Stop
        }
    } catch {
        if ($ErrCheckFlg) {
            throw
        } else {
            # 空配列返却
            return @()
        }
    }
}
<#
.SYNOPSIS
  ファイル/フォルダ名 一致判定
.DESCRIPTION
  指定のフォルダ配下一覧を取得、返す
.PARAMETER Path
  取得対象フォルダの絶対パス
.PARAMETER Recurse
  配下も削除する場合、記載する
.RETURNS
  ファイル&フォルダ一覧
.EXAMPLE
  RemoveItem -Path "C:\Temp\新しいフォルダー"
.EXAMPLE
  RemoveItem -Path "C:\Temp\新しいフォルダー" -Recurse
#>
function IsMatchItemNm {
    param($name)
    # === [前方一致/後方一致/部分一致] いずれか一致しない場合 ==========
    if ($MatchPrefix   -ne "" -and -not $name.StartsWith($MatchPrefix)) { return $false }
    if ($MatchSuffix   -ne "" -and -not $name.EndsWith($MatchSuffix))   { return $false }
    if ($MatchContains -ne "" -and -not $name.Contains($MatchContains)) { return $false }
    return $true
}
