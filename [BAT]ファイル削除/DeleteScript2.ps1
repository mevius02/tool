### ========================================
###  削除設定エリア(ユーザー編集はここだけ)
### ========================================
# ●エラーチェック有無設定(削除対象が存在しない場合エラー終了とするか)
# (true=終了する / false=終了しない)
$ErrCheckFlg = $true
# ●ファイル削除
$DelFilePaths = @(
    "C:\Users\meviu\Desktop\TOOL\[BAT]ファイル削除\newFolder\新しいフォルダー1\FILE3.txt"
)
# ●フォルダ内だけ削除
$DelContentsPaths = @(
)
# ●フォルダごと削除
$DelFolderPaths = @(
)
# ●削除対象指定(フォルダ内だけ削除に適用される)
# (all=全て / file=ファイルだけ削除 / folder=フォルダだけ削除)
$DelType = "all"
# === 削除対象の名称一致検索(設定しない場合""空セット) ==========
# ●前方一致
$MatchPrefix = ""
# ●後方一致
$MatchSuffix = ""
# ●部分一致
$MatchContains = ""
### ========================================

# ●指定パス一覧
$AllPaths = @()
# ●存在パス一覧(ファイル)
$ExistsFilePaths = @()
# ●存在パス一覧(ファイル内)
$ExistsContentsPaths = @()
# ●存在パス一覧(フォルダごと)
$ExistsFolderPaths = @()
# ●未存在パス一覧(エラー出力用)
$ErrPaths = @()

### ========================================
###  関数
### ========================================
# ■■■ ログ出力(INFO) ■■■■■■■■■■
function LogInfo {
    param([string]$Msg)
    Write-Host "[" -NoNewline
    Write-Host "INFO" -ForegroundColor Blue -NoNewline
    Write-Host "] " $Msg
}
# ■■■ ログ出力(WARN) ■■■■■■■■■■
function LogWarn {
    param([string]$Msg)
    Write-Host "[" -NoNewline
    Write-Host "WARN" -ForegroundColor Yellow -NoNewline
    Write-Host "] " $Msg
}
# ■■■ ログ出力(ERROR) ■■■■■■■■■■
function LogError {
    param([string]$Msg)
    Write-Host "[" -NoNewline
    Write-Host "ERROR" -ForegroundColor Red -NoNewline
    Write-Host "]" $Msg
}

# ■■■ ゴミ箱削除 ■■■■■■■■■■
function RemoveItem {
    param([string]$Path, [switch]$Recurse)
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
        }
    }
}

# ■■■ 配下一覧 取得 ■■■■■■■■■■
function GetChildItems {
    param([string]$Path, [switch]$Recurse)
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

# ■■■ ファイル/フォルダ名 一致判定 ■■■■■■■■■■
function IsMatchItemNm {
    param($name)
    # === [前方一致/後方一致/部分一致] いずれか一致しない場合 ==========
    if ($MatchPrefix   -ne "" -and -not $name.StartsWith($MatchPrefix)) { return $false }
    if ($MatchSuffix   -ne "" -and -not $name.EndsWith($MatchSuffix))   { return $false }
    if ($MatchContains -ne "" -and -not $name.Contains($MatchContains)) { return $false }
    return $true
}

### ================================
###  メイン処理
### ================================
LogInfo -Msg ""
$fileNm = Split-Path $PSCommandPath -Leaf
LogInfo -Msg "■■■■■ 実行PS: $fileNm ■■■■■■■■■■"

if ($ErrCheckFlg)      { LogInfo -Msg "指定パス未存在時エラー終了: ON" }
if (-not $ErrCheckFlg) { LogWarn -Msg "指定パス未存在時エラー終了: OFF" }
if ($DelType -eq "all")    { LogInfo -Msg "削除対象: 全て" }
if ($DelType -eq "file")   { LogInfo -Msg "削除対象: ファイルのみ" }
if ($DelType -eq "folder") { LogInfo -Msg "削除対象: フォルダのみ" }

# 指定パスまとめ
$AllPaths += $DelFilePaths
$AllPaths += $DelContentsPaths
$AllPaths += $DelFolderPaths

### === エラーチェック ==========
# 存在パス/未存在パス まとめ
foreach ($path in $DelFilePaths) {
    if (-not (Test-Path -LiteralPath $path)) {
        $ErrPaths += $path
    } else {
        $ExistsFilePaths += $path
    }
}
foreach ($path in $DelContentsPaths) {
    if (-not (Test-Path -LiteralPath $path)) {
        $ErrPaths += $path
    } else {
        $ExistsContentsPaths += $path
    }
}
foreach ($path in $DelFolderPaths) {
    if (-not (Test-Path -LiteralPath $path)) {
        $ErrPaths += $path
    } else {
        $ExistsFolderPaths += $path
    }
}
# 出力
if ($ErrPaths.Count -gt 0) {
    LogInfo -Msg ""
    LogError -Msg "指定パスが存在しません"
    foreach ($errPath in $ErrPaths) {
        LogError -Msg $errPath
    }
    # エラー終了設定の場合
    if ($ErrCheckFlg) {
        exit 1
    }
}

### === 削除 ==========
# --- ファイル ----------
if ($ExistsFilePaths) {
    LogInfo -Msg ""
    LogInfo -Msg "=== ファイル削除 ==========" 
}
foreach ($path in $ExistsFilePaths) {
    LogInfo -Msg $path
    # 削除
    RemoveItem -Path $path
}

# --- フォルダ内 ----------
if ($ExistsContentsPaths) {
    LogInfo -Msg ""
    LogInfo -Msg "=== フォルダ内削除 =========="
}
foreach ($path in $ExistsContentsPaths) {
    # ●削除成功フラグ
    $deletedFlg = $false
    LogInfo -Msg $path
    # 配下の削除一覧 取得
    $items = GetChildItems -Path $path
    foreach ($item in $items) {
        # 名称一致条件に一致しない場合
        if (-not (IsMatchItemNm $item.Name)) { continue }
        # ファイルだけ削除 & 対象がフォルダの場合
        if ($DelType -eq "file" -and $item.PSIsContainer) { continue }
        # フォルダだけ削除 & 対象がファイルの場合
        if ($DelType -eq "folder" -and -not $item.PSIsContainer) { continue }
        $deletedFlg = $true
        if ($item.PSIsContainer) {
            $subItems = GetChildItems -Path $item.FullName -Recurse
            foreach ($subItem in $subItems) {
                LogInfo -Msg $($subItem.FullName)
            }
            RemoveItem -Path $item.FullName -Recurse
        } else {
            RemoveItem -Path $item.FullName
        }
    }
    # 削除対象無しの場合
    if (-not $deletedFlg) {
        LogWarn -Msg "フォルダ内に削除対象がありません"
    }
}

# --- フォルダごと ----------
if ($ExistsFolderPaths) {
    LogInfo -Msg ""
    LogInfo -Msg "=== フォルダごと削除 =========="
}
foreach ($path in $ExistsFolderPaths) {
    LogInfo -Msg $path
    $items = GetChildItems -Path $path -Recurse
    foreach ($item in $items) {
        LogInfo -Msg $($item.FullName)
    }
    RemoveItem -Path $path -Recurse
}
