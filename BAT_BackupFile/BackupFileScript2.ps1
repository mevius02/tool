### ========================================
###  Backup設定エリア(ユーザー編集はここだけ)
### ========================================
# ●Backup対象(ファイル/フォルダ)パス(絶対パス)
$SourcePaths = @(
    "C:\Users\meviu\Desktop\TOOL\BAT_BackupFile\newFolder_test\新しいフォルダー1\FILE1.txt"
    , "C:\Users\meviu\Desktop\TOOL\BAT_BackupFile\newFolder_test\新しいフォルダー2\新しいフォルダー2-1"
)
# ●Backup先パス
$DestPath = "C:\Users\meviu\Desktop\TOOL\BAT_BackupFile\newFolder"
# ●タイムスタンプ付与設定
# (none=付与無し / pre=先頭に付与(YYYYMMDDHHMMSS_ファイル名) / sfx=末尾に付与(ファイル名_YYYYMMDDHHMMSS))
$AddTimestamp = "none"
# ●最新のみフラグ(Backup対象フォルダ内に複数存在する場合、名前降順の1番目のみBackup対象とする)
#  (例:Backup対象パス="C:\Users\Desktop\log"
#      配下一覧:C:\Users\Desktop\log\Log_20260101
#               C:\Users\Desktop\log\Log_20260201
#               C:\Users\Desktop\log\Log_20260301
#      Log_20260301のみ、Backup対象とする)
#  (true=最新のみコピー / false=すべてコピー)
$MostRecentOnlyFlg = $false
# === Backup対象の名称一致検索 ==========
#  (指定した(ファイル名/フォルダ名)が存在しない場合、と一致する場合、Backup対象とする)
#  (例:Backup対象パス="C:\Users\Desktop\log\java_log_202601010300.log"
#      前方一致="java_log_"
#      後方一致=".log"
#      該当した(ファイル/フォルダ)すべて、Backup対象とする)
#  (設定しない場合""空セット)
# ●前方一致
$MatchPrefix = ""
# ●後方一致
$MatchSuffix = ".log"
# ●部分一致
$MatchContains = ""
### ========================================

# ●Backup対象パス一覧
$BackupTgtPaths = @()
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

### === エラーチェック ==========
# Backup対象 指定チェック
if (-not $SourcePaths) {
    LogError -Msg "Backup対象が指定されていません"
    exit 1
}
# Backup対象 存在チェック
foreach ($path in $SourcePaths) {
    # 完全一致
    if (Test-Path -LiteralPath $path) {
        $BackupTgtPaths += $path
        continue
    }
    # 部分一致
    if ($MatchPrefix -ne "" -or $MatchSuffix -ne "" -or $MatchContains -ne "") {
        # 直前パス
        $BaseDir = Split-Path $path -Parent
        # ファイル/フォルダ名一覧 取得
        $items = GetChildItems -Path $BaseDir
        # 一致対象パス一覧
        $matchPaths = @()
        foreach ($item in $items) {
            if (IsMatchItemNm $item.Name) {
                $matchPaths += $item.FullName
            }
        }
        if (0 -lt $matchPaths.Count) {
            $BackupTgtPaths += $matchPaths
        } else {
            $ErrPaths += $path
        }
    } else {
        $ErrPaths += $path
    }
}
# 出力
if (0 -eq $BackupTgtPaths.Count -or 0 -lt $ErrPaths.Count) {
    LogError -Msg "Backup対象が存在しません"
    foreach ($errPath in $ErrPaths) {
        LogError -Msg $errPath
    }
    exit 1
}
# Backup先 指定チェック
if (-not $DestPath -or $DestPath.Trim() -eq "") {
    LogError -Msg "Backup先が指定されていません"
    exit 1
}
# Backup先 存在チェック
if (-not (Test-Path -LiteralPath $DestPath)) {
    LogError -Msg "Backup先が存在しません"
    LogError -Msg $DestPath
    exit 1
}

### === 最新のみ指定の場合 ==========
if ($MostRecentOnlyFlg) {
    # ファイル指定が含まれている場合
    foreach ($path in $BackupTgtPaths) {
        $item = Get-Item -LiteralPath $path
        if (-not $item.PSIsContainer) {
            LogError -Msg "最新のみフラグON の場合、ファイル指定はできません"
            LogError -Msg $path
            exit 1
        }
    }
    $tmpTgtPaths = $BackupTgtPaths
    $BackupTgtPaths = @()
    foreach ($path in $tmpTgtPaths) {
        # フォルダ配下の一覧取得
        $items = GetChildItems -Path $path
        # 一覧有りの場合
        if (0 -lt $items.Count) {
            # 名前降順でソートし、先頭1件だけ対象
            $mostRecentItem = $items | Sort-Object Name -Descending | Select-Object -First 1
            # BackupTgtPaths をその1件だけに置き換える
            $BackupTgtPaths += @($mostRecentItem.FullName)
        }
    }
}

### === Backup ==========
LogInfo -Msg ""
LogInfo -Msg "=== Backup ==========" 

# タイムスタンプ作成
$timestamp = (Get-Date).ToString("yyyyMMddHHmmss")

# コピー処理
foreach ($path in $BackupTgtPaths) {
    LogInfo -Msg $path
    $item = Get-Item -LiteralPath $path.TrimEnd('\') -ErrorAction Stop
    $backupName = Split-Path $path -Leaf
    # 先頭付与の場合
    if ($AddTimestamp -eq "pre") { $backupName = "${timestamp}_${backupName}" }
    # 末尾付与の場合
    if ($AddTimestamp -eq "sfx") {
        # ファイルの場合
        if ($item -is [System.IO.FileInfo]) {
            # ファイル名(拡張子無し)
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($backupName)
            # 拡張子
            $ext = [System.IO.Path]::GetExtension($backupName)
            $backupName = "${baseName}_${timestamp}$ext"
        # フォルダの場合
        } else {
            $backupName = "${backupName}_${timestamp}"
        }
    }
    # Backup先パス 作成
    $dest = Join-Path $DestPath $backupName
    # 配下一覧出力(フォルダの場合)
    if ($item.PSIsContainer) {
        $childItems = Get-ChildItem -LiteralPath $item.FullName -Recurse
        foreach ($child in $childItems) {
            LogInfo -Msg $child.FullName
        }
    }
    # 既に存在する場合の場合
    if (Test-Path -LiteralPath $dest) {
        LogWarn -Msg "既に存在するためスキップしました"
        LogWarn -Msg $dest
        continue
    }
    # コピー
    try {
        # フォルダの場合
        if ($item.PSIsContainer) {
            Copy-Item -LiteralPath $item.FullName -Destination $dest -Recurse -Force -ErrorAction Stop
        # ファイルの場合
        } else {
            Copy-Item -LiteralPath $item.FullName -Destination $dest -Force -ErrorAction Stop
        }
    } catch {
        LogError -Msg "コピー中にエラーが発生しました"
        LogError -Msg $path.FullName
        LogError -Msg $_.Exception.Message
        exit 1
    }
}
