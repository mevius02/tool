# ============================================================
# LogInfo / LogWarn / LogError（実行確認）
# ============================================================
Describe "Write-Log (比較表示)" {

    BeforeAll {
        Import-Module (Join-Path $PSScriptRoot "..\DeleteFileModule1.psm1") -Force
    }

    InModuleScope DeleteFileModule1 {

        Mock Write-Host {
            param($Object, $ForegroundColor, $NoNewline)
            $script:Captured += $Object
        }

        It "INFO の出力文字列を比較して表示する" {
            $script:Captured = @()

            Write-Log INFO "msg"

            $expected = "[INFO]  msg"
            $actual   = ($script:Captured -join "")

            Write-Verbose "期待値: $expected"
            Write-Verbose "実際値: $actual"

            if ($expected -eq $actual) {
                Write-Verbose "→ 一致: OK"
            } else {
                Write-Verbose "→ 不一致: NG"
            }

            $actual | Should Be $expected
        }

    }
}

# ============================================================
# RemoveItem（正常系）
# ============================================================
Describe "RemoveItem" {

    It "ファイルを削除できる" {
        $tmp = New-TemporaryFile
        RemoveItem -Path $tmp.FullName
        Test-Path $tmp.FullName | Should Be $false
    }

    It "フォルダを-Recurseで削除できる" {
        $dir = Join-Path $env:TEMP "testdir_rm1"
        if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
        New-Item -ItemType Directory -Path $dir | Out-Null
        New-Item -ItemType File -Path (Join-Path $dir "test.txt") | Out-Null

        RemoveItem -Path $dir -Recurse
        Test-Path $dir | Should Be $false
    }
}

# ============================================================
# Exists
# ============================================================
Describe "Exists" {

    It "存在するパスなら True を返す" {
        $tmp = New-TemporaryFile
        Exists -Path $tmp.FullName | Should Be $true
    }

    It "存在しないパスなら False を返す" {
        Exists -Path "C:\NoSuchFile_$(Get-Random)" | Should Be $false
    }
}

# ============================================================
# IsMatchItemNm（Prefix / Suffix / Contains 全分岐）
# ============================================================
Describe "IsMatchItemNm" {

    It "Prefix が一致する場合 True" {
        $global:MatchPrefix   = "pre"
        $global:MatchSuffix   = ""
        $global:MatchContains = ""
        IsMatchItemNm "pre_abc" | Should Be $true
    }

    It "Prefix が一致しない場合 False" {
        $global:MatchPrefix   = "pre"
        $global:MatchSuffix   = ""
        $global:MatchContains = ""
        IsMatchItemNm "xxx" | Should Be $false
    }

    It "Suffix が一致する場合 True" {
        $global:MatchPrefix   = ""
        $global:MatchSuffix   = "suf"
        $global:MatchContains = ""
        IsMatchItemNm "abc_suf" | Should Be $true
    }

    It "Suffix が一致しない場合 False" {
        $global:MatchPrefix   = ""
        $global:MatchSuffix   = "suf"
        $global:MatchContains = ""
        IsMatchItemNm "xxx" | Should Be $false
    }

    It "Contains が一致する場合 True" {
        $global:MatchPrefix   = ""
        $global:MatchSuffix   = ""
        $global:MatchContains = "mid"
        IsMatchItemNm "aaa_mid_bbb" | Should Be $true
    }

    It "Contains が一致しない場合 False" {
        $global:MatchPrefix   = ""
        $global:MatchSuffix   = ""
        $global:MatchContains = "mid"
        IsMatchItemNm "xxx" | Should Be $false
    }
}

# ============================================================
# GetChildItems（正常系・例外系）
# ============================================================
Describe "GetChildItems" {

    It "Recurse なしで一覧取得できる" {
        $dir = Join-Path $env:TEMP "gci_test1"
        if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
        New-Item -ItemType Directory -Path $dir | Out-Null
        New-Item -ItemType File -Path (Join-Path $dir "a.txt") | Out-Null

        $items = GetChildItems -Path $dir
        $items.Name | Should Be "a.txt"
    }

    It "Recurse ありで一覧取得できる" {
        $dir = Join-Path $env:TEMP "gci_test2"
        if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }

        # 親ディレクトリ作成
        New-Item -ItemType Directory -Path $dir | Out-Null

        # サブディレクトリ作成（Out-Null を使わない）
        $sub = New-Item -ItemType Directory -Path (Join-Path $dir "sub")

        # ファイル作成
        New-Item -ItemType File -Path (Join-Path $sub.FullName "b.txt") | Out-Null

        # テスト対象
        $items = GetChildItems -Path $dir -Recurse | Where-Object { -not $_.PSIsContainer }

        # 検証
        $items.Name | Should Be "b.txt"
    }

    It "存在しないパスなら空配列を返す（ErrCheckFlg=0）" {
        $global:ErrCheckFlg = $false
        $items = GetChildItems -Path "C:\NoSuchDir_$(Get-Random)"
        $items.Count | Should Be 0
    }

    It "存在しないパスなら例外を投げる（ErrCheckFlg=1）" {
        $global:ErrCheckFlg = $true
        { GetChildItems -Path "C:\NoSuchDir_$(Get-Random)" } | Should Throw
    }
}

# ============================================================
# RemoveItem（例外系）
# ============================================================
Describe "RemoveItem (例外)" {

    It "ErrCheckFlg=1 のとき例外を投げる" {
        $global:ErrCheckFlg = $true
        { RemoveItem -Path "C:\NoSuchFile_$(Get-Random)" } | Should Throw
    }

    It "ErrCheckFlg=0 のとき例外を投げない" {
        $global:ErrCheckFlg = $false
        { RemoveItem -Path "C:\NoSuchFile_$(Get-Random)" } | Should Not Throw
    }
}
