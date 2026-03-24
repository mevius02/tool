'変数未宣言エラー設定ON
Option Explicit
'------------------------------------------------------------
'概要 : Sheet[TMPL]作成
'引数 : -
'戻値 : -
'備考 : CSVファイルの入力Sheetを作成するマクロ
'------------------------------------------------------------
Sub CreateSheet_CreateTemplateCsv()
    '処理中描画OFF設定
    SetDrawProcess False

    '●Loopカウンタ [i/j/k]
    Dim i As Long, j As Long, k As Long
    '●作成Sheet
    Dim newSheet As Variant
    '●作成Sheet名
    Dim newSheetNm As String: newSheetNm = "TMPL"
    '●範囲
    Dim rng As Range
    '●図形
    Dim shp As Shape
    '●存在有無
    Dim exists As Boolean

    '■■■■■ Sheet存在チェック■■■■■
    If IsExistsSheet(newSheetNm) Then
        MsgBox "すでにSheet [" & newSheetNm & "] が存在します"
        Exit Sub
    End If

    '■■■■■ Sheet作成&詳細設定(1) ■■■■■
    Set newSheet = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Worksheets(1))
    With newSheet
        '●Sheet名
        .Name = newSheetNm
        '●Tab色
        .Tab.Color = RGB(189, 215, 238) '#BDD7EE
        '●Cell固定文字列
        .Range("A1").Value = "FILE名"
        .Range("A2").Value = "CSV/TSV"
        .Range("A3").Value = "HEADER"
        .Range("A4").Value = "論理/物理"
        .Range("A9").Value = "No"
        .Range("A10:A109").FormulaR1C1 = "=ROW()-9"
        .Range("B1").Value = "ファイル名"
        '●ドロップダウン==========
        .Range("B2").Validation.Delete
        .Range("B2").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="CSV,TSV"
        .Range("B2").Value = "CSV"
        .Range("B3").Validation.Delete
        .Range("B3").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Y,N"
        .Range("B3").Value = "Y"
        .Range("B4").Validation.Delete
        .Range("B4").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="論理,物理"
        .Range("B4").Value = "論理"
        '==============================
        .Range("C1").Value = "No"
        .Range("C2").Value = "項目名(論理)"
        .Range("C3").Value = "項目名(物理)"
        .Range("C4").Value = "ダブクォ付与"
        .Range("C5").Value = "備考1"
        .Range("C6").Value = "備考2"
        .Range("C7").Value = "備考3"
        .Range("C8").Value = "備考4"
        .Range("C9").Value = "備考5"
        .Range("D1:BA1").FormulaR1C1 = "=COLUMN()-3"
        For i = 1 To 50
            .Cells(2, i + 3).Value = "項目名" + CStr(i)
            .Cells(3, i + 3).Value = "COL_NM_" + CStr(i)
        Next i
        .Range("D4:BA4").Value = "Y"
        '●Cell背景色
        .Range("A1:A4").Interior.Color = RGB(255, 192, 0)     '#FFC000
        .Range("B1:B4").Interior.Color = RGB(255, 242, 204)   '#FFF2CC
        .Range("A5:B8").Interior.Color = RGB(128, 128, 128)   '#808080
        .Range("A9:A109").Interior.Color = RGB(217, 217, 217) '#D9D9D9
        .Range("B9").Interior.Color = RGB(217, 217, 217)      '#D9D9D9
        .Range("C1:C9").Interior.Color = RGB(217, 217, 217)   '#D9D9D9
    End With

    '■■■■■ 詳細設定(2) ■■■■■
    newSheet.Activate
    ActiveWindow.Zoom = 70 '表示倍率
    ActiveWindow.DisplayGridlines = False '目盛線
    With newSheet
        '●列幅
        .Columns("A").ColumnWidth = 10
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 13
        .Range("D:BA").ColumnWidth = 15
        '●文字揃え
        .Cells.VerticalAlignment = xlTop                '上揃え
        .Cells.HorizontalAlignment = xlLeft             '左揃え
        .Columns("A").HorizontalAlignment = xlCenter    '文字揃え(横)
        .Range("B1:C9").HorizontalAlignment = xlCenter  '文字揃え(横)
        .Range("D1:BA4").HorizontalAlignment = xlCenter '文字揃え(横)
        .Range("B1").HorizontalAlignment = xlLeft       '文字揃え(横)
        '●フォント
        .Cells.Font.Name = "HGSｺﾞｼｯｸM"
        '●Cell結合
        .Range("B10:C109").Merge Across:=True           'Cell結合(横)
        '●表示形式
        .Range("D10:BA109").NumberFormat = "@"          '表示書式[文字列]
        '●グループ化
        .Range("B1:B1").EntireColumn.Group              '列グループ化
    End With
    '●罫線
    Set rng = Range("C5:BA9")
    With rng.Borders(xlInsideHorizontal)
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    Set rng = Range("B10:BA109")
    With rng.Borders(xlInsideHorizontal)
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    '●ウィンドウ枠の固定
    newSheet.Range("D10").Select
    ActiveWindow.FreezePanes = True
    '●条件付き書式(奇数行背景色)
    Set rng = Range("B10:BA109")
    rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISODD(ROW())").Interior.Color = RGB(242, 242, 242)

    '========== マクロボタン 作成 ==========
    Set shp = newSheet.Shapes.AddShape(msoShapeBevel, 0, 0, Application.CentimetersToPoints(4.2), Application.CentimetersToPoints(1.5))
    With shp
        .Fill.ForeColor.RGB = RGB(91, 155, 213)                            '塗りつぶし色 #5B9BD5
        .line.ForeColor.RGB = RGB(91, 155, 213)                            '枠線色 #5B9BD5
        .line.Weight = 1                                                   '枠線太さ
        .TextFrame2.TextRange.Text = "FILE 作成"                           '表示文字列
        .TextFrame2.TextRange.Font.NameFarEast = "游ゴシック本文"          'フォント
        .TextFrame2.TextRange.Font.Name = "Yu Gothic"                      'フォント
        .TextFrame2.TextRange.Font.Size = 16                               '文字サイズ
        .TextFrame2.TextRange.Font.Bold = msoTrue                          '太字
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(189, 215, 238) '文字色 #BDD7EE
        .TextFrame2.VerticalAnchor = msoAnchorMiddle                       '文字揃え(横)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter   '文字揃え(縱)
        .OnAction = "CreateCsvFile"                                        'マクロ登録
        .Top = 57.5
        .Left = 65
    End With

    '処理中描画ON設定
    SetDrawProcess True
End Sub
'------------------------------------------------------------
'概要 : CSV or TSVファイル作成
'引数 : -
'戻値 : -
'備考 : -
'------------------------------------------------------------
Sub CreateCsvFile()
    '処理中描画OFF設定
    SetDrawProcess False

    '●Loopカウンタ [i/j/k]
    Dim i As Long, j As Long, k As Long
    '●タイムスタンプ
    Dim timestamp As String: timestamp = Format(Now, "YYYYMMDDHHMMSS")
    '●一時保持用(String)
    Dim tmpStr As String

    '●Sheet
    Dim ws As Variant: Set ws = ActiveSheet
    '●Sheet名
    Dim wsNm As String: wsNm = ws.Name
    '●Index(開始列/開始行/終了列/終了行)
    Dim colStartIdx As Long
    Dim rowStartIdx As Long
    Dim colEndIdx As Long
    Dim rowEndIdx As Long
    '●Cel値
    Dim cellVal As String

    '●作成ファイル
    Dim cratFile As Object
    '●ファイルObject
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    '●ファイルPath
    Dim filePath As String: filePath = ThisWorkbook.Path & "\"

    '●1行内容
    Dim line As String

    '●区切り文字(CSV/TSV)
    Dim delimiter As String

    '●ファイル名
    Dim fileNm As String
    '●CSV/TSVフラグ(True=CSV/False=TSV)
    Dim csvTsvFlg As Boolean
    '●Header有無フラグ(True=有/False=無)
    Dim headerFlg As Boolean
    '●論理/物理フラグ(True=論理/False=物理)
    Dim physiLogiFlg As Boolean

    '■■■■■ Index セット ■■■■■
    '●開始列
    colStartIdx = 4
    '●開始行
    rowStartIdx = 10
    '●終了列
    colEndIdx = 4
    Do While Not IsEmpty(ws.Cells(1, colEndIdx).Value2)
        colEndIdx = colEndIdx + 1
    Loop
    colEndIdx = colEndIdx - 1
    '●終了行
    rowEndIdx = 10
    Do While Not IsEmpty(ws.Cells(rowEndIdx, 1).Value2)
        rowEndIdx = rowEndIdx + 1
    Loop
    rowEndIdx = rowEndIdx - 1

    '■■■■■ 設定値 取得 ■■■■■
    Dim settingVals As Variant: settingVals = ws.Range("B1:B4").Value2
    '========== CSV/TSVフラグ ==========
    tmpStr = settingVals(2, 1)
    If tmpStr = "CSV" Then
        csvTsvFlg = True
        delimiter = ","
    ElseIf tmpStr = "TSV" Then
        csvTsvFlg = False
        delimiter = vbTab
    Else
        MsgBox "(B2) に CSV or TSV を入力してください"
        Exit Sub
    End If

    '========== ファイル名 ==========
    '未入力の場合 Sheet名 & 存在する場合 タイムスタンプ付与
    tmpStr = settingVals(1, 1)
    If IsEmpty(tmpStr) Then
        If csvTsvFlg Then
            fileNm = wsNm & ".csv"
        Else
            fileNm = wsNm & ".tsv"
        End If
    Else
        If csvTsvFlg And Right(tmpStr, 4) = ".csv" Or Not csvTsvFlg And Right(tmpStr, 4) = ".tsv" Then
            fileNm = tmpStr
        ElseIf csvTsvFlg Then
            fileNm = tmpStr & ".csv"
        Else
            fileNm = tmpStr & ".tsv"
        End If
    End If
    'ファイル存在チェック
    If fso.FileExists(filePath & fileNm) Then
        fileNm = Left(fileNm, Len(fileNm) - 4) & "_" & timestamp & Right(fileNm, 4)
    End If

    '========== Header有無フラグ ==========
    tmpStr = settingVals(3, 1)
    If tmpStr = "Y" Then
        headerFlg = True
    ElseIf tmpStr = "N" Then
        headerFlg = False
    Else
        MsgBox "(B3) に Y or N を入力してください"
        Exit Sub
    End If

    '========== 論理/物理フラグ ==========
    tmpStr = settingVals(4, 1)
    If tmpStr = "論理" Then
        physiLogiFlg = True
    ElseIf tmpStr = "物理" Then
        physiLogiFlg = False
    Else
        MsgBox "(B4) に 論理 or 物理 を入力してください"
        Exit Sub
    End If

    '■■■■■ 入力値 取得 ■■■■■
    '●Header
    Dim headerVals As Variant
    If physiLogiFlg Then
        headerVals = ws.Range(ws.Range("D2"), ws.Cells(2, colEndIdx)).Value2
    Else
        headerVals = ws.Range(ws.Range("D3"), ws.Cells(3, colEndIdx)).Value2
    End If
    '●ダブルクォート有無
    Dim dblVals As Variant: dblVals = ws.Range(ws.Range("D4"), ws.Cells(4, colEndIdx)).Value2
    '●Contents
    Dim contentsVals As Variant: contentsVals = ws.Range(ws.Range("D10"), ws.Cells(rowEndIdx, colEndIdx)).Value2

    '■■■■■ ファイル作成 ■■■■■
    Set cratFile = fso.CreateTextFile(filePath & fileNm, True)

    '========== Header ==========
    If headerFlg Then
        line = ""
        For i = 1 To UBound(headerVals, 2)
            If Trim$(CStr(dblVals(1, i))) = "Y" Then
                line = line & """" & Trim$(CStr(headerVals(1, i))) & """" & delimiter
            Else
                line = line & Trim$(CStr(headerVals(1, i))) & delimiter
            End If
        Next i
        '末尾区切り除去
        line = Left(line, Len(line) - Len(delimiter))
        cratFile.WriteLine line
    End If

    '========== Contents ==========
    For i = 1 To UBound(contentsVals, 1)
        '進捗度ステータスバー表示
        Application.StatusBar = "Create Contents : " & i & " / " & rowEndIdx - 9

        line = ""
        For j = 1 To UBound(contentsVals, 2)
            cellVal = CStr(contentsVals(i, j))
            If Trim$(CStr(dblVals(1, j))) = "Y" Then
                line = line & """" & ConvertCsvVal(cellVal) & """" & delimiter
            Else
                line = line & ConvertCsvVal(cellVal) & delimiter
            End If
        Next j
        '末尾区切り除去
        line = Left(line, Len(line) - Len(delimiter))
        cratFile.WriteLine line
    Next i

    cratFile.Close

    MsgBox "ファイルを作成しました"

    'ステータスバー初期化
    Application.StatusBar = False

    '処理中描画ON設定
    SetDrawProcess True
End Sub
'------------------------------------------------------------
'概要 : 処理中の描画を設定する
'引数 : flg(True=描画ON / False=描画OFF)
'戻値 : -
'備考 : -
'------------------------------------------------------------
Private Sub SetDrawProcess(ByVal flg As Boolean)
    Application.ScreenUpdating = flg
    Application.DisplayAlerts = flg
End Sub
'------------------------------------------------------------
'概要 : sheetNmのSheetが存在するか判定し、結果を返す
'引数 : sheetNm : Sheet名
'戻値 : 判定結果(True=存在する / False=存在しない)
'備考 : -
'------------------------------------------------------------
Private Function IsExistsSheet(sheetNm As String) As Boolean
    '一時的にエラー処理OFF設定
    On Error Resume Next
    IsExistsSheet = Not Worksheets(sheetNm) Is Nothing
    'エラー処理ON設定
    On Error GoTo 0
End Function
'------------------------------------------------------------
'概要 : 対象が配列に含まれるか判定し、結果を返す
'引数 : word : 検索対象
'     : arr : 検索配列
'戻値 : 判定結果(True=含まれる / False=含まれない)
'備考 : -
'------------------------------------------------------------
Private Function IsInArray(ByVal word As String, ByVal arr As Variant) As Boolean
    Dim arrVal As Variant
    For Each arrVal In arr
        If word = arrVal Then
            IsInArray = True
            Exit Function
        End If
    Next arrVal
End Function
'------------------------------------------------------------
'概要 : 空か判定し、結果を返す
'       strの桁が0 or "" の場合、空判定
'引数 : str : チェック対象文字列
'戻値 : 判定結果(True=空 / False=空でない)
'備考 : -
'------------------------------------------------------------
Private Function IsEmpty(str As String) As Boolean
    If LenB(str) = 0 Or str = "" Then
        IsEmpty = True
    End If
End Function
'------------------------------------------------------------
'概要 : 列Indexから列名(アルファベット)を返す
'引数 : colIdx : 列Index
'戻値 : 列名(アルファベット)
'備考 : -
'------------------------------------------------------------
Private Function GetColNm(colIdx As Long) As String
    GetColNm = Split(Cells(1, colIdx).Address(True, False), "$")(0)
End Function
'------------------------------------------------------------
'概要 : CSV出力値を変換し返す
'引数 : str : 文字列
'戻値 : 変換後出力値
'備考 : -
'------------------------------------------------------------
Private Function ConvertCsvVal(str As String) As String
    '========== 改行 ==========
    str = Replace(str, vbCrLf, "\n")
    str = Replace(str, vbCr, "\n")
    str = Replace(str, vbLf, "\n")
    '========== ダブルクォーテーション ==========
    str = Replace(str, """", """""")
    ConvertCsvVal = str
End Function
