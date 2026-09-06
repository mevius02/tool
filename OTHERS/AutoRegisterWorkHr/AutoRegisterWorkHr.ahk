#Requires AutoHotkey v2.0
;; ============================================================
;;  TeamSpilitの勤怠自動入力(1か月分)
;;   ・WEBブラウザを開いて実行し、自動で(登録 or クリア)をする
;;   ・勤務時間、休憩時間は変数で日単位に設定可能(下の時間指定エリアを編集する)
;;      (変数一括作成Excelツールもあり)
;;   ・処理開始はAHKをダブルクリック起動後[F1]キーを押す
;;   ・強制終了は[Esc]キーを押す
;;   ・処理時間は、普通のPCで約1分ほど
;;
;;  手順 (ユーザー作業は1～12,15)
;;    1, 勤怠入力画面まで進める
;;    2, このAHKファイルをダブルクリックで起動する
;;    3, 勤怠入力画面(ブラウザ)を最前面に表示する
;;    4, [F1]キーを押してAHK処理を開始する
;;    5, ブラウザが(表示最大化 & 拡大率100% & 一番上にスクロール)される
;;    6, 実行するモードを選択する (登録 or クリア)
;;    7, 現在の表示行数を選択する (15未満 or 15以上)
;;        (10行未満の場合でも[F11]の全画面表示で10行以上見えればOK)
;;    8, 登録月の日数を選択する (28 or 29 or 30 or 31 or その他)
;;    9, POPUPが閉じた判定をするため、ヘッダーの白色部分以外をクリックでX/Y軸を指定する
;;        (紺色の部分などをクリックする)
;;    10, 各行の[出社]セル座標を特定するため、X軸をクリックで指定する
;;         ([出社]セルの中央あたりをクリックする)
;;    11, 各行の[出社]セル座標を特定するため、Y軸をクリックで指定する
;;         (1行目の[申請]セルの十字ボタンの中央あたりをクリックする)
;;           → 5行目まで繰り返し入力受付
;;    12, 1～5行分の平均値を行間とする (クリックした座標の平均値なので、クリック場所が適当だとずれる)
;;    13, 自動で10 or 15日まで登録を繰り返す
;;         登録  ：([出社]セルクリック → POPUP開く → [勤務時間/休憩時間]セット → [登録]ボタンクリック → POPUP閉じる)
;;         クリア：([出社]セルクリック → POPUP開く → TAB移動 → [リセット]ボタン → [確認]ボタンクリック → Shift+TAB移動 → [登録]ボタンクリック → POPUP閉じる)
;;    14, 下にnホイール分スクロールされる
;;    15, 再度、n日目のY軸をクリックで指定する
;;         (n行目の[申請]セルの十字ボタンの中央あたりをクリックする)
;;    16, 自動でn～31日まで登録を繰り返す
;; ------------------------------------------------------------
;;    Ex, 土日祝の判定は行の背景色が土日祝の色かで判定
;;    Ex, POPUP表示したかの判定は、POPUPの[出社]入力項目にフォーカスがあたったかで判定
;;    Ex, POPUP閉じたかの判定は、一番上にスクロールし、ヘッダー色がオーバーレイの色かで判定
;; ============================================================
;; 画面全体の絶対座標を基準にする設定
CoordMode("Mouse", "Screen")
CoordMode("ToolTip", "Screen")
;; ==================================================
;;  変数 ([勤務/休憩]時間指定エリア)
;; ==================================================
;; 勤務(開始/終了)時間
startWork := "09:00"
endWork := "18:00"
;; 休憩(開始/終了)時間
startRest := "12:00"
endRest := "13:00"
;; ==================================================
;;  変数 (利用ユーザーも変更可)
;; ==================================================
;; スクロール回数
;; (入力後半はブラウザをスクロールしないと行が見えないので、スクロールする)
;; (1スクロールの行数設定がPC初期値以外の場合、任意調整必要)
scrollCnt11 := 3 ;; 11-20行目用
scrollCnt16 := 4 ;; 16-31行目用
scrollCnt21 := 6 ;; 21-31行目用
;; ========== 実行中に変更可 ==========
;; クリアモードフラグ(true:登録/false:クリア)
clearFlg := false
;; 10行単位フラグ
;; (入力行がスクロールしないと見えない場合、画面のサイズ、拡大率によって10行までしか見えない用)
;; (true:10行単位/false:15行単位)
row10Flg := true
;; 月の日数
dayCnt := 0
;; ==================================================
;;  変数 (開発ユーザーの場合、変更可)
;; ==================================================
;; [出社]座標
colX := 0
rowY1 := 0
rowY2 := 0
rowY3 := 0
rowY4 := 0
rowY5 := 0
rowY10 := 0
rowY16 := 0
rowY20 := 0
;; 土日祝の背景色
holidayColor1 := 0xDEEFFF ;;#DEEFFF 土/祝
holidayColor2 := 0xEFCFBD ;;#EFCFBD 日
;; オーバーレイ座標 / 色
rayX := 0
rayY := 0
rayColor := 0x256499 ;#256499
;; イベント発火防止用
captureClick := false
clickX := 0
clickY := 0
;; Tips座標
tipsX := 20
tipsY := 20
;; スキップフラグ
skipFlg := true

;; ==================================================
;;  ESC：アプリ強制終了
;; ==================================================
Esc:: {
    MsgBox("AutoHotkeyを強制終了します", "FINISH", "T1") ;; 1秒だけ通知を表示
    ExitApp()
}

;; ==================================================
;;  F1：全体処理 開始
;; ==================================================
F1:: {
    global clearFlg := false
    global row10Flg := true
    global dayCnt := 0
    global skipFlg := true

    ;; ブラウザが最前面でアクティブでない場合
    GroupAdd("Browsers", "ahk_exe chrome.exe")
    GroupAdd("Browsers", "ahk_exe msedge.exe")
    GroupAdd("Browsers", "ahk_exe firefox.exe")
    activeProc := WinGetProcessName("A")
    if !(activeProc = "chrome.exe" || activeProc = "msedge.exe" || activeProc = "firefox.exe") {
        MsgBox("Webブラウザをアクティブにしてから [F1] を押して開始して下さい", "ERROR", "iconX")
        ExitApp()
    }

    ;; === ブラウザ事前準備 ==========
    ;; フルスクリーン表示 &  拡大率100% セット & 一番上にスクロール
    WinMaximize("A")
    Sleep(20)
    Send("^0")
    Sleep(20)
    Send("{Home}")

    ;; === 入力受付 ==========
    ;; 処理モード
    SelectForm_ProcessMode(&clearFlg)
    ;; 表示行数
    SelectForm_RowCnt(&row10Flg)
    ;; 月の日数
    SelectForm_DayCnt(&dayCnt)
    ;; ヘッダー背景色 (XY座標)
    ToolTip(GetMsg(1, true, 7, 1, 0), tipsX, tipsY)
    SetClickPos(&rayX, &rayY)
    ;; [出社]列の縦位置 (X座標)
    ToolTip(GetMsg(2, true, 7, 2, 0), tipsX, tipsY)
    SetClickPos(&colX, &colY)
    ;; 1～5行目の横位置 (Y座標)
    ToolTip(GetMsg(3, true, 7, 3, 1), tipsX, tipsY)
    SetClickPos(&rowX1, &rowY1)
    ToolTip(GetMsg(3, true, 7, 4, 2), tipsX, tipsY)
    SetClickPos(&rowX2, &rowY2)
    ToolTip(GetMsg(3, true, 7, 5, 3), tipsX, tipsY)
    SetClickPos(&rowX3, &rowY3)
    ToolTip(GetMsg(3, true, 7, 6, 4), tipsX, tipsY)
    SetClickPos(&rowX4, &rowY4)
    ToolTip(GetMsg(3, true, 7, 7, 5), tipsX, tipsY)
    SetClickPos(&rowX5, &rowY5)
    ;; ツールチップ消去
    ToolTip()

    ;; 行間 セット (平均値)
    lineSpace := ((rowY2 - rowY1) + (rowY3 - rowY2) + (rowY4 - rowY3) + (rowY5 - rowY4)) / 4

    ;; === メイン処理 ==========
    ;; 1か月の日数分Loop
    loop dayCnt {
        ;; 10行単位で入力繰り返しの場合
        if (row10Flg) {
            if (A_Index == 11 or A_Index == 21) {
                ;; 一番上にスクロール
                Send("{Home}")
                Sleep(500) ;; 0.5秒
            }
            if (A_Index == 11 or A_Index == 21 or !skipFlg) {
                scrollCnt := A_Index <= 20 ? scrollCnt11 : scrollCnt21
                ;; 下スクロール
                loop scrollCnt {
                    Send("{WheelDown}")
                    Sleep(20) ;; 0.02秒
                }
            }
            ;; 10行以内の場合
            if (A_Index <= 10) {
                rowY := rowY1 + ((A_Index - 1) * lineSpace)
                MouseMove(colX, rowY) ;; 不安定動作解消用
                skipFlg := true
                ;; 11-20行の場合
            } else if (A_Index <= 20) {
                ;; 11行目の横位置 指定(Y座標)
                if (A_Index == 11) {
                    ToolTip(GetMsg(3, false, 0, 0, A_Index), tipsX, tipsY)
                    SetClickPos(&rowX11, &rowY11)
                    ToolTip()
                }
                rowY := rowY11 + ((A_Index - 11) * lineSpace)
                skipFlg := false
                MouseMove(colX, rowY) ;; 不安定動作解消用
                ;; 21行以上の場合
            } else {
                ;; 21行目の横位置 指定(Y座標)
                if (A_Index == 21) {
                    ToolTip(GetMsg(3, false, 0, 0, A_Index), tipsX, tipsY)
                    SetClickPos(&rowX21, &rowY21)
                    ToolTip()
                }
                rowY := rowY21 + ((A_Index - 21) * lineSpace)
                skipFlg := false
                MouseMove(colX, rowY) ;; 不安定動作解消用
            }
        } else {
            if (A_Index == 16) {
                ;; 一番上にスクロール
                Send("{Home}")
                Sleep(500) ;; 0.5秒
            }
            if (A_Index == 16 or !skipFlg) {
                ;; 下スクロール
                loop scrollCnt16 {
                    Send("{WheelDown}")
                    Sleep(20) ;; 0.02秒
                }
            }
            ;; 15行以内の場合
            if (A_Index <= 15) {
                rowY := rowY1 + ((A_Index - 1) * lineSpace)
                skipFlg := true
                MouseMove(colX, rowY) ;; 不安定動作解消用
            } else {
                ;; 16行目の横位置 指定(Y座標)
                if (A_Index == 16) {
                    ToolTip(GetMsg(3, false, 0, 0, A_Index), tipsX, tipsY)
                    SetClickPos(&rowX16, &rowY16)
                    ToolTip()
                }
                rowY := rowY16 + ((A_Index - 16) * lineSpace)
                skipFlg := false
                MouseMove(colX, rowY) ;; 不安定動作解消用
            }
        }
        ;; n行目の[出社]セルの背景色 取得
        tgtColor := PixelGetColor(colX, rowY, "RGB")
        ;; [出社]入力項目の有無判定 ([出社]セルの背景色が土日祝の色ではない場合)
        if (tgtColor != holidayColor1 and tgtColor != holidayColor2) {
            ;; === [出社]セルをクリック ==========
            Click(colX, rowY)
            Sleep(300) ;; 0.3秒
            success := false
            ;; === POPUP表示 判定 ==========
            ;; リトライ(0.25秒 × 120回 = 約30秒)
            loop 120 {
                ;; クリップボード クリア
                A_Clipboard := ""
                ;; [出社]入力欄に入力してコピー
                Send("^a{Backspace}")
                Sleep(50) ;; 0.05秒
                Send("{Text}" startWork)
                Sleep(50) ;; 0.05秒
                Send("^a^c")
                ;; コピー可否 判定
                if ClipWait(0.05) {
                    ;; 入力した文字列と一致する場合(入力欄にコピペできた場合)
                    if (A_Clipboard == startWork) {
                        success := true
                        break
                    }
                }
                Sleep(250) ;; 0.25秒
            }
            ;; === POPUP表示 不可の場合 ==========
            if (!success) {
                MsgBox(A_Index "日目：POPUPが表示されませんでした", "ERROR", "IconX")
                ExitApp()
            }

            ;; === 入力処理 ==========
            ;; クリアの場合
            if (clearFlg) {
                ;; TAB移動 → リセットボタン → 確認ボタン
                loop 14 {
                    Send("{Tab}")
                }
                Sleep(250) ;; 0.25秒
                Send("{Enter}")
                Sleep(250) ;; 0.25秒
                Send("{Enter}")
                Sleep(250) ;; 0.25秒
                ;; Shift+TAB移動 → 登録ボタン → 確認ボタン
                loop 2 {
                    Send("+{Tab}")
                }
                Send("{Enter}")
                Sleep(250) ;; 0.25秒
                Send("{Enter}")
            } else {
                ;; TAB移動 → 勤務終了
                Send("{Tab}")
                Send("{Text}" endWork)
                ;; TAB移動 → 休憩開始
                Send("{Tab}")
                Send("{Text}" startRest)
                ;; TAB移動 → 休憩終了
                Send("{Tab}")
                Send("{Text}" endRest)
                ;; TAB移動 → 登録ボタン
                loop 9 {
                    Send("{Tab}")
                }
                Send("{Enter}")
            }
            ;; 一番上にスクロール
            Sleep(500) ;; 0.5秒
            Send("{Home}")
            ;; === POPUP閉じる 判定 ==========
            ;; リトライ(0.25秒 × 120回 = 約30秒)
            success := false
            loop 120 {
                ;; ヘッダーの背景色 取得
                tgtColor := PixelGetColor(rayX, rayY, "RGB")
                ;; オーバーレイ無しの背景色と一致する場合
                if (rayColor == tgtColor) {
                    success := true
                    break
                }
                Sleep(250) ;; 0.25秒
            }
            ;; === POPUP閉じる 不可の場合 ==========
            if (!success) {
                MsgBox(A_Index "日目：POPUPが閉じませんでした", "ERROR", "IconX")
                ExitApp()
            }
        } else {
            skipFlg := true
        }
    }
    if (clearFlg) {
        MsgBox("クリア完了しました！！", "FINISH")
    } else {
        MsgBox("登録完了しました！！", "FINISH")
    }
    ExitApp()
}

;; ==================================================
;;  処理モード 受付
;; ==================================================
SelectForm_ProcessMode(&clearFlg) {
    dayGui := Gui("-SysMenu +AlwaysOnTop", "")
    dayGui.SetFont("s10", "Noto Sans JP")
    dayGui.AddText("x0 y0 w250 Center", "実行するモードを選択して下さい")
    btn10 := dayGui.AddButton("x30 y30 w80 h40", "登録")
    btn15 := dayGui.AddButton("x140 y30 w80 h40", "クリア")
    selectFlg := 0
    btn10.OnEvent("Click", (*) => selectFlg := 1)
    btn15.OnEvent("Click", (*) => selectFlg := 2)
    dayGui.Show("w250 h90")
    while (selectFlg = 0) {
        Sleep(50)
    }
    dayGui.Destroy()
    clearFlg := selectFlg == 1 ? false : true
}

;; ==================================================
;;  表示行数 受付
;; ==================================================
SelectForm_RowCnt(&row10Flg) {
    dayGui := Gui("-SysMenu +AlwaysOnTop", "")
    dayGui.SetFont("s10", "Noto Sans JP")
    dayGui.AddText("x0 y0 w250 Center", "現在表示されている行数は`n１５行以上ですか？")
    btn10 := dayGui.AddButton("x30 y50 w80 h40", "15未満")
    btn15 := dayGui.AddButton("x140 y50 w80 h40", "15以上")
    selectFlg := 0
    btn10.OnEvent("Click", (*) => selectFlg := 1)
    btn15.OnEvent("Click", (*) => selectFlg := 2)
    dayGui.Show("w250 h110")
    while (selectFlg = 0) {
        Sleep(50)
    }
    dayGui.Destroy()
    row10Flg := selectFlg == 1 ? true : false
}

;; ==================================================
;;  月の日数 受付
;; ==================================================
SelectForm_DayCnt(&dayCnt) {
    dayGui := Gui("-SysMenu +AlwaysOnTop", "")
    dayGui.SetFont("s10", "Noto Sans JP")
    dayGui.AddText("x25 y0 w200 Center", "月の日数を選択して下さい")
    btn28 := dayGui.AddButton("x40 y30 w80 h40", "28日")
    btn29 := dayGui.AddButton("x130 y30 w80 h40", "29日")
    btn30 := dayGui.AddButton("x40 y80 w80 h40", "30日")
    btn31 := dayGui.AddButton("x130 y80 w80 h40", "31日")
    selectDayCnt := 0
    btn28.OnEvent("Click", (*) => selectDayCnt := 28)
    btn29.OnEvent("Click", (*) => selectDayCnt := 29)
    btn30.OnEvent("Click", (*) => selectDayCnt := 30)
    btn31.OnEvent("Click", (*) => selectDayCnt := 31)
    dayGui.Show("w250 h140")
    while (selectDayCnt = 0) {
        Sleep(50)
    }
    dayGui.Destroy()
    dayCnt := Number(selectDayCnt)
}

;; ==================================================
;;  クリックした座標を取得する(引数にセットする)
;;   引数1 : x : X軸座標
;;   引数2 : y : Y軸座標
;; ==================================================
SetClickPos(&x, &y) {
    global captureClick, clickX, clickY
    ;; クリック受付 開始
    captureClick := true
    ;; ユーザーがクリックするまで待機
    while (captureClick) {
        Sleep(10) ;; 0.01秒
    }
    ;; 座標を返す
    x := clickX
    y := clickY
}
;; ==================================================
;;  座標受付クリックのWEBブラウザイベント発火防止
;; ==================================================
#HotIf captureClick
LButton:: {
    global captureClick, clickX, clickY
    ;; クリック座標 取得
    MouseGetPos(&clickX, &clickY)
    ;; クリック受付 終了
    captureClick := false
    ;; ボタンが離れるまで待機
    KeyWait("LButton", "U")
    ;; returnでブラウザイベント無し
    return
}
#HotIf

;; ==================================================
;;  メッセージを取得
;; ==================================================
GetMsg(msgNo, args1, args2, args3, args4) {
    ;; 準備
    headerFlg := args1   ;; ヘッダー有無
    allStepCnt := args2  ;; 全STEP数
    crntStepCnt := args3 ;; 現在STEP数
    rowNum := args4      ;; 行数
    mk1 := ""            ;; マーク1
    mk2 := ""            ;; マーク2
    loop crntStepCnt - 1 {
        mk1 := mk1 " ■"
    }
    loop allStepCnt - crntStepCnt + 1 {
        mk2 := mk2 " □"
    }
    ;; 作成
    msg := "◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆`n"
    if (headerFlg) {
        msg .= "◆ STEP" NumberToFullWidth(crntStepCnt) "/" NumberToFullWidth(allStepCnt) " [" mk1 mk2 " ]　`n"
    }
    switch msgNo {
        case 1:
            msg .= "◆ 勤務表のヘッダーの白色以外の部分をどこでもいいのでクリックして下さい (X/Y座標の指定)　`n"
            msg .= "◆ ※ 終了する場合は [ESC] キーを押して下さい　`n"
        case 2:
            msg .= "◆ 勤務表の [出社] 列の位置をクリックで教えて下さい (X座標の指定)　`n"
            msg .= "◆ [出社]列の中央あたりをクリックする　`n"
            msg .= "◆ ※ 終了する場合は [ESC] キーを押して下さい　`n"
        case 3:
            msg .= "◆ 勤務表の" NumberToFullWidth(rowNum) "日目の行の位置をクリックで教えて下さい (Y座標の指定)　`n"
            msg .= "◆ " NumberToFullWidth(rowNum) "行目の[申請]セルの十字ボタンの中央あたりをクリックする　`n"
            msg .= "◆ ※ 終了する場合は [ESC] キーを押して下さい　`n"
    }
    msg .= "◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆ ◆`n"
    return msg
}

;; ==================================================
;;  半角数字を全角数字に変換する
;; ==================================================
NumberToFullWidth(inputVal) {
    ;; 変換用のマップを定義
    fullWidMap := Map(
        "0", "０", "1", "１", "2", "２", "3", "３", "4", "４",
        "5", "５", "6", "６", "7", "７", "8", "８", "9", "９")
    result := ""
    ;; 1文字ずつループして全角に置換
    loop parse, String(inputVal) {
        if fullWidMap.Has(A_LoopField) {
            result .= fullWidMap[A_LoopField]
        } else {
            ;; 数字以外の文字はそのまま結合
            result .= A_LoopField
        }
    }
    return result
}
