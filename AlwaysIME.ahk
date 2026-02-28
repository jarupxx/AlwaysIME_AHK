; ============================================================
; AlwaysIME_AHK
; アルファベットキー入力時にIMEを自動で「ひらがな」モードへ切替
; AutoHotKey v2 対応
; ============================================================

#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; ============================================================
; 設定：対象アプリの実行ファイル名リスト（小文字で記載）
; 例："notepad++.exe", "code.exe", "chrome.exe"
; ============================================================
global TargetApps := [
    "notepad++.exe",
    "hidemaru.exe",
    "code.exe",
    "chrome.exe",
    "msedge.exe"
]

; ============================================================
; ログ設定
; ============================================================
global LogFilePath := A_ScriptDir "\AlwaysIME_AHK.log"  ; ログファイルパス
global LogMaxLines := 500                                 ; 最大行数（超えたら .old へローテーション）

; ============================================================
; ログ出力関数
; level: "INFO" / "WARN" / "ERROR"
; ============================================================
Log(msg, level := "INFO") {
    timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    line := "[" timestamp "] [" level "] " msg
    try {
        ; 行数超過チェック → ローテーション
        if FileExist(LogFilePath) {
            lineCount := 0
            loop read LogFilePath
                lineCount++
            if (lineCount >= LogMaxLines) {
                archivePath := LogFilePath . ".old"
                if FileExist(archivePath)
                    FileDelete archivePath
                FileMove LogFilePath, archivePath
            }
        }
        FileAppend line "`n", LogFilePath, "UTF-8"
    } catch as e {
        ; ログ書き込み失敗時はサイレントに無視（無限再帰防止）
    }
}

; ============================================================
; IME制御定数
; ============================================================
global WM_IME_CONTROL    := 0x283
global IMC_GETOPENSTATUS := 0x005
global IMC_SETOPENSTATUS := 0x006

; ConversionMode ビットフラグ
global IME_CMODE_NATIVE    := 1   ; ひらがな/カタカナ（日本語）
global IME_CMODE_KATAKANA  := 2   ; カタカナ
global IME_CMODE_FULLSHAPE := 8   ; 全角
global IME_CMODE_ROMAN     := 16  ; ローマ字入力

; MS-IME 2024以降（IME_CMODE_ROMANビットが立たない）
; 00: × IMEが無効                        0000 0000
; 03: カ 半角カナ                         0000 0011
; 08: Ａ 全角英数                         0000 1000
; 09: あ ひらがな（漢字変換モード）       0000 1001
; 11:    全角カナ                         0000 1011
global CModeMS_HankakuKana := IME_CMODE_KATAKANA | IME_CMODE_NATIVE
global CModeMS_ZenkakuEisu := IME_CMODE_FULLSHAPE
global CModeMS_Hiragana    := IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE
global CModeMS_ZenkakuKana := IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE

; ConversionModeの値を人間が読める文字列に変換
CModeName(cmode) {
    if (cmode = CModeMS_Hiragana)
        return "あ ひらがな"
    if (cmode = CModeMS_ZenkakuKana)
        return "ア 全角カナ"
    if (cmode = CModeMS_ZenkakuEisu)
        return "Ａ 全角英数"
    if (cmode = CModeMS_HankakuKana)
        return "ｶ 半角カナ"
    if (cmode = 0)
        return "× IME無効"
    return "? 不明(" cmode ")"
}

; ============================================================
; IME制御関数
; ============================================================

; IMEのON/OFF状態を取得（1=ON, 0=OFF）
IME_GET(WinTitle := "A") {
    hwnd := WinExist(WinTitle)
    if (hwnd = 0) {
        Log("IME_GET: hwnd が見つかりません", "WARN")
        return 0
    }
    hwndIme := DllCall("imm32\ImmGetDefaultIMEWnd", "Ptr", hwnd, "Ptr")
    return DllCall("SendMessage", "Ptr", hwndIme, "UInt", WM_IME_CONTROL, "Ptr", IMC_GETOPENSTATUS, "Ptr", 0, "Ptr")
}

; IMEをONにする（状態が変化した場合のみログ出力）
IME_ON(WinTitle := "A") {
    hwnd := WinExist(WinTitle)
    if (hwnd = 0) {
        Log("IME_ON: hwnd が見つかりません", "WARN")
        return
    }
    hwndIme := DllCall("imm32\ImmGetDefaultIMEWnd", "Ptr", hwnd, "Ptr")
    DllCall("SendMessage", "Ptr", hwndIme, "UInt", WM_IME_CONTROL, "Ptr", IMC_SETOPENSTATUS, "Ptr", 1, "Ptr")
    Log("IME状態変化: OFF → ON")
}

; IMEをひらがなモードに設定（ConversionModeが変化した場合のみログ出力）
IME_SetHiragana(WinTitle := "A") {
    hwnd := WinExist(WinTitle)
    if (hwnd = 0) {
        Log("IME_SetHiragana: hwnd が見つかりません", "WARN")
        return
    }
    hImc := DllCall("imm32\ImmGetContext", "Ptr", hwnd, "Ptr")
    if (hImc = 0) {
        Log("IME_SetHiragana: ImmGetContext 失敗 (hwnd=" hwnd ")", "WARN")
        return
    }
    ; 変更前のConversionModeを取得
    beforeMode := 0
    DllCall("imm32\ImmGetConversionStatus", "Ptr", hImc, "UInt*", &beforeMode, "UInt*", 0)
    ; ConversionMode: CModeMS_Hiragana = ひらがな（MS-IME 2024以降）
    DllCall("imm32\ImmSetConversionStatus", "Ptr", hImc, "UInt", CModeMS_Hiragana, "UInt", 0)
    DllCall("imm32\ImmReleaseContext", "Ptr", hwnd, "Ptr", hImc)
    ; 変化があった場合のみログ出力
    if (beforeMode != CModeMS_Hiragana)
        Log("IME状態変化: " CModeName(beforeMode) " → " CModeName(CModeMS_Hiragana))
}

; 対象アプリかどうかチェック
IsTargetApp() {
    try {
        processName := WinGetProcessName("A")
        processName := StrLower(processName)
        for app in TargetApps {
            if (processName = app)
                return true
        }
    }
    return false
}

; IMEをONかつひらがなモードにして、キーを送信する共通処理
EnsureHiraganaAndSend(key) {
    if !IsTargetApp()
        return false   ; 対象外アプリ → 通常処理に任せる

    ; キー入力ログ（アプリ名付き）
    try {
        processName := WinGetProcessName("A")
        Log("キー入力: `"" key "`" ← " processName)
    }

    ; IMEがOFFなら先にONにする（IME_ON内で状態変化ログを出力）
    if (IME_GET("A") = 0) {
        IME_ON("A")
        Sleep 30
    }

    ; ひらがなモードに切替（変化があればIME_SetHiragana内でログ出力）
    IME_SetHiragana("A")
    Sleep 10

    ; キーを送信
    SendInput key
    return true
}

; ============================================================
; アルファベットキーのフック（a-z / A-Z）
; 対象アプリでのみ動作
; ============================================================

#HotIf IsTargetApp()

a::EnsureHiraganaAndSend("a")
b::EnsureHiraganaAndSend("b")
c::EnsureHiraganaAndSend("c")
d::EnsureHiraganaAndSend("d")
e::EnsureHiraganaAndSend("e")
f::EnsureHiraganaAndSend("f")
g::EnsureHiraganaAndSend("g")
h::EnsureHiraganaAndSend("h")
i::EnsureHiraganaAndSend("i")
j::EnsureHiraganaAndSend("j")
k::EnsureHiraganaAndSend("k")
l::EnsureHiraganaAndSend("l")
m::EnsureHiraganaAndSend("m")
n::EnsureHiraganaAndSend("n")
o::EnsureHiraganaAndSend("o")
p::EnsureHiraganaAndSend("p")
q::EnsureHiraganaAndSend("q")
r::EnsureHiraganaAndSend("r")
s::EnsureHiraganaAndSend("s")
t::EnsureHiraganaAndSend("t")
u::EnsureHiraganaAndSend("u")
v::EnsureHiraganaAndSend("v")
w::EnsureHiraganaAndSend("w")
x::EnsureHiraganaAndSend("x")
y::EnsureHiraganaAndSend("y")
z::EnsureHiraganaAndSend("z")

; 大文字（Shift+アルファベット）も対応
+a::EnsureHiraganaAndSend("A")
+b::EnsureHiraganaAndSend("B")
+c::EnsureHiraganaAndSend("C")
+d::EnsureHiraganaAndSend("D")
+e::EnsureHiraganaAndSend("E")
+f::EnsureHiraganaAndSend("F")
+g::EnsureHiraganaAndSend("G")
+h::EnsureHiraganaAndSend("H")
+i::EnsureHiraganaAndSend("I")
+j::EnsureHiraganaAndSend("J")
+k::EnsureHiraganaAndSend("K")
+l::EnsureHiraganaAndSend("L")
+m::EnsureHiraganaAndSend("M")
+n::EnsureHiraganaAndSend("N")
+o::EnsureHiraganaAndSend("O")
+p::EnsureHiraganaAndSend("P")
+q::EnsureHiraganaAndSend("Q")
+r::EnsureHiraganaAndSend("R")
+s::EnsureHiraganaAndSend("S")
+t::EnsureHiraganaAndSend("T")
+u::EnsureHiraganaAndSend("U")
+v::EnsureHiraganaAndSend("V")
+w::EnsureHiraganaAndSend("W")
+x::EnsureHiraganaAndSend("X")
+y::EnsureHiraganaAndSend("Y")
+z::EnsureHiraganaAndSend("Z")

#HotIf

; ============================================================
; 起動・終了ログ／トレイメニュー設定
; ============================================================
Log("=== AlwaysIME_AHK 起動 === (LogFile: " LogFilePath ")")
OnExit((*) => Log("=== AlwaysIME_AHK 終了 ==="))

A_TrayMenu.Delete()
A_TrayMenu.Add("対象アプリ一覧を表示", ShowTargetApps)
A_TrayMenu.Add("対象アプリを追加", AddTargetApp)
A_TrayMenu.Add()
A_TrayMenu.Add("ログファイルを開く", OpenLogFile)
A_TrayMenu.Add("ログファイルを削除", DeleteLogFile)
A_TrayMenu.Add()
A_TrayMenu.Add("終了", (*) => ExitApp())
A_TrayMenu.Default := "対象アプリ一覧を表示"
TraySetIcon(A_AhkPath, 2)

; ============================================================
; トレイメニュー関数
; ============================================================

ShowTargetApps(*) {
    appList := ""
    for i, app in TargetApps
        appList .= i ". " app "`n"
    MsgBox appList, "現在の対象アプリ一覧", "OK"
}

AddTargetApp(*) {
    result := InputBox("追加するアプリの実行ファイル名を入力してください`n例: notepad++.exe", "対象アプリを追加")
    if (result.Result = "OK" && result.Value != "") {
        newApp := StrLower(Trim(result.Value))
        for app in TargetApps {
            if (app = newApp) {
                MsgBox newApp " はすでに登録されています。", "情報"
                return
            }
        }
        TargetApps.Push(newApp)
        Log("対象アプリを追加: " newApp)
        MsgBox newApp " を追加しました。", "完了"
    }
}

OpenLogFile(*) {
    if FileExist(LogFilePath)
        Run LogFilePath
    else
        MsgBox "ログファイルがまだ作成されていません。`n" LogFilePath, "情報"
}

DeleteLogFile(*) {
    if !FileExist(LogFilePath) {
        MsgBox "ログファイルが存在しません。", "情報"
        return
    }
    result := MsgBox("ログファイルを削除しますか？`n" LogFilePath, "確認", "YesNo")
    if (result = "Yes") {
        FileDelete LogFilePath
        MsgBox "削除しました。", "完了"
    }
}
