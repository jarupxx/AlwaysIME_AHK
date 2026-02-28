; ============================================================
; AlwaysIME_AHK
; キー入力時にIMEを自動制御する常駐スクリプト
; AutoHotKey v2 対応
; ============================================================

#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; ============================================================
; ログ設定
; ============================================================
global LogFilePath := A_ScriptDir "\AlwaysIME_AHK.log"
global LogMaxLines := 500

Log(msg, level := "INFO") {
    timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    line := "[" timestamp "] [" level "] " msg
    try {
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
    } catch {
    }
}

; ============================================================
; IME制御定数
; ============================================================
global WM_IME_CONTROL    := 0x283
global IMC_GETOPENSTATUS := 0x005
global IMC_SETOPENSTATUS := 0x006

global IME_CMODE_NATIVE    := 1
global IME_CMODE_KATAKANA  := 2
global IME_CMODE_FULLSHAPE := 8
global IME_CMODE_ROMAN     := 16

global CModeMS_HankakuKana := IME_CMODE_KATAKANA | IME_CMODE_NATIVE
global CModeMS_ZenkakuEisu := IME_CMODE_FULLSHAPE
global CModeMS_Hiragana    := IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE
global CModeMS_ZenkakuKana := IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE

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
; アプリ・タイトルごとの設定
; ============================================================

; IMEを制御しないアプリ（完全一致・小文字）
global IgnoreApps := [
    ; "example.exe",
]

; キー入力のたびにIME-OFFを強制するアプリ（完全一致・小文字）
global ForceOffApps := [
    ; "putty.exe",
]

; タイトルの一部にマッチしたらIME-OFFにするパターン（正規表現）
; アプリ問わず全体で有効
global TitleOffPatterns := [
    "\.cs$",
    "\.js$",
    "\.ts$",
    "\.py$",
    "\.ahk$",
]

; タイトル変化の検出から除外するタグパターン（正規表現）
; マッチした部分を取り除いてからタイトルを比較する
global TitleIgnoreTags := [
    "\s*[\(\[『「][\*●○＊]?更新[済]?[\)\]』」]",
    "\s*\*$",
    "\s*•$",
]

; ============================================================
; IME制御再開の判定に使う状態
; ============================================================
global LastProcessName := ""
global LastWindowTitle := ""
global IMEControlled   := false

; 未入力タイムアウト（ミリ秒）
global IdleTimeoutMs := 5 * 60 * 1000

; ============================================================
; タイトルからIgnoreTagsを取り除いて正規化する
; ============================================================
NormalizeTitle(title) {
    result := title
    for pattern in TitleIgnoreTags
        result := RegExReplace(result, pattern, "")
    return Trim(result)
}

; ============================================================
; IMEをONにする
; ============================================================
IME_ON(hwnd) {
    hwndIme := DllCall("imm32\ImmGetDefaultIMEWnd", "Ptr", hwnd, "Ptr")
    if (hwndIme = 0) {
        Log("IME_ON: ImmGetDefaultIMEWnd 失敗 (hwnd=" hwnd ")", "WARN")
        return
    }
    DllCall("SendMessage", "Ptr", hwndIme, "UInt", WM_IME_CONTROL, "Ptr", IMC_SETOPENSTATUS, "Ptr", 1, "Ptr")
    Log("IME状態変化: → ON")
}

; ============================================================
; IMEをOFFにする
; ============================================================
IME_OFF(hwnd) {
    hwndIme := DllCall("imm32\ImmGetDefaultIMEWnd", "Ptr", hwnd, "Ptr")
    if (hwndIme = 0) {
        Log("IME_OFF: ImmGetDefaultIMEWnd 失敗 (hwnd=" hwnd ")", "WARN")
        return
    }
    DllCall("SendMessage", "Ptr", hwndIme, "UInt", WM_IME_CONTROL, "Ptr", IMC_SETOPENSTATUS, "Ptr", 0, "Ptr")
    Log("IME状態変化: → OFF")
}

; ============================================================
; ConversionModeをひらがなにセット（変化があればログ）
; ============================================================
IME_SetHiragana(hwnd) {
    hImc := DllCall("imm32\ImmGetContext", "Ptr", hwnd, "Ptr")
    if (hImc = 0) {
        Log("IME_SetHiragana: ImmGetContext 失敗 (hwnd=" hwnd ")", "WARN")
        return
    }
    beforeMode := 0
    DllCall("imm32\ImmGetConversionStatus", "Ptr", hImc, "UInt*", &beforeMode, "UInt*", 0)
    DllCall("imm32\ImmSetConversionStatus", "Ptr", hImc, "UInt", CModeMS_Hiragana, "UInt", 0)
    DllCall("imm32\ImmReleaseContext", "Ptr", hwnd, "Ptr", hImc)
    if (beforeMode != CModeMS_Hiragana)
        Log("IME状態変化: " CModeName(beforeMode) " → " CModeName(CModeMS_Hiragana))
}

; ============================================================
; IME制御を再開すべきか判定し、状態を更新する
; 戻り値: true = 制御を実行すべき / false = 制御スキップ
; ============================================================
ShouldControl(processName, normalizedTitle) {
    global LastProcessName, LastWindowTitle, IMEControlled, IdleTimeoutMs

    idleMs := A_TimeIdlePhysical
    if (idleMs >= IdleTimeoutMs) {
        Log("再開トリガー: 未入力タイムアウト (" Round(idleMs/1000) "秒)")
        IMEControlled := false
    }

    if (processName != LastProcessName) {
        Log("再開トリガー: アプリ変化 (" LastProcessName " → " processName ")")
        IMEControlled := false
        LastProcessName := processName
        LastWindowTitle := normalizedTitle
        return true
    }

    if (normalizedTitle != LastWindowTitle) {
        Log("再開トリガー: タイトル変化 (`"" LastWindowTitle "`" → `"" normalizedTitle "`")")
        IMEControlled := false
        LastWindowTitle := normalizedTitle
        return true
    }

    if IMEControlled
        return false

    return true
}

; ============================================================
; キー入力ごとに呼ばれるメイン処理
; ============================================================
HandleKeyInput(key) {
    hwnd := WinExist("A")
    if (hwnd = 0)
        return

    processName := StrLower(WinGetProcessName("A"))
    rawTitle    := WinGetTitle("A")
    normTitle   := NormalizeTitle(rawTitle)

    Log("キー入力: `"" key "`" app=" processName " title=`"" normTitle "`"")

    ; IgnoreApps: 制御しない
    for app in IgnoreApps {
        if (processName = app) {
            Log("スキップ: IgnoreApps に一致 (" processName ")")
            SendInput key
            return
        }
    }

    ; ForceOffApps: 再開条件が揃ったときだけIME-OFFにして以後維持
    for app in ForceOffApps {
        if (processName = app) {
            if ShouldControl(processName, normTitle) {
                IME_OFF(hwnd)
                global IMEControlled := true
                global LastProcessName := processName
                global LastWindowTitle := normTitle
            }
            SendInput key
            return
        }
    }

    ; TitleOffPatterns: タイトルにマッチしたらIME-OFF
    for pattern in TitleOffPatterns {
        if RegExMatch(rawTitle, pattern) {
            Log("IME-OFF: TitleOffPattern に一致 (`"" pattern "`")")
            IME_OFF(hwnd)
            global IMEControlled := true
            global LastProcessName := processName
            global LastWindowTitle := normTitle
            SendInput key
            return
        }
    }

    ; IME制御を再開すべきか判定
    if !ShouldControl(processName, normTitle) {
        SendInput key
        return
    }

    ; IME-ON かつ ひらがなモードに設定
    IME_ON(hwnd)
    IME_SetHiragana(hwnd)

    global IMEControlled := true
    global LastProcessName := processName
    global LastWindowTitle := normTitle

    SendInput key
}

; ============================================================
; 全キーのフック（a-z / A-Z）
; #InputLevel 1 により、SendLevel 0（デフォルト）で送信した
; キーはこのホットキーに再捕捉されない（無限ループ防止）
; ============================================================

#InputLevel 1

a::HandleKeyInput("a")
b::HandleKeyInput("b")
c::HandleKeyInput("c")
d::HandleKeyInput("d")
e::HandleKeyInput("e")
f::HandleKeyInput("f")
g::HandleKeyInput("g")
h::HandleKeyInput("h")
i::HandleKeyInput("i")
j::HandleKeyInput("j")
k::HandleKeyInput("k")
l::HandleKeyInput("l")
m::HandleKeyInput("m")
n::HandleKeyInput("n")
o::HandleKeyInput("o")
p::HandleKeyInput("p")
q::HandleKeyInput("q")
r::HandleKeyInput("r")
s::HandleKeyInput("s")
t::HandleKeyInput("t")
u::HandleKeyInput("u")
v::HandleKeyInput("v")
w::HandleKeyInput("w")
x::HandleKeyInput("x")
y::HandleKeyInput("y")
z::HandleKeyInput("z")

+a::HandleKeyInput("A")
+b::HandleKeyInput("B")
+c::HandleKeyInput("C")
+d::HandleKeyInput("D")
+e::HandleKeyInput("E")
+f::HandleKeyInput("F")
+g::HandleKeyInput("G")
+h::HandleKeyInput("H")
+i::HandleKeyInput("I")
+j::HandleKeyInput("J")
+k::HandleKeyInput("K")
+l::HandleKeyInput("L")
+m::HandleKeyInput("M")
+n::HandleKeyInput("N")
+o::HandleKeyInput("O")
+p::HandleKeyInput("P")
+q::HandleKeyInput("Q")
+r::HandleKeyInput("R")
+s::HandleKeyInput("S")
+t::HandleKeyInput("T")
+u::HandleKeyInput("U")
+v::HandleKeyInput("V")
+w::HandleKeyInput("W")
+x::HandleKeyInput("X")
+y::HandleKeyInput("Y")
+z::HandleKeyInput("Z")

; ============================================================
; 起動・終了ログ／トレイメニュー設定
; ============================================================
Log("=== AlwaysIME_AHK 起動 === (LogFile: " LogFilePath ")")
OnExit((*) => Log("=== AlwaysIME_AHK 終了 ==="))

A_TrayMenu.Delete()
A_TrayMenu.Add("設定を表示", ShowConfig)
A_TrayMenu.Add()
A_TrayMenu.Add("ログファイルを開く", OpenLogFile)
A_TrayMenu.Add("ログファイルを削除", DeleteLogFile)
A_TrayMenu.Add()
A_TrayMenu.Add("終了", (*) => ExitApp())
A_TrayMenu.Default := "設定を表示"
TraySetIcon(A_AhkPath, 2)

; ============================================================
; トレイメニュー関数
; ============================================================

ShowConfig(*) {
    msg := "=== IME制御しないアプリ (IgnoreApps) ===`n"
    if (IgnoreApps.Length = 0)
        msg .= "  (なし)`n"
    for i, v in IgnoreApps
        msg .= "  " i ". " v "`n"

    msg .= "`n=== IME-OFFを強制するアプリ (ForceOffApps) ===`n"
    if (ForceOffApps.Length = 0)
        msg .= "  (なし)`n"
    for i, v in ForceOffApps
        msg .= "  " i ". " v "`n"

    msg .= "`n=== IME-OFFにするタイトルパターン (TitleOffPatterns) ===`n"
    if (TitleOffPatterns.Length = 0)
        msg .= "  (なし)`n"
    for i, v in TitleOffPatterns
        msg .= "  " i ". " v "`n"

    msg .= "`n=== タイトル変化から除外するタグ (TitleIgnoreTags) ===`n"
    if (TitleIgnoreTags.Length = 0)
        msg .= "  (なし)`n"
    for i, v in TitleIgnoreTags
        msg .= "  " i ". " v "`n"

    msg .= "`n未入力タイムアウト: " (IdleTimeoutMs // 1000) " 秒"

    MsgBox msg, "AlwaysIME_AHK 設定一覧", "OK"
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
