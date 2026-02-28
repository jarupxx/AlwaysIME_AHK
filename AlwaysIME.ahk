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

; ============================================================
; IME制御関数
; ============================================================

; IMEのON/OFF状態を取得（1=ON, 0=OFF）
IME_GET(WinTitle := "A") {
    hwnd := WinExist(WinTitle)
    if (hwnd = 0)
        return 0
    hwndIme := DllCall("imm32\ImmGetDefaultIMEWnd", "Ptr", hwnd, "Ptr")
    return DllCall("SendMessage", "Ptr", hwndIme, "UInt", WM_IME_CONTROL, "Ptr", IMC_GETOPENSTATUS, "Ptr", 0, "Ptr")
}

; IMEをONにする
IME_ON(WinTitle := "A") {
    hwnd := WinExist(WinTitle)
    if (hwnd = 0)
        return
    hwndIme := DllCall("imm32\ImmGetDefaultIMEWnd", "Ptr", hwnd, "Ptr")
    DllCall("SendMessage", "Ptr", hwndIme, "UInt", WM_IME_CONTROL, "Ptr", IMC_SETOPENSTATUS, "Ptr", 1, "Ptr")
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

    ; IMEがOFFなら先にONにする
    if (IME_GET("A") = 0) {
        IME_ON("A")
        Sleep 30
    }

    ; キーを送信
    SendInput key
    return true
}

; ============================================================
; アルファベットキーのフック（a-z）
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
; トレイメニュー設定
; ============================================================
A_TrayMenu.Delete()
A_TrayMenu.Add("対象アプリ一覧を表示", ShowTargetApps)
A_TrayMenu.Add("対象アプリを追加", AddTargetApp)
A_TrayMenu.Add()
A_TrayMenu.Add("終了", (*) => ExitApp())
A_TrayMenu.Default := "対象アプリ一覧を表示"
TraySetIcon(A_AhkPath, 2)  ; AHKアイコンを使用

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
        ; 重複チェック
        for app in TargetApps {
            if (app = newApp) {
                MsgBox newApp " はすでに登録されています。", "情報"
                return
            }
        }
        TargetApps.Push(newApp)
        MsgBox newApp " を追加しました。", "完了"
    }
}
