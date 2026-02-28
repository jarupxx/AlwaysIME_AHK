; ============================================================
; AlwaysIME_AHK
; キー入力時にIMEを自動制御する常駐スクリプト
; AutoHotKey v2 対応
; ============================================================

#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; ============================================================
; 設定ファイル・アイコン
; ============================================================
global ConfigFilePath := A_ScriptDir "\AlwaysIME_AHK.ini"

; ============================================================
; アイコン設定
; 次にキーを押したらIMEがどう制御されるかをトレイで予告表示
; ファイルが存在しない場合はAHK組み込みアイコンで代替
; ============================================================

; アイコンファイルパス（スクリプトと同じフォルダに配置）
; icon_ime_on.ico      : 次キー → IME-ONにする（通常動作）
; icon_ime_off.ico     : 次キー → IME-OFFにする
; icon_ignore.ico      : 次キー → IME制御なし（IgnoreApps）
global IconImeOn      := A_ScriptDir "\Resources\icon_ime_on.ico"
global IconImeOff     := A_ScriptDir "\Resources\icon_ime_off.ico"
global IconIgnore     := A_ScriptDir "\Resources\icon_ignore.ico"

; AHK組み込みアイコンのフォールバック番号（.icoが存在しない場合に使用）
global IconFallbackMap := Map(
    "ime_on",     [A_AhkPath, 2],
    "ime_off",    [A_AhkPath, 7],
    "ignore",     [A_AhkPath, 8]
)

; 現在表示中のアイコンモード（無駄な再設定を防ぐ）
global CurrentIconMode := ""

; トレイアイコンを更新する
; mode: "ime_on" / "ime_off" / "ignore"
UpdateTrayIcon(mode) {
    global CurrentIconMode
    if (CurrentIconMode = mode)
        return
    CurrentIconMode := mode

    icoPath := (mode = "ime_on")  ? IconImeOn
             : (mode = "ime_off") ? IconImeOff
             : IconIgnore

    if FileExist(icoPath) {
        TraySetIcon icoPath
    } else {
        fb := IconFallbackMap[mode]
        TraySetIcon fb[1], fb[2]
    }

    tip := (mode = "ime_on")  ? "AlwaysIME_AHK: 次キーでIME-ONにします"
         : (mode = "ime_off") ? "AlwaysIME_AHK: 次キーでIME-OFFにします"
         : "AlwaysIME_AHK: このアプリはIME制御対象外"
    A_IconTip := tip
}

; ============================================================
; アクティブウィンドウの変化を検知してアイコンを先行更新する
; WinEventHook: EVENT_SYSTEM_FOREGROUND と EVENT_OBJECT_NAMECHANGE
; ============================================================

RefreshIconForActiveWindow() {
    hwnd := WinExist("A")
    if (hwnd = 0)
        return
    try {
        processName := StrLower(WinGetProcessName("A"))
        rawTitle    := WinGetTitle("A")
    } catch {
        return
    }
    normTitle := NormalizeTitle(rawTitle)

    for app in IgnoreApps {
        if (processName = app) {
            UpdateTrayIcon("ignore")
            return
        }
    }
    for app in ForceOffApps {
        if (processName = app) {
            UpdateTrayIcon("ime_off")
            return
        }
    }
    for pattern in TitleOffPatterns {
        if RegExMatch(rawTitle, pattern) {
            UpdateTrayIcon("ime_off")
            return
        }
    }

    idleMs := A_TimeIdlePhysical
    if (IMEControlled
        && processName = LastProcessName
        && normTitle   = LastWindowTitle
        && idleMs < IdleTimeoutMs) {
        UpdateTrayIcon("ime_on")
        return
    }

    UpdateTrayIcon("ime_on")
}

; WinEventHook コールバック
WinEventProc(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime) {
    if (idObject != 0)
        return
    RefreshIconForActiveWindow()
}

; WinEventHook を登録する
; EVENT_SYSTEM_FOREGROUND = 0x0003（アプリ切替）
; EVENT_OBJECT_NAMECHANGE  = 0x800C（タイトル変化）
SetupWinEventHook() {
    global WinEventHookHandle1, WinEventHookHandle2, WinEventProcRef
    WinEventProcRef := CallbackCreate(WinEventProc, "F", 7)
    WinEventHookHandle1 := DllCall("SetWinEventHook",
        "UInt", 0x0003, "UInt", 0x0003,
        "Ptr",  0, "Ptr", WinEventProcRef,
        "UInt", 0, "UInt", 0, "UInt", 0, "Ptr")
    WinEventHookHandle2 := DllCall("SetWinEventHook",
        "UInt", 0x800C, "UInt", 0x800C,
        "Ptr",  0, "Ptr", WinEventProcRef,
        "UInt", 0, "UInt", 0, "UInt", 0, "Ptr")
    Log("WinEventHook 登録完了 (h1=" WinEventHookHandle1 " h2=" WinEventHookHandle2 ")")
}

TeardownWinEventHook() {
    global WinEventHookHandle1, WinEventHookHandle2, WinEventProcRef
    if WinEventHookHandle1
        DllCall("UnhookWinEvent", "Ptr", WinEventHookHandle1)
    if WinEventHookHandle2
        DllCall("UnhookWinEvent", "Ptr", WinEventHookHandle2)
    if WinEventProcRef
        CallbackFree(WinEventProcRef)
    Log("WinEventHook 解除")
}

OnExitHandler(*) {
    TeardownWinEventHook()
    Log("=== AlwaysIME_AHK 終了 ===")
}

; ============================================================
; 設定ファイルの読み書き
; ============================================================

; INIファイルから設定を読み込む
; セクション [ListName] の各行を配列として返す
LoadConfig() {
    if !FileExist(ConfigFilePath) {
        Log("設定ファイルが見つかりません。デフォルト値を使用します: " ConfigFilePath)
        return
    }
    Log("設定ファイルを読み込みます: " ConfigFilePath)

    global IgnoreApps, ForceOffApps, TitleOffPatterns, TitleIgnoreTags, IdleTimeoutMs

    IgnoreApps      := ReadIniList("IgnoreApps")
    ForceOffApps    := ReadIniList("ForceOffApps")
    TitleOffPatterns := ReadIniList("TitleOffPatterns")
    TitleIgnoreTags := ReadIniList("TitleIgnoreTags")

    timeoutSec := IniRead(ConfigFilePath, "General", "IdleTimeoutSec", "300")
    IdleTimeoutMs := Integer(timeoutSec) * 1000

    Log("設定読み込み完了 (IdleTimeout=" Round(IdleTimeoutMs/1000) "秒"
      . " IgnoreApps=" IgnoreApps.Length
      . " ForceOffApps=" ForceOffApps.Length
      . " TitleOffPatterns=" TitleOffPatterns.Length
      . " TitleIgnoreTags=" TitleIgnoreTags.Length ")")
}

; INIファイルへ設定を書き出す
SaveConfig() {
    Log("設定ファイルを保存します: " ConfigFilePath)

    ; ファイルを一旦削除して新規作成
    if FileExist(ConfigFilePath)
        FileDelete ConfigFilePath

    IniWrite Round(IdleTimeoutMs / 1000), ConfigFilePath, "General", "IdleTimeoutSec"

    WriteIniList("IgnoreApps",       IgnoreApps)
    WriteIniList("ForceOffApps",     ForceOffApps)
    WriteIniList("TitleOffPatterns", TitleOffPatterns)
    WriteIniList("TitleIgnoreTags",  TitleIgnoreTags)

    Log("設定ファイルを保存しました")
}

; [Section] の各エントリ（Item1, Item2, ...）を配列として読み込む
ReadIniList(section) {
    arr := []
    i := 1
    loop {
        val := IniRead(ConfigFilePath, section, "Item" i, "")
        if (val = "")
            break
        arr.Push(val)
        i++
    }
    return arr
}

; 配列を [Section] の Item1, Item2, ... として書き出す
WriteIniList(section, arr) {
    for i, v in arr
        IniWrite v, ConfigFilePath, section, "Item" i
}


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

; IMEを制御しないアプリ（完全一致）
global IgnoreApps := [
    "AutoHotkey64.exe",
]

; キー入力のたびにIME-OFFを強制するアプリ（完全一致）
global ForceOffApps := [
    ; "WindowsTerminal.exe",
]

; タイトルの一部にマッチしたらIME-OFFにするパターン（正規表現）
; アプリ問わず全体で有効
global TitleOffPatterns := [
    "\.cs",
    "\.js",
    "\.ts",
    "\.py",
    "\.ahk",
]

; タイトル変化の検出から除外するタグパターン（正規表現）
; マッチした部分を取り除いてからタイトルを比較する
global TitleIgnoreTags := [
    "\(更新\)",
    "\s*\*$",
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
    imeEnabled := DllCall("SendMessage", "Ptr", hwndIme, "UInt", WM_IME_CONTROL, "Ptr", IMC_GETOPENSTATUS, "Ptr", 0, "Ptr")
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
    else {
        Log("IME_SetHiragana: ImmGetContext 成功 (hwnd=" hwnd ")", "INFO")
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
            UpdateTrayIcon("ignore")
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
            UpdateTrayIcon("ime_off")
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
            UpdateTrayIcon("ime_off")
            SendInput key
            return
        }
    }

    ; IME制御を再開すべきか判定
    if !ShouldControl(processName, normTitle) {
        UpdateTrayIcon("ime_on")
        SendInput key
        return
    }

    ; IME-ON かつ ひらがなモードに設定
    IME_ON(hwnd)
    IME_SetHiragana(hwnd)
    UpdateTrayIcon("ime_on")

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
LoadConfig()
OnExit(OnExitHandler)

A_TrayMenu.Delete()
A_TrayMenu.Add("設定を表示", ShowConfig)
A_TrayMenu.Add()
A_TrayMenu.Add("ログファイルを開く", OpenLogFile)
A_TrayMenu.Add("ログファイルを削除", DeleteLogFile)
A_TrayMenu.Add()
A_TrayMenu.Add("終了", (*) => ExitApp())
A_TrayMenu.Default := "設定を表示"
UpdateTrayIcon("ime_on")
SetupWinEventHook()

; ============================================================
; トレイメニュー関数
; ============================================================

; ============================================================
; 設定画面
; 左：カテゴリーリスト  右：入力エリア
; ============================================================

; カテゴリー定義（表示名 / 説明文 / 入力タイプ）
; type: "list" = 1行1項目テキストエリア / "seconds" = 数値スピナー
global ConfigCategories := [
    Map(
        "key",   "IgnoreApps",
        "label", "制御しないアプリ",
        "desc",  "IMEを一切操作しないアプリ。`n実行ファイル名を1行1件入力。`n例: AutoHotkey64.exe",
        "type",  "list"
    ),
    Map(
        "key",   "ForceOffApps",
        "label", "IME-OFFアプリ",
        "desc",  "アプリ切替・タイトル変化・タイムアウト時にIME-OFFにするアプリ。`n実行ファイル名を1行1件入力。`n例: WindowsTerminal.exe",
        "type",  "list"
    ),
    Map(
        "key",   "TitleOffPatterns",
        "label", "IME-OFFタイトルパターン",
        "desc",  "ウィンドウタイトルにマッチしたらIME-OFFにする正規表現。`n1行1パターン。全アプリ共通。`n例: \.cs",
        "type",  "list"
    ),
    Map(
        "key",   "TitleIgnoreTags",
        "label", "タイトル除外タグ",
        "desc",  "タイトル変化検出の際に無視する部分の正規表現。`nマッチした部分を除去してから変化を判定。`n例: \(更新\)",
        "type",  "list"
    ),
    Map(
        "key",   "IdleTimeoutMs",
        "label", "未入力タイムアウト",
        "desc",  "この時間キーボード・マウス未入力が続いたらIME制御を再開する。`n秒単位で入力。",
        "type",  "seconds"
    ),
]

; 現在選択中のカテゴリーインデックス
global ConfigSelectedIndex := 1

ShowConfig(*) {
    global ConfigSelectedIndex

    W  := 640
    H  := 400
    LW := 160
    SP := 8
    RX := LW + SP * 2
    RW := W - RX - SP
    LH := H - SP * 2
    DescH  := 56
    ItemLH := H - SP*3 - DescH - 32 - 36
    BtnY   := H - 36

    ; ---- 現在カテゴリーの項目を保持する内部配列 ----
    currentItems := []
    editingIdx   := 0   ; 編集中の行番号（0=新規追加モード）

    ; ---- ウィンドウ生成 ----
    cfgGui := Gui("+Resize -MinimizeBox", "AlwaysIME_AHK 設定")
    cfgGui.SetFont("s9", "Meiryo UI")
    cfgGui.BackColor := "F0F0F0"

    cfgGui.Add("Text", "x0 y0 w" LW " h" H " BackgroundE8E8E8", "")

    labels := []
    for cat in ConfigCategories
        labels.Push(cat["label"])
    catList := cfgGui.Add("ListBox",
        "x" SP " y" SP " w" LW-SP*2 " h" LH " vCategoryList BackgroundE8E8E8 -E0x200",
        labels)
    catList.Choose(ConfigSelectedIndex)

    descLabel := cfgGui.Add("Text",
        "x" RX " y" SP " w" RW " h" DescH " vDescLabel BackgroundTrans", "")

    itemList := cfgGui.Add("ListBox",
        "x" RX " y" SP+DescH " w" RW-80 " h" ItemLH " vItemList")

    BX := RX + RW - 76
    btnAdd    := cfgGui.Add("Button", "x" BX " y" SP+DescH    " w72 h24", "追加")
    btnEdit   := cfgGui.Add("Button", "x" BX " y" SP+DescH+28 " w72 h24", "編集")
    btnDelete := cfgGui.Add("Button", "x" BX " y" SP+DescH+56 " w72 h24", "削除")
    btnUp     := cfgGui.Add("Button", "x" BX " y" SP+DescH+88 " w72 h24", "↑ 上へ")
    btnDown   := cfgGui.Add("Button", "x" BX " y" SP+DescH+116 " w72 h24", "↓ 下へ")

    inputY := SP + DescH + ItemLH + SP
    valLabel  := cfgGui.Add("Text", "x" RX " y" inputY+4 " w28 h20 BackgroundTrans vValLabel", "値：")
    inputEdit := cfgGui.Add("Edit",
        "x" RX+30 " y" inputY " w" RW-80-34 " h24 vInputEdit")
    btnAddOK := cfgGui.Add("Button",
        "x" RX+RW-76 " y" inputY " w72 h24", "リストへ追加")

    spinRow  := cfgGui.Add("Text",
        "x" RX " y" SP+DescH " w60 h24 vSpinLabel Hidden BackgroundTrans", "秒数：")
    spinCtrl := cfgGui.Add("Edit",
        "x" RX+64 " y" SP+DescH " w80 h24 vSpinCtrl Hidden Number")
    spinUpDown := cfgGui.Add("UpDown", "vSpinUpDown Range10-3600 Hidden", 300)

    cfgGui.Add("Text", "x0 y" BtnY-8 " w" W " h2 BackgroundTrans +0x10")
    btnSave   := cfgGui.Add("Button", "x" RX      " y" BtnY " w120 h28 Default", "保存して閉じる")
    btnApply  := cfgGui.Add("Button", "x" RX+124  " y" BtnY " w80  h28",          "適用")
    btnCancel := cfgGui.Add("Button", "x" W-SP-80 " y" BtnY " w80  h28",          "キャンセル")

    catList.OnEvent("Change",        OnCategoryChange)
    btnAdd.OnEvent("Click",          OnStartAdd)
    btnEdit.OnEvent("Click",         OnStartEdit)
    btnAddOK.OnEvent("Click",        OnCommitItem)
    btnDelete.OnEvent("Click",       OnDeleteItem)
    btnUp.OnEvent("Click",           OnMoveUp)
    btnDown.OnEvent("Click",         OnMoveDown)
    itemList.OnEvent("DoubleClick",  OnStartEdit)
    btnSave.OnEvent("Click",         OnSave)
    btnApply.OnEvent("Click",        OnApply)
    btnCancel.OnEvent("Click",       (*) => cfgGui.Destroy())
    cfgGui.OnEvent("Close",          (*) => cfgGui.Destroy())

    RenderCategory(ConfigSelectedIndex)
    cfgGui.Show("w" W " h" H)

    ; ----------------------------------------------------------
    ; itemList を currentItems の内容で再描画し、sel行を選択
    ; ----------------------------------------------------------
    RefreshListBox(sel := 0) {
        itemList.Delete()
        if (currentItems.Length > 0)
            itemList.Add(currentItems)
        if (sel > 0 && sel <= currentItems.Length)
            itemList.Choose(sel)
        else if (currentItems.Length > 0)
            itemList.Choose(1)
    }

    ; ----------------------------------------------------------
    ; カテゴリー切替
    ; ----------------------------------------------------------
    OnCategoryChange(ctrl, *) {
        idx := ctrl.Value
        if (idx = 0)
            return
        FlushItemList(ConfigSelectedIndex)
        editingIdx    := 0
        btnAddOK.Text := "追加"
        inputEdit.Value := ""
        ConfigSelectedIndex := idx
        RenderCategory(idx)
    }

    ; ----------------------------------------------------------
    ; 右ペインを指定カテゴリーで描画
    ; ----------------------------------------------------------
    RenderCategory(idx) {
        cat := ConfigCategories[idx]
        descLabel.Value := cat["desc"]

        if (cat["type"] = "seconds") {
            itemList.Visible   := false
            btnAdd.Visible     := false
            btnEdit.Visible    := false
            btnDelete.Visible  := false
            btnUp.Visible      := false
            btnDown.Visible    := false
            valLabel.Visible   := false
            inputEdit.Visible  := false
            btnAddOK.Visible   := false
            spinRow.Visible    := true
            spinCtrl.Visible   := true
            spinUpDown.Visible := true
            spinCtrl.Value     := Round(IdleTimeoutMs / 1000)
        } else {
            itemList.Visible   := true
            btnAdd.Visible     := true
            btnEdit.Visible    := true
            btnDelete.Visible  := true
            btnUp.Visible      := true
            btnDown.Visible    := true
            valLabel.Visible   := true
            inputEdit.Visible  := true
            btnAddOK.Visible   := true
            spinRow.Visible    := false
            spinCtrl.Visible   := false
            spinUpDown.Visible := false

            currentItems := []
            for v in GetGlobalArray(cat["key"])
                currentItems.Push(v)
            inputEdit.Value := ""
            RefreshListBox(1)
        }
    }

    ; ----------------------------------------------------------
    ; 新規追加モードに入る
    ; ----------------------------------------------------------
    OnStartAdd(*) {
        editingIdx := 0
        inputEdit.Value := ""
        btnAddOK.Text   := "追加"
        inputEdit.Focus()
    }

    ; ----------------------------------------------------------
    ; 編集モードに入る（ボタンまたはダブルクリック）
    ; ----------------------------------------------------------
    OnStartEdit(*) {
        idx := itemList.Value
        if (idx = 0)
            return
        editingIdx      := idx
        inputEdit.Value := currentItems[idx]
        btnAddOK.Text   := "更新"
        inputEdit.Focus()
    }

    ; ----------------------------------------------------------
    ; 追加 or 更新を確定する（btnAddOK）
    ; ----------------------------------------------------------
    OnCommitItem(*) {
        val := Trim(inputEdit.Value)
        if (val = "")
            return
        if (editingIdx = 0) {
            ; 新規追加
            currentItems.Push(val)
            RefreshListBox(currentItems.Length)
        } else {
            ; 既存項目を更新
            currentItems[editingIdx] := val
            RefreshListBox(editingIdx)
            editingIdx := 0
        }
        inputEdit.Value := ""
        btnAddOK.Text   := "追加"
        inputEdit.Focus()
    }

    ; ----------------------------------------------------------
    ; 選択中の項目を currentItems から削除して再描画
    ; ----------------------------------------------------------
    OnDeleteItem(*) {
        idx := itemList.Value
        if (idx = 0)
            return
        currentItems.RemoveAt(idx)
        RefreshListBox(Min(idx, currentItems.Length))
    }

    ; ----------------------------------------------------------
    ; 選択中の項目を1つ上へ
    ; ----------------------------------------------------------
    OnMoveUp(*) {
        idx := itemList.Value
        if (idx <= 1)
            return
        tmp := currentItems[idx - 1]
        currentItems[idx - 1] := currentItems[idx]
        currentItems[idx]     := tmp
        RefreshListBox(idx - 1)
    }

    ; ----------------------------------------------------------
    ; 選択中の項目を1つ下へ
    ; ----------------------------------------------------------
    OnMoveDown(*) {
        idx := itemList.Value
        if (idx = 0 || idx >= currentItems.Length)
            return
        tmp := currentItems[idx + 1]
        currentItems[idx + 1] := currentItems[idx]
        currentItems[idx]     := tmp
        RefreshListBox(idx + 1)
    }

    ; ----------------------------------------------------------
    ; currentItems をグローバル変数に書き戻す
    ; ----------------------------------------------------------
    FlushItemList(idx) {
        cat := ConfigCategories[idx]
        if (cat["type"] = "seconds") {
            global IdleTimeoutMs := Integer(spinCtrl.Value) * 1000
        } else {
            arr := []
            for v in currentItems
                arr.Push(v)
            SetGlobalArray(cat["key"], arr)
        }
    }

    OnSave(*) {
        FlushItemList(ConfigSelectedIndex)
        SaveConfig()
        Log("設定を変更しました")
        cfgGui.Destroy()
    }

    OnApply(*) {
        FlushItemList(ConfigSelectedIndex)
        SaveConfig()
        Log("設定を適用しました")
    }
}

; グローバル配列をキー名で取得
GetGlobalArray(key) {
    if (key = "IgnoreApps")
        return IgnoreApps
    if (key = "ForceOffApps")
        return ForceOffApps
    if (key = "TitleOffPatterns")
        return TitleOffPatterns
    if (key = "TitleIgnoreTags")
        return TitleIgnoreTags
    return []
}

; グローバル配列をキー名でセット
SetGlobalArray(key, arr) {
    if (key = "IgnoreApps") {
        global IgnoreApps := arr
    } else if (key = "ForceOffApps") {
        global ForceOffApps := arr
    } else if (key = "TitleOffPatterns") {
        global TitleOffPatterns := arr
    } else if (key = "TitleIgnoreTags") {
        global TitleIgnoreTags := arr
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
