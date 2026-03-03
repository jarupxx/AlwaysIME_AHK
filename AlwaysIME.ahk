; ============================================================
; AlwaysIME_AHK.ahk
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

; ============================================================
; アクティブウィンドウに対してIME制御 + アイコン更新を行う
; WinEventHook から SetTimer 経由で遅延呼び出しされる
; ============================================================

RefreshActiveWindow() {
    ; SetTimer の1回起動なのでまず自分自身を解除
    SetTimer RefreshActiveWindow, 0

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

    ; タイトル空欄スキップ（上級者向け設定）
    if (SkipEmptyTitle && rawTitle = "") {
        UpdateTrayIcon("ignore")
        return
    }

    ; IgnoreApps: 制御しない
    for app in IgnoreApps {
        if (processName = app) {
            UpdateTrayIcon("ignore")
            return
        }
    }

    ; ForceOffApps: 再開条件が揃ったときだけIME-OFF
    for app in ForceOffApps {
        if (processName = app) {
            UpdateTrayIcon("ime_off")
            if ShouldControl(processName, normTitle) {
                IME_OFF(hwnd)
                global IMEControlled  := true
                global LastProcessName := processName
                global LastWindowTitle := normTitle
            }
            return
        }
    }

    ; TitleOffPatterns: タイトルにマッチしたらIME-OFF
    for pattern in TitleOffPatterns {
        if RegExMatch(rawTitle, pattern) {
            UpdateTrayIcon("ime_off")
            if ShouldControl(processName, normTitle) {
                IME_OFF(hwnd)
                global IMEControlled  := true
                global LastProcessName := processName
                global LastWindowTitle := normTitle
            }
            return
        }
    }

    ; 制御済み判定（アプリ・タイトル変化なし＆タイムアウト未到達）
    idleMs := A_TimeIdlePhysical
    if (IMEControlled
        && processName = LastProcessName
        && normTitle   = LastWindowTitle
        && idleMs < IdleTimeoutMs) {
        UpdateTrayIcon("ime_on")
        return
    }

    ; 上記以外 → IME-ON + ひらがなモードに先行設定
    UpdateTrayIcon("ime_on")
    IME_ON(hwnd)
    IME_SetHiragana(hwnd)
    global IMEControlled  := true
    global LastProcessName := processName
    global LastWindowTitle := normTitle
    Log("WinEvent先行制御: IME-ON app=" processName " title=`"" normTitle "`"", "DEBUG")
}

; WinEventHook コールバック
; DLLコールバックスレッドから呼ばれるため SetTimer で
; AHKメインスレッドに処理を委譲する（遅延50ms）
WinEventProc(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime) {
    if (idObject != 0)
        return
    SetTimer RefreshActiveWindow, -50
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
    if MsImeSettingsEnabled
        RestoreMsImeRegistry()
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

    skipVal := IniRead(ConfigFilePath, "Advanced", "SkipEmptyTitle", "1")
    SkipEmptyTitle := (skipVal = "1")

    logVal := IniRead(ConfigFilePath, "Advanced", "EnableLog", "0")
    EnableLog := (logVal = "1")

    confirmVal := IniRead(ConfigFilePath, "Advanced", "ConfirmExit", "1")
    ConfirmExit := (confirmVal = "1")

    global MsImeSettingsEnabled, SpaceInitVal, SpaceTargetVal, PunctInitVal, PunctTargetVal
    msimeVal := IniRead(ConfigFilePath, "MsIme", "Enabled", "0")
    MsImeSettingsEnabled := false   ; 読み込み後に EnableMsImeSettings() で有効化する
    SpaceInitVal   := Integer(IniRead(ConfigFilePath, "MsIme", "SpaceInitVal",   "0"))
    SpaceTargetVal := Integer(IniRead(ConfigFilePath, "MsIme", "SpaceTargetVal", "2"))
    PunctInitVal   := Integer(IniRead(ConfigFilePath, "MsIme", "PunctInitVal",   "1"))
    PunctTargetVal := Integer(IniRead(ConfigFilePath, "MsIme", "PunctTargetVal", "0"))

    Log("設定読み込み完了 (IdleTimeout=" Round(IdleTimeoutMs/1000) "秒"
      . " IgnoreApps=" IgnoreApps.Length
      . " ForceOffApps=" ForceOffApps.Length
      . " TitleOffPatterns=" TitleOffPatterns.Length
      . " TitleIgnoreTags=" TitleIgnoreTags.Length
      . " MsIme=" msimeVal ")")

    ; MS-IME設定は他の初期化が終わった後に有効化する必要があるため
    ; トレイメニュー構築後に呼ばれるよう SetTimer で遅延実行する
    if (msimeVal = "1")
        SetTimer(() => EnableMsImeSettings(), -1)
}

; INIファイルへ設定を書き出す
SaveConfig() {
    Log("設定ファイルを保存します: " ConfigFilePath)

    ; ファイルを一旦削除して新規作成
    if FileExist(ConfigFilePath)
        FileDelete ConfigFilePath

    IniWrite Round(IdleTimeoutMs / 1000), ConfigFilePath, "General", "IdleTimeoutSec"
    IniWrite (SkipEmptyTitle ? "1" : "0"), ConfigFilePath, "Advanced", "SkipEmptyTitle"
    IniWrite (EnableLog     ? "1" : "0"), ConfigFilePath, "Advanced", "EnableLog"
    IniWrite (ConfirmExit   ? "1" : "0"), ConfigFilePath, "Advanced", "ConfirmExit"

    IniWrite (MsImeSettingsEnabled ? "1" : "0"), ConfigFilePath, "MsIme", "Enabled"
    IniWrite SpaceInitVal,   ConfigFilePath, "MsIme", "SpaceInitVal"
    IniWrite SpaceTargetVal, ConfigFilePath, "MsIme", "SpaceTargetVal"
    IniWrite PunctInitVal,   ConfigFilePath, "MsIme", "PunctInitVal"
    IniWrite PunctTargetVal, ConfigFilePath, "MsIme", "PunctTargetVal"

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
    if !EnableLog
        return
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
; MS-IME 入力設定（レジストリ制御）
; ============================================================

; レジストリパス
global MSIME_REG_PATH := "HKCU\Software\Microsoft\IME\15.0\IMEJP\MSIME"

; --- スペース制御 (InputSpace) ---
; 0: 現在の入力モード相当（ここでは全角扱い）
; 1: 常に全角
; 2: 常に半角
global IME_AUTO_WIDTH_SPACE := 0
global IME_FULL_WIDTH_SPACE := 1
global IME_HALF_WIDTH_SPACE := 2
; インデックス 1-based: [1]=値0, [2]=値1, [3]=値2
global InputSpaceLabels := ["現在の入力モード", "常に全角", "常に半角"]

; MS-IME入力設定機能の有効/無効（初期値：オフ）
global MsImeSettingsEnabled := false

; 起動時に読み込んだ元の値（終了時復元用）
global OrigInputSpace   := -1   ; -1 = 未取得
global OrigOption1      := -1

; 現在スクリプトが書き込んでいる値（-1 = 未変更）
global CurInputSpace    := -1
global CurOption1       := -1

; ユーザー設定：スペースの初期値と切替先（レジストリ値 0-2）
; 起動時にレジストリから読んだ値を初期値として使う（-1 = レジストリ実値を使用）
global SpaceInitVal   := -1   ; -1 = レジストリ実値をそのまま使う
global SpaceTargetVal := IME_FULL_WIDTH_SPACE    ; 切替先デフォルト: 常に全角

; --- 句読点制御 (option1 ビットマスク) ---
; ビット位置: (value >> 16) & 0x3
; 0: ，．  1: 、。  2: 、．  3: ，。
global IME_COMMA_PERIOD  := 0
global IME_TOUTEN_KUTENN := 1
global IME_TOUTEN_PERIOD := 2
global IME_COMMA_KUTENN  := 3
global PunctuationLabels := ["，．", "、。", "、．", "，。"]

; ビットマスクのシフト量とマスク
global OPTION1_PUNCT_SHIFT := 16
global OPTION1_PUNCT_MASK  := 0x3

; ユーザー設定：句読点の初期値と切替先（0-3）
global PunctInitVal   := -1   ; -1 = レジストリ実値をそのまま使う
global PunctTargetVal := IME_COMMA_KUTENN    ; 切替先デフォルト: ，。

; ============================================================
; MS-IME レジストリ読み書きユーティリティ
; ============================================================

; InputSpace レジストリ値を読む（失敗時は -1）
ReadInputSpace() {
    try {
        val := RegRead(MSIME_REG_PATH, "InputSpace")
        return Integer(val)
    } catch {
        return -1
    }
}

; InputSpace レジストリ値を書く
WriteInputSpace(val) {
    try {
        RegWrite val, "REG_DWORD", MSIME_REG_PATH, "InputSpace"
        Log("WriteInputSpace: " val)
        return true
    } catch as e {
        Log("WriteInputSpace 失敗: " e.Message, "WARN")
        return false
    }
}

; option1 レジストリ値を読む（失敗時は -1）
ReadOption1() {
    try {
        val := RegRead(MSIME_REG_PATH, "option1")
        return Integer(val)
    } catch {
        return -1
    }
}

; option1 の句読点ビット(bits 17-16)を val(0-3) に設定して書く
WriteOption1Punct(punctMode) {
    try {
        current := ReadOption1()
        if (current = -1) {
            Log("WriteOption1Punct: option1 読み取り失敗", "WARN")
            return false
        }
        ; 既存のビット17-16をクリアして新値をセット
        newVal := (current & ~(OPTION1_PUNCT_MASK << OPTION1_PUNCT_SHIFT))
               | ((punctMode & OPTION1_PUNCT_MASK) << OPTION1_PUNCT_SHIFT)
        RegWrite newVal, "REG_DWORD", MSIME_REG_PATH, "option1"
        Log("WriteOption1Punct: punctMode=" punctMode " option1: " current " → " newVal)
        return true
    } catch as e {
        Log("WriteOption1Punct 失敗: " e.Message, "WARN")
        return false
    }
}

; option1 から現在の句読点モード(0-3)を読む
ReadOption1Punct() {
    val := ReadOption1()
    if (val = -1)
        return -1
    return (val >> OPTION1_PUNCT_SHIFT) & OPTION1_PUNCT_MASK
}

; ============================================================
; MS-IME入力設定を有効化（レジストリ保存 + 初期切替）
; ============================================================
EnableMsImeSettings() {
    global MsImeSettingsEnabled, OrigInputSpace, OrigOption1, CurInputSpace, CurOption1
    global SpaceInitVal, SpaceTargetVal, PunctInitVal, PunctTargetVal

    ; 元の値を保存（初回のみ）
    if (OrigInputSpace = -1) {
        OrigInputSpace := ReadInputSpace()
        Log("MS-IME設定有効化: OrigInputSpace=" OrigInputSpace)
    }
    if (OrigOption1 = -1) {
        OrigOption1 := ReadOption1()
        Log("MS-IME設定有効化: OrigOption1=" OrigOption1)
    }

    MsImeSettingsEnabled := true

    ; スペース：設定された初期値をレジストリに書き込む
    ; SpaceInitVal = -1 → レジストリ実値のまま（変更しない）
    ; SpaceInitVal >= 0 → 指定値を書き込む
    if (SpaceInitVal >= 0) {
        CurInputSpace := SpaceInitVal
        WriteInputSpace(CurInputSpace)
    } else if (CurInputSpace = -1) {
        ; 初期値未指定かつ初回: レジストリ実値をCurに記録するだけ
        CurInputSpace := (OrigInputSpace >= 0) ? OrigInputSpace : IME_FULL_WIDTH_SPACE
    }

    ; 句読点：設定された初期値をレジストリに書き込む
    if (PunctInitVal >= 0) {
        CurOption1 := PunctInitVal
        WriteOption1Punct(CurOption1)
    } else if (CurOption1 = -1) {
        initPunct := ReadOption1Punct()
        CurOption1 := (initPunct >= 0) ? initPunct : IME_TOUTEN_KUTENN
    }

    RebuildMsImeMenu()
    Log("MS-IME入力設定 有効化完了 space=" CurInputSpace "→" SpaceTargetVal " punct=" CurOption1 "→" PunctTargetVal)
}

; ============================================================
; MS-IME入力設定を無効化（レジストリ復元）
; ============================================================
DisableMsImeSettings() {
    global MsImeSettingsEnabled
    MsImeSettingsEnabled := false
    RestoreMsImeRegistry()
    RebuildMsImeMenu()
    Log("MS-IME入力設定 無効化")
}

; レジストリを元の値に復元する（終了時・無効化時）
RestoreMsImeRegistry() {
    global OrigInputSpace, OrigOption1, CurInputSpace, CurOption1

    if (OrigInputSpace >= 0) {
        WriteInputSpace(OrigInputSpace)
        Log("レジストリ復元: InputSpace=" OrigInputSpace)
    }
    if (OrigOption1 >= 0) {
        try {
            RegWrite OrigOption1, "REG_DWORD", MSIME_REG_PATH, "option1"
            Log("レジストリ復元: option1=" OrigOption1)
        } catch as e {
            Log("option1 復元失敗: " e.Message, "WARN")
        }
    }

    CurInputSpace := -1
    CurOption1    := -1
}

; ============================================================
; スペース切替（メニュー操作ごと：現在値 ↔ 切替先 を交互）
; ============================================================
ToggleInputSpace(*) {
    global CurInputSpace, MsImeSettingsEnabled, SpaceInitVal, SpaceTargetVal
    if !MsImeSettingsEnabled
        return
    ; 現在が切替先なら初期値へ、そうでなければ切替先へ
    initVal := (SpaceInitVal >= 0) ? SpaceInitVal : CurInputSpace
    CurInputSpace := (CurInputSpace = SpaceTargetVal) ? initVal : SpaceTargetVal
    WriteInputSpace(CurInputSpace)
    UpdateMsImeMenu()
    Log("スペース切替 → " CurInputSpace " (" InputSpaceLabels[CurInputSpace + 1] ")")
}

; ============================================================
; 句読点切替（メニュー操作ごと：現在値 ↔ 切替先 を交互）
; ============================================================
TogglePunctuation(*) {
    global CurOption1, MsImeSettingsEnabled, PunctInitVal, PunctTargetVal
    if !MsImeSettingsEnabled
        return
    ; 現在が切替先なら初期値へ、そうでなければ切替先へ
    initVal := (PunctInitVal >= 0) ? PunctInitVal : CurOption1
    CurOption1 := (CurOption1 = PunctTargetVal) ? initVal : PunctTargetVal
    WriteOption1Punct(CurOption1)
    UpdateMsImeMenu()
    Log("句読点切替 → " CurOption1 " (" PunctuationLabels[CurOption1 + 1] ")")
}

; ============================================================
; トレイメニューに切替項目を動的に追加・削除する
;
; 無効時の構成（固定）:
;   1: 設定を表示
;   2: セパレータ
;   3: ログファイルを開く
;   4: ログファイルを削除
;   5: セパレータ
;   6: 終了
;
; 有効時の構成（切替項目を終了の直前に挿入）:
;   1: 設定を表示
;   2: セパレータ
;   3: ログファイルを開く
;   4: ログファイルを削除
;   5: セパレータ
;   6: 句読点切替
;   7: スペース切替
;   8: セパレータ
;   9: 終了
; ============================================================
RebuildMsImeMenu() {
    global MsImeSettingsEnabled, TrayMenuHasMsImeItems

    if MsImeSettingsEnabled {
        if !TrayMenuHasMsImeItems {
            ; 「終了」の前（位置6）にセパレータ・切替項目を挿入
            ; Insert は指定位置の「前」に追加される
            A_TrayMenu.Insert("6&")                              ; セパレータ
            A_TrayMenu.Insert("6&", "句読点切替", TogglePunctuation)
            A_TrayMenu.Insert("6&", "スペース切替", ToggleInputSpace)
            TrayMenuHasMsImeItems := true
        }
        UpdateMsImeMenu()
    } else {
        if TrayMenuHasMsImeItems {
            ; 挿入した3項目（スペース切替・句読点切替・セパレータ）を削除
            ; 有効時の構成: 6=スペース切替 7=句読点切替 8=セパレータ
            A_TrayMenu.Delete("8&")
            A_TrayMenu.Delete("7&")
            A_TrayMenu.Delete("6&")
            TrayMenuHasMsImeItems := false
        }
    }
}

; ============================================================
; 切替項目のラベルを現在値で更新する（有効時のみ）
; ============================================================
UpdateMsImeMenu() {
    global MsImeSettingsEnabled, CurInputSpace, CurOption1
    global SpaceTargetVal, PunctTargetVal

    if !MsImeSettingsEnabled
        return

    ; スペース切替（位置6）
    curSpaceLabel  := (CurInputSpace >= 0 && CurInputSpace <= 2)
        ? InputSpaceLabels[CurInputSpace + 1] : "─"
    nextSpaceLabel := (SpaceTargetVal >= 0 && SpaceTargetVal <= 2)
        ? InputSpaceLabels[SpaceTargetVal + 1] : "─"
    spaceArrow := (CurInputSpace = SpaceTargetVal) ? "← 戻す" : "→ " nextSpaceLabel
    A_TrayMenu.Rename("6&", "スペース [" curSpaceLabel "]  " spaceArrow)

    ; 句読点切替（位置7）
    curPunctLabel  := (CurOption1 >= 0 && CurOption1 <= 3)
        ? PunctuationLabels[CurOption1 + 1] : "─"
    nextPunctLabel := (PunctTargetVal >= 0 && PunctTargetVal <= 3)
        ? PunctuationLabels[PunctTargetVal + 1] : "─"
    punctArrow := (CurOption1 = PunctTargetVal) ? "← 戻す" : "→ " nextPunctLabel
    A_TrayMenu.Rename("7&", "句読点 [" curPunctLabel "]  " punctArrow)
}

; MS-IME入力設定の有効/無効をトグル
ToggleMsImeSettings(*) {
    global MsImeSettingsEnabled
    if MsImeSettingsEnabled
        DisableMsImeSettings()
    else
        EnableMsImeSettings()
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
    ; "\(更新\)",
    "\s*\*$",
]

; ============================================================
; 上級者向け設定
; ============================================================

; タイトルが空欄のウィンドウはIMEを制御しない
global SkipEmptyTitle := true

; .logファイルを生成する（初期値：しない）
; オンにすると AlwaysIME_AHK.log へ動作ログを記録する
global EnableLog := false

; トレイメニューから終了するとき確認メッセージを表示する（初期値：オン）
global ConfirmExit := true

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
; フォーカスを持つ hwnd を返す
; タブモード等でフレームhwndとフォーカスhwndが異なる場合に対応
; GetGUIThreadInfo でフォーカスウィンドウを取得し、
; 失敗した場合はフォールバックとして渡された hwnd をそのまま返す
; ============================================================
GetFocusedHwnd(hwnd) {
    ; GUITHREADINFO 構造体（cbSize=72 固定）
    ; typedef struct tagGUITHREADINFO {
    ;   DWORD cbSize, flags;
    ;   HWND hwndActive, hwndFocus, hwndCapture, hwndMenuOwner, hwndMoveSize, hwndCaret;
    ;   RECT rcCaret;
    ; }
    ; hwndFocus のオフセット = 4 + 4 + 8 = 16 バイト目（64bit: Ptr=8bytes）
    cbSize := 72
    buf := Buffer(cbSize, 0)
    NumPut("UInt", cbSize, buf, 0)

    threadId := DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "Ptr", 0, "UInt")
    ret := DllCall("GetGUIThreadInfo", "UInt", threadId, "Ptr", buf.Ptr, "Int")
    if !ret {
        Log("GetFocusedHwnd: GetGUIThreadInfo 失敗 (hwnd=" hwnd ")", "WARN")
        return hwnd
    }

    ; hwndFocus はオフセット 16（cbSize:4 + flags:4 + hwndActive:8）
    focusedHwnd := NumGet(buf, 16, "Ptr")
    if (focusedHwnd = 0)
        return hwnd

    Log("GetFocusedHwnd: frame=" hwnd " focused=" focusedHwnd, "DEBUG")
    return focusedHwnd
}

; ============================================================
; IMEをONにする
; ============================================================
IME_ON(hwnd) {
    hwnd := GetFocusedHwnd(hwnd)
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
    hwnd := GetFocusedHwnd(hwnd)
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
    hwnd := GetFocusedHwnd(hwnd)
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

    Log("キー入力: `"" key "`" app=" processName " title=`"" normTitle "`"", "DEBUG")

    ; タイトル空欄スキップ（上級者向け設定）
    if (SkipEmptyTitle && rawTitle = "") {
        Log("スキップ: タイトル空欄 (" processName ")", "DEBUG")
        UpdateTrayIcon("ignore")
        SendInput key
        return
    }

    ; IgnoreApps: 制御しない
    for app in IgnoreApps {
        if (processName = app) {
            Log("スキップ: IgnoreApps に一致 (" processName ")", "DEBUG")
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

; MS-IME入力設定サブメニューは廃止。切替項目はメインメニューに直接追加する。
global TrayMenuHasMsImeItems := false

A_TrayMenu.Delete()
A_TrayMenu.Add("設定を表示", ShowConfig)
A_TrayMenu.Add()
A_TrayMenu.Add("ログファイルを開く", OpenLogFile)
A_TrayMenu.Add("ログファイルを削除", DeleteLogFile)
A_TrayMenu.Add()
A_TrayMenu.Add("終了", OnMenuExit)
A_TrayMenu.Default := "設定を表示"
UpdateTrayIcon("ime_on")
SetupWinEventHook()

; ============================================================
; トレイメニュー関数
; ============================================================

OnMenuExit(*) {
    if ConfirmExit {
        result := MsgBox("AlwaysIME_AHK を終了しますか？", "確認", "YesNo Icon?")
        if (result != "Yes")
            return
    }
    ExitApp()
}

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
    Map(
        "key",   "Advanced",
        "label", "上級者向け設定",
        "desc",  "動作に詳しい方向けの設定です。`n通常は変更不要です。",
        "type",  "advanced"
    ),
    Map(
        "key",   "MsIme",
        "label", "MS-IME入力設定",
        "desc",  "MS-IMEのスペース・句読点をレジストリ経由で制御します。`n有効にするとトレイメニューから切替できます。`n終了時はレジストリを元の値に復元します。",
        "type",  "msime"
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

    ; ---- 右ペイン：上級者向けチェックボックス ----
    chkSkipEmptyTitle := cfgGui.Add("Checkbox",
        "x" RX " y" SP+DescH    " w" RW " h24 vChkSkipEmpty Hidden",
        "タイトルが空欄のウィンドウはIMEを制御しない（推奨）")
    chkDebugLog := cfgGui.Add("Checkbox",
        "x" RX " y" SP+DescH+30 " w" RW " h24 vChkDebugLog Hidden",
        ".logファイルを生成する（動作ログをファイルに記録する）")
    chkConfirmExit := cfgGui.Add("Checkbox",
        "x" RX " y" SP+DescH+60 " w" RW " h24 vChkConfirmExit Hidden",
        "トレイメニューから終了するとき確認メッセージを表示する")

    ; ---- MS-IMEパネル（msimeタイプ専用） ----
    msimePanelChk := cfgGui.Add("Checkbox",
        "x" RX " y" SP+DescH " w" RW " h24 vMsImePanelChk Hidden",
        "MS-IME入力設定を有効にする")

    ; 左カラム：スペース設定  右カラム：句読点設定
    HalfW := (RW - 16) // 2
    ColL  := RX
    ColR  := RX + HalfW + 16
    LblW  := 52                          ; ラベル幅（「初期値：」「切替先：」）
    DDW   := Round((HalfW - LblW - 4) * 0.8)   ; ドロップダウン幅（元より2割減）
    DDX   := LblW + 4                   ; ラベルからDDまでのオフセット

    cfgGui.Add("Text", "x" ColL " y" SP+DescH+34 " w" HalfW " h18 vMsImeLblSpace Hidden BackgroundTrans",
        "── スペース ──────────")
    cfgGui.Add("Text", "x" ColL " y" SP+DescH+56 " w" LblW " h20 vMsImeLblSpaceInit Hidden BackgroundTrans",
        "初期値：")
    msimeSpaceInit := cfgGui.Add("DropDownList",
        "x" ColL+DDX " y" SP+DescH+53 " w" DDW " h120 vMsImeSpaceInit Hidden",
        InputSpaceLabels)
    cfgGui.Add("Text", "x" ColL " y" SP+DescH+84 " w" LblW " h20 vMsImeLblSpaceTo Hidden BackgroundTrans",
        "切替先：")
    msimeSpaceTo := cfgGui.Add("DropDownList",
        "x" ColL+DDX " y" SP+DescH+81 " w" DDW " h120 vMsImeSpaceTo Hidden",
        InputSpaceLabels)

    cfgGui.Add("Text", "x" ColR " y" SP+DescH+34 " w" HalfW " h18 vMsImeLblPunct Hidden BackgroundTrans",
        "── 句読点 ──────────")
    cfgGui.Add("Text", "x" ColR " y" SP+DescH+56 " w" LblW " h20 vMsImeLblPunctInit Hidden BackgroundTrans",
        "初期値：")
    msimePunctInit := cfgGui.Add("DropDownList",
        "x" ColR+DDX " y" SP+DescH+53 " w" DDW " h120 vMsImePunctInit Hidden",
        PunctuationLabels)
    cfgGui.Add("Text", "x" ColR " y" SP+DescH+84 " w" LblW " h20 vMsImeLblPunctTo Hidden BackgroundTrans",
        "切替先：")
    msimePunctTo := cfgGui.Add("DropDownList",
        "x" ColR+DDX " y" SP+DescH+81 " w" DDW " h120 vMsImePunctTo Hidden",
        PunctuationLabels)

    cfgGui.Add("Text",
        "x" RX " y" SP+DescH+116 " w" RW " h48 vMsImeHint Hidden BackgroundTrans",
        "初期値：有効化時にMS-IMEへ書き込む設定値。「現在の入力モード」は変更なし。`n切替先：トレイ右クリック → MS-IME入力設定 から切替ボタンを押したときの値。`nもう一度押すと初期値に戻ります。終了時はレジストリを元の値に復元します。")

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

        ; 全コントロールを一旦隠す
        HideAll() {
            itemList.Visible          := false
            btnAdd.Visible            := false
            btnEdit.Visible           := false
            btnDelete.Visible         := false
            btnUp.Visible             := false
            btnDown.Visible           := false
            valLabel.Visible          := false
            inputEdit.Visible         := false
            btnAddOK.Visible          := false
            spinRow.Visible           := false
            spinCtrl.Visible          := false
            spinUpDown.Visible        := false
            chkSkipEmptyTitle.Visible := false
            chkDebugLog.Visible       := false
            chkConfirmExit.Visible    := false
            ; msimeパネル
            for ctrl in [msimePanelChk,
                         msimeSpaceInit, msimeSpaceTo,
                         msimePunctInit, msimePunctTo]
                ctrl.Visible := false
            for v in ["MsImeLblSpace","MsImeLblSpaceInit","MsImeLblSpaceTo",
                      "MsImeLblPunct","MsImeLblPunctInit","MsImeLblPunctTo",
                      "MsImeHint"]
                cfgGui[v].Visible := false
        }

        if (cat["type"] = "seconds") {
            HideAll()
            spinRow.Visible    := true
            spinCtrl.Visible   := true
            spinUpDown.Visible := true
            spinCtrl.Value     := Round(IdleTimeoutMs / 1000)
        } else if (cat["type"] = "advanced") {
            HideAll()
            chkSkipEmptyTitle.Visible := true
            chkSkipEmptyTitle.Value   := SkipEmptyTitle ? 1 : 0
            chkDebugLog.Visible       := true
            chkDebugLog.Value         := EnableLog ? 1 : 0
            chkConfirmExit.Visible    := true
            chkConfirmExit.Value      := ConfirmExit ? 1 : 0
        } else if (cat["type"] = "msime") {
            HideAll()
            msimePanelChk.Visible := true
            msimePanelChk.Value   := MsImeSettingsEnabled ? 1 : 0
            for ctrl in [msimeSpaceInit, msimeSpaceTo, msimePunctInit, msimePunctTo]
                ctrl.Visible := true
            for v in ["MsImeLblSpace","MsImeLblSpaceInit","MsImeLblSpaceTo",
                      "MsImeLblPunct","MsImeLblPunctInit","MsImeLblPunctTo",
                      "MsImeHint"]
                cfgGui[v].Visible := true
            ; DropDownListに現在の設定値を反映（1-based。レジストリ値+1がインデックス）
            ; SpaceInitVal が -1（未設定）のときは 0（現在の入力モード）扱いで index=1
            siVal := (SpaceInitVal < 0) ? 0 : SpaceInitVal
            piVal := (PunctInitVal < 0) ? 0 : PunctInitVal
            msimeSpaceInit.Choose(siVal + 1)
            msimeSpaceTo.Choose(SpaceTargetVal + 1)
            msimePunctInit.Choose(piVal + 1)
            msimePunctTo.Choose(PunctTargetVal + 1)
        } else {
            HideAll()
            itemList.Visible   := true
            btnAdd.Visible     := true
            btnEdit.Visible    := true
            btnDelete.Visible  := true
            btnUp.Visible      := true
            btnDown.Visible    := true
            valLabel.Visible   := true
            inputEdit.Visible  := true
            btnAddOK.Visible   := true

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
        } else if (cat["type"] = "advanced") {
            global SkipEmptyTitle := (chkSkipEmptyTitle.Value = 1)
            global EnableLog      := (chkDebugLog.Value = 1)
            global ConfirmExit    := (chkConfirmExit.Value = 1)
        } else if (cat["type"] = "msime") {
            ; DropDownList.Value は 1-based → レジストリ値 = Value - 1
            global SpaceInitVal   := msimeSpaceInit.Value - 1   ; 1→0, 2→1, 3→2
            global SpaceTargetVal := msimeSpaceTo.Value - 1
            global PunctInitVal   := msimePunctInit.Value - 1
            global PunctTargetVal := msimePunctTo.Value - 1
            ; 有効/無効のトグル
            newMsIme := (msimePanelChk.Value = 1)
            if (newMsIme != MsImeSettingsEnabled) {
                if newMsIme
                    EnableMsImeSettings()
                else
                    DisableMsImeSettings()
            } else if MsImeSettingsEnabled {
                UpdateMsImeMenu()
            }
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
