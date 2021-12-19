VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UXUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] UX/UIユーティリティクラス
'* [詳  細] ダイアログ等、UX/UIに関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'* [参  考]
'*          http://tetsucom.blogspot.com/2011/03/vba_621.html
'*
'* @author Bankura
'* Copyright (c) 2019-2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function GetInputState Lib "user32" () As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindowAsync Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare PtrSafe Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
Private Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As LongPtr
Private Declare PtrSafe Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare PtrSafe Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As LongPtr, phiconSmall As LongPtr, ByVal nIcons As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

'******************************************************************************
'* Enum定義
'******************************************************************************
Public Enum NOTIFYICONINFO
    NIIF_NONE = &H0      'アイコンなし
    NIIF_INFO = &H1      '「情報」アイコン
    NIIF_WARNING = &H2   '「警告」アイコン
    NIIF_ERROR = &H3     '「エラー」アイコン
    NIIF_USER = &H4
    NIIF_ICON_MASK = &HF
    NIIF_NOSOUND = &H10
End Enum

'******************************************************************************
'* 定数定義
'******************************************************************************
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1 'uCallbackMessage
Private Const NIF_ICON = &H2    'hIcon
Private Const NIF_TIP = &H4     'szTip
Private Const NIF_STATE = &H8   'Version 5.0 以降： dwState, dwStateMask
Private Const NIF_INFO = &H10   'Version 5.0 以降： szInfo, uTimeout, szInfoTitle, dwInfoFlags
Private Const NIF_GUID = &H20   'Version 6.0 以降： 予約されています。
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504
Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_THICKFRAME = &H40000

'******************************************************************************
'構造体定義
'******************************************************************************
Private Type GUID
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(0 To 7) As Byte
End Type

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As LongPtr
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As LongPtr
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
    guidItem As GUID
End Type


'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mTime As Variant
Private mIcon As NOTIFYICONDATA

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] DoEvents実行処理
'* [詳  細] DoEventsを最適なタイミングで実行する。
'*
'******************************************************************************
Public Sub CheckEvents()
    If GetInputState() Or (DateDiff("s", mTime, time) > 3) Then
        DoEvents
        mTime = time
    End If
End Sub

'******************************************************************************
'* [概  要] WaitFor
'* [詳  細] 指定した秒だけ停止する関数
'*
'* @param second 停止する秒数
'*
'******************************************************************************
Public Sub WaitFor(ByVal second As Integer)
    Dim futureTime As Date
    futureTime = DateAdd("s", second, Now)
    While Now < futureTime
        Sleep 100
        CheckEvents
    Wend
End Sub

'******************************************************************************
'* [概  要] ShowOpenFileDialog メソッド
'* [詳  細] ファイルダイアログを表示する。
'*          シートを指定している場合、選択したファイルパスを指定した
'*          シートのセルに設定する。
'*
'* @param sheetName シート名
'* @param pos セル位置（A1形式）
'* @return String 選択したファイルパス
'*
'******************************************************************************
Public Function ShowOpenFileDialog(Optional sheetName As String, Optional pos As String = "A1") As String
    Dim fileName: fileName = Application.ShowOpenFileDialog
    If fileName <> False Then
        ShowOpenFileDialog = fileName
        If sheetName <> "" Then
            ThisWorkbook.Worksheets(sheetName).Range(pos).Value = fileName
        End If
    End If
End Function

'******************************************************************************
'* [概  要] ShowCsvFileDialog メソッド
'* [詳  細] CSVファイルを選択するダイアログを表示します｡
'*
'* @param initialFile 初期表示用ファイルパス。省略可能。
'* @return 選択したファイルのパス
'*
'******************************************************************************
Public Function ShowCsvFileDialog(Optional ByVal initialFile As String) As String
    ShowCsvFileDialog = ""
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv"
        .InitialFileName = initialFile
        .AllowMultiSelect = False
        If .Show = True Then
            ShowCsvFileDialog = .SelectedItems(1)
        End If
    End With
End Function

'******************************************************************************
'* [概  要] ShowCsvSaveFileDialog メソッド
'* [詳  細] CSVファイルを保存するダイアログを表示する｡
'*
'* @param initialFile 初期表示用ファイルパス。省略可能。
'* @return 保存ファイルのパス
'*
'******************************************************************************
Public Function ShowCsvSaveFileDialog(Optional ByVal initialFile As String) As String
    Dim fileName As Variant
    fileName = Application.GetSaveAsFilename(initialFile, "CSVファイル,*.csv")
    If fileName = False Then
        Exit Function
    End If
    ShowCsvSaveFileDialog = CStr(fileName)
End Function

'******************************************************************************
'* [概  要] ShowFolderDialog メソッド
'* [詳  細] フォルダダイアログを表示する｡
'*
'* @param sheetName シート名
'* @param pos セル位置（A1形式）
'* @return String 選択したファイルパス
'*
'******************************************************************************
Public Function ShowFolderDialog(Optional sheetName As String, Optional pos As String = "A1") As String
    On Error GoTo ErrorHandler
    
    'フォルダダイアログ表示
    Dim objPath As Object
    Set objPath = Base.GetShell().BrowseForFolder(0, "フォルダを選択してください", &H1)
    If Not objPath Is Nothing Then
        ShowFolderDialog = CStr(objPath.items.Item.Path)
        If sheetName <> "" Then
            ThisWorkbook.Worksheets(sheetName).Range(pos).Value = CStr(objPath.items.Item.Path)
        End If
    End If
    Exit Function
    
ErrorHandler:
    Call MsgBox("指定できないフォルダです。" & vbNewLine & "別のフォルダを指定してください。", vbOKOnly)
End Function

'******************************************************************************
'* [概  要] ShowForeground
'* [詳  細] 指定されたウィンドウを最前面化する関数
'*
'* @param  hwnd 対象ウィンドウのハンドル番号
'*
'******************************************************************************
Public Sub ShowForeground(ByVal hWnd As LongPtr)
    If IsIconic(hWnd) Then
        Call ShowWindowAsync(hWnd, SW_RESTORE)
    End If
    Call SetForegroundWindow(hWnd)
End Sub

'******************************************************************************
'* [概  要] ShowForegroundMax
'* [詳  細] 指定されたウィンドウを最前面化・最大化する関数
'*
'* @param hwnd 対象ウィンドウのハンドル番号
'*
'******************************************************************************
Public Sub ShowForegroundMax(ByVal hWnd As LongPtr)
    Call ShowWindowAsync(hWnd, SW_SHOWMAXIMIZED)
    Call SetForegroundWindow(hWnd)
End Sub

'******************************************************************************
'* [概  要] ShowMinimizedNoActive
'* [詳  細] 指定されたウィンドウを最小化（アクティブ化しない）する関数
'*
'* @param hwnd 対象ウィンドウのハンドル番号
'*
'******************************************************************************
Public Sub ShowMinimizedNoActive(ByVal hWnd As LongPtr)
    If Not IsIconic(hWnd) Then
        Call ShowWindowAsync(hWnd, SW_SHOWMINNOACTIVE)
    End If
End Sub

'******************************************************************************
'* [概  要] DoEventsForPaint
'* [詳  細] DoEvents代替処理。WM_PAINTメッセージのみを処理する。
'*
'******************************************************************************
Public Sub DoEventsForPaint()
    Dim message  As msg
    If PeekMessage(message, 0, WM_PAINT, WM_PAINT, PM_REMOVE) Then
        Call TranslateMessage(message)
        Call DispatchMessage(message)
    End If
End Sub

'******************************************************************************
'* [概  要] ChangeNumLockOn
'* [詳  細] NumLockキーがオフの場合はオンにする。
'*
'******************************************************************************
Public Sub ChangeNumLockOn()
    Dim bState  As Boolean
    Dim keys(0 To 255) As Byte
    
    Call GetKeyboardState(keys(0))
    bState = keys(VK_NUMLOCK)
    
    If Not bState Then
        Call keybd_event(VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0)
        Call keybd_event(VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    End If
End Sub

'******************************************************************************
'* [概  要] ChangeMousePoint
'* [詳  細] マウスカーソルの位置を変更する。
'*
'******************************************************************************
Public Sub ChangeMousePoint(x As Long, y As Long)
    Dim mousePoint As POINTAPI
    Dim mousePSet As Long
    
    ' マウスカーソル位置指定
    mousePSet = SetCursorPos(x, y)
    
'    ' マウスカーソル位置取得
'    Call GetCursorPos(mousePoint)
'    With mousePoint
'        x = .x
'        y = .y
'    End With
End Sub

'******************************************************************************
'* [概  要] NotifyToast
'* [詳  細] システムの通知領域にトースト通知を表示する。
'*          タスクトレイにアイコン追加、トースト通知、アイコン削除の
'*          処理を一括して行う。
'*
'* @param msg                トースト通知するメッセージ
'* @param toastTitle         トースト通知のタイトル
'* @param toastIcon          アイコンの種類（NIIF_INFO：アイコンなし、NIIF_INFO：情報、NIIF_WARNING：警告、NIIF_ERROR：エラー など）
'* @param toastTimeOutSecond トースト通知がタイムアウトするまでの時間(秒)
'* @param hwnd               アイコンの親となるウィンドウハンドル
'* @param toolTip            アイコンの上にマウスを持ってきたときに表示されるメッセージ
'* @param iconResourceFile   アイコンのリソース(DLL,EXEファイル。~EXCEL.EXEなど)
'*
'******************************************************************************
Public Sub NotifyToast(ByVal msg As String, _
                       Optional ByVal toastTitle As String = "Notifyicon", _
                       Optional ByVal toastIcon As NOTIFYICONINFO, _
                       Optional ByVal toastTimeOutSecond As Long = 10, _
                       Optional ByVal hWnd As LongPtr = 0, _
                       Optional ByVal toolTip As String = "Notifyicon", _
                       Optional ByVal iconResourceFile As String = "")
    If hWnd = 0 Then
        hWnd = Application.hWnd
    End If
    Call AddIcon(hWnd, toolTip, iconResourceFile)
    Call ShowBalloon(msg, toastTitle, toastIcon, toastTimeOutSecond)
    Me.WaitFor 3    ' 通知前にアイコンを削除すると通知元の表示がおかしくなるため、3秒待機
    Call DeleteIcon
End Sub

'******************************************************************************
'* [概  要] AddIcon
'* [詳  細] アイコンをタスクトレイに追加する。
'*
'* @param hwnd             アイコンの親となるウィンドウハンドル
'* @param toolTip          アイコンの上にマウスを持ってきたときに表示されるメッセージ
'* @param iconResourceFile アイコンのリソース(DLL,EXEファイル。~EXCEL.EXEなど)
'* @return Long Shell_NotifyIcon関数の戻り値
'*
'******************************************************************************
Public Function AddIcon(ByVal hWnd As LongPtr, Optional ByVal toolTip As String = "Notifyicon", Optional ByVal iconResourceFile As String = "") As Long
    Dim iconSmall As LongPtr
    If iconResourceFile = "" Then
        iconResourceFile = Application.Path & "\" & "EXCEL.EXE"
    End If
    
    Call GetExtractIcon(iconResourceFile, iconSmall)
    
    With mIcon
        .cbSize = NOTIFYICONDATA_V3_SIZE
        .hWnd = hWnd
        .uID = 0
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO Or NIF_GUID
        .hIcon = iconSmall
        .szTip = toolTip & vbNullChar
        .szInfoTitle = vbNullChar
        .szInfo = vbNullChar
        .uTimeout = 0
        .dwInfoFlags = 0
    End With
    
    DrawMenuBar hWnd
    
    AddIcon = Shell_NotifyIcon(NIM_ADD, mIcon)
End Function

'******************************************************************************
'* [概  要] ShowBalloon
'* [詳  細] バルーンチップ（トースト通知）を表示する。
'* [注  意] 実行前に、AddIconでタスクトレイにアイコンを追加する必要がある。
'*
'* @param msg                  バルーンに表示するメッセージ
'* @param balloonTitle         バルーンタイトル
'* @param balloonIcon          NOTIFYICONINFO構造体
'* @param balloonTimeOutSecond バルーン表示がタイムアウトするまでの時間(秒)
'* @return Long Shell_NotifyIcon関数の戻り値
'*
'******************************************************************************
Public Function ShowBalloon(ByVal msg As String, Optional ByVal balloonTitle As String = "Notifyicon", Optional ByVal balloonIcon As NOTIFYICONINFO, Optional ByVal balloonTimeOutSecond As Long = 10) As Long
    With mIcon
        .uFlags = NIF_INFO
        .szInfoTitle = balloonTitle & vbNullChar
        .szInfo = msg & vbNullChar
        .uTimeout = balloonTimeOutSecond * 1000 ' タイムアウト時間(ms)
        .dwInfoFlags = balloonIcon
    End With

    ShowBalloon = Shell_NotifyIcon(NIM_MODIFY, mIcon)
End Function

'******************************************************************************
'* [概  要] ModifyIcon
'* [詳  細] アイコンのリソースを変更する。
'* [注  意] 実行前に、AddIconでタスクトレイにアイコンを追加する必要がある。
'*
'* @param iconResourceFile アイコンのリソース(DLL,EXEファイル。~EXCEL.EXEなど)
'* @return Long Shell_NotifyIcon関数の戻り値
'*
'******************************************************************************
Public Function ModifyIcon(ByVal iconResourceFile) As Long
    Dim iconSmall As LongPtr
    Call GetExtractIcon(iconResourceFile, iconSmall)
    
    With mIcon
        .hIcon = iconSmall
        .uFlags = NIF_ICON
    End With
    
    ModifyIcon = Shell_NotifyIcon(NIM_MODIFY, mIcon)
End Function

'******************************************************************************
'* [概  要] DeleteIcon
'* [詳  細] 作成したアイコンを削除する。
'* [注  意] 実行前に、AddIconでタスクトレイにアイコンを追加する必要がある。
'*
'******************************************************************************
Public Sub DeleteIcon()
    With mIcon
        .uFlags = NIF_ICON
        .uID = 0
    End With
    
    Call Shell_NotifyIcon(NIM_DELETE, mIcon)
End Sub

'******************************************************************************
'* [概  要] GetExtractIcon
'* [詳  細] ExtractIconEx関数のラッパーメソッド。
'*
'* @param iconResourceFile アイコンリソースファイル
'* @param iconSmall        小さなアイコンのハンドルを格納する変数
'* @param iconIndex        リソースから取得するアイコンのインデックス
'* @return Long ExtractIconEx関数の戻り値
'*
'******************************************************************************
Private Function GetExtractIcon(ByVal iconResourceFile As String, ByRef iconSmall As LongPtr, Optional ByVal iconIndex As Long = 0) As Long
    Dim iconLarge As LongPtr

    If Dir(iconResourceFile) = "" Then
        Debug.Print iconResourceFile & "が見つかりません。"
        GetExtractIcon = -1
        Exit Function
    End If
        
    GetExtractIcon = ExtractIconEx(iconResourceFile, iconIndex, iconLarge, iconSmall, 1)
End Function

'******************************************************************************
'* [概  要] FormResizeable
'* [詳  細] フォームをサイズ変更可能（かつ最小、最大ボタンを追加）とする。
'* [備  考] Formクラスから呼び出すこと。
'*
'******************************************************************************
Public Sub FormResizeable()
    Dim hWnd As LongPtr: hWnd = GetActiveWindow()
    Dim style As Long: style = GetWindowLong(hWnd, GWL_STYLE)
    
    ' ウインドウサイズ可変＋最小ボタン＋最大ボタンを追加
    style = style Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX

    Call SetWindowLong(hWnd, GWL_STYLE, style)
End Sub

'******************************************************************************
'* [概  要] IsVBEWindowActive
'* [詳  細] アクティブウィンドウがVBEかどうか判定する。
'*
'* @return Boolean アクティブウィンドウがVBEかどうか
'******************************************************************************
Public Function IsVBEWindowActive() As Boolean
    IsVBEWindowActive = (GetActiveWindow() = FindWindow("wndclass_desked_gsk", vbNullString))
End Function

'End
