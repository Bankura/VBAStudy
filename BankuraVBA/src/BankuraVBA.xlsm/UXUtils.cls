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
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
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

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************

'******************************************************************************
'構造体定義
'******************************************************************************

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mTime As Variant

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
    
    'ファイルダイアログ表示
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

