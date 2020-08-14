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
Private mWinApi As WinAPI

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mWinApi = Base.GetWinAPI()
End Sub
Private Sub Class_Terminate()
    Set mWinApi = Nothing
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
    If mWinApi.GetInputState() Or (DateDiff("s", mTime, time) > 3) Then
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
        mWinApi.Sleep 100
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
            ThisWorkbook.Worksheets(sheetName).Range(pos).value = fileName
        End If
    End If
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
        ShowFolderDialog = CStr(objPath.Items.Item.path)
        If sheetName <> "" Then
            ThisWorkbook.Worksheets(sheetName).Range(pos).value = CStr(objPath.Items.Item.path)
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
    If mWinApi.IsIconic(hWnd) Then
        Call mWinApi.ShowWindowAsync(hWnd, SW_RESTORE)
    End If
    Call mWinApi.SetForegroundWindow(hWnd)
End Sub

'******************************************************************************
'* [概  要] ShowForegroundMax
'* [詳  細] 指定されたウィンドウを最前面化・最大化する関数
'*
'* @param hwnd 対象ウィンドウのハンドル番号
'*
'******************************************************************************
Public Sub ShowForegroundMax(ByVal hWnd As LongPtr)
    Call mWinApi.ShowWindowAsync(hWnd, SW_SHOWMAXIMIZED)
    Call mWinApi.SetForegroundWindow(hWnd)
End Sub

'******************************************************************************
'* [概  要] ShowMinimizedNoActive
'* [詳  細] 指定されたウィンドウを最小化（アクティブ化しない）する関数
'*
'* @param hwnd 対象ウィンドウのハンドル番号
'*
'******************************************************************************
Public Sub ShowMinimizedNoActive(ByVal hWnd As LongPtr)
    If Not mWinApi.IsIconic(hWnd) Then
        Call mWinApi.ShowWindowAsync(hWnd, SW_SHOWMINNOACTIVE)
    End If
End Sub

'******************************************************************************
'* [概  要] DoEventsForPaint
'* [詳  細] DoEvents代替処理。WM_PAINTメッセージのみを処理する。
'*
'******************************************************************************
Public Sub DoEventsForPaint()
    Dim message  As msg
    If mWinApi.PeekMessage(message, 0, WM_PAINT, WM_PAINT, PM_REMOVE) Then
        Call mWinApi.TranslateMessage(message)
        Call mWinApi.DispatchMessage(message)
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
    
    Call mWinApi.GetKeyboardState(keys(0))
    bState = keys(VK_NUMLOCK)
    
    If Not bState Then
        Call mWinApi.keybd_event(VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0)
        Call mWinApi.keybd_event(VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
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
    mousePSet = mWinApi.SetCursorPos(x, y)
    
'    ' マウスカーソル位置取得
'    Call mWinApi.GetCursorPos(mousePoint)
'    With mousePoint
'        x = .x
'        y = .y
'    End With
End Sub

