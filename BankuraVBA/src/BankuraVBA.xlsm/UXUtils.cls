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
'*          http://home.att.ne.jp/zeta/gen/excel/c04p06.htm
'*
'* @author Bankura
'* Copyright (c) 2019-2022 Bankura
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
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDc As Long) As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

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
Private Const NIF_MESSAGE = &H1 ' uCallbackMessage
Private Const NIF_ICON = &H2    ' hIcon
Private Const NIF_TIP = &H4     ' szTip
Private Const NIF_STATE = &H8   ' Version 5.0 以降： dwState, dwStateMask
Private Const NIF_INFO = &H10   ' Version 5.0 以降： szInfo, uTimeout, szInfoTitle, dwInfoFlags
Private Const NIF_GUID = &H20   ' Version 6.0 以降： 予約されています。
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504
Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_THICKFRAME = &H40000
Private Const PPI As Long = 72 ' Excel内で用いるポイント:72dpi
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const KEY_PRESSED = -32768
Private Const VK_LBUTTON = &H1 ' LeftClick
Private Const VK_RBUTTON = &H2 ' RightClick

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
'*-----------------------------------------------------------------------------
'* DPI プロパティ（読み取り専用）
'*
'* [補  足]
'* ・DPIとは、解像度を示す単位。
'* ・dots per inchの略で、1インチの幅の中にどれだけのドットを表現できるかを表す。
'* ・Windowsでは96dpiが標準とされている。
'*-----------------------------------------------------------------------------
Public Property Get DPI() As Long
    'デスクトップのウィンドウハンドル
    Dim hWnd As Long: hWnd = GetDesktopWindow()
  
    'デバイスコンテキストを取得
    Dim hDc As Long: hDc = GetDC(hWnd)
  
    'デバイス固有情報を取得
    DPI = GetDeviceCaps(hDc, LOGPIXELSX)
  
    'デバイスコンテキストを解放
    Call ReleaseDC(hWnd, hDc)
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] RepeatDoEvents
'* [詳  細] DoEventsを繰り返し実行し、画面描画停止の抑止を確実に行う。
'*
'* @param repCount DoEventsを行う回数（1カウントにつきDoEventsは2回実施）
'* @param withSleep Sleep処理を行うか（True:Yes）
'* @param sleepMSec Sleepの時間（millisecond）
'******************************************************************************
Public Sub RepeatDoEvents(Optional ByVal repCount As Long = 3, _
                          Optional ByVal withSleep As Boolean = True, _
                          Optional ByVal sleepMSec As Long = 20)
    Dim i As Long
    For i = 1 To repCount
        DoEvents
        DoEvents
        If withSleep Then Sleep sleepMSec
    Next
End Sub

'******************************************************************************
'* [概  要] DoEvents実行処理
'* [詳  細] DoEventsを最適なタイミングで実行する。
'*
'* @param second DoEventsを行う間隔（秒）
'******************************************************************************
Public Sub CheckEvents(Optional ByVal second As Integer = 3)
    If GetInputState() Or (DateDiff("s", mTime, time) > second) Then
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
'* @param fFilter   ファイルの種類を指定する文字列(ファイルフィルタ文字列)を指定
'*                  （Ex. Excelファイル,*.xlsx）
'* @param fIndex    fFilterの何番目のペアを既定値とするか（1始まり）
'* @param dlgTitle  ダイアログボックスのタイトル
'* @param initialFile 初期表示用ファイルパス。省略可能。
'* @param sheetName シート名
'* @param pos セル位置（A1形式）
'* @return String 選択したファイルパス
'*
'******************************************************************************
Public Function ShowOpenFileDialog(Optional fFilter As String, Optional fIndex As Long, Optional dlgTitle As String, Optional ByVal initialFile As String, Optional sheetName As String, Optional pos As String = "A1") As String
    If initialFile <> "" Then
        If IO.FolderExists(initialFile) Then
            Call FileUtils.ChangeDirectory(initialFile)
        Else
            With Application.FileDialog(msoFileDialogOpen)
                .Filters.Clear
                If fFilter <> "" Then
                    Dim vFiltersList As Variant: vFiltersList = Split(fFilter, ",")
                    If MathUtils.IsEven(ArrayUtils.GetLength(vFiltersList)) Then
                        Dim i As Long
                        For i = LBound(vFiltersList) To UBound(vFiltersList) Step 2
                            .Filters.Add vFiltersList(i), vFiltersList(i + 1)
                        Next
                        .FilterIndex = fIndex
                    End If
                End If
                .InitialFileName = initialFile
                .InitialView = msoFileDialogViewDetails
                .Title = dlgTitle
                .AllowMultiSelect = False
                If .Show = True Then
                    ShowOpenFileDialog = .SelectedItems(1)
                    If sheetName <> "" Then
                        ThisWorkbook.Worksheets(sheetName).Range(pos).Value = ShowOpenFileDialog
                    End If
                End If
            End With
            Exit Function
        End If
    End If
    ' FileDialogだとファイル入力欄にフォルダ名が表示されるため、初期表示用ファイルパスがフォルダの場合
    ' GetOpenFilename で対応
    Dim FileName: FileName = Application.GetOpenFilename(fFilter, fIndex, dlgTitle)
    If FileName <> False Then
        ShowOpenFileDialog = FileName
        If sheetName <> "" Then
            ThisWorkbook.Worksheets(sheetName).Range(pos).Value = FileName
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
    ShowCsvFileDialog = ShowOpenFileDialog("CSVファイル,*.csv", , , initialFile)
End Function

'******************************************************************************
'* [概  要] ShowSaveFileDialog メソッド
'* [詳  細] ファイルを保存するダイアログを表示する｡
'*
'* @param fFilter   ファイルの種類を指定する文字列(ファイルフィルタ文字列)を指定
'*                  （Ex. Excelファイル,*.xlsx）
'* @param fIndex    fFilterの何番目のペアを既定値とするか（1始まり）
'* @param dlgTitle  ダイアログボックスのタイトル
'* @param initialFile 初期表示用ファイルパス。省略可能。
'* @return String 保存ファイルのパス
'*
'******************************************************************************
Public Function ShowSaveFileDialog(Optional fFilter As String, Optional fIndex As Long, Optional dlgTitle As String, Optional ByVal initialFile As String, Optional sheetName As String, Optional pos As String = "A1") As String
    If initialFile <> "" Then
        If IO.FolderExists(initialFile) Then
            Call FileUtils.ChangeDirectory(initialFile)
        Else
            If fFilter = "" Then
                With Application.FileDialog(msoFileDialogSaveAs)
    
                    .InitialFileName = initialFile
                    .InitialView = msoFileDialogViewDetails
                    .Title = dlgTitle
                    If .Show = True Then
                        ShowSaveFileDialog = .SelectedItems(1)
                    End If
                End With
                Exit Function
            End If
        End If
    End If
    ' FileDialogだとフィルターが使用できない、かつファイル入力欄にフォルダ名が表示されるため、
    ' フィルター指定時と、初期表示用ファイルパスがフォルダの場合は、GetSaveAsFilename で対応
    Dim FileName: FileName = Application.GetSaveAsFilename(initialFile, fFilter, fIndex, dlgTitle)
    If FileName <> False Then
        ShowSaveFileDialog = FileName
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
    ShowCsvSaveFileDialog = ShowSaveFileDialog("CSVファイル,*.csv", , , initialFile)
End Function

'******************************************************************************
'* [概  要] ShowFolderDialog メソッド
'* [詳  細] フォルダダイアログを表示する｡
'*
'* @param dlgTitle      ダイアログボックスのタイトル
'* @param initialFolder 初期表示用フォルダパス。省略可能。
'* @param sheetName     シート名
'* @param pos           セル位置（A1形式）
'* @return String 選択したフォルダパス
'*
'******************************************************************************
Public Function ShowFolderDialog(Optional dlgTitle As String = "フォルダを開く", Optional ByVal initialFolder As String, Optional sheetName As String, Optional pos As String = "A1") As String
    With Application.FileDialog(msoFileDialogFolderPicker)

        .InitialFileName = initialFolder
        .InitialView = msoFileDialogViewDetails
        .Title = dlgTitle
        If .Show = True Then
            ShowFolderDialog = .SelectedItems(1)
            If sheetName <> "" Then
                ThisWorkbook.Worksheets(sheetName).Range(pos).Value = .SelectedItems(1)
            End If
        End If
    End With
End Function

'******************************************************************************
'* [概  要] ShowFolderTreeDialog メソッド
'* [詳  細] フォルダツリーダイアログを表示する｡
'*
'* @param dlgMessage ダイアログボックスのメッセージ
'* @param rootFolder ルートフォルダパス。省略可能。
'* @param sheetName シート名
'* @param pos セル位置（A1形式）
'* @return String 選択したフォルダパス
'*
'******************************************************************************
Public Function ShowFolderTreeDialog(Optional dlgMessage As String = "フォルダを選択してください", Optional ByVal rootFolder As Variant, Optional sheetName As String, Optional pos As String = "A1") As String
    On Error GoTo ErrorHandler
    
    'フォルダダイアログ表示
    Dim objPath As Object
    Set objPath = Base.GetShell().BrowseForFolder(0, dlgMessage, &H1, rootFolder)
    If Not objPath Is Nothing Then
        ShowFolderTreeDialog = CStr(objPath.items.Item.Path)
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
'* [概  要] ChangeMousePixelPosition
'* [詳  細] マウスカーソルの位置を変更する。
'*
'* @param x 横位置（Pixel）
'* @param y 縦位置（Pixel）
'******************************************************************************
Public Sub ChangeMousePixelPosition(x As Long, y As Long)
    ' マウスカーソル位置指定
    Call SetCursorPos(x, y)
End Sub

'******************************************************************************
'* [概  要] GetMousePixelPosition
'* [詳  細] マウスカーソルの位置を取得する。
'*
'* @param x 横位置（Pixel）
'* @param y 縦位置（Pixel）
'******************************************************************************
Public Sub GetMousePixelPosition(x As Long, y As Long)
    Dim mousePoint As POINTAPI

    ' マウスカーソル位置取得
    Call GetCursorPos(mousePoint)
    With mousePoint
        x = .x
        y = .y
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeMousePointPosition
'* [詳  細] マウスカーソルの位置を変更する（Point指定）。
'*
'* @param x 横位置（Point）
'* @param y 縦位置（Point）
'******************************************************************************
Public Sub ChangeMousePointPosition(x As Long, y As Long)
    Dim srcPixel As POINTAPI, destPoint As POINTAPI
    With srcPixel
        .x = x
        .y = y
    End With
    destPoint = ChangeScreenPixelToXlPoint(srcPixel)
    
    ' マウスカーソル位置指定
    Call SetCursorPos(destPoint.x, destPoint.y)
End Sub

'******************************************************************************
'* [概  要] GetMousePointPosition
'* [詳  細] マウスカーソルの位置を取得する。
'*
'* @param x 横位置（Point）
'* @param y 縦位置（Point）
'******************************************************************************
Public Sub GetMousePointPosition(x As Long, y As Long)
    Dim mousePixel As POINTAPI, destPoint As POINTAPI

    ' マウスカーソル位置取得
    Call GetCursorPos(mousePixel)
    destPoint = ChangeScreenPixelToXlPoint(mousePixel)
    With destPoint
        x = .x
        y = .y
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeXlPointToScreenPixel
'* [詳  細] Excel内で使う座標（単位：Point）から、スクリーン座標（単位：Pixel）
'*          に変換する。
'*          ActiveWindow情報シート上の位置（単位：Point）を、
'*          Windows 画面上の位置（単位：Pixel）に変換するため
'*          A1 セルの左上の位置をWindows 画面上のPixel位置で取得し、
'*          それに、シート上の位置（単位：Point）をPixelに変換して加算する。
'*
'* @param xlPoint Excel座標（Point）
'* @return POINTAPI Screen座標（Pixel）
'******************************************************************************
Public Function ChangeXlPointToScreenPixel(xlPoint As POINTAPI) As POINTAPI
    Dim xR1C1 As Long: xR1C1 = ActiveWindow.PointsToScreenPixelsX(0)
    Dim yR1C1 As Long: yR1C1 = ActiveWindow.PointsToScreenPixelsY(0)

    ChangeXlPointToScreenPixel.x = xR1C1 + PointToPixel(xlPoint.x)
    ChangeXlPointToScreenPixel.y = yR1C1 + PointToPixel(xlPoint.y)
End Function

'******************************************************************************
'* [概  要] ChangeScreenPixelToXlPoint
'* [詳  細] スクリーン座標（単位：Pixel）から、Excel内で使う座標（単位：Point）
'*          に変換する。
'*          Windows 画面上の位置（単位：Pixel）を、
'*          ActiveWindow情報シート上の位置（単位：Point）に変換するため
'*          A1 セルの左上の位置をWindows 画面上のPixel位置で取得し、
'*          シート上の位置（単位：Pixel）から減算後、Pointに変換する。
'*
'* @param xlPoint Excel座標（Point）
'* @return POINTAPI Screen座標（Pixel）
'******************************************************************************
Public Function ChangeScreenPixelToXlPoint(scPixel As POINTAPI) As POINTAPI
    Dim xR1C1 As Long: xR1C1 = ActiveWindow.PointsToScreenPixelsX(0)
    Dim yR1C1 As Long: yR1C1 = ActiveWindow.PointsToScreenPixelsY(0)

    ChangeScreenPixelToXlPoint.x = PixelToPoint(scPixel.x - xR1C1)
    ChangeScreenPixelToXlPoint.y = PixelToPoint(scPixel.y - yR1C1)
End Function

'******************************************************************************
'* [概  要] PointToPixel
'* [詳  細] Point値から、Pixel値に変換する。
'*
'* @param pointValue Pointの値
'* @return Long Pixelの値
'******************************************************************************
Public Function PointToPixel(ByVal pointValue As Long) As Long
    PointToPixel = (pointValue / PPI * DPI) * (ActiveWindow.Zoom / 100)
End Function

'******************************************************************************
'* [概  要] PixelToPoint
'* [詳  細] Pixel値から、Point値に変換する。
'*
'* @param pointValue Pixelの値
'* @return Long Pointの値
'******************************************************************************
Public Function PixelToPoint(ByVal pixelValue As Long) As Long
    PixelToPoint = pixelValue * PPI / DPI ' * (100 / ActiveWindow.Zoom)
End Function

'******************************************************************************
'* [概  要] IsMouseLeftClicked
'* [詳  細] マウスの左クリックが行われたか判定する。
'*
'* @return Boolean 左クリックが行われたか（True:Yes, False:No）
'******************************************************************************
Public Function IsMouseLeftClicked() As Boolean
    IsMouseLeftClicked = IsKeyPressed(VK_LBUTTON)
End Function

'******************************************************************************
'* [概  要] IsMouseRightClicked
'* [詳  細] マウスの右クリックが行われたか判定する。
'*
'* @return Boolean 右クリックが行われたか（True:Yes, False:No）
'******************************************************************************
Public Function IsMouseRightClicked() As Boolean
    IsMouseRightClicked = IsKeyPressed(VK_RBUTTON)
End Function

'******************************************************************************
'* [概  要] IsKeyPressed
'* [詳  細] 指定したキーが押されているか判定する。
'*
'* @return Boolean 右クリックが行われたか（True:Yes, False:No）
'******************************************************************************
Public Function IsKeyPressed(keyCd As Long) As Boolean
    IsKeyPressed = (GetAsyncKeyState(keyCd) And KEY_PRESSED) = KEY_PRESSED
End Function

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
