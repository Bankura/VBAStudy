Attribute VB_Name = "WinApiFunctions"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowsAPI関連関数モジュール
'* [詳  細] WindowsAPIに渡すコールバック関数等のWindowsAPI関連の処理を定義する。
'*
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (lpsz As Any, lpiid As Any) As Long
Private Declare PtrSafe Function ObjectFromLresult Lib "oleacc" (ByVal lResult As LongPtr, riid As Any, ByVal wParam As LongPtr, ppvObject As Any) As LongPtr

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 構造体定義
'******************************************************************************
Private Type WbkDtl
    phwnd   As LongPtr
    hwnd    As LongPtr
    wkb     As Excel.Workbook
End Type

'******************************************************************************
'* 定数定義
'******************************************************************************
Private Const OBJID_NATIVEOM = &HFFFFFFF0
Private Const OBJID_CLIENT = &HFFFFFFFC
Private Const IID_IMdcList = "{8BD21D23-EC42-11CE-9E0D-00AA006002F3}"
Private Const IID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
Private Const IID_IDispatch = "{00020400-0000-0000-C000-000000000046}"

'******************************************************************************
'* 変数定義
'******************************************************************************
Private tmpProcessId As Long
Private tmpHwnd As LongPtr
Private tmpParentHwnd As LongPtr
Public wd() As WbkDtl

'******************************************************************************
'* [概  要] PrintCaptionAndProcessMain
'* [詳  細] ウィンドウのキャプションとプロセス名を表示する。
'*
'******************************************************************************
Public Sub PrintCaptionAndProcessMain()
    Dim lngRtnCode  As Long

    lngRtnCode = EnumWindows(AddressOf EnumerateWindow, 0&)
End Sub

'******************************************************************************
'* [概  要] EnumerateWindow
'* [詳  細] ウィンドウを列挙するためのコールバック関数。
'*
'* @param hwnd ウィンドウハンドル
'* @param lParam lParam
'* @return Boolean
'******************************************************************************
Public Function EnumerateWindow(ByVal hwnd As LongPtr, lParam As Long) As Boolean

    If IsWindowVisible(hwnd) Then
        Call PrintCaptionAndProcess(hwnd)
    End If
    EnumerateWindow = True
End Function

'******************************************************************************
'* [概  要] PrintCaptionAndProcess
'* [詳  細] ウィンドウのキャプションとプロセス名を表示する。
'*
'* @param hwnd ウィンドウハンドル
'******************************************************************************
Private Sub PrintCaptionAndProcess(ByVal hwnd As LongPtr)
    Dim strClassBuff As String * 128
    Dim strTextBuff  As String * 516
    Dim strClass     As String
    Dim strText      As String
    Dim lngRtnCode   As Long
    Dim lngProcesID  As Long
    
    ' クラス名取得
    lngRtnCode = GetClassName(hwnd, strClassBuff, Len(strClassBuff))
    strClass = Left(strClassBuff, InStr(strClassBuff, vbNullChar) - 1)
    
    ' ウィンドウのキャプションを取得
    lngRtnCode = GetWindowText(hwnd, strTextBuff, Len(strTextBuff))
    strText = Left(strTextBuff, InStr(strTextBuff, vbNullChar) - 1)
    
    ' プロセスIDを取得
    lngRtnCode = GetWindowThreadProcessId(hwnd, lngProcesID)

    Debug.Print "PID:" & lngProcesID, "hwnd:" & hwnd, "クラス名:" & strClass, "Caption:" & strText
End Sub

'******************************************************************************
'* [概  要] GetExcelBookProc
'* [詳  細] 別プロセスExcelブックオブジェクト取得関数メイン。
'*          取得したExcelプロセス情報は、グローバル変数の wd 構造体に格納される。
'*
'* @return Boolean 処理結果（True:正常 False：異常）
'******************************************************************************
Public Function GetExcelBookProc() As Boolean
    Dim lngRtnCode  As Long
    Dim i           As Long
    Dim iArr()      As Integer
    Dim wD2()       As WbkDtl
    Dim cnt         As Long

    Erase wd
    ' ワークブックのウィンドウハンドルを取得
    lngRtnCode = EnumWindows(AddressOf EnumWindowsProc, 0)
    
    On Error Resume Next
    Dim cktemp As Long
    cktemp = UBound(wd)
    
    If Err.Number = 0 Then
        On Error GoTo 0
        cnt = 0
        For i = 0 To UBound(wd)
            If GetExcelBook(wd(i)) Then
                ReDim Preserve iArr(cnt)
                iArr(cnt) = i
                cnt = cnt + 1
            End If
        Next
        
        ReDim wD2(0 To UBound(iArr)) As WbkDtl
        For i = 0 To UBound(iArr)
            wD2(i).hwnd = wd(iArr(i)).hwnd
            wD2(i).phwnd = wd(iArr(i)).phwnd
            Set wD2(i).wkb = wd(iArr(i)).wkb
        Next

        Erase wd
        ReDim wd(0 To UBound(wD2)) As WbkDtl
        For i = 0 To UBound(wD2)
            wd(i).hwnd = wD2(i).hwnd
            wd(i).phwnd = wD2(i).phwnd
            Set wd(i).wkb = wD2(i).wkb
        Next
        Erase wD2
        GetExcelBookProc = True
    Else
        Err.Clear
        On Error GoTo 0
        GetExcelBookProc = False
    End If
End Function

'******************************************************************************
'* [概  要] EnumWindowsProc
'* [詳  細] EnumWindows APIコールバック関数
'*
'* @param hwnd ハンドル
'* @param lParam LPARAM
'* @return Long 処理結果
'******************************************************************************
Public Function EnumWindowsProc(ByVal hwnd As LongPtr, _
                                ByVal lParam As LongPtr) As Long

    Dim strClassBuff As String * 128
    Dim strClass     As String
    Dim lngRtnCode   As Long
    Dim lngThreadID  As Long
    Dim lngProcesID  As Long
    
    ' クラス名取得
    lngRtnCode = GetClassName(hwnd, strClassBuff, Len(strClassBuff))
    strClass = Left(strClassBuff, InStr(strClassBuff, vbNullChar) - 1)
    
    If strClass = "XLMAIN" Then
        tmpParentHwnd = hwnd
        ' 子ウィンドウを列挙
        lngRtnCode = EnumChildWindows(hwnd, _
                         AddressOf EnumChildSubProc, lParam)
                        
    End If
EnumPass:
    EnumWindowsProc = True
End Function

'******************************************************************************
'* [概  要] EnumChildSubProc
'* [詳  細] EnumChildWindows APIコールバック関数
'*
'* @param hwndChild 子ウィンドウハンドル
'* @param lParam LPARAM
'* @return Long 処理結果
'******************************************************************************
Public Function EnumChildSubProc(ByVal hwndChild As LongPtr, _
                                ByVal lParam As Long) As Long

    Dim strClassBuff As String * 128
    Dim strClass     As String
    Dim strTextBuff  As String * 516
    Dim strText      As String
    Dim lngRtnCode   As Long
    
    ' クラス名取得
    lngRtnCode = GetClassName(hwndChild, strClassBuff, Len(strClassBuff))
    strClass = Left(strClassBuff, InStr(strClassBuff, vbNullChar) - 1)
    If strClass = "EXCEL7" Then
        ' テキストをバッファにする
        lngRtnCode = GetWindowText(hwndChild, strTextBuff, Len(strTextBuff))
        strText = Left(strTextBuff, InStr(strTextBuff, vbNullChar) - 1)
        If InStr(1, strText, ".xla") = 0 Then
            On Error Resume Next
            Dim cktemp As LongPtr
            cktemp = UBound(wd)
        
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                ReDim wd(0)
                wd(0).phwnd = tmpParentHwnd
                wd(0).hwnd = hwndChild
            Else
                On Error GoTo 0
                ReDim Preserve wd(UBound(wd) + 1)
                wd(UBound(wd)).phwnd = tmpParentHwnd
                wd(UBound(wd)).hwnd = hwndChild
            End If
        End If

    End If
EnumChildPass:
    EnumChildSubProc = True
End Function

'******************************************************************************
'* [概  要] GetExcelBook
'* [詳  細] 別プロセスExcelブックオブジェクト取得関数
'*
'* @param wDl ハンドル＋Excelブック構造体
'*
'******************************************************************************
Private Function GetExcelBook(wDl As WbkDtl) As Boolean
    Dim bytID()     As Byte
    Dim IID(0 To 3) As LongPtr
    Dim lngResult   As LongPtr
    Dim lngRtnCode  As LongPtr

    Dim wbw         As Excel.Window
    GetExcelBook = False
    
    If IsWindow(wDl.hwnd) = 0 Then Exit Function
    lngResult = SendMessage(wDl.hwnd, WM_GETOBJECT, 0, ByVal OBJID_NATIVEOM)
    If lngResult Then
        bytID = IID_IDispatch & vbNullChar

        IIDFromString bytID(0), IID(0)
        lngRtnCode = ObjectFromLresult(lngResult, IID(0), 0, wbw)
        If Not wbw Is Nothing Then Set wDl.wkb = wbw.Parent
        GetExcelBook = True
    End If
    
End Function

'******************************************************************************
'* [概  要] GetEnumWindowProcessId
'* [詳  細] ウィンドウのプロセスIDを取得するためのコールバック関数。
'*
'* @param hwnd ウィンドウハンドル
'* @param lParam lParam
'* @return Boolean
'******************************************************************************
Public Function GetEnumWindowProcessId(ByVal hwnd As LongPtr, lParam As Long) As Boolean
    Dim lngRtnCode  As Long
    Dim lngProcesID  As Long
    If IsWindowVisible(hwnd) Then
      '  プロセスIDを取得
        lngRtnCode = GetWindowThreadProcessId(hwnd, lngProcesID)
        ' 指定したプロセスIDと一致する場合はウィンドウハンドルを設定
        If lngProcesID = tmpProcessId Then
            tmpHwnd = hwnd
        End If
    End If
    GetEnumWindowProcessId = True
End Function

'******************************************************************************
'* [概  要] GetHwndByPid
'* [詳  細] 指定したプロセスIDに対応したウィンドウハンドルを取得する。
'*          使用不可：うまく動かない。原因解明のため残しておく。
'*
'* @param pId プロセスID
'* @return LongPtr ウィンドウハンドル
'******************************************************************************
Public Function GetHwndByPid(pid As Long) As LongPtr

    Dim hwnd As LongPtr
    Dim pIdLast As Long
    GetHwndByPid = 0
    
    'デスクトップWindowの子WindowのうちトップレベルのWindowを取得
    hwnd = GetDesktopWindow()
    hwnd = GetWindow(hwnd, 5)

    Do While (0 <> hwnd)
        pIdLast = 0
        Call GetWindowThreadProcessId(hwnd, pIdLast)
'        Debug.Print "hwnd:" & hwnd & " pIdLast:" & pIdLast & "  pid:" & pid
        If (pid = pIdLast) Then
            GetHwndByPid = hwnd
            Exit Do
        End If
        '次のWindowハンドルを取得
        hwnd = GetWindow(hwnd, 2)
    Loop
End Function

'******************************************************************************
'* [概  要] GetHwndByPid
'* [詳  細] 表示されているウィンドウのハンドルをプロセスIDから取得する。
'*         （EnumWindowsを利用）
'*
'* @param pid プロセスID
'* @return LongPtr ウィンドウハンドル
'******************************************************************************
Public Function GetHwndByPid2(pid As Long) As LongPtr
    Dim lngRtnCode  As Long
    tmpProcessId = 0
    tmpHwnd = 0
    
    tmpProcessId = pid
    lngRtnCode = EnumWindows(AddressOf GetEnumWindowProcessId, 0)
    GetHwndByPid2 = tmpHwnd
End Function
