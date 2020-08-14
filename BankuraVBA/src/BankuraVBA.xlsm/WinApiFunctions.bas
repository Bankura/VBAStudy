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

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 構造体定義
'******************************************************************************
Private Type WbkDtl
    phwnd   As LongPtr
    hWnd    As LongPtr
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
Public wD() As WbkDtl

'******************************************************************************
' [関数名] EnumerateWindow
' [説　明] ウィンドウを列挙するためのコールバック関数。
' [引　数] hwnd ウィンドウハンドル
'          lParam lParam
' [戻り値] Boolean
'******************************************************************************
Public Function EnumerateWindow(ByVal hWnd As LongPtr, lParam As Long) As Boolean

    If GetWinAPI.IsWindowVisible(hWnd) Then
        Call PrintCaptionAndProcess(hWnd)
    End If
    EnumerateWindow = True
End Function

'******************************************************************************
' [関数名] PrintCaptionAndProcess
' [説　明] ウィンドウのキャプションとプロセス名を表示する。
' [引　数] hwnd ウィンドウハンドル
'          lParam lParam
' [戻り値] Boolean
'******************************************************************************
Private Sub PrintCaptionAndProcess(ByVal hWnd As LongPtr)
    Dim strClassBuff As String * 128
    Dim strTextBuff  As String * 516
    Dim strClass     As String
    Dim strText      As String
    Dim lngRtnCode   As Long
    Dim lngProcesID  As Long
    
    ' クラス名取得
    lngRtnCode = GetWinAPI.GetClassName(hWnd, strClassBuff, Len(strClassBuff))
    strClass = Left(strClassBuff, InStr(strClassBuff, vbNullChar) - 1)
    
    ' ウィンドウのキャプションを取得
    lngRtnCode = GetWinAPI.GetWindowText(hWnd, strTextBuff, Len(strTextBuff))
    strText = Left(strTextBuff, InStr(strTextBuff, vbNullChar) - 1)
    
    ' プロセスIDを取得
    lngRtnCode = GetWinAPI.GetWindowThreadProcessId(hWnd, lngProcesID)

    Debug.Print "PID:" & lngProcesID, "hwnd:" & hWnd, "クラス名:" & strClass, "Caption:" & strText
End Sub


'******************************************************************************
' [関数名] GetEnumWindowProcessId
' [説　明] ウィンドウのプロセスIDを取得するためのコールバック関数。
' [引　数] hwnd ウィンドウハンドル
'          lParam lParam
' [戻り値] Boolean
'******************************************************************************
Public Function GetEnumWindowProcessId(ByVal hWnd As LongPtr, lParam As Long) As Boolean
    Dim lngRtnCode  As Long
    Dim lngProcesID  As Long
    If GetWinAPI.IsWindowVisible(hWnd) Then
      '  プロセスIDを取得
        lngRtnCode = GetWinAPI.GetWindowThreadProcessId(hWnd, lngProcesID)
        ' 指定したプロセスIDと一致する場合はウィンドウハンドルを設定
        If lngProcesID = tmpProcessId Then
            tmpHwnd = hWnd
        End If
    End If
    GetEnumWindowProcessId = True
End Function

'******************************************************************************
' [関数名] PrintCaptionAndProcessMain
' [説　明] ウィンドウのキャプションとプロセス名を表示する。
' [引　数] なし
' [戻り値] なし
'******************************************************************************
Public Sub PrintCaptionAndProcessMain()
    Dim lngRtnCode  As Long

    lngRtnCode = GetWinAPI.EnumWindows(AddressOf EnumerateWindow, 0&)
End Sub


' 使用不可：うまく動かない。原因解明のため残しておく。
'******************************************************************************
' [関数名] GetHwndByPid
' [説　明] 指定したプロセスIDに対応したウィンドウハンドルを取得する。
' [引　数] pId プロセスID
' [戻り値] LongPtr ウィンドウハンドル
'******************************************************************************
Public Function GetHwndByPid(pid As Long) As LongPtr

    Dim hWnd As LongPtr
    Dim pIdLast As Long
    GetHwndByPid = 0
    
    'デスクトップWindowの子WindowのうちトップレベルのWindowを取得
    hWnd = GetWinAPI.GetDesktopWindow()
    hWnd = GetWinAPI.GetWindow(hWnd, 5)

    Do While (0 <> hWnd)
        pIdLast = 0
        Call GetWinAPI.GetWindowThreadProcessId(hWnd, pIdLast)
'        Debug.Print "hwnd:" & hwnd & " pIdLast:" & pIdLast & "  pid:" & pid
        If (pid = pIdLast) Then
            GetHwndByPid = hWnd
            Exit Do
        End If
        '次のWindowハンドルを取得
        hWnd = GetWinAPI.GetWindow(hWnd, 2)
    Loop
End Function

'******************************************************************************
' [関数名] GetHwndByPid
' [説　明] 表示されているウィンドウのハンドルをプロセスIDから取得する。
'          （EnumWindowsを利用）
' [引　数] pid
' [戻り値] LongPtr
'******************************************************************************
Public Function GetHwndByPid2(pid As Long) As LongPtr
    Dim lngRtnCode  As Long
    tmpProcessId = 0
    tmpHwnd = 0
    
    tmpProcessId = pid
    lngRtnCode = GetWinAPI.EnumWindows(AddressOf GetEnumWindowProcessId, 0)
    GetHwndByPid2 = tmpHwnd
End Function


'******************************************************************************
'* [概  要] EnumChildSubProc
'* [詳  細] EnumChildWindows APIコールバック関数
'*
'* @param hwndChild 子ウィンドウハンドル
'* @param lParam LPARAM
'* @return Long 処理結果
'******************************************************************************
Public Function EnumChildSubProc(ByVal hWndChild As LongPtr, _
                                ByVal lParam As Long) As Long

    Dim strClassBuff As String * 128
    Dim strClass     As String
    Dim strTextBuff  As String * 516
    Dim strText      As String
    Dim lngRtnCode   As Long
    
    ' クラス名取得
    lngRtnCode = GetWinAPI.GetClassName(hWndChild, strClassBuff, Len(strClassBuff))
    strClass = Left(strClassBuff, InStr(strClassBuff, vbNullChar) - 1)
    If strClass = "EXCEL7" Then
        ' テキストをバッファにする
        lngRtnCode = GetWinAPI.GetWindowText(hWndChild, strTextBuff, Len(strTextBuff))
        strText = Left(strTextBuff, InStr(strTextBuff, vbNullChar) - 1)
        If InStr(1, strText, ".xla") = 0 Then
            If Sgn(wD) = 0 Then
                ReDim wD(0)
                wD(0).phwnd = tmpParentHwnd
                wD(0).hWnd = hWndChild
            Else
                ReDim Preserve wD(UBound(wD) + 1)
                wD(UBound(wD)).phwnd = tmpParentHwnd
                wD(UBound(wD)).hWnd = hWndChild
            End If
        End If

    End If
EnumChildPass:
    EnumChildSubProc = True
End Function

'******************************************************************************
'* [概  要] EnumWindowsProc
'* [詳  細] EnumWindows APIコールバック関数
'*
'* @param hwnd ハンドル
'* @param lParam LPARAM
'* @return Long 処理結果
'******************************************************************************
Public Function EnumWindowsProc(ByVal hWnd As LongPtr, _
                                ByVal lParam As LongPtr) As Long

    Dim strClassBuff As String * 128
    Dim strClass     As String
    Dim lngRtnCode   As Long
    Dim lngThreadID  As Long
    Dim lngProcesID  As Long
    
    ' クラス名取得
    lngRtnCode = GetWinAPI.GetClassName(hWnd, strClassBuff, Len(strClassBuff))
    strClass = Left(strClassBuff, InStr(strClassBuff, vbNullChar) - 1)
    
    If strClass = "XLMAIN" Then
        tmpParentHwnd = hWnd
        ' 子ウィンドウを列挙
        lngRtnCode = GetWinAPI.EnumChildWindows(hWnd, _
                         AddressOf EnumChildSubProc, lParam)
                        
    End If
EnumPass:
    EnumWindowsProc = True
End Function

'******************************************************************************
'* [概  要] GetExcelBookProc
'* [詳  細] 別プロセスExcelブックオブジェクト取得関数メイン
'*
'* @return Boolean 処理結果（True:正常 False：異常）
'******************************************************************************
Public Function GetExcelBookProc() As Boolean
    Dim lngRtnCode  As Long
    Dim i           As Long
    Dim iArr()      As Integer
    Dim wD2()       As WbkDtl
    Dim cnt         As Long

    Erase wD
    ' ワークブックのウィンドウハンドルを取得
    lngRtnCode = GetWinAPI.EnumWindows(AddressOf EnumWindowsProc, 0)
    If Sgn(wD) <> 0 Then
        cnt = 0
        For i = 0 To UBound(wD)
            If GetExcelBook(wD(i)) Then
                ReDim Preserve iArr(cnt)
                iArr(cnt) = i
                cnt = cnt + 1
            End If
        Next
        
        ReDim wD2(0 To UBound(iArr)) As WbkDtl
        For i = 0 To UBound(iArr)
            wD2(i).hWnd = wD(iArr(i)).hWnd
            wD2(i).phwnd = wD(iArr(i)).phwnd
            Set wD2(i).wkb = wD(iArr(i)).wkb
        Next

        Erase wD
        ReDim wD(0 To UBound(wD2)) As WbkDtl
        For i = 0 To UBound(wD2)
            wD(i).hWnd = wD2(i).hWnd
            wD(i).phwnd = wD2(i).phwnd
            Set wD(i).wkb = wD2(i).wkb
        Next
        Erase wD2
        GetExcelBookProc = True
    Else
        GetExcelBookProc = False
    End If

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
    
    If GetWinAPI.IsWindow(wDl.hWnd) = 0 Then Exit Function
    lngResult = GetWinAPI.SendMessage(wDl.hWnd, WM_GETOBJECT, 0, OBJID_NATIVEOM)
    If lngResult Then
        bytID = IID_IDispatch & vbNullChar

        GetWinAPI.IIDFromString bytID(0), IID(0)
        lngRtnCode = GetWinAPI.ObjectFromLresult(lngResult, IID(0), 0, wbw)
        If Not wbw Is Nothing Then Set wDl.wkb = wbw.Parent
        GetExcelBook = True
    End If
    
End Function
