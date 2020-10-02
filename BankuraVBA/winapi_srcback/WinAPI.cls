VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPIクラス
'* [詳  細] WindowAPIの機能を提供する。
'*
'* [参  考] Win32API_PtrSafe.TXT
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
'* 内部変数定義
'******************************************************************************

Private mWinApiAdvapi As WinApiAdvapi
Private mWinApiComdlg As WinApiComdlg
Private mWinApiGdi As WinApiGdi
Private mWinApiImm As WinApiImm
Private mWinApiKernel As WinApiKernel
Private mWinApiKernelWinBase As WinApiKernelWinbase
Private mWinApiLz As WinApiLz
Private mWinApiMpr As WinApiMpr
Private mWinApiNetapi As WinApiNetapi
Private mWinApiOle As WinApiOle
Private mWinApiOleacc As WinApiOleacc
Private mWinApiOleaut As WinApiOleaut
Private mWinApiShell As WinApiShell
Private mWinApiUrlmon As WinApiUrlmon
Private mWinApiUser As WinApiUser
Private mWinApiVersion As WinApiVersion
Private mWinApiWininet As WinApiWininet
Private mWinApiWinmm As WinApiWinmm
Private mWinApiWinspool As WinApiWinspool


'******************************************************************************
'* プロパティ定義
'******************************************************************************

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mWinApiAdvapi = New WinApiAdvapi
    Set mWinApiComdlg = New WinApiComdlg
    Set mWinApiGdi = New WinApiGdi
    Set mWinApiImm = New WinApiImm
    Set mWinApiKernel = New WinApiKernel
    Set mWinApiKernelWinBase = New WinApiKernelWinbase
    Set mWinApiLz = New WinApiLz
    Set mWinApiMpr = New WinApiMpr
    Set mWinApiNetapi = New WinApiNetapi
    Set mWinApiOle = New WinApiOle
    Set mWinApiOleacc = New WinApiOleacc
    Set mWinApiOleaut = New WinApiOleaut
    Set mWinApiShell = New WinApiShell
    Set mWinApiUrlmon = New WinApiUrlmon
    Set mWinApiUser = New WinApiUser
    Set mWinApiVersion = New WinApiVersion
    Set mWinApiWininet = New WinApiWininet
    Set mWinApiWinmm = New WinApiWinmm
    Set mWinApiWinspool = New WinApiWinspool
End Sub
Private Sub Class_Terminate()
    Set mWinApiAdvapi = Nothing
    Set mWinApiComdlg = Nothing
    Set mWinApiGdi = Nothing
    Set mWinApiImm = Nothing
    Set mWinApiKernel = Nothing
    Set mWinApiKernelWinBase = Nothing
    Set mWinApiLz = Nothing
    Set mWinApiMpr = Nothing
    Set mWinApiNetapi = Nothing
    Set mWinApiOle = Nothing
    Set mWinApiOleacc = Nothing
    Set mWinApiOleaut = Nothing
    Set mWinApiShell = Nothing
    Set mWinApiUrlmon = Nothing
    Set mWinApiUser = Nothing
    Set mWinApiVersion = Nothing
    Set mWinApiWininet = Nothing
    Set mWinApiWinmm = Nothing
    Set mWinApiWinspool = Nothing
End Sub

'*=============================================================================
'* ole32
'*=============================================================================

'******************************************************************************
'* [概  要] IIDFromString
'* [詳  細] WindowsAPI「 IIDFromString 」ラッパーメソッド。
'*          Lib "ole32"
'*
'******************************************************************************
Friend Function IIDFromString(lpsz As Byte, lpiid As LongPtr) As Long
    IIDFromString = mWinApiOle.WinApi_IIDFromString(lpsz, lpiid)
End Function


'*=============================================================================
'* oleacc
'*=============================================================================
'******************************************************************************
'* [概  要] ObjectFromLresult
'* [詳  細] WindowsAPI「 ObjectFromLresult 」ラッパーメソッド。
'*          Lib "oleacc"
'*
'******************************************************************************
Friend Function ObjectFromLresult(ByVal lResult As LongPtr, riid As LongPtr, ByVal wParam As LongPtr, ppvObject As Object) As LongPtr
    ObjectFromLresult = mWinApiOleacc.WinApi_ObjectFromLresult(lResult, riid, wParam, ppvObject)
End Function

'*=============================================================================
'* oleaut
'*=============================================================================

'******************************************************************************
'* [概  要] DispCallFunc
'* [詳  細] WindowsAPI「 DispCallFunc 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function DispCallFunc(ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc_ As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Integer, ByRef prgpvarg As LongPtr, ByRef pvargResult As Variant) As Long
    DispCallFunc = mWinApiOleaut.WinApi_DispCallFunc(pvInstance, oVft, cc_, vtReturn, cActuals, prgvt, prgpvarg, pvargResult)
End Function

'******************************************************************************
'* [概  要] SafeArrayAllocDescriptor
'* [詳  細] WindowsAPI「 SafeArrayAllocDescriptor 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function SafeArrayAllocDescriptor(ByVal cDims As Long, ByRef ppsaOut() As Variant) As Long
    SafeArrayAllocDescriptor = mWinApiOleaut.WinApi_SafeArrayAllocDescriptor(cDims, ppsaOut)
End Function

'*=============================================================================
'* urlmon
'*=============================================================================
'******************************************************************************
'* [概  要] URLDownloadToFile
'* [詳  細] WindowsAPI「 URLDownloadToFile 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function URLDownloadToFile(ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    URLDownloadToFile = mWinApiUrlmon.WinApi_URLDownloadToFile(pCaller, szURL, szFileName, dwReserved, lpfnCB)
End Function


'*=============================================================================
'* wininet
'*=============================================================================
'******************************************************************************
'* [概  要] DeleteUrlCacheEntry
'* [詳  細] WindowsAPI「 DeleteUrlCacheEntry 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function DeleteUrlCacheEntry(ByVal lpszUrlName As String) As Long
    DeleteUrlCacheEntry = mWinApiWininet.WinApi_DeleteUrlCacheEntry(lpszUrlName)
End Function


'*=============================================================================
'* shell32
'*=============================================================================

'******************************************************************************
'* [概  要] SHCreateDirectoryEx
'* [詳  細] WindowsAPI「 SHCreateDirectoryEx 」ラッパーメソッド。
'*          Lib "shell32"
'*
'******************************************************************************
Friend Function SHCreateDirectoryEx(ByVal hWnd As LongPtr, ByVal pszPath As String, ByVal psa As Long) As LongPtr
    SHCreateDirectoryEx = mWinApiShell.WinApi_SHCreateDirectoryEx(hWnd, pszPath, psa)
End Function

'*=============================================================================
'* kernel32
'*=============================================================================

'******************************************************************************
'* [概  要] Sleep
'* [詳  細] WindowsAPI「 Sleep 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'* @param ms 待機時間（ミリ秒）
'*
'******************************************************************************
Friend Sub Sleep(ByVal ms As Long)
    Call mWinApiKernel.WinApi_Sleep(ms)
End Sub

'******************************************************************************
'* [概  要] CopyMemory
'* [詳  細] WindowsAPI「 CopyMemory 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'******************************************************************************
Friend Sub CopyMemory(Destination As LongPtr, source As LongPtr, ByVal Length As LongPtr)
    Call mWinApiKernel.WinApi_CopyMemory(Destination, source, Length)
End Sub

'******************************************************************************
'* [概  要] CopyMemoryByRef
'* [詳  細] WindowsAPI「 CopyMemory 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'******************************************************************************
Friend Sub CopyMemoryByRef(ByRef Destination As LongPtr, ByRef source As LongPtr, ByVal Length As LongPtr, _
                           Optional destByVal As Boolean = False, Optional sourceByVal As Boolean = False)
    Call mWinApiKernel.WinApi_CopyMemoryByRef(Destination, source, Length, destByVal, sourceByVal)
End Sub

'******************************************************************************
'* [概  要] GetPrivateProfileString
'* [詳  細] WindowsAPI「 GetPrivateProfileString 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'******************************************************************************
Friend Function GetPrivateProfileString(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    GetPrivateProfileString = mWinApiKernel.WinApi_GetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFileName)
End Function

'******************************************************************************
'* [概  要] CreateFile
'* [詳  細] WindowsAPI「 CreateFile 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'******************************************************************************
Friend Function CreateFile(ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
    CreateFile = mWinApiKernel.WinApi_CreateFile(lpFileName, dwDesiredAccess, dwShareMode, lpSecurityAttributes, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile)
End Function

'******************************************************************************
'* [概  要] CloseHandle
'* [詳  細] WindowsAPI「 CloseHandle 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'******************************************************************************
Friend Function CloseHandle(ByVal hObject As LongPtr) As Long
    CloseHandle = mWinApiKernel.WinApi_CloseHandle(hObject)
End Function

'******************************************************************************
'* [概  要] LocalFileTimeToFileTime
'* [詳  細] WindowsAPI「 LocalFileTimeToFileTime 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'******************************************************************************
Friend Function LocalFileTimeToFileTime(lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
    LocalFileTimeToFileTime = mWinApiKernel.WinApi_LocalFileTimeToFileTime(lpLocalFileTime, lpFileTime)
End Function

'******************************************************************************
'* [概  要] SystemTimeToFileTime
'* [詳  細] WindowsAPI「 SystemTimeToFileTime 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'******************************************************************************
Friend Function SystemTimeToFileTime(lpSystemTime As SystemTime, lpFileTime As FILETIME) As Long
    SystemTimeToFileTime = mWinApiKernel.WinApi_SystemTimeToFileTime(lpSystemTime, lpFileTime)
End Function

'******************************************************************************
'* [概  要] SetFileTime
'* [詳  細] WindowsAPI「 SetFileTime 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'******************************************************************************
Friend Function SetFileTime(ByVal hFile As LongPtr, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long
    SetFileTime = mWinApiKernel.WinApi_SetFileTime(hFile, lpCreationTime, lpLastAccessTime, lpLastWriteTime)
End Function

'******************************************************************************
'* [概  要] GetFileTime
'* [詳  細] WindowsAPI「 GetFileTime 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'******************************************************************************
Friend Function GetFileTime(ByVal hFile As LongPtr, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long
    GetFileTime = mWinApiKernel.WinApi_GetFileTime(hFile, lpCreationTime, lpLastAccessTime, lpLastWriteTime)
End Function


'*=============================================================================
'* user32
'*=============================================================================

'******************************************************************************
'* [概  要] GetInputState
'* [詳  細] WindowsAPI「 GetInputState 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function GetInputState() As Long
    GetInputState = mWinApiUser.WinApi_GetInputState()
End Function

'******************************************************************************
'* [概  要] SetForegroundWindow
'* [詳  細] WindowsAPI「 SetForegroundWindow 」ラッパーメソッド。
'*          ウィンドウを強制的に最前面に表示させるAPI。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function SetForegroundWindow(ByVal hWnd As LongPtr) As Long
    SetForegroundWindow = mWinApiUser.WinApi_SetForegroundWindow(hWnd)
End Function

'******************************************************************************
'* [概  要] IsIconic
'* [詳  細] WindowsAPI「 IsIconic 」ラッパーメソッド。
'*          ウィンドウが最小化されているか調べるAPI。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function IsIconic(ByVal hWnd As LongPtr) As Long
    IsIconic = mWinApiUser.WinApi_IsIconic(hWnd)
End Function
    
'******************************************************************************
'* [概  要] ShowWindowAsync
'* [詳  細] WindowsAPI「 ShowWindowAsync 」ラッパーメソッド。
'*          ウィンドウを元の大きさに戻すAPI。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function ShowWindowAsync(ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    ShowWindowAsync = mWinApiUser.WinApi_ShowWindowAsync(hWnd, nCmdShow)
End Function

'******************************************************************************
'* [概  要] PeekMessage
'* [詳  細] WindowsAPI「 PeekMessage 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function PeekMessage(lpMsg As msg, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
    PeekMessage = mWinApiUser.WinApi_PeekMessage(lpMsg, hWnd, wMsgFilterMin, wMsgFilterMax, wRemoveMsg)
End Function

'******************************************************************************
'* [概  要] TranslateMessage
'* [詳  細] WindowsAPI「 TranslateMessage 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function TranslateMessage(lpMsg As msg) As Long
    TranslateMessage = mWinApiUser.WinApi_TranslateMessage(lpMsg)
End Function

'******************************************************************************
'* [概  要] DispatchMessage
'* [詳  細] WindowsAPI「 DispatchMessage 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function DispatchMessage(lpMsg As msg) As LongPtr
    DispatchMessage = mWinApiUser.WinApi_DispatchMessage(lpMsg)
End Function

'******************************************************************************
'* [概  要] keybd_event
'* [詳  細] WindowsAPI「 keybd_event 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Sub keybd_event(ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
    Call mWinApiUser.WinApi_keybd_event(bVk, bScan, dwFlags, dwExtraInfo)
End Sub

'******************************************************************************
'* [概  要] GetKeyboardState
'* [詳  細] WindowsAPI 「GetKeyboardState 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function GetKeyboardState(pbKeyState As Byte) As Long
    GetKeyboardState = mWinApiUser.WinApi_GetKeyboardState(pbKeyState)
End Function

'******************************************************************************
'* [概  要] SetCursorPos
'* [詳  細] WindowsAPI「 SetCursorPos 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function SetCursorPos(ByVal x As Long, ByVal y As Long) As Long
    SetCursorPos = mWinApiUser.WinApi_SetCursorPos(x, y)
End Function

'******************************************************************************
'* [概  要] GetCursorPos
'* [詳  細] WindowsAPI「 GetCursorPos 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function GetCursorPos(lpPoint As POINTAPI) As Long
    GetCursorPos = mWinApiUser.WinApi_GetCursorPos(lpPoint)
End Function

'******************************************************************************
'* [概  要] EnumChildWindows
'* [詳  細] WindowsAPI「 EnumChildWindows 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function EnumChildWindows(ByVal hwndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    EnumChildWindows = mWinApiUser.WinApi_EnumChildWindows(hwndParent, lpEnumFunc, lParam)
End Function

'******************************************************************************
'* [概  要] EnumWindows
'* [詳  細] WindowsAPI「 EnumWindows 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function EnumWindows(ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    EnumWindows = mWinApiUser.WinApi_EnumWindows(lpEnumFunc, lParam)
End Function

'******************************************************************************
'* [概  要] GetClassName
'* [詳  細] WindowsAPI「 GetClassName 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function GetClassName(ByVal hWnd As LongPtr, ByRef lpClassName As String, ByVal nMaxCount As Long) As Long
    GetClassName = mWinApiUser.WinApi_GetClassName(hWnd, lpClassName, nMaxCount)
End Function

'******************************************************************************
'* [概  要] GetDesktopWindow
'* [詳  細] WindowsAPI「 GetDesktopWindow 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function GetDesktopWindow() As LongPtr
    GetDesktopWindow = mWinApiUser.WinApi_GetDesktopWindow()
End Function

'******************************************************************************
'* [概  要] GetWindow
'* [詳  細] WindowsAPI「 GetWindow 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function GetWindow(ByVal hWnd As LongPtr, ByVal wCmd As Long) As LongPtr
    GetWindow = mWinApiUser.WinApi_GetWindow(hWnd, wCmd)
End Function

'******************************************************************************
'* [概  要] GetWindowText
'* [詳  細] WindowsAPI「 GetWindowText 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function GetWindowText(ByVal hWnd As LongPtr, ByRef lpString As String, ByVal cch As Long) As Long
    GetWindowText = mWinApiUser.WinApi_GetWindowText(hWnd, lpString, cch)
End Function

'******************************************************************************
'* [概  要] GetWindowThreadProcessId
'* [詳  細] WindowsAPI「 GetWindowThreadProcessId 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function GetWindowThreadProcessId(ByVal hWnd As LongPtr, lpdwProcessId As Long) As Long
    GetWindowThreadProcessId = mWinApiUser.WinApi_GetWindowThreadProcessId(hWnd, lpdwProcessId)
End Function

'******************************************************************************
'* [概  要] IsWindow
'* [詳  細] WindowsAPI「 IsWindow 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function IsWindow(ByVal hWnd As LongPtr) As LongPtr
    IsWindow = mWinApiUser.WinApi_IsWindow(hWnd)
End Function

'******************************************************************************
'* [概  要] SendMessage
'* [詳  細] WindowsAPI「 SendMessage 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function SendMessage(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As LongPtr) As LongPtr
    SendMessage = mWinApiUser.WinApi_SendMessage(hWnd, wMsg, wParam, lParam)
End Function

'******************************************************************************
'* [概  要] IsWindowVisible
'* [詳  細] WindowsAPI「 IsWindowVisible 」ラッパーメソッド。
'*          Lib "user32"
'*
'******************************************************************************
Friend Function IsWindowVisible(ByVal hWnd As LongPtr) As Long
    IsWindowVisible = mWinApiUser.WinApi_IsWindowVisible(hWnd)
End Function

