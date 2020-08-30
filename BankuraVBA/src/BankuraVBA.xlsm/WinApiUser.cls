VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI user32クラス
'* [詳  細] WindowAPIのuser32.dllの機能を提供する。
'*
'* [参  考] Win32API_PtrSafe.TXT
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************

Private Declare PtrSafe Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal Flags As Long) As LongPtr
Private Declare PtrSafe Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As LongPtr, ByVal Flags As Long) As LongPtr
Private Declare PtrSafe Function UnloadKeyboardLayout Lib "user32" (ByVal hkl As LongPtr) As Long
Private Declare PtrSafe Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Private Declare PtrSafe Function GetThreadDesktop Lib "user32" (ByVal dwThread As Long) As LongPtr

Private Declare PtrSafe Function GetProcessWindowStation Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetUserObjectSecurity Lib "user32" (ByVal hObj As LongPtr, pSIRequested As Long, pSd As SECURITY_DESCRIPTOR) As Long
Private Declare PtrSafe Function GetUserObjectSecurity Lib "user32" (ByVal hObj As LongPtr, pSIRequested As Long, pSd As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long

Private Declare PtrSafe Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
' Message Function Templates
Private Declare PtrSafe Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare PtrSafe Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
Private Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As LongPtr
Private Declare PtrSafe Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Private Declare PtrSafe Function RegisterHotKey Lib "user32" (ByVal hWnd As LongPtr, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare PtrSafe Function UnregisterHotKey Lib "user32" (ByVal hWnd As LongPtr, ByVal id As Long) As Long
Private Declare PtrSafe Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Declare PtrSafe Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Private Declare PtrSafe Function GetMessagePos Lib "user32" () As Long
Private Declare PtrSafe Function GetMessageTime Lib "user32" () As Long
Private Declare PtrSafe Function GetMessageExtraInfo Lib "user32" () As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As LongPtr) As LongPtr
Private Declare PtrSafe Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function SendMessageCallback Lib "user32" Alias "SendMessageCallbackA" (ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal lpResultCallBack As LongPtr, ByVal dwData As LongPtr) As Long
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function PostThreadMessage Lib "user32" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long

Private Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare PtrSafe Function ReplyMessage Lib "user32" (ByVal lReply As LongPtr) As Long
Private Declare PtrSafe Function WaitMessage Lib "user32" () As Long
Private Declare PtrSafe Function WaitForInputIdle Lib "user32" (ByVal hProcess As LongPtr, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Private Declare PtrSafe Function InSendMessage Lib "user32" () As Long

Private Declare PtrSafe Sub InitCommonControls Lib "COMCTL32" ()
Private Declare PtrSafe Function ImageList_Create Lib "COMCTL32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As LongPtr
Private Declare PtrSafe Function ImageList_AddIcon Lib "COMCTL32" (ByVal himl As LongPtr, ByVal hIcon As LongPtr) As Long
Private Declare PtrSafe Function ImageList_GetIcon Lib "COMCTL32" (ByVal HIMAGELIST As LongPtr, ByVal ImgIndex As Long, hbmMask As Long) As LongPtr
Private Declare PtrSafe Function ImageList_Draw Lib "COMCTL32" (ByVal HIMAGELIST As LongPtr, ByVal ImgIndex As Long, ByVal hDCDest As LongPtr, ByVal xDest As Long, ByVal yDest As Long, ByVal lStyle As Long) As Long

Private Declare PtrSafe Function GetDoubleClickTime Lib "user32" () As Long
Private Declare PtrSafe Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Private Declare PtrSafe Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Private Declare PtrSafe Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As LongPtr) As Long
Private Declare PtrSafe Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As LongPtr, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long

Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function IsMenu Lib "user32" (ByVal hMenu As LongPtr) As Long
Private Declare PtrSafe Function IsChild Lib "user32" (ByVal hwndParent As LongPtr, ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function FlashWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal bInvert As Long) As Long
Private Declare PtrSafe Function ShowOwnedPopups Lib "user32" (ByVal hWnd As LongPtr, ByVal fShow As Long) As Long

Private Declare PtrSafe Function OpenIcon Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function GetWindowPlacement Lib "user32" (ByVal hWnd As LongPtr, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare PtrSafe Function SetWindowPlacement Lib "user32" (ByVal hWnd As LongPtr, lpwndpl As WINDOWPLACEMENT) As Long

Private Declare PtrSafe Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As LongPtr
Private Declare PtrSafe Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As LongPtr, ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As LongPtr
Private Declare PtrSafe Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As LongPtr) As Long

Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function AnyPopup Lib "user32" () As Long
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function EndDialog Lib "user32" (ByVal hDlg As LongPtr, ByVal nResult As LongPtr) As Long
Private Declare PtrSafe Function GetDlgItem Lib "user32" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long) As LongPtr
Private Declare PtrSafe Function SetDlgItemInt Lib "user32" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long, ByVal wValue As Long, ByVal bSigned As Long) As Long
Private Declare PtrSafe Function GetDlgItemInt Lib "user32" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long, ByVal lpTranslated As LongPtr, ByVal bSigned As Long) As Long
Private Declare PtrSafe Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
Private Declare PtrSafe Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function CheckDlgButton Lib "user32" Alias "CheckDLGButtonA" (ByVal hDlg As LongPtr, ByVal nIDButton As Long, ByVal wCheck As Long) As Long
Private Declare PtrSafe Function CheckRadioButton Lib "user32" Alias "CheckRadioButtonA" (ByVal hDlg As LongPtr, ByVal nIDFirstButton As Long, ByVal nIDLastButton As Long, ByVal nIDCheckButton As Long) As Long
Private Declare PtrSafe Function IsDlgButtonChecked Lib "user32" (ByVal hDlg As LongPtr, ByVal nIDButton As Long) As Long
Private Declare PtrSafe Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetNextDlgGroupItem Lib "user32" (ByVal hDlg As LongPtr, ByVal hCtl As LongPtr, ByVal bPrevious As Long) As LongPtr
Private Declare PtrSafe Function GetNextDlgTabItem Lib "user32" (ByVal hDlg As LongPtr, ByVal hCtl As LongPtr, ByVal bPrevious As Long) As LongPtr
Private Declare PtrSafe Function GetDlgCtrlID Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetDialogBaseUnits Lib "user32" () As Long
Private Declare PtrSafe Function DefDlgProc Lib "user32" Alias "DefDlgProcA" (ByVal hDlg As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Private Declare PtrSafe Function CallMsgFilter Lib "user32" Alias "CallMsgFilterA" (lpMsg As msg, ByVal nCode As Long) As Long

' Clipboard Manager Functions
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardOwner Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetClipboardViewer Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardViewer Lib "user32" () As LongPtr
Private Declare PtrSafe Function ChangeClipboardChain Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndNext As LongPtr) As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" Alias "SetClipboardDataA" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare PtrSafe Function CountClipboardFormats Lib "user32" () As Long
Private Declare PtrSafe Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function GetPriorityClipboardFormat Lib "user32" (lpPriorityList As Long, ByVal nCount As Long) As Long
Private Declare PtrSafe Function GetOpenClipboardWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare PtrSafe Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare PtrSafe Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Private Declare PtrSafe Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Private Declare PtrSafe Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Private Declare PtrSafe Function CharUpperBuff Lib "user32" Alias "CharUpperBuffA" (ByVal lpsz As String, ByVal cchLength As Long) As Long
Private Declare PtrSafe Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Private Declare PtrSafe Function CharLowerBuff Lib "user32" Alias "CharLowerBuffA" (ByVal lpsz As String, ByVal cchLength As Long) As Long
Private Declare PtrSafe Function CharNext Lib "user32" Alias "CharNextA" (ByVal lpsz As String) As String
Private Declare PtrSafe Function CharPrev Lib "user32" Alias "CharPrevA" (ByVal lpszStart As String, ByVal lpszCurrent As String) As String

' Language dependent Routines
Private Declare PtrSafe Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Private Declare PtrSafe Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Private Declare PtrSafe Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Private Declare PtrSafe Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long

Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr

' Keyboard Information Routines
Private Declare PtrSafe Function GetKBCodePage Lib "user32" () As Long
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare PtrSafe Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare PtrSafe Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
Private Declare PtrSafe Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare PtrSafe Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
Private Declare PtrSafe Function ToUnicode Lib "user32" (ByVal wVirtKey As Long, ByVal wScanCode As Long, lpKeyState As Byte, ByVal pwszBuff As String, ByVal cchBuff As Long, ByVal wFlags As Long) As Long

Private Declare PtrSafe Function OemKeyScan Lib "user32" (ByVal wOemChar As Long) As Long
Private Declare PtrSafe Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer

Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)

Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As LongPtr)
Private Declare PtrSafe Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

Private Declare PtrSafe Function GetInputState Lib "user32" () As Long
Private Declare PtrSafe Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long
Private Declare PtrSafe Function GetCapture Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetCapture Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long

Private Declare PtrSafe Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As LongPtr, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long

' Windows Functions
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

Private Declare PtrSafe Function IsWindowUnicode Lib "user32" (ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal fEnable As Long) As Long
Private Declare PtrSafe Function IsWindowEnabled Lib "user32" (ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function LoadAccelerators Lib "user32" Alias "LoadAcceleratorsA" (ByVal hInstance As LongPtr, ByVal lpTableName As String) As LongPtr
Private Declare PtrSafe Function CreateAcceleratorTable Lib "user32" Alias "CreateAcceleratorTableA" (lpaccl As ACCEL, ByVal cEntries As Long) As LongPtr
Private Declare PtrSafe Function DestroyAcceleratorTable Lib "user32" (ByVal haccel As LongPtr) As Long
Private Declare PtrSafe Function CopyAcceleratorTable Lib "user32" Alias "CopyAcceleratorTableA" (ByVal hAccelSrc As LongPtr, lpAccelDst As ACCEL, ByVal cAccelEntries As Long) As Long
Private Declare PtrSafe Function TranslateAccelerator Lib "user32" Alias "TranslateAcceleratorA" (ByVal hWnd As LongPtr, ByVal hAccTable As LongPtr, lpMsg As msg) As Long

Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare PtrSafe Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As LongPtr, ByVal lpString As String) As LongPtr
Private Declare PtrSafe Function LoadMenuIndirect Lib "user32" Alias "LoadMenuIndirectA" (ByVal lpMenuTemplate As LongPtr) As LongPtr
Private Declare PtrSafe Function GetMenu Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function SetMenu Lib "user32" (ByVal hWnd As LongPtr, ByVal hMenu As LongPtr) As Long
Private Declare PtrSafe Function HiliteMenuItem Lib "user32" (ByVal hWnd As LongPtr, ByVal hMenu As LongPtr, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long
Private Declare PtrSafe Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As LongPtr, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare PtrSafe Function GetMenuState Lib "user32" (ByVal hMenu As LongPtr, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As LongPtr, ByVal bRevert As Long) As LongPtr
Private Declare PtrSafe Function CreateMenu Lib "user32" () As LongPtr
Private Declare PtrSafe Function CreatePopupMenu Lib "user32" () As LongPtr
Private Declare PtrSafe Function DestroyMenu Lib "user32" (ByVal hMenu As LongPtr) As Long
Private Declare PtrSafe Function CheckMenuItem Lib "user32" (ByVal hMenu As LongPtr, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Private Declare PtrSafe Function EnableMenuItem Lib "user32" (ByVal hMenu As LongPtr, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare PtrSafe Function GetSubMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPos As Long) As LongPtr
Private Declare PtrSafe Function GetMenuItemID Lib "user32" (ByVal hMenu As LongPtr, ByVal nPos As Long) As Long
Private Declare PtrSafe Function GetMenuItemCount Lib "user32" (ByVal hMenu As LongPtr) As Long

Private Declare PtrSafe Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As LongPtr, ByVal lpNewItem As Any) As Long
Private Declare PtrSafe Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As LongPtr, ByVal wFlags As Long, ByVal wIDNewItem As LongPtr, ByVal lpNewItem As Any) As Long
Private Declare PtrSafe Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As LongPtr, ByVal lpString As Any) As Long
Private Declare PtrSafe Function RemoveMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As LongPtr, ByVal hBitmapChecked As LongPtr) As Long
Private Declare PtrSafe Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Private Declare PtrSafe Function TrackPopupMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As LongPtr, lprc As RECT) As Long

Private Declare PtrSafe Function DrawIcon Lib "user32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal hIcon As LongPtr) As Long

Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As LongPtr, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare PtrSafe Function TabbedTextOut Lib "user32" Alias "TabbedTextOutA" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long
Private Declare PtrSafe Function GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentA" (ByVal hdc As LongPtr, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long

Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetActiveWindow Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function WindowFromDC Lib "user32" (ByVal hdc As LongPtr) As LongPtr

Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDCEx Lib "user32" (ByVal hWnd As LongPtr, ByVal hrgnclip As LongPtr, ByVal fdwOptions As Long) As LongPtr

Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long

Private Declare PtrSafe Function BeginPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As LongPtr
Private Declare PtrSafe Function EndPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As Long
Private Declare PtrSafe Function GetUpdateRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare PtrSafe Function GetUpdateRgn Lib "user32" (ByVal hWnd As LongPtr, ByVal hRgn As LongPtr, ByVal fErase As Long) As Long
Private Declare PtrSafe Function ExcludeUpdateRgn Lib "user32" (ByVal hdc As LongPtr, ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare PtrSafe Function ValidateRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function InvalidateRgn Lib "user32" (ByVal hWnd As LongPtr, ByVal hRgn As LongPtr, ByVal bErase As Long) As Long
Private Declare PtrSafe Function ValidateRgn Lib "user32" (ByVal hWnd As LongPtr, ByVal hRgn As LongPtr) As Long

Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hWnd As LongPtr, lprcUpdate As RECT, ByVal hrgnUpdate As LongPtr, ByVal fuRedraw As Long) As Long

Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hwndLock As LongPtr) As Long

Private Declare PtrSafe Function ScrollWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RECT, lpClipRect As RECT) As Long
Private Declare PtrSafe Function ScrollDC Lib "user32" (ByVal hdc As LongPtr, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As LongPtr, lprcUpdate As RECT) As Long
Private Declare PtrSafe Function ScrollWindowEx Lib "user32" (ByVal hWnd As LongPtr, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As LongPtr, lprcUpdate As RECT, ByVal fuScroll As Long) As Long

Private Declare PtrSafe Function SetScrollPos Lib "user32" (ByVal hWnd As LongPtr, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare PtrSafe Function GetScrollPos Lib "user32" (ByVal hWnd As LongPtr, ByVal nBar As Long) As Long
Private Declare PtrSafe Function SetScrollRange Lib "user32" (ByVal hWnd As LongPtr, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Private Declare PtrSafe Function GetScrollRange Lib "user32" (ByVal hWnd As LongPtr, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Private Declare PtrSafe Function ShowScrollBar Lib "user32" (ByVal hWnd As LongPtr, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Declare PtrSafe Function EnableScrollBar Lib "user32" (ByVal hWnd As LongPtr, ByVal wSBflags As Long, ByVal wArrows As Long) As Long

Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal hData As LongPtr) As Long
Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As LongPtr, ByVal lpString As String) As LongPtr
Private Declare PtrSafe Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As LongPtr, ByVal lpString As String) As LongPtr

Private Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function AdjustWindowRect Lib "user32" (lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long) As Long
Private Declare PtrSafe Function AdjustWindowRectEx Lib "user32" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long

Private Declare PtrSafe Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare PtrSafe Function MessageBoxEx Lib "user32" Alias "MessageBoxExA" (ByVal hWnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long
Private Declare PtrSafe Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Private Declare PtrSafe Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare PtrSafe Function GetCursor Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetClipCursor Lib "user32" (lprc As RECT) As Long

Private Declare PtrSafe Function CreateCaret Lib "user32" (ByVal hWnd As LongPtr, ByVal hBitmap As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare PtrSafe Function GetCaretBlinkTime Lib "user32" () As Long
Private Declare PtrSafe Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long
Private Declare PtrSafe Function DestroyCaret Lib "user32" () As Long
Private Declare PtrSafe Function HideCaret Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ShowCaret Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetCaretPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hWnd As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hWnd As LongPtr, lpPoint As POINTAPI) As Long

Private Declare PtrSafe Function MapWindowPoints Lib "user32" (ByVal hwndFrom As LongPtr, ByVal hwndTo As LongPtr, lppt As Any, ByVal cPoints As Long) As Long

Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal point As LongPtr) As LongPtr
Private Declare PtrSafe Function ChildWindowFromPoint Lib "user32" (ByVal hWnd As LongPtr, ByVal point As LongPtr) As LongPtr
Private Declare PtrSafe Function ChildWindowFromPointEx Lib "user32" (ByVal hWnd As LongPtr, ByVal pt As LongPtr, ByVal un As Long) As LongPtr
Private Declare PtrSafe Function DragDetect Lib "user32" (ByVal hWnd As LongPtr, ByVal pt As LongPtr) As Long
Private Declare PtrSafe Function MenuItemFromPoint Lib "user32" (ByVal hWnd As LongPtr, ByVal hMenu As LongPtr, ByVal ptScreen As LongPtr) As Long

Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long

Private Declare PtrSafe Function DrawFocusRect Lib "user32" (ByVal hdc As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hdc As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function FrameRect Lib "user32" (ByVal hdc As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function InvertRect Lib "user32" (ByVal hdc As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare PtrSafe Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare PtrSafe Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare PtrSafe Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare PtrSafe Function UnionRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare PtrSafe Function SubtractRect Lib "user32" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
Private Declare PtrSafe Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare PtrSafe Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
Private Declare PtrSafe Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long

Private Declare PtrSafe Function GetWindowWord Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Integer
Private Declare PtrSafe Function SetWindowWord Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare PtrSafe Function GetClassWord Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetClassWord Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr

Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClassLongPtr Lib "user32" Alias "GetClassLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr

Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetTopWindow Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As LongPtr, ByVal wFlag As Long) As LongPtr

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, lpdwProcessId As Long) As Long

Private Declare PtrSafe Function GetLastActivePopup Lib "user32" (ByVal hwndOwnder As LongPtr) As LongPtr

Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr

' Resource Loading Routines
Private Declare PtrSafe Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As LongPtr, ByVal lpBitmapName As String) As LongPtr
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As String) As LongPtr
Private Declare PtrSafe Function CreateCursor Lib "user32" (ByVal hInstance As LongPtr, ByVal nXhotspot As Long, ByVal nYhotspot As Long, ByVal nWidth As Long, ByVal nHeight As Long, lpANDbitPlane As Any, lpXORbitPlane As Any) As LongPtr
Private Declare PtrSafe Function DestroyCursor Lib "user32" (ByVal hCursor As LongPtr) As Long
Private Declare PtrSafe Function CopyCursor Lib "user32" (ByVal hcur As LongPtr) As LongPtr

Private Declare PtrSafe Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As LongPtr, ByVal lpIconName As String) As LongPtr
Private Declare PtrSafe Function CreateIcon Lib "user32" (ByVal hInstance As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Byte, ByVal nBitsPixel As Byte, lpANDbits As Byte, lpXORbits As Byte) As LongPtr
Private Declare PtrSafe Function DestroyIcon Lib "user32" (ByVal hIcon As LongPtr) As Long
Private Declare PtrSafe Function LookupIconIdFromDirectory Lib "user32" (presbits As Byte, ByVal fIcon As Long) As Long
Private Declare PtrSafe Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As LongPtr
Private Declare PtrSafe Function CopyIcon Lib "user32" (ByVal hIcon As LongPtr) As LongPtr
Private Declare PtrSafe Function GetIconInfo Lib "user32" (ByVal hIcon As LongPtr, piconinfo As ICONINFO) As Long

Private Declare PtrSafe Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As LongPtr, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

' Dialog Manager Routines
Private Declare PtrSafe Function IsDialogMessage Lib "user32" Alias "IsDialogMessageA" (ByVal hDlg As LongPtr, lpMsg As msg) As Long
Private Declare PtrSafe Function MapDialogRect Lib "user32" (ByVal hDlg As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function DlgDirList Lib "user32" Alias "DlgDirListA" (ByVal hDlg As LongPtr, ByVal lpPathSpec As String, ByVal nIDListBox As Long, ByVal nIDStaticPath As Long, ByVal wFileType As Long) As Long

Private Declare PtrSafe Function DlgDirSelectEx Lib "user32" Alias "DlgDirSelectExA" (ByVal hWndDlg As LongPtr, ByVal lpszPath As String, ByVal cbPath As Long, ByVal idListBox As Long) As Long
Private Declare PtrSafe Function DlgDirListComboBox Lib "user32" Alias "DlgDirListComboBoxA" (ByVal hDlg As LongPtr, ByVal lpPathSpec As String, ByVal nIDComboBox As Long, ByVal nIDStaticPath As Long, ByVal wFileType As Long) As Long
Private Declare PtrSafe Function DlgDirSelectComboBoxEx Lib "user32" Alias "DlgDirSelectComboBoxExA" (ByVal hWndDlg As LongPtr, ByVal lpszPath As String, ByVal cbPath As Long, ByVal idComboBox As Long) As Long

Private Declare PtrSafe Function DefFrameProc Lib "user32" Alias "DefFrameProcA" (ByVal hWnd As LongPtr, ByVal hWndMDIClient As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function DefMDIChildProc Lib "user32" Alias "DefMDIChildProcA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Private Declare PtrSafe Function TranslateMDISysAccel Lib "user32" (ByVal hWndClient As LongPtr, lpMsg As msg) As Long

Private Declare PtrSafe Function ArrangeIconicWindows Lib "user32" (ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function CreateMDIWindow Lib "user32" Alias "CreateMDIWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As LongPtr, ByVal hInstance As LongPtr, ByVal lParam As LongPtr) As LongPtr

Private Declare PtrSafe Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As LongPtr, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As LongPtr) As Long

Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Declare PtrSafe Function DdeSetQualityOfService Lib "user32" (ByVal hWndClient As LongPtr, pqosNew As SECURITY_QUALITY_OF_SERVICE, pqosPrev As SECURITY_QUALITY_OF_SERVICE) As Long
Private Declare PtrSafe Function ImpersonateDdeClientWindow Lib "user32" (ByVal hWndClient As LongPtr, ByVal hWndServer As LongPtr) As Long
Private Declare PtrSafe Function PackDDElParam Lib "user32" (ByVal msg As Long, ByVal uiLo As LongPtr, ByVal uiHi As LongPtr) As LongPtr
Private Declare PtrSafe Function UnpackDDElParam Lib "user32" (ByVal msg As Long, ByVal lParam As LongPtr, puiLo As LongPtr, puiHi As LongPtr) As Long
Private Declare PtrSafe Function FreeDDElParam Lib "user32" (ByVal msg As Long, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function ReuseDDElParam Lib "user32" (ByVal lParam As LongPtr, ByVal msgIn As Long, ByVal msgOut As Long, ByVal uiLo As LongPtr, ByVal uiHi As LongPtr) As LongPtr

Private Declare PtrSafe Function DdeUninitialize Lib "user32" (ByVal idInst As Long) As Long

' conversation enumeration functions
Private Declare PtrSafe Function DdeConnectList Lib "user32" (ByVal idInst As Long, ByVal hszService As LongPtr, ByVal hszTopic As LongPtr, ByVal hConvList As LongPtr, pCC As CONVCONTEXT) As LongPtr
Private Declare PtrSafe Function DdeQueryNextServer Lib "user32" (ByVal hConvList As LongPtr, ByVal hConvPrev As LongPtr) As LongPtr
Private Declare PtrSafe Function DdeDisconnectList Lib "user32" (ByVal hConvList As LongPtr) As Long

' conversation control functions
Private Declare PtrSafe Function DdeConnect Lib "user32" (ByVal idInst As Long, ByVal hszService As LongPtr, ByVal hszTopic As LongPtr, pCC As CONVCONTEXT) As LongPtr
Private Declare PtrSafe Function DdeDisconnect Lib "user32" (ByVal hConv As LongPtr) As Long
Private Declare PtrSafe Function DdeReconnect Lib "user32" (ByVal hConv As LongPtr) As LongPtr
Private Declare PtrSafe Function DdeQueryConvInfo Lib "user32" (ByVal hConv As LongPtr, ByVal idTransaction As Long, pConvInfo As CONVINFO) As Long
Private Declare PtrSafe Function DdeSetUserHandle Lib "user32" (ByVal hConv As LongPtr, ByVal id As Long, ByVal hUser As LongPtr) As Long
Private Declare PtrSafe Function DdeAbandonTransaction Lib "user32" (ByVal idInst As Long, ByVal hConv As LongPtr, ByVal idTransaction As Long) As Long

' app server interface functions
Private Declare PtrSafe Function DdePostAdvise Lib "user32" (ByVal idInst As Long, ByVal hszTopic As LongPtr, ByVal hszItem As LongPtr) As Long
Private Declare PtrSafe Function DdeEnableCallback Lib "user32" (ByVal idInst As Long, ByVal hConv As LongPtr, ByVal wCmd As Long) As Long
Private Declare PtrSafe Function DdeImpersonateClient Lib "user32" (ByVal hConv As LongPtr) As Long

Private Declare PtrSafe Function DdeNameService Lib "user32" (ByVal idInst As Long, ByVal hsz1 As Long, ByVal hsz2 As Long, ByVal afCmd As Long) As LongPtr

' app client interface functions
Private Declare PtrSafe Function DdeClientTransaction Lib "user32" (pData As Byte, ByVal cbData As Long, ByVal hConv As LongPtr, ByVal hszItem As LongPtr, ByVal wFmt As Long, ByVal wType As Long, ByVal dwTimeout As Long, pdwResult As Long) As LongPtr

' data transfer functions
Private Declare PtrSafe Function DdeCreateDataHandle Lib "user32" (ByVal idInst As Long, pSrc As Byte, ByVal cb As Long, ByVal cbOff As Long, ByVal hszItem As LongPtr, ByVal wFmt As Long, ByVal afCmd As Long) As LongPtr
Private Declare PtrSafe Function DdeAddData Lib "user32" (ByVal hData As LongPtr, pSrc As Byte, ByVal cb As Long, ByVal cbOff As Long) As LongPtr
Private Declare PtrSafe Function DdeGetData Lib "user32" (ByVal hData As LongPtr, pDst As Byte, ByVal cbMax As Long, ByVal cbOff As Long) As Long
Private Declare PtrSafe Function DdeAccessData Lib "user32" (ByVal hData As LongPtr, pcbDataSize As Long) As LongPtr
Private Declare PtrSafe Function DdeUnaccessData Lib "user32" (ByVal hData As LongPtr) As Long
Private Declare PtrSafe Function DdeFreeDataHandle Lib "user32" (ByVal hData As LongPtr) As Long

Private Declare PtrSafe Function DdeGetLastError Lib "user32" (ByVal idInst As Long) As Long

Private Declare PtrSafe Function DdeCreateStringHandle Lib "user32" Alias "DdeCreateStringHandleA" (ByVal idInst As Long, ByVal psz As String, ByVal iCodePage As Long) As LongPtr

Private Declare PtrSafe Function DdeQueryString Lib "user32" Alias "DdeQueryStringA" (ByVal idInst As Long, ByVal hsz As LongPtr, ByVal psz As String, ByVal cchMax As Long, ByVal iCodePage As Long) As Long

Private Declare PtrSafe Function DdeFreeStringHandle Lib "user32" (ByVal idInst As Long, ByVal hsz As LongPtr) As Long
Private Declare PtrSafe Function DdeKeepStringHandle Lib "user32" (ByVal idInst As Long, ByVal hsz As LongPtr) As Long
Private Declare PtrSafe Function DdeCmpStringHandles Lib "user32" (ByVal hsz1 As LongPtr, ByVal hsz2 As LongPtr) As Long

Private Declare PtrSafe Function DdeInitialize Lib "user32" Alias "DdeInitializeA" (pidInst As Long, ByVal pfnCallback As LongPtr, ByVal afCmd As Long, ByVal ulRes As Long) As Long
Private Declare PtrSafe Function BroadcastSystemMessage Lib "user32" (ByVal dw As Long, pdw As Long, ByVal un As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function CascadeWindows Lib "user32" (ByVal hwndParent As LongPtr, ByVal wHow As Long, lpRect As RECT, ByVal cKids As Long, lpKids As LongPtr) As Integer
Private Declare PtrSafe Function ChangeMenu Lib "user32" Alias "ChangeMenuA" (ByVal hMenu As LongPtr, ByVal cmd As Long, ByVal lpszNewItem As String, ByVal cmdInsert As Long, ByVal Flags As Long) As Long
Private Declare PtrSafe Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As LongPtr, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Private Declare PtrSafe Function CloseDesktop Lib "user32" (ByVal hDesktop As LongPtr) As Long
Private Declare PtrSafe Function CloseWindowStation Lib "user32" (ByVal hWinSta As LongPtr) As Long
Private Declare PtrSafe Function CopyImage Lib "user32" (ByVal handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
Private Declare PtrSafe Function CreateDesktop Lib "user32" Alias "CreateDesktopA" (ByVal lpszDesktop As String, ByVal lpszDevice As String, pDevmode As DEVMODE, ByVal dwFlags As Long, ByVal dwDesiredAccess As Long, lpsa As SECURITY_ATTRIBUTES) As LongPtr
Private Declare PtrSafe Function CreateDialogIndirectParam Lib "user32" Alias "CreateDialogIndirectParamA" (ByVal hInstance As LongPtr, lpTemplate As DLGTEMPLATE, ByVal hwndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateDialogParam Lib "user32" Alias "CreateDialogParamA" (ByVal hInstance As LongPtr, ByVal lpName As String, ByVal hwndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal lParamInit As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateIconFromResource Lib "user32" (presbits As Byte, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long) As LongPtr
Private Declare PtrSafe Function DialogBoxIndirectParam Lib "user32" Alias "DialogBoxIndirectParamA" (ByVal hInstance As LongPtr, hDialogTemplate As DLGTEMPLATE, ByVal hwndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As LongPtr
Private Declare PtrSafe Function DragObject Lib "user32" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal un As Long, ByVal dw As LongPtr, ByVal hCursor As LongPtr) As Long
Private Declare PtrSafe Function DrawAnimatedRects Lib "user32" (ByVal hWnd As LongPtr, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Private Declare PtrSafe Function DrawCaption Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr, pcRect As RECT, ByVal un As Long) As Long
Private Declare PtrSafe Function DrawEdge Lib "user32" (ByVal hdc As LongPtr, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare PtrSafe Function DrawFrameControl Lib "user32" (ByVal hdc As LongPtr, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare PtrSafe Function DrawIconEx Lib "user32" (ByVal hdc As LongPtr, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As LongPtr, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As LongPtr, ByVal diFlags As Long) As Long
Private Declare PtrSafe Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As LongPtr, ByVal hBrush As LongPtr, ByVal lpDrawStateProc As LongPtr, ByVal lParam As LongPtr, ByVal wParam As LongPtr, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare PtrSafe Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Private Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hwndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hWinSta As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As LongPtr, ByVal lpfn As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumPropsEx Lib "user32" Alias "EnumPropsExA" (ByVal hWnd As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumProps Lib "user32" Alias "EnumPropsA" (ByVal hWnd As LongPtr, ByVal lpEnumFunc As LongPtr) As Long
Private Declare PtrSafe Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr

Private Declare PtrSafe Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As LongPtr) As Long
Private Declare PtrSafe Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As LongPtr
Private Declare PtrSafe Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As LongPtr) As Long
Private Declare PtrSafe Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As LongPtr, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Private Declare PtrSafe Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As LongPtr, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare PtrSafe Function GetMenuItemRect Lib "user32" (ByVal hWnd As LongPtr, ByVal hMenu As LongPtr, ByVal uItem As Long, lprcItem As RECT) As Long
Private Declare PtrSafe Function GetScrollInfo Lib "user32" (ByVal hWnd As LongPtr, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare PtrSafe Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function GetUserObjectInformation Lib "user32" Alias "GetUserObjectInformationA" (ByVal hObj As LongPtr, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Private Declare PtrSafe Function GetWindowContextHelpId Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowRgn Lib "user32" (ByVal hWnd As LongPtr, ByVal hRgn As LongPtr) As Long
Private Declare PtrSafe Function GrayString Lib "user32" Alias "GrayStringA" (ByVal hdc As LongPtr, ByVal hBrush As LongPtr, ByVal lpOutputFunc As LongPtr, ByVal lpData As LongPtr, ByVal nCount As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare PtrSafe Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As LongPtr, ByVal un As Long, ByVal bool As Long, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare PtrSafe Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As LongPtr, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
Private Declare PtrSafe Function LookupIconIdFromDirectoryEx Lib "user32" (presbits As Byte, ByVal fIcon As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Declare PtrSafe Function MapVirtualKeyEx Lib "user32" Alias "MapVirtualKeyExA" (ByVal uCode As Long, ByVal uMapType As Long, ByVal dwhkl As LongPtr) As Long
Private Declare PtrSafe Function MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectA" (lpMsgBoxParams As MSGBOXPARAMS) As Long
Private Declare PtrSafe Function OpenDesktop Lib "user32" Alias "OpenDesktopA" (ByVal lpszDesktop As String, ByVal dwFlags As Long, ByVal fInherit As Long, ByVal dwDesiredAccess As Long) As LongPtr
Private Declare PtrSafe Function OpenInputDesktop Lib "user32" (ByVal dwFlags As Long, ByVal fInherit As Long, ByVal dwDesiredAccess As Long) As LongPtr
Private Declare PtrSafe Function OpenWindowStation Lib "user32" Alias "OpenWindowStationA" (ByVal lpszWinSta As String, ByVal fInherit As Long, ByVal dwDesiredAccess As Long) As LongPtr
Private Declare PtrSafe Function PaintDesktop Lib "user32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Private Declare PtrSafe Function SetMenuContextHelpId Lib "user32" (ByVal hMenu As LongPtr, ByVal dw As Long) As Long
Private Declare PtrSafe Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As LongPtr, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare PtrSafe Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As LongPtr, ByVal un As Long, ByVal bool As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare PtrSafe Function SetMessageExtraInfo Lib "user32" (ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function SetMessageQueue Lib "user32" (ByVal cMessagesMax As Long) As Long
Private Declare PtrSafe Function SetProcessWindowStation Lib "user32" (ByVal hWinSta As LongPtr) As Long
Private Declare PtrSafe Function SetScrollInfo Lib "user32" (ByVal hWnd As LongPtr, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Long) As Long
Private Declare PtrSafe Function SetSystemCursor Lib "user32" (ByVal hcur As LongPtr, ByVal id As Long) As Long
Private Declare PtrSafe Function SetThreadDesktop Lib "user32" (ByVal hDesktop As LongPtr) As Long
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function SetUserObjectInformation Lib "user32" Alias "SetUserObjectInformationA" (ByVal hObj As LongPtr, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long) As Long
Private Declare PtrSafe Function SetWindowContextHelpId Lib "user32" (ByVal hWnd As LongPtr, ByVal dw As Long) As Long
Private Declare PtrSafe Function SetWindowRgn Lib "user32" (ByVal hWnd As LongPtr, ByVal hRgn As LongPtr, ByVal bRedraw As Long) As Long
Private Declare PtrSafe Function SetWindowsHook Lib "user32" Alias "SetWindowsHookA" (ByVal nFilterType As Long, ByVal pfnFilterProc As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function ShowWindowAsync Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SwitchDesktop Lib "user32" (ByVal hDesktop As LongPtr) As Long
Private Declare PtrSafe Function TileWindows Lib "user32" (ByVal hwndParent As LongPtr, ByVal wHow As Long, lpRect As RECT, ByVal cKids As Long, lpKids As LongPtr) As Integer
Private Declare PtrSafe Function ToAsciiEx Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpKeyState As Byte, lpChar As Integer, ByVal uFlags As Long, ByVal dwhkl As LongPtr) As Long
Private Declare PtrSafe Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As LongPtr, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hWnd As LongPtr, lpTPMParams As TPMPARAMS) As Long
Private Declare PtrSafe Function UnhookWindowsHook Lib "user32" (ByVal nCode As Long, ByVal pfnFilterProc As LongPtr) As Long
Private Declare PtrSafe Function VkKeyScanEx Lib "user32" Alias "VkKeyScanExA" (ByVal ch As Byte, ByVal dwhkl As LongPtr) As Integer

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As LongPtr)
' Copies a POINTAPI into a LongPtr.  For an API requiring a ByVal POINTAPI parameter,
' this LongPtr can be passed in instead.  Example API's include WindowFromPoint,
' ChildWindowFromPoint, ChildWindowFromPointEx, DragDetect, and MenuItemFromPoint.
Friend Function PointToLongPtr(point As POINTAPI) As LongPtr
    Dim ll As LongPtr
    Dim cbLongPtr As LongPtr
    
    cbLongPtr = LenB(ll)
    
    ' make sure the contents will fit
    If LenB(point) = cbLongPtr Then
        CopyMemory ll, point, cbLongPtr
    End If
    
    PointToLongPtr = ll
End Function

'******************************************************************************
'* [概  要] WinApi_GetInputState
'* [詳  細] WindowsAPI「 GetInputState 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_GetInputState() As Long
    WinApi_GetInputState = GetInputState()
End Function

'******************************************************************************
'* [概  要] WinApi_SetForegroundWindow
'* [詳  細] WindowsAPI「 SetForegroundWindow 」ラッパーメソッド。
'*          ウィンドウを強制的に最前面に表示させるAPI。
'*
'******************************************************************************
Friend Function WinApi_SetForegroundWindow(ByVal hWnd As LongPtr) As Long
    WinApi_SetForegroundWindow = SetForegroundWindow(hWnd)
End Function

'******************************************************************************
'* [概  要] WinApi_IsIconic
'* [詳  細] WindowsAPI「 IsIconic 」ラッパーメソッド。
'*          ウィンドウが最小化されているか調べるAPI。
'*
'******************************************************************************
Friend Function WinApi_IsIconic(ByVal hWnd As LongPtr) As Long
    WinApi_IsIconic = IsIconic(hWnd)
End Function
    
'******************************************************************************
'* [概  要] WinApi_ShowWindowAsync
'* [詳  細] WindowsAPI「 ShowWindowAsync 」ラッパーメソッド。
'*          ウィンドウを元の大きさに戻すAPI。
'*
'******************************************************************************
Friend Function WinApi_ShowWindowAsync(ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    WinApi_ShowWindowAsync = ShowWindowAsync(hWnd, nCmdShow)
End Function

'******************************************************************************
'* [概  要] WinApi_PeekMessage
'* [詳  細] WindowsAPI「 PeekMessage 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_PeekMessage(lpMsg As msg, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
    WinApi_PeekMessage = PeekMessage(lpMsg, hWnd, wMsgFilterMin, wMsgFilterMax, wRemoveMsg)
End Function

'******************************************************************************
'* [概  要] WinApi_TranslateMessage
'* [詳  細] WindowsAPI「 TranslateMessage 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_TranslateMessage(lpMsg As msg) As Long
    WinApi_TranslateMessage = TranslateMessage(lpMsg)
End Function

'******************************************************************************
'* [概  要] WinApi_DispatchMessage
'* [詳  細] WindowsAPI「 DispatchMessage 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_DispatchMessage(lpMsg As msg) As LongPtr
    WinApi_DispatchMessage = DispatchMessage(lpMsg)
End Function

'******************************************************************************
'* [概  要] WinApi_keybd_event
'* [詳  細] WindowsAPI「 keybd_event 」ラッパーメソッド。
'*
'******************************************************************************
Friend Sub WinApi_keybd_event(ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
    Call keybd_event(bVk, bScan, dwFlags, dwExtraInfo)
End Sub

'******************************************************************************
'* [概  要] WinApi_GetKeyboardState
'* [詳  細] WindowsAPI 「GetKeyboardState 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_GetKeyboardState(pbKeyState As Byte) As Long
    WinApi_GetKeyboardState = GetKeyboardState(pbKeyState)
End Function

'******************************************************************************
'* [概  要] WinApi_SetCursorPos
'* [詳  細] WindowsAPI「 SetCursorPos 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_SetCursorPos(ByVal x As Long, ByVal y As Long) As Long
    WinApi_SetCursorPos = SetCursorPos(x, y)
End Function

'******************************************************************************
'* [概  要] WinApi_GetCursorPos
'* [詳  細] WindowsAPI「 GetCursorPos 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_GetCursorPos(lpPoint As POINTAPI) As Long
    WinApi_GetCursorPos = GetCursorPos(lpPoint)
End Function

'******************************************************************************
'* [概  要] WinApi_EnumChildWindows
'* [詳  細] WindowsAPI「 EnumChildWindows 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_EnumChildWindows(ByVal hwndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    WinApi_EnumChildWindows = EnumChildWindows(hwndParent, lpEnumFunc, ByVal lParam)
End Function

'******************************************************************************
'* [概  要] WinApi_EnumWindows
'* [詳  細] WindowsAPI「 EnumWindows 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_EnumWindows(ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    WinApi_EnumWindows = EnumWindows(lpEnumFunc, ByVal lParam)
End Function

'******************************************************************************
'* [概  要] WinApi_GetClassName
'* [詳  細] WindowsAPI「 GetClassName 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_GetClassName(ByVal hWnd As LongPtr, ByRef lpClassName As String, ByVal nMaxCount As Long) As Long
    WinApi_GetClassName = GetClassName(hWnd, lpClassName, nMaxCount)
End Function

'******************************************************************************
'* [概  要] WinApi_GetDesktopWindow
'* [詳  細] WindowsAPI「 GetDesktopWindow 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_GetDesktopWindow() As LongPtr
    WinApi_GetDesktopWindow = GetDesktopWindow()
End Function

'******************************************************************************
'* [概  要] WinApi_GetWindow
'* [詳  細] WindowsAPI「 GetWindow 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_GetWindow(ByVal hWnd As LongPtr, ByVal wCmd As Long) As LongPtr
    WinApi_GetWindow = GetWindow(hWnd, wCmd)
End Function

'******************************************************************************
'* [概  要] WinApi_GetWindowText
'* [詳  細] WindowsAPI「 GetWindowText 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_GetWindowText(ByVal hWnd As LongPtr, ByRef lpString As String, ByVal cch As Long) As Long
    WinApi_GetWindowText = GetWindowText(hWnd, lpString, cch)
End Function

'******************************************************************************
'* [概  要] WinApi_GetWindowThreadProcessId
'* [詳  細] WindowsAPI「 GetWindowThreadProcessId 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_GetWindowThreadProcessId(ByVal hWnd As LongPtr, lpdwProcessId As Long) As Long
    WinApi_GetWindowThreadProcessId = GetWindowThreadProcessId(hWnd, lpdwProcessId)
End Function

'******************************************************************************
'* [概  要] WinApi_IsWindow
'* [詳  細] WindowsAPI「 IsWindow 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_IsWindow(ByVal hWnd As LongPtr) As LongPtr
    WinApi_IsWindow = IsWindow(hWnd)
End Function

'******************************************************************************
'* [概  要] WinApi_SendMessage
'* [詳  細] WindowsAPI「 SendMessage 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_SendMessage(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As LongPtr) As LongPtr
    WinApi_SendMessage = SendMessage(hWnd, wMsg, wParam, ByVal lParam)
End Function
'******************************************************************************
'* [概  要] WinApi_IsWindowVisible
'* [詳  細] WindowsAPI「 IsWindowVisible 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_IsWindowVisible(ByVal hWnd As LongPtr) As Long
    WinApi_IsWindowVisible = IsWindowVisible(hWnd)
End Function
