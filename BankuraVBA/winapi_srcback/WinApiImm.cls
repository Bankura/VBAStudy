VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiImm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI imm クラス
'* [詳  細] WindowAPIのimm32.dllの機能を提供する。
'*
'* [参  考] Win32API_PtrSafe.TXT
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
' IMM.H - Input Method Manager definitions
Private Declare PtrSafe Function ImmInstallIME Lib "imm32.dll" Alias "ImmInstallIMEA" (ByVal lpszIMEFileName As String, ByVal lpszLayoutText As String) As LongPtr
Private Declare PtrSafe Function ImmGetDefaultIMEWnd Lib "imm32.dll" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As LongPtr, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Private Declare PtrSafe Function ImmGetIMEFileName Lib "imm32.dll" Alias "ImmGetIMEFileNameA" (ByVal hkl As LongPtr, ByVal lpStr As String, ByVal uBufLen As Long) As Long
Private Declare PtrSafe Function ImmGetProperty Lib "imm32.dll" (ByVal hkl As LongPtr, ByVal dw As Long) As Long
Private Declare PtrSafe Function ImmIsIME Lib "imm32.dll" (ByVal hkl As LongPtr) As Long
Private Declare PtrSafe Function ImmSimulateHotKey Lib "imm32.dll" (ByVal hWnd As LongPtr, ByVal dw As Long) As Long
Private Declare PtrSafe Function ImmCreateContext Lib "imm32.dll" () As LongPtr
Private Declare PtrSafe Function ImmDestroyContext Lib "imm32.dll" (ByVal himc As LongPtr) As Long
Private Declare PtrSafe Function ImmGetContext Lib "imm32.dll" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ImmReleaseContext Lib "imm32.dll" (ByVal hWnd As LongPtr, ByVal himc As LongPtr) As Long
Private Declare PtrSafe Function ImmAssociateContext Lib "imm32.dll" (ByVal hWnd As LongPtr, ByVal himc As LongPtr) As LongPtr
Private Declare PtrSafe Function ImmGetCompositionString Lib "imm32.dll" Alias "ImmGetCompositionStringA" (ByVal himc As LongPtr, ByVal dw As Long, lpv As Any, ByVal dw2 As Long) As Long
Private Declare PtrSafe Function ImmSetCompositionString Lib "imm32.dll" Alias "ImmSetCompositionStringA" (ByVal himc As LongPtr, ByVal dwIndex As Long, lpComp As Any, ByVal dw As Long, lpRead As Any, ByVal dw2 As Long) As Long
Private Declare PtrSafe Function ImmGetCandidateListCount Lib "imm32.dll" Alias "ImmGetCandidateListCountA" (ByVal himc As LongPtr, lpdwListCount As Long) As Long
Private Declare PtrSafe Function ImmGetCandidateList Lib "imm32.dll" Alias "ImmGetCandidateListA" (ByVal himc As LongPtr, ByVal deIndex As Long, lpCandidateList As CANDIDATELIST, ByVal dwBufLen As Long) As Long
Private Declare PtrSafe Function ImmGetGuideLine Lib "imm32.dll" Alias " ImmGetGuideLineA" (ByVal himc As LongPtr, ByVal dwIndex As Long, ByVal lpStr As String, ByVal dwBufLen As Long) As Long
Private Declare PtrSafe Function ImmGetConversionStatus Lib "imm32.dll" (ByVal himc As LongPtr, lpdw As Long, lpdw2 As Long) As Long
Private Declare PtrSafe Function ImmSetConversionStatus Lib "imm32.dll" (ByVal himc As LongPtr, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Private Declare PtrSafe Function ImmGetOpenStatus Lib "imm32.dll" (ByVal himc As LongPtr) As Long
Private Declare PtrSafe Function ImmSetOpenStatus Lib "imm32.dll" (ByVal himc As LongPtr, ByVal b As Long) As Long
Private Declare PtrSafe Function ImmGetCompositionFont Lib "imm32.dll" Alias "ImmGetCompositionFontA" (ByVal himc As LongPtr, lpLogFont As LOGFONT) As Long
Private Declare PtrSafe Function ImmSetCompositionFont Lib "imm32.dll" Alias "ImmSetCompositionFontA" (ByVal himc As LongPtr, lpLogFont As LOGFONT) As Long
Private Declare PtrSafe Function ImmConfigureIME Lib "imm32.dll" (ByVal hkl As LongPtr, ByVal hWnd As LongPtr, ByVal dw As Long, lpData As Any) As Long
Private Declare PtrSafe Function ImmEscape Lib "imm32.dll" Alias "ImmEscapeA" (ByVal hkl As LongPtr, ByVal himc As LongPtr, ByVal un As Long, lpv As Any) As LongPtr
Private Declare PtrSafe Function ImmGetConversionList Lib "imm32.dll" Alias "ImmGetConversionListA" (ByVal hkl As LongPtr, ByVal himc As LongPtr, ByVal lpsz As String, lpCandidateList As CANDIDATELIST, ByVal dwBufLen As Long, ByVal uFlag As Long) As Long
Private Declare PtrSafe Function ImmNotifyIME Lib "imm32.dll" (ByVal himc As LongPtr, ByVal dwAction As Long, ByVal dwIndex As Long, ByVal dwValue As Long) As Long
Private Declare PtrSafe Function ImmGetStatusWindowPos Lib "imm32.dll" (ByVal himc As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function ImmSetStatusWindowPos Lib "imm32.dll" (ByVal himc As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function ImmGetCompositionWindow Lib "imm32.dll" (ByVal himc As LongPtr, lpCompositionForm As COMPOSITIONFORM) As Long
Private Declare PtrSafe Function ImmSetCompositionWindow Lib "imm32.dll" (ByVal himc As LongPtr, lpCompositionForm As COMPOSITIONFORM) As Long
Private Declare PtrSafe Function ImmGetCandidateWindow Lib "imm32.dll" (ByVal himc As LongPtr, ByVal dw As Long, lpCandidateForm As CANDIDATEFORM) As Long
Private Declare PtrSafe Function ImmSetCandidateWindow Lib "imm32.dll" (ByVal himc As LongPtr, lpCandidateForm As CANDIDATEFORM) As Long
Private Declare PtrSafe Function ImmIsUIMessage Lib "imm32.dll" Alias "ImmIsUIMessageA" (ByVal hWnd As LongPtr, ByVal un As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function ImmGetVirtualKey Lib "imm32.dll" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ImmRegisterWord Lib "imm32.dll" Alias "ImmRegisterWordA" (ByVal hkl As LongPtr, ByVal lpszReading As String, ByVal dw As Long, ByVal lpszRegister As String) As Long
Private Declare PtrSafe Function ImmUnregisterWord Lib "imm32.dll" Alias "ImmUnregisterWordA" (ByVal hkl As LongPtr, ByVal lpszReading As String, ByVal dw As Long, ByVal lpszUnregister As String) As Long
Private Declare PtrSafe Function ImmGetRegisterWordStyle Lib "imm32.dll" Alias " ImmGetRegisterWordStyleA" (ByVal hkl As LongPtr, ByVal nItem As Long, lpStyleBuf As STYLEBUF) As Long
Private Declare PtrSafe Function ImmEnumRegisterWord Lib "imm32.dll" Alias "ImmEnumRegisterWordA" (ByVal hkl As LongPtr, ByVal RegisterWordEnumProc As LongPtr, ByVal lpszReading As String, ByVal dw As Long, ByVal lpszRegister As String, lpv As Any) As Long


