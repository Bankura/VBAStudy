Attribute VB_Name = "WScriptEx"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WScript���b�v�E�g�����W���[��
'* [��  ��] WScript��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�AWScript���g�p����
'*          ���[�e�B���e�B��񋟂���B
'*          ���b�v����WScript���C�u�����͈ȉ��̂��̂Ƃ���B
'*              [name] Windows Script Host Object Model
'*              [library name] IWshRuntimeLibrary
'*              [dll] C:\Windows\System32\wshom.ocx
'* [��  �l] Scripting�Ƌ��ʂ���N���X�͏��O����B
'* [�Q  �l]
'*  <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364455(v=msdn.10)?redirectedfrom=MSDN>
'*  <https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2003/cc738350(v=ws.10)?redirectedfrom=MSDN>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* WshExecStatus
'*
'*-----------------------------------------------------------------------------
Public Enum WshExecStatus
    WshFailed = 2
    WshFinished = 1
    WshRunning = 0
End Enum

'*-----------------------------------------------------------------------------
'* WshWindowStyle
'*
'*-----------------------------------------------------------------------------
Public Enum WshWindowStyle
    WshHide = 0
    WshMaximizedFocus = 3
    WshMinimizedFocus = 2
    WshMinimizedNoFocus = 6
    WshNormalFocus = 1
    WshNormalNoFocus = 4
End Enum

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************
Sub Echo(text As String)
    Debug.Print text
End Sub

Sub Sleep(time As Long)
    Application.Wait Now() + time / 86400000
End Sub


