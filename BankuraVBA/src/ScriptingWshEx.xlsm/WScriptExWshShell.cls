VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WScriptExWshShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WScript.WshShell�N���X�̃��b�v�E�g���N���X
'* [��  ��] WScript.WshShell�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          WScript.WshShell�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' WScript.WshShell�N���X�ɂ��āB
' WindowsScriptHost�̃V�F���I�u�W�F�N�g�BDos�R�}���h�̎��s�A�A�v���P�[�V�����̋N���ȂǂɎg�p����B
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364436(v=msdn.10)>
'
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As IWshRuntimeLibrary.WshShell
Private mOrigin As Object
    
'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************
'*-----------------------------------------------------------------------------
'* �I���W�i��Object �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get Origin() As Object
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Object)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* CurrentDirectory �v���p�e�B
'*
'* [��  ��]
'* �E���݃A�N�e�B�u�ɂȂ��Ă���f�B���N�g�����擾�܂��͕ύX���܂��B
'*-----------------------------------------------------------------------------
Property Get CurrentDirectory() As String
    CurrentDirectory = mOrigin.CurrentDirectory
End Property
Property Let CurrentDirectory(arg As String)
    mOrigin.CurrentDirectory = arg
End Property

'*-----------------------------------------------------------------------------
'* Environment �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EWshEnvironment �I�u�W�F�N�g (���ϐ��̃R���N�V����) ��Ԃ��܂��B
'* �E<https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364502%28v%3dmsdn.10%29>
'*
'* @param Type_ �ȗ��\�B���ϐ��̎�ނ��w�肵�܂��B
'*              ���ϐ��̑��ݏꏊ�������l�Ƃ��� System�AUser�AVolatile�A
'*              �܂��� Process �̂����ꂩ���g�p�\�B�w�肵�Ȃ��ꍇ��System�B
'*
'*-----------------------------------------------------------------------------
Property Get Environment(Optional Type_) As WScriptExWshEnvironment
    Dim wrap As WScriptExWshEnvironment: Set wrap = New WScriptExWshEnvironment
    If IsMissing(Type_) Then
        Set wrap.Origin = mOrigin.Environment
    Else
        Set wrap.Origin = mOrigin.Environment(Type_)
    End If
    Set Environment = wrap
End Property

'*-----------------------------------------------------------------------------
'* SpecialFolders �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �ESpecialFolders �I�u�W�F�N�g (����t�H���_�̃R���N�V����) ��Ԃ��܂��B
'* �E<https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364490%28v%3dmsdn.10%29>
'*-----------------------------------------------------------------------------
Property Get SpecialFolders() As WScriptExWshCollection
    Dim wrap As WScriptExWshCollection: Set wrap = New WScriptExWshCollection
    Set wrap.Origin = mOrigin.SpecialFolders
    Set SpecialFolders = wrap
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("WScript.Shell")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] AppActivate ���\�b�h
'* [��  ��] AppActivate �̃��b�p�[���\�b�h�B
'*          �A�v���P�[�V���� �E�B���h�E���A�N�e�B�u�ɂ��܂��B
'* [�Q  �l] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364396%28v%3dmsdn.10%29>
'*
'* @param App �A�N�e�B�u�ɂ���A�v���P�[�V�������w�肵�܂��B
'*            ���̈����ɂ́A�A�v���P�[�V�����̃^�C�g�� �o�[�ɕ\�������
'*            �^�C�g�������񂩁A�A�v���P�[�V�����̃v���Z�X ID ���w��ł��܂��B
'* @param Wait �ȗ��\�B
'* @return Boolean
'*
'******************************************************************************
Public Function AppActivate(App, Optional Wait) As Boolean
    If IsMissing(Wait) Then
        AppActivate = mOrigin.AppActivate(App)
    Else
        AppActivate = mOrigin.AppActivate(App, Wait)
    End If
End Function

'******************************************************************************
'* [�T  �v] CreateShortcut ���\�b�h
'* [��  ��] CreateShortcut �̃��b�p�[���\�b�h�B
'*          �V���[�g�J�b�g�܂��� URL �V���[�g�J�b�g�ւ̃I�u�W�F�N�g�Q�Ƃ��쐬���܂��B
'* [�Q  �l] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364400%28v%3dmsdn.10%29>
'*
'* @param PathLink �쐬����V���[�g�J�b�g�̃p�X�ƃt�@�C����
'* @return WScriptExWshShortcut�܂���WScriptExWshURLShortcut�I�u�W�F�N�g
'*
'******************************************************************************
Public Function CreateShortcut(PathLink As String) As Object
    Dim obj As Object
    Set obj = mOrigin.CreateShortcut(PathLink)
    If TypeName(obj) = "WshShortcut" Then
        Dim wrap As WScriptExWshShortcut: Set wrap = New WScriptExWshShortcut
        Set wrap.Origin = obj
        Set CreateShortcut = wrap
    ElseIf TypeName(obj) = "WshURLShortcut" Then
        Dim wrap2 As WScriptExWshURLShortcut: Set wrap2 = New WScriptExWshURLShortcut
        Set wrap2.Origin = obj
        Set CreateShortcut = wrap2
    Else
        Set CreateShortcut = obj
    End If
End Function

'******************************************************************************
'* [�T  �v] Exec ���\�b�h
'* [��  ��] Exec �̃��b�p�[���\�b�h�B
'*          �q�R�}���h�V�F���ŃA�v���P�[�V���������s���܂��B
'*          �A�v���P�[�V�������� StdIn/StdOut/StdErr �X�g���[���ɃA�N�Z�X
'*          �ł��܂��B
'* [�Q  �l] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364356%28v%3dmsdn.10%29>
'*
'* @param Command �X�N���v�g�̎��s�Ɏg�p����R�}���h ���C��������������l�ł��B
'*                �R�}���h �v�����v�g������͂���ꍇ�ƑS�������R�}���h ���C��
'*                ���w�肵�܂��B
'* @return WScriptExWshExec
'*
'******************************************************************************
Public Function Exec(Command As String) As WScriptExWshExec
    Dim wrap As WScriptExWshExec: Set wrap = New WScriptExWshExec
    Set wrap.Origin = mOrigin.Exec(Command)
    Set Exec = wrap
End Function

'******************************************************************************
'* [�T  �v] ExpandEnvironmentStrings ���\�b�h
'* [��  ��] ExpandEnvironmentStrings �̃��b�p�[���\�b�h�B
'*          ���ϐ���W�J�����l��Ԃ��܂��B
'*          ExpandEnvironmentStrings ���\�b�h�́APROCESS ���̈�ɒ�`���ꂽ
'*          ���ϐ��݂̂�W�J���܂��B���ϐ����� "%" �ň͂ޕK�v������܂��B
'*          �܂��A���ϐ����̑啶���Ə������͋�ʂ���܂���B
'* [�Q  �l] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364405%28v%3dmsdn.10%29>
'*
'* @param Src �W�J������ϐ��̖��O
'* @return ���ϐ���W�J�����l
'*
'******************************************************************************
Public Function ExpandEnvironmentStrings(Src As String) As String
    ExpandEnvironmentStrings = mOrigin.ExpandEnvironmentStrings(Src)
End Function

'******************************************************************************
'* [�T  �v] LogEvent ���\�b�h
'* [��  ��] LogEvent �̃��b�p�[���\�b�h�B
'*          �C�x���g �G���g�������O �t�@�C���ɒǉ����܂��B
'* [�Q  �l] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364408%28v%3dmsdn.10%29>
'*
'* @param Type_ �C�x���g�̎�ނ�\�����鐔�l
'* @param Message ���O �G���g���̃e�L�X�g
'* @param Target �ȗ��\�B�C�x���g ���O���i�[����R���s���[�^�V�X�e���̖��O
'*               ������������l�B����l�̓��[�J���̃R���s���[�^ �V�X�e���B
'* @return �C�x���g������ɋL�^���ꂽ�ꍇ�� TRUE�A����ȊO�̏ꍇ�� FALSE
'*
'******************************************************************************
Public Function LogEvent(Type_, Message As String, Optional Target As String) As Boolean
    LogEvent = mOrigin.LogEvent(Type_, Message, Target)
End Function

'******************************************************************************
'* [�T  �v] Popup ���\�b�h
'* [��  ��] Popup �̃��b�p�[���\�b�h�B
'*          �|�b�v�A�b�v ���b�Z�[�W �{�b�N�X�Ƀe�L�X�g��\�����܂��B
'* [�Q  �l] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364428%28v%3dmsdn.10%29>
'*
'* @param Text �|�b�v�A�b�v �E�B���h�E�ɕ\������e�L�X�g
'* @param SecondsToWait �ȗ��\�B�|�b�v�A�b�v �E�B���h�E�����܂őҋ@����
'*                      �b���B�[�� (����l) ���w�肵���ꍇ�A�|�b�v�A�b�v
'*                      ���b�Z�[�W�{�b�N�X�̓��[�U�[���E�B���h�E�����܂�
'*                      �����ƕ\������܂��B
'* @param Title �ȗ��\�B�|�b�v�A�b�v�E�B���h�E�̃^�C�g���ɕ\������e�L�X�g�B
'*              �ȗ�����ƁA"Windows Script Host" �ƂȂ�܂��B
'* @param Type_ �ȗ��\�B�|�b�v�A�b�v���b�Z�[�W �{�b�N�X�ɕ\������{�^����
'*              �A�C�R���̎�ނ��������l�B�I�������ނɂ���ă��b�Z�[�W
'*              �{�b�N�X�̎g���������܂�܂��B
'* @return ���b�Z�[�W �{�b�N�X�I�����ɃN���b�N����{�^���̔ԍ������������l
'*
'******************************************************************************
Public Function Popup(text As String, Optional SecondsToWait, Optional Title, Optional Type_) As Long
    If IsMissing(SecondsToWait) Then
        If IsMissing(Title) Then
             If IsMissing(Type_) Then
                Popup = mOrigin.Popup(text)
            Else
                Popup = mOrigin.Popup(text, , , Type_)
            End If
        ElseIf IsMissing(Type_) Then
            Popup = mOrigin.Popup(text, , Title)
        Else
            Popup = mOrigin.Popup(text, , Title, Type_)
        End If
    Else
        If IsMissing(Title) Then
             If IsMissing(Type_) Then
                Popup = mOrigin.Popup(text, SecondsToWait)
            Else
                Popup = mOrigin.Popup(text, SecondsToWait, , Type_)
            End If
        ElseIf IsMissing(Type_) Then
            Popup = mOrigin.Popup(text, SecondsToWait, Title)
        Else
            Popup = mOrigin.Popup(text, SecondsToWait, Title, Type_)
        End If
    End If
End Function

'******************************************************************************
'* [�T  �v] RegDelete ���\�b�h
'* [��  ��] RegDelete �̃��b�p�[���\�b�h�B
'*          ���W�X�g������w�肳�ꂽ�L�[�܂��͒l���폜���܂��B
'* [�Q  �l] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364415%28v%3dmsdn.10%29>
'*
'* @param Name ���W�X�g�����ō폜����L�[�܂��͒l�̖��O������������l
'*
'******************************************************************************
Public Sub RegDelete(Name As String)
    Call mOrigin.RegDelete(Name)
End Sub

'******************************************************************************
'* [�T  �v] RegRead ���\�b�h
'* [��  ��] RegRead �̃��b�p�[���\�b�h�B
'*          ���W�X�g�����̃L�[���܂��͒l���̒l��Ԃ��܂��B
'* [�Q  �l] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364416%28v%3dmsdn.10%29>
'*
'* @param Name �ǂݎ��L�[�܂��͒l�̖��O
'* @return ���� 5 ��ނ̒l��Ԃ��܂��B
'*           �^            ����                        �`��
'*           -----------------------------------------------------------
'*           REG_SZ        ������ł��B                ������
'*           REG_DWORD     ���l�ł��B                  ����
'*           REG_BINARY    2�i���̒l�ł��B             VBArray (����)
'*           REG_EXPAND_SZ �W�J�\�ȕ�����ł��
'*                         (�� : "%windir%\\calc.exe") ������
'*           REG_MULTI_SZ  ������̔z��ł��B          VBArray (������)
'*
'******************************************************************************
Public Function RegRead(Name As String)
    RegRead = mOrigin.RegRead(Name)
End Function

'******************************************************************************
'* [�T  �v] RegWrite ���\�b�h
'* [��  ��] RegWrite �̃��b�p�[���\�b�h�B
'*          �V�����L�[�̍쐬�A�V�����l���̊����L�[�ւ̒ǉ� (����ђl�̐ݒ�)�A
'*          �����̒l���̒l�ύX�Ȃǂ��s���܂��B
'* [�Q  �l] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364417(v=msdn.10)>
'*
'* @param Name �쐬�A�ǉ��A�ύX����L�[���A�l���A�܂��͒l������������l
'* @param Value �쐬����L�[�̖��O�A�����̃L�[�ɒǉ�����l�̖��O�A�܂��͊���
'*              �̒l���ɐݒ肷��l
'* @param Type_ �ȗ��\�B���W�X�g���ɕۑ�����l�̃f�[�^�^�B
'*
'******************************************************************************
Public Sub RegWrite(Name As String, Value, Optional Type_)
    If IsMissing(Type_) Then
        Call mOrigin.RegWrite(Name, Value)
    Else
        Call mOrigin.RegWrite(Name, Value, Type_)
    End If
End Sub

'******************************************************************************
'* [�T  �v] Run ���\�b�h
'* [��  ��] Run �̃��b�p�[���\�b�h�B
'*          �V�����v���Z�X���Ńv���O���������s���܂��B
'* [�Q  �l] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364421(v=msdn.10)>
'*
'* @param Command ���s����R�}���h ���C��������������l�B
'*                ���̈����ɂ́A���s�\�t�@�C���ɓn���ׂ��p�����[�^�����ׂ�
'*                �܂߂�K�v������܂��B
'* @param WindowStyle �ȗ��\�B�v���O�����̃E�B���h�E�̊O�ς����������l�B
'* @param WaitOnReturn �ȗ��\�B�X�N���v�g���̎��̃X�e�[�g�����g�ɐi�܂���
'*                     �v���O�����̎��s���I������܂ŃX�N���v�g��ҋ@�����邩
'*                     �ǂ����������u�[���l�B
'*                     WaitOnReturn �� TRUE ���w�肷��ƁA�v���O�����̎��s��
'*                     �I������܂ŃX�N���v�g�̎��s�͒��f����ARun ���\�b�h��
'*                     �A�v���P�[�V��������Ԃ����C�ӂ̃G���[�R�[�h��Ԃ�
'*                     �܂��BFALSE ���w�肷��ƁA�v���O�������J�n����� Run
'*                     ���\�b�h�͑����ɕ��A���Ď����I�� 0 ��Ԃ��܂�
'* @return ���s���ʂ����������l�i�G���[�R�[�h���j
'*
'******************************************************************************
Public Function Run(Command As String, Optional WindowStyle, Optional WaitOnReturn) As Long
    If IsMissing(WindowStyle) Then
        If IsMissing(WaitOnReturn) Then
            Run = mOrigin.Run(Command)
        Else
            Run = mOrigin.Run(Command, , WaitOnReturn)
        End If
    ElseIf IsMissing(WaitOnReturn) Then
        Run = mOrigin.Run(Command, WindowStyle)
    Else
        Run = mOrigin.Run(Command, WindowStyle, WaitOnReturn)
    End If
End Function

'******************************************************************************
'* [�T  �v] SendKeys ���\�b�h
'* [��  ��] SendKeys �̃��b�p�[���\�b�h�B
'*          �L�[�{�[�h������͂����Ƃ��̂悤�ɁA1 �ȏ�̃L�[ �X�g���[�N��
'*          �A�N�e�B�u�ȃE�B���h�E�ɑ���܂��B
'* [�Q  �l] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364423(v=msdn.10)>
'*
'* @param Keys ���M����L�[�X�g���[�N�𕶎��񎮂Ŏw��
'* @param Wait �ȗ��\�B
'*
'******************************************************************************
Public Sub SendKeys(Keys As String, Optional Wait)
    If IsMissing(Wait) Then
        Call mOrigin.SendKeys(Keys)
    Else
        Call mOrigin.SendKeys(Keys, Wait)
    End If
End Sub

'*-----------------------------------------------------------------------------
'* �g�����\�b�h
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [�T  �v] SendKeysCtlF ���\�b�h
'* [��  ��] SendKeys ���\�b�h�� Ctl + F �L�[�𑗐M����B
'*
'* @param Wait �ȗ��\�B
'*
'******************************************************************************
Public Sub SendKeysCtlF(Optional Wait As Long = 100)
    Call mOrigin.SendKeys("^f")
    Sleep Wait
End Sub

'******************************************************************************
'* [�T  �v] SendKeysESC ���\�b�h
'* [��  ��] SendKeys ���\�b�h�� ESC �L�[�𑗐M����B
'*
'* @param Wait �ȗ��\�B
'*
'******************************************************************************
Public Sub SendKeysESC(Optional Wait As Long = 100)
    Call mOrigin.SendKeys("{ESC}")
    Sleep Wait
End Sub

'******************************************************************************
'* [�T  �v] SendKeysHOME ���\�b�h
'* [��  ��] SendKeys ���\�b�h�� HOME �L�[�𑗐M����B
'*
'* @param Wait �ȗ��\�B
'*
'******************************************************************************
Public Sub SendKeysHOME(Optional Wait As Long = 100)
    Call mOrigin.SendKeys("{HOME}")
    Sleep Wait
End Sub

'******************************************************************************
'* [�T  �v] SendKeysShiftEND ���\�b�h
'* [��  ��] SendKeys ���\�b�h�� Shift + END �L�[�𑗐M����B
'*
'* @param Wait �ȗ��\�B
'*
'******************************************************************************
Public Sub SendKeysShiftEND(Optional Wait As Long = 100)
    Call mOrigin.SendKeys("+{END}")
    Sleep Wait
End Sub

'******************************************************************************
'* [�T  �v] SendKeysCtlC ���\�b�h
'* [��  ��] SendKeys ���\�b�h�� Ctl + C �L�[�𑗐M����B
'*
'* @param Wait �ȗ��\�B
'*
'******************************************************************************
Public Sub SendKeysCtlC(Optional Wait As Long = 100)
    Call mOrigin.SendKeys("^c")
    Sleep Wait
End Sub

'******************************************************************************
'* [�T  �v] SendKeysAltF4 ���\�b�h
'* [��  ��] SendKeys ���\�b�h�� Alt + F4 �L�[�𑗐M����B
'*
'* @param Wait �ȗ��\�B
'*
'******************************************************************************
Public Sub SendKeysAltF4(Optional Wait As Long = 100)
    Call mOrigin.SendKeys("%{F4}")
    Sleep Wait
End Sub

'******************************************************************************
'* [�T  �v] Is32BitProcessorForApp ���\�b�h
'* [��  ��] �g�p����A�v���P�[�V������32�r�b�g�����`�F�b�N���܂��B
'*
'* @return �`�F�b�N���ʁiTrue: 32Bit�AFalse: 64bit�j
'*
'******************************************************************************
Public Function Is32BitProcessorForApp() As Boolean
    Dim proc As String: proc = Me.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    If proc = "x86" Then
       Is32BitProcessorForApp = True
    Else
       Is32BitProcessorForApp = False
    End If
End Function

'******************************************************************************
'* [�T  �v] Is32BitProcessor ���\�b�h
'* [��  ��] �g�p����[���̃v���Z�b�T��32�r�b�g�����`�F�b�N���܂��B
'*
'* @return �`�F�b�N���ʁiTrue: 32Bit�AFalse: 64bit�j
'*
'******************************************************************************
Public Function Is32BitProcessor() As Boolean
    Dim proc As String: proc = Me.ExpandEnvironmentStrings("%PROCESSOR_ARCHITEW6432%")
    If proc = "x86" Then
       Is32BitProcessor = True
    Else
       Is32BitProcessor = False
    End If
End Function

'******************************************************************************
'* [�T  �v] CreateObject32bit ���\�b�h
'* [��  ��] 32�r�b�g����Object�𐶐����܂��B
'* [�Q  �l] <https://github.com/vocho/vbs/blob/a5c3ee608103638678c983da00ec290c4b8ab90c/CreateObject32bit.vbs>
'*
'* @param strClassName �����Ώۂ̃N���X���B"Shell.Application"���B
'* @return 32�r�b�g��Object
'*
'******************************************************************************
Public Function CreateObject32bit(ByVal strClassName As String) As Variant
    If Is32BitProcessorForApp Then
     Set CreateObject32bit = CreateObject(strClassName)
     Exit Function
    End If

    Dim shell As Object: Set shell = CreateObject("Shell.Application")
    shell.Windows().Item(0).PutProperty strClassName, Nothing
    
    ' �ꎞ�X�N���v�g�R�}���h�e�L�X�g����
    Dim strScriptCodes As String
    strScriptCodes = "CreateObject(""Shell.Application"").Windows().Item(0).PutProperty """ & strClassName & """, CreateObject(""" & strClassName & """)" & vbNewLine & _
                     "Set objExec = CreateObject(""WScript.Shell"").Exec(""MSHTA.EXE -"")" & vbNewLine & _
                     "Set objWMIService = GetObject(""winmgmts:"")" & vbNewLine & _
                     "lngCurrentPID = objWMIService.Get(""Win32_Process.Handle="" & objExec.ProcessID).ParentProcessID" & vbNewLine & _
                     "objExec.Terminate" & vbNewLine & _
                     "lngParentPID = objWMIService.Get(""Win32_Process.Handle="" & lngCurrentPID).ParentProcessID" & vbNewLine & _
                     "Do While objWMIService.ExecQuery(""SELECT * FROM Win32_Process WHERE ProcessID="" & lngParentPID).Count<>0" & vbNewLine & _
                     "    WScript.Sleep 1000" & vbNewLine & _
                     "Loop" & vbNewLine & _
                     "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbNewLine & _
                     "If objFSO.FileExists(WScript.ScriptFullName) Then objFSO.DeleteFile WScript.ScriptFullName" & vbNewLine & _
                     ""

    ' �ꎞ�X�N���v�g�t�@�C���쐬
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim strTempFile As String
    With fso
        Do
            strTempFile = .BuildPath(.GetSpecialFolder(2), .GetTempName() & ".vbs")
        Loop While .FileExists(strTempFile)
        With .OpenTextFile(strTempFile, 2, True)
            .WriteLine strScriptCodes
            .Close
        End With
    End With
    
    ' �ꎞ�X�N���v�g�t�@�C�����s(32bit)
    With Me.Environment("Process")
        .Item("SysWOW64") = fso.BuildPath(.Item("SystemRoot"), "SysWOW64")
        .Item("WScriptName") = fso.GetFileName("C:\WINDOWS\SysWOW64\cscript.exe")
        .Item("WScriptWOW64") = fso.BuildPath(.Item("SysWOW64"), .Item("WScriptName"))
        .Item("Run") = .Item("WScriptWOW64") & " """ & strTempFile & """"
         Me.Run .Item("Run"), True
    End With
    
    ' �I�u�W�F�N�g�󂯎��
    Do
        Set CreateObject32bit = shell.Windows().Item(0).GetProperty(strClassName)
    Loop While CreateObject32bit Is Nothing
End Function

'******************************************************************************
'* [�T  �v] CreateObjectFromWsh ���\�b�h
'* [��  ��] Wsh�o�R��Object�𐶐����܂��B
'*
'* @param strClassName �����Ώۂ̃N���X���B"Shell.Application"���B
'* @param strTypeName  �^�C�v���B
'* @return Object
'*
'******************************************************************************
Public Function CreateObjectFromWsh(ByVal strClassName As String, ByVal strTypeName As String) As Variant
    Dim shell As Object: Set shell = CreateObject("Shell.Application")
    shell.Windows().Item(0).PutProperty strClassName, Nothing

    ' �ꎞ�X�N���v�g�R�}���h�e�L�X�g����
    Dim strScriptCodes As String
    strScriptCodes = "CreateObject(""Shell.Application"").Windows().Item(0).PutProperty """ & strClassName & """, WScript.CreateObject(""" & strClassName & """, """ & strTypeName & """)" & vbNewLine & _
                     "Set objExec = CreateObject(""WScript.Shell"").Exec(""MSHTA.EXE -"")" & vbNewLine & _
                     "Set objWMIService = GetObject(""winmgmts:"")" & vbNewLine & _
                     "lngCurrentPID = objWMIService.Get(""Win32_Process.Handle="" & objExec.ProcessID).ParentProcessID" & vbNewLine & _
                     "objExec.Terminate" & vbNewLine & _
                     "lngParentPID = objWMIService.Get(""Win32_Process.Handle="" & lngCurrentPID).ParentProcessID" & vbNewLine & _
                     "Do While objWMIService.ExecQuery(""SELECT * FROM Win32_Process WHERE ProcessID="" & lngParentPID).Count<>0" & vbNewLine & _
                     "    WScript.Sleep 1000" & vbNewLine & _
                     "Loop" & vbNewLine & _
                     "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbNewLine & _
                     "If objFSO.FileExists(WScript.ScriptFullName) Then objFSO.DeleteFile WScript.ScriptFullName" & vbNewLine & _
                     ""
    ' �ꎞ�X�N���v�g�t�@�C���쐬
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim strTempFile
    With fso
        Do
            strTempFile = .BuildPath(.GetSpecialFolder(2), .GetTempName() & ".vbs")
        Loop While .FileExists(strTempFile)
        With .OpenTextFile(strTempFile, 2, True)
            .WriteLine strScriptCodes
            .Close
        End With
    End With

    ' �ꎞ�X�N���v�g�t�@�C�����s
    Me.Run "cscript.exe """ & strTempFile & """", True
    
    ' �I�u�W�F�N�g�󂯎��
    Do
        Set CreateObjectFromWsh = shell.Windows().Item(0).GetProperty(strClassName)
    Loop While CreateObjectFromWsh Is Nothing
End Function
