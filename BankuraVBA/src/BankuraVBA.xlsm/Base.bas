Attribute VB_Name = "Base"
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] BankuraVBA���ʊ�Ճ��W���[��
'* [��  ��] ���ʂŎg�p����v���V�[�W����񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI��`
'******************************************************************************

'******************************************************************************
'* Enum��`
'******************************************************************************


'******************************************************************************
'* �\���̒�`
'******************************************************************************

'******************************************************************************
'* �萔��`
'******************************************************************************
Public Const MAX_INT As Integer = 32767
Public Const MAX_LONG As Long = 2147483647
Public Const TWIP As Long = 567

#If Win64 Then
    Public Const LONGPTR_SIZE = 8
#Else
    Public Const LONGPTR_SIZE = 4
#End If

' For Wsh
Public Const WshHide = 0
Public Const ForReading = 1
    
'******************************************************************************
'* �ϐ���`
'******************************************************************************
Private mDisplayAlerts As Boolean
Private mScreenUpdating As Boolean
Private mCalculation As Long
Private mEnableEvents As Boolean
Private mRegExp As Object
Private mShell As Object
Private mWshNetwork As Object
Private mSc As Object
Private mSettingInfo As SettingInfo
Private mWinApi As WinAPI

'******************************************************************************
'* �v���V�[�W����`
'******************************************************************************

'******************************************************************************
'* [�T  �v] ChangeDisplayWorkbookTabs
'* [��  ��] �V�[�g���o���̕\���E��\����؂�ւ���B
'*
'******************************************************************************
Public Sub ChangeDisplayWorkbookTabs()
    With ActiveWindow
        If .DisplayWorkbookTabs Then
            '�V�[�g���o�����\��
            .DisplayWorkbookTabs = False
        Else
            '�V�[�g���o����\��
            .DisplayWorkbookTabs = True
        End If
    End With
End Sub

'******************************************************************************
'* [�T  �v] ChangeDisplayGridlines
'* [��  ��] �r���̕\���E��\����؂�ւ���B
'*
'******************************************************************************
Public Sub ChangeDisplayGridlines()
    With ActiveWindow
        If .DisplayGridlines Then
            '�r�����\��
            .DisplayGridlines = False
        Else
            '�r����\��
            .DisplayGridlines = True
        End If
    End With
End Sub

'******************************************************************************
'* [�T  �v] ChangeDisplayHeadings
'* [��  ��] �s��ԍ��̕\���E��\����؂�ւ���B
'*
'******************************************************************************
Public Sub ChangeDisplayHeadings()
    With ActiveWindow
        If .DisplayHeadings Then
            '�s��ԍ����\��
            .DisplayHeadings = False
        Else
            '�s��ԍ���\��
            .DisplayHeadings = True
        End If
    End With
End Sub

'******************************************************************************
'* [�T  �v] ChangeDisplayHeadings
'* [��  ��] �X�N���[���o�[�̕\���E��\����؂�ւ���B
'*
'*******************************************************************************
Public Sub ChangeDisplayScrollBar()
    With ActiveWindow
        If .DisplayHorizontalScrollBar Then
            '�X�N���[���o�[���\��
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
        Else
            '�X�N���[���o�[��\��
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
        End If
    End With
End Sub

'******************************************************************************
'* [�T  �v] ChangeReferenceStyle
'* [��  ��] Excel�Q�ƌ`����؂�ւ���B
'*
'******************************************************************************
Public Sub ChangeReferenceStyle()
    With Application
        If .ReferenceStyle = xlA1 Then
            .ReferenceStyle = xlR1C1
        Else
            .ReferenceStyle = xlA1
        End If
    End With
End Sub

'******************************************************************************
'* [�T  �v] �G���[�����B
'* [��  ��] �G���[�������̏������s���B
'*
'******************************************************************************
Public Sub ErrorProcess()
    Debug.Print "�G���[���� Number: " & err.Number & " Source: " & err.Source & " Description: " & err.Description
End Sub

'******************************************************************************
'* [�T  �v] �J�n�����B
'* [��  ��] �����̃X�s�[�h����̂��߁AExcel�̐ݒ��ύX����B
'*
'******************************************************************************
Public Sub StartProcess()
    Call SaveApplicationProperties
    With Application
        .Cursor = xlWait
        .DisplayAlerts = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
End Sub

'******************************************************************************
'* [�T  �v] �I�������B
'* [��  ��] �����̃X�s�[�h����̂��ߕύX����Excel�̐ݒ�����ɖ߂��B
'*
'******************************************************************************
Public Sub EndProcess()
    With Application
        .Cursor = xlDefault
        .DisplayAlerts = mDisplayAlerts
        .ScreenUpdating = mScreenUpdating
        .Calculation = mCalculation
        .EnableEvents = mEnableEvents
        .StatusBar = False
    End With
End Sub

'******************************************************************************
'* [�T  �v] Application�ݒ�ޔ������B
'* [��  ��] Application�̐ݒ�������o�ϐ��ɑޔ�����B
'*
'******************************************************************************
Public Sub SaveApplicationProperties()
    With Application
        mDisplayAlerts = .DisplayAlerts
        mScreenUpdating = .ScreenUpdating
        mCalculation = .Calculation
        mEnableEvents = .EnableEvents
    End With
End Sub

'******************************************************************************
'* [�T  �v] ���K�\���I�u�W�F�N�g�擾�����B
'* [��  ��] ���K�\���I�u�W�F�N�g���擾����B�������̏ꍇ��������B
'*
'******************************************************************************
Public Function GetRegExp() As Object
    If mRegExp Is Nothing Then
        Set mRegExp = CreateObject("VBScript.RegExp")
    End If
    Set GetRegExp = mRegExp
End Function

'******************************************************************************
'* [�T  �v] Shell�I�u�W�F�N�g�擾�����B
'* [��  ��] Shell�I�u�W�F�N�g���擾����B�������̏ꍇ��������B
'*
'******************************************************************************
Public Function GetShell() As Object
    If mShell Is Nothing Then
        Set mShell = CreateObject("Shell.Application")
    End If
    Set GetShell = mShell
End Function

'******************************************************************************
'* [�T  �v] WScript.Network�I�u�W�F�N�g�擾�����B
'* [��  ��] WScript.Network�I�u�W�F�N�g���擾����B�������̏ꍇ��������B
'*
'******************************************************************************
Public Function GetWshNetwork() As Object
    If mWshNetwork Is Nothing Then
        Set mWshNetwork = CreateObject("WScript.Network")
    End If
    Set GetWshNetwork = mWshNetwork
End Function

'******************************************************************************
'* [�T  �v] ScriptControl�I�u�W�F�N�g�擾�����B
'* [��  ��] ScriptControl�I�u�W�F�N�g���擾����B�������̏ꍇ��������B
'*
'******************************************************************************
Public Function GetScriptControl() As Object
    If mSc Is Nothing Then
        Set mSc = CreateObject32bit("MSScriptControl.ScriptControl")
    End If
    Set GetScriptControl = mSc
End Function


'******************************************************************************
'* [�T  �v] CDO.Message�I�u�W�F�N�g���������B
'* [��  ��] CDO.Message�I�u�W�F�N�g�𐶐�����B
'*
'******************************************************************************
Public Function CreateCDOMessage() As Object
    Set CreateCDOMessage = CreateObject("CDO.Message")
End Function

'******************************************************************************
'* [�T  �v] WinAPI�I�u�W�F�N�g�擾�����B
'* [��  ��] WinAPI�I�u�W�F�N�g���擾����B�������̏ꍇ��������B
'*
'******************************************************************************
Public Function GetWinAPI() As WinAPI
    If mWinApi Is Nothing Then
        Set mWinApi = New WinAPI
    End If
    Set GetWinAPI = mWinApi
End Function

'******************************************************************************
'* [�T  �v] �ݒ���I�u�W�F�N�g�擾�����B
'* [��  ��] �ݒ���I�u�W�F�N�g���擾����B�������̏ꍇ��������B
'*
'******************************************************************************
Public Function GetSettingInfo() As SettingInfo
    If mSettingInfo Is Nothing Then
        Set mSettingInfo = New SettingInfo
    End If
    Set GetSettingInfo = mSettingInfo
End Function

'*******************************************************************************
'* [�T  �v] �R���s���[�^���ݒ菈��
'* [��  ��] ���s�[���̃R���s���[�^�����擾�B
'*
'* @param String �R���s���[�^��
'*
'*******************************************************************************
Public Function GetComputerName() As String
    GetComputerName = Core.Wsh.ComputerName
End Function

'******************************************************************************
'* [�T  �v] ���s�A�v���P�[�V�������菈��
'* [��  ��] ���s�A�v���P�[�V������Excel�����肷��B
'*
'* @param Boolean �������ʁiTrue:Excel False�FExcel�ȊO�j
'*
'******************************************************************************
Public Function CheckXlApplication() As Boolean
    CheckXlApplication = InStr(Application.Name, "Excel") > 0
End Function

'******************************************************************************
'* [�T  �v] Is32BitProcessorForApp
'* [��  ��] �g�p����A�v���P�[�V������32�r�b�g�����`�F�b�N����B
'*
'* @return �`�F�b�N���ʁiTrue: 32Bit�AFalse: 64bit�j
'*
'******************************************************************************
Public Function Is32BitProcessorForApp() As Boolean
    Dim proc As String: proc = Wsh.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    If proc = "x86" Then
       Is32BitProcessorForApp = True
    Else
       Is32BitProcessorForApp = False
    End If
End Function

'******************************************************************************
'* [�T  �v] Is32BitProcessor
'* [��  ��] �g�p����[���̃v���Z�b�T��32�r�b�g�����`�F�b�N����B
'*
'* @return �`�F�b�N���ʁiTrue: 32Bit�AFalse: 64bit�j
'*
'******************************************************************************
Public Function Is32BitProcessor() As Boolean
    Dim proc As String: proc = Wsh.ExpandEnvironmentStrings("%PROCESSOR_ARCHITEW6432%")
    If proc = "x86" Then
       Is32BitProcessor = True
    Else
       Is32BitProcessor = False
    End If
End Function

'******************************************************************************
'* [�T  �v] CreateObject32bit
'* [��  ��] 32�r�b�g����Object�𐶐�����B
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
    
    Base.GetShell.Windows().Item(0).PutProperty strClassName, Nothing
    
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
    With IO.fso
        Dim strTempFile As String
        Do
            strTempFile = .BuildPath(.GetSpecialFolder(2), .GetTempName() & ".vbs")
        Loop While .FileExists(strTempFile)
        With .OpenTextFile(strTempFile, 2, True)
            .WriteLine strScriptCodes
            .Close
        End With
    End With
    
    ' �ꎞ�X�N���v�g�t�@�C�����s(32bit)
    With Core.Wsh.Environment("Process")
        .Item("SysWOW64") = IO.fso.BuildPath(.Item("SystemRoot"), "SysWOW64")
        .Item("WScriptName") = IO.fso.GetFileName("C:\WINDOWS\SysWOW64\cscript.exe")
        .Item("WScriptWOW64") = IO.fso.BuildPath(.Item("SysWOW64"), .Item("WScriptName"))
        .Item("Run") = .Item("WScriptWOW64") & " """ & strTempFile & """"
         Core.Wsh.Run .Item("Run"), True
    End With
    
    ' �I�u�W�F�N�g�󂯎��
    Do
        Set CreateObject32bit = Base.GetShell.Windows().Item(0).GetProperty(strClassName)
    Loop While CreateObject32bit Is Nothing
End Function

