VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SCExScriptControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] MSScriptControl.ScriptControl �N���X�̃��b�v�E�g���N���X
'* [��  ��] MSScriptControl.ScriptControl �N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          MSScriptControl.ScriptControl �N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' MSScriptControl.ScriptControl �N���X�ɂ��āB
' Control to host scripting engines that understand the ActiveX Scripting interface.
'
' [�C�x���g]
' Event Error()
'     Event fired when any error occurs in the scripting engine
' Event Timeout()
'     Event fired when a script is aborted because of a timeout
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As MSScriptControl.ScriptControl
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
'* AllowUI �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EEnable or disable display of the UI.
'*-----------------------------------------------------------------------------
Property Get AllowUI() As Boolean
    AllowUI = mOrigin.AllowUI
End Property

'*-----------------------------------------------------------------------------
'* CodeObject �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EObject exposed by the scripting engine that contains methods and properties defined in the code added to the global module.
'*-----------------------------------------------------------------------------
Property Get CodeObject() As Object
    Set CodeObject = mOrigin.CodeObject
End Property
    
'*-----------------------------------------------------------------------------
'* Error �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe last error reported by the scripting engine.
'*-----------------------------------------------------------------------------
Property Get Error() As SCExError
    Dim wrap As SCExError: Set wrap = New SCExError
    Set wrap.Origin = mOrigin.Error
    Set Error = wrap
End Property

'*-----------------------------------------------------------------------------
'* Language �v���p�e�B
'*
'* [��  ��]
'* �ELanguage engine to use.
'*   "JScript",  "VBScript", "JavaScript"
'*-----------------------------------------------------------------------------
Property Get Language() As String
    Language = mOrigin.Language
End Property
Property Let Language(arg As String)
    mOrigin.Language = arg
End Property

'*-----------------------------------------------------------------------------
'* Modules �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �ECollection of modules for the ScriptControl.
'*-----------------------------------------------------------------------------
Property Get Modules() As SCExModules
    Dim wrap As SCExModules: Set wrap = New SCExModules
    Set wrap.Origin = mOrigin.Modules
    Set Modules = wrap
End Property

'*-----------------------------------------------------------------------------
'* Procedures �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �ECollection of procedures that are defined in the global module.
'*-----------------------------------------------------------------------------
Property Get Procedures() As SCExProcedures
    Dim wrap As SCExProcedures: Set wrap = New SCExProcedures
    Set wrap.Origin = mOrigin.Procedures
    Set Procedures = wrap
End Property

'*-----------------------------------------------------------------------------
'* SitehWnd �v���p�e�B
'*
'* [��  ��]
'* �EhWnd used as a parent for displaying UI.
'*-----------------------------------------------------------------------------
Property Get SitehWnd() As Long
    SitehWnd = mOrigin.SitehWnd
End Property
Property Let SitehWnd(arg As Long)
    mOrigin.SitehWnd = arg
End Property
    
'*-----------------------------------------------------------------------------
'* State �v���p�e�B
'*
'* [��  ��]
'* �EState of the control.
'*-----------------------------------------------------------------------------
Property Get State() As ScriptControlStates
    State = mOrigin.State
End Property
Property Let State(arg As ScriptControlStates)
    mOrigin.State = arg
End Property

'*-----------------------------------------------------------------------------
'* Timeout �v���p�e�B
'*
'* [��  ��]
'* �ELength of time in milliseconds that a script can execute before being considered hung.
'*-----------------------------------------------------------------------------
Property Get Timeout() As Long
    Timeout = mOrigin.Timeout
End Property
Property Let Timeout(arg As Long)
    mOrigin.Timeout = arg
End Property

'*-----------------------------------------------------------------------------
'* UseSafeSubset �v���p�e�B
'*
'* [��  ��]
'* �EForce script to execute in safe mode and disallow potentially harmful actions.
'*-----------------------------------------------------------------------------
Property Get UseSafeSubset() As Boolean
    UseSafeSubset = mOrigin.UseSafeSubset
End Property
Property Let UseSafeSubset(arg As Boolean)
    mOrigin.UseSafeSubset = arg
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject32bit("MSScriptControl.ScriptControl")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] AddCode ���\�b�h
'* [��  ��] AddCode �̃��b�p�[���\�b�h�B
'*          Add code to the global module.
'*
'* @param Code xx
'*
'******************************************************************************
Sub AddCode(Code As String)
    Call mOrigin.AddCode(Code)
End Sub

'******************************************************************************
'* [�T  �v] AddObject ���\�b�h
'* [��  ��] AddObject �̃��b�p�[���\�b�h�B
'*          Add an object to the global namespace of the scripting engine.
'*
'* @param Name xx
'* @param Object xx
'* @param AddMembers xx
'*
'******************************************************************************
Sub AddObject(Name As String, Object As Object, Optional AddMembers As Boolean = False)
    Call mOrigin.AddObject(Name, Object, AddMembers)
End Sub
  
'******************************************************************************
'* [�T  �v] Eval ���\�b�h
'* [��  ��] Eval �̃��b�p�[���\�b�h�B
'*          Evaluate an expression within the context of the global module.
'*
'* @param Expression xx
'* @return Variant
'*
'******************************************************************************
Function Eval(Expression As String)
    Eval = mOrigin.Eval(Expression)
End Function

'******************************************************************************
'* [�T  �v] ExecuteStatement ���\�b�h
'* [��  ��] ExecuteStatement �̃��b�p�[���\�b�h�B
'*          Execute a statement within the context of the global module.
'*
'* @param Statement xx
'*
'******************************************************************************
Sub ExecuteStatement(Statement As String)
    Call mOrigin.ExecuteStatement(Statement)
End Sub

'******************************************************************************
'* [�T  �v] Reset ���\�b�h
'* [��  ��] Reset �̃��b�p�[���\�b�h�B
'*          Reset the scripting engine to a newly created state.
'*
'******************************************************************************
Sub Reset()
    Call mOrigin.Reset
End Sub

'******************************************************************************
'* [�T  �v] Run ���\�b�h
'* [��  ��] Run �̃��b�p�[���\�b�h�B
'*          Call a procedure defined in the global module.
'*
'* @param ProcedureName xx
'* @param Parameters() xx
'* @return Variant
'*
'******************************************************************************
Function Run(ProcedureName As String, ParamArray Parameters() As Variant)
    Select Case UBound(Parameters)
        Case LBound(Parameters)
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)))
        Case LBound(Parameters) + 1
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1))
        Case LBound(Parameters) + 2
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2))
        Case LBound(Parameters) + 3
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3))
        Case LBound(Parameters) + 4
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3), _
                                             Parameters(LBound(Parameters) + 4))
        Case LBound(Parameters) + 5
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3), _
                                             Parameters(LBound(Parameters) + 4), _
                                             Parameters(LBound(Parameters) + 5))
        Case LBound(Parameters) + 6
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3), _
                                             Parameters(LBound(Parameters) + 4), _
                                             Parameters(LBound(Parameters) + 5), _
                                             Parameters(LBound(Parameters) + 6))
        Case LBound(Parameters) + 7
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3), _
                                             Parameters(LBound(Parameters) + 4), _
                                             Parameters(LBound(Parameters) + 5), _
                                             Parameters(LBound(Parameters) + 6), _
                                             Parameters(LBound(Parameters) + 7))
        Case LBound(Parameters) + 8
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3), _
                                             Parameters(LBound(Parameters) + 4), _
                                             Parameters(LBound(Parameters) + 5), _
                                             Parameters(LBound(Parameters) + 6), _
                                             Parameters(LBound(Parameters) + 7), _
                                             Parameters(LBound(Parameters) + 8))
        Case LBound(Parameters) + 9
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3), _
                                             Parameters(LBound(Parameters) + 4), _
                                             Parameters(LBound(Parameters) + 5), _
                                             Parameters(LBound(Parameters) + 6), _
                                             Parameters(LBound(Parameters) + 7), _
                                             Parameters(LBound(Parameters) + 8), _
                                             Parameters(LBound(Parameters) + 9))
    End Select
End Function


'*-----------------------------------------------------------------------------
'* �g�����\�b�h
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [�T  �v] EncodeURI ���\�b�h
'* [��  ��] URI�G���R�[�h�������s���܂��B
'* [�Q  �l] <https://github.com/vocho/vbs/blob/a5c3ee608103638678c983da00ec290c4b8ab90c/CreateObject32bit.vbs>
'*          <https://srad.jp/~black-hole/journal/516080/>
'*
'* @param strURI ���G���R�[�hURI
'* @return �G���R�[�h�ς�URIt
'*
'******************************************************************************
Function EncodeURI(ByVal strURI As String) As String
    With Me
        .Language = "JScript"
        EncodeURI = .CodeObject.encodeURIComponent(strURI)
    End With
End Function

'******************************************************************************
'* [�T  �v] DecodeURI ���\�b�h
'* [��  ��] URI�f�R�[�h�������s���܂��B
'* [�Q  �l] <https://github.com/vocho/vbs/blob/a5c3ee608103638678c983da00ec290c4b8ab90c/CreateObject32bit.vbs>
'*          <https://srad.jp/~black-hole/journal/516080/>
'*
'* @param strURI ���f�R�[�hURI
'* @return �f�R�[�h�ς�URI
'*
'******************************************************************************
Function DecodeURI(ByVal strURI As String) As String
    With Me
        .Language = "JScript"
        DecodeURI = .CodeObject.decodeURIComponent(strURI)
    End With
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
Private Function CreateObject32bit(strClassName As String)
    If Is32BitProcessorForApp Then
     Set CreateObject32bit = CreateObject(strClassName)
     Exit Function
    End If
    
    Dim shell As Object: Set shell = CreateObject("Shell.Application")
    shell.Windows().Item(0).PutProperty strClassName, Nothing
    ' �ꎞ�X�N���v�g�R�}���h�e�L�X�g����
    Dim strScript As String
    strScript = "CreateObject(""Shell.Application"").Windows().Item(0).PutProperty """ & strClassName & """, CreateObject(""" & strClassName & """)" & vbNewLine & _
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
            strTempFile = .BuildPath(.GetSpecialFolder(2), .GetTempName() & ".vbs") ' Const TemporaryFolder = 2
        Loop While .FileExists(strTempFile)
        With .OpenTextFile(strTempFile, 2, True) ' ForWriting = 2, ForAppending = 8
            .WriteLine strScript
            .Close
        End With
    End With
    
    ' �ꎞ�X�N���v�g�t�@�C�����s(32bit)
    With CreateObject("WScript.Shell").Environment("Process")
        .Item("SysWOW64") = fso.BuildPath(.Item("SystemRoot"), "SysWOW64")
        .Item("WScriptName") = fso.GetFileName("C:\WINDOWS\SysWOW64\cscript.exe")
        .Item("WScriptWOW64") = fso.BuildPath(.Item("SysWOW64"), .Item("WScriptName"))
        .Item("Run") = .Item("WScriptWOW64") & " """ & strTempFile & """"
         CreateObject("WScript.Shell").Run .Item("Run"), True
    End With
    
    ' �I�u�W�F�N�g�󂯎��
    Do
        Set CreateObject32bit = shell.Windows().Item(0).GetProperty(strClassName)
    Loop While CreateObject32bit Is Nothing
End Function

'******************************************************************************
'* [�T  �v] Is32BitProcessorForApp ���\�b�h
'* [��  ��] �g�p����A�v���P�[�V������32�r�b�g�����`�F�b�N���܂��B
'*
'* @return �`�F�b�N���ʁiTrue: 32Bit�AFalse: 64bit�j
'*
'******************************************************************************
Private Function Is32BitProcessorForApp() As Boolean
    Dim proc As String: proc = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    If proc = "x86" Then
       Is32BitProcessorForApp = True
    Else
       Is32BitProcessorForApp = False
    End If
End Function