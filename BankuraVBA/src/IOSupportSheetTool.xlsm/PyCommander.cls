VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PyCommander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] Python�R�}���h���s�N���X
'* [��  ��] Python�̎��s�Ɋւ��鏈����񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2020-2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI��`
'******************************************************************************

'******************************************************************************
'* Enum��`
'******************************************************************************

'******************************************************************************
'* �萔��`
'******************************************************************************

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mCommander As Commander
Private mPythonInstallPath As String

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************
'*-----------------------------------------------------------------------------
'* PythonInstallPath �v���p�e�B
'*
'* Python�̃C���X�g�[����p�X
'*-----------------------------------------------------------------------------
Public Property Get PythonInstallPath() As String
    PythonInstallPath = mPythonInstallPath
End Property
Public Property Let PythonInstallPath(arg As String)
    mPythonInstallPath = arg
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Call SetPythonInstallPathFromRegistry
    Call SetPythonEnviron

    Set mCommander = New Commander
    Call CommandMode
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] CommandMode
'* [��  ��] �R�}���h�����s���郂�[�h�ɕύX����B
'*
'******************************************************************************
Public Sub CommandMode()
    mCommander.BaseCommand = "cmd.exe /c " & mPythonInstallPath & "\python.exe -c ""{0}"" > {1}"
End Sub

'******************************************************************************
'* [�T  �v] ScriptMode
'* [��  ��] �X�N���v�g�����s���郂�[�h�ɕύX����B
'*
'******************************************************************************
Public Sub ScriptMode()
    mCommander.BaseCommand = "cmd.exe /c " & mPythonInstallPath & "\python.exe {0} > {1}"
End Sub

'******************************************************************************
'* [�T  �v] SetPythonEnviron
'* [��  ��] Python�̊��ϐ���ݒ肷��B
'*          ���ϐ��uPath�v��Python�̃p�X��ǉ�����B
'*
'******************************************************************************
Public Sub SetPythonEnviron()
    If mPythonInstallPath <> "" Then
        Call Base.AppendEnvItem("Path", mPythonInstallPath & "\Scripts")
        Call Base.AppendEnvItem("Path", mPythonInstallPath & "\Library\bin")
        Call Base.AppendEnvItem("Path", mPythonInstallPath)
    End If
End Sub

'******************************************************************************
'* [�T  �v] WriteScriptAndRun
'* [��  ��] �X�N���v�g���쐬���A���s�������ʂ��擾����B
'*
'* @param scriptString Script������A�܂���1�����z��
'* @param recieveResultString ���ʂ𕶎���Ŏ󂯎�邩
'*                            �iTrue:String, False:�z��j
'* @return Variant �R�}���h���s����
'******************************************************************************
Public Function WriteScriptAndRun(scriptString, Optional recieveResultString As Boolean = True) As Variant
    Dim tempScrPath As String: tempScrPath = FileUtils.GetTempFilePath(, ".py")
    Call FileUtils.WriteUTF8TextFile(tempScrPath, scriptString)
    
    WriteScriptAndRun = ExecScript(tempScrPath, recieveResultString)

    Call IO.fso.DeleteFile(tempScrPath, True)
End Function

'******************************************************************************
'* [�T  �v] ExecCommand
'* [��  ��] �R�}���h�����s�������ʂ��擾����B
'*
'* @param scriptString Script������A�܂���1�����z��
'* @param recieveResultString ���ʂ𕶎���Ŏ󂯎�邩
'*                            �iTrue:String, False:�z��j
'* @return Variant �R�}���h���s����
'******************************************************************************
Public Function ExecCommand(commandString As String, Optional recieveResultString As Boolean = True) As Variant
    Call CommandMode
    ExecCommand = Exec(commandString, recieveResultString)
End Function

'******************************************************************************
'* [�T  �v] ExecScript
'* [��  ��] �X�N���v�g�����s�������ʂ��擾����B
'*
'* @param scriptPath �X�N���v�g�̃p�X
'* @param recieveResultString ���ʂ𕶎���Ŏ󂯎�邩
'*                            �iTrue:String, False:�z��j
'* @return Variant �X�N���v�g���s����
'******************************************************************************
Public Function ExecScript(ScriptPath As String, Optional recieveResultString As Boolean = True) As Variant
    Call ScriptMode
    ExecScript = Exec(ScriptPath, recieveResultString)
End Function

'******************************************************************************
'* [�T  �v] Exec
'* [��  ��] �w�肵���R�}���h�܂��̓X�N���v�g�����s�������ʂ��擾����B
'* [��  �l] ���O�� CommandMode / ScriptMode ���Ăяo���ă��[�h���w�肷��z��B
'*
'* @param execString ���s������i�R�}���h�܂��̓X�N���v�g�j
'* @param recieveResultString ���ʂ𕶎���Ŏ󂯎�邩
'*                            �iTrue:String, False:�z��j
'* @return Variant �X�N���v�g���s����
'******************************************************************************
Public Function Exec(execString As String, Optional recieveResultString As Boolean = True) As Variant
    If recieveResultString Then
        Exec = mCommander.GetCommandResult(execString)
    Else
        Exec = mCommander.GetCommandResultAsArray(execString)
    End If
End Function

'******************************************************************************
'* [�T  �v] SetPythonInstallPathFromRegistry
'* [��  ��] ���W�X�g������Python�̃C���X�g�[���p�X���擾���A�ݒ肷��B
'* [�Q  �l] https://www.python.org/dev/peps/pep-0514/
'*
'******************************************************************************
Private Sub SetPythonInstallPathFromRegistry()
    Const PYTHON_REG_PATH As String = "SOFTWARE\Python\PythonCore"
    Const PYTHON_REG_WOW_PATH As String = "SOFTWARE\Wow6432Node\Python\PythonCore"
    
    Dim stdRegProv As Object: Set stdRegProv = CreateStdRegProv()
    Dim regPythonInstallPathString As String
    Dim regRoot As Long

    Dim vArr
    stdRegProv.EnumKey HKEY_CURRENT_USER, PYTHON_REG_PATH, vArr
    If IsNull(vArr) Then
        stdRegProv.EnumKey HKEY_LOCAL_MACHINE, PYTHON_REG_PATH, vArr
        If IsNull(vArr) Then
            stdRegProv.EnumKey HKEY_LOCAL_MACHINE, PYTHON_REG_WOW_PATH, vArr
            regPythonInstallPathString = PYTHON_REG_WOW_PATH & "\{0}\InstallPath"
            regRoot = HKEY_LOCAL_MACHINE
        Else
            regPythonInstallPathString = PYTHON_REG_PATH & "\{0}\InstallPath"
            regRoot = HKEY_LOCAL_MACHINE
        End If
    Else
        regPythonInstallPathString = PYTHON_REG_PATH & "\{0}\InstallPath"
        regRoot = HKEY_CURRENT_USER
    End If
    
    If Not IsNull(vArr) Then
        ArrayUtils.Sort vArr, False
        stdRegProv.GetStringValue regRoot, StringUtils.PlaceHolder(regPythonInstallPathString, CStr(vArr(0))), , mPythonInstallPath
    End If
End Sub