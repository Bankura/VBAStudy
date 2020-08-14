VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PowerShellCommander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] PowerShell�R�}���h���s�N���X
'* [��  ��] PowerShell�R�}���h�i�X�N���v�g���b�g�j�̎��s�Ɋւ��鏈����񋟂���B
'*
'* [�Q  �l]�@https://thom.hateblo.jp/entry/2017/01/31/012913
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
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

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mCommander = New Commander
    mCommander.BaseCommand = "powershell -ExecutionPolicy RemoteSigned -Command Invoke-Expression ""{0} | Out-File -filePath {1} -encoding Default"""
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] GetCommandResultAsTextStream
'* [��  ��] �R�}���h�����s���A���s�������ʂ��ꎞ�t�@�C���ɕۑ�����B
'*          �ꎞ�t�@�C��������s���ʂ�TextStream�Ŏ擾����B
'*
'* @param commandString �R�}���h������
'* @param myTempPath �ꎞ�t�@�C���p�X�i�W���̃p�X�ȊO�ɂ������ꍇ�Ɏw��j
'* @return Object(Of Scripting.TextStream)
'******************************************************************************
Public Function GetCommandResultAsTextStream(commandString, Optional myTempPath) As Object
    Set GetCommandResultAsTextStream = mCommander.GetCommandResultAsTextStream(commandString, myTempPath)
End Function

'******************************************************************************
'* [�T  �v] GetCommandResult
'* [��  ��] �R�}���h�����s���A���s�������ʕ�������擾����B
'*
'* @param commandString �R�}���h������
'* @return String �R�}���h���s����
'******************************************************************************
Public Function GetCommandResult(commandString) As String
    GetCommandResult = mCommander.GetCommandResult(commandString)
End Function

'******************************************************************************
'* [�T  �v] GetCommandResultAsArray
'* [��  ��]�R�}���h�����s���A���s�������ʂ��s���ɕ������������z��Ŏ擾����B
'*
'* @param commandString �R�}���h������
'* @return String() �R�}���h���s����
'******************************************************************************
Public Function GetCommandResultAsArray(commandString) As String()
    GetCommandResultAsArray = mCommander.GetCommandResultAsArray(commandString)
End Function
