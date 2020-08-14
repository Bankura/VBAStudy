VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClipBoardUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] ClipBoard���[�e�B���e�B�N���X
'* [��  ��] ClipBoard�Ɋւ��郆�[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'* [�Q  �l] https://www.ka-net.org/blog/?p=7537
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
    mCommander.BaseCommand = "cmd.exe /c {0} > {1}"
End Sub
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] SetClipBoard
'* [��  ��] �N���b�v�{�[�h�ɕ�������i�[����B
'*
'* @param str �i�[������
'*
'******************************************************************************
Public Sub SetClipBoard(ByVal str As String)
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .text = str
        .SelStart = 0
        .SelLength = .textLength
        .Copy
    End With
End Sub

'******************************************************************************
'* [�T  �v] GetClipBoard
'* [��  ��] �N���b�v�{�[�h���當������擾����B
'*
'* @return String �擾������
'*
'******************************************************************************
Public Function GetClipBoard() As String
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        If .CanPaste = True Then .Paste
        GetClipBoard = .text
    End With
End Function