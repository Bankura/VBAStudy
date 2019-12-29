VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptingExTextStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] Scripting.TextStream�N���X�̃��b�v�E�g���N���X
'* [��  ��] Scripting.TextStream�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          Scripting.TextStream�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' Scripting.TextStream�N���X�ɂ��āB
' <https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/drive-object>
' ����̃f�B�X�N �h���C�u�܂��̓l�b�g���[�N���L�̃v���p�e�B�ɃA�N�Z�X�ł��܂��B
'
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As Scripting.TextStream
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
    'If Not mOrigin Is Nothing Then
    '    mOrigin.Close
    'End If
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* AtEndOfLine �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E���݂̈ʒu�͍s�̖������ǂ����������܂��B
'*-----------------------------------------------------------------------------
Property Get AtEndOfLine() As Boolean
    AtEndOfLine = mOrigin.AtEndOfLine
End Property

'*-----------------------------------------------------------------------------
'* AtEndOfStream �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E���݂̈ʒu�̓X�g���[���̍Ōォ�ǂ����������܂��B
'*-----------------------------------------------------------------------------
Property Get AtEndOfStream() As Boolean
    AtEndOfStream = mOrigin.AtEndOfStream
End Property

'*-----------------------------------------------------------------------------
'* Column �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E���݂̌����B
'*-----------------------------------------------------------------------------
Property Get Column() As Long
    Column = mOrigin.Column
End Property

'*-----------------------------------------------------------------------------
'* Line �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E���݂̍s�ԍ��B
'*-----------------------------------------------------------------------------
Property Get Line() As Long
    Line = mOrigin.Line
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Scripting.TextStream")
End Sub
Private Sub Class_Terminate()
    'If Not mOrigin Is Nothing Then
    '    mOrigin.Close
    'End If
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] CloseStream���\�b�h
'* [��  ��] Close�̃��b�p�[���\�b�h�B
'*          �e�L�X�g �X�g���[�������B
'******************************************************************************
Sub CloseStream()
    Call mOrigin.Close
End Sub

'******************************************************************************
'* [�T  �v] Read���\�b�h
'* [��  ��] Read�̃��b�p�[���\�b�h�B
'*          �w�肵�����̕����𕶎���ɓǂݍ��݂܂��B
'*
'* @param Characters
'* @return  �w�肵�����̕���
'******************************************************************************
Function Read(Characters As Long) As String
    Read = mOrigin.Read(Characters)
End Function

'******************************************************************************
'* [�T  �v] ReadAll���\�b�h
'* [��  ��] ReadAll�̃��b�p�[���\�b�h�B
'*          �X�g���[���S�̂𕶎���ɓǂݍ��݂܂��B
'*
'* @return �X�g���[���S�̂̕�����
'******************************************************************************
Function ReadAll() As String
    ReadAll = mOrigin.ReadAll()
End Function

'******************************************************************************
'* [�T  �v] ReadLine���\�b�h
'* [��  ��] ReadLine�̃��b�p�[���\�b�h�B
'*          �s�S�̂𕶎���ɓǂݍ��݂܂��B
'*
'* @return �s�S�̂̕�����
'******************************************************************************
Function ReadLine() As String
    ReadLine = mOrigin.ReadLine()
End Function

'******************************************************************************
'* [�T  �v] Skip���\�b�h
'* [��  ��] Skip�̃��b�p�[���\�b�h�B
'*          �w�肵�����������X�L�b�v���܂��
'*
'* @param Characters ������
'******************************************************************************
Sub Skip(Characters As Long)
    Call mOrigin.Skip(Characters)
End Sub

'******************************************************************************
'* [�T  �v] SkipLine���\�b�h
'* [��  ��] SkipLine�̃��b�p�[���\�b�h�B
'*          �s���X�L�b�v���܂��B
'*
'******************************************************************************
Sub SkipLine()
    Call mOrigin.SkipLine
End Sub

'******************************************************************************
'* [�T  �v] WriteText���\�b�h
'* [��  ��] Write�̃��b�p�[���\�b�h�B
'*          �X�g���[���ɕ�������������݂܂��B
'*
'* @param Text �������ރe�L�X�g
'******************************************************************************
Sub WriteText(text As String)
    Call mOrigin.Write(text)
End Sub

'******************************************************************************
'* [�T  �v] WriteBlankLines���\�b�h
'* [��  ��] WriteBlankLines�̃��b�p�[���\�b�h�B
'*          �X�g���[���ɕ����s�̋󔒍s���������݂܂��B
'*
'* @param Lines �󔒍s��
'******************************************************************************
Sub WriteBlankLines(Lines As Long)
    Call mOrigin.WriteBlankLines(Lines)
End Sub

'******************************************************************************
'* [�T  �v] WriteLine���\�b�h
'* [��  ��] WriteLine�̃��b�p�[���\�b�h�B
'*          �X�g���[���ɕ�����ƍs���R�[�h���������݂܂��B
'*
'* @param Text �ȗ��\�B�������ރe�L�X�g�B
'******************************************************************************
Sub WriteLine(Optional text As String)
    If IsMissing(text) Then
        Call mOrigin.WriteLine
    Else
        Call mOrigin.WriteLine(text)
    End If
End Sub

'*-----------------------------------------------------------------------------
'* �g�����\�b�h
'*-----------------------------------------------------------------------------
