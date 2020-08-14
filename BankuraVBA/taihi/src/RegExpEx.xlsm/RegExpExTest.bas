Attribute VB_Name = "RegExpExTest"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] RegExpEx�e�X�g�p���W���[��
'* [��  ��] RegExpEx�e�X�g�p���W���[��
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* [�T  �v] RegExpExRegExp ��Test�B
'* [��  ��] RegExpExRegExp ��Test�p�����B
'* [�Q  �l] <https://vbabeginner.net/vba%E3%81%A7%E6%AD%A3%E8%A6%8F%E8%A1%A8%E7%8F%BE%E3%82%92%E5%88%A9%E7%94%A8%E3%81%99%E3%82%8B/>
'*
'******************************************************************************
Sub RegExpExRegExpTest()
    Dim reg As RegExpExRegExp
    Set reg = New RegExpExRegExp
    Dim s As String

    Dim oMatches As RegExpExMatchCollection
    Dim oMatch As RegExpExMatch
    Dim i As Long
    Dim iCount
    Dim iFirstIndex
    Dim iLength
    Dim sValue
    Dim oSubMatches As RegExpExSubMatches
    Dim iSub As Long
    Dim iSubCount
    
    '// �����Ώە�����
    s = "abcd1234efgh5678ijkl9012"
    
    '// ���������ݒ�
    reg.Global_ = True              ' �����͈́iTrue�F������̍Ō�܂Ō����AFalse�F�ŏ��̈�v�܂Ō����j
    reg.IgnoreCase = True           ' �啶���������̋�ʁiTrue�F��ʂ��Ȃ��AFalse�F��ʂ���j
    reg.Pattern = "([a-z]+)(\d+)"   ' �����p�^�[���i�����ł͘A�����鐔�������������ɐݒ�j
    
    '// �������s
    Set oMatches = reg.Execute(s)
    
    '// ������v�������擾
    iCount = oMatches.Count
    
    '// ������v�����������[�v
    For i = 0 To iCount - 1
        '// �R���N�V�����̌����[�v�I�u�W�F�N�g���擾
        Set oMatch = oMatches.Item(i)
        
        '// �ŏ��̌�����v�ʒu
        iFirstIndex = oMatch.FirstIndex
        '// ������v������̒���
        iLength = oMatch.Length
        '// ������v������
        sValue = oMatch.Value
        
        Debug.Print "�ŏ�������v�ʒu�F" & iFirstIndex & " �����F" & iLength & " ������F" & sValue
        
        '// ������v
        Set oSubMatches = oMatch.SubMatches
        
        '// �T�u�\���i�ۊ��ʂň͂��Ă��錟�������j����Ă��鐔���擾
        iSubCount = oSubMatches.Count
        
        '// �T�u�\���̐��������[�v
        For iSub = 0 To iSubCount - 1
            Debug.Print "�T�u�\����v������F" & oSubMatches.Item(iSub)
        Next
    Next
End Sub


