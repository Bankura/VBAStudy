Attribute VB_Name = "GoogleSearchMain"
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] Selenium�g�pGoogle�������W���[��
'* [��  ��] Google�������s�����W���[���B
'*
'* @author Bankura
'* Copyright (c) 2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* WindowsAPI��`
'******************************************************************************
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'******************************************************************************
'* �ϐ���`
'******************************************************************************

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] GoogleSearchMain ���\�b�h
'* [��  ��] �e�X�g���\�b�h�B
'******************************************************************************
Sub GoogleSearchMain()
    Dim gSearch As New GoogleSearch
    
    ' ���O�`�F�b�N���s
    If Not gSearch.WebDriver.ChromePreCheck() Then
        Debug.Print "���O�`�F�b�N���s"
        Exit Sub
    End If
    
    ' �����ݒ�
    gSearch.SearchWord = "�{�D���̉�����`�i���ɂȂ邽�߂ɂ͎�i��I��ł����܂���` ���m�̖�I"
    gSearch.GoogleSearchType = GSEARCH_BOOK
    gSearch.UseFilter = True
    gSearch.MaxSearchCount = 20
    
    ' Chrome�N��
    gSearch.Go

    ' �����ƕ\��
    gSearch.SearchAndShow
    
    'gSearch.WebDriver.CloseWindow
End Sub
