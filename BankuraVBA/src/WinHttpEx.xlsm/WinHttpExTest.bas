Attribute VB_Name = "WinHttpExTest"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WinHttpEx�e�X�g�p���W���[��
'* [��  ��] WinHttpEx�e�X�g�p���W���[��
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* [�T  �v] WinHttpExWinHttpRequest ��Test�B
'* [��  ��] WinHttpExWinHttpRequest ��Test�p�����B
'*
'******************************************************************************
Sub WinHttpExWinHttpRequestTest()

    Dim whr As WinHttpExWinHttpRequest
    Set whr = New WinHttpExWinHttpRequest
    
    ' HTTP ���N�G�X�g�̑��M
    'whr.OpenConn "GET", "http://www.google.com/", False
    'whr.Send
    
    ' �X�e�[�^�X�ƃ��X�|���X�{�f�B�̎擾
    'result = whr.Status & " " & _
    '         whr.StatusText & vbCrLf & _
    '         whr.GetAllResponseHeaders & vbCrLf & _
    '         whr.ResponseText
    
    Dim dto As RequestDto
    Dim response As ResponseDto
    Set dto = New RequestDto
    dto.Url = "http://www.google.com/"
    
    Set response = whr.GetReq(dto)
    With response
        Debug.Print .StatusCd & " " & _
                    .StatusText & vbCrLf & _
                    .Headers & vbCrLf & _
                    .Body
    End With
End Sub

'******************************************************************************
'* [�T  �v] RequestDto ��Test�B
'* [��  ��] RequestDto ��Test�p�����B
'*
'******************************************************************************
Sub RequestDtoTest()

    Dim dto As RequestDto
    Set dto = New RequestDto
    dto.Url = "http://localhost:8080/test"
    dto.SetRequestParam "param1", "value1"
    dto.SetRequestParam "param2", "value��2"
    Dim result As String
    result = dto.Url
    
    Debug.Print result
End Sub

