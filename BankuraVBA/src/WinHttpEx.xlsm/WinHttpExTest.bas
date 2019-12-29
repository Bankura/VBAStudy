Attribute VB_Name = "WinHttpExTest"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WinHttpExテスト用モジュール
'* [詳  細] WinHttpExテスト用モジュール
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* [概  要] WinHttpExWinHttpRequest のTest。
'* [詳  細] WinHttpExWinHttpRequest のTest用処理。
'*
'******************************************************************************
Sub WinHttpExWinHttpRequestTest()

    Dim whr As WinHttpExWinHttpRequest
    Set whr = New WinHttpExWinHttpRequest
    
    ' HTTP リクエストの送信
    'whr.OpenConn "GET", "http://www.google.com/", False
    'whr.Send
    
    ' ステータスとレスポンスボディの取得
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
'* [概  要] RequestDto のTest。
'* [詳  細] RequestDto のTest用処理。
'*
'******************************************************************************
Sub RequestDtoTest()

    Dim dto As RequestDto
    Set dto = New RequestDto
    dto.Url = "http://localhost:8080/test"
    dto.SetRequestParam "param1", "value1"
    dto.SetRequestParam "param2", "valueの2"
    Dim result As String
    result = dto.Url
    
    Debug.Print result
End Sub

