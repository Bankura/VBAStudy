VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] HttpResponse�N���X
'* [��  ��] HttpResponse���̃f�[�^�N���X�B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mStatusCd As String
Private mStatusTxt As String
Private mBody As String
Private mHeaders As String

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************
'Status Code
Public Property Get StatusCd() As String
    StatusCd = mStatusCd
End Property
Public Property Let StatusCd(arg As String)
    mStatusCd = arg
End Property

'Status Text
Public Property Get StatusText() As String
    StatusText = mStatusTxt
End Property
Public Property Let StatusText(arg As String)
    mStatusTxt = arg
End Property

'Response Body
Public Property Get Body() As String
    Body = mBody
End Property
Public Property Let Body(arg As String)
    mBody = arg
End Property

'Response Headers
Public Property Get Headers() As String
    Headers = mHeaders
End Property
Public Property Let Headers(arg As String)
    mHeaders = arg
End Property
