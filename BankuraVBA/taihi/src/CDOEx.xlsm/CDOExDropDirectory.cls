VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDOExDropDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CDO.DropDirectoryクラスのラップ・拡張クラス
'* [詳  細] CDO.DropDirectoryクラスのWrapperとしての機能を提供する他、
'*          CDO.DropDirectoryクラスを使用したユーティリティを提供する。
'*
'* [参  考] <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' CDO.DropDirectoryクラスについて。
' <xx>
' Defines an object used to access messages located on the file system.
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As CDO.DropDirectory
Private mOrigin As Object

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* オリジナルObject プロパティ
'*-----------------------------------------------------------------------------
Property Get Origin() As Object
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Object)
    Set mOrigin = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("CDO.DropDirectory")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] GetMessages メソッド
'* [詳  細] GetMessages のラッパーメソッド。
'*          Returns a collection of messages contained in the specified
'*          directory on the file system.
'*          The default location is the SMTP drop directory.
'*
'* @param DirName  xx
'* @return IMessages
'******************************************************************************
Function GetMessages(Optional DirName As String) As CDOExIMessages
    Dim wrap As CDOExIMessages: Set wrap = New CDOExIMessages
    Set wrap.Origin = mOrigin.GetMessages(DirName)
    Set GetMessages = wrap
End Function
