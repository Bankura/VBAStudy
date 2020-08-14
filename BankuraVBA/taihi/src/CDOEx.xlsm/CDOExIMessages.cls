VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDOExIMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CDO.IMessagesクラスのラップ・拡張クラス
'* [詳  細] CDO.IMessagesクラスのWrapperとしての機能を提供する他、
'*          CDO.IMessagesクラスを使用したユーティリティを提供する。
'*
'* [参  考] <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' CDO.IMessagesクラスについて。
' <xx>
' Defines methods and properties used to manage a collection of Message objects on the file system. Returned by IDropDirectory.GetMessages.
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As CDO.IMessages
Private mOrigin As Object
Private mKeys As Collection

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
    Set mKeys = New Collection
    Dim myItem As Object
    For Each myItem In mOrigin
        Dim wrap As CDOExMessage: Set wrap = New CDOExMessage
        Set wrap.Origin = myItem
        mKeys.Add wrap, wrap
    Next
End Property

'*-----------------------------------------------------------------------------
'* Count プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The number of message objects in the collection.
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property

'*-----------------------------------------------------------------------------
'* FileName プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Returns the name of the file containing the specified message in the file system.
'*-----------------------------------------------------------------------------
Property Get FileName(var) As String
    FileName = mOrigin.FileName(var)
End Property

'*-----------------------------------------------------------------------------
'* Item プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・Returns the message specified by index from the collection.
'*-----------------------------------------------------------------------------
Property Get Item(Index As Long) As CDOExMessage
Attribute Item.VB_UserMemId = 0
    Dim wrap As CDOExMessage: Set wrap = New CDOExMessage
    Set wrap.Origin = mOrigin.Item(Index)
    Set Item = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("CDO.IMessages")
    'Set mKeys = New Collection
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
    Set mKeys = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************
'******************************************************************************
'* [概  要] Delete メソッド
'* [詳  細] Delete のラッパーメソッド。
'*          Deletes the specified message object in the collection.
'*
'* @param Index xx
'*
'******************************************************************************
Sub Delete(Index As Long)
    Call mOrigin.Delete(Index)

    Dim i As Long
    For i = 1 To mKeys.Count
        If mKeys(i) = Index Then
            mKeys.Remove i
            Exit Sub
        End If
    Next
End Sub

'******************************************************************************
'* [概  要] DeleteAll メソッド
'* [詳  細] DeleteAll のラッパーメソッド。
'*          Deletes all message objects in the collection.
'*
'******************************************************************************
Sub DeleteAll()
    Call mOrigin.DeleteAll
    Set mKeys = Nothing
    Set mKeys = New Collection
End Sub

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] NewEnum メソッド
'* [詳  細] For Each 構文でイテレーターを返すためのメソッド。
'*
'* @return イテレーター
'******************************************************************************
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mKeys.[_NewEnum]
End Function

