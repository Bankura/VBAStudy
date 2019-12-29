VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDOExIBodyParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CDO.IBodyPartsクラスのラップ・拡張クラス
'* [詳  細] CDO.IBodyPartsクラスのWrapperとしての機能を提供する他、
'*          CDO.IBodyPartsクラスを使用したユーティリティを提供する。
'*
'* [参  考] <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' CDO.IBodyPartsクラスについて。
' <xx>
' Defines methods and properties used to manage a collection of BodyPart objects.
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As CDO.IBodyParts
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
        Dim wrap As CDOExIBodyPart: Set wrap = New CDOExIBodyPart
        Set wrap.Origin = myItem
        mKeys.Add wrap, wrap
    Next
End Property

'*-----------------------------------------------------------------------------
'* Count プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Returns the number of BodyPart objects in the collection.
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property

'*-----------------------------------------------------------------------------
'* Item プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・The specified BodyPart object in the collection.
'*-----------------------------------------------------------------------------
Property Get Item(Index As Long) As CDOExIBodyPart
Attribute Item.VB_UserMemId = 0
    Dim wrap As CDOExIBodyPart: Set wrap = New CDOExIBodyPart
    Set wrap.Origin = mOrigin.Item(Index)
    Set Item = wrap
End Property

    
'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("CDO.IBodyParts")
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
'* [概  要] Add メソッド
'* [詳  細] Add のラッパーメソッド。
'*          Adds a BodyPart object to the collection at the specified index, and returns the newly added object.
'*
'* @param Index
'* @return IBodyPart
'*
'******************************************************************************
Function Add(Optional Index As Long = -1) As CDOExIBodyPart
    Dim wrap As CDOExIBodyPart: Set wrap = New CDOExIBodyPart
    Set wrap.Origin = mOrigin.Add(Index)
    Set Add = wrap
    mKeys.Add wrap, wrap
End Function


'******************************************************************************
'* [概  要] Delete メソッド
'* [詳  細] Delete のラッパーメソッド。
'*          Deletes the specified BodyPart object from the collection. Can use the index or a reference to the object.
'*
'* @param varBP xx
'*
'******************************************************************************
Sub Delete(varBP)
    Call mOrigin.Delete(varBP)

    Dim i As Long
    For i = 1 To mKeys.Count
        If mKeys(i) = varBP Then
            mKeys.Remove i
            Exit Sub
        End If
    Next
End Sub

'******************************************************************************
'* [概  要] DeleteAll メソッド
'* [詳  細] DeleteAll のラッパーメソッド。
'*          Deletes all BodyPart objects in the collection.
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

