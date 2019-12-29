VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemPropertySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemPropertySetクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemPropertySetクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemPropertySetクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemPropertySetクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393805(v=vs.80)>
' A collection of Properties
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemPropertySet
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
        Dim wrap As WMIExSWbemProperty: Set wrap = New WMIExSWbemProperty
        Set wrap.Original = myItem
        mKeys.Add wrap, myItem.Name
    Next myItem
End Property

'*-----------------------------------------------------------------------------
'* Count プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The number of items in this collection.
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemPropertySet")
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
'*          Add a Property to this collection.
'*
'* @param strName xx
'* @param iCimType xx
'* @param bIsArray xx
'* @param iFlags xx
'* @return WMIExSWbemProperty
'*
'******************************************************************************
Function Add(strName As String, iCimType As WbemCimtypeEnum, Optional bIsArray As Boolean = False, Optional iFlags As Long) As WMIExSWbemProperty
    Dim wrap As WMIExSWbemProperty: Set wrap = New WMIExSWbemProperty
    Set wrap.Original = mOrigin.Add(strName, iCimType, bIsArray, iFlags)
    Set Add = wrap
    mKeys.Add wrap, strName
End Function

'******************************************************************************
'* [概  要] Item メソッド（既定メンバー）
'* [詳  細] Item のラッパーメソッド。
'*          Get a named Property from this collection.
'*
'* @param strName xx
'* @param iFlags xx
'* @return WMIExSWbemProperty
'*
'******************************************************************************
Function Item(strName As String, Optional iFlags As Long) As WMIExSWbemProperty
Attribute Item.VB_UserMemId = 0
    Dim wrap As WMIExSWbemProperty: Set wrap = New WMIExSWbemProperty
    Set wrap.Original = mOrigin.Item(strName, iFlags)
    Set Item = wrap
End Function

'******************************************************************************
'* [概  要] Remove メソッド
'* [詳  細] Remove のラッパーメソッド。
'*          Remove a Property from this collection.
'*
'* @param strName xx
'* @param iFlags xx
'*
'******************************************************************************
Sub Remove(strName As String, Optional iFlags As Long)
    Call mOrigin.Remove(strName, iFlags)
    
    Dim i As Long
    For i = 1 To mKeys.Count
        If mKeys(i) = strName Then
            mKeys.Remove i
            Exit Sub
        End If
    Next
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

