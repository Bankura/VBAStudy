VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemNamedValueSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemNamedValueSetクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemNamedValueSetクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemNamedValueSetクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemNamedValueSetクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393732(v=vs.80)>
' A collection of Named Values
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemNamedValueSet
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
        Dim wrap As WMIExSWbemNamedValue: Set wrap = New WMIExSWbemNamedValue
        Set wrap.Origin = myItem
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
    Set mOrigin = CreateObject("WbemScripting.SWbemNamedValueSet")
    Set mKeys = New Collection
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
'*          Add a named value to this collection.
'*
'* @param strName xx
'* @param varValue xx
'* @param iFlags xx
'* @return WMIExSWbemNamedValue
'*
'******************************************************************************
Function Add(strName As String, varValue, Optional iFlags As Long) As WMIExSWbemNamedValue
    Dim wrap As WMIExSWbemNamedValue: Set wrap = New WMIExSWbemNamedValue
    Set wrap.Origin = mOrigin.Add(strName, varValue, iFlags)
    Set Add = wrap
    mKeys.Add wrap, strName
End Function

'******************************************************************************
'* [概  要] Clone メソッド
'* [詳  細] Clone のラッパーメソッド。
'*          Make a copy of this collection.
'*
'* @return WMIExSWbemNamedValueSet
'*
'******************************************************************************
Function Clone() As WMIExSWbemNamedValueSet
    Dim wrap As WMIExSWbemNamedValueSet: Set wrap = New WMIExSWbemNamedValueSet
    Set wrap.Origin = mOrigin.Clone()
    Set Clone = wrap
End Function

'******************************************************************************
'* [概  要] DeleteAll メソッド
'* [詳  細] DeleteAll のラッパーメソッド。
'*          Delete all items in this collection.
'*
'******************************************************************************
Sub DeleteAll()
    Call mOrigin.DeleteAll
    Set mKeys = Nothing
    Set mKeys = New Collection
End Sub

'******************************************************************************
'* [概  要] Item メソッド（既定メンバー）
'* [詳  細] Item のラッパーメソッド。
'*          Get a named value from this Collection.
'*
'* @param strName xx
'* @param iFlags xx
'* @return WMIExSWbemNamedValue
'*
'******************************************************************************
Function Item(strName As String, Optional iFlags As Long) As WMIExSWbemNamedValue
Attribute Item.VB_UserMemId = 0
    Dim wrap As WMIExSWbemNamedValue: Set wrap = New WMIExSWbemNamedValue
    Set wrap.Origin = mOrigin.Item(strName, iFlags)
    Set Item = wrap
End Function

'******************************************************************************
'* [概  要] Remove メソッド
'* [詳  細] Remove のラッパーメソッド。
'*          Remove a named value from this collection.
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
