VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemQualifierSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemQualifierSetクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemQualifierSetクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemQualifierSetクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemQualifierSetクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393818(v=vs.80)>
' A collection of Qualifiers
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemQualifierSet
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
        Dim wrap As WMIExSWbemQualifier: Set wrap = New WMIExSWbemQualifier
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
    'Set mOrigin = CreateObject("WbemScripting.SWbemQualifierSet")
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
'*          Add a Qualifier to this collection.
'*
'* @param strName xx
'* @param varVal xx
'* @param bPropagatesToSubclass xx
'* @param bPropagatesToInstance xx
'* @param bIsOverridable xx
'* @param iFlags xx
'* @return WMIExSWbemQualifier
'*
'******************************************************************************
Function Add(strName As String, _
             varVal, _
             Optional bPropagatesToSubclass As Boolean = True, _
             Optional bPropagatesToInstance As Boolean = True, _
             Optional bIsOverridable As Boolean = True, _
             Optional iFlags As Long) As WMIExSWbemQualifier
    Dim wrap As WMIExSWbemQualifier: Set wrap = New WMIExSWbemQualifier
    Set wrap.Origin = mOrigin.Add(strName, varVal, bPropagatesToSubclass, bPropagatesToInstance, bIsOverridable, iFlags)
    Set Add = wrap
    mKeys.Add wrap, strName
End Function
    
'******************************************************************************
'* [概  要] Item メソッド（既定メンバー）
'* [詳  細] Item のラッパーメソッド。
'*          Get a named Qualifier from this collection.
'*
'* @param Name xx
'* @param iFlags xx
'* @return WMIExSWbemProperty
'*
'******************************************************************************
Function Item(Name As String, Optional iFlags As Long) As WMIExSWbemQualifier
Attribute Item.VB_UserMemId = 0
    Dim wrap As WMIExSWbemQualifier: Set wrap = New WMIExSWbemQualifier
    Set wrap.Origin = mOrigin.Item(Name, iFlags)
    Set Item = wrap
End Function
 
'******************************************************************************
'* [概  要] Remove メソッド
'* [詳  細] Remove のラッパーメソッド。
'*          Remove a Qualifier from this collection.
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

