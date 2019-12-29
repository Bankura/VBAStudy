VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemRefresher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemRefresherクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemRefresherクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemRefresherクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemRefresherクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393838(v=vs.80)>
' Refresher
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemRefresher
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

'*-----------------------------------------------------------------------------
'* AutoReconnect プロパティ
'*
'* [補  足]
'* ・Whether to attempt auto-reconnection to a remote provider.
'*-----------------------------------------------------------------------------
Property Get AutoReconnect() As Boolean
    AutoReconnect = mOrigin.AutoReconnect
End Property
Property Let AutoReconnect(arg As Boolean)
    mOrigin.AutoReconnect = arg
End Property

'*-----------------------------------------------------------------------------
'* Count プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The number of items in this refresher.
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("WbemScripting.SWbemRefresher")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Add メソッド
'* [詳  細] Add のラッパーメソッド。
'*         Add a refreshable instance to this refresher.
'*
'* @param objWbemServices xx
'* @param bsInstancePath xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemRefreshableItem
'*
'******************************************************************************
Function Add(objWbemServices As WMIExSWbemServicesEx, bsInstancePath As String, Optional iFlags As Long, Optional objWbemNamedValueSet As Object) As WMIExSWbemRefreshableItem
    Dim wrap As WMIExSWbemRefreshableItem: Set wrap = New WMIExSWbemRefreshableItem
    Set wrap.Origin = mOrigin.Add(objWbemServices.Origin, bsInstancePath, iFlags, objWbemNamedValueSet)
    Set Add = wrap
End Function

'******************************************************************************
'* [概  要] AddEnum メソッド
'* [詳  細] AddEnum のラッパーメソッド。
'*          Add a refreshable enumerator to this refresher.
'*
'* @param objWbemServices xx
'* @param bsInstancePath xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemRefreshableItem
'*
'******************************************************************************
Function AddEnum(objWbemServices As WMIExSWbemServicesEx, bsClassName As String, Optional iFlags As Long, Optional objWbemNamedValueSet As Object) As WMIExSWbemRefreshableItem
    Dim wrap As WMIExSWbemRefreshableItem: Set wrap = New WMIExSWbemRefreshableItem
    Set wrap.Origin = mOrigin.AddEnum(objWbemServices.Origin, bsClassName, iFlags, objWbemNamedValueSet)
    Set AddEnum = wrap
End Function

'******************************************************************************
'* [概  要] DeleteAll メソッド
'* [詳  細] DeleteAll のラッパーメソッド。
'*          Delete all items in this collection.
'*
'******************************************************************************
Sub DeleteAll()
    Call mOrigin.DeleteAll
End Sub
    
'******************************************************************************
'* [概  要] Item メソッド（既定メンバー）
'* [詳  細] Item のラッパーメソッド。
'*          Get an item from this refresher.
'*
'* @param Name xx
'* @param iFlags xx
'* @return WMIExSWbemProperty
'*
'******************************************************************************
Function Item(iIndex As Long) As WMIExSWbemRefreshableItem
Attribute Item.VB_UserMemId = 0
    Dim wrap As WMIExSWbemRefreshableItem: Set wrap = New WMIExSWbemRefreshableItem
    Set wrap.Origin = mOrigin.Item(iIndex)
    Set Item = wrap
End Function

'******************************************************************************
'* [概  要] Remove メソッド
'* [詳  細] Remove のラッパーメソッド。
'*          Refresh all items in this collection.
'*
'* @param strName xx
'* @param iFlags xx
'*
'******************************************************************************
Sub Refresh(Optional iFlags As Long)
    Call mOrigin.Refresh(iFlags)
End Sub
    
'******************************************************************************
'* [概  要] Remove メソッド
'* [詳  細] Remove のラッパーメソッド。
'*          Remove an item from this refresher.
'*
'* @param iIndex xx
'* @param iFlags xx
'*
'******************************************************************************
Sub Remove(iIndex As Long, Optional iFlags As Long)
    Call mOrigin.Remove(iIndex, iFlags)
End Sub

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------
