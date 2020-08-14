VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemRefreshableItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemRefreshableItemクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemRefreshableItemクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemRefreshableItemクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemRefreshableItemクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393831(v=vs.80)>
' A single item from a Refresher
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemRefreshableItem
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
'* Index プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The index of this item in the parent refresher.
'*-----------------------------------------------------------------------------
Property Get Index() As Long
    Index = mOrigin.Index
End Property

'*-----------------------------------------------------------------------------
'* IsSet プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Whether this item represents a single object or an object set.
'*-----------------------------------------------------------------------------
Property Get IsSet() As Boolean
    IsSet = mOrigin.IsSet
End Property
    
'*-----------------------------------------------------------------------------
'* Object プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The Object.
'*-----------------------------------------------------------------------------
Property Get Object() As WMIExSWbemObjectEx
    Dim wrap As WMIExSWbemObjectEx: Set wrap = New WMIExSWbemObjectEx
    Set wrap.Origin = mOrigin.Object
    Set Object = wrap
End Property
    
'*-----------------------------------------------------------------------------
'* ObjectSet プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The object set.
'*-----------------------------------------------------------------------------
Property Get ObjectSet() As WMIExSWbemObjectSet
    Dim wrap As WMIExSWbemObjectSet: Set wrap = New WMIExSWbemObjectSet
    Set wrap.Origin = mOrigin.ObjectSet
    Set ObjectSet = wrap
End Property
    
'*-----------------------------------------------------------------------------
'* Refresher プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The parent refresher.
'*-----------------------------------------------------------------------------
Property Get Refresher() As WMIExSWbemRefresher
    Dim wrap As WMIExSWbemRefresher: Set wrap = New WMIExSWbemRefresher
    Set wrap.Origin = mOrigin.Refresher
    Set Refresher = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemRefreshableItem")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Remove メソッド
'* [詳  細] Remove のラッパーメソッド。
'*          Remove this item from the parent refresher.
'*
'* @param iFlags xx
'*
'******************************************************************************
Sub Remove(Optional iFlags As Long)
    Call mOrigin.Remove(iFlags)
End Sub

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------



