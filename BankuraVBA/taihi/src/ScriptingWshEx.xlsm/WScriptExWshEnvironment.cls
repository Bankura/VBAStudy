VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WScriptExWshEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WScript.WshEnvironmentクラスのラップ・拡張クラス
'* [詳  細] WScript.WshEnvironmentクラスのWrapperとしての機能を提供する他、
'*          WScript.WshEnvironmentクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WScript.WshEnvironmentクラスについて。
' Windows 環境変数のコレクションへのアクセスを提供します。
' WshEnvironment オブジェクトは環境変数のコレクションであり、WshShell オブジェ
' クトの Environment プロパティによって返されます。
' このコレクションには、環境変数セット全体 (名前付き変数および名前なし変数)
' が含まれます。
' 個々の環境変数 (およびその値) をこのコレクションから取得するには、環境変数
' 名をインデックスとして使用します。
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364435(v=msdn.10)>。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As IWshRuntimeLibrary.WshEnvironment
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
    Dim v
    For Each v In mOrigin
        'mKeys.Add Split(v, "=")(1), Split(v, "=")(0)
        mKeys.Add v
    Next
End Property

'*-----------------------------------------------------------------------------
'* Item プロパティ（既定メンバー）
'*
'* [補  足]
'* ・指定した名称に対する項目を設定。
'*-----------------------------------------------------------------------------
Property Get Item(Name As String) As String
Attribute Item.VB_UserMemId = 0
    Item = mOrigin.Item(Name)
End Property
Property Let Item(Name As String, arg As String)
     mOrigin.Item(Name) = arg
     'mKeys.Add arg, Name
     mKeys.Add Name & "=" & arg
End Property

'*-----------------------------------------------------------------------------
'* length プロパティ（読み取り専用）
'*
'* [補  足]
'* ・コレクション内のオブジェクト数を示す。
'*-----------------------------------------------------------------------------
Property Get length() As Long
    length = mOrigin.length
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WScript.Environment")
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
'* [概  要] Countメソッド
'* [詳  細] Countのラッパーメソッド。
'*          コレクションの要素数を返却。
'*
'* @return コレクションの要素数
'*
'******************************************************************************
Function Count() As Long
    Count = mOrigin.Count
End Function

'******************************************************************************
'* [概  要] Removeメソッド
'* [詳  細] Removeのラッパーメソッド。
'*          コレクションの要素を除外。
'*
'* @param Name コレクション内から除外する要素の名称
'*
'******************************************************************************
Sub Remove(Name As String)
    Call mOrigin.Remove(Name)
    Dim i As Long
    For i = 1 To mKeys.Count
        'If mKeys(i) = Name Then
        If Split(mKeys(i), "=")(0) = Name Then
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
    'Set NewEnum = mOrigin.[_NewEnum]
    Set NewEnum = mKeys.[_NewEnum]
End Function
