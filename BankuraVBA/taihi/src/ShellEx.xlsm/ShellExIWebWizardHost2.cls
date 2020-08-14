VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellExIWebWizardHost2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Shell32.IWebWizardHost2クラスのラップ・拡張クラス
'* [詳  細] Shell32.IWebWizardHost2クラスのWrapperとしての機能を提供する他、
'*          Shell32.IWebWizardHost2クラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Shell32.IWebWizardHost2クラスについて。
' <xxxxxxxxxxxxxx>
'
' IWebWizardHost2 interface
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mOrigin As shell32.IWebWizardHost2
'Private mOrigin As Object

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
'* Caption プロパティ
'*
'* [補  足]
'* ・
'*-----------------------------------------------------------------------------
Property Get Caption() As String
    Caption = mOrigin.Caption
End Property
Property Let Caption(arg As String)
    mOrigin.Caption = arg
End Property

'*-----------------------------------------------------------------------------
'* Property プロパティ
'*
'* [補  足]
'* ・
'*-----------------------------------------------------------------------------
Property Get Property(bstrPropertyName As String) As Variant
    If IsObject(mOrigin.Property(bstrPropertyName)) Then
        Set Property = mOrigin.Property(bstrPropertyName)
    Else
        Property = mOrigin.Property(bstrPropertyName)
    End If
End Property
Property Set Property(bstrPropertyName As String, arg As Variant)
    If IsObject(arg) Then
        Set mOrigin.Property(bstrPropertyName) = arg
    Else
        mOrigin.Property(bstrPropertyName) = arg
    End If
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Shell32.IWebWizardHost2")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Cancel メソッド
'* [詳  細] Cancel のラッパーメソッド。
'*
'******************************************************************************
Sub Cancel()
    Call mOrigin.Cancel
End Sub

'******************************************************************************
'* [概  要] FinalBack メソッド（既定メンバー）
'* [詳  細] FinalBack のラッパーメソッド。
'*
'******************************************************************************
Sub FinalBack()
Attribute FinalBack.VB_UserMemId = 0
    Call mOrigin.FinalBack
End Sub

'******************************************************************************
'* [概  要] FinalNext メソッド
'* [詳  細] FinalNext のラッパーメソッド。
'*
'******************************************************************************
Sub FinalNext()
    Call mOrigin.FinalNext
End Sub

'******************************************************************************
'* [概  要] SetHeaderText メソッド
'* [詳  細] SetHeaderText のラッパーメソッド。
'*
'* @param bstrHeaderTitle
'* @param bstrHeaderSubtitle
'******************************************************************************
Sub SetHeaderText(bstrHeaderTitle As String, bstrHeaderSubtitle As String)
    Call mOrigin.SetHeaderText(bstrHeaderTitle, bstrHeaderSubtitle)
End Sub

'******************************************************************************
'* [概  要] SetWizardButtons メソッド
'* [詳  細] SetWizardButtons のラッパーメソッド。
'*
'* @param vfEnableBack
'* @param vfEnableNext
'* @param vfLastPage
'******************************************************************************
Sub SetWizardButtons(vfEnableBack As Boolean, vfEnableNext As Boolean, vfLastPage As Boolean)
    Call mOrigin.SetWizardButtons(vfEnableBack, vfEnableNext, vfLastPage)
End Sub

'******************************************************************************
'* [概  要] SignString メソッド
'* [詳  細] SignString のラッパーメソッド。
'*
'* @param Value
'* @return String
'******************************************************************************
Function SignString(Value As String) As String
    SignString = mOrigin.SignString(Value)
End Function
