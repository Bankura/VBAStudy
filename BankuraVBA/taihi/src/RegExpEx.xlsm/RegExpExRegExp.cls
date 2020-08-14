VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegExpExRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] VBScript.RegExpクラスのラップ・拡張クラス
'* [詳  細] VBScript.RegExpクラスのWrapperとしての機能を提供する他、
'*          VBScript.RegExpクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' VBScript.RegExpクラスについて。
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc392403(v=msdn.10)?redirectedfrom=MSDN>
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As VBScript_RegExp_55.RegExp
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
'* Global_ プロパティ
'*
'* [補  足]
'* ・｡
'*-----------------------------------------------------------------------------
Property Get Global_() As Boolean
    Global_ = mOrigin.Global
End Property
Property Let Global_(arg As Boolean)
    mOrigin.Global = arg
End Property

'*-----------------------------------------------------------------------------
'* IgnoreCase プロパティ
'*
'* [補  足]
'* ・｡
'*-----------------------------------------------------------------------------
Property Get IgnoreCase() As Boolean
    IgnoreCase = mOrigin.IgnoreCase
End Property
Property Let IgnoreCase(arg As Boolean)
    mOrigin.IgnoreCase = arg
End Property

'*-----------------------------------------------------------------------------
'* Multiline プロパティ
'*
'* [補  足]
'* ・｡
'*-----------------------------------------------------------------------------
Property Get Multiline() As Boolean
    Multiline = mOrigin.Multiline
End Property
Property Let Multiline(arg As Boolean)
    mOrigin.Multiline = arg
End Property

'*-----------------------------------------------------------------------------
'* Pattern プロパティ
'*
'* [補  足]
'* ・｡
'*-----------------------------------------------------------------------------
Property Get Pattern() As String
    Pattern = mOrigin.Pattern
End Property
Property Let Pattern(arg As String)
    mOrigin.Pattern = arg
End Property
    
'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("VBScript.RegExp")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Execute メソッド
'* [詳  細] Execute のラッパーメソッド。
'*
'*
'* @param sourceString xx
'* @return Object
'*
'******************************************************************************
Function Execute(sourceString As String) As RegExpExMatchCollection
    Dim wrap As RegExpExMatchCollection: Set wrap = New RegExpExMatchCollection
    Set wrap.Origin = mOrigin.Execute(sourceString)
    Set Execute = wrap
End Function

'******************************************************************************
'* [概  要] Replace メソッド
'* [詳  細] Replace のラッパーメソッド。
'*
'*
'* @param sourceString xx
'* @param replaceVar xx
'* @return Object
'*
'******************************************************************************
Function Replace(sourceString As String, replaceVar) As String
    Replace = mOrigin.Replace(sourceString, replaceVar)
End Function

'******************************************************************************
'* [概  要] Test メソッド
'* [詳  細] Test のラッパーメソッド。
'*
'*
'* @param sourceString xx
'* @return Boolean
'*
'******************************************************************************
Function Test(sourceString As String) As Boolean
    Test = mOrigin.Test(sourceString)
End Function
    
'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------
