VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptingExDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Scripting.Drivesクラスのラップ・拡張クラス
'* [詳  細] Scripting.DrivesクラスのWrapperとしての機能を提供する他、
'*          Scripting.Drivesクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Scripting.Drivesクラスについて。
' ドライブ文字に関連付けられた、ドライブのコレクション。
' 利用可能なすべてのドライブの読み取り専用のコレクションです。
' <https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/drives-collection>
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As Scripting.Drives
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
    Dim myDrive As Object
    For Each myDrive In mOrigin
        Dim wrap As ScriptingExDrive: Set wrap = New ScriptingExDrive
        Set wrap.Origin = myDrive
        mKeys.Add wrap, myDrive.DriveLetter
    Next myDrive
End Property

'*-----------------------------------------------------------------------------
'* Count プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ドライブの数。
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property

'*-----------------------------------------------------------------------------
'* Item プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・ドライブを取得する。
'* ・Set wrap.Origin = mOrigin.Item(Key) とすると以下のエラーが発生。
'*   実行時エラー '5': プロシージャの呼び出し、または引数が不正です。
'*
'*   DrivesオブジェクトのItemメソッドは、参照設定せずに使用する場合、
'*   上記のように引数に変数を使用するとエラーが発生する模様。
'*   理由は不明。固定値であれば動作するため、固定値を渡すように対処した。
'*-----------------------------------------------------------------------------
Property Get Item(Key) As ScriptingExDrive
Attribute Item.VB_UserMemId = 0
    Dim wrap As ScriptingExDrive: Set wrap = New ScriptingExDrive
    'Set wrap.Origin = mOrigin.Item(Key)
    Select Case Key
        Case "A", "a", "A:", "a:"
            Set wrap.Origin = mOrigin.Item("A:")
        Case "B", "b", "A:", "a:"
            Set wrap.Origin = mOrigin.Item("B:")
        Case "C", "c", "C:", "c:"
            Set wrap.Origin = mOrigin.Item("C:")
        Case "D", "d", "D:", "d:"
            Set wrap.Origin = mOrigin.Item("D:")
        Case "E", "e", "E:", "e:"
            Set wrap.Origin = mOrigin.Item("E:")
        Case "F", "f", "F:", "f:"
            Set wrap.Origin = mOrigin.Item("F:")
        Case "G", "g", "G:", "g:"
            Set wrap.Origin = mOrigin.Item("G:")
        Case "H", "h", "H:", "h:"
            Set wrap.Origin = mOrigin.Item("H:")
        Case "I", "i", "I:", "i:"
            Set wrap.Origin = mOrigin.Item("I:")
        Case "J", "j", "J:", "j:"
            Set wrap.Origin = mOrigin.Item("J:")
        Case "K", "k", "K:", "k:"
            Set wrap.Origin = mOrigin.Item("K:")
        Case "L", "l", "L:", "l:"
            Set wrap.Origin = mOrigin.Item("L:")
        Case "M", "m", "M:", "m:"
            Set wrap.Origin = mOrigin.Item("M:")
        Case "N", "n", "N:", "n:"
            Set wrap.Origin = mOrigin.Item("N:")
        Case "O", "o", "O:", "o:"
            Set wrap.Origin = mOrigin.Item("O:")
        Case "P", "p", "P:", "p:"
            Set wrap.Origin = mOrigin.Item("P:")
        Case "Q", "q", "Q:", "q:"
            Set wrap.Origin = mOrigin.Item("Q:")
        Case "R", "r", "R:", "r:"
            Set wrap.Origin = mOrigin.Item("R:")
        Case "S", "s", "S:", "s:"
            Set wrap.Origin = mOrigin.Item("S:")
        Case "T", "t", "T:", "t:"
            Set wrap.Origin = mOrigin.Item("T:")
        Case "U", "u", "U:", "u:"
            Set wrap.Origin = mOrigin.Item("U:")
        Case "V", "v", "V:", "V:"
            Set wrap.Origin = mOrigin.Item("V:")
        Case "W", "w", "W:", "w:"
            Set wrap.Origin = mOrigin.Item("W:")
        Case "X", "x", "X:", "x:"
            Set wrap.Origin = mOrigin.Item("X:")
        Case "Y", "y", "Y:", "y:"
            Set wrap.Origin = mOrigin.Item("Y:")
        Case "Z", "z", "Z:", "z:"
            Set wrap.Origin = mOrigin.Item("Z:")
    End Select
    Set Item = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Scripting.Drives")
    'Set mKeys = New Collection
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
    Set mKeys = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************
'なし

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
