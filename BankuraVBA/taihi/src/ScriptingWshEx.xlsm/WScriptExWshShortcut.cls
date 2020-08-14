VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WScriptExWshShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WScript.WshShortcutクラスのラップ・拡張クラス
'* [詳  細] WScript.WshShortcutクラスのWrapperとしての機能を提供する他、
'*          WScript.WshShortcutクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WScript.WshShortcutクラスについて。
' ショートカットへのオブジェクト参照を作成します。
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364438(v=msdn.10)>
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As IWshRuntimeLibrary.WshShortcut
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
'* Arguments プロパティ
'*
'* [補  足]
'* ・ Arguments 。
'*-----------------------------------------------------------------------------
Property Get Arguments() As String
    Arguments = mOrigin.Arguments
End Property
Property Let Arguments(arg As String)
    mOrigin.Arguments = arg
End Property

'*-----------------------------------------------------------------------------
'* Description プロパティ
'*
'* [補  足]
'* ・ショートカットの説明を返します。
'*-----------------------------------------------------------------------------
Property Get Description() As String
    Description = mOrigin.Description
End Property
Property Let Description(arg As String)
    mOrigin.Description = arg
End Property

'* FullName プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ショートカット オブジェクトのリンク先への絶対パスを返します。
'*-----------------------------------------------------------------------------
Property Get FullName() As String
    FullName = mOrigin.FullName
End Property

'*-----------------------------------------------------------------------------
'* Hotkey プロパティ
'*
'* [補  足]
'* ・キーの組み合わせをショートカットに割り当てるか、割り当てられたキーの
'*   組み合わせを識別します。
'*-----------------------------------------------------------------------------
Property Get Hotkey() As String
    Hotkey = mOrigin.Hotkey
End Property
Property Let Hotkey(arg As String)
    mOrigin.Hotkey = arg
End Property

'*-----------------------------------------------------------------------------
'* IconLocation プロパティ
'*
'* [補  足]
'* ・アイコンをショートカットに割り当てるか、割り当てられたアイコンを識別します。
'*-----------------------------------------------------------------------------
Property Get IconLocation() As String
    IconLocation = mOrigin.IconLocation
End Property
Property Let IconLocation(arg As String)
    mOrigin.IconLocation = arg
End Property

'*-----------------------------------------------------------------------------
'* RelativePath プロパティ
'*
'* [補  足]
'* ・ RelativePath 。
'*-----------------------------------------------------------------------------
Property Get RelativePath() As String
    RelativePath = mOrigin.RelativePath
End Property
Property Let RelativePath(arg As String)
    mOrigin.RelativePath = arg
End Property

'*-----------------------------------------------------------------------------
'* TargetPath プロパティ
'*
'* [補  足]
'* ・ショートカットの実行可能ファイルへのパスです。
'*-----------------------------------------------------------------------------
Property Get TargetPath() As String
    TargetPath = mOrigin.TargetPath
End Property
Property Let TargetPath(arg As String)
    mOrigin.TargetPath = arg
End Property

'*-----------------------------------------------------------------------------
'* WindowStyle プロパティ
'*
'* [補  足]
'* ・ウィンドウスタイルをショートカットに割り当てるか、ショートカットの
'*   ウィンドウ スタイルの種類を識別します。
'*     1   ウィンドウをアクティブにして表示します｡
'*         ウィンドウが最小化または最大化されている場合は､元のサイズと位置に戻ります｡
'*     3   ウィンドウをアクティブにし､最大化ウィンドウとして表示します｡
'*     7   ウィンドウを最小化し､次に上位となるウィンドウをアクティブにします｡
'*
'*-----------------------------------------------------------------------------
Property Get WindowStyle() As Long
    WindowStyle = mOrigin.WindowStyle
End Property
Property Let WindowStyle(arg As Long)
    mOrigin.WindowStyle = arg
End Property

'*-----------------------------------------------------------------------------
'* WorkingDirectory プロパティ
'*
'* [補  足]
'* ・作業ディレクトリをショートカットに割り当てるか、ショートカットの作業
'*   ディレクトリを識別します。
'*-----------------------------------------------------------------------------
Property Get WorkingDirectory() As String
    WorkingDirectory = mOrigin.WorkingDirectory
End Property
Property Let WorkingDirectory(arg As String)
    mOrigin.WorkingDirectory = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WScript.WshShortcut")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Save メソッド
'* [詳  細] Save のラッパーメソッド。
'*          ショートカット オブジェクトを保存します。
'*
'******************************************************************************
Sub Save()
    Call mOrigin.Save
End Sub

'******************************************************************************
'* [概  要] Load メソッド
'* [詳  細] Load のラッパーメソッド。
'*          隠しメソッド。
'*
'* @param PathLink
'******************************************************************************
Sub Load(PathLink As String)
    Call mOrigin.Load(PathLink)
End Sub
