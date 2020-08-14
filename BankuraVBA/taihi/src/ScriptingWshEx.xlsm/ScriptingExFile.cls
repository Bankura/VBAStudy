VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptingExFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Scripting.Fileクラスのラップ・拡張クラス
'* [詳  細] Scripting.FileクラスのWrapperとしての機能を提供する他、
'*          Scripting.Fileクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Scripting.Fileクラスについて。
' <https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/file-object>
' ファイルのすべてのプロパティにアクセスできます。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As Scripting.file
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
'* Attributes プロパティ
'*
'* [補  足]
'* ・ファイルの属性。
'*-----------------------------------------------------------------------------
Property Get Attributes() As FileAttribute
    Attributes = mOrigin.Attributes
End Property
Property Let Attributes(arg As FileAttribute)
    mOrigin.Attributes = arg
End Property

'*-----------------------------------------------------------------------------
'* DateCreated プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ファイル作成日。
'*-----------------------------------------------------------------------------
Property Get DateCreated() As Date
    DateCreated = mOrigin.DateCreated
End Property

'*-----------------------------------------------------------------------------
'* DateLastAccessed プロパティ（読み取り専用）
'*
'* [補  足]
'* ・最終ファイルアクセス日。
'*-----------------------------------------------------------------------------
Property Get DateLastAccessed() As Date
    DateLastAccessed = mOrigin.DateLastAccessed
End Property

'*-----------------------------------------------------------------------------
'* DateLastModified プロパティ（読み取り専用）
'*
'* [補  足]
'* ・最終ファイルアクセス日。
'*-----------------------------------------------------------------------------
Property Get DateLastModified() As Date
    DateLastModified = mOrigin.DateLastModified
End Property

'*-----------------------------------------------------------------------------
'* Drive プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ファイルを含むドライブを取得します。
'*-----------------------------------------------------------------------------
Property Get Drive() As ScriptingExDrive
    Dim wrap As ScriptingExDrive: Set wrap = New ScriptingExDrive
    Set wrap.Origin = mOrigin.Drive
    Set Drive = wrap
End Property

'*-----------------------------------------------------------------------------
'* Name プロパティ
'*
'* [補  足]
'* ・ファイル名を取得します。
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property
Property Let Name(arg As String)
    mOrigin.Name = arg
End Property

'*-----------------------------------------------------------------------------
'* ParentFolder プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ファイルを含むフォルダを取得します。
'*-----------------------------------------------------------------------------
Property Get ParentFolder() As ScriptingExFolder
    Dim wrap As ScriptingExFolder: Set wrap = New ScriptingExFolder
    Set wrap.Origin = mOrigin.ParentFolder
    Set ParentFolder = wrap
End Property

'*-----------------------------------------------------------------------------
'* Path プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・パス。
'*-----------------------------------------------------------------------------
Property Get Path() As String
Attribute Path.VB_UserMemId = 0
    Path = mOrigin.Path
End Property

'*-----------------------------------------------------------------------------
'* ShortName プロパティ（読み取り専用）
'*
'* [補  足]
'* ・短い名前。
'*-----------------------------------------------------------------------------
Property Get ShortName() As String
    ShortName = mOrigin.ShortName
End Property

'*-----------------------------------------------------------------------------
'* ShortPath プロパティ（読み取り専用）
'*
'* [補  足]
'* ・短いパス。
'*-----------------------------------------------------------------------------
Property Get ShortPath() As String
    ShortPath = mOrigin.ShortPath
End Property

'*-----------------------------------------------------------------------------
'* Size プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ファイルの容量。
'*-----------------------------------------------------------------------------
Property Get Size() As Variant
    Size = mOrigin.Size
End Property

'*-----------------------------------------------------------------------------
'* Type プロパティ（読み取り専用）
'*
'* [補  足]
'* ・タイプ記述 。
'*-----------------------------------------------------------------------------
Property Get Type_() As Variant
    Type_ = mOrigin.Type
End Property


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Scripting.File")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Copyメソッド
'* [詳  細] Copyのラッパーメソッド。
'*          指定したファイルをコピーします｡
'*
'* @param Destination
'* @param OverWriteFiles  省略可能。
'*
'******************************************************************************
Sub Copy(Destination As String, Optional OverWriteFiles As Boolean = True)
    Call mOrigin.Copy(Destination, OverWriteFiles)
End Sub

'******************************************************************************
'* [概  要] Deleteメソッド
'* [詳  細] Deleteのラッパーメソッド。
'*         指定したファイルを削除します｡
'*
'* @param Force 省略可能。
'*
'******************************************************************************
Sub Delete(Optional Force As Boolean = False)
    Call mOrigin.Delete(Force)
End Sub

'******************************************************************************
'* [概  要] Moveメソッド
'* [詳  細] Moveのラッパーメソッド。
'*          指定したファイルを移動します｡
'*
'* @param Destination
'*
'******************************************************************************
Sub Move(Destination As String)
    Call mOrigin.Move(Destination)
End Sub


'******************************************************************************
'* [概  要] OpenAsTextStreamメソッド
'* [詳  細] OpenAsTextStreamのラッパーメソッド。
'*          ファイルを TextStream オブジェクトとして開きます｡
'*
'* @param IOMode 省略可能。
'* @param Format 省略可能。
'* @return ScriptingExTextStream ファイルのストリーム
'*
'******************************************************************************
Function OpenAsTextStream(Optional IOMode As IOMode = ForReading, Optional Format As Tristate = TristateFalse) As ScriptingExTextStream
    Dim wrap As ScriptingExTextStream: Set wrap = New ScriptingExTextStream
    Set wrap.Origin = mOrigin.OpenAsTextStream(IOMode, Format)
    Set OpenAsTextStream = wrap
End Function
