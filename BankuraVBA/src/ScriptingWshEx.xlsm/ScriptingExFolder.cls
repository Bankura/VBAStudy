VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptingExFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Scripting.Folderクラスのラップ・拡張クラス
'* [詳  細] Scripting.FolderクラスのWrapperとしての機能を提供する他、
'*          Scripting.Folderクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Scripting.Folderクラスについて。
' <https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/folder-object>
' フォルダーのすべてのプロパティにアクセスできます。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As Scripting.Folder
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
'* ・フォルダの属性。
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
'* ・フォルダ作成日。
'*-----------------------------------------------------------------------------
Property Get DateCreated() As Date
    DateCreated = mOrigin.DateCreated
End Property

'*-----------------------------------------------------------------------------
'* DateLastAccessed プロパティ（読み取り専用）
'*
'* [補  足]
'* ・最終フォルダアクセス日。
'*-----------------------------------------------------------------------------
Property Get DateLastAccessed() As Date
    DateLastAccessed = mOrigin.DateLastAccessed
End Property

'*-----------------------------------------------------------------------------
'* DateLastModified プロパティ（読み取り専用）
'*
'* [補  足]
'* ・最終フォルダアクセス日。
'*-----------------------------------------------------------------------------
Property Get DateLastModified() As Date
    DateLastModified = mOrigin.DateLastModified
End Property

'*-----------------------------------------------------------------------------
'* Drive プロパティ（読み取り専用）
'*
'* [補  足]
'* ・フォルダを含むドライブを取得します。
'*-----------------------------------------------------------------------------
Property Get Drive() As ScriptingExDrive
    Dim wrap As ScriptingExDrive: Set wrap = New ScriptingExDrive
    Set wrap.Origin = mOrigin.Drive
    Set Drive = wrap
End Property

'*-----------------------------------------------------------------------------
'* Files プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ファイルのコレクションを取得します。
'*-----------------------------------------------------------------------------
Property Get files() As ScriptingExFiles
    Dim wrap As ScriptingExFiles: Set wrap = New ScriptingExFiles
    Set wrap.Origin = mOrigin.files
    Set files = wrap
End Property

'*-----------------------------------------------------------------------------
'* IsRootFolder プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ルート フォルダの場合には True を返します。
'*-----------------------------------------------------------------------------
Property Get IsRootFolder() As Boolean
    IsRootFolder = mOrigin.IsRootFolder
End Property

'*-----------------------------------------------------------------------------
'* Name プロパティ
'*
'* [補  足]
'* ・フォルダ名を取得します。
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
'* ・1 つ上のフォルダを取得します。
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
'* ・フォルダのパス。
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
'* ・ファイルおよびサブフォルダの容量。
'*-----------------------------------------------------------------------------
Property Get Size() As Variant
    Size = mOrigin.Size
End Property

'*-----------------------------------------------------------------------------
'* SubFolders プロパティ（読み取り専用）
'*
'* [補  足]
'* ・フォルダのコレクションを取得します。
'*-----------------------------------------------------------------------------
Property Get SubFolders() As ScriptingExFolders
    Dim wrap As ScriptingExFolders: Set wrap = New ScriptingExFolders
    Set wrap.Origin = mOrigin.SubFolders
    Set SubFolders = wrap
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
    'Set mOrigin = CreateObject("Scripting.Folder")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Copy メソッド
'* [詳  細] Copy のラッパーメソッド。
'*          指定したフォルダをコピーします｡
'*
'* @param Destination
'* @param OverWriteFiles  省略可能。
'*
'******************************************************************************
Sub Copy(Destination As String, Optional OverWriteFiles As Boolean = True)
    Call mOrigin.Copy(Destination, OverWriteFiles)
End Sub

'******************************************************************************
'* [概  要] CreateTextFile メソッド
'* [詳  細] CreateTextFile のラッパーメソッド。
'*          TextStream オブジェクトとしてファイルを作成します｡
'*
'* @param FileName
'* @param Overwrite 省略可能。
'* @param Unicode 省略可能。
'* @return ScriptingExTextStream ファイルのストリーム
'*
'******************************************************************************
Function CreateTextFile(fileName As String, Optional Overwrite As Boolean = True, Optional Unicode As Boolean = False) As ScriptingExTextStream
    Dim wrap As ScriptingExTextStream: Set wrap = New ScriptingExTextStream
    Set wrap.Origin = mOrigin.CreateTextFile(fileName, Overwrite, Unicode)
    Set CreateTextFile = wrap
End Function

'******************************************************************************
'* [概  要] Delete メソッド
'* [詳  細] Delete のラッパーメソッド。
'*         指定したフォルダを削除します｡
'*
'* @param Force 省略可能。
'*
'******************************************************************************
Sub Delete(Optional Force As Boolean = False)
    Call mOrigin.Delete(Force)
End Sub

'******************************************************************************
'* [概  要] Move メソッド
'* [詳  細] Move のラッパーメソッド。
'*          指定したフォルダを移動します｡
'*
'* @param Destination
'*
'******************************************************************************
Sub Move(Destination As String)
    Call mOrigin.Move(Destination)
End Sub
