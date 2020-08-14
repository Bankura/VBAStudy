VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptingExFileSystemObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Scripting.FileSystemObjectクラスのラップ・拡張クラス
'* [詳  細] Scripting.FileSystemObjectクラスのWrapperとしての機能を提供する他、
'*          Scripting.FileSystemObjectクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Scripting.FileSystemObjectクラスについて。
' <https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/filesystemobject-object>
' コンピューターのファイルシステムへのアクセスを提供します。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As Scripting.FileSystemObject
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
'* Drives プロパティ（読み取り専用）
'*
'* [補  足]
'* ・使用可能な容量を取得を取得する。
'*-----------------------------------------------------------------------------
Property Get Drives() As ScriptingExDrives
    Dim wrap As ScriptingExDrives: Set wrap = New ScriptingExDrives
    Set wrap.Origin = mOrigin.Drives
    Set Drives = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("Scripting.FileSystemObject")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] BuildPath メソッド
'* [詳  細] BuildPath のラッパーメソッド。
'*          既存のパスおよび名前からパスを作成します｡
'*
'* @param Path
'* @param Name
'* @return パス
'*
'******************************************************************************
Function BuildPath(Path As String, Name As String) As String
    BuildPath = mOrigin.BuildPath(Path, Name)
End Function

'******************************************************************************
'* [概  要] CopyFile メソッド
'* [詳  細] CopyFile のラッパーメソッド。
'*         ファイルをコピーします｡
'*
'* @param Source
'* @param Destination
'* @param OverWriteFiles 省略可能。
'*
'******************************************************************************
Sub CopyFile(Source As String, Destination As String, Optional OverWriteFiles As Boolean = True)
    Call mOrigin.CopyFile(Source, Destination, OverWriteFiles)
End Sub

'******************************************************************************
'* [概  要] CopyFolder メソッド
'* [詳  細] CopyFolder のラッパーメソッド。
'*          フォルダをコピーします｡
'*
'* @param Source
'* @param Destination
'* @param OverWriteFiles 省略可能。
'*
'******************************************************************************
Sub CopyFolder(Source As String, Destination As String, Optional OverWriteFiles As Boolean = True)
    Call mOrigin.CopyFolder(Source, Destination, OverWriteFiles)
End Sub

'******************************************************************************
'* [概  要] CreateFolder メソッド
'* [詳  細] CreateFolder のラッパーメソッド。
'*          フォルダを作成します｡
'*
'* @param Path
'* @return ScriptingExFolder
'*
'******************************************************************************
Function CreateFolder(Path As String) As ScriptingExFolder
    Dim wrap As ScriptingExFolder: Set wrap = New ScriptingExFolder
    Set wrap.Origin = mOrigin.CreateFolder(Path)
    Set CreateFolder = wrap
End Function

'******************************************************************************
'* [概  要] CreateTextFile メソッド
'* [詳  細] CreateTextFile のラッパーメソッド。
'*          TextStream オブジェクトとしてファイルを作成します｡
'*
'* @param FileName
'* @param Overwrite 省略可能。
'* @param Unicode 省略可能。
'* @return ScriptingExTextStream
'*
'******************************************************************************
Function CreateTextFile(fileName As String, Optional Overwrite As Boolean = True, Optional Unicode As Boolean = False) As ScriptingExTextStream
    Dim wrap As ScriptingExTextStream: Set wrap = New ScriptingExTextStream
    Set wrap.Origin = mOrigin.CreateTextFile(fileName, Overwrite, Unicode)
    Set CreateTextFile = wrap
End Function

'******************************************************************************
'* [概  要] DeleteFile メソッド
'* [詳  細] DeleteFile のラッパーメソッド。
'*         ファイルを削除します｡
'*
'* @param FileSpec
'* @param Force 省略可能。
'*
'******************************************************************************
Sub DeleteFile(FileSpec As String, Optional Force As Boolean = False)
    Call mOrigin.DeleteFile(FileSpec, Force)
End Sub

'******************************************************************************
'* [概  要] DeleteFolder メソッド
'* [詳  細] DeleteFolder のラッパーメソッド。
'*          フォルダを削除します｡
'*
'* @param FolderSpec
'* @param Force 省略可能。
'*
'******************************************************************************
Sub DeleteFolder(FolderSpec As String, Optional Force As Boolean = False)
    Call mOrigin.DeleteFolder(FolderSpec, Force)
End Sub


'******************************************************************************
'* [概  要] DriveExists メソッド
'* [詳  細] DriveExists のラッパーメソッド。
'*          ディスク ドライブまたはネットワーク ドライブが存在するかどうか
'*          判定します｡
'*
'* @param DriveSpec
'* @return 判定結果
'*
'******************************************************************************
Function DriveExists(DriveSpec As String) As Boolean
    DriveExists = mOrigin.DriveExists(DriveSpec)
End Function

'******************************************************************************
'* [概  要] FileExists メソッド
'* [詳  細] FileExists のラッパーメソッド。
'*          ファイルが存在するかどうか判定します｡
'*
'* @param FileSpec
'* @return 判定結果
'*
'******************************************************************************
Function FileExists(FileSpec As String) As Boolean
    FileExists = mOrigin.FileExists(FileSpec)
End Function

'******************************************************************************
'* [概  要] FolderExists メソッド
'* [詳  細] FolderExists のラッパーメソッド。
'*          パスが存在するかどうか判定します｡
'*
'* @param FolderSpec
'* @return 判定結果
'*
'******************************************************************************
Function FolderExists(FolderSpec As String) As Boolean
    FolderExists = mOrigin.FolderExists(FolderSpec)
End Function

'******************************************************************************
'* [概  要] GetAbsolutePathName メソッド
'* [詳  細] GetAbsolutePathName のラッパーメソッド。
'*          パスの基準表現を返します｡
'*
'* @param Path
'* @return 絶対パス
'*
'******************************************************************************
Function GetAbsolutePathName(Path As String) As String
    GetAbsolutePathName = mOrigin.GetAbsolutePathName(Path)
End Function

'******************************************************************************
'* [概  要] GetBaseName メソッド
'* [詳  細] GetBaseName のラッパーメソッド。
'*          パスのベース名を返します｡
'*
'* @param Path
'* @return パスのベース名
'*
'******************************************************************************
Function GetBaseName(Path As String) As String
    GetBaseName = mOrigin.GetBaseName(Path)
End Function

'******************************************************************************
'* [概  要] GetDrive メソッド
'* [詳  細] GetDrive のラッパーメソッド。
'*          ディスクドライブ名またはネットワークドライブのUNC 名を取得します｡
'*
'* @param DriveSpec
'* @return ScriptingExDrive ディスクドライブ名／ネットワークドライブのUNC名
'*
'******************************************************************************
Function GetDrive(DriveSpec As String) As ScriptingExDrive
    Dim wrap As ScriptingExDrive: Set wrap = New ScriptingExDrive
    Set wrap.Origin = mOrigin.GetDrive(DriveSpec)
    Set GetDrive = wrap
End Function


'******************************************************************************
'* [概  要] GetDriveName メソッド
'* [詳  細] GetDriveName のラッパーメソッド。
'*          パスのドライブ名を返します｡
'*
'* @param Path
'* @return パスのドライブ名
'*
'******************************************************************************
Function GetDriveName(Path As String) As String
    GetDriveName = mOrigin.GetDriveName(Path)
End Function


'******************************************************************************
'* [概  要] GetExtensionName メソッド
'* [詳  細] GetExtensionName のラッパーメソッド。
'*          パスの拡張子を返します｡
'*
'* @param Path
'* @return パスの拡張子
'*
'******************************************************************************
Function GetExtensionName(Path As String) As String
    GetExtensionName = mOrigin.GetExtensionName(Path)
End Function

'******************************************************************************
'* [概  要] GetFile メソッド
'* [詳  細] GetFile のラッパーメソッド。
'*         ファイルを取得します｡
'*
'* @param FilePath
'* @return ScriptingExFile ファイル
'*
'******************************************************************************
Function GetFile(FilePath As String) As ScriptingExFile
    Dim wrap As ScriptingExFile: Set wrap = New ScriptingExFile
    Set wrap.Origin = mOrigin.GetFile(FilePath)
    Set GetFile = wrap
End Function

'******************************************************************************
'* [概  要] GetFileName メソッド
'* [詳  細] GetFileName のラッパーメソッド。
'*         パスのファイル名を返します｡
'*
'* @param Path
'* @return ファイル名
'*
'******************************************************************************
Function GetFileName(Path As String) As String
    GetFileName = mOrigin.GetFileName(Path)
End Function

'******************************************************************************
'* [概  要] GetFileVersion メソッド
'* [詳  細] GetFileVersion のラッパーメソッド。
'*         Retrieve the file version of the specified file into a string
'*
'* @param FileName
'* @return file version
'*
'******************************************************************************
Function GetFileVersion(fileName As String) As String
    GetFileVersion = mOrigin.GetFileVersion(fileName)
End Function

'******************************************************************************
'* [概  要] GetFolder メソッド
'* [詳  細] GetFolder のラッパーメソッド。
'*         フォルダを取得します｡
'*
'* @param FilePath
'* @return ScriptingExFolder フォルダ
'*
'******************************************************************************
Function GetFolder(FolderPath As String) As ScriptingExFolder
    Dim wrap As ScriptingExFolder: Set wrap = New ScriptingExFolder
    Set wrap.Origin = mOrigin.GetFolder(FolderPath)
    Set GetFolder = wrap
End Function

'******************************************************************************
'* [概  要] GetParentFolderName メソッド
'* [詳  細] GetParentFolderName のラッパーメソッド。
'*         1 つ上のフォルダのパスを返します｡
'*
'* @param Path
'* @return 1つ上のフォルダパス
'*
'******************************************************************************
Function GetParentFolderName(Path As String) As String
    GetParentFolderName = mOrigin.GetParentFolderName(Path)
End Function

'******************************************************************************
'* [概  要] GetSpecialFolder メソッド
'* [詳  細] GetSpecialFolder のラッパーメソッド。
'*         各システムフォルダの位置を取得します｡
'*
'* @param SpecialFolder
'* @return ScriptingExFolder 各システムフォルダの位置
'*
'******************************************************************************
Function GetSpecialFolder(SpecialFolder As SpecialFolderConst) As ScriptingExFolder
    Dim wrap As ScriptingExFolder: Set wrap = New ScriptingExFolder
    Set wrap.Origin = mOrigin.GetSpecialFolder(SpecialFolder)
    Set GetSpecialFolder = wrap
End Function

'******************************************************************************
'* [概  要] GetStandardStream メソッド
'* [詳  細] GetStandardStream のラッパーメソッド。
'*         各システムフォルダの位置を取得します｡
'*
'* @param StandardStreamType
'* @param Unicode 省略可能。
'* @return ScriptingExTextStream 各システムフォルダの位置
'*
'******************************************************************************
Function GetStandardStream(StandardStreamType As StandardStreamTypes, Optional Unicode As Boolean = False) As ScriptingExTextStream
    Dim wrap As ScriptingExTextStream: Set wrap = New ScriptingExTextStream
    Set wrap.Origin = mOrigin.GetStandardStream(StandardStreamType, Unicode)
    Set GetStandardStream = wrap
End Function

'******************************************************************************
'* [概  要] GetTempName メソッド
'* [詳  細] GetTempName のラッパーメソッド。
'*         一時ファイルの名前として使用する名前を作成します｡
'*
'* @return 一時ファイルの名前
'*
'******************************************************************************
Function GetTempName() As String
    GetTempName = mOrigin.GetTempName()
End Function

'******************************************************************************
'* [概  要] MoveFile メソッド
'* [詳  細] MoveFile のラッパーメソッド。
'*          ファイルを移動します｡
'*
'* @param Source
'* @param Destination
'*
'******************************************************************************
Sub MoveFile(Source As String, Destination As String)
    Call mOrigin.MoveFile(Source, Destination)
End Sub

'******************************************************************************
'* [概  要] MoveFolder メソッド
'* [詳  細] MoveFolder のラッパーメソッド。
'*          フォルダを移動します｡
'*
'* @param Source
'* @param Destination
'*
'******************************************************************************
Sub MoveFolder(Source As String, Destination As String)
    Call mOrigin.MoveFolder(Source, Destination)
End Sub

'******************************************************************************
'* [概  要] OpenTextFile メソッド
'* [詳  細] OpenTextFile のラッパーメソッド。
'*          ファイルを TextStream オブジェクトとして開きます｡
'*
'* @param FileName
'* @param IOMode 省略可能。
'* @param Create 省略可能。
'* @param Format 省略可能。
'* @return ScriptingExTextStream ファイルストリーム
'*
'******************************************************************************
Function OpenTextFile(fileName As String, _
                      Optional IOMode As IOMode = ForReading, _
                      Optional Create As Boolean = False, _
                      Optional Format As Tristate = TristateFalse) As ScriptingExTextStream
    Dim wrap As ScriptingExTextStream: Set wrap = New ScriptingExTextStream
    Set wrap.Origin = mOrigin.OpenTextFile(fileName, IOMode, Create, Format)
    Set OpenTextFile = wrap
End Function

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------
