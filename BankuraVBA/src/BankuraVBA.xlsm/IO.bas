Attribute VB_Name = "IO"
'''+----                                                                   --+
'''|                             Ariawase 0.9.0                              |
'''|                Ariawase is free library for VBA cowboys.                |
'''|          The Project Page: https://github.com/vbaidiot/Ariawase         |
'''+--                                                                   ----+
Option Explicit
Option Private Module

''' @seealso Scripting.FileSystemObject http://msdn.microsoft.com/ja-jp/library/cc409798.aspx
''' @seealso ADODB.Stream http://msdn.microsoft.com/ja-jp/library/cc364272.aspx

Public Enum TristateEnum
    UseDefault = -2
    True_ = -1
    False_ = 0
End Enum

Public Enum OpenFileEnum
    ForReading = 1
    ForWriting = 2
    ForAppending = 8
End Enum

Public Enum FileAttrEnum
    ReadOnly = 1
    Hidden = 2
    System = 4
    Volume = 8
    Directory = 16
    Archive = 32
    Alias = 64
    Compressed = 128
End Enum

Public Enum DriveTypeEnum
    Removable = 1
    Fixed = 2
    Network = 3
    CDROM = 4
    RAMDisk = 5
End Enum

Public Enum SpFolderEnum
    WindowsFolder = 0
    SystemFolder = 1
    TemporaryFolder = 2
End Enum

Public Enum StreamTypeEnum
    adTypeBinary = 1
    adTypeText = 2
End Enum

Public Enum LineSeparatorsEnum
    adCRLF = -1
    adCR = 13
    adLF = 10
End Enum

Public Enum StreamOpenOptionsEnum
    adOpenStreamUnspecified = -1
    adOpenStreamAsync = 1
    adOpenStreamFromRecord = 4
End Enum

Public Enum ConnectModeEnum
    adModeUnknown = 0
    adModeRead = 1
    adModeWrite = 2
    adModeReadWrite = 3
    adModeShareDenyRead = 4
    adModeShareDenyWrite = 8
    adModeShareExclusive = 12
    adModeShareDenyNone = 16
    adModeRecursive = &H400000
End Enum

Public Enum ObjectStateEnum
    adStateClosed = 0
    adStateOpen = 1
    adStateConnecting = 2
    adStateExecuting = 4
    adStateFetching = 8
End Enum

Public Enum SaveOptionsEnum
    adSaveCreateNotExist = 1
    adSaveCreateOverWrite = 2
End Enum

Public Enum StreamReadEnum
    adReadLine = -2
    adReadAll = -1
End Enum

Public Enum StreamWriteEnum
    adWriteChar = 0
    adWriteLine = 1
End Enum

'*-----------------------------------------------------------------------------
'* StandardStreamTypes
'*   Add By Bankura
'*-----------------------------------------------------------------------------
Public Enum StandardStreamTypes
    StdErr = 2
    StdIn = 0
    StdOut = 1
End Enum

Private xxFso As Object 'Is Scripting.FileSystemObject
Private xxMimeCharsets As Variant '(Of Array(Of String))

''' @return As Object Is Scripting.FileSystemObject
Public Property Get fso() As Object
    If xxFso Is Nothing Then Set xxFso = CreateObject("Scripting.FileSystemObject")
    Set fso = xxFso
End Property

''' @return As String
Public Property Get ExecPath() As String
    Dim app As Object: Set app = Application
    Select Case app.name
        Case "Microsoft Word":   ExecPath = app.MacroContainer.Path
        Case "Microsoft Excel":  ExecPath = app.ThisWorkbook.Path
        Case "Microsoft Access": ExecPath = app.CurrentProject.Path
        Case Else: Err.Raise 17
    End Select
End Property

''' @return As Variant(Of Array(Of String))
Public Property Get MimeCharsets() As Variant
    If IsEmpty(xxMimeCharsets) Then
        Dim stdRegProv As Object: Set stdRegProv = CreateStdRegProv()
        stdRegProv.EnumKey HKEY_CLASSES_ROOT, "MIME\Database\Charset", xxMimeCharsets
    End If
    MimeCharsets = xxMimeCharsets
End Property

''' @param propType As Integer Is StreamTypeEnum
''' @param propCharset As String In MimeCharsets
''' @param propLineSeparator As Integer Is LineSeparatorsEnum
''' @return As Object Is ADODB.Stream
Public Function CreateADODBStream( _
    Optional ByVal propType As StreamTypeEnum = adTypeText, _
    Optional ByVal propCharset As String = "Unicode", _
    Optional ByVal propLineSeparator As LineSeparatorsEnum = adCRLF _
    ) As Object
    
    Set CreateADODBStream = CreateObject("ADODB.Stream")
    With CreateADODBStream
        .charSet = propCharset
        .lineSeparator = propLineSeparator
        .Type = propType
    End With
End Function

Public Function BomSize(ByVal chrset As String) As Integer
    Select Case LCase(chrset)
        Case "utf-8":             BomSize = 3
        Case "utf-16", "unicode": BomSize = 2
        Case Else:                BomSize = 0
    End Select
End Function

Public Sub SaveToFileWithoutBom( _
    ByVal strm As Object, ByVal fPath As String, ByVal opSave As SaveOptionsEnum _
    )
    
    If TypeName(strm) <> "Stream" Then Err.Raise 13
    If strm.Type <> adTypeText Then Err.Raise 5
    
    Dim strmZ As Object: Set strmZ = CreateADODBStream(adTypeBinary)
    strmZ.Open
    
    Dim chrset As String: chrset = strm.charSet
    Dim lnsep As Integer: lnsep = strm.lineSeparator
    strm.Type = adTypeBinary
    strm.Position = BomSize(chrset)
    
    strmZ.Write strm.Read(adReadAll)
    strmZ.Position = 0
    strmZ.SaveToFile fPath, opSave
    strmZ.Close
    
    strm.Position = 0
    strm.Type = adTypeText
    strm.charSet = chrset
    strm.lineSeparator = lnsep
End Sub

Public Sub RemoveBom( _
    ByVal fPath As String, ByVal chrset As String, ByVal linsep As LineSeparatorsEnum _
    )
    
    Dim strm As Object: Set strm = CreateADODBStream(chrset, linsep)
    strm.Open
    strm.LoadFromFile fPath
    SaveToFileWithoutBom strm, fPath, adSaveCreateOverWrite
    strm.Close
End Sub

Public Function ChangeCharset(ByVal strm As Object, ByVal chrset As String) As Object
    If TypeName(strm) <> "Stream" Then Err.Raise 13
    If strm.Type <> adTypeText Then Err.Raise 5
    
    Dim strmZ As Object: Set strmZ = CreateADODBStream(adTypeText, chrset, strm.lineSeparator)
    strmZ.Open
    
    If strm.State = adStateClosed Then strm.Open
    strm.copyto strmZ
    strm.Close
    
    strmZ.Position = 0
    Set ChangeCharset = strmZ
End Function

Public Sub ChangeFileCharset( _
    ByVal fPath As String, ByVal crrChrset As String, ByVal chgChrset As String _
    )
    
    Dim strm As Object: Set strm = CreateADODBStream(adTypeText, crrChrset)
    strm.Open
    strm.LoadFromFile fPath
    Set strm = ChangeCharset(strm, chgChrset)
    strm.SaveToFile fPath, adSaveCreateOverWrite
    strm.Close
End Sub

Public Function ChangeLineSeparator( _
    ByVal strm As Object, ByVal linsep As LineSeparatorsEnum _
    ) As Object
    
    If TypeName(strm) <> "Stream" Then Err.Raise 13
    If strm.Type <> adTypeText Then Err.Raise 5
    
    Dim strmZ As Object: Set strmZ = CreateADODBStream(strm.charSet, linsep)
    strmZ.Open
    
    If strm.State = adStateClosed Then strm.Open
    strm.Position = 0
    While Not strm.EOS: strmZ.WriteText strm.ReadText(adReadLine), adWriteLine: Wend
    strm.Close
    
    strmZ.Position = 0
    Set ChangeLineSeparator = strmZ
End Function

Public Sub ChangeFileLineSeparator( _
    ByVal fPath As String, ByVal chrset As String, _
    ByVal crrLinsep As LineSeparatorsEnum, ByVal chgLinsep As LineSeparatorsEnum _
    )
    
    Dim strm As Object: Set strm = CreateADODBStream(chrset, crrLinsep)
    strm.Open
    strm.LoadFromFile fPath
    Set strm = ChangeLineSeparator(strm, chgLinsep)
    strm.SaveToFile fPath, adSaveCreateOverWrite
    strm.Close
End Sub

Public Function IsPathRooted(ByVal fPath As String) As Boolean
    Dim s As String
    s = Left(fPath, 1)
    If s = "\" Or s = "/" Then
        IsPathRooted = True
        GoTo Escape
    End If
    s = Mid(fPath, 2, 1)
    If s = ":" Then
        IsPathRooted = True
        GoTo Escape
    End If
    IsPathRooted = False
    
Escape:
End Function

Public Function GetSpecialFolder(ByVal spFolder As Variant) As String
    If IsNumeric(spFolder) Then
        GetSpecialFolder = fso.GetSpecialFolder(spFolder)
    ElseIf VarType(spFolder) = vbString Then
        GetSpecialFolder = Wsh.SpecialFolders(spFolder)
    Else
        Err.Raise 13
    End If
End Function

Public Function GetTempFilePath( _
    Optional ByVal tdir As String, Optional extName As String = ".tmp" _
    ) As String
    
    If StrPtr(tdir) = 0 Then tdir = fso.GetSpecialFolder(TemporaryFolder)
    Do
        GetTempFilePath = fso.BuildPath(tdir, Replace(fso.GetTempName(), ".tmp", extName))
    Loop While fso.FileExists(GetTempFilePath)
End Function

Public Function GetUniqueFileName( _
    ByVal fPath As String, Optional delim As String = "_" _
    ) As String
    
    Dim d As String: d = fso.GetParentFolderName(fPath)
    Dim b As String: b = fso.GetBaseName(fPath) & delim
    Dim x As String: x = "." & fso.GetExtensionName(fPath)
    
    Dim n As Long: n = 0
    While fso.FileExists(fPath)
        n = n + 1
        fPath = fso.BuildPath(d, b & CStr(n) & x)
    Wend
    GetUniqueFileName = fPath
End Function

Public Function GetAllFolders(ByVal folderPath As String) As Variant
    Dim ret As Collection: Set ret = New Collection
    GetAllFoldersImpl folderPath, ret
    GetAllFolders = ClctToArr(ret)
End Function
Private Sub GetAllFoldersImpl(ByVal folderPath As String, ByVal ret As Collection)
    Dim d As Object: Set d = fso.GetFolder(folderPath)
    
    Dim sd As Object
    For Each sd In d.SubFolders
        ret.Add sd.Path
        GetAllFoldersImpl sd.Path, ret
    Next
End Sub

Public Function GetAllFiles(ByVal folderPath As String) As Variant
    Dim ret As Collection: Set ret = New Collection
    GetAllFilesImpl folderPath, ret
    GetAllFiles = ClctToArr(ret)
End Function
Private Sub GetAllFilesImpl(ByVal folderPath As String, ByVal ret As Collection)
    Dim d As Object: Set d = fso.GetFolder(folderPath)
    
    Dim fl As Object
    For Each fl In d.Files: ret.Add fl.Path: Next
    
    Dim sd As Object
    For Each sd In d.SubFolders: GetAllFilesImpl sd.Path, ret: Next
End Sub

Public Sub CreateFolderTree(ByVal folderPath As String)
    If Not fso.DriveExists(fso.GetDriveName(folderPath)) Then Err.Raise 5
    CreateFolderTreeImpl folderPath
End Sub
Private Sub CreateFolderTreeImpl(ByVal folderPath As String)
    If fso.FolderExists(folderPath) Then GoTo Escape
    CreateFolderTreeImpl fso.GetParentFolderName(folderPath)
    fso.CreateFolder folderPath
    
Escape:
End Sub


'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* 拡張メソッド
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* FileSystemObject ラッパープロシージャ
'******************************************************************************
'*-----------------------------------------------------------------------------
'* Drives プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Drivesオブジェクトを取得する。
'*-----------------------------------------------------------------------------
Public Property Get Drives() As Object
    Set Drives = fso.Drives
End Property

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
Public Function BuildPath(ByVal Path As String, ByVal name As String) As String
    BuildPath = fso.BuildPath(Path, name)
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
Public Sub CopyFile(ByVal source As String, ByVal destination As String, Optional ByVal overWriteFiles As Boolean = True)
    Call fso.CopyFile(source, destination, overWriteFiles)
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
Public Sub CopyFolder(ByVal source As String, ByVal destination As String, Optional ByVal overWriteFiles As Boolean = True)
    Call fso.CopyFolder(source, destination, overWriteFiles)
End Sub

'******************************************************************************
'* [概  要] CreateFolder メソッド
'* [詳  細] CreateFolder のラッパーメソッド。
'*          フォルダを作成します｡
'*
'* @param Path
'* @return Folder
'*
'******************************************************************************
Public Function CreateFolder(ByVal Path As String) As Object
    Set CreateFolder = fso.CreateFolder(Path)
End Function

'******************************************************************************
'* [概  要] CreateTextFile メソッド
'* [詳  細] CreateTextFile のラッパーメソッド。
'*          TextStream オブジェクトとしてファイルを作成します｡
'*
'* @param FileName
'* @param Overwrite 省略可能。
'* @param Unicode 省略可能。
'* @return TextStream
'*
'******************************************************************************
Public Function CreateTextFile(ByVal fileName As String, Optional ByVal overwrite As Boolean = True, Optional ByVal unicode As Boolean = False) As Object
    Set CreateTextFile = fso.CreateTextFile(fileName, overwrite, unicode)
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
Public Sub DeleteFile(ByVal fileSpec As String, Optional ByVal force As Boolean = False)
    Call fso.DeleteFile(fileSpec, force)
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
Public Sub DeleteFolder(ByVal folderSpec As String, Optional ByVal force As Boolean = False)
    Call fso.DeleteFolder(folderSpec, force)
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
Public Function DriveExists(ByVal driveSpec As String) As Boolean
    DriveExists = fso.DriveExists(driveSpec)
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
Public Function FileExists(ByVal fileSpec As String) As Boolean
    FileExists = fso.FileExists(fileSpec)
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
Public Function FolderExists(ByVal folderSpec As String) As Boolean
    FolderExists = fso.FolderExists(folderSpec)
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
Public Function GetAbsolutePathName(ByVal Path As String) As String
    GetAbsolutePathName = fso.GetAbsolutePathName(Path)
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
Public Function GetBaseName(ByVal Path As String) As String
    GetBaseName = fso.GetBaseName(Path)
End Function

'******************************************************************************
'* [概  要] GetDrive メソッド
'* [詳  細] GetDrive のラッパーメソッド。
'*          ディスクドライブ名またはネットワークドライブのUNC 名を取得します｡
'*
'* @param DriveSpec
'* @return Drive ディスクドライブ名／ネットワークドライブのUNC名
'*
'******************************************************************************
Public Function GetDrive(ByVal driveSpec As String) As Object
    Set GetDrive = fso.GetDrive(driveSpec)
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
Public Function GetDriveName(ByVal Path As String) As String
    GetDriveName = fso.GetDriveName(Path)
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
Public Function GetExtensionName(ByVal Path As String) As String
    GetExtensionName = fso.GetExtensionName(Path)
End Function

'******************************************************************************
'* [概  要] GetFile メソッド
'* [詳  細] GetFile のラッパーメソッド。
'*         ファイルを取得します｡
'*
'* @param FilePath
'* @return File ファイル
'*
'******************************************************************************
Public Function GetFile(ByVal filePath As String) As Object
    Set GetFile = fso.GetFile(filePath)
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
Public Function GetFileName(ByVal Path As String) As String
    GetFileName = fso.GetFileName(Path)
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
Public Function GetFileVersion(ByVal fileName As String) As String
    GetFileVersion = fso.GetFileVersion(fileName)
End Function

'******************************************************************************
'* [概  要] GetFolder メソッド
'* [詳  細] GetFolder のラッパーメソッド。
'*         フォルダを取得します｡
'*
'* @param FolderPath
'* @return Folder フォルダ
'*
'******************************************************************************
Public Function GetFolder(ByVal folderPath As String) As Object
    Set GetFolder = fso.GetFolder(folderPath)
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
Public Function GetParentFolderName(ByVal Path As String) As String
    GetParentFolderName = fso.GetParentFolderName(Path)
End Function

'重複のためコメントアウト
'******************************************************************************
'* [概  要] GetSpecialFolder メソッド
'* [詳  細] GetSpecialFolder のラッパーメソッド。
'*         各システムフォルダの位置を取得します｡
'*
'* @param SpecialFolder
'* @return Folder 各システムフォルダの位置
'*
'******************************************************************************
'Public Function GetSpecialFolder(SpecialFolder As SpFolderEnum) As Object
'    Set GetSpecialFolder = fso.GetSpecialFolder(SpecialFolder)
'End Function

'******************************************************************************
'* [概  要] GetStandardStream メソッド
'* [詳  細] GetStandardStream のラッパーメソッド。
'*         指定した標準の TextStream オブジェクトを返します｡
'*
'* @param StandardStreamType
'* @param Unicode 省略可能。
'* @return TextStream 標準のTextStreamオブジェクト
'*
'******************************************************************************
Public Function GetStandardStream(ByVal standardStreamType As StandardStreamTypes, Optional ByVal unicode As Boolean = False) As Object
    Set GetStandardStream = fso.GetStandardStream(standardStreamType, unicode)
End Function

'******************************************************************************
'* [概  要] GetTempName メソッド
'* [詳  細] GetTempName のラッパーメソッド。
'*         一時ファイルの名前として使用する名前を作成します｡
'*
'* @return 一時ファイルの名前
'*
'******************************************************************************
Public Function GetTempName() As String
    GetTempName = fso.GetTempName()
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
Public Sub MoveFile(ByVal source As String, ByVal destination As String)
    Call fso.MoveFile(source, destination)
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
Public Sub MoveFolder(ByVal source As String, ByVal destination As String)
    Call fso.MoveFolder(source, destination)
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
'* @return TextStream ファイルストリーム
'*
'******************************************************************************
Public Function OpenTextFile(ByVal fileName As String, _
                      Optional ByVal IOMode As OpenFileEnum = ForReading, _
                      Optional ByVal create As Boolean = False, _
                      Optional ByVal Format As TristateEnum = False_) As Object
    Set OpenTextFile = fso.OpenTextFile(fileName, IOMode, create, Format)
End Function
