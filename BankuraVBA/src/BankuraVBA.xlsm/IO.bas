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
'* �g�����\�b�h
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* FileSystemObject ���b�p�[�v���V�[�W��
'******************************************************************************
'*-----------------------------------------------------------------------------
'* Drives �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EDrives�I�u�W�F�N�g���擾����B
'*-----------------------------------------------------------------------------
Public Property Get Drives() As Object
    Set Drives = fso.Drives
End Property

'******************************************************************************
'* [�T  �v] BuildPath ���\�b�h
'* [��  ��] BuildPath �̃��b�p�[���\�b�h�B
'*          �����̃p�X����і��O����p�X���쐬���܂��
'*
'* @param Path
'* @param Name
'* @return �p�X
'*
'******************************************************************************
Public Function BuildPath(ByVal Path As String, ByVal name As String) As String
    BuildPath = fso.BuildPath(Path, name)
End Function

'******************************************************************************
'* [�T  �v] CopyFile ���\�b�h
'* [��  ��] CopyFile �̃��b�p�[���\�b�h�B
'*         �t�@�C�����R�s�[���܂��
'*
'* @param Source
'* @param Destination
'* @param OverWriteFiles �ȗ��\�B
'*
'******************************************************************************
Public Sub CopyFile(ByVal source As String, ByVal destination As String, Optional ByVal overWriteFiles As Boolean = True)
    Call fso.CopyFile(source, destination, overWriteFiles)
End Sub

'******************************************************************************
'* [�T  �v] CopyFolder ���\�b�h
'* [��  ��] CopyFolder �̃��b�p�[���\�b�h�B
'*          �t�H���_���R�s�[���܂��
'*
'* @param Source
'* @param Destination
'* @param OverWriteFiles �ȗ��\�B
'*
'******************************************************************************
Public Sub CopyFolder(ByVal source As String, ByVal destination As String, Optional ByVal overWriteFiles As Boolean = True)
    Call fso.CopyFolder(source, destination, overWriteFiles)
End Sub

'******************************************************************************
'* [�T  �v] CreateFolder ���\�b�h
'* [��  ��] CreateFolder �̃��b�p�[���\�b�h�B
'*          �t�H���_���쐬���܂��
'*
'* @param Path
'* @return Folder
'*
'******************************************************************************
Public Function CreateFolder(ByVal Path As String) As Object
    Set CreateFolder = fso.CreateFolder(Path)
End Function

'******************************************************************************
'* [�T  �v] CreateTextFile ���\�b�h
'* [��  ��] CreateTextFile �̃��b�p�[���\�b�h�B
'*          TextStream �I�u�W�F�N�g�Ƃ��ăt�@�C�����쐬���܂��
'*
'* @param FileName
'* @param Overwrite �ȗ��\�B
'* @param Unicode �ȗ��\�B
'* @return TextStream
'*
'******************************************************************************
Public Function CreateTextFile(ByVal fileName As String, Optional ByVal overwrite As Boolean = True, Optional ByVal unicode As Boolean = False) As Object
    Set CreateTextFile = fso.CreateTextFile(fileName, overwrite, unicode)
End Function

'******************************************************************************
'* [�T  �v] DeleteFile ���\�b�h
'* [��  ��] DeleteFile �̃��b�p�[���\�b�h�B
'*         �t�@�C�����폜���܂��
'*
'* @param FileSpec
'* @param Force �ȗ��\�B
'*
'******************************************************************************
Public Sub DeleteFile(ByVal fileSpec As String, Optional ByVal force As Boolean = False)
    Call fso.DeleteFile(fileSpec, force)
End Sub

'******************************************************************************
'* [�T  �v] DeleteFolder ���\�b�h
'* [��  ��] DeleteFolder �̃��b�p�[���\�b�h�B
'*          �t�H���_���폜���܂��
'*
'* @param FolderSpec
'* @param Force �ȗ��\�B
'*
'******************************************************************************
Public Sub DeleteFolder(ByVal folderSpec As String, Optional ByVal force As Boolean = False)
    Call fso.DeleteFolder(folderSpec, force)
End Sub


'******************************************************************************
'* [�T  �v] DriveExists ���\�b�h
'* [��  ��] DriveExists �̃��b�p�[���\�b�h�B
'*          �f�B�X�N �h���C�u�܂��̓l�b�g���[�N �h���C�u�����݂��邩�ǂ���
'*          ���肵�܂��
'*
'* @param DriveSpec
'* @return ���茋��
'*
'******************************************************************************
Public Function DriveExists(ByVal driveSpec As String) As Boolean
    DriveExists = fso.DriveExists(driveSpec)
End Function

'******************************************************************************
'* [�T  �v] FileExists ���\�b�h
'* [��  ��] FileExists �̃��b�p�[���\�b�h�B
'*          �t�@�C�������݂��邩�ǂ������肵�܂��
'*
'* @param FileSpec
'* @return ���茋��
'*
'******************************************************************************
Public Function FileExists(ByVal fileSpec As String) As Boolean
    FileExists = fso.FileExists(fileSpec)
End Function

'******************************************************************************
'* [�T  �v] FolderExists ���\�b�h
'* [��  ��] FolderExists �̃��b�p�[���\�b�h�B
'*          �p�X�����݂��邩�ǂ������肵�܂��
'*
'* @param FolderSpec
'* @return ���茋��
'*
'******************************************************************************
Public Function FolderExists(ByVal folderSpec As String) As Boolean
    FolderExists = fso.FolderExists(folderSpec)
End Function

'******************************************************************************
'* [�T  �v] GetAbsolutePathName ���\�b�h
'* [��  ��] GetAbsolutePathName �̃��b�p�[���\�b�h�B
'*          �p�X�̊�\����Ԃ��܂��
'*
'* @param Path
'* @return ��΃p�X
'*
'******************************************************************************
Public Function GetAbsolutePathName(ByVal Path As String) As String
    GetAbsolutePathName = fso.GetAbsolutePathName(Path)
End Function

'******************************************************************************
'* [�T  �v] GetBaseName ���\�b�h
'* [��  ��] GetBaseName �̃��b�p�[���\�b�h�B
'*          �p�X�̃x�[�X����Ԃ��܂��
'*
'* @param Path
'* @return �p�X�̃x�[�X��
'*
'******************************************************************************
Public Function GetBaseName(ByVal Path As String) As String
    GetBaseName = fso.GetBaseName(Path)
End Function

'******************************************************************************
'* [�T  �v] GetDrive ���\�b�h
'* [��  ��] GetDrive �̃��b�p�[���\�b�h�B
'*          �f�B�X�N�h���C�u���܂��̓l�b�g���[�N�h���C�u��UNC �����擾���܂��
'*
'* @param DriveSpec
'* @return Drive �f�B�X�N�h���C�u���^�l�b�g���[�N�h���C�u��UNC��
'*
'******************************************************************************
Public Function GetDrive(ByVal driveSpec As String) As Object
    Set GetDrive = fso.GetDrive(driveSpec)
End Function


'******************************************************************************
'* [�T  �v] GetDriveName ���\�b�h
'* [��  ��] GetDriveName �̃��b�p�[���\�b�h�B
'*          �p�X�̃h���C�u����Ԃ��܂��
'*
'* @param Path
'* @return �p�X�̃h���C�u��
'*
'******************************************************************************
Public Function GetDriveName(ByVal Path As String) As String
    GetDriveName = fso.GetDriveName(Path)
End Function


'******************************************************************************
'* [�T  �v] GetExtensionName ���\�b�h
'* [��  ��] GetExtensionName �̃��b�p�[���\�b�h�B
'*          �p�X�̊g���q��Ԃ��܂��
'*
'* @param Path
'* @return �p�X�̊g���q
'*
'******************************************************************************
Public Function GetExtensionName(ByVal Path As String) As String
    GetExtensionName = fso.GetExtensionName(Path)
End Function

'******************************************************************************
'* [�T  �v] GetFile ���\�b�h
'* [��  ��] GetFile �̃��b�p�[���\�b�h�B
'*         �t�@�C�����擾���܂��
'*
'* @param FilePath
'* @return File �t�@�C��
'*
'******************************************************************************
Public Function GetFile(ByVal filePath As String) As Object
    Set GetFile = fso.GetFile(filePath)
End Function

'******************************************************************************
'* [�T  �v] GetFileName ���\�b�h
'* [��  ��] GetFileName �̃��b�p�[���\�b�h�B
'*         �p�X�̃t�@�C������Ԃ��܂��
'*
'* @param Path
'* @return �t�@�C����
'*
'******************************************************************************
Public Function GetFileName(ByVal Path As String) As String
    GetFileName = fso.GetFileName(Path)
End Function

'******************************************************************************
'* [�T  �v] GetFileVersion ���\�b�h
'* [��  ��] GetFileVersion �̃��b�p�[���\�b�h�B
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
'* [�T  �v] GetFolder ���\�b�h
'* [��  ��] GetFolder �̃��b�p�[���\�b�h�B
'*         �t�H���_���擾���܂��
'*
'* @param FolderPath
'* @return Folder �t�H���_
'*
'******************************************************************************
Public Function GetFolder(ByVal folderPath As String) As Object
    Set GetFolder = fso.GetFolder(folderPath)
End Function

'******************************************************************************
'* [�T  �v] GetParentFolderName ���\�b�h
'* [��  ��] GetParentFolderName �̃��b�p�[���\�b�h�B
'*         1 ��̃t�H���_�̃p�X��Ԃ��܂��
'*
'* @param Path
'* @return 1��̃t�H���_�p�X
'*
'******************************************************************************
Public Function GetParentFolderName(ByVal Path As String) As String
    GetParentFolderName = fso.GetParentFolderName(Path)
End Function

'�d���̂��߃R�����g�A�E�g
'******************************************************************************
'* [�T  �v] GetSpecialFolder ���\�b�h
'* [��  ��] GetSpecialFolder �̃��b�p�[���\�b�h�B
'*         �e�V�X�e���t�H���_�̈ʒu���擾���܂��
'*
'* @param SpecialFolder
'* @return Folder �e�V�X�e���t�H���_�̈ʒu
'*
'******************************************************************************
'Public Function GetSpecialFolder(SpecialFolder As SpFolderEnum) As Object
'    Set GetSpecialFolder = fso.GetSpecialFolder(SpecialFolder)
'End Function

'******************************************************************************
'* [�T  �v] GetStandardStream ���\�b�h
'* [��  ��] GetStandardStream �̃��b�p�[���\�b�h�B
'*         �w�肵���W���� TextStream �I�u�W�F�N�g��Ԃ��܂��
'*
'* @param StandardStreamType
'* @param Unicode �ȗ��\�B
'* @return TextStream �W����TextStream�I�u�W�F�N�g
'*
'******************************************************************************
Public Function GetStandardStream(ByVal standardStreamType As StandardStreamTypes, Optional ByVal unicode As Boolean = False) As Object
    Set GetStandardStream = fso.GetStandardStream(standardStreamType, unicode)
End Function

'******************************************************************************
'* [�T  �v] GetTempName ���\�b�h
'* [��  ��] GetTempName �̃��b�p�[���\�b�h�B
'*         �ꎞ�t�@�C���̖��O�Ƃ��Ďg�p���閼�O���쐬���܂��
'*
'* @return �ꎞ�t�@�C���̖��O
'*
'******************************************************************************
Public Function GetTempName() As String
    GetTempName = fso.GetTempName()
End Function

'******************************************************************************
'* [�T  �v] MoveFile ���\�b�h
'* [��  ��] MoveFile �̃��b�p�[���\�b�h�B
'*          �t�@�C�����ړ����܂��
'*
'* @param Source
'* @param Destination
'*
'******************************************************************************
Public Sub MoveFile(ByVal source As String, ByVal destination As String)
    Call fso.MoveFile(source, destination)
End Sub

'******************************************************************************
'* [�T  �v] MoveFolder ���\�b�h
'* [��  ��] MoveFolder �̃��b�p�[���\�b�h�B
'*          �t�H���_���ړ����܂��
'*
'* @param Source
'* @param Destination
'*
'******************************************************************************
Public Sub MoveFolder(ByVal source As String, ByVal destination As String)
    Call fso.MoveFolder(source, destination)
End Sub

'******************************************************************************
'* [�T  �v] OpenTextFile ���\�b�h
'* [��  ��] OpenTextFile �̃��b�p�[���\�b�h�B
'*          �t�@�C���� TextStream �I�u�W�F�N�g�Ƃ��ĊJ���܂��
'*
'* @param FileName
'* @param IOMode �ȗ��\�B
'* @param Create �ȗ��\�B
'* @param Format �ȗ��\�B
'* @return TextStream �t�@�C���X�g���[��
'*
'******************************************************************************
Public Function OpenTextFile(ByVal fileName As String, _
                      Optional ByVal IOMode As OpenFileEnum = ForReading, _
                      Optional ByVal create As Boolean = False, _
                      Optional ByVal Format As TristateEnum = False_) As Object
    Set OpenTextFile = fso.OpenTextFile(fileName, IOMode, create, Format)
End Function
