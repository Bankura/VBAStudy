VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] File/Folder���샆�[�e�B���e�B�N���X
'* [��  ��] �t�@�C��/�t�H���_����Ɋւ��郆�[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'* [�Q  �l]
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum��`
'******************************************************************************

'******************************************************************************
'* �萔��`
'******************************************************************************

'******************************************************************************
'* �\���̒�`
'******************************************************************************

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mWinApi As WinAPI

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mWinApi = Base.GetWinAPI()
End Sub
Private Sub Class_Terminate()
    Set mWinApi = Nothing
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] �f�B���N�g���p�X�������t�^�����B
'* [��  ��] �f�B���N�g���p�X�̖����ɕ������i\�j���Ȃ���Εt�^���s���B
'*
'* @param strDirPath �f�B���N�g���p�X
'* @return �������t���f�B���N�g���p�X
'******************************************************************************
Public Function AddPathSeparator(strDirPath As String) As String
    If Right(strDirPath, 1) <> "\" Then
        AddPathSeparator = strDirPath & "\"
    Else
        AddPathSeparator = strDirPath
    End If
End Function

'******************************************************************************
'* [�T  �v] FileExists ���\�b�h
'* [��  ��] �t�@�C�������݂��邩�ǂ�������
'*
'* @param targetFilePath �t�@�C���p�X
'* @return ���茋�ʁiTrue:���݁j
'*
'******************************************************************************
Public Function FileExists(ByVal targetFilePath As String) As Boolean
    FileExists = IO.fso.FileExists(targetFilePath)
End Function

'******************************************************************************
'* [�T  �v] DirExists
'* [��  ��] �t�H���_���݃`�F�b�N
'*
'* @param strDirPath �`�F�b�N�Ώۂ̃t�H���_�p�X
'* @return Boolean �������ʁiTrue:���� False�F�ُ�j
'******************************************************************************
Public Function DirExists(strDirPath As String) As Boolean
    On Error GoTo ErrorHandler

    '�t�H���_�����݂��Ȃ��ꍇ�͍쐬���邩�m�F����
    If Dir(strDirPath, vbDirectory) = "" Then
        Dim ret As Long: ret = MsgBox("�ȉ��̃t�H���_�͑��݂��܂���B�쐬���܂����H" & _
                                   vbCrLf & vbCrLf & strDirPath, vbOKCancel, "�t�H���_�쐬")
        If ret = vbOK Then
            Dim dirRet As LongPtr: dirRet = mWinApi.SHCreateDirectoryEx(0, strDirPath, 0)
            If dirRet = 0 Then
                DirExists = True
                Exit Function
            Else
                DirExists = False
                MsgBox "�t�H���_���쐬�ł��܂���ł����B"
                Exit Function
            End If
        End If
        DirExists = False
        
    Else
        DirExists = True
    End If

    Exit Function
    
ErrorHandler:
    DirExists = False
End Function

'******************************************************************************
'* [�T  �v] ReadCsv
'* [��  ��] CSV�t�@�C���ǂݍ��݊֐��i���s�R�[�h��CRLF�j
'*
'* @param strCsvFilePath CSV�t�@�C���p�X
'* @param objOutWorksheet �f�[�^�\�t��V�[�g�I�u�W�F�N�g
'* @param lngStartRow �V�[�g�\�t�J�n�s
'* @param lngStartCol �V�[�g�\�t�J�n��
'* @return Long �Ǎ��s��
'*
'* [��  �l] ���p�~�\��iUTF-8�ɑΉ����Ă��Ȃ����߁j
'*
'******************************************************************************
Public Function ReadCsv(strCsvFilePath As String, objOutWorksheet As Worksheet, _
                        lngStartRow As Long, lngStartCol As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim buf As String, tmp As Variant, n As Long, i As Long
    n = 0
  
    Open strCsvFilePath For Input As #1
    Do Until EOF(1)
        Line Input #1, buf
        tmp = Split(buf, ",")
        objOutWorksheet.Cells(lngStartRow + n, lngStartCol).Resize(1, UBound(tmp) + lngStartCol - 1).value = tmp
        n = n + 1
    Loop

    Close #1
    ReadCsv = n
    Exit Function
    
ErrorHandler:
    ReadCsv = -1
End Function

'******************************************************************************
'* [�T  �v] ReadCsvLf
'* [��  ��] CSV�t�@�C���ǂݍ��݊֐��i���s�R�[�h��LF�j
'*
'* @param strCsvFilePath CSV�t�@�C���p�X
'* @param objOutWorksheet �f�[�^�\�t��V�[�g�I�u�W�F�N�g
'* @param lngStartRow �V�[�g�\�t�J�n�s
'* @param lngStartCol �V�[�g�\�t�J�n��
'* @return Long �Ǎ��s��
'*
'* [��  �l] ���p�~�\��iUTF-8�ɑΉ����Ă��Ȃ����߁j
'*
'******************************************************************************
Public Function ReadCsvLf(strCsvFilePath As String, objOutWorksheet As Worksheet, _
                          lngStartRow As Long, lngStartCol As Long) As Long
                          
    On Error GoTo ErrorHandler
    Dim buf() As Byte
    Dim tmp As Variant, tmp2 As Variant, n As Long, i As Long
  
    Open strCsvFilePath For Binary As #1
        ReDim buf(1 To LOF(1))
        Get #1, , buf
    Close #1
  
    tmp = Split(StrConv(buf, vbUnicode), vbLf)
    For n = 0 To UBound(tmp)
        If tmp(n) = "" Then
            n = n - 1
            Exit For
        End If
        tmp2 = Split(tmp(n), ",")
        objOutWorksheet.Cells(lngStartRow + n, lngStartCol).Resize(1, UBound(tmp2) + lngStartCol - 1).value = tmp2
    Next n
  
    ReadCsvLf = n + 1

    Exit Function
    
ErrorHandler:
    ReadCsvLf = -1
End Function


'******************************************************************************
'* [�T  �v] SaveToFileWithoutBom
'* [��  ��] �X�g���[������BOM����菜���ăt�@�C���ɕۑ�����B
'*          Ariawase��IO���W���[�����b�p�[���\�b�h�B
'*
'* @param strm ADODB.Stream�I�u�W�F�N�g
'* @param fpath �t�@�C���p�X
'* @param opSave �ۑ��I�v�V�����iSaveOptionsEnum�񋓑̂̒l�j
'*
'******************************************************************************
Public Sub SaveToFileWithoutBom( _
    ByVal strm As Object, ByVal fpath As String, ByVal opSave As SaveOptionsEnum _
    )
    Call IO.SaveToFileWithoutBom(strm, fpath, opSave)
End Sub

'******************************************************************************
'* [�T  �v] RemoveBom
'* [��  ��] �t�@�C������BOM����菜���ĕۑ�����B
'*          Ariawase��IO���W���[�����b�p�[���\�b�h�B
'*
'* @param fpath �t�@�C���p�X
'* @param chrset �����R�[�h
'* @param linsep ���s�R�[�h
'*
'******************************************************************************
Public Sub RemoveBom( _
    ByVal fpath As String, ByVal chrset As String, ByVal linsep As LineSeparatorsEnum _
    )
    Call IO.RemoveBom(fpath, chrset, linsep)
End Sub

'******************************************************************************
'* [�T  �v] ChangeCharset
'* [��  ��] �X�g���[���̕����R�[�h��ύX����B
'*          Ariawase��IO���W���[�����b�p�[���\�b�h�B
'*
'* @param strm ADODB.Stream�I�u�W�F�N�g
'* @param chrset �ύX�敶���R�[�h
'* @return Object �������ʁiTrue:���� False�F�ُ�j
'*
'******************************************************************************
Public Function ChangeCharset(ByVal strm As Object, ByVal chrset As String) As Object
    ChangeCharset = IO.ChangeCharset(strm, chrset)
End Function

'******************************************************************************
'* [�T  �v] ChangeFileCharset
'* [��  ��] �t�@�C���̕����R�[�h��ύX���ĕۑ�����B
'*          Ariawase��IO���W���[�����b�p�[���\�b�h�B
'*
'* @param fpath �t�@�C���p�X
'* @param crrChrset �����R�[�h
'* @param chgChrset �ύX�㕶���R�[�h
'*
'******************************************************************************
Public Sub ChangeFileCharset( _
    ByVal fpath As String, ByVal crrChrset As String, ByVal chgChrset As String _
    )
    Call IO.ChangeFileCharset(fpath, crrChrset, chgChrset)
End Sub

'******************************************************************************
'* [�T  �v] ChangeLineSeparator
'* [��  ��] �X�g���[���̉��s�R�[�h��ύX����B
'*          Ariawase��IO���W���[�����b�p�[���\�b�h�B
'*
'* @param strm ADODB.Stream�I�u�W�F�N�g
'* @param linsep �ύX����s�R�[�h
'* @return Object ADODB.Stream�I�u�W�F�N�g�i���s�R�[�h�ύX��j
'*
'******************************************************************************
Public Function ChangeLineSeparator( _
    ByVal strm As Object, ByVal linsep As LineSeparatorsEnum _
    ) As Object
    ChangeLineSeparator = IO.ChangeLineSeparator(strm, linsep)
End Function

'******************************************************************************
'* [�T  �v] ChangeFileLineSeparator
'* [��  ��] �t�@�C���̉��s�R�[�h��ύX���ĕۑ�����B
'*          Ariawase��IO���W���[�����b�p�[���\�b�h�B
'*
'* @param fpath �t�@�C���p�X
'* @param chrset �����R�[�h
'* @param crrLinsep ���s�R�[�h
'* @param chgLinsep �ύX����s�R�[�h
'*
'******************************************************************************
Public Sub ChangeFileLineSeparator( _
    ByVal fpath As String, ByVal chrset As String, _
    ByVal crrLinsep As LineSeparatorsEnum, ByVal chgLinsep As LineSeparatorsEnum _
    )
    Call IO.ChangeFileLineSeparator(fpath, chrset, crrLinsep, chgLinsep)
End Sub

'******************************************************************************
'* [�T  �v] IsPathRooted
'* [��  ��] ��΃p�X���ǂ������`�F�b�N����B
'*          Ariawase��IO���W���[�����b�p�[���\�b�h�B
'*
'* @param fpath �t�@�C��/�t�H���_�p�X
'* @return Boolean �`�F�b�N���ʁiTrue:�Y���j
'*
'******************************************************************************
Public Function IsPathRooted(ByVal fpath As String) As Boolean
    IsPathRooted = IO.IsPathRooted(fpath)
End Function

'******************************************************************************
'* [�T  �v] GetSpecialFolder
'* [��  ��] Windows�̓���f�B���N�g���p�X���擾����B
'*          Ariawase��IO���W���[�����b�p�[���\�b�h�B
'*
'* @param spFolder IO.SpFolderEnum�񋓑̂̒l
'*                 WshShell.SpecialFolders�Ɏw��\�ȕ�����
'* @return String ����f�B���N�g���p�X
'* @seealso https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/getspecialfolder-method
'* @seealso https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364490(v=msdn.10)
'*
'******************************************************************************
Public Function GetSpecialFolder(ByVal spFolder As Variant) As String
    GetSpecialFolder = IO.GetSpecialFolder(spFolder)
End Function

'******************************************************************************
'* [�T  �v] GetTempFilePath
'* [��  ��] ��ӂƂȂ�Temp�t�@�C���p�X�𐶐�����B
'*          Ariawase��IO���W���[�����b�p�[���\�b�h�B
'*
'* @param tdir Temp�t�@�C���쐬�f�B���N�g���i�ȗ����͕W��Temp�t�H���_�j
'* @param extName Temp�t�@�C���g���q�i�ȗ�����.tmp�j
'* @param willCreateFile True�̏ꍇ�t�@�C���쐬���s��
'* @return String Temp�t�@�C���p�X
'*
'******************************************************************************
Public Function GetTempFilePath( _
    Optional ByVal tdir As String, Optional extName As String = ".tmp", _
    Optional willCreateFile As Boolean = False) As String
    GetTempFilePath = IO.GetTempFilePath(tdir, extName)
    If willCreateFile Then
        Call IO.fso.CreateTextFile(GetTempFilePath)
    End If
End Function

'******************************************************************************
'* [�T  �v] GetUniqueFileName
'* [��  ��] �w�肵���t�@�C���p�X�̃t�@�C�������݂���ꍇ�A�t�@�C���x�[�X������
'*          �Ƀf���~�^ + �����������āA��ӂƂȂ�t�@�C�����𐶐�����B
'*          Ariawase��IO���W���[�����b�p�[���\�b�h�B
'*
'* @param fpath �t�@�C����
'* @param delim �f���~�^
'* @return String ��ӂƂȂ�t�@�C����
'*
'******************************************************************************
Public Function GetUniqueFileName( _
    ByVal fpath As String, Optional delim As String = "_" _
    ) As String
    GetUniqueFileName = IO.GetUniqueFileName(fpath, delim)
End Function

'******************************************************************************
'* [�T  �v] GetAllFolders
'* [��  ��] �w�肵���p�X�z���̂��ׂẴt�H���_�i�T�u�t�H���_���܂߁j��
'*          �t�H���_�p�X��������擾���AVariant�z��Ƃ��ĕԋp����B
'*          Ariawase��IO���W���[�����b�p�[���\�b�h�B
'*
'* @param folderPath �t�H���_�p�X
'* @return Variant �t�H���_�p�X�̃��X�g�i1�����z��j
'*
'******************************************************************************
Public Function GetAllFolders(ByVal folderPath As String) As Variant
    GetAllFolders = IO.GetAllFolders(folderPath)
End Function

'******************************************************************************
'* [�T  �v] GetAllFiles
'* [��  ��] �w�肵���p�X�z���̂��ׂẴt�@�C���i�T�u�t�H���_���܂߁j��
'*          �t�@�C���p�X��������擾���AVariant�z��Ƃ��ĕԋp����B
'*          Ariawase��IO���W���[�����b�p�[���\�b�h�B
'*
'* @param folderPath �t�H���_�p�X
'* @return Variant �t�@�C���p�X�̃��X�g�i1�����z��j
'*
'******************************************************************************
Public Function GetAllFiles(ByVal folderPath As String) As Variant
    GetAllFiles = IO.GetAllFiles(folderPath)
End Function

'******************************************************************************
'* [�T  �v] CreateFolderTree
'* [��  ��] �w�肵���p�X�̊K�w���ɑ��݂��Ȃ��t�H���_������΍쐬����B
'*          Ariawase��IO���W���[�����b�p�[���\�b�h�B
'*
'* @param folderPath �t�H���_�p�X
'*
'******************************************************************************
Public Sub CreateFolderTree(ByVal folderPath As String)
    Call IO.CreateFolderTree(folderPath)
End Sub

'******************************************************************************
'* [�T  �v] GetFileTime
'* [��  ��] FileTime ���擾����
'*
'* @param dtSetting �ݒ����
'* @return FileTime �t�@�C������
'*
'******************************************************************************
Private Function GetFileTime(ByVal dtSetting As Date) As FILETIME
    Dim tSystemTime As SystemTime

    With tSystemTime
        .wYear = Year(dtSetting)
        .wMonth = Month(dtSetting)
        .wDayOfWeek = Weekday(dtSetting)
        .wDay = Day(dtSetting)
        .wHour = hour(dtSetting)
        .wMinute = Minute(dtSetting)
        .wSecond = second(dtSetting)
    End With

    Dim tLocalTime As FILETIME
    Call mWinApi.SystemTimeToFileTime(tSystemTime, tLocalTime)

    Dim tFileTime As FILETIME
    Call mWinApi.LocalFileTimeToFileTime(tLocalTime, tFileTime)

    GetFileTime = tFileTime
End Function

'******************************************************************************
'* [�T  �v] GetFileHandle
'* [��  ��] �t�@�C���̃n���h�����擾����
'*
'* @param stFilePath �Ώۃt�@�C���̃p�X
'* @return LongPtr �n���h���ԍ�
'*
'******************************************************************************
Private Function GetFileHandle(ByVal stFilePath As String) As LongPtr
    GetFileHandle = mWinApi.CreateFile( _
                        stFilePath, GENERIC_READ Or GENERIC_WRITE, _
                        FILE_SHARE_READ, 0, OPEN_EXISTING, _
                        FILE_ATTRIBUTE_NORMAL, 0 _
                    )
                   
End Function

'******************************************************************************
'* [�T  �v] SetFileTimeProc
'* [��  ��] �t�@�C���̓������w�肵�����t�Ǝ��Ԃɐݒ肷��
'*
'* @param stFilePath �Ώۃt�@�C���̃p�X
'* @param dtCreateTime �ݒ肷����t�Ǝ���
'* @param onCreationTime �쐬�����ɐݒ肷�邩
'* @param onLastWriteTime �X�V�����ɐݒ肷�邩
'* @param onLastAccessTime �A�N�Z�X�����ɐݒ肷�邩
'*
'******************************************************************************
Private Sub SetFileTimeProc(ByVal stFilePath As String, ByVal dtCreateTime As Date, _
                            ByVal onCreationTime As Boolean, _
                            ByVal onLastWriteTime As Boolean, _
                            ByVal onLastAccessTime As Boolean)
    ' FileTime ���擾����
    Dim tFileTime As FILETIME
    tFileTime = GetFileTime(dtCreateTime)

    ' �t�@�C���̃n���h�����擾����
    Dim cFileHandle As LongPtr
    cFileHandle = GetFileHandle(stFilePath)

    ' �t�@�C���̃n���h�����擾�ł����ꍇ�̂݁u�쐬�����v���X�V����
    If cFileHandle >= 0 Then
        Dim tNullable As FILETIME
        Dim tCreationTime As FILETIME
        Dim tLastAccessTime As FILETIME
        Dim tLastWriteTime As FILETIME
        
        If onCreationTime Then tCreationTime = tFileTime Else tCreationTime = tNullable
        If onLastAccessTime Then tLastAccessTime = tFileTime Else tLastAccessTime = tNullable
        If onLastWriteTime Then tLastWriteTime = tFileTime Else tLastWriteTime = tNullable

        Call mWinApi.SetFileTime(cFileHandle, _
                                 tCreationTime, _
                                 tLastAccessTime, _
                                 tLastWriteTime _
                        )
        Call mWinApi.CloseHandle(cFileHandle)
    End If
End Sub

'******************************************************************************
'* [�T  �v] SetCreationTime
'* [��  ��] �t�@�C���̍쐬�������w�肵�����t�Ǝ��Ԃɐݒ肷��
'*
'* @param stFilePath �Ώۃt�@�C���̃p�X
'* @param dtCreateTime �쐬�����ɐݒ肷����t�Ǝ���
'*
'******************************************************************************
Public Sub SetCreationTime(ByVal stFilePath As String, ByVal dtCreateTime As Date)
    Call SetFileTimeProc(stFilePath, dtCreateTime, True, False, False)
End Sub

'******************************************************************************
'* [�T  �v] SetLastWriteTime
'* [��  ��] �t�@�C���̍X�V�������w�肵�����t�Ǝ��Ԃɐݒ肷��
'*
'* @param stFilePath �Ώۃt�@�C���̃p�X
'* @param dtUpdateTime �X�V�����ɐݒ肷����t�Ǝ�
'*
'******************************************************************************
Public Sub SetLastWriteTime(ByVal stFilePath As String, ByVal dtUpdateTime As Date)
    Call SetFileTimeProc(stFilePath, dtUpdateTime, False, False, True)
End Sub

'******************************************************************************
'* [�T  �v] SetLastAccessTime
'* [��  ��] �t�@�C���̃A�N�Z�X�������w�肵�����t�Ǝ��Ԃɐݒ肷��
'*
'* @param stFilePath �Ώۃt�@�C���̃p�X
'* @param dtAccessTime �A�N�Z�X�����ɐݒ肷����t�Ǝ���
'*
'******************************************************************************
Public Sub SetLastAccessTime(ByVal stFilePath As String, ByVal dtAccessTime As Date)
    Call SetFileTimeProc(stFilePath, dtAccessTime, False, True, False)
End Sub

'******************************************************************************
'* [�T  �v] SetAllTime
'* [��  ��] �t�@�C���̍쐬�����E�X�V�����E�A�N�Z�X�������w�肵�����t�Ǝ��Ԃ�
'          �ݒ肷��
'*
'* @param stFilePath �Ώۃt�@�C���̃p�X
'* @param dtSetTime �����ɐݒ肷����t�Ǝ���
'*
'******************************************************************************
Public Sub SetAllTime(ByVal stFilePath As String, ByVal dtSetTime As Date)
    Call SetFileTimeProc(stFilePath, dtSetTime, True, True, True)
End Sub

'******************************************************************************
'* [�T  �v] �t�@�C���G���R�[�h�ꊇ�ϊ������B
'* [��  ��] �w�肵���t�H���_���̃t�@�C���̃G���R�[�h���ꊇ�ϊ�����B
'*
'* @param targetFolderName �ΏۂƂȂ�t�H���_�̃t���p�X
'* @param srcEncode �ύX���G���R�[�h
'* @param destEncode �ύX��G���R�[�h
'* @param bomInclude BOM�L���i�ȗ��B�K���False:BOM���j
'******************************************************************************
Public Sub ChangeFilesEncode(targetFolderName As String, srcEncode As String, destEncode As String, _
                            Optional bomInclude As Boolean = False)
    Dim fso, oFolder, oFiles, oFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFolder = fso.GetFolder(targetFolderName)
    Set oFiles = oFolder.Files
    
    For Each oFile In oFiles
        Dim filePath As String
        filePath = fso.GetAbsolutePathName(AddPathSeparator(targetFolderName) & oFile.Name)
         
        ' �t�@�C���G���R�[�h�ϊ�
        Call ChangeFileEncode(filePath, srcEncode, destEncode, bomInclude)
    Next
End Sub

'******************************************************************************
'* [�T  �v] �t�@�C���G���R�[�h�ϊ������B
'* [��  ��] �w�肵���t�@�C���̃G���R�[�h��ϊ�����B
'*
'* @param filePath �ΏۂƂȂ�t�@�C���̃t���p�X
'* @param srcEncode �ύX���G���R�[�h
'* @param destEncode �ύX��G���R�[�h
'* @param bomInclude BOM�L���i�ȗ��B�K���False:BOM���j
'******************************************************************************
Public Sub ChangeFileEncode(filePath As String, srcEncode As String, destEncode As String, _
                            Optional bomInclude As Boolean = False)
    Dim adoStream1 As Object, adoStream2 As Object
    Set adoStream1 = IO.CreateADODBStream(adTypeText, srcEncode)
    Set adoStream2 = IO.CreateADODBStream(adTypeText, destEncode)
         
    ' �ύX���t�@�C��Stream�Ǎ�
    With adoStream1
        .Open
        .LoadFromFile filePath
    End With
     
    ' �ύX��t�@�C��Stream�Ǎ�
    adoStream2.Open
     
    ' �G���R�[�h�ϊ�
    adoStream1.CopyTo adoStream2
    If Not bomInclude Then ExcludeBom adoStream2  ' BOM�Ή�
    adoStream2.SaveToFile filePath, adSaveCreateOverWrite '�t�@�C���㏑�w��
     
    ' Stream�N���[�Y
    adoStream2.Close
    adoStream1.Close
End Sub

'******************************************************************************
'* [�T  �v] �t�@�C���Ǎ��E�������ݏ����B
'* [��  ��] �w�肵���Ǎ��t�@�C���̃f�[�^��ʃt�@�C���ɏ������ށB
'* [�Q  �l] ��e�ʃf�[�^�̓ǂݍ��݂ɂ��Ă͈ȉ��̃T�C�g���Q�l�ɂ����B
'*          <https://mussyu1204.myhome.cx/wordpress/it/?p=720>
'*
'* @param srcFilePath �Ǎ��t�@�C���̃t���p�X
'* @param srcEncode �Ǎ����G���R�[�h
'* @param srcSep �Ǎ������s�R�[�h
'* @param destFilePath �����t�@�C���̃t���p�X
'* @param destEncode ������G���R�[�h
'* @param destSep ��������s�R�[�h
'* @param func �s�ҏW�����p�֐����B
'*             �ȉ��̂悤�Ɉ����ɕ�����A�߂�l�ɕ������Ԃ��֐������w��B
'*             funcName(row As String) As String
'*             �w�肵�Ȃ��i�󕶎��j�ꍇ�́A�s�ҏW�͍s��Ȃ��B
'*             Func�I�u�W�F�N�g���w��\�B
'*
'* @param chunkSize �`�����N�T�C�Y�B���̃T�C�Y�𒴂���Ǎ��f�[�^�̏ꍇ�́A
'*                  �`�����N�T�C�Y���Ƃɕ������ď������s���B
'* @param bomInclude BOM�L���i�ȗ��B�K���False:BOM���j
'******************************************************************************
Public Sub ReadAndWrite(srcFilePath As String, srcEncode As String, srcSep As LineSeparatorsEnum, _
                        destFilePath As String, destEncode As String, destSep As LineSeparatorsEnum, _
                        Optional Func As Variant, _
                        Optional chunkSize As Long = 2048, _
                        Optional bomInclude As Boolean = False)

    Dim inStream As Object, outStream As Object
    Set inStream = IO.CreateADODBStream(adTypeText, srcEncode, srcSep)
    Set outStream = IO.CreateADODBStream(adTypeText, destEncode, destSep)

    With inStream
        .Open
        .LoadFromFile srcFilePath
    End With
     
    outStream.Open

    Dim lines As Variant, lastLine As String
    
    ' �Ǎ��f�[�^�̃T�C�Y���w��T�C�Y���傫���ꍇ�͕��������i�������j���{
    If inStream.size > chunkSize Then
        Do Until inStream.EOS
            Dim tmp As String: tmp = inStream.ReadText(chunkSize)
            lines = Split(tmp, vbLf)
             
            Dim lineCnt As Long: lineCnt = UBound(lines)
            lines(0) = lastLine + lines(0)
            
            Dim i As Long
            For i = 0 To (lineCnt - 1)
                lines(i) = Replace(lines(i), vbCr, "")
                
                lines(i) = StringUtils.EditTextByFunc(lines(i), Func)
                
                outStream.WriteText CStr(lines(i)), adWriteLine
            Next
             
            lastLine = lines(lineCnt)
        Loop
        If lastLine <> "" Then
            outStream.WriteText lastLine, adWriteLine
        End If
    Else
        Do Until inStream.EOS
            Dim tmpLine As String: tmpLine = inStream.ReadText(adReadLine)
            
            tmpLine = StringUtils.EditTextByFunc(tmpLine, Func)
            
            outStream.WriteText tmpLine, adWriteChar
        Loop
    End If

    ' �t�@�C���ۑ�
    outStream.SaveToFile destFilePath, adSaveCreateOverWrite
     
    inStream.Close
    outStream.Close
End Sub

'******************************************************************************
'* [�T  �v] BOM���O�����B
'* [��  ��] UTF-8�e�L�X�g��BOM�����O����B
'* [���@�l] �Ăяo���O��Stream��Open����Ă��邱�ƁB
'*          WriteText�ďo������A�܂��̓t�@�C���ۑ��O�Ɏ��s���邱�ƁB
'*
'* @param streamObj Open�ς�Stream�I�u�W�F�N�g
'* @return BOM���O�f�[�^
'******************************************************************************
Public Function ExcludeBom(streamObj As Object) As Byte()
    With streamObj
        .Position = 0
        .Type_ = adTypeBinary
        .Position = IO.BomSize(streamObj.charset)  'BOM��΂�
        ExcludeBom = .Read
        .Position = 0
        .Write ExcludeBom
        .SetEOS
    End With
End Function


'******************************************************************************
'* [�T  �v] �e�L�X�g�t�@�C���Ǎ������B
'* [��  ��] �e�L�X�g�t�@�C���̓Ǎ����s���i�ꊇ�Ǎ��j�B
'*
'* @param filePath �ΏۂƂȂ�e�L�X�g�t�@�C���̃t���p�X
'* @param cSet �L�����Z�b�g�i�����R�[�h�w��j
'*             �ȗ����͎������ʁi"_autodetect_all"�j
'*               ���g�p�\��Charset��F
'*                 utf-8�Ashift_jis�Aeuc-jp�AISO-2022-JP�Aunicode�i����j
'*                 �uHKEY_CLASSES_ROOT\MIME\Database\Charset�v�Q�ƁB
'* @return �Ǎ��e�L�X�g�f�[�^
'******************************************************************************
Public Function ReadTextFile(filePath As String, Optional cSet As String = "_autodetect_all") As String
    Dim data As String
    With IO.CreateADODBStream(adTypeText, cSet)
        .Mode = adModeReadWrite
        .Open
        .LoadFromFile filePath '�t�@�C���w��
        .Position = 0
        data = .ReadText()
        .Close
    End With

    ReadTextFile = data
End Function

'******************************************************************************
'* [�T  �v] UTF-8�e�L�X�g�t�@�C���Ǎ������B
'* [��  ��] UTF-8�e�L�X�g�t�@�C���̓Ǎ����s���i�ꊇ�Ǎ��j�B
'*
'* @param filePath �ΏۂƂȂ�e�L�X�g�t�@�C���̃t���p�X
'* @return �Ǎ��e�L�X�g�f�[�^
'******************************************************************************
Public Function ReadUTF8TextFile(filePath As String) As String
    ReadUTF8TextFile = ReadTextFile(filePath, "UTF-8")
End Function

'******************************************************************************
'* [�T  �v] �e�L�X�g�t�@�C���Ǎ������B
'* [��  ��] �e�L�X�g�t�@�C�����s�P�ʂœǍ��݁AVariant�z��ɐݒ肵�ĕԋp����B
'*
'* @param filePath �ΏۂƂȂ�e�L�X�g�t�@�C���̃t���p�X
'* @param cSet �L�����Z�b�g�i�����R�[�h�w��j
'*             �ȗ����͎������ʁi"_autodetect_all"�j
'*               ���g�p�\��Charset��F
'*                 utf-8�Ashift_jis�Aeuc-jp�AISO-2022-JP�Aunicode�i����j
'*                 �uHKEY_CLASSES_ROOT\MIME\Database\Charset�v�Q�ƁB
'* @param separator ���s�R�[�h�i�ȗ��B�K���CRLF�j
'* @return �Ǎ��e�L�X�g�f�[�^�iVariant�z��j
'******************************************************************************
Public Function ReadTextFileToVArray(filePath As String, _
                                     Optional cSet As String = "_autodetect_all", _
                                     Optional separator As LineSeparatorsEnum = LineSeparatorsEnum.adCRLF) As Variant

    Dim varr(), lineCount As Long, cnt As Long
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeText, cSet, separator)
    With stream
        .Mode = adModeReadWrite
        .Open
        .LoadFromFile filePath '�t�@�C���w��
        lineCount = CountLine(stream)
        ReDim Preserve varr(lineCount - 1)
        .Position = 0
        Do While Not .EOS
            varr(cnt) = .ReadText(adReadLine)
            cnt = cnt + 1
        Loop
        .Close
    End With

    ReadTextFileToVArray = varr
End Function

'******************************************************************************
'* [�T  �v] �e�L�X�g�̍s���擾�����B
'* [��  ��] �e�L�X�g�̍s�����擾����B
'* [���@�l] �Ăяo���O��Stream��Open����Ă��邱�ƁB
'*
'* @param stream Open�ς�Stream�I�u�W�F�N�g
'* @return �s��
'******************************************************************************
Public Function CountLine(stream As Object) As Long
    Dim lineCount As Long
    With stream
        .Position = 0
        Do While Not .EOS
            .SkipLine
            lineCount = lineCount + 1
        Loop
    End With

    CountLine = lineCount
End Function
  
'******************************************************************************
'* [�T  �v] �e�L�X�g�t�@�C���s���擾�����B
'* [��  ��] �e�L�X�g�t�@�C����Ǎ��݁A�s����ԋp����B
'*
'* @param filePath �ΏۂƂȂ�e�L�X�g�t�@�C���̃t���p�X
'* @param cSet �L�����Z�b�g�i�����R�[�h�w��j
'*             �ȗ����͎������ʁi"_autodetect_all"�j
'*               ���g�p�\��Charset��F
'*                 utf-8�Ashift_jis�Aeuc-jp�AISO-2022-JP�Aunicode�i����j
'*                 �uHKEY_CLASSES_ROOT\MIME\Database\Charset�v�Q�ƁB
'* @param separator ���s�R�[�h�i�ȗ��B�K���CRLF�j
'* @return �s��
'******************************************************************************
Public Function GetTextFileLineCount(filePath As String, _
                                      Optional cSet As String = "_autodetect_all", _
                                      Optional separator As LineSeparatorsEnum = adCRLF) As Long
    ' �t�@�C�������݂��Ȃ��ꍇ��-1�ԋp
    If (IO.fso.FileExists(filePath) = False) Then
        GetTextFileLineCount = -1
        Exit Function
    End If
    If Not separator = adCR Then
        ' �ǋL���[�h�ŊJ���s�����J�E���g
        GetTextFileLineCount = IO.fso.OpenTextFile(filePath, ForAppending).line - 1
        Exit Function
    End If
    
    Dim lineCount As Long
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeText, cSet, separator)
    With stream
        .Mode = adModeReadWrite
        .Open
        .LoadFromFile filePath '�t�@�C���w��
        lineCount = .CountLine(stream)
        .Close
    End With

    GetTextFileLineCount = lineCount
End Function

'******************************************************************************
'* [�T  �v] �e�L�X�g�t�@�C�����������B
'* [��  ��] �e�L�X�g�t�@�C���̏������s���B
'*
'* @param filePath �ΏۂƂȂ�e�L�X�g�t�@�C���̃t���p�X
'* @param data �����e�L�X�g�f�[�^
'* @param cSet �L�����Z�b�g�i�����R�[�h�w��j
'*               ���g�p�\��Charset��F
'*                 utf-8�Ashift_jis�Aeuc-jp�AISO-2022-JP�Aunicode�i����j
'*                 �uHKEY_CLASSES_ROOT\MIME\Database\Charset�v�Q�ƁB
'* @param willAppend �ǋL�L���i�ȗ��B�K���False:�ǋL�Ȃ��j
'* @param bomInclude BOM�L���i�ȗ��B�K���False:BOM���j
'******************************************************************************
Public Sub WriteTextFile(filePath As String, data As String, cSet As String, _
                         Optional willAppend As Boolean = False, _
                         Optional bomInclude As Boolean = False)
                         
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeText, cSet)
    With stream
        .Mode = adModeReadWrite
        .Open
        If willAppend Then
            .AppendMode stream, filePath
        End If
        .WriteText data, adWriteChar
        If Not bomInclude Then ExcludeBom stream  ' BOM�Ή�
            
        .SaveToFile filePath, adSaveCreateOverWrite '�t�@�C���㏑�w��
        .Close
    End With
End Sub

'******************************************************************************
'* [�T  �v] UTF-8�e�L�X�g�t�@�C�����������B
'* [��  ��] UTF-8�e�L�X�g�t�@�C���̏������s���B
'*
'* @param filePath �ΏۂƂȂ�e�L�X�g�t�@�C���̃t���p�X
'* @param data �����e�L�X�g�f�[�^
'* @param willAppend �ǋL�L���i�ȗ��B�K���False:�ǋL�Ȃ��j
'* @param bomInclude BOM�L���i�ȗ��B�K���False:BOM���j
'******************************************************************************
Public Sub WriteUTF8TextFile(filePath As String, data As String, _
                             Optional willAppend As Boolean = False, _
                             Optional bomInclude As Boolean = False)
     Call WriteTextFile(filePath, data, "UTF-8", willAppend, bomInclude)
End Sub

'******************************************************************************
'* [�T  �v] �e�L�X�g�t�@�C�����������B
'* [��  ��] �e�L�X�g�t�@�C���ɔz��f�[�^���s�P�ʂŏ������s���B
'*
'* @param filePath �ΏۂƂȂ�e�L�X�g�t�@�C���̃t���p�X
'* @param vArray �����z��f�[�^�iVariant�j
'* @param cSet �L�����Z�b�g�i�����R�[�h�w��j
'*               ���g�p�\��Charset��F
'*                 utf-8�Ashift_jis�Aeuc-jp�AISO-2022-JP�Aunicode�i����j
'*                 �uHKEY_CLASSES_ROOT\MIME\Database\Charset�v�Q�ƁB
'* @param separator ���s�R�[�h�i�ȗ��B�K���CRLF�j
'* @param willAppend �ǋL�L���i�ȗ��B�K���False:�ǋL�Ȃ��j
'* @param bomInclude BOM�L���i�ȗ��B�K���False:BOM���j
'******************************************************************************
Public Sub WriteTextFileFromVArray(filePath As String, vArray As Variant, cSet As String, _
                                   Optional separator As LineSeparatorsEnum = adCRLF, _
                                   Optional willAppend As Boolean = False, _
                                   Optional bomInclude As Boolean = False)
    Dim tmp
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeText, cSet, separator)
    With stream
        .Mode = adModeReadWrite
        .Open
        If willAppend Then
            .AppendMode stream, filePath
        End If
        For Each tmp In vArray
            .WriteText CStr(tmp), adWriteLine
        Next
        If Not bomInclude Then ExcludeBom stream  ' BOM�Ή�
        .SaveToFile filePath, adSaveCreateOverWrite '�t�@�C���㏑�w��
        .Close
    End With
End Sub

'******************************************************************************
'* [�T  �v] �t�@�C���A�y���h�����B
'* [��  ��] �t�@�C���ǋL���邽�߂̏������s���B
'*          �t�@�C���̑��݃`�F�b�N���s���A���݂���ꍇ�͑Ώۃt�@�C����
'*          �|�C���^���I�[�ֈړ�����B
'* [���@�l] �Ăяo���O��Stream��Open����Ă��邱�ƁB
'*
'* @param filePath �����ݑΏۂƂȂ�e�L�X�g�t�@�C���̃t���p�X
'******************************************************************************
Public Sub AppendMode(stream As Object, filePath As String)
    If IO.fso.FileExists(filePath) Then
        With stream
            .LoadFromFile filePath
            .Position = .size '�|�C���^���I�[��
        End With
    End If
End Sub

'******************************************************************************
'* [�T  �v] BOM�L���`�F�b�N�����B
'* [��  ��] UTF-8�e�L�X�g�t�@�C����BOM�����邩�`�F�b�N���s���B
'*
'* @param filePath �ΏۂƂȂ�e�L�X�g�t�@�C���̃t���p�X
'* @param �`�F�b�N���ʁitrue:BOM�L�Afalse:BOM���j
'******************************************************************************
Function IsUTF8BomIncluded(filePath As String) As Boolean
    Dim buf() As Byte
    Dim suspected As String
    Dim v As Variant

    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeBinary)
    With stream
        .Open
        .LoadFromFile filePath
        buf = .Read(3) 'BOM�ɂ�����擪3�o�C�g���擾
        .Close
    End With

    For Each v In buf
        suspected = suspected & Hex(v)
    Next
    
    If suspected = "EFBBBF" Then
        IsUTF8BomIncluded = True
    Else
        IsUTF8BomIncluded = False
    End If
End Function

'******************************************************************************
'* [�T  �v] �t�@�C���_���v�����B
'* [��  ��] �t�@�C����Ǎ��݁A16�i�_���v�����f�[�^��ԋp����B
'*
'* @param filePath �ΏۂƂȂ�e�L�X�g�t�@�C���̃t���p�X
'* @return 16�i�_���v�����f�[�^
'******************************************************************************
Public Function ReadFileToDump(filePath As String) As String
    Dim ret As StringEx, buf As String, i As Long
    Set ret = New StringEx
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeBinary)
    With stream
        .Mode = adModeReadWrite
        .Type_ = adTypeBinary
        .Open
        .LoadFromFile filePath '�t�@�C���w��
        Do While Not .EOS
            buf = .Read(16)
 
            For i = 1 To LenB(buf)
                Dim tmp As String
                tmp = MidB(buf, i, 1)
                tmp = AscB(tmp)
                tmp = Hex(tmp)
                tmp = UCase(tmp)
                tmp = Right("0" & tmp, 2)
                Call ret.Append(tmp).Append(" ")
            Next
             Call ret.Append(vbCrLf)
        Loop
        .Close
    End With

    ReadFileToDump = ret.ToString
End Function