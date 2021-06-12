VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvFileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IRowEditor

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] CSVFile�����N���X
'* [��  ��] CSVFile�����p�̃��[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mAdoStream As Object
Private mCharSet As String
Private mLineSeparator As LineSeparatorsEnum
Private mHeaderExists As Boolean
Private mEnclosureChar As String
Private mDelimiter As String
Private mWillRemoveNewlineCode As Boolean
Private mFilePath As String
Private mValidFormat As Boolean
Private mBOM As Boolean
Private mRowCount As Long
Private mColumnCount As Long
Private mProgressReporter As IProgressReporter
Private mReportUnit As Long
Private mItemNames As Collection

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* FilePath �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get FilePath() As String
    FilePath = mFilePath
End Property
Property Let FilePath(arg As String)
    mFilePath = arg
End Property

'*-----------------------------------------------------------------------------
'* Charset �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get CharSet() As String
    CharSet = mCharSet
End Property
Property Let CharSet(arg As String)
    mCharSet = arg
End Property

'*-----------------------------------------------------------------------------
'* LineSeparator �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get LineSeparator() As LineSeparatorsEnum
    LineSeparator = mLineSeparator
End Property
Property Let LineSeparator(arg As LineSeparatorsEnum)
    mLineSeparator = arg
End Property

'*-----------------------------------------------------------------------------
'* HeaderExists �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get HeaderExists() As Boolean
    HeaderExists = mHeaderExists
End Property
Property Let HeaderExists(arg As Boolean)
    mHeaderExists = arg
End Property

'*-----------------------------------------------------------------------------
'* ItemNames �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get ItemNames() As Collection
    Set ItemNames = mItemNames
End Property
Property Set ItemNames(arg As Collection)
    Set mItemNames = arg
End Property

'*-----------------------------------------------------------------------------
'* EnclosureChar �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get EnclosureChar() As String
    EnclosureChar = mEnclosureChar
End Property
Property Let EnclosureChar(arg As String)
    mEnclosureChar = arg
End Property

'*-----------------------------------------------------------------------------
'* Delimiter �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get Delimiter() As String
    Delimiter = mDelimiter
End Property
Property Let Delimiter(arg As String)
    mDelimiter = arg
End Property

'*-----------------------------------------------------------------------------
'* WillRemoveNewlineCode �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get WillRemoveNewlineCode() As Boolean
    Delimiter = mWillRemoveNewlineCode
End Property
Property Let WillRemoveNewlineCode(arg As Boolean)
    mWillRemoveNewlineCode = arg
End Property

'*-----------------------------------------------------------------------------
'* BOM �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get BOM() As Boolean
    BOM = mBOM
End Property
Property Let BOM(arg As Boolean)
    mBOM = arg
End Property

'*-----------------------------------------------------------------------------
'* ValidFormat �v���p�e�B�i�ǂݎ���p�j
'*-----------------------------------------------------------------------------
Property Get ValidFormat() As Boolean
    ValidFormat = mValidFormat
End Property

'*-----------------------------------------------------------------------------
'* ProgressReporter �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get ProgressReporter() As IProgressReporter
    Set ProgressReporter = mProgressReporter
End Property
Property Set ProgressReporter(arg As IProgressReporter)
    Set mProgressReporter = arg
End Property

'*-----------------------------------------------------------------------------
'* ReportUnit  �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get ReportUnit() As Long
    ReportUnit = mReportUnit
End Property
Property Let ReportUnit(arg As Long)
    mReportUnit = arg
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mAdoStream = CreateObject("ADODB.Stream")
    mCharSet = "UTF-8"
    mLineSeparator = LineSeparatorsEnum.adLF
    mHeaderExists = False
    mEnclosureChar = ""
    mDelimiter = ","
    mValidFormat = True
    mBOM = False
    mRowCount = 0
    mColumnCount = 0
    mReportUnit = 500
End Sub
Private Sub Class_Terminate()
    If Not mAdoStream Is Nothing Then
        If mAdoStream.State <> adStateClosed Then
            Call mAdoStream.Close
        End If
    End If
    Set mAdoStream = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] FileExists ���\�b�h
'* [��  ��] �t�@�C�������݂��邩�ǂ������肷��
'*
'* @param targetFilePath �t�@�C���p�X�B�ȗ��\�B
'* @return ���茋��
'*
'******************************************************************************
Public Function FileExists(Optional ByVal targetFilePath As String) As Boolean
    Dim tempPath As String: tempPath = mFilePath
    If targetFilePath <> "" Then
        tempPath = targetFilePath
    End If
    FileExists = FileUtils.FileExists(tempPath)
End Function

'******************************************************************************
'* [�T  �v] ShowCsvSaveFileDialog ���\�b�h
'* [��  ��] CSV�t�@�C����ۑ�����_�C�A���O��\������
'*
'* @param initialFile �����\���p�t�@�C���p�X�B�ȗ��\�B
'* @return �ۑ��t�@�C���̃p�X
'*
'******************************************************************************
Public Function ShowCsvSaveFileDialog(Optional ByVal initialFile As String) As String
    mFilePath = UXUtils.ShowCsvSaveFileDialog(initialFile)
    ShowCsvSaveFileDialog = mFilePath
End Function

'******************************************************************************
'* [�T  �v] �e�L�X�g�t�@�C�����������B
'* [��  ��] �e�L�X�g�t�@�C���Ƀf�[�^���s�P�ʂŏ����ށB
'*
'* @param records �����f�[�^�iRecordFormat/Array2DEx/Array(,)�j
'******************************************************************************
Public Sub WriteFile(records As Variant)

    ' RowEnumerator�̐ݒ�
    Dim rEnumerator As RowEnumerator: Set rEnumerator = New RowEnumerator
    Set rEnumerator.RowEditor = Me
    rEnumerator.SourceData = records
    rEnumerator.DestTarget = mAdoStream
    rEnumerator.IsActionable = False
    
    If ValidateUtils.IsArray2DEx(records) Then
        rEnumerator.MaxReportCount = records.RowLength
    ElseIf ArrayUtils.Is2DArray(records) Then
        rEnumerator.MaxReportCount = ArrayUtils.GetLength(records)
    ElseIf ValidateUtils.IsRecordFormat(records) Then
        rEnumerator.MaxReportCount = records.RecordSet.Count
    Else
        Err.Raise 9999, "CsvFileWriter#WriteFile", "�Ή����Ă��Ȃ������f�[�^�ł��B[" & TypeName(records) & "]"
    End If

    ' ProgressController�̐ݒ�
    Dim pController As ProgressController: Set pController = New ProgressController
    If mProgressReporter Is Nothing Then
        pController.IsReportable = False
    Else
        Set pController.ProgressReporter = mProgressReporter
        pController.IsReportable = True
        rEnumerator.ReportUnit = mReportUnit
    End If

    ' �����J�n
    On Error GoTo ErrorHandler
    pController.Start rEnumerator

    Exit Sub
ErrorHandler:
    If Not Err.Number = 9997 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

'******************************************************************************
'* [�T  �v] �w�b�_�ҏW�����B
'* [��  ��] ���R�[�h����CSV������i�w�b�_�j�ɕҏW����B
'*
'* @param records ���R�[�h���
'* @return CSV������i�w�b�_�j
'******************************************************************************
Private Function EditHeader(records) As String
    Dim itm As Item, ret As String, first As Boolean
    first = True
    Dim vData
    If ValidateUtils.IsArray2DEx(records) Or ArrayUtils.Is2DArray(records) Then
        Dim vName
        For Each vName In mItemNames
            If first Then
                ret = EditItem(CStr(vName))
                first = False
            Else
                ret = ret & mDelimiter & EditItem(CStr(vName))
            End If
        Next
    ElseIf ValidateUtils.IsRecordFormat(records) Then
        For Each itm In records.RecordDefinition
            If itm.OutputTarget Then
                If first Then
                    ret = EditItem(itm.Name)
                    first = False
                Else
                    ret = ret & mDelimiter & EditItem(itm.Name)
                End If
            End If
        Next
    End If
    EditHeader = ret
End Function

'******************************************************************************
'* [�T  �v] ���R�[�h�ҏW�����B
'* [��  ��] ���R�[�h����CSV������ɕҏW����B
'*
'* @param record ���R�[�h���
'* @return CSV������
'******************************************************************************
Private Function EditRecord(record) As String
    Dim sInfo As SettingInfo: Set sInfo = Base.GetSettingInfo()
           
    Dim itm As Item, ret As String, val As String, first As Boolean
    first = True
    If ValidateUtils.IsArrayEx(record) Then
        Dim colidx As Long
        For colidx = 0 To record.LastIndex
            If first Then
                ret = EditItem(CStr(record.Item(colidx)))
                first = False
            Else
                ret = ret & mDelimiter & EditItem(CStr(record.Item(colidx)))
            End If
        Next
    ElseIf ArrayUtils.Is1DArray(record) Then
        Dim v
        For Each v In record
            If first Then
                ret = EditItem(CStr(v))
                first = False
            Else
                ret = ret & mDelimiter & EditItem(CStr(v))
            End If
        Next
    ElseIf ValidateUtils.IsCollection(record) Then
        For Each itm In record
            If itm.OutputTarget Then
                itm.forOutput = True
                val = itm.EditedValue
                If first Then
                    ret = EditItem(val)
                    first = False
                Else
                    ret = ret & mDelimiter & EditItem(val)
                End If
            End If
        Next
    End If
    EditRecord = ret
End Function

'******************************************************************************
'* [�T  �v] ���ڒl�ҏW�����B
'* [��  ��] ���ڒl������s�R�[�h�����O���A�_�u���N�H�[�g���G�X�P�[�v����B
'*
'* @param val ���ڒl
'* @return �ҏW�㍀�ڒl
'******************************************************************************
Private Function EditItem(val As String) As String
    Dim ret As String
    ret = val
    If mWillRemoveNewlineCode Then
        ret = StringUtils.RemoveLineSeparator(ret)
    End If
    If mEnclosureChar = """" Then ret = StringUtils.EscapeDoubleQuote(ret)
    ret = mEnclosureChar & ret & mEnclosureChar
    EditItem = ret
End Function

'******************************************************************************
'* [�T  �v] StartProcess
'* [��  ��] �J�n���̏������s���B
'*
'******************************************************************************
Private Sub StartProcess(enumerator As Object)
    Dim rEnumerator As RowEnumerator
    Set rEnumerator = enumerator
    
    Dim records
    If IsObject(rEnumerator.SourceData) Then
        Set records = rEnumerator.SourceData
    Else
        records = rEnumerator.SourceData
    End If

    With rEnumerator.DestTarget
        .mode = adModeReadWrite
        .Type = adTypeText
        .CharSet = mCharSet
        .LineSeparator = mLineSeparator
        .Open
        Dim recText As String
        '�w�b�_�s�o��
        If mHeaderExists Then
            recText = EditHeader(records)
            .WriteText recText, adWriteLine
        End If
    End With

    ' �f�[�^�L���`�F�b�N
    Dim nodata As Boolean
    If ValidateUtils.IsArray2DEx(records) Then
        nodata = records.IsEmptyArray
    ElseIf ArrayUtils.Is2DArray(records) Then
        nodata = ArrayUtils.IsEmptyArray(records)
    ElseIf ValidateUtils.IsRecordFormat(records) Then
         nodata = CollectionUtils.IsEmpty(records.RecordSet)
    End If
    If nodata Then
        If UCase(mCharSet) = "UTF-8" And Not mBOM Then
            Call FileUtils.ExcludeBom(rEnumerator.DestTarget)
        End If
        With rEnumerator.DestTarget
            .SaveToFile mFilePath, adSaveCreateOverWrite '�t�@�C���㏑�w��
            .Close
        End With
        Err.Raise 9997, "CsvFileWriter#StartProcess", "��f�[�^�ł��B"
    End If
    
    If ValidateUtils.IsArray2DEx(records) Then
        rEnumerator.MaxReportCount = records.RowLength
        rEnumerator.StartRow = 0
        rEnumerator.EndRow = records.RowMaxIndex
        mRowCount = records.RowLength
        mColumnCount = records.ColLength
    ElseIf ArrayUtils.Is2DArray(records) Then
        rEnumerator.MaxReportCount = ArrayUtils.GetLength(records)
        rEnumerator.StartRow = LBound(records)
        rEnumerator.EndRow = UBound(records)
        mRowCount = ArrayUtils.GetLength(records)
        mColumnCount = ArrayUtils.GetLength(records, 2)
    ElseIf ValidateUtils.IsRecordFormat(records) Then
        rEnumerator.MaxReportCount = records.RecordSet.Count
        rEnumerator.StartRow = 1
        rEnumerator.EndRow = records.RecordSet.Count
        mRowCount = records.RecordSet.Count
        If mHeaderExists Then
            mColumnCount = records.RecordDefinition.Count
        Else
            mColumnCount = records.RecordSet(1).Count
        End If
    End If
End Sub

'******************************************************************************
'* [�T  �v] Edit
'* [��  ��] �s�ҏW�������s���B
'*
'******************************************************************************
Private Sub Edit(enumerator As Object)
    Dim rEnumerator As RowEnumerator
    Set rEnumerator = enumerator

    Dim vData, recText As String
    Dim record

    If ValidateUtils.IsArray2DEx(rEnumerator.SourceData) Then
        Set record = rEnumerator.SourceData.XRow(rEnumerator.CurrentRow)
    ElseIf ArrayUtils.Is2DArray(rEnumerator.SourceData) Then
        record = ArrayUtils.Copy2DArrayRow(rEnumerator.SourceData, rEnumerator.CurrentRow)
    ElseIf ValidateUtils.IsRecordFormat(rEnumerator.SourceData) Then
        Set record = rEnumerator.SourceData.RecordSet(rEnumerator.CurrentRow)
    End If

    recText = EditRecord(record)
    rEnumerator.DestTarget.WriteText recText, adWriteLine
End Sub

'******************************************************************************
'* [�T  �v] EndProcess
'* [��  ��] �I�����̏������s���B
'*
'******************************************************************************
Private Sub EndProcess(enumerator As Object)
    Dim rEnumerator As RowEnumerator
    Set rEnumerator = enumerator
    
    With rEnumerator.DestTarget
        If UCase(mCharSet) = "UTF-8" And Not mBOM Then
            Call FileUtils.ExcludeBom(rEnumerator.DestTarget)
        End If
        .SaveToFile mFilePath, adSaveCreateOverWrite '�t�@�C���㏑�w��
        .Close
    End With
End Sub

'******************************************************************************
'* �C���^�t�F�[�X����
'******************************************************************************
'*-----------------------------------------------------------------------------
'* IRowEditor#StartProcess
'*
'*-----------------------------------------------------------------------------
Private Sub IRowEditor_StartProcess(enumerator As Object)
    StartProcess enumerator
End Sub
'*-----------------------------------------------------------------------------
'* IRowEditor#Edit
'*
'*-----------------------------------------------------------------------------
Private Sub IRowEditor_Edit(enumerator As Object)
    Edit enumerator
End Sub
'*-----------------------------------------------------------------------------
'* IRowEditor#EndProcess
'*
'*-----------------------------------------------------------------------------
Private Sub IRowEditor_EndProcess(enumerator As Object)
    EndProcess enumerator
End Sub
