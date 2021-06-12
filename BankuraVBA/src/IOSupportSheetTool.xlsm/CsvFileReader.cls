VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvFileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IRowEditor

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] CSVFile�Ǎ��N���X
'* [��  ��] CSVFile�Ǎ��p�̃��[�e�B���e�B��񋟂���B
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
Private mSize As Long
Private mHeaderExists As Boolean
Private mQuotExists As Boolean
Private mFilePath As String
Private mValidFormat As Boolean
Private mRowCount As Long
Private mColumnCount As Long
Private mProgressReporter As IProgressReporter
Private mReportUnit As Long

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
    If FileExists(arg) Then
        mSize = IO.GetFile(mFilePath).size
    End If
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
'* Size �v���p�e�B�i�ǂݎ���p�j
'*-----------------------------------------------------------------------------
Property Get size() As Long
    size = mSize
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
'* QuotExists �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get QuotExists() As Boolean
    QuotExists = mQuotExists
End Property
Property Let QuotExists(arg As Boolean)
    mQuotExists = arg
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
    mLineSeparator = adLF
    mHeaderExists = False
    mQuotExists = False
    mValidFormat = True
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
'* [��  ��] �t�@�C�������݂��邩�ǂ������肵�܂��
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
'* [�T  �v] IsEmpty ���\�b�h
'* [��  ��] �t�@�C������t�@�C�����ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Public Function IsEmpty() As Boolean
    IsEmpty = (mSize = 0)
End Function

'******************************************************************************
'* [�T  �v] ShowCsvFileDialog ���\�b�h
'* [��  ��] CSV�t�@�C����I������_�C�A���O��\�����܂��
'*
'* @param initialFile �����\���p�t�@�C���p�X�B�ȗ��\�B
'* @return �I�������t�@�C���̃p�X
'*
'******************************************************************************
Public Function ShowCsvFileDialog(Optional ByVal initialFile As String) As String
    mFilePath = UXUtils.ShowCsvFileDialog(initialFile)
    mSize = IO.GetFile(mFilePath).size
    ShowCsvFileDialog = mFilePath
End Function

'******************************************************************************
'* [�T  �v] CSV�t�@�C���Ǎ������B
'* [��  ��] CSV�t�@�C�����s�P�ʂœǍ��݁AArray2DEx�I�u�W�F�N�g�ɐݒ肵�ĕԋp����B
'*
'* @return Array2DEx �Ǎ��e�L�X�g�f�[�^
'******************************************************************************
Public Function Read() As Array2DEx
    Dim arr2dex As Array2DEx: Set arr2dex = New Array2DEx
    Dim lSize As Long
    With mAdoStream
        .mode = adModeReadWrite
        .Type = adTypeText
        .CharSet = mCharSet
        .LineSeparator = mLineSeparator
        .Open
        .LoadFromFile mFilePath
        lSize = Me.CountLine
        
        ' ��t�@�C���͋�Variant�z��ŕԋp
        If (mHeaderExists And lSize <= 1) Or (Not mHeaderExists And lSize = 0) Then
            .Close
            Set Read = arr2dex
            Exit Function
        End If
        .Position = 0
    End With

    ' RowEnumerator�̐ݒ�
    Dim rEnumerator As RowEnumerator: Set rEnumerator = New RowEnumerator
    Set rEnumerator.RowEditor = Me
    rEnumerator.SourceData = mAdoStream
    rEnumerator.DestTarget = arr2dex
    rEnumerator.IsActionable = False
    rEnumerator.MaxReportCount = lSize
    rEnumerator.EndRow = lSize
    mRowCount = lSize
    
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

    Set Read = arr2dex
    Exit Function
ErrorHandler:
    If Err.Number = 9998 Then
        mValidFormat = False
    ElseIf Err.Number = 9997 Then
        Set Read = arr2dex
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

'******************************************************************************
'* [�T  �v] CSV�t�@�C���Ǎ������B
'* [��  ��] CSV�t�@�C�����s�P�ʂœǍ��݁AVariant�z��ɐݒ肵�ĕԋp����B
'*
'* @return �Ǎ��e�L�X�g�f�[�^�iVariant�z��j
'******************************************************************************
Public Function ReadToVArray() As Variant
    Dim arr2dex As Array2DEx
    Set arr2dex = Read()
    ReadToVArray = arr2dex.To2DArray
End Function

'******************************************************************************
'* [�T  �v] CSV�s���������B
'* [��  ��] CSV��������J���}���Ƃɕ�������Variant�z��Ɋi�[���A�ԋp����B
'*
'* @param line CSV�s
'* @param quotExist �_�u���N�H�[�g���ݗL���iTrue�F����j�B�ȗ��\�B
'* @param delimiter �f���~�^�B�ȗ��\�B
'* @return ������Variant�z��
'******************************************************************************
Public Function SplitCsvRow(line As String, Optional ByVal QuotExists As Boolean = False, Optional Delimiter As String = ",") As Variant
    SplitCsvRow = StringUtils.SplitCsvRow(line, QuotExists, Delimiter)
End Function

'******************************************************************************
'* [�T  �v] �e�L�X�g�̍s���擾�����B
'* [��  ��] �e�L�X�g�̍s�����擾����B
'* [���@�l] �Ăяo���O��Stream��Open����Ă��邱�ƁB
'*
'* @return �s��
'******************************************************************************
Public Function CountLine() As Long
    CountLine = FileUtils.CountLine(mAdoStream)
End Function

'******************************************************************************
'* [�T  �v] StartProcess
'* [��  ��] �J�n���̏������s���B
'*
'******************************************************************************
Private Sub StartProcess(enumerator As Object)
    Dim rEnumerator As RowEnumerator, arr2dex As Array2DEx
    Set rEnumerator = enumerator

    Set arr2dex = rEnumerator.DestTarget
    Dim vRow
    With rEnumerator.SourceData
        ' 1�s�ڏ���
        vRow = SplitCsvRow(.ReadText(adReadLine), True)
         If Not mHeaderExists Then
            arr2dex.AddRow vRow
        End If
        mColumnCount = UBound(vRow)
    End With
    rEnumerator.StartRow = 2
End Sub

'******************************************************************************
'* [�T  �v] Edit
'* [��  ��] �s�ҏW�������s���B
'*
'******************************************************************************
Private Sub Edit(enumerator As Object)
    Dim rEnumerator As RowEnumerator, arr2dex As Array2DEx
    Set rEnumerator = enumerator
    Set arr2dex = rEnumerator.DestTarget

    Dim vRow
    With rEnumerator.SourceData
        vRow = SplitCsvRow(.ReadText(adReadLine), mQuotExists)
        If UBound(vRow) <> mColumnCount Then
            .Close
            Err.Raise 9998, "CsvFileReader#Edit", "�s���̍��ڐ����s�����ł��B"
        End If
        arr2dex.AddRow vRow
    End With
End Sub

'******************************************************************************
'* [�T  �v] EndProcess
'* [��  ��] �I�����̏������s���B
'*
'******************************************************************************
Private Sub EndProcess(enumerator As Object)
    enumerator.SourceData.Close
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