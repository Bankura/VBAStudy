VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] File�Ǎ��N���X
'* [��  ��] File�Ǎ��p�̃��[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* Enum��`
'******************************************************************************
'*-----------------------------------------------------------------------------
'* �I�u�W�F�N�g���J���Ă��邩���Ă��邩�A�f�[�^ �\�[�X�ɐڑ������A�R�}���h��
'* ���s�����A�܂��̓f�[�^���擾������\���܂��B
'*-----------------------------------------------------------------------------
Private Enum ObjectStateEnum
    adStateClosed = 0     '�I�u�W�F�N�g�����Ă��邱�Ƃ������܂��B
    adStateOpen = 1       '�I�u�W�F�N�g���J���Ă��邱�Ƃ������܂��B
    adStateConnecting = 2 '�I�u�W�F�N�g���ڑ����ł��邱�Ƃ������܂��B
    adStateExecuting = 4  '�I�u�W�F�N�g���R�}���h�����s���ł��邱�Ƃ������܂��B
    adStateFetching = 8   '�I�u�W�F�N�g�̍s���擾���ł��邱�Ƃ������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Stream �I�u�W�F�N�g�ɕۑ�����f�[�^�̎�ނ�\���܂��B
'*
'*-----------------------------------------------------------------------------
Private Enum StreamTypeEnum
    adTypeBinary = 1 '�o�C�i�� �f�[�^�������܂��B
    adTypeText = 2   '����l�BCharset �Ŏw�肳�ꂽ�����Z�b�g�̃e�L�X�g �f�[�^�������܂��B
End Enum

'*-----------------------------------------------------------------------------
'* Stream �I�u�W�F�N�g����A�X�g���[���S�̂�ǂݎ�邩�A�܂��͎��̍s��ǂݎ�邩��\���܂��B
'*
'*-----------------------------------------------------------------------------
Private Enum StreamReadEnum
    adReadAll = -1  '����l�B���݂̈ʒu���� EOS �}�[�J�[�����ɁA���ׂẴo�C�g���X�g���[������ǂݎ��܂��B����́A�o�C�i�� �X�g���[�� (Type �� adTypeBinary) �ɗB��L���� StreamReadEnum �l�ł��B
    adReadLine = -2 '�X�g���[�����玟�̍s��ǂݎ��܂� (LineSeparator �v���p�e�B�Ŏw��)�B
End Enum

'*-----------------------------------------------------------------------------
'* Connection ���̃f�[�^�̕ҏW�ARecord �̃I�[�v���A�܂��� Record ����� Stream
'*  �I�u�W�F�N�g�� Mode �v���p�e�B�̒l�̎w��ɑ΂��錠����\���܂��B
'*-----------------------------------------------------------------------------
Private Enum ConnectModeEnum
    adModeRead = 1
    adModeReadWrite = 3
    adModeRecursive = 4194304
    adModeShareDenyNone = 16
    adModeShareDenyRead = 4
    adModeShareDenyWrite = 8
    adModeShareExclusive = 12
    adModeUnknown = 0
    adModeWrite = 2
End Enum

'*-----------------------------------------------------------------------------
'* �e�L�X�g Stream �I�u�W�F�N�g�̍s��؂�L���Ɏg���Ă��镶����\���܂��B
'*
'*-----------------------------------------------------------------------------
Public Enum LineSeparatorEnum
    adCR = 13
    adCRLF = -1
    adLF = 10
End Enum

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mFso As Object
Private mFile As Object
Private mAdoStream As Object
Private mCharSet As String
Private mLineSeparator As LineSeparatorEnum
Private mSize As Long
Private mHeaderExists As Boolean
Private mFilePath As String
Private mValidFormat As Boolean

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* FilePath �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get filePath() As String
    filePath = mFilePath
End Property
Property Let filePath(arg As String)
    mFilePath = arg
    If FileExists(arg) Then
        Set mFile = mFso.GetFile(mFilePath)
        mSize = mFile.Size
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
Property Get LineSeparator() As LineSeparatorEnum
    LineSeparator = mLineSeparator
End Property
Property Let LineSeparator(arg As LineSeparatorEnum)
    mLineSeparator = arg
End Property

'*-----------------------------------------------------------------------------
'* Size �v���p�e�B�i�ǂݎ���p�j
'*-----------------------------------------------------------------------------
Property Get Size() As Long
    Size = mSize
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
'* ValidFormat �v���p�e�B�i�ǂݎ���p�j
'*-----------------------------------------------------------------------------
Property Get ValidFormat() As Boolean
    ValidFormat = mValidFormat
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mFso = CreateObject("Scripting.FileSystemObject")
    Set mAdoStream = CreateObject("ADODB.Stream")
    mCharSet = "UTF-8"
    mLineSeparator = adLF
    mHeaderExists = False
    mValidFormat = True
End Sub
Private Sub Class_Terminate()
    If Not mAdoStream Is Nothing Then
        If mAdoStream.State <> adStateClosed Then
            Call mAdoStream.Close
        End If
    End If
    Set mAdoStream = Nothing
    Set mFso = Nothing
    Set mFile = Nothing
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
    FileExists = mFso.FileExists(tempPath)
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
    ShowCsvFileDialog = ""
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "CSV�t�@�C��", "*.csv"
        .InitialFileName = initialFile
        .AllowMultiSelect = False
        If .Show = True Then
            mFilePath = .SelectedItems(1)
            Set mFile = mFso.GetFile(mFilePath)
            mSize = mFile.Size
            ShowCsvFileDialog = mFilePath
        End If
    End With
End Function


'******************************************************************************
'* [�T  �v] �e�L�X�g�t�@�C���Ǎ������B
'* [��  ��] �e�L�X�g�t�@�C�����s�P�ʂœǍ��݁AVariant�z��ɐݒ肵�ĕԋp����B
'*
'* @return �Ǎ��e�L�X�g�f�[�^�iVariant�z��j
'******************************************************************************
Public Function ReadTextFileToVArray() As Variant

    Dim lSize As Long, lCount As Long, i As Long, vRow, vArr()
    With mAdoStream
        .Mode = adModeReadWrite
        .Type = adTypeText
        .CharSet = mCharSet
        .LineSeparator = mLineSeparator
        .Open
        .LoadFromFile mFilePath
        lSize = Me.CountLine
        
        '��t�@�C���͋�Variant�z��ŕԋp
        If (mHeaderExists And lSize <= 1) Or (Not mHeaderExists And lSize = 0) Then
            .Close
            Exit Function
        End If
        '1�s�ڏ���
        .Position = 0
        vRow = SplitCsvRow(.ReadText(adReadLine), True)
         If Not mHeaderExists Then
            ReDim vArr(0 To lSize - 1, 0 To UBound(vRow))
            For i = LBound(vRow) To UBound(vRow)
                vArr(lCount, i) = vRow(i)
            Next
            lCount = lCount + 1
        Else
            ReDim vArr(0 To lSize - 2, 0 To UBound(vRow))
        End If
        Dim ColumnCount As Long: ColumnCount = UBound(vRow)
        
        Dim statusBarTxt As String: statusBarTxt = "CSV�Ǎ��������F"
        Application.StatusBar = statusBarTxt
         
        '2�s�ڈȍ~����
        Do While Not .EOS
            vRow = SplitCsvRow(.ReadText(adReadLine), True)
            If UBound(vRow) <> ColumnCount Then
                .Close
                ReadTextFileToVArray = Empty
                mValidFormat = False
                Exit Function
            End If

            If lCount Mod 5000 = 0 Then
                statusBarTxt = statusBarTxt & "��"
                Application.StatusBar = statusBarTxt
                Call CheckEvents
            End If
            
            For i = LBound(vRow) To UBound(vRow)
                vArr(lCount, i) = vRow(i)
            Next
            lCount = lCount + 1
        Loop
        .Close
    End With

    ReadTextFileToVArray = vArr
End Function

'******************************************************************************
'* [�T  �v] CSV�s���������B
'* [��  ��] CSV��������J���}���Ƃɕ�������Variant�z��Ɋi�[���A�ԋp����B
'*          ���ڗ��[�́u"�v���폜����B
'*
'* @param line CSV�s
'* @param quotExist �_�u���N�H�[�g���ݗL���iTrue�F����j�B�ȗ��\�B
'* @param delimiter �f���~�^�B�ȗ��\�B
'* @return ������Variant�z��
'******************************************************************************
Public Function SplitCsvRow(line As String, Optional ByVal quotExist As Boolean = False, Optional delimiter As String = ",") As Variant
    If quotExist Then
        Dim lQuoteCount As Long, lItemCount As Long, i As Long, tmpTxt As String, v()
        For i = 1 To Len(line)
            Select Case Mid(line, i, 1)
                Case delimiter
                    If lQuoteCount Mod 2 = 0 Then
                        lItemCount = lItemCount + 1
                    End If
                Case """"
                    lQuoteCount = lQuoteCount + 1
            End Select
        Next
        ReDim v(lItemCount) As Variant
        lQuoteCount = 0: lItemCount = 0
        For i = 1 To Len(line)
            Select Case Mid(line, i, 1)
                Case delimiter
                    If lQuoteCount Mod 2 = 0 Then
                        v(lItemCount) = EditIncludeQuotItem(tmpTxt)
                        tmpTxt = ""
                        lItemCount = lItemCount + 1
                    Else
                        tmpTxt = tmpTxt & Mid(line, i, 1)
                    End If
                Case """"
                    lQuoteCount = lQuoteCount + 1
                    tmpTxt = tmpTxt & Mid(line, i, 1)
                Case Else
                    tmpTxt = tmpTxt & Mid(line, i, 1)
            End Select
        Next
        If tmpTxt <> "" Then
            v(lItemCount) = EditIncludeQuotItem(tmpTxt)
        End If
        SplitCsvRow = v
    Else
        SplitCsvRow = Split(line, delimiter)
    End If
End Function

'******************************************************************************
'* [�T  �v] �_�u���N�H�[�g���荀�ڕҏW�����B
'* [��  ��] �G�X�P�[�v����Ă���_�u���N�H�[�g�u""�v���u"�v�ɕϊ����A
'*          ���ڗ��[�́u"�v���폜����B
'*
'* @param text �_�u���N�H�[�g���荀��
'* @return �ҏW�㍀��
'******************************************************************************
Private Function EditIncludeQuotItem(ByVal text As String) As String
    text = Replace(text, """""", """")
    If Left(text, 1) = """" And Right(text, 1) = """" Then
        If Len(text) <= 2 Then
            EditIncludeQuotItem = ""
        Else
            EditIncludeQuotItem = Mid(text, 2, Len(text) - 2)
            
        End If
    End If
End Function

'******************************************************************************
'* [�T  �v] �e�L�X�g�̍s���擾�����B
'* [��  ��] �e�L�X�g�̍s�����擾����B
'* [���@�l] �Ăяo���O��Stream��Open����Ă��邱�ƁB
'*
'* @return �s��
'******************************************************************************
Public Function CountLine() As Long
    Dim lineCount As Long
    With mAdoStream
        .Position = 0
        Do While Not .EOS
            .SkipLine
            lineCount = lineCount + 1
        Loop
    End With

    CountLine = lineCount
End Function
