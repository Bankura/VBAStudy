Attribute VB_Name = "Main"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] �c�[�����C���������W���[��
'* [��  ��] �c�[���̃��C���������s�����W���[���B
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* �萔��`
'******************************************************************************
'�c�[����
Public Const TOOL_NAME As String = "���͎x���c�[��"
'�c�[���p�X���[�h
Public Const TOOL_PASSWORD As String = "1234"
'�c�[���V�[�g��
Public Const TOOL_SHEET_NAME As String = "data"
'����CSV�t�@�C���ݒ�V�[�g��
Public Const INPUTCSV_SHEET_NAME As String = "inputcsv_setting"
'�c�[���t�H�[���i�f�[�^�V�[�g�j�ݒ�V�[�g��
Public Const FORM_SHEET_NAME As String = "form_setting"
'HELP�i�g�����j�V�[�g��
Public Const HELP_SHEET_NAME As String = "help"
'�ݒ�V�[�g���
Private Const SETTING_SHEET_NAME As String = "setting"
Private Const SETTING_SH_START_ROW As Long = 4
Private Const SETTING_SH_START_COL As Long = 4

'******************************************************************************
'* Enum��`
'******************************************************************************
'�Ȃ�

'******************************************************************************
'* �ϐ���`
'******************************************************************************
Private mSettingInfo As SettingInfo
Private csStartRow As Long
Private csStartCol As Long
Private csItemCount As Long
Private fmStartRow As Long
Private fmStartCol As Long
Private fmItemCount As Long
Public dsStartRow As Long
Public dsStartCol As Long
Public dsKoubanCol As Long
Public dsItemCount As Long

'******************************************************************************
'* �֐���`
'******************************************************************************

'******************************************************************************
'* [�T  �v] �����������B
'* [��  ��] �������̏������s���B
'*
'******************************************************************************
Public Sub MyInit()
    Base.SettingSheetName = SETTING_SHEET_NAME
    Base.SettingSheetStartRow = SETTING_SH_START_ROW
    Base.SettingSheetStartCol = SETTING_SH_START_COL
    
    Set mSettingInfo = Base.GetSettingInfo
    Call Base.SaveApplicationProperties
    csStartRow = mSettingInfo.GetSettingValue("InputCsvSettingStartRowNo")
    csStartCol = mSettingInfo.GetSettingValue("InputCsvSettingStartColNo")
    csItemCount = mSettingInfo.GetSettingValue("InputCsvSettingItemCount")
    fmStartRow = mSettingInfo.GetSettingValue("FormSettingStartRowNo")
    fmStartCol = mSettingInfo.GetSettingValue("FormSettingStartColNo")
    fmItemCount = mSettingInfo.GetSettingValue("FormSettingItemCount")
    dsStartRow = mSettingInfo.GetSettingValue("DataSheetStartRowNo")
    dsStartCol = mSettingInfo.GetSettingValue("DataSheetStartColNo")
    dsKoubanCol = mSettingInfo.GetSettingValue("DataSheetKoubanColNo")
    dsItemCount = mSettingInfo.GetSettingValue("DataSheetItemCount")
End Sub

'******************************************************************************
'* [�T  �v] �t�@�C���Ǎ��{�^�������������B
'* [��  ��] CSV�t�@�C����ǂݍ��݃f�[�^�V�[�g�ɏo�͂���B
'*
'******************************************************************************
Public Sub ReadFileButton_Click()
    On Error GoTo ErrorHandler
    Call MyInit
    Dim mysheet As WorkSheetEx
    Set mysheet = Core.Init(New WorkSheetEx, TOOL_SHEET_NAME)
        
    Dim fReader As CsvFileReader
    Set fReader = New CsvFileReader
    fReader.QuotExists = True
    
    'Dim myReporter As SBProgressReporter: Set myReporter = New SBProgressReporter
    Dim myReporter As FormProgressReporter: Set myReporter = New FormProgressReporter
    myReporter.BaseMessage = "CSV�Ǎ�������"
    Set fReader.ProgressReporter = myReporter
    
    ' �_�C�A���O�\��
    fReader.ShowCsvFileDialog
    
    ' �t�@�C�����I������Ă���ΓǍ������s
    If fReader.FileExists Then
        Call MyStartProcess
        
        ' ����CSV��`���Ǎ�
        Dim rf As RecordFormat
        Set rf = GetDefinedRecordFormatFromSheet(INPUTCSV_SHEET_NAME, csStartRow, csStartCol, csItemCount)

        ' CSV�t�@�C���Ǎ�
        fReader.HeaderExists = True
        Dim csvData As Array2DEx: Set csvData = fReader.Read
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        DoEvents
        
        If csvData.IsEmptyArray() Then
            If fReader.ValidFormat Then
                Call Err.Raise(9999, "CSV�t�@�C���Ǎ�����", "�Ǎ��t�@�C���Ƀf�[�^������܂���B")
            Else
                Call Err.Raise(9999, "CSV�t�@�C���Ǎ�����", "�Ǎ��t�@�C���̃t�H�[�}�b�g���s���ł��B")
            End If
        End If

        ' �f�[�^����
        myReporter.BaseMessage = "CSV�f�[�^���ؒ�"
        Set rf.ProgressReporter = myReporter
        If Not rf.Validate(csvData) Then
            DebugUtils.Show rf.ErrMessage
            Call Err.Raise(9999, "CSV�t�@�C���Ǎ�����", "�Ǎ��t�@�C���̃t�H�[�}�b�g���s���ł��B")
        End If
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        DoEvents
        
        ' �t�H�[�����ڒ�`���Ǎ�
        Dim formRecDef As RecordFormat
        Set formRecDef = GetDefinedRecordFormatFromSheet(FORM_SHEET_NAME, fmStartRow, fmStartCol, fmItemCount)
        
        ' ���ڒ�`�Ɋ�Â�CSV�f�[�^���V�[�g�o�͗p�t�H�[���f�[�^�ɕϊ�
        myReporter.BaseMessage = "�V�[�g�f�[�^�ϊ�������"
        Set formRecDef.ProgressReporter = myReporter
        Dim formData As Array2DEx: Set formData = formRecDef.Convert(csvData)
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        DoEvents

        '�V�[�g���N���A
        Call mysheet.ClearActualUsedRange(dsStartRow, dsKoubanCol, dsItemCount)
        Call mysheet.DeleteNoUsedRange(dsStartRow)
        Call UXUtils.CheckEvents
        
        '�V�[�g�ɏo��
        Call mysheet.ImportArray(formData, dsStartRow, dsStartCol)
        Call UXUtils.CheckEvents
        Call mysheet.NumbersToIndexCells(dsStartRow, dsKoubanCol, formData.RowLength)
        Call UXUtils.CheckEvents

        Call MyEndProcess
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        mysheet.Cells(dsStartRow, dsStartCol).Select
        MsgBox "CSV�t�@�C���̓Ǎ����������܂����", vbOKOnly + vbInformation, TOOL_NAME
    End If

    Exit Sub
ErrorHandler:
    Call MyEndProcess
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �t�@�C���o�̓{�^�������������B
'* [��  ��] �f�[�^�V�[�g��ǂݍ���CSV�t�@�C���ɏo�͂���B
'*
'******************************************************************************
Public Sub OutputFileButton_Click()
    On Error GoTo ErrorHandler
    Call MyInit
    
    Dim mysheet As WorkSheetEx
    Set mysheet = Core.Init(New WorkSheetEx, TOOL_SHEET_NAME)
    
    ' ���ڒ�`���Ǎ�
    Dim rf As RecordFormat
    Set rf = GetDefinedRecordFormatFromSheet(FORM_SHEET_NAME, fmStartRow, fmStartCol, fmItemCount)
    
    ' ���ڃf�[�^�Ǎ��E�f�[�^����
    Dim formData As Array2DEx: Set formData = CheckDataSheet(mysheet, rf)
    If formData Is Nothing Then
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        Exit Sub
    End If

    ' �t�@�C���o�͏���
    Dim fWriter As CsvFileWriter
    Set fWriter = New CsvFileWriter
    fWriter.EnclosureChar = """"
    fWriter.WillRemoveNewlineCode = True

    ' �o�͐�I���_�C�A���O�\��
    If fWriter.ShowCsvSaveFileDialog <> "" Then
        Dim ret As Long: ret = vbYes
        If fWriter.FileExists Then
            ret = MsgBox("���Ƀt�@�C�������݂��܂��B" & vbCrLf & "�㏑�����Ă�낵���ł����B", vbYesNo + vbQuestion, TOOL_NAME)
        End If

        If ret = vbYes Then
            Call MyStartProcess
            
            ' ���ڃf�[�^�ҏW
            Dim myReporter As FormProgressReporter: Set myReporter = New FormProgressReporter
            myReporter.BaseMessage = "���ڃf�[�^�ҏW��"
            Set rf.ProgressReporter = myReporter
            Dim editedFormData As Array2DEx
            Set editedFormData = rf.Convert(formData, True)
            
            ' �t�@�C���o��
            'Dim myReporter As SBProgressReporter: Set myReporter = New SBProgressReporter
            myReporter.BaseMessage = "CSV�o�͏�����"
            Set fWriter.ProgressReporter = myReporter
            fWriter.HeaderExists = True
            Set fWriter.ItemNames = rf.ItemNames
            Call fWriter.WriteFile(editedFormData)

            Call MyEndProcess
            AppActivate ThisWorkbook.Name
            mysheet.Activate
            MsgBox "CSV�t�@�C���̏o�͂��������܂����", vbOKOnly + vbInformation, TOOL_NAME
        End If
    End If

    Exit Sub
ErrorHandler:
    Call MyEndProcess
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �s�폜�{�^�������������B
'* [��  ��] �f�[�^�V�[�g��ǂݍ���CSV�t�@�C���ɏo�͂���B
'*
'******************************************************************************
Public Sub DeleteRowButton_Click()
    On Error GoTo ErrorHandler
    Call MyInit
    Dim mysheet As WorkSheetEx
    Set mysheet = Core.Init(New WorkSheetEx, TOOL_SHEET_NAME)

    If ActiveWindow.RangeSelection.row >= dsStartRow Then
        Call MyStartProcess
        
        '�I��͈͂̍s���폜
        ActiveWindow.RangeSelection.EntireRow.Delete
        
        '���ԐU�蒼��
        Dim lMaxRow As Long: lMaxRow = mysheet.GetFinalKeyRow(dsKoubanCol)
        Call mysheet.NumbersToIndexCells(dsStartRow, dsKoubanCol, lMaxRow - dsStartRow + 1)
        
        Call MyEndProcess
    End If
    Exit Sub

ErrorHandler:
    Call MyEndProcess
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    Call MyErrorProcess
End Sub


'******************************************************************************
'* [�T  �v] �`�F�b�N�{�^�������������B
'* [��  ��] �f�[�^�̃`�F�b�N���s���B
'*
'******************************************************************************
Sub CheckButton_Click()
    On Error GoTo ErrorHandler
    Call MyInit
    
    Dim mysheet As WorkSheetEx
    Set mysheet = Core.Init(New WorkSheetEx, TOOL_SHEET_NAME)

    ' ���ڒ�`���Ǎ�
    Dim rf As RecordFormat
    Set rf = GetDefinedRecordFormatFromSheet(FORM_SHEET_NAME, fmStartRow, fmStartCol, fmItemCount)
    
    '���ڃf�[�^�Ǎ��E�f�[�^����
    Dim formData As Array2DEx: Set formData = CheckDataSheet(mysheet, rf)
    If formData Is Nothing Then
        Exit Sub
    End If

    '���b�Z�[�W�\��
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    mysheet.Cells(dsStartRow, dsStartCol).Select
    MsgBox "�`�F�b�N���������܂����" + vbNewLine + "��肠��܂���B", vbOKOnly + vbInformation, TOOL_NAME
    
    Exit Sub

ErrorHandler:
    Call MyEndProcess
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �N���A�{�^�������������B
'* [��  ��] �f�[�^�V�[�g���N���A����B
'*
'******************************************************************************
Public Sub ClearButton_Click()
    On Error GoTo ErrorHandler
    Call MyInit

    Call MyStartProcess
    Dim mysheet As WorkSheetEx
    Set mysheet = Core.Init(New WorkSheetEx, TOOL_SHEET_NAME)
    
    '�t�B���^�I������
    If Not mysheet.AutoFilter Is Nothing Then
        If mysheet.FilterMode Then
            mysheet.ShowAllData
        End If
    End If
    
    '�V�[�g���N���A
    Call mysheet.ClearActualUsedRange(dsStartRow, dsKoubanCol, dsItemCount)
    Call mysheet.DeleteNoUsedRange(dsStartRow)
    
    Call MyEndProcess
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    mysheet.Cells(dsStartRow, dsStartCol).Select

    Exit Sub
ErrorHandler:
    Call MyEndProcess
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �g�����{�^�������������B
'* [��  ��] help�V�[�g�ֈړ�����B
'*
'******************************************************************************
Sub GotoHelpButton_Click()
    On Error GoTo ErrorHandler

    Call XlWorkSheetUtils.GotoSheet(HELP_SHEET_NAME)
    Exit Sub

ErrorHandler:
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �߂�{�^�������������B
'* [��  ��] �f�[�^�V�[�g�ɖ߂�B
'*
'******************************************************************************
Sub ReturnFromHelpButton_Click()
    On Error GoTo ErrorHandler
    Call XlWorkSheetUtils.GotoSheet(TOOL_SHEET_NAME)
    
    Exit Sub

ErrorHandler:
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �ݒ�V�[�g�u�X�V�v�{�^�������������B
'* [��  ��] �ݒ�����X�V����B
'*
'******************************************************************************
Sub UpdateSettingButton_Click()
    On Error GoTo ErrorHandler
    Call MyInit
    Set mSettingInfo = New SettingInfo
    
    Exit Sub

ErrorHandler:
    Call MyErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �t�H�[���f�[�^�擾�E���؏����B
'* [��  ��] �t�H�[���i�V�[�g�j���獀�ڃf�[�^���擾�����ڒ�`�������ƂɌ��؂��s���B
'*
'* @param mysheet    ���[�N�V�[�g
'* @param rf         ���ڒ�`���
'* @return Array2DEx ���R�[�h�f�[�^���
'*
'******************************************************************************
Function CheckDataSheet(mysheet As WorkSheetEx, rf As RecordFormat) As Array2DEx
    Call MyInit
    Call MyStartProcess
    
    ' ���ڃf�[�^�Ǎ�
    Dim formData As Array2DEx
    Set formData = mysheet.GetActualUsedRangeToArray2DEx(dsStartRow, dsStartCol, rf.ColumnCount, dsKoubanCol)
    If formData.IsEmptyArray Then
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        Call MyEndProcess
        MsgBox "�f�[�^�����͂���Ă��܂���B", vbOKOnly + vbExclamation, TOOL_NAME
        Set CheckDataSheet = Nothing
        Exit Function
    End If

    ' ���ԐU�蒼��
    Call mysheet.ClearActualUsedRange(dsStartRow, dsKoubanCol, 1)
    Call mysheet.NumbersToIndexCells(dsStartRow, dsKoubanCol, formData.RowLength)
    
    ' �f�[�^����
    Dim myReporter As FormProgressReporter: Set myReporter = New FormProgressReporter
    myReporter.BaseMessage = "�f�[�^���ؒ�"
    Set rf.ProgressReporter = myReporter
    If Not rf.Validate(formData, True) Then
        AppActivate ThisWorkbook.Name
        mysheet.Activate
        mysheet.Cells(dsStartRow + rf.ErrRowNo - 1, dsStartCol + rf.ErrColNo - 1).Select
        Call MyEndProcess
        MsgBox rf.ErrMessage, vbOKOnly + vbExclamation, TOOL_NAME
        Set CheckDataSheet = Nothing
        Exit Function
    End If
    Set CheckDataSheet = formData
    AppActivate ThisWorkbook.Name
    mysheet.Activate
    Call MyEndProcess
    Exit Function
End Function

'******************************************************************************
'* [�T  �v] �G���[�����B
'* [��  ��] �G���[�������̏������s���B
'*
'******************************************************************************
Public Sub MyErrorProcess()
    Call ErrorProcess
    
    If Err.Number = 9999 Then
        MsgBox Err.Description, vbOKOnly + vbExclamation, TOOL_NAME
    ElseIf Err.Number = 3004 Then
        MsgBox "�t�@�C���֏������߂܂���ł����B" & vbNewLine & _
               "�ʃv���O�����Ńt�@�C�����J���Ă���Ȃǂ̌������l�����܂��B" & vbNewLine & _
                "���m�F���������B", vbOKOnly + vbExclamation, TOOL_NAME
    Else
        MsgBox "�V�X�e���G���[���������܂����", vbOKOnly + vbCritical, TOOL_NAME
    End If
End Sub

'******************************************************************************
'* [�T  �v] �J�n�����B
'* [��  ��] �����̃X�s�[�h����̂��߁AExcel�̐ݒ��ύX����B
'*
'******************************************************************************
Public Sub MyStartProcess()
    Call StartProcess

    '�V�[�g�ی����
    Call XlWorkSheetUtils.UnprotectSheet(TOOL_SHEET_NAME, TOOL_PASSWORD)
End Sub

'******************************************************************************
'* [�T  �v] �I�������B
'* [��  ��] �����̃X�s�[�h����̂��ߕύX����Excel�̐ݒ�����ɖ߂��B
'*
'******************************************************************************
Public Sub MyEndProcess()
    Call EndProcess
    
    '�V�[�g�ی�
    Call XlWorkSheetUtils.ProtectSheet(TOOL_SHEET_NAME, TOOL_PASSWORD)
End Sub

'******************************************************************************
'* [�T  �v] ���ڕ\���R�[�h��`���擾�E�ݒ菈���B
'* [��  ��] worksheet�̍��ڕ\���烌�R�[�h��`�����擾����
'*
'* @param defSheetName ���ڕ\���[�N�V�[�g��
'* @param lStartRow ���ڕ\�f�[�^�J�n�s�ԍ�
'* @param lStartCol ���ڕ\�f�[�^�J�n��ԍ�
'* @param itemCount ���ڗ�
'* @return RecordFormat ���R�[�h��`���
'*
'******************************************************************************
Public Function GetDefinedRecordFormatFromSheet(defSheetName As String, lStartRow As Long, lStartCol As Long, Optional colCount As Long) As RecordFormat
    Dim mysheet As WorkSheetEx
    Set mysheet = Core.Init(New WorkSheetEx, defSheetName)
    
    Dim vArr: vArr = mysheet.ExportArray(lStartRow, lStartCol, colCount)
    If IsEmpty(vArr) Then
        Set GetDefinedRecordFormatFromSheet = Nothing
        Exit Function
    End If
    Set GetDefinedRecordFormatFromSheet = GetDefinedRecordFormat(vArr)
End Function

'******************************************************************************
'* [�T  �v] ���R�[�h��`���擾�����B
'* [��  ��] ���ڒ�`���i2�����z��j���烌�R�[�h��`�����쐬�E�擾����
'*
'* @param vArr Variant�^2�����z��i���ڕ\�f�[�^�j
'* @return RecordFormat ���R�[�h��`���
'*
'******************************************************************************
Private Function GetDefinedRecordFormat(vArr) As RecordFormat
    Dim itmdefs As Collection: Set itmdefs = New Collection
    Dim i As Long
    For i = LBound(vArr, 1) To UBound(vArr, 1)
        itmdefs.Add DefineItem(vArr, i)
    Next
    Set GetDefinedRecordFormat = Core.Init(New RecordFormat, itmdefs)
End Function

'******************************************************************************
'* [�T  �v] ���ڐݒ菈���B
'* [��  ��] Item�ɒ�`����ݒ肷��
'*
'* @param vArr Variant�^2�����z��i���ڕ\�f�[�^�j
'* @param rownum �z��s�i1�����Y�����j
'* @return Item ��`�ςݍ���
'*
'******************************************************************************
Private Function DefineItem(vArr, rowNum As Long) As Item
    Dim itm As Item
    Set itm = New Item
    itm.Name = vArr(rowNum, 1)
    If vArr(rowNum, 2) = "��" Then
        itm.required = True
    End If
    Select Case vArr(rowNum, 3)
        Case "���p"
            itm.Attr = AttributeEnum.attrHalf
        Case "���p�p��"
            itm.Attr = AttributeEnum.attrHalfAlphaNumeric
        Case "���p�p���L��"
            itm.Attr = AttributeEnum.attrHalfAlphaNumericSymbol
        Case "���l"
            itm.Attr = AttributeEnum.attrNumeric
        Case "�S�p�J�^�J�i"
            itm.Attr = AttributeEnum.attrZenKatakana
        Case "�S�p�Ђ炪��"
            itm.Attr = AttributeEnum.attrZenHiragana
        Case "���t"
            itm.Attr = AttributeEnum.attrDate
        Case "�X�֔ԍ�"
            itm.Attr = AttributeEnum.attrZipCode
        Case "�d�b�ԍ�"
            itm.Attr = AttributeEnum.attrTelNo
        Case "���[���A�h���X"
            itm.Attr = AttributeEnum.attrMailAddress
        Case Else
            itm.Attr = AttributeEnum.attrString
    End Select
    Select Case vArr(rowNum, 4)
        Case "�Œ�"
            itm.KindOfDigits = KindOfDigitsEnum.digitFixed
        Case "�ȓ�"
            itm.KindOfDigits = KindOfDigitsEnum.digitWithin
        Case "�͈�"
            itm.KindOfDigits = KindOfDigitsEnum.digitRange
        Case Else
            itm.KindOfDigits = KindOfDigitsEnum.digitNone
    End Select
    If vArr(rowNum, 5) <> "" And VBA.IsNumeric(vArr(rowNum, 5)) Then
        itm.MinNumOfDigits = CLng(vArr(rowNum, 5))
    End If
    If vArr(rowNum, 6) <> "" And VBA.IsNumeric(vArr(rowNum, 6)) Then
        itm.MaxNumOfDigits = CLng(vArr(rowNum, 6))
    End If
    itm.Pattern = vArr(rowNum, 7)
    If UBound(vArr, 2) = 13 Then
        If vArr(rowNum, 8) <> "" And VBA.IsNumeric(vArr(rowNum, 8)) Then
            itm.InputColNo = vArr(rowNum, 8)
        End If
        Select Case vArr(rowNum, 9)
            Case "�}�X�^�ϊ��iCode��Value�j"
                itm.InitValueKind = EditKindEnum.mstCodeToValue
            Case "�}�X�^�ϊ��iValue��Code�j"
                itm.InitValueKind = EditKindEnum.mstValueToCode
            Case "�f�t�H���g"
                itm.InitValueKind = EditKindEnum.useDefaultValue
            Case Else
                itm.InitValueKind = EditKindEnum.edtNone
        End Select
        itm.InitValue = vArr(rowNum, 10)
        If vArr(rowNum, 11) = "��" Then
            itm.OutputTarget = True
        End If
        Select Case vArr(rowNum, 12)
            Case "�}�X�^�ϊ��iCode��Value�j"
                itm.OutputEditKind = EditKindEnum.mstCodeToValue
            Case "�}�X�^�ϊ��iValue��Code�j"
                itm.OutputEditKind = EditKindEnum.mstValueToCode
            Case "�f�t�H���g"
                itm.OutputEditKind = EditKindEnum.useDefaultValue
            Case Else
                itm.OutputEditKind = EditKindEnum.edtNone
        End Select
        itm.OutputEditValue = vArr(rowNum, 13)
    End If
    Set DefineItem = itm
End Function

