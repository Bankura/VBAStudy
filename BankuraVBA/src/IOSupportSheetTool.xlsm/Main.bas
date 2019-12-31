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
'* WinowsAPI�֐���`
'******************************************************************************
#If VBA7 Then
    '�v���O������C�ӂ̎��Ԃ����ҋ@������API�֐�
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
    '�C�x���g�L���[�ҋ@���̃C�x���g�`�F�b�NAPI
    Public Declare PtrSafe Function GetInputState Lib "user32" () As Long
#Else
    '�v���O������C�ӂ̎��Ԃ����ҋ@������API�֐�
    Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
    '�C�x���g�L���[�ҋ@���̃C�x���g�`�F�b�NAPI
    Public Declare Function GetInputState Lib "user32" () As Long
#End If

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

'******************************************************************************
'* Enum��`
'******************************************************************************
'�Ȃ�

'******************************************************************************
'* �ϐ���`
'******************************************************************************
Private mDisplayAlerts As Boolean
Private mScreenUpdating As Boolean
Private mCalculation As Long
Private mEnableEvents As Boolean
Private mRegExp As Object
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
Private dsItemCount As Long

Private mTime As Variant

'******************************************************************************
'* �֐���`
'******************************************************************************

'******************************************************************************
'* [�T  �v] �����������B
'* [��  ��] �������̏������s���B
'*
'******************************************************************************
Public Sub Init()
    Call GetSettingInfo
    Call SaveApplicationProperties
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
    Call Init

    Dim fReader As FileReader
    Set fReader = New FileReader

    '�_�C�A���O�\��
    fReader.ShowCsvFileDialog
    
    '�t�@�C�����I������Ă���ΓǍ������s
    If fReader.FileExists Then
        Call StartProcess
        
        '����CSV��`���Ǎ�
        Dim rf As RecordFormat: Set rf = New RecordFormat
        Call rf.GetItemDataFromSheet(ThisWorkbook.Sheets(INPUTCSV_SHEET_NAME), csStartRow, csStartCol, csItemCount)
        
        'CSV�t�@�C���Ǎ�
        fReader.HeaderExists = True
        Dim vArr: vArr = fReader.ReadTextFileToVArray
        
        If IsEmpty(vArr) Then
            If fReader.ValidFormat Then
                Call Err.Raise(9999, "CSV�t�@�C���Ǎ�����", "�Ǎ��t�@�C���Ƀf�[�^������܂���B")
            Else
                Call Err.Raise(9999, "CSV�t�@�C���Ǎ�����", "�Ǎ��t�@�C���̃t�H�[�}�b�g���s���ł��B")
            End If
        End If
        
        '�f�o�b�O�o��
        'Call PrintVariantArray(vArr)
        
        '�f�[�^����
        If Not rf.Validate(vArr) Then
            Call Err.Raise(9999, "CSV�t�@�C���Ǎ�����", "�Ǎ��t�@�C���̃t�H�[�}�b�g���s���ł��B")
        End If
        
        '�t�H�[�����ڒ�`���Ǎ�
        Dim formRecDef As RecordFormat: Set formRecDef = New RecordFormat
        Call formRecDef.GetItemDataFromSheet(ThisWorkbook.Sheets(FORM_SHEET_NAME), fmStartRow, fmStartCol, fmItemCount)
        
        '�t�H�[�����ڒ�`�Ɋ�Â�CSV�f�[�^���V�[�g�o�͗p�t�H�[���f�[�^�ɕϊ�
        Dim vFormArr: vFormArr = formRecDef.GetFormVariantData(vArr)
        Call CheckEvents

        '�V�[�g���N���A
        Dim mysheet As Worksheet
        Set mysheet = ThisWorkbook.Sheets(TOOL_SHEET_NAME)
        Call ClearActualUsedRangeFromSheet(mysheet, dsStartRow, dsKoubanCol, dsItemCount)
        Call DeleteNoUsedRange(mysheet, dsStartRow)
        Call CheckEvents
        
        '�V�[�g�ɏo��
        Call InjectVariantArrayToCells(mysheet, vFormArr, dsStartRow, dsStartCol)
        Call CheckEvents
        Call InjectNumbersToIndexCells(mysheet, dsStartRow, dsKoubanCol, UBound(vFormArr, 1) - LBound(vFormArr, 1) + 1)
        Call CheckEvents

        Call EndProcess
        
        mysheet.Cells(dsStartRow, dsStartCol).Select
        MsgBox "CSV�t�@�C���̓Ǎ����������܂����", vbOKOnly + vbInformation, TOOL_NAME
    End If

    Exit Sub
ErrorHandler:
    Call EndProcess
    Call ErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �t�@�C���o�̓{�^�������������B
'* [��  ��] �f�[�^�V�[�g��ǂݍ���CSV�t�@�C���ɏo�͂���B
'*
'******************************************************************************
Public Sub OutputFileButton_Click()
    On Error GoTo ErrorHandler
    Call Init
    
    Dim mysheet As Worksheet
    Set mysheet = ThisWorkbook.Sheets(TOOL_SHEET_NAME)
    
    Dim fWriter As FileWriter
    Set fWriter = New FileWriter
    
    '���ڃf�[�^�Ǎ��E�f�[�^����
    Dim rf As RecordFormat: Set rf = CheckDataSheet(mysheet)
    If rf Is Nothing Then
        Exit Sub
    End If

    '�f�o�b�O�o��
    'Call PrintRecordSet(rf)

    '�o�͐�I���_�C�A���O�\��
    If fWriter.ShowCsvSaveFileDialog <> "" Then
        Dim ret As Long: ret = vbYes
        If fWriter.FileExists Then
            ret = MsgBox("���Ƀt�@�C�������݂��܂��B" & vbCrLf & "�㏑�����Ă�낵���ł����B", vbYesNo + vbQuestion, TOOL_NAME)
        End If
    
        If ret = vbYes Then
            Call StartProcess
            
            '�t�@�C���o��
            fWriter.HeaderExists = True
            Call fWriter.WriteTextFileFromRecordSet(rf)

            Call EndProcess
            
            MsgBox "CSV�t�@�C���̏o�͂��������܂����", vbOKOnly + vbInformation, TOOL_NAME
        End If
    End If

    Exit Sub
ErrorHandler:
    Call EndProcess
    Call ErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �`�F�b�N�{�^�������������B
'* [��  ��] �f�[�^�̃`�F�b�N���s���B
'*
'******************************************************************************
Sub CheckButton_CLick()
    On Error GoTo ErrorHandler
    Call Init
    
    Dim mysheet As Worksheet
    Set mysheet = ThisWorkbook.Sheets(TOOL_SHEET_NAME)
    
    '���ڃf�[�^�Ǎ��E�f�[�^����
    Dim rf As RecordFormat: Set rf = CheckDataSheet(mysheet)
    If rf Is Nothing Then
        Exit Sub
    End If

    '���b�Z�[�W�\��
    mysheet.Cells(dsStartRow, dsStartCol).Select
    MsgBox "�`�F�b�N���������܂����" + vbNewLine + "��肠��܂���B", vbOKOnly + vbInformation, TOOL_NAME
    
    Exit Sub

ErrorHandler:
    Call EndProcess
    Call ErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �N���A�{�^�������������B
'* [��  ��] �f�[�^�V�[�g���N���A����B
'*
'******************************************************************************
Public Sub ClearButton_Click()
    On Error GoTo ErrorHandler
    Call Init

    '�V�[�g���N���A
    Call StartProcess
    Dim mysheet As Worksheet
    Set mysheet = ThisWorkbook.Sheets(TOOL_SHEET_NAME)
    Call ClearActualUsedRangeFromSheet(mysheet, dsStartRow, dsKoubanCol, dsItemCount)
    Call DeleteNoUsedRange(mysheet, dsStartRow)
    Call EndProcess
    mysheet.Cells(dsStartRow, dsStartCol).Select

    Exit Sub
ErrorHandler:
    Call EndProcess
    Call ErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �g�����{�^�������������B
'* [��  ��] help�V�[�g�ֈړ�����B
'*
'******************************************************************************
Sub GotoHelpButton_Click()
    On Error GoTo ErrorHandler

    Call GotoSheet(HELP_SHEET_NAME)
    Exit Sub

ErrorHandler:
    Call ErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �߂�{�^�������������B
'* [��  ��] �f�[�^�V�[�g�ɖ߂�B
'*
'******************************************************************************
Sub ReturnFromHelpButton_Click()
    On Error GoTo ErrorHandler
    Call GotoSheet(TOOL_SHEET_NAME)
    
    Exit Sub

ErrorHandler:
    Call ErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �ݒ�V�[�g�u�X�V�v�{�^�������������B
'* [��  ��] �ݒ�����X�V����B
'*
'******************************************************************************
Sub UpdateSettingButton_Click()
    On Error GoTo ErrorHandler
    Call Init
    Set mSettingInfo = New SettingInfo
    
    Exit Sub

ErrorHandler:
    Call ErrorProcess
End Sub

'******************************************************************************
'* [�T  �v] �t�H�[���f�[�^�擾�E���؏����B
'* [��  ��] �t�H�[���i�V�[�g�j���獀�ڃf�[�^���擾�����ڒ�`�������ƂɌ��؂��s���B
'*
'* @param mysheet ���[�N�V�[�g
'* @return ���R�[�h�f�[�^���
'*
'******************************************************************************
Function CheckDataSheet(mysheet As Worksheet) As RecordFormat
    Call Init
    Call StartProcess
    
    '���ڒ�`���Ǎ�
    Dim rf As RecordFormat: Set rf = New RecordFormat
    Call rf.GetItemDataFromSheet(ThisWorkbook.Sheets(FORM_SHEET_NAME), fmStartRow, fmStartCol, fmItemCount)

    '���ڃf�[�^�Ǎ�
    Dim readOk As Boolean: readOk = rf.GetRecordDataFromSheet(mysheet, dsStartRow, dsStartCol, rf.ColumnCount, dsKoubanCol)
    
    '���ԐU�蒼��
    Call ClearActualUsedRangeFromSheet(mysheet, dsStartRow, dsKoubanCol, 1)
    Call InjectNumbersToIndexCells(mysheet, dsStartRow, dsKoubanCol, rf.DataRowCount)
    
    '�f�[�^����
    If Not readOk Then
        Call EndProcess
        mysheet.Cells(6 + rf.ErrRowNo - 1, 3 + rf.ErrColNo - 1).Select
        MsgBox rf.ErrMessage, vbOKOnly + vbExclamation, TOOL_NAME
        Set CheckDataSheet = Nothing
        Exit Function
    End If
    Set CheckDataSheet = rf
    
    Call EndProcess
    Exit Function

End Function

'******************************************************************************
'* [�T  �v] �V�[�g�\��t�������B
'* [��  ��] Variant�z��f�[�^���V�[�g�ɏo�͂���B
'*
'* @param dataSheet ���[�N�V�[�g
'* @param vArray Variant�z��f�[�^
'* @param lStartRow �f�[�^�J�n�s�ԍ�
'* @param lStartCol �f�[�^�J�n��ԍ�
'*
'******************************************************************************
Private Sub InjectVariantArrayToCells(ByVal dataSheet As Worksheet, ByVal vArray, lStartRow As Long, lStartCol As Long)
    dataSheet.Cells(lStartRow, lStartCol).Resize(UBound(vArray, 1) + 1, UBound(vArray, 2) + 1).Value = vArray
End Sub

'******************************************************************************
'* [�T  �v] ���Ԑݒ菈���B
'* [��  ��] ���ԂɘA�Ԃ��o�͂���B
'*
'* @param dataSheet ���[�N�V�[�g
'* @param lStartRow �f�[�^�J�n�s�ԍ�
'* @param lStartCol �f�[�^�J�n��ԍ�
'* @param rowNum �ԍ���
'*
'******************************************************************************
Private Sub InjectNumbersToIndexCells(ByVal dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, rownum As Long)
    With dataSheet
        .Cells(lStartRow, lStartCol) = 1
        If rownum > 1 Then
            .Cells(lStartRow, lStartCol).AutoFill _
              Destination:=Range(.Cells(lStartRow, lStartCol), .Cells(lStartRow + rownum - 1, lStartCol)), Type:=xlLinearTrend
        End If
    End With
End Sub

'******************************************************************************
'* [�T  �v] �G���[�����B
'* [��  ��] �G���[�������̏������s���B
'*
'******************************************************************************
Public Sub ErrorProcess()
    Debug.Print "�G���[���� Number: " & Err.Number & " Source: " & Err.Source & " Description: " & Err.Description
    
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
Public Sub StartProcess()
    Call SaveApplicationProperties
    
    '�V�[�g�ی����
    Call UnprotectSheet
    
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
End Sub

'******************************************************************************
'* [�T  �v] �I�������B
'* [��  ��] �����̃X�s�[�h����̂��ߕύX����Excel�̐ݒ�����ɖ߂��B
'*
'******************************************************************************
Public Sub EndProcess()
    With Application
        .DisplayAlerts = mDisplayAlerts
        .ScreenUpdating = mScreenUpdating
        .Calculation = mCalculation
        .EnableEvents = mEnableEvents
        .StatusBar = False
    End With
    
    '�V�[�g�ی�
    Call ProtectSheet
End Sub

'******************************************************************************
'* [�T  �v] �V�[�g�ی���������B
'* [��  ��] �V�[�g�̕ی����������B
'*
'******************************************************************************
Public Sub UnprotectSheet()
    '�V�[�g�ی����
    If TOOL_PASSWORD = "" Then
        ThisWorkbook.Sheets(TOOL_SHEET_NAME).Unprotect
    Else
        ThisWorkbook.Sheets(TOOL_SHEET_NAME).Unprotect Password:=TOOL_PASSWORD
    End If
End Sub

'******************************************************************************
'* [�T  �v] �V�[�g�ی쏈���B
'* [��  ��] �V�[�g�̕ی������B
'*
'******************************************************************************
Public Sub ProtectSheet()
    With ThisWorkbook.Sheets(TOOL_SHEET_NAME)
        .EnableOutlining = True  '�A�E�g���C���L��
        .EnableAutoFilter = True '�I�[�g�t�B���^�L��
        
        '�V�[�g�ی�
        If TOOL_PASSWORD = "" Then
            .Protect Contents:=True, UserInterfaceOnly:=True
        Else
            .Protect Contents:=True, UserInterfaceOnly:=True, Password:=TOOL_PASSWORD
        End If
    End With
End Sub

'******************************************************************************
'* [�T  �v] Application�ݒ�ޔ������B
'* [��  ��] Application�̐ݒ�������o�ϐ��ɑޔ�����B
'*
'******************************************************************************
Public Sub SaveApplicationProperties()
    With Application
        mDisplayAlerts = .DisplayAlerts
        mScreenUpdating = .ScreenUpdating
        mCalculation = .Calculation
        mEnableEvents = .EnableEvents
    End With
End Sub

'******************************************************************************
'* [�T  �v] ���K�\���I�u�W�F�N�g�擾�����B
'* [��  ��] ���K�\���I�u�W�F�N�g���擾����B�������̏ꍇ��������B
'*
'******************************************************************************
Public Function GetRegExp() As Object
    If mRegExp Is Nothing Then
        Set mRegExp = CreateObject("VBScript.RegExp")
    End If
    Set GetRegExp = mRegExp
End Function

'******************************************************************************
'* [�T  �v] �ݒ���I�u�W�F�N�g�擾�����B
'* [��  ��] �ݒ���I�u�W�F�N�g���擾����B�������̏ꍇ��������B
'*
'******************************************************************************
Public Function GetSettingInfo() As SettingInfo
    If mSettingInfo Is Nothing Then
        Set mSettingInfo = New SettingInfo
    End If
    Set GetSettingInfo = mSettingInfo
End Function


'******************************************************************************
'* [�T  �v] GotoSheet
'* [��  ��] �A�N�e�B�u�ȃu�b�N�̎w�肵���V�[�g�E�A�h���X�ֈړ�����B
'*
'* @param sheetName �ړ���V�[�g��
'* @param strAddr �ړ���Z���̃A�h���X
'******************************************************************************
Public Sub GotoSheet(SheetName As String, Optional strAddr As String = "A1")
    ThisWorkbook.Activate
    ThisWorkbook.Worksheets(SheetName).Select
    ThisWorkbook.Worksheets(SheetName).Range(strAddr).Activate
End Sub


'******************************************************************************
'* [�T  �v] �\���擾�����B
'* [��  ��] worksheet�̕\��������擾���AVariant�z���ԋp���܂��
'*
'* @param dataSheet ���[�N�V�[�g
'* @param lStartRow �f�[�^�J�n�s�ԍ�
'* @param lStartCol �f�[�^�J�n��ԍ�
'* @param itemCount ���ڗ�
'*
'******************************************************************************
Public Function GetVariantDataFromSheet(dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, Optional colCount As Long)
    Dim lMaxRow As Long: lMaxRow = dataSheet.Cells(Rows.Count, lStartCol).End(xlUp).row
    Dim lMaxCol As Long
    If colCount = 0 Then
        lMaxCol = Cells(lStartRow, Columns.Count).End(xlToLeft).Column
    Else
        lMaxCol = lStartCol + colCount - 1
    End If
    
    '���R�[�h�����݂��Ȃ��ꍇ
    If lMaxRow < lStartRow Or lMaxCol < lStartCol Then
        GetVariantDataFromSheet = Empty
        Exit Function
    End If
    
    Dim vArr: vArr = dataSheet.Range(dataSheet.Cells(lStartRow, lStartCol), dataSheet.Cells(lMaxRow, lMaxCol))
    
    GetVariantDataFromSheet = vArr
End Function


'******************************************************************************
'* [�T  �v] �g�p�Z���͈̓N���A�����B
'* [��  ��] worksheet�̃f�[�^�\�̎g�p�Z���͈͂��N���A���܂��
'*
'* @param dataSheet data�\���[�N�V�[�g
'* @param lStartRow data�\�f�[�^�J�n�s�ԍ�
'* @param lStartCol data�\�f�[�^�J�n��ԍ�
'* @param itemCount ���ڗ�
'* @param ignoreColnum �����ΏۊO�̗�ԍ�
'*
'******************************************************************************
Public Sub ClearActualUsedRangeFromSheet(dataSheet As Worksheet, _
                                         lStartRow As Long, _
                                         lStartCol As Long, _
                                         Optional colCount As Long, _
                                         Optional ignoreColnum As Long)
    Dim rng As Range
    Set rng = GetActualUsedRangeFromSheet(dataSheet, lStartRow, lStartCol, colCount, ignoreColnum)
    If rng Is Nothing Then
        Exit Sub
    End If
    rng.ClearContents
End Sub

'******************************************************************************
'* [�T  �v] ���g�p�͈͍s�폜�����B
'* [��  ��] worksheet�̃f�[�^�\�̖��g�p�͈͍s���폜���܂��iUsedRange���k���j�
'*
'* @param dataSheet data�\���[�N�V�[�g
'* @param lStartRow data�\�f�[�^�J�n�s�ԍ�
'*
'******************************************************************************
Public Sub DeleteNoUsedRange(dataSheet As Worksheet, lStartRow As Long)
    Dim delStartRow As Long
    Dim delEndRow As Long
    
    Dim rng As Range
    Set rng = GetActualUsedRangeFromSheet(dataSheet, lStartRow, 1)
    If rng Is Nothing Then
        delStartRow = lStartRow
    Else
        delStartRow = rng.Item(rng.Count).row + 1
    End If
    delEndRow = dataSheet.UsedRange.Item(dataSheet.UsedRange.Count).row
    
    If delStartRow > delEndRow Then
        Exit Sub
    End If
    With dataSheet
        .Range(.Rows(delStartRow), .Rows(delEndRow)).Delete
    End With
End Sub

'******************************************************************************
'* [�T  �v] �g�p�Z���͈͎擾�����B
'* [��  ��] worksheet�̃f�[�^�\�̎g�p�Z���͈͂��擾����
'*
'* @param dataSheet data�\���[�N�V�[�g
'* @param lStartRow data�\�f�[�^�J�n�s�ԍ�
'* @param lStartCol data�\�f�[�^�J�n��ԍ�
'* @param itemCount ���ڗ�
'* @param ignoreColnum �����ΏۊO�̗�ԍ�
'* @return �g�p�Z���͈�
'*
'******************************************************************************
Public Function GetActualUsedRangeFromSheet(dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, Optional colCount As Long, Optional ignoreColnum As Long) As Range
    Dim lMaxRow As Long: lMaxRow = GetFinalRow(dataSheet, ignoreColnum)
    Dim lMaxCol As Long
    If colCount = 0 Then
        lMaxCol = GetFinalCol(dataSheet)
    Else
        lMaxCol = lStartCol + colCount - 1
    End If

    '���R�[�h�����݂��Ȃ��ꍇ
    If lMaxRow < lStartRow Or lMaxCol < lStartCol Then
        Set GetActualUsedRangeFromSheet = Nothing
        Exit Function
    End If
    
    Set GetActualUsedRangeFromSheet = dataSheet.Range(dataSheet.Cells(lStartRow, lStartCol), dataSheet.Cells(lMaxRow, lMaxCol))
End Function

'******************************************************************************
'* [�T  �v] �ŏI�s�擾�����B
'* [��  ��] Worksheet��UsedRange�������瑖�����A�ŏI�s�ԍ����擾����
'*
'* @param dataSheet ���[�N�V�[�g
'* @param ignoreColnum �����ΏۊO�̗�ԍ�
'* @return �ŏI�s�ԍ�
'*
'******************************************************************************
Public Function GetFinalRow(ByVal dataSheet As Worksheet, Optional ignoreColnum As Long) As Long
    Dim ret As Long
    Dim i As Long, cnta As Long
    With dataSheet.UsedRange
        For i = .Rows.Count To 1 Step -1
            cnta = WorksheetFunction.counta(.Rows(i))
            If cnta > 0 Then
                If cnta <> 1 Then
                    ret = i
                    Exit For
                Else
                    If ignoreColnum > 0 Then
                        If .Cells(i, ignoreColnum) = "" Then
                            ret = i
                            Exit For
                        End If
                    Else
                        ret = i
                        Exit For
                    End If
                End If
            End If
        Next
        If ret > 0 Then
            ret = ret + .row - 1
        End If
    End With
    GetFinalRow = ret
End Function

'******************************************************************************
'* [�T  �v] �ŏI��擾�����B
'* [��  ��] Worksheet��UsedRange���E���瑖�����A�ŏI��ԍ����擾����
'*
'* @param dataSheet ���[�N�V�[�g
'* @return �ŏI��ԍ�
'*
'******************************************************************************
Public Function GetFinalCol(ByVal dataSheet As Worksheet) As Long
    Dim ret As Long
    Dim i As Long
    With dataSheet.UsedRange
        For i = .Columns.Count To 1 Step -1
            If WorksheetFunction.counta(.Columns(i)) > 0 Then
                ret = i
                Exit For
            End If
        Next
        If ret > 0 Then
            ret = ret + .Column - 1
        End If
    End With
    GetFinalCol = ret
End Function


'******************************************************************************
'* [�T  �v] Variant�z��f�o�b�O�o�͏���
'* [��  ��] Variant�z��̓��e���C�~�f�B�G�C�g�E�B���h�E�ɏo�͂���B
'*
'* @param vArr Variant�z��
'******************************************************************************
Private Sub PrintVariantArray(vArr)
    Dim i As Long, j As Long, tmp As String
    For i = LBound(vArr, 1) To UBound(vArr, 1)
        For j = LBound(vArr, 2) To UBound(vArr, 2)
            tmp = tmp & " | " & vArr(i, j)
        Next
        Debug.Print tmp
        tmp = ""
    Next
End Sub

'******************************************************************************
'* [�T  �v] Variant�z��f�o�b�O�o�͏���
'* [��  ��] Variant�z��̓��e���C�~�f�B�G�C�g�E�B���h�E�ɏo�͂���B
'*
'* @param vArr Variant�z��
'******************************************************************************
Private Sub PrintRecordSet(rf As RecordFormat)
    Dim record As Collection, itm As Item, tmp As String
    For Each record In rf.RecordSet
        For Each itm In record
            tmp = tmp & " | " & itm.Value
        Next
        Debug.Print tmp
        tmp = ""
    Next
End Sub

Public Sub CheckEvents()
    If GetInputState() Or (DateDiff("s", mTime, Time) > 3) Then
        DoEvents
        mTime = Time
    End If
End Sub
