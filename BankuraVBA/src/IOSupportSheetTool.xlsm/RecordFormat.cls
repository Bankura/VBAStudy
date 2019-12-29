VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] ���R�[�h�t�H�[�}�b�g�N���X
'* [��  ��] ���ڂ̃t�H�[�}�b�g��`�������R�[�h�P�ʂő��˂��N���X�B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mCol As Collection
Private mRecordSetCol As Collection
Private mColumnCount As Long
Private mDataRowCount As Long

Private mErrMessage As String
Private mErrRowNo As Long
Private mErrColNo As Long

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* Item �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get Item(index) As Item
    Item = mCol.Item(index)
End Property
Property Let Item(index, arg As Item)
    mCol.Item(index) = arg
End Property

'*-----------------------------------------------------------------------------
'* RecordSet �v���p�e�B�i�ǂݎ���p�j
'*-----------------------------------------------------------------------------
Property Get RecordSet() As Collection
    Set RecordSet = mRecordSetCol
End Property

'*-----------------------------------------------------------------------------
'* RecordDefinition �v���p�e�B�i�ǂݎ���p�j
'*-----------------------------------------------------------------------------
Property Get RecordDefinition() As Collection
    Set RecordDefinition = mCol
End Property

'*-----------------------------------------------------------------------------
'* ColumnCount �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get ColumnCount() As Long
    ColumnCount = mColumnCount
End Property
Property Let ColumnCount(arg As Long)
    mColumnCount = arg
End Property

'*-----------------------------------------------------------------------------
'* DataRowCount �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get DataRowCount() As Long
    DataRowCount = mDataRowCount
End Property
Property Let DataRowCount(arg As Long)
    mDataRowCount = arg
End Property

'*-----------------------------------------------------------------------------
'* ErrMessage �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get ErrMessage() As String
    ErrMessage = mErrMessage
End Property
Property Let ErrMessage(arg As String)
    mErrMessage = arg
End Property

'*-----------------------------------------------------------------------------
'* ErrRowNo �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get ErrRowNo() As Long
    ErrRowNo = mErrRowNo
End Property
Property Let ErrRowNo(arg As Long)
    mErrRowNo = arg
End Property

'*-----------------------------------------------------------------------------
'* mErrColNo �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get ErrColNo() As Long
    ErrColNo = mErrColNo
End Property
Property Let ErrColNo(arg As Long)
    mErrColNo = arg
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mRecordSetCol = New Collection
End Sub
Private Sub Class_Terminate()
    Set mCol = Nothing
    Set mRecordSetCol = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] ���ڕ\���擾�E�ݒ菈���B
'* [��  ��] worksheet�̍��ڕ\��������擾���A�R���N�V�����ɍ��ڂ��ꊇ�ݒ肵�܂��
'*
'* @param dataSheet ���ڕ\���[�N�V�[�g
'* @param lStartRow ���ڕ\�f�[�^�J�n�s�ԍ�
'* @param lStartCol ���ڕ\�f�[�^�J�n��ԍ�
'* @param itemCount ���ڗ�
'*
'******************************************************************************
Public Function GetItemDataFromSheet(dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, Optional colCount As Long) As Boolean
    mDataRowCount = 0
    Dim vArr: vArr = GetVariantDataFromSheet(dataSheet, lStartRow, lStartCol, colCount)
    If IsEmpty(vArr) Then
        GetItemDataFromSheet = False
        Exit Function
    End If
    mDataRowCount = UBound(vArr, 1) - LBound(vArr, 1) + 1
    
    Call SetItems(vArr)
    
    GetItemDataFromSheet = True
End Function


'******************************************************************************
'* [�T  �v] ���ڈꊇ�ݒ菈���B
'* [��  ��] ���ڈꊇ�ݒ肵�܂��
'*
'* @param vArr Variant�^2�����z��i���ڕ\�f�[�^�j
'*
'******************************************************************************
Public Sub SetItems(vArr)
    Dim i As Long

    For i = LBound(vArr, 1) To UBound(vArr, 1)
        Call AddItem(vArr, i)
    Next
    mColumnCount = UBound(vArr, 1) - LBound(vArr, 1) + 1
End Sub

'******************************************************************************
'* [�T  �v] ���ڐݒ菈���B
'* [��  ��] Item�ɒl��ݒ肵�R���N�V�����ɒǉ����܂��
'*
'* @param vArr Variant�^2�����z��i���ڕ\�f�[�^�j
'* @param rownum �z��s�i1�����Y�����j
'*
'******************************************************************************
Public Sub AddItem(vArr, rownum As Long)
    Dim itm As Item
    Set itm = New Item
    itm.Name = vArr(rownum, 1)
    If vArr(rownum, 2) = "��" Then
        itm.Required = True
    End If
    Select Case vArr(rownum, 3)
        Case "���p"
            itm.Attr = AttributeEnum.attrHalf
        Case "���p�p��"
            itm.Attr = AttributeEnum.attrHalfAlphaNumeric
        Case "���l"
            itm.Attr = AttributeEnum.attrNumeric
        Case "���t"
            itm.Attr = AttributeEnum.attrDate
        Case Else
            itm.Attr = AttributeEnum.attrString
    End Select
    Select Case vArr(rownum, 4)
        Case "�Œ�"
            itm.KindOfDigits = KindOfDigitsEnum.digitFixed
        Case "�ȓ�"
            itm.KindOfDigits = KindOfDigitsEnum.digitWithin
        Case "�͈�"
            itm.KindOfDigits = KindOfDigitsEnum.digitRange
        Case Else
            itm.KindOfDigits = KindOfDigitsEnum.digitNone
    End Select
    If vArr(rownum, 5) <> "" And VBA.IsNumeric(vArr(rownum, 5)) Then
        itm.MinNumOfDigits = CLng(vArr(rownum, 5))
    End If
    If vArr(rownum, 6) <> "" And VBA.IsNumeric(vArr(rownum, 6)) Then
        itm.MaxNumOfDigits = CLng(vArr(rownum, 6))
    End If
    itm.Pattern = vArr(rownum, 7)
    If UBound(vArr, 2) = 11 Then
        If vArr(rownum, 8) <> "" And VBA.IsNumeric(vArr(rownum, 8)) Then
            itm.InputCsvColNo = vArr(rownum, 8)
        End If
        itm.InitValueKind = vArr(rownum, 9)
        itm.InitValue = vArr(rownum, 10)
        If vArr(rownum, 11) = "��" Or vArr(rownum, 11) = "" Then
            itm.InputAllowed = True
        End If
    End If
    mCol.Add itm
End Sub

'******************************************************************************
'* [�T  �v] ���؏����B
'* [��  ��] ���R�[�h�f�[�^�����؂��܂��
'*
'* @param vArr Variant�^2�����z��i���R�[�h�f�[�^�j
'* @return ���،���
'*
'******************************************************************************
Public Function Validate(vArr)
    If mColumnCount <> (UBound(vArr, 2) - LBound(vArr, 2) + 1) Then
        Validate = False
        Exit Function
    End If
    Dim statusBarTxt As String: statusBarTxt = "CSV�f�[�^���ؒ��F"
    
    Dim i As Long, j As Long, itm As Item, index As Long
    index = 1
    For i = LBound(vArr, 1) To UBound(vArr, 1)
    
        Application.StatusBar = statusBarTxt
        If i Mod 5000 = 0 Then
            statusBarTxt = statusBarTxt & "��"
            Application.StatusBar = statusBarTxt
            Call CheckEvents
        End If
            
        For j = LBound(vArr, 2) To UBound(vArr, 2)
            Set itm = mCol(index)
            itm.Value = vArr(i, j)
            If Not itm.Validate Then
                Validate = False
                Call SetErrorMessage(itm, i, j)
                Exit Function
            End If
            index = index + 1
        Next
        index = 1
    Next
    Validate = True
End Function

'******************************************************************************
'* [�T  �v] ���b�Z�[�W�ݒ菈���B
'* [��  ��] �G���[�������ڂ���G���[���b�Z�[�W���擾���Đݒ肵�܂��
'*
'* @param itm �G���[��������
'* @param rownum �G���[�����f�[�^�s�ԍ�
'* @param colnum �G���[�����f�[�^��ԍ�
'*
'******************************************************************************
Private Sub SetErrorMessage(ByVal itm As Item, ByVal rownum As Long, ByVal colnum As Long)
    mErrMessage = "[�s" + CStr(rownum) + " ��" + CStr(colnum) + "]  " + itm.Name + vbNewLine + vbNewLine + _
          itm.GetErrorMessage
    mErrRowNo = rownum
    mErrColNo = colnum
End Sub

'******************************************************************************
'* [�T  �v] ���R�[�h���擾�E�ݒ菈���B
'* [��  ��] worksheet�̃f�[�^�\��������擾���A�R���N�V�����ɍ��ڂ�
'*          �ꊇ�ݒ肵�܂��
'*
'* @param dataSheet data�\���[�N�V�[�g
'* @param lStartRow data�\�f�[�^�J�n�s�ԍ�
'* @param lStartCol data�\�f�[�^�J�n��ԍ�
'* @param itemCount ���ڗ�
'*
'******************************************************************************
Public Function GetRecordDataFromSheet(dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, Optional colCount As Long, Optional ignoreColnum As Long) As Boolean
    mDataRowCount = 0
    Dim rng As Range: Set rng = GetActualUsedRangeFromSheet(dataSheet, lStartRow, lStartCol, colCount, ignoreColnum)
    If rng Is Nothing Then
        mErrMessage = "�f�[�^�����͂���Ă��܂���B"
        mErrRowNo = 1
        mErrColNo = 1
        GetRecordDataFromSheet = False
        Exit Function
    End If
    Dim vArr: vArr = rng
    mDataRowCount = UBound(vArr, 1) - LBound(vArr, 1) + 1
    
    GetRecordDataFromSheet = SetDataItems(vArr)
End Function


'******************************************************************************
'* [�T  �v] ���R�[�h���ꊇ�ݒ菈���B
'* [��  ��] ���R�[�h�����ꊇ�ݒ肵�܂��
'*
'* @param vArr Variant�^2�����z��i���R�[�h���f�[�^�j
'*
'******************************************************************************
Public Function SetDataItems(vArr) As Boolean
    Dim i As Long, j As Long, itm As Item, index As Long
    Dim recordCol As Collection
    For i = LBound(vArr, 1) To UBound(vArr, 1)
        Set recordCol = New Collection

        index = 1
        For j = LBound(vArr, 2) To UBound(vArr, 2)
            Set itm = CopyItem(mCol(index))
            itm.Value = vArr(i, j)
            
            If Not itm.Validate Then
                SetDataItems = False
                Call SetErrorMessage(itm, i, j)
                Exit Function
            End If
            recordCol.Add itm
            index = index + 1
        Next
        mRecordSetCol.Add recordCol
    Next
    SetDataItems = True
End Function

'******************************************************************************
'* [�T  �v] Item�R�s�[�����B
'* [��  ��] Item�I�u�W�F�N�g�̃R�s�[��ԋp���܂��
'*
'* @param itm Item�I�u�W�F�N�g
'* @return �R�s�[����Item�I�u�W�F�N�g
'*
'******************************************************************************
Public Function CopyItem(itm As Item) As Item
    Dim newItem As Item: Set newItem = New Item
    newItem.Name = itm.Name
    newItem.Value = itm.Value
    newItem.Required = itm.Required
    newItem.Attr = itm.Attr
    newItem.KindOfDigits = itm.KindOfDigits
    newItem.MinNumOfDigits = itm.MinNumOfDigits
    newItem.MaxNumOfDigits = itm.MaxNumOfDigits
    newItem.Pattern = itm.Pattern
    newItem.InitValueKind = itm.InitValueKind
    newItem.InitValue = itm.InitValue
    newItem.InputCsvColNo = itm.InputCsvColNo
    newItem.InputAllowed = itm.InputAllowed

    Set CopyItem = newItem
End Function

'******************************************************************************
'* [�T  �v] �t�H�[���pVariant�z��擾�����B
'* [��  ��] �Ǎ����R�[�h�f�[�^���t�H�[����`���Ɋ�Â��A�t�H�[���p�f�[�^
'*          �ɕϊ����ԋp���܂��
'*
'* @param vArr Variant�^2�����z��i�Ǎ����R�[�h�f�[�^�j
'* @return Variant�^2�����z��i�t�H�[���p�f�[�^�j
'*
'******************************************************************************
Public Function GetFormVariantData(vArr) As Variant
    Dim vFormArr()
    ReDim vFormArr(0 To UBound(vArr, 1) - LBound(vArr, 1), 0 To mColumnCount - 1)
    
    Dim sInfo As SettingInfo: Set sInfo = GetSettingInfo()
    Dim statusBarTxt As String: statusBarTxt = "�V�[�g�f�[�^�ϊ��������F"

    Dim i As Long, j As Long, itm As Item, index As Long, val As String
    For i = LBound(vFormArr, 1) To UBound(vFormArr, 1)
        
        Application.StatusBar = statusBarTxt
        If i Mod 5000 = 0 Then
            statusBarTxt = statusBarTxt & "��"
            Application.StatusBar = statusBarTxt
            Call CheckEvents
        End If
        
        index = 1
        For j = LBound(vFormArr, 2) To UBound(vFormArr, 2)
            Set itm = mCol(index)
            
            '�l���擾
            If itm.InputCsvColNo = 0 Then
                val = ""
            Else
                val = vArr(i, itm.InputCsvColNo - 1)
            End If
            
            If itm.InitValueKind <> "" Then
                If itm.InitValueKind = "�f�t�H���g" Then
                    val = itm.InitValue
                ElseIf itm.InitValueKind = "�}�X�^�ϊ�" Then
                    val = sInfo.GetMasterValueByCode(itm.InitValue, val)
                End If
            End If
            vFormArr(i, j) = val
            index = index + 1
        Next
    Next
    GetFormVariantData = vFormArr
End Function