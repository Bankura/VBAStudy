VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] ���ڃN���X
'* [��  ��] ���ڂ̃t�H�[�}�b�g��`����ݒ肷��B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* Enum��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* �����̎�ށB
'*
'*-----------------------------------------------------------------------------
Public Enum AttributeEnum
    attrString
    attrHalf
    attrHalfAlphaNumeric
    attrHalfAlphaNumericSymbol
    attrNumeric
    attrZenKatakana
    attrZenHiragana
    attrDate
    attrZipCode
    attrTelNo
    attrMailAddress
End Enum

'*-----------------------------------------------------------------------------
'* �����̎�ށB
'*
'*-----------------------------------------------------------------------------
Public Enum KindOfDigitsEnum
    digitNone
    digitFixed
    digitWithin
    digitRange
End Enum

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mName As String
Private mValue As String
Private mRequired As Boolean
Private mAttribute As AttributeEnum
Private mKindOfDigits As KindOfDigitsEnum
Private mMinNumOfDigits As Long
Private mMaxNumOfDigits As Long
Private mPattern As String
Private mInputCsvColNo As Long
Private mInitValueKind As String
Private mInitValue As String
Private mOutputTarget As Boolean
Private mOutputEditKind As String
Private mOutputEditValue As String

Private mValidOk As Boolean
Private mRequiredOk As Boolean
Private mAttributeOk As Boolean
Private mSizeOk As Boolean
Private mPatternOk As Boolean
Private mMasterExistsOk As Boolean

Private mForOutput As Boolean

Private mRegExp As Object

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************
'*-----------------------------------------------------------------------------
'* Name �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mName
End Property
Property Let Name(arg As String)
    mName = arg
End Property

'*-----------------------------------------------------------------------------
'* Value �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get Value() As String
    Value = mValue
End Property
Property Let Value(arg As String)
    mValue = arg
End Property

'*-----------------------------------------------------------------------------
'* Required �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get Required() As Boolean
    Required = mRequired
End Property
Property Let Required(arg As Boolean)
    mRequired = arg
End Property

'*-----------------------------------------------------------------------------
'* Attr �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get Attr() As AttributeEnum
    Attr = mAttribute
End Property
Property Let Attr(arg As AttributeEnum)
    mAttribute = arg
End Property

'*-----------------------------------------------------------------------------
'* KindOfDigits �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get KindOfDigits() As KindOfDigitsEnum
    KindOfDigits = mKindOfDigits
End Property
Property Let KindOfDigits(arg As KindOfDigitsEnum)
    mKindOfDigits = arg
End Property

'*-----------------------------------------------------------------------------
'* MinNumOfDigits �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get MinNumOfDigits() As Long
    MinNumOfDigits = mMinNumOfDigits
End Property
Property Let MinNumOfDigits(arg As Long)
    mMinNumOfDigits = arg
End Property

'*-----------------------------------------------------------------------------
'* MaxNumOfDigits �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get MaxNumOfDigits() As Long
    MaxNumOfDigits = mMaxNumOfDigits
End Property
Property Let MaxNumOfDigits(arg As Long)
    mMaxNumOfDigits = arg
End Property

'*-----------------------------------------------------------------------------
'* Pattern �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get Pattern() As String
    Pattern = mPattern
End Property
Property Let Pattern(arg As String)
    mPattern = arg
End Property

'*-----------------------------------------------------------------------------
'* InputCsvColNo �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get InputCsvColNo() As Long
    InputCsvColNo = mInputCsvColNo
End Property
Property Let InputCsvColNo(arg As Long)
    mInputCsvColNo = arg
End Property

'*-----------------------------------------------------------------------------
'* InitValueKind �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get InitValueKind() As String
    InitValueKind = mInitValueKind
End Property
Property Let InitValueKind(arg As String)
    mInitValueKind = arg
End Property

'*-----------------------------------------------------------------------------
'* InitValue �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get InitValue() As String
    InitValue = mInitValue
End Property
Property Let InitValue(arg As String)
    mInitValue = arg
End Property

'*-----------------------------------------------------------------------------
'* OutputTarget �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get OutputTarget() As Boolean
    OutputTarget = mOutputTarget
End Property
Property Let OutputTarget(arg As Boolean)
    mOutputTarget = arg
End Property

'*-----------------------------------------------------------------------------
'* OutputEditKind �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get OutputEditKind() As String
    OutputEditKind = mOutputEditKind
End Property
Property Let OutputEditKind(arg As String)
    mOutputEditKind = arg
End Property

'*-----------------------------------------------------------------------------
'* OutputEditValue �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get OutputEditValue() As String
    OutputEditValue = mOutputEditValue
End Property
Property Let OutputEditValue(arg As String)
    mOutputEditValue = arg
End Property

'*-----------------------------------------------------------------------------
'* ForOutput �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get ForOutput() As Boolean
    ForOutput = mForOutput
End Property
Property Let ForOutput(arg As Boolean)
    mForOutput = arg
End Property

'*-----------------------------------------------------------------------------
'* ValidOk �v���p�e�B�i�ǂݎ���p�j
'*-----------------------------------------------------------------------------
Property Get ValidOk() As Boolean
    ValidOk = mValidOk
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    On Error Resume Next
    Set mRegExp = GetRegExp
    On Error GoTo 0
    If mRegExp Is Nothing Then
        Set mRegExp = CreateObject("VBScript.RegExp")
    End If
    mRequired = False
    mValidOk = False
    mRequiredOk = False
    mAttributeOk = False
    mSizeOk = False
    mPatternOk = False
    mOutputTarget = False
End Sub
Private Sub Class_Terminate()
    Set mRegExp = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] �ݒ�l�o���f�[�g�����B
'* [��  ��] Value�̒l�����؂��܂��
'*
'* @return ���،���
'*
'******************************************************************************
Public Function Validate() As Boolean
    mValidOk = False
    mRequiredOk = False
    mSizeOk = False
    mAttributeOk = False
    mPatternOk = False
    mMasterExistsOk = False
    
    mRequiredOk = ValidateRequired()
    mSizeOk = ValidateNumOfDigits()
    mAttributeOk = ValidateAttribute()
    mPatternOk = ValidatePattern()
    mMasterExistsOk = ValidateMasterExists()
    
    mValidOk = mRequiredOk And mSizeOk And mAttributeOk And mPatternOk And mMasterExistsOk
    Validate = mValidOk
End Function

'******************************************************************************
'* [�T  �v] �K�{�`�F�b�N�����B
'* [��  ��] Value���K�{���ڂ��A�ݒ肳��Ă��邩�ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Public Function ValidateRequired() As Boolean
    If mRequired Then
        If Trim(mValue) = "" Then
            ValidateRequired = False
            Exit Function
        End If
    End If
    ValidateRequired = True
End Function

'******************************************************************************
'* [�T  �v] �����`�F�b�N�����B
'* [��  ��] Value�̌������Ó������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Public Function ValidateNumOfDigits() As Boolean
    If Len(mValue) > 0 Then
        Select Case mKindOfDigits
            Case digitFixed
                If Len(mValue) <> mMaxNumOfDigits Then
                    ValidateNumOfDigits = False
                    Exit Function
                End If
            Case digitWithin
                If Len(mValue) > mMaxNumOfDigits Then
                    ValidateNumOfDigits = False
                    Exit Function
                End If
            Case digitRange
                If (Len(mValue) > mMaxNumOfDigits) Or (Len(mValue) < mMinNumOfDigits) Then
                    ValidateNumOfDigits = False
                    Exit Function
                End If
        End Select
    End If
    ValidateNumOfDigits = True
End Function

'******************************************************************************
'* [�T  �v] �����`�F�b�N�����B
'* [��  ��] Value�̑������Ó������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Public Function ValidateAttribute() As Boolean
    Select Case mAttribute
        Case attrHalf
            If Not IsHalf() Then
                ValidateAttribute = False
                Exit Function
            End If
        Case attrHalfAlphaNumeric
            If Not IsHalfAlphaNumeric() Then
                ValidateAttribute = False
                Exit Function
            End If
        Case attrHalfAlphaNumericSymbol
            If Not IsHalfAlphaNumericSymbol() Then
                ValidateAttribute = False
                Exit Function
            End If
        Case attrNumeric
            If Not IsNumeric() Then
                ValidateAttribute = False
                Exit Function
            End If
        Case attrZenKatakana
            If Not IsZenKatakana() Then
                ValidateAttribute = False
                Exit Function
            End If
        Case attrZenHiragana
            If Not IsZenHiragana() Then
                ValidateAttribute = False
                Exit Function
            End If
        Case attrDate
            If Not IsDate() Then
                ValidateAttribute = False
                Exit Function
            End If
        Case attrZipCode
            If Not IsZipCode() Then
                ValidateAttribute = False
                Exit Function
            End If
        Case attrTelNo
            If Not IsTelNo() Then
                ValidateAttribute = False
                Exit Function
            End If
        Case attrMailAddress
            If Not IsMailAddress() Then
                ValidateAttribute = False
                Exit Function
            End If
    End Select
    ValidateAttribute = True
End Function

'******************************************************************************
'* [�T  �v] �p�^�[���`�F�b�N�����B
'* [��  ��] Value���p�^�[���ʂ肩�ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Private Function ValidatePattern() As Boolean
    If Len(mValue) > 0 And mPattern <> "" Then
        With mRegExp
            .Pattern = mPattern
            .IgnoreCase = False
            .Global = True
            ValidatePattern = .Test(mValue)
        End With
        Exit Function
    End If
    ValidatePattern = True
End Function

'******************************************************************************
'* [�T  �v] �}�X�^���݃`�F�b�N�����B
'* [��  ��] Value���}�X�^�ɑ��݂��邩�ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Private Function ValidateMasterExists() As Boolean
    If mValue = "" Then
        ValidateMasterExists = True
        Exit Function
    End If
    
    Dim val As String
    Dim sInfo As SettingInfo: Set sInfo = GetSettingInfo()

    If mForOutput Then
        If mOutputEditValue <> "" Then
            If mOutputEditKind = "�}�X�^�ϊ��iCode��Value�j" Then
                val = sInfo.GetMasterValueByCode(mOutputEditValue, mValue)
            ElseIf mOutputEditKind = "�}�X�^�ϊ��iValue��Code�j" Then
                val = sInfo.GetMasterCodeByValue(mOutputEditValue, mValue)
            End If
            If val = "" Then
                ValidateMasterExists = False
                Exit Function
            End If
        End If
    Else
        If mInitValue <> "" Then
            If mInitValueKind = "�}�X�^�ϊ��iCode��Value�j" Then
                val = sInfo.GetMasterValueByCode(mInitValue, mValue)
            ElseIf mInitValueKind = "�}�X�^�ϊ��iValue��Code�j" Then
                val = sInfo.GetMasterCodeByValue(mInitValue, mValue)
            End If
            If val = "" Then
                ValidateMasterExists = False
                Exit Function
            End If
        End If
    End If
    ValidateMasterExists = True
End Function
            
'******************************************************************************
'* [�T  �v] ���p�`�F�b�N�����B
'* [��  ��] Value�����p�������ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Private Function IsHalf() As Boolean
    Dim checkValue As String
    If Len(mValue) > 0 Then
        checkValue = StrConv(mValue, vbFromUnicode)
        If Len(mValue) <> LenB(checkValue) Then
            IsHalf = False
            Exit Function
        End If
    End If
    IsHalf = True
End Function

'******************************************************************************
'* [�T  �v] ���p�p�����`�F�b�N�����B
'* [��  ��] Value�����p�p�������ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Private Function IsHalfAlphaNumeric() As Boolean
    Dim i As Long
    If Len(mValue) > 0 Then
        For i = 1 To Len(mValue)
            If Not Mid(mValue, i, 1) Like "[0-9a-zA-Z]" Then
                IsHalfAlphaNumeric = False
                Exit Function
            End If
        Next
    End If
    IsHalfAlphaNumeric = True
End Function

'******************************************************************************
'* [�T  �v] ���p�p�����L���`�F�b�N�����B
'* [��  ��] Value�����p�p�����L�����ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Private Function IsHalfAlphaNumericSymbol() As Boolean
    If Len(mValue) > 0 Then
        With mRegExp
            .Pattern = "^[a-zA-Z0-9!-/:-@\[-`{-~]*$"
            .IgnoreCase = False
            .Global = True
            IsHalfAlphaNumericSymbol = .Test(mValue)
        End With
        Exit Function
    End If
    IsHalfAlphaNumericSymbol = True
End Function

'******************************************************************************
'* [�T  �v] ���l�`�F�b�N�����B
'* [��  ��] Value�����l���ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Private Function IsNumeric() As Boolean
    Dim i As Long
    If Len(mValue) > 0 Then
        For i = 1 To Len(mValue)
            If Not Mid(mValue, i, 1) Like "[0-9]" Then
                IsNumeric = False
                Exit Function
            End If
        Next
    End If
    IsNumeric = True
End Function

'******************************************************************************
'* [�T  �v] �S�p�J�^�J�i�`�F�b�N�����B
'* [��  ��] Value���S�p�J�^�J�i���ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Private Function IsZenKatakana() As Boolean
    If Len(mValue) > 0 Then
        With mRegExp
            .Pattern = "^[�@-�����[]*$"
            .IgnoreCase = False
            .Global = True
            IsZenKatakana = .Test(mValue)
        End With
        Exit Function
    End If
    IsZenKatakana = True
End Function

'******************************************************************************
'* [�T  �v] �S�p�Ђ炪�ȃ`�F�b�N�����B
'* [��  ��] Value���S�p�Ђ炪�Ȃ��ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Private Function IsZenHiragana() As Boolean
    If Len(mValue) > 0 Then
        With mRegExp
            .Pattern = "^[��-��[]*$"
            .IgnoreCase = False
            .Global = True
            IsZenHiragana = .Test(mValue)
        End With
        Exit Function
    End If
    IsZenHiragana = True
End Function

'******************************************************************************
'* [�T  �v] ���t�`�F�b�N�����B
'* [��  ��] Value�����t�iyyyy/mm/dd�j���ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Private Function IsDate() As Boolean
    Dim i As Long, tmp As String, tmpDate As String
    If Len(mValue) > 0 Then
        For i = 1 To Len(mValue)
            tmp = Mid(mValue, i, 1)
            If i <> 5 And i <> 8 And (tmp Like "[0-9]") Then
                tmpDate = tmpDate & tmp
            ElseIf (i = 5 Or i = 8) And (tmp <> "/") Then
                IsDate = False
                Exit Function
            End If
        Next
        If Len(tmpDate) <> 8 Then
            IsDate = False
            Exit Function
        End If
        tmpDate = Format(tmpDate, "####/##/##")
        IsDate = VBA.IsDate(tmpDate)
        Exit Function
    End If
    IsDate = True
End Function

'******************************************************************************
'* [�T  �v] �X�֔ԍ��`�F�b�N�����B
'* [��  ��] Value���X�֔ԍ����ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Private Function IsZipCode() As Boolean
    If Len(mValue) > 0 Then
        With mRegExp
            .Pattern = "^\d{7}$|^\d{3}-\d{4}$"
            .IgnoreCase = False
            .Global = True
            IsZipCode = .Test(mValue)
        End With
        Exit Function
    End If
    IsZipCode = True
End Function

'******************************************************************************
'* [�T  �v] �d�b�ԍ��`�F�b�N�����B
'* [��  ��] Value���d�b�ԍ����ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Private Function IsTelNo() As Boolean
    If Len(mValue) > 0 Then
        With mRegExp
            .Pattern = "^0\d{9,10}$|^0\d{2,3}-\d{1,4}-\d{4}$"
            .IgnoreCase = False
            .Global = True
            IsTelNo = .Test(mValue)
        End With
        Exit Function
    End If
    IsTelNo = True
End Function

'******************************************************************************
'* [�T  �v] ���[���A�h���X�`�F�b�N�����B
'* [��  ��] Value�����[���A�h���X���ǂ������肵�܂��
'*
'* @return ���茋��
'*
'******************************************************************************
Private Function IsMailAddress() As Boolean
    If Len(mValue) > 0 Then
        With mRegExp
            .Pattern = "^[a-zA-Z0-9.!#$%&'*+\/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$"
            .IgnoreCase = False
            .Global = True
            IsMailAddress = .Test(mValue)
        End With
        Exit Function
    End If
    IsMailAddress = True
End Function

'******************************************************************************
'* [�T  �v] �G���[���b�Z�[�W�擾�����B
'* [��  ��] �G���[�󋵂ɉ������G���[���b�Z�[�W�𐶐����ԋp���܂��
'*
'* @return �G���[���b�Z�[�W
'*
'******************************************************************************
Public Function GetErrorMessage() As String
    Dim msg As String
    If mValidOk Then
        GetErrorMessage = ""
    Else
        If Not mRequiredOk Then
            msg = EditMessage(msg, "�E{0}�����͂���Ă��܂���B", Array(mName))
        End If
        If Not mAttributeOk Then
            Dim attrName As String
            Select Case mAttribute
                Case attrHalf
                    attrName = "���p"
                Case attrHalfAlphaNumeric
                    attrName = "���p�p��"
                Case attrHalfAlphaNumericSymbol
                    attrName = "���p�p���L��"
                Case attrNumeric
                    attrName = "���l"
                Case attrZenKatakana
                    attrName = "�S�p�J�^�J�i"
                Case attrZenHiragana
                    attrName = "�S�p�Ђ炪��"
                Case attrDate
                    attrName = "���t"
                Case attrZipCode
                    attrName = "�X�֔ԍ�"
                Case attrTelNo
                    attrName = "�d�b�ԍ�"
                Case attrMailAddress
                    attrName = "���[���A�h���X"
            End Select
            msg = EditMessage(msg, "�E{0}��{1}�œ��͂��Ă��������B", Array(mName, attrName))
        End If
        If Not mSizeOk Then
            Dim digitTxt As String
            Select Case mKindOfDigits
                Case digitFixed
                    msg = EditMessage(msg, "�E{0}��{1}���œ��͂��Ă��������B", Array(mName, CStr(mMaxNumOfDigits)))
                Case digitWithin
                    msg = EditMessage(msg, "�E{0}��{1}���ȓ��œ��͂��Ă��������B", Array(mName, CStr(mMaxNumOfDigits)))
                Case digitRange
                    msg = EditMessage(msg, "�E{0}��{1}���ȏ�{2}���ȓ��œ��͂��Ă��������B", Array(mName, CStr(mMinNumOfDigits), CStr(mMaxNumOfDigits)))
            End Select
        End If
        If Not mPatternOk Then
            msg = EditMessage(msg, "�E{0}�̌`�����قȂ�܂��B", Array(mName))
        End If
        If Not mMasterExistsOk Then
            msg = EditMessage(msg, "�E{0}���I���\�Ȓl�ł͂���܂���B", Array(mName))
        End If
        GetErrorMessage = msg
    End If

End Function

'******************************************************************************
'* [�T  �v] ���b�Z�[�W�{���ҏW�����B
'* [��  ��] ���b�Z�[�W�{���ɉ��s�������ă��b�Z�[�W��A���t�^���܂��
'*
'* @param body ���b�Z�[�W�{��
'* @param msg ���b�Z�[�W�i{0}{1}�Ƃ����������w�肠��j
'* @param padVals ��������
'* @return �G���[���b�Z�[�W
'*
'******************************************************************************
Private Function EditMessage(ByVal body As String, ByVal msg As String, ByVal padVals) As String
    If body = "" Then
        EditMessage = FormatMessage(msg, padVals)
    Else
        EditMessage = body + vbNewLine + FormatMessage(msg, padVals)
    End If
End Function

'*******************************************************************************
'* [�T  �v] ���b�Z�[�W�ҏW�����B
'* [��  ��] ���b�Z�[�W�ɒl�𖄍��ݕԋp���܂��
'*
'* @param msg ���b�Z�[�W�i{0}{1}�Ƃ����������w�肠��j
'* @param padVals ��������
'* @return �ҏW���b�Z�[�W
'*
'*******************************************************************************
Private Function FormatMessage(msg As String, padVals) As String
    Dim i As Integer, tmpTxt As String
    tmpTxt = msg
    For i = 0 To UBound(padVals)
        tmpTxt = Replace(tmpTxt, "{" & i & "}", padVals(i))
    Next i
    FormatMessage = tmpTxt
End Function