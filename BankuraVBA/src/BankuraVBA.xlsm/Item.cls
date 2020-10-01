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
'* [機能名] 項目クラス
'* [詳  細] 項目のフォーマット定義情報を設定する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* Enum定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* 属性の種類。
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
'* 桁数の種類。
'*
'*-----------------------------------------------------------------------------
Public Enum KindOfDigitsEnum
    digitNone
    digitFixed
    digitWithin
    digitRange
End Enum

'******************************************************************************
'* 内部変数定義
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
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* Name プロパティ
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mName
End Property
Property Let Name(arg As String)
    mName = arg
End Property

'*-----------------------------------------------------------------------------
'* Value プロパティ
'*-----------------------------------------------------------------------------
Property Get Value() As String
    Value = mValue
End Property
Property Let Value(arg As String)
    mValue = arg
End Property

'*-----------------------------------------------------------------------------
'* Required プロパティ
'*-----------------------------------------------------------------------------
Property Get required() As Boolean
    required = mRequired
End Property
Property Let required(arg As Boolean)
    mRequired = arg
End Property

'*-----------------------------------------------------------------------------
'* Attr プロパティ
'*-----------------------------------------------------------------------------
Property Get Attr() As AttributeEnum
    Attr = mAttribute
End Property
Property Let Attr(arg As AttributeEnum)
    mAttribute = arg
End Property

'*-----------------------------------------------------------------------------
'* KindOfDigits プロパティ
'*-----------------------------------------------------------------------------
Property Get KindOfDigits() As KindOfDigitsEnum
    KindOfDigits = mKindOfDigits
End Property
Property Let KindOfDigits(arg As KindOfDigitsEnum)
    mKindOfDigits = arg
End Property

'*-----------------------------------------------------------------------------
'* MinNumOfDigits プロパティ
'*-----------------------------------------------------------------------------
Property Get MinNumOfDigits() As Long
    MinNumOfDigits = mMinNumOfDigits
End Property
Property Let MinNumOfDigits(arg As Long)
    mMinNumOfDigits = arg
End Property

'*-----------------------------------------------------------------------------
'* MaxNumOfDigits プロパティ
'*-----------------------------------------------------------------------------
Property Get MaxNumOfDigits() As Long
    MaxNumOfDigits = mMaxNumOfDigits
End Property
Property Let MaxNumOfDigits(arg As Long)
    mMaxNumOfDigits = arg
End Property

'*-----------------------------------------------------------------------------
'* Pattern プロパティ
'*-----------------------------------------------------------------------------
Property Get Pattern() As String
    Pattern = mPattern
End Property
Property Let Pattern(arg As String)
    mPattern = arg
End Property

'*-----------------------------------------------------------------------------
'* InputCsvColNo プロパティ
'*-----------------------------------------------------------------------------
Property Get InputCsvColNo() As Long
    InputCsvColNo = mInputCsvColNo
End Property
Property Let InputCsvColNo(arg As Long)
    mInputCsvColNo = arg
End Property

'*-----------------------------------------------------------------------------
'* InitValueKind プロパティ
'*-----------------------------------------------------------------------------
Property Get InitValueKind() As String
    InitValueKind = mInitValueKind
End Property
Property Let InitValueKind(arg As String)
    mInitValueKind = arg
End Property

'*-----------------------------------------------------------------------------
'* InitValue プロパティ
'*-----------------------------------------------------------------------------
Property Get InitValue() As String
    InitValue = mInitValue
End Property
Property Let InitValue(arg As String)
    mInitValue = arg
End Property

'*-----------------------------------------------------------------------------
'* OutputTarget プロパティ
'*-----------------------------------------------------------------------------
Property Get OutputTarget() As Boolean
    OutputTarget = mOutputTarget
End Property
Property Let OutputTarget(arg As Boolean)
    mOutputTarget = arg
End Property

'*-----------------------------------------------------------------------------
'* OutputEditKind プロパティ
'*-----------------------------------------------------------------------------
Property Get OutputEditKind() As String
    OutputEditKind = mOutputEditKind
End Property
Property Let OutputEditKind(arg As String)
    mOutputEditKind = arg
End Property

'*-----------------------------------------------------------------------------
'* OutputEditValue プロパティ
'*-----------------------------------------------------------------------------
Property Get OutputEditValue() As String
    OutputEditValue = mOutputEditValue
End Property
Property Let OutputEditValue(arg As String)
    mOutputEditValue = arg
End Property

'*-----------------------------------------------------------------------------
'* ForOutput プロパティ
'*-----------------------------------------------------------------------------
Property Get ForOutput() As Boolean
    ForOutput = mForOutput
End Property
Property Let ForOutput(arg As Boolean)
    mForOutput = arg
End Property

'*-----------------------------------------------------------------------------
'* ValidOk プロパティ（読み取り専用）
'*-----------------------------------------------------------------------------
Property Get ValidOk() As Boolean
    ValidOk = mValidOk
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mRegExp = Base.GetRegExp
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
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] 設定値バリデート処理。
'* [詳  細] Valueの値を検証します｡
'*
'* @return 検証結果
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
'* [概  要] 必須チェック処理。
'* [詳  細] Valueが必須項目か、設定されているかどうか判定します｡
'*
'* @return 判定結果
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
'* [概  要] 桁数チェック処理。
'* [詳  細] Valueの桁数が妥当か判定します｡
'*
'* @return 判定結果
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
'* [概  要] 属性チェック処理。
'* [詳  細] Valueの属性が妥当か判定します｡
'*
'* @return 判定結果
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
'* [概  要] パターンチェック処理。
'* [詳  細] Valueがパターン通りかどうか判定します｡
'*
'* @return 判定結果
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
'* [概  要] マスタ存在チェック処理。
'* [詳  細] Valueがマスタに存在するかどうか判定します｡
'*
'* @return 判定結果
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
            If mOutputEditKind = "マスタ変換（Code→Value）" Then
                val = sInfo.GetMasterValueByCode(mOutputEditValue, mValue)
            ElseIf mOutputEditKind = "マスタ変換（Value→Code）" Then
                val = sInfo.GetMasterCodeByValue(mOutputEditValue, mValue)
            End If
            If val = "" Then
                ValidateMasterExists = False
                Exit Function
            End If
        End If
    Else
        If mInitValue <> "" Then
            If mInitValueKind = "マスタ変換（Code→Value）" Then
                val = sInfo.GetMasterValueByCode(mInitValue, mValue)
            ElseIf mInitValueKind = "マスタ変換（Value→Code）" Then
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
'* [概  要] 半角チェック処理。
'* [詳  細] Valueが半角文字かどうか判定します｡
'*
'* @return 判定結果
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
'* [概  要] 半角英数字チェック処理。
'* [詳  細] Valueが半角英数字かどうか判定します｡
'*
'* @return 判定結果
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
'* [概  要] 半角英数字記号チェック処理。
'* [詳  細] Valueが半角英数字記号かどうか判定します｡
'*
'* @return 判定結果
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
'* [概  要] 数値チェック処理。
'* [詳  細] Valueが数値かどうか判定します｡
'*
'* @return 判定結果
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
'* [概  要] 全角カタカナチェック処理。
'* [詳  細] Valueが全角カタカナかどうか判定します｡
'*
'* @return 判定結果
'*
'******************************************************************************
Private Function IsZenKatakana() As Boolean
    If Len(mValue) > 0 Then
        With mRegExp
            .Pattern = "^[ァ-ンヴー]*$"
            .IgnoreCase = False
            .Global = True
            IsZenKatakana = .Test(mValue)
        End With
        Exit Function
    End If
    IsZenKatakana = True
End Function

'******************************************************************************
'* [概  要] 全角ひらがなチェック処理。
'* [詳  細] Valueが全角ひらがなかどうか判定します｡
'*
'* @return 判定結果
'*
'******************************************************************************
Private Function IsZenHiragana() As Boolean
    If Len(mValue) > 0 Then
        With mRegExp
            .Pattern = "^[ぁ-んー]*$"
            .IgnoreCase = False
            .Global = True
            IsZenHiragana = .Test(mValue)
        End With
        Exit Function
    End If
    IsZenHiragana = True
End Function

'******************************************************************************
'* [概  要] 日付チェック処理。
'* [詳  細] Valueが日付（yyyy/mm/dd）かどうか判定します｡
'*
'* @return 判定結果
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
'* [概  要] 郵便番号チェック処理。
'* [詳  細] Valueが郵便番号かどうか判定します｡
'*
'* @return 判定結果
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
'* [概  要] 電話番号チェック処理。
'* [詳  細] Valueが電話番号かどうか判定します｡
'*
'* @return 判定結果
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
'* [概  要] メールアドレスチェック処理。
'* [詳  細] Valueがメールアドレスかどうか判定します｡
'*
'* @return 判定結果
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
'* [概  要] エラーメッセージ取得処理。
'* [詳  細] エラー状況に応じたエラーメッセージを生成し返却します｡
'*
'* @return エラーメッセージ
'*
'******************************************************************************
Public Function GetErrorMessage() As String
    Dim msg As String
    If mValidOk Then
        GetErrorMessage = ""
    Else
        If Not mRequiredOk Then
            msg = EditMessage(msg, "・{0}が入力されていません。", Array(mName))
        End If
        If Not mAttributeOk Then
            Dim attrName As String
            Select Case mAttribute
                Case attrHalf
                    attrName = "半角"
                Case attrHalfAlphaNumeric
                    attrName = "半角英数"
                Case attrHalfAlphaNumericSymbol
                    attrName = "半角英数記号"
                Case attrNumeric
                    attrName = "数値"
                Case attrZenKatakana
                    attrName = "全角カタカナ"
                Case attrZenHiragana
                    attrName = "全角ひらがな"
                Case attrDate
                    attrName = "日付"
                Case attrZipCode
                    attrName = "郵便番号"
                Case attrTelNo
                    attrName = "電話番号"
                Case attrMailAddress
                    attrName = "メールアドレス"
            End Select
            msg = EditMessage(msg, "・{0}は{1}で入力してください。", Array(mName, attrName))
        End If
        If Not mSizeOk Then
            Dim digitTxt As String
            Select Case mKindOfDigits
                Case digitFixed
                    msg = EditMessage(msg, "・{0}は{1}桁で入力してください。", Array(mName, CStr(mMaxNumOfDigits)))
                Case digitWithin
                    msg = EditMessage(msg, "・{0}は{1}桁以内で入力してください。", Array(mName, CStr(mMaxNumOfDigits)))
                Case digitRange
                    msg = EditMessage(msg, "・{0}は{1}桁以上{2}桁以内で入力してください。", Array(mName, CStr(mMinNumOfDigits), CStr(mMaxNumOfDigits)))
            End Select
        End If
        If Not mPatternOk Then
            msg = EditMessage(msg, "・{0}の形式が異なります。", Array(mName))
        End If
        If Not mMasterExistsOk Then
            msg = EditMessage(msg, "・{0}が選択可能な値ではありません。", Array(mName))
        End If
        GetErrorMessage = msg
    End If

End Function

'******************************************************************************
'* [概  要] メッセージ本文編集処理。
'* [詳  細] メッセージ本文に改行を加えてメッセージを連結付与します｡
'*
'* @param body メッセージ本体
'* @param msg メッセージ（{0}{1}といった埋込指定あり）
'* @param padVals 埋込文字
'* @return エラーメッセージ
'*
'******************************************************************************
Private Function EditMessage(ByVal Body As String, ByVal msg As String, ByVal padVals) As String
    If Body = "" Then
        EditMessage = FormatMessage(msg, padVals)
    Else
        EditMessage = Body + vbNewLine + FormatMessage(msg, padVals)
    End If
End Function

'*******************************************************************************
'* [概  要] メッセージ編集処理。
'* [詳  細] メッセージに値を埋込み返却します｡
'*
'* @param msg メッセージ（{0}{1}といった埋込指定あり）
'* @param padVals 埋込文字
'* @return 編集メッセージ
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
