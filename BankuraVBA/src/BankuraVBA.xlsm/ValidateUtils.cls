VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidateUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] 検証ユーティリティクラス
'* [詳  細] 検証を行うユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2019-2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* メソッド定義
'******************************************************************************
'******************************************************************************
'* [概  要] IsNotMissing
'* [詳  細] Optional引数が省略されていないことをチェック｡
'*
'* @param arg 引数
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsNotMissing(ByVal arg) As Boolean
    IsNotMissing = Not IsMissing(arg)
End Function

Public Function IsString(var) As Boolean
    IsString = (TypeName(var) = "String")
End Function

Public Function IsStringArray(var) As Boolean
    IsStringArray = (TypeName(var) = "String()")
End Function

Public Function IsNotString(var) As Boolean
    IsNotString = Not IsString(var)
End Function

Public Function IsNotStringArray(var) As Boolean
    IsNotStringArray = Not IsStringArray(var)
End Function

Public Function IsNumericType(var) As Boolean
    IsNumericType = IsInteger(var) Or IsLong(var) Or IsSingle(var) Or IsDouble(var) Or IsCurrency(var) Or IsDecimal(var)
End Function

Public Function IsNotNumericType(var) As Boolean
    IsNotNumericType = Not IsNumericType(var)
End Function

Public Function IsInteger(var) As Boolean
    IsInteger = (TypeName(var) = "Integer")
End Function

Public Function IsNotInteger(var) As Boolean
    IsNotInteger = Not IsInteger(var)
End Function

Public Function IsIntegerArray(var) As Boolean
    IsIntegerArray = (TypeName(var) = "Integer()")
End Function

Public Function IsNotIntegerArray(var) As Boolean
    IsNotIntegerArray = Not IsIntegerArray(var)
End Function

Public Function IsLong(var) As Boolean
    IsLong = (TypeName(var) = "Long")
End Function

Public Function IsNotLong(var) As Boolean
    IsNotLong = Not IsLong(var)
End Function

Public Function IsLongArray(var) As Boolean
    IsLongArray = (TypeName(var) = "Long()")
End Function

Public Function IsNotLongArray(var) As Boolean
    IsNotLongArray = Not IsLongArray(var)
End Function

Public Function IsSingle(var) As Boolean
    IsSingle = (TypeName(var) = "Single")
End Function

Public Function IsSingleArray(var) As Boolean
    IsSingleArray = (TypeName(var) = "Single()")
End Function

Public Function IsNotSingle(var) As Boolean
    IsNotSingle = Not IsSingle(var)
End Function

Public Function IsNotSingleArray(var) As Boolean
    IsNotSingleArray = Not IsSingleArray(var)
End Function

Public Function IsDouble(var) As Boolean
    IsDouble = (TypeName(var) = "Double")
End Function

Public Function IsDoubleArray(var) As Boolean
    IsDoubleArray = (TypeName(var) = "Double()")
End Function

Public Function IsNotDouble(var) As Boolean
    IsNotDouble = Not IsDouble(var)
End Function

Public Function IsNotDoubleArray(var) As Boolean
    IsNotDoubleArray = Not IsDoubleArray(var)
End Function

Public Function IsCurrency(var) As Boolean
    IsCurrency = (TypeName(var) = "Currency")
End Function

Public Function IsCurrencyArray(var) As Boolean
    IsCurrencyArray = (TypeName(var) = "Currency()")
End Function

Public Function IsNotCurrency(var) As Boolean
    IsNotCurrency = Not IsCurrency(var)
End Function

Public Function IsNotCurrencyArray(var) As Boolean
    IsNotCurrencyArray = Not IsCurrencyArray(var)
End Function

Public Function IsDecimal(var) As Boolean
    IsDecimal = (TypeName(var) = "Decimal")
End Function

Public Function IsDecimalArray(var) As Boolean
    IsDecimalArray = (TypeName(var) = "Decimal()")
End Function

Public Function IsNotDecimal(var) As Boolean
    IsNotDecimal = Not IsDecimal(var)
End Function

Public Function IsNotDecimalArray(var) As Boolean
    IsNotDecimalArray = Not IsDecimalArray(var)
End Function

Public Function IsDateType(var) As Boolean
    IsDateType = (TypeName(var) = "Date")
End Function

Public Function IsDateTypeArray(var) As Boolean
    IsDateTypeArray = (TypeName(var) = "Date()")
End Function

Public Function IsNotDateType(var) As Boolean
    IsNotDateType = Not IsDateType(var)
End Function

Public Function IsNotDateTypeArray(var) As Boolean
    IsNotDateTypeArray = Not IsDateTypeArray(var)
End Function

Public Function IsBoolean(var) As Boolean
    IsBoolean = (TypeName(var) = "Boolean")
End Function

Public Function IsBooleanArray(var) As Boolean
    IsBooleanArray = (TypeName(var) = "Boolean()")
End Function

Public Function IsNotBoolean(var) As Boolean
    IsNotBoolean = Not IsBoolean(var)
End Function

Public Function IsNotBooleanArray(var) As Boolean
    IsNotBooleanArray = Not IsBooleanArray(var)
End Function

Public Function IsVariant(var) As Boolean
    IsVariant = (TypeName(var) = "Variant")
End Function

Public Function IsVariantArray(var) As Boolean
    IsVariantArray = (TypeName(var) = "Variant()")
End Function

Public Function IsNotVariant(var) As Boolean
    IsNotVariant = Not IsVariant(var)
End Function

Public Function IsNotVariantArray(var) As Boolean
    IsNotVariantArray = Not IsVariantArray(var)
End Function

Public Function IsError(var) As Boolean
    IsError = (TypeName(var) = "Error")
End Function

Public Function IsNotError(var) As Boolean
    IsNotError = Not IsError(var)
End Function

Public Function IsNotEmpty(var) As Boolean
    IsNotEmpty = Not VBA.IsEmpty(var)
End Function

Public Function IsFalse(flg As Boolean) As Boolean
    IsFalse = Not flg
End Function

Public Function IsArray(val) As Boolean
    IsArray = VBA.IsArray(val)
End Function

Public Function IsNotArray(val) As Boolean
    IsNotArray = Not Me.IsArray(val)
End Function

Public Function IsNullOrEmpty(var) As Boolean
    IsNullOrEmpty = (VBA.IsNull(var) Or VBA.IsEmpty(var))
End Function

Public Function IsNotNullOrEmpty(var) As Boolean
    IsNotNullOrEmpty = Not IsNullOrEmpty(var)
End Function

Public Function IsEmptyOrArray(var) As Boolean
    IsEmptyOrArray = (VBA.IsEmpty(var) Or VBA.IsArray(var))
End Function

Public Function IsNotEmptyOrArray(var) As Boolean
    IsNotEmptyOrArray = Not IsEmptyOrArray(var)
End Function

Public Function IsNullOrEmptyOrArray(var) As Boolean
    IsNullOrEmptyOrArray = (VBA.IsNull(var) Or VBA.IsEmpty(var) Or VBA.IsArray(var))
End Function

Public Function IsNotNullOrEmptyOrArray(var) As Boolean
    IsNotNullOrEmptyOrArray = Not IsNullOrEmptyOrArray(var)
End Function

Public Function IsPrimitive(var)
    IsPrimitive = IsNotNullOrEmptyOrArray(var) And IsNotObject(var)
End Function

Public Function IsNotPrimitive(var)
    IsNotPrimitive = Not IsPrimitive(var)
End Function

Public Function IsNotNull(var) As Boolean
    IsNotNull = Not VBA.IsNull(var)
End Function

Public Function IsNotObject(obj) As Boolean
    IsNotObject = Not VBA.IsObject(obj)
End Function

Public Function IsUnknown(obj) As Boolean
    IsUnknown = (TypeName(obj) = "Unknown")
End Function

Public Function IsNotUnknown(obj) As Boolean
    IsNotUnknown = Not IsUnknown(obj)
End Function

Public Function IsNothing(obj) As Boolean
    IsNothing = (obj Is Nothing)
End Function

Public Function IsNotNothing(obj) As Boolean
    IsNotNothing = Not IsNothing(obj)
End Function

Public Function IsCollection(obj) As Boolean
    IsCollection = (TypeName(obj) = "Collection")
End Function

Public Function IsNotCollection(obj) As Boolean
    IsNotCollection = Not IsCollection(obj)
End Function

Public Function IsDictionary(obj) As Boolean
    IsDictionary = (TypeName(obj) = "Dictionary")
End Function

Public Function IsNotDictionary(obj) As Boolean
    IsNotDictionary = Not IsDictionary(obj)
End Function

Public Function IsDictionaryEx(obj) As Boolean
    IsDictionaryEx = (TypeName(obj) = "DictionaryEx")
End Function

Public Function IsNotDictionaryEx(obj) As Boolean
    IsNotDictionaryEx = Not IsDictionaryEx(obj)
End Function

Public Function IsWScriptShell(obj) As Boolean
    IsWScriptShell = (TypeName(obj) = "IWshShell3") Or _
                     (TypeName(obj) = "IWshShell2") Or _
                     (TypeName(obj) = "IWshShell")
End Function

Public Function IsNotWScriptShell(obj) As Boolean
    IsNotWScriptShell = Not IsWScriptShell(obj)
End Function

Public Function IsWmi(obj) As Boolean
    IsWmi = (TypeName(obj) = "SWbemLocator")
End Function

Public Function IsNotWmi(obj) As Boolean
    IsNotWmi = Not IsWmi(obj)
End Function

Public Function IsScriptControl(obj) As Boolean
    IsScriptControl = (TypeName(obj) = "ScriptControl")
End Function

Public Function IsNotScriptControl(obj) As Boolean
    IsNotScriptControl = Not IsScriptControl(obj)
End Function

Public Function IsShell(obj) As Boolean
    IsShell = (TypeName(obj) = "IShellDispatch6") Or _
              (TypeName(obj) = "IShellDispatch5") Or _
              (TypeName(obj) = "IShellDispatch4") Or _
              (TypeName(obj) = "IShellDispatch3") Or _
              (TypeName(obj) = "IShellDispatch2") Or _
              (TypeName(obj) = "IShellDispatch") Or _
              (TypeName(obj) = "Shell")
End Function

Public Function IsNotShell(obj) As Boolean
    IsNotShell = Not IsShell(obj)
End Function

Public Function IsFolderItem(obj) As Boolean
    IsFolderItem = (TypeName(obj) = "FolderItem2") Or _
              (TypeName(obj) = "FolderItem")
End Function

Public Function IsNotFolderItem(obj) As Boolean
    IsNotFolderItem = Not IsFolderItem(obj)
End Function

Public Function IsRegExp(obj) As Boolean
    IsRegExp = (TypeName(obj) = "IRegExp2")
End Function

Public Function IsNotRegExp(obj) As Boolean
    IsNotRegExp = Not IsRegExp(obj)
End Function

Public Function IsStringEx(obj) As Boolean
    IsStringEx = (TypeName(obj) = "StringEx")
End Function

Public Function IsNotStringEx(obj) As Boolean
    IsNotStringEx = Not IsStringEx(obj)
End Function

Public Function IsArrayEx(obj) As Boolean
    IsArrayEx = (TypeName(obj) = "ArrayEx")
End Function

Public Function IsNotArrayEx(obj) As Boolean
    IsNotArrayEx = Not IsArrayEx(obj)
End Function

Public Function IsTuple(obj) As Boolean
    IsTuple = (TypeName(obj) = "Tuple")
End Function

Public Function IsNotTuple(obj) As Boolean
    IsNotTuple = Not IsTuple(obj)
End Function

Public Function IsFunc(obj) As Boolean
    IsFunc = (TypeName(obj) = "Func")
End Function

Public Function IsNotFunc(obj) As Boolean
    IsNotFunc = Not IsFunc(obj)
End Function

Public Function IsArray2DEx(obj) As Boolean
    IsArray2DEx = (TypeName(obj) = "Array2DEx")
End Function

Public Function IsNotArray2DEx(obj) As Boolean
    IsNotArray2DEx = Not IsArray2DEx(obj)
End Function

Public Function IsArrayExOfArrayEx(obj) As Boolean
    IsArrayExOfArrayEx = False
    If IsArrayEx(obj) Then
        If obj.IsEmptyArray Then
            IsArrayExOfArrayEx = True
        Else
            IsArrayExOfArrayEx = IsArrayEx(obj.Item(0))
        End If
    End If
End Function

Public Function IsNotArrayExOfArrayEx(obj) As Boolean
    IsNotArrayExOfArrayEx = Not IsArrayExOfArrayEx(obj)
End Function

Public Function IsArrayOfArrayEx(var) As Boolean
    IsArrayOfArrayEx = False
    If IsArray(var) Then
        If ArrayUtils.IsEmptyArray(var) Then
            IsArrayOfArrayEx = True
        Else
            IsArrayOfArrayEx = IsArrayEx(var(LBound(var)))
        End If
    End If
End Function

Public Function IsNotArrayOfArrayEx(obj) As Boolean
    IsNotArrayOfArrayEx = Not IsArrayOfArrayEx(obj)
End Function

Public Function IsRecordFormat(obj) As Boolean
    IsRecordFormat = (TypeName(obj) = "RecordFormat")
End Function

Public Function IsNotRecordFormat(obj) As Boolean
    IsNotRecordFormat = Not IsRecordFormat(obj)
End Function

Public Function IsParamArrayMissing(ByVal arr) As Boolean
    IsParamArrayMissing = LBound(arr) > UBound(arr)
End Function

'******************************************************************************
'* [概  要] パターンチェック処理。
'* [詳  細] 値がパターン通りかどうか判定する｡
'*
'* @param val 検証値
'* @param patternStr パターン文字列（正規表現）
'* @param igCase 大文字小文字を区別するか
'* @return 判定結果（True：合致）、
'*         検証・パターン文字列が設定されていない場合False
'*
'******************************************************************************
Public Function ValidatePattern(ByVal val As String, ByVal patternStr As String, Optional ByVal igCase As Boolean = False) As Boolean
    If Len(val) > 0 And patternStr <> "" Then
        With Base.GetRegExp
            .Pattern = patternStr
            .IgnoreCase = igCase
            .Global = True
            ValidatePattern = .test(val)
        End With
        Exit Function
    End If
    ValidatePattern = False
End Function
            
'******************************************************************************
'* [概  要] 半角チェック処理。
'* [詳  細] 値が半角文字かどうか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsHalf(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    Dim checkValue As String
    If Len(val) > 0 Then
        checkValue = StrConv(val, vbFromUnicode)
        If Len(val) <> LenB(checkValue) Then
            IsHalf = False
            Exit Function
        End If
    ElseIf required Then
        IsHalf = False
        Exit Function
    End If
    IsHalf = True
End Function

'******************************************************************************
'* [概  要] 半角チェック処理。
'* [詳  細] 値が半角文字のみではないか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsNotHalf(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    If Len(val) > 0 Then
    
        IsNotHalf = Not IsHalf(val, True)
        Exit Function
    ElseIf required Then
        IsNotHalf = False
        Exit Function
    End If
    IsNotHalf = True
End Function

'******************************************************************************
'* [概  要] 半角英数字チェック処理。
'* [詳  細] 値が半角英数字かどうか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsHalfAlphaNumeric(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    Dim i As Long
    If Len(val) > 0 Then
        For i = 1 To Len(val)
            If Not Mid(val, i, 1) Like "[0-9a-zA-Z]" Then
                IsHalfAlphaNumeric = False
                Exit Function
            End If
        Next
    ElseIf required Then
        IsHalfAlphaNumeric = False
        Exit Function
    End If
    IsHalfAlphaNumeric = True
End Function

'******************************************************************************
'* [概  要] 半角英数字チェック処理。
'* [詳  細] 値が半角英数字のみではないか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsNotHalfAlphaNumeric(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    If Len(val) > 0 Then
        IsNotHalfAlphaNumeric = Not IsHalfAlphaNumeric(val, True)
        Exit Function
    ElseIf required Then
        IsNotHalfAlphaNumeric = False
        Exit Function
    End If
    IsNotHalfAlphaNumeric = True
End Function

'******************************************************************************
'* [概  要] 半角英数字記号チェック処理。
'* [詳  細] 値が半角英数字記号かどうか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsHalfAlphaNumericSymbol(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    If Len(val) > 0 Then
        IsHalfAlphaNumericSymbol = ValidatePattern(val, "^[a-zA-Z0-9!-/:-@\[-`{-~]*$")
        Exit Function
    ElseIf required Then
        IsHalfAlphaNumericSymbol = False
        Exit Function
    End If
    IsHalfAlphaNumericSymbol = True
End Function

'******************************************************************************
'* [概  要] 半角英数字記号チェック処理。
'* [詳  細] 値が半角英数字記号のみではないか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsNotHalfAlphaNumericSymbol(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    If Len(val) > 0 Then
        IsNotHalfAlphaNumericSymbol = Not IsHalfAlphaNumericSymbol(val, True)
        Exit Function
    ElseIf required Then
        IsNotHalfAlphaNumericSymbol = False
        Exit Function
    End If
    IsNotHalfAlphaNumericSymbol = True
End Function

'******************************************************************************
'* [概  要] 数値チェック処理。
'* [詳  細] 値が数字かどうか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsNumber(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    Dim i As Long
    If Len(val) > 0 Then
        For i = 1 To Len(val)
            If Not Mid(val, i, 1) Like "[0-9]" Then
                IsNumber = False
                Exit Function
            End If
        Next
    ElseIf required Then
        IsNumber = False
        Exit Function
    End If
    IsNumber = True
End Function

'******************************************************************************
'* [概  要] 数値チェック処理。
'* [詳  細] 値が数字以外かどうか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsNotNumber(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    Dim i As Long
    If Len(val) > 0 Then
        IsNotNumber = Not IsNumber(val, True)
        Exit Function
    ElseIf required Then
        IsNotNumber = False
        Exit Function
    End If
    IsNotNumber = True
End Function

'******************************************************************************
'* [概  要] 全角カタカナチェック処理。
'* [詳  細] 値が全角カタカナかどうか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsZenKatakana(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    If Len(val) > 0 Then
        IsZenKatakana = ValidatePattern(val, "^[ァ-ンヴー]*$")
        Exit Function
    ElseIf required Then
        IsZenKatakana = False
        Exit Function
    End If
    IsZenKatakana = True
End Function

'******************************************************************************
'* [概  要] 全角カタカナチェック処理。
'* [詳  細] 値が全角カタカナ以外かどうか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsNotZenKatakana(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    If Len(val) > 0 Then
        IsNotZenKatakana = Not IsZenKatakana(val, True)
        Exit Function
    ElseIf required Then
        IsNotZenKatakana = False
        Exit Function
    End If
    IsNotZenKatakana = True
End Function

'******************************************************************************
'* [概  要] 全角ひらがなチェック処理。
'* [詳  細] 値が全角ひらがなかどうか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG））
'*
'******************************************************************************
Public Function IsZenHiragana(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    If Len(val) > 0 Then
        IsZenHiragana = ValidatePattern(val, "^[ぁ-んー]*$")
        Exit Function
    ElseIf required Then
        IsZenHiragana = False
        Exit Function
    End If
    IsZenHiragana = True
End Function

'******************************************************************************
'* [概  要] 全角ひらがなチェック処理。
'* [詳  細] 値が全角ひらがな以外かどうか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsNotZenHiragana(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    If Len(val) > 0 Then
        IsNotZenHiragana = Not IsZenHiragana(val, True)
        Exit Function
    ElseIf required Then
        IsNotZenHiragana = False
        Exit Function
    End If
    IsNotZenHiragana = True
End Function

'******************************************************************************
'* [概  要] 日付チェック処理。
'* [詳  細] 値が日付（yyyy/mm/dd）かどうか判定する｡
'*
'* @param val 検証値
'* @param separator 日付セパレータ。デフォルト「/」
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsDateString(ByVal val As String, Optional ByVal separator As String = "/", Optional ByVal required As Boolean = False) As Boolean
    Dim i As Long, tmp As String, tmpDate As String
    If Len(val) > 0 Then
        For i = 1 To Len(val)
            tmp = Mid(val, i, 1)
            If separator = "" Or (i <> 5 And i <> 8 And (tmp Like "[0-9]")) Then
                tmpDate = tmpDate & tmp
            ElseIf i = 5 Or i = 8 Then
                If separator <> "" And (tmp <> separator) Then
                    IsDateString = False
                    Exit Function
                End If
            End If
        Next
        If Len(tmpDate) <> 8 Then
            IsDateString = False
            Exit Function
        End If
        tmpDate = Format(tmpDate, "####/##/##")
        IsDateString = VBA.IsDate(tmpDate)
        Exit Function
    ElseIf required Then
        IsDateString = False
        Exit Function
    End If
    IsDateString = True
End Function

'******************************************************************************
'* [概  要] 日付チェック処理。
'* [詳  細] 値が日付（yyyy/mm/dd）以外かどうか判定する｡
'*
'* @param val 検証値
'* @param separator 日付セパレータ。デフォルト「/」
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsNotDateString(ByVal val As String, Optional ByVal separator As String = "/", Optional ByVal required As Boolean = False) As Boolean
    If Len(val) > 0 Then
        IsNotDateString = Not IsDateString(val, separator, True)
        Exit Function
    ElseIf required Then
        IsNotDateString = False
        Exit Function
    End If
    IsNotDateString = True
End Function

'******************************************************************************
'* [概  要] 日時チェック処理。
'* [詳  細] 値が日付（yyyy/mm/dd hh:nn:ss）かどうか判定する｡
'*
'* @param val 検証値
'* @param separator 日付セパレータ。デフォルト「/」
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsDateTimeString(ByVal val As String, Optional ByVal separator As String = "/", Optional ByVal required As Boolean = False) As Boolean
    Dim i As Long, tmp As String, tmpDate As String
    If Len(val) > 0 Then
        For i = 1 To Len(val)
            tmp = Mid(val, i, 1)
            If separator = "" Or _
               (i <> 5 And i <> 8 And i <> 11 And i <> 14 And i <> 17 And (tmp Like "[0-9]")) Then
                tmpDate = tmpDate & tmp
            ElseIf i = 5 Or i = 8 Then
                If separator <> "" And (tmp <> separator) Then
                    IsDateTimeString = False
                    Exit Function
                End If
            ElseIf i = 11 Then
                If separator <> "" And (tmp <> " ") Then
                    IsDateTimeString = False
                    Exit Function
                End If
            ElseIf i <> 14 And i <> 17 Then
                If separator <> "" And (tmp <> ":") Then
                    IsDateTimeString = False
                    Exit Function
                End If
            End If
        Next
        If Len(tmpDate) <> 14 Then
            IsDateTimeString = False
            Exit Function
        End If
        tmpDate = Format(tmpDate, "####/##/## ##:##:##")
        IsDateTimeString = VBA.IsDate(tmpDate)
        Exit Function
    ElseIf required Then
        IsDateTimeString = False
        Exit Function
    End If
    IsDateTimeString = True
End Function

'******************************************************************************
'* [概  要] 日時チェック処理。
'* [詳  細] 値が日付（yyyy/mm/dd hh:nn:ss）以外かどうか判定する｡
'*
'* @param val 検証値
'* @param separator 日付セパレータ。デフォルト「/」
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsNotDateTimeString(ByVal val As String, Optional ByVal separator As String = "/", Optional ByVal required As Boolean = False) As Boolean
    If Len(val) > 0 Then
        IsNotDateTimeString = Not IsDateTimeString(val, separator, True)
        Exit Function
    ElseIf required Then
        IsNotDateTimeString = False
        Exit Function
    End If
    IsNotDateTimeString = True
End Function

'******************************************************************************
'* [概  要] 郵便番号チェック処理。
'* [詳  細] 値が郵便番号かどうか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsZipCode(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    If Len(val) > 0 Then
        IsZipCode = ValidatePattern(val, "^\d{7}$|^\d{3}-\d{4}$")
        Exit Function
    ElseIf required Then
        IsZipCode = False
        Exit Function
    End If
    IsZipCode = True
End Function

'******************************************************************************
'* [概  要] 電話番号チェック処理。
'* [詳  細] 値が電話番号かどうか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsTelNo(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    If Len(val) > 0 Then
        IsTelNo = ValidatePattern(val, "^0\d{9,10}$|^0\d{2,3}-\d{1,4}-\d{4}$")
        Exit Function
    ElseIf required Then
        IsTelNo = False
        Exit Function
    End If
    IsTelNo = True
End Function

'******************************************************************************
'* [概  要] メールアドレスチェック処理。
'* [詳  細] 値がメールアドレスかどうか判定する｡
'*
'* @param val 検証値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsMailAddress(ByVal val As String, Optional ByVal required As Boolean = False) As Boolean
    If Len(val) > 0 Then
        IsMailAddress = ValidatePattern(val, "^[a-zA-Z0-9.!#$%&'*+\/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$")
        Exit Function
    ElseIf required Then
        IsMailAddress = False
        Exit Function
    End If
    IsMailAddress = True
End Function

'******************************************************************************
'* [概  要] 配列空チェック
'* [詳  細] 指定された配列内の要素のどれかが空かチェックを行う。
'*
'* @param arr Variant配列
'* @return Boolean チェック結果（True：どれかが空、False：全て値あり）
'******************************************************************************
Public Function IsEmptyAny(ByRef arr As Variant) As Boolean
    If IsNotArray(arr) Then
        If arr = "" Then
            IsEmptyAny = True
            Exit Function
        End If
        IsEmptyAny = False
        Exit Function
    End If
    Dim v
    For Each v In arr
        If v = "" Then
            IsEmptyAny = True
            Exit Function
        End If
    Next
    IsEmptyAny = False
End Function

'******************************************************************************
'* [概  要] 配列空チェック
'* [詳  細] 指定された配列内の要素の全てが空かチェックを行う。
'*
'* @param arr Variant配列
'* @return Boolean チェック結果（True：全てが空、False：どれかに値あり）
'******************************************************************************
Public Function IsEmptyAll(ByRef arr As Variant) As Boolean
    If IsNotArray(arr) Then
        If arr = "" Then
            IsEmptyAll = True
            Exit Function
        End If
        IsEmptyAll = False
        Exit Function
    End If
    Dim v
    For Each v In arr
        If v <> "" Then
            IsEmptyAll = False
            Exit Function
        End If
    Next
    IsEmptyAll = True
End Function

'******************************************************************************
'* [概  要] 配列値設定チェック
'* [詳  細] 指定された配列内の要素のどれかに値が設定されているかチェックを行う。
'*
'* @param arr Variant配列
'* @return Boolean チェック結果（True：どれかに設定、False：全て設定なし）
'******************************************************************************
Public Function IsFilledAny(ByRef arr As Variant) As Boolean
    IsFilledAny = Not IsEmptyAll(arr)
End Function

'******************************************************************************
'* [概  要] 配列値設定チェック
'* [詳  細] 指定された配列内の要素の全てに値が設定されているかチェックを行う。
'*
'* @param arr Variant配列
'* @return Boolean チェック結果（True：全て設定、False：どれかが空）
'******************************************************************************
Public Function IsFilledAll(ByRef arr As Variant) As Boolean
    IsFilledAll = Not IsEmptyAny(arr)
End Function

'******************************************************************************
'* [概  要] 過去日チェック処理。
'* [詳  細] 値が現在日より過去日かどうか判定する｡
'*
'* @param val 検証値
'* @param timeIgnore 時刻を無視するか（Falseを指定すると時刻も考慮する）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsPastDay(ByVal val As Date, Optional timeIgnore As Boolean = True) As Boolean
    IsPastDay = IsBeforeDay(val, Date, timeIgnore)
End Function

'******************************************************************************
'* [概  要] 過去日チェック処理。
'* [詳  細] 値が指定した日付より過去日かどうか判定する｡
'*
'* @param val 検証値
'* @param baseDate 基準日
'* @param timeIgnore 時刻を無視するか（Falseを指定すると時刻も考慮する）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsBeforeDay(ByVal val As Date, ByVal baseDate As Date, Optional timeIgnore As Boolean = True) As Boolean
    If timeIgnore Then
        val = DateValue(val)
        baseDate = DateValue(baseDate)
    End If

    If val < baseDate Then
        IsBeforeDay = True
        Exit Function
    End If
    IsBeforeDay = False
End Function

'******************************************************************************
'* [概  要] 未来日チェック処理。
'* [詳  細] 値が現在日より未来日かどうか判定する｡
'*
'* @param val 検証値
'* @param timeIgnore 時刻を無視するか（Falseを指定すると時刻も考慮する）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsFutureDay(ByVal val As Date, Optional timeIgnore As Boolean = True) As Boolean
    IsFutureDay = IsAfterDay(val, Date, timeIgnore)
End Function

'******************************************************************************
'* [概  要] 未来日チェック処理。
'* [詳  細] 値が指定した日付より未来日かどうか判定する｡
'*
'* @param val 検証値
'* @param baseDate 基準日
'* @param timeIgnore 時刻を無視するか（Falseを指定すると時刻も考慮する）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsAfterDay(ByVal val As Date, ByVal baseDate As Date, Optional timeIgnore As Boolean = True) As Boolean
    If timeIgnore Then
        val = DateValue(val)
        baseDate = DateValue(baseDate)
    End If

    If val > baseDate Then
        IsAfterDay = True
        Exit Function
    End If
    IsAfterDay = False
End Function

'******************************************************************************
'* [概  要] 当日チェック処理。
'* [詳  細] 値が現在日かどうか判定する｡
'*
'* @param val 検証値
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsToday(ByVal val As Date) As Boolean
    val = DateValue(val)
    If val = Date Then
        IsToday = True
        Exit Function
    End If
    IsToday = False
End Function

'******************************************************************************
'* [概  要] 整数範囲チェック。
'* [詳  細] 指定された値の範囲の整数であるかチェックを行う。
'*
'* @param checkVal チェック対象
'* @param startVal 開始数値
'* @param endVal   終了数値
'* @param required 必須か（デフォルトFalse）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function ValidateNumRange(ByVal checkVal As Variant, _
                                 ByVal startVal As Long, _
                                 ByVal endVal As Long, _
                                 Optional ByVal required As Boolean = False) As Boolean
    If required Then
        If IsNotPrimitive(checkVal) Then
            ValidateNumRange = False
            Exit Function
        End If
        If CStr(checkVal) = "" Then
            ValidateNumRange = False
            Exit Function
        End If
    End If
    
    ' 数字チェック
    If IsNotNumber(CStr(checkVal)) Then
        ValidateNumRange = False
        Exit Function
    End If
    
    ' 範囲チェック
    Dim lChkVal As Long
    lChkVal = CLng(checkVal)
    If lChkVal < startVal Or lChkVal > endVal Then
        ValidateNumRange = False
        Exit Function
    End If
    
    ValidateNumRange = True
End Function

'******************************************************************************
'* [概  要] 固定桁数チェック
'* [詳  細] 指定された値の固定桁数チェックを行う。
'*
'* @param checkVal チェック対象
'          pLength チェック桁数
'* @return Boolean チェック結果（True：False：NG）
'******************************************************************************
Public Function ValidateFixedLength(ByVal checkVal As Variant, _
                                    ByVal pLength As Long) As Boolean
    If IsNotPrimitive(checkVal) Then
        ValidateFixedLength = False
        Exit Function
    End If
    
    If Not (Len(checkVal) = pLength) Then
        ValidateFixedLength = False
        Exit Function
    End If
    ValidateFixedLength = True
End Function

'******************************************************************************
'* [概  要] 最大桁数チェック
'* [詳  細] 指定された値の最大桁数チェックを行う。
'*
'* @param checkVal チェック対象
'          pLength チェック桁数
'* @return Boolean チェック結果（True：False：NG）
'******************************************************************************
Public Function ValidateMaxLength(ByVal checkVal As Variant, _
                                  ByVal pLength As Long) As Boolean
    If IsNotPrimitive(checkVal) Then
        ValidateMaxLength = False
        Exit Function
    End If

    If Len(checkVal) > pLength Then
        ValidateMaxLength = False
        Exit Function
    End If

    ValidateMaxLength = True
End Function

'******************************************************************************
'* [概  要] 最小桁数チェック
'* [詳  細] 指定された値の最小桁数チェックを行う。
'*
'* @param checkVal チェック対象
'          pLength チェック桁数
'* @return Boolean チェック結果（True：False：NG）
'******************************************************************************
Public Function ValidateMinLength(ByVal checkVal As Variant, _
                                  ByVal pLength As Long) As Boolean
    If IsNotPrimitive(checkVal) Then
        ValidateMinLength = False
        Exit Function
    End If

    If Len(checkVal) < pLength Then
        ValidateMinLength = False
        Exit Function
    End If

    ValidateMinLength = True
End Function

'******************************************************************************
'* [概  要] 範囲桁数チェック
'* [詳  細] 指定された値の範囲桁数チェックを行う。
'*
'* @param checkVal チェック対象
'* @param minLength 最小桁数
'* @param maxLength 最大桁数
'* @return Boolean チェック結果（True：False：NG）
'******************************************************************************
Public Function ValidateRangeLength(ByVal checkVal As Variant, _
                                    ByVal minLength As Long, _
                                    ByVal maxLength As Long) As Boolean
    ValidateRangeLength = ValidateMinLength(checkVal, minLength) And _
                          ValidateMaxLength(checkVal, maxLength)
End Function

'******************************************************************************
'* [概  要] セルアドレス形式チェック
'* [詳  細] セルアドレス形式チェックを行う。Excel2007以降が対象。
'*
'* @param val チェック対象
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function ValidateCellAddress(ByVal val As Variant) As Boolean
    If IsNotString(val) Then
        ValidateCellAddress = False
        Exit Function
    End If

    Dim v
    For Each v In Split(val, ":")
        Dim cntAlph As Integer: cntAlph = 0
        Dim iLen As Integer: iLen = 0
        Dim i As Integer: i = 0
        
        iLen = Len(v)
        If iLen < 2 Or iLen > 10 Then
            ValidateCellAddress = False
            Exit Function
        End If
        
        ' 先頭からのアルファベットの数をチェック
        For i = 1 To iLen
            If Not (Mid(v, i, 1) Like "[A-Z]") Then
                cntAlph = i - 1
                Exit For
            End If
        Next i
        If cntAlph = 0 Or cntAlph > 3 Or cntAlph = iLen Then
            ValidateCellAddress = False
            Exit Function
        End If
        
        ' アルファベット部分と数値部分を切分け
        Dim strTmpAlph As String, strTmpNum As String
        strTmpAlph = Left(CStr(v), cntAlph)
        strTmpNum = Mid(CStr(v), cntAlph + 1)
        
        ' アルファベット3文字の場合はアルファベット部分をチェック
        If cntAlph = 3 Then
            If ValidatePattern(strTmpAlph, "[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]") Then
                ValidateCellAddress = False
                Exit Function
            End If
        End If
        
        ' 数値部分のチェック
        If Not ValidateNumRange(strTmpNum, 1, 1048576) Then
            ValidateCellAddress = False
            Exit Function
        End If
    Next
    ValidateCellAddress = True
End Function

'******************************************************************************
'* [概  要] 設定順チェック
'* [詳  細] 指定した配列Aの要素全てに値が設定されていない状態で、
'*          指定した配列Bの要素のいづれかに値が設定されている場合NGとする。
'*
'* @param firstArr 配列A
'* @param secondArr 配列B
'* @return Boolean チェック結果（True：OK、False：NG）
'******************************************************************************
Public Function ValidateSettingOrder(ByRef firstArr As Variant, _
                                    ByRef secondArr As Variant) As Boolean
    ValidateSettingOrder = Not (IsEmptyAll(firstArr) And IsFilledAny(secondArr))
End Function

