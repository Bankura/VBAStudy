VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VariantUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Variant関連ユーティリティクラス
'* [詳  細] Variant操作を行うユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************


'******************************************************************************
'* 変数定義
'******************************************************************************

'******************************************************************************
'* メソッド定義
'******************************************************************************
'******************************************************************************
'* [概  要] EmptyStringIfMissing
'* [詳  細] Optional引数が省略されている場合、空文字列に変換する｡
'*
'* @param arg 引数（Optional引数を渡す）
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function EmptyStringIfMissing(ByVal arg) As String
    If IsMissing(arg) Then
        EmptyStringIfMissing = ""
    End If
    EmptyStringIfMissing = arg
End Function

'******************************************************************************
'* [概  要] DefaultIfMissing
'* [詳  細] Optional引数が省略されている場合、デフォルト値に変換する｡
'*
'* @param arg 引数（Optional引数を渡す）
'* @param defaultValue デフォルト値
'* @return Variant 変換後Variant
'*
'******************************************************************************
Public Function DefaultIfMissing(ByVal arg, ByVal defaultValue) As Variant
    If IsMissing(arg) Then
        DefaultIfMissing = defaultValue
    End If
    DefaultIfMissing = arg
End Function


'*******************************************************************************
'* [概  要] ToJsonString
'* [詳  細] オブジェクト、プリミティブ型を文字列変換する。
'*          JSON形式を意識した文字列に変換するが、完全にJSON形式
'*          に準拠はしていない。
'*
'* @param txt Variant型データ
'* @return String 結果文字列
'*
'*******************************************************************************
Public Function ToString(ByVal var As Variant) As String
    If IsObject(var) Then
        ToString = ObjectToString(var)
        
    ElseIf IsArray(var) Then
        ToString = ArrayUtils.ToString(var)
        
    ElseIf ValidateUtils.IsString(var) Then
        ToString = """" & CStr(var) & """"
        
    ElseIf IsEmpty(var) Then
        ToString = "Empty"
        
    ElseIf IsNull(var) Then
        ToString = "Null"
        
    ElseIf IsError(var) Then
        ToString = "Error"
        
    Else
        ToString = CStr(var)
        
    End If
End Function

'*******************************************************************************
'* [概  要] ObjectToString
'* [詳  細] オブジェクトを文字列変換する。
'*          JSON形式を意識した文字列に変換するが、完全にJSON形式
'*          に準拠はしていない。
'*
'* @param var Variant型データ
'* @return String 結果文字列
'*
'*******************************************************************************
Public Function ObjectToString(ByVal var As Variant) As String
    If IsObject(var) Then
        If ValidateUtils.IsNothing(var) Then
            ObjectToString = "Nothing"
        
        ElseIf ValidateUtils.IsArrayEx(var) Then
            ObjectToString = ArrayUtils.ToString(var.ToArray)
        
        ElseIf ValidateUtils.IsStringEx(var) Then
            ObjectToString = """" & var.ToString & """"
        
        ElseIf ValidateUtils.IsCollection(var) Then
            ObjectToString = CollectionUtils.CollectionToString(var)
    
        ElseIf ValidateUtils.IsDictionary(var) Or ValidateUtils.IsDictionaryEx(var) Then
            ObjectToString = CollectionUtils.DictionaryToString(var)
        
        Else
            ObjectToString = "(" & TypeName(var) & ")"
        End If
    Else
        ObjectToString = "(no obj)"
    End If
End Function


