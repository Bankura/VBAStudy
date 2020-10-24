VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MathUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] 数値演算ユーティリティクラス
'* [詳  細] 数値演算を行うユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'*
'* [依　存]
'*
'* [参  考] https://github.com/RelaxTools/Hidennotare/wiki/StringUtils
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


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
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
End Sub
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] LShift
'* [詳  細] 左シフト｡
'*
'* @param val 値
'* @param digit 左シフトする桁数
'* @return String シフト結果
'******************************************************************************
Public Function LShift(ByVal val As Long, ByVal digit As Long) As Long
    LShift = val * (2 ^ digit)
End Function

'******************************************************************************
'* [概  要] RShift
'* [詳  細] 右シフト｡
'*
'* @param val 値
'* @param digit 右シフトする桁数
'* @return String シフト結果
'******************************************************************************
Public Function RShift(val As Long, digit As Long) As Long
    RShift = val \ (2 ^ digit)
End Function

'******************************************************************************
'* [概  要] Min
'* [詳  細] 指定した要素のうち最小の要素を返却する｡
'*
'* @param arr() 対象要素（複数）
'* @return Variant 最小の要素
'******************************************************************************
Public Function Min(ParamArray parr() As Variant) As Variant
    Dim tmp: tmp = parr
    Dim arr: arr = ArrayUtils.ToVArrayFlattened(tmp)
    Dim ret: ret = Empty
    
    Dim ub As Variant: ub = UBound(arr)
    If ub < 0 Then
        Min = ret
        Exit Function
    End If
    
    Dim i As Long
    If IsObject(arr(0)) Then
        Set ret = arr(0)
        For i = 1 To ub
            If Compare(arr(i), ret) = -1 Then Set ret = arr(i)
        Next
    Else
        Let ret = arr(0)
        For i = 1 To ub
            If Compare(arr(i), ret) = -1 Then Let ret = arr(i)
        Next
    End If
    Min = ret
End Function

'******************************************************************************
'* [概  要] Max
'* [詳  細] 指定した要素のうち最大の要素を返却する｡
'*
'* @param arr() 対象要素（複数）
'* @return Variant 最大の要素
'******************************************************************************
Public Function Max(ParamArray parr() As Variant) As Variant
    Dim tmp: tmp = parr
    Dim arr: arr = ArrayUtils.ToVArrayFlattened(tmp)
    Dim ret: ret = Empty
    
    Dim ub As Variant: ub = UBound(arr)
    If ub < 0 Then
        Max = ret
        Exit Function
    End If
    
    Dim i As Long
    If IsObject(arr(0)) Then
        Set ret = arr(0)
        For i = 1 To ub
            If Compare(arr(i), ret) = 1 Then Set ret = arr(i)
        Next
    Else
        Let ret = arr(0)
        For i = 1 To ub
            If Compare(arr(i), ret) = 1 Then Let ret = arr(i)
        Next
    End If
    Max = ret
End Function

'******************************************************************************
'* [概  要] IsInt
'* [詳  細] 整数かどうかチェック｡
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param num 対象Variant値
'* @return Boolean チェック結果（True：整数）
'******************************************************************************
Public Function IsInt(ByVal num As Variant) As Boolean
    IsInt = Core.IsInt(num)
End Function

'******************************************************************************
'* [概  要] ARound
'* [詳  細] 四捨五入等の端数処理を行う。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param num 対象の小数を含む数値
'* @param digits 端数処理を行う小数第n位のn+1を示す
'* @param rndup 切り上げを行う数（デフォルト:5 四捨五入）
'* @return Variant(Of Decimal) 端数処理を行った結果の数値
'******************************************************************************
Public Function ARound( _
    ByVal num As Variant, Optional ByVal digits As Integer = 0, Optional rndup As Integer = 5 _
    ) As Variant
    ARound = Core.ARound(num, digits, rndup)
End Function
