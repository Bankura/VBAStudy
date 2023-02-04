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
'* 構造体定義
'******************************************************************************
' 分数型
Private Type Fraction
    numerator As Long        ' 分子
    denominator As Long      ' 分母
End Type

'******************************************************************************
'* 定数定義
'******************************************************************************

'******************************************************************************
'* 変数定義
'******************************************************************************

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* PI プロパティ（読み取り専用）
'*
'* [補  足]
'* ・円周率を返す。
'*-----------------------------------------------------------------------------
Property Get PI() As Double
    PI = 4 * Atn(1)
End Property

'*-----------------------------------------------------------------------------
'* Radianプロパティ（読み取り専用）
'*
'* [補  足]
'* ・ラジアンを返す。
'*-----------------------------------------------------------------------------
Property Get Radian() As Double
    Radian = PI / 180
End Property

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
'* [概  要] IsEven
'* [詳  細] 偶数かどうかチェック｡
'*
'* @param num 対象数値
'* @return Boolean チェック結果（True：整数）
'******************************************************************************
Public Function IsEven(ByVal num As Long) As Boolean
    IsEven = (num Mod 2 = 0)
End Function

'******************************************************************************
'* [概  要] IsOdd
'* [詳  細] 奇数かどうかチェック｡
'*
'* @param num 対象数値
'* @return Boolean チェック結果（True：整数）
'******************************************************************************
Public Function IsOdd(ByVal num As Long) As Boolean
    IsOdd = (num Mod 2 = 1)
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

'******************************************************************************
'* [概  要] Sum
'* [詳  細] 指定した要素の合計値を返却する｡
'*
'* @param arr() 対象要素（複数）
'* @return Variant 最大の要素
'******************************************************************************
Public Function Sum(ParamArray parr() As Variant) As Double
    Dim tmp: tmp = parr
    Dim arr: arr = ArrayUtils.ToVArrayFlattened(tmp)

    Dim total As Double, v
    For Each v In arr
        total = total + CDbl(v)
    Next
    Sum = total
End Function

'******************************************************************************
'* [概  要] CalcVSphere
'* [詳  細] 球体の体積を計算する｡
'*
'* @param radius 半径
'* @return Double 球体の体積
'******************************************************************************
Public Function CalcVSphere(radius As Double) As Double
    CalcVSphere = 4 * PI * radius ^ 3 / 3
End Function


'******************************************************************************
'* [概  要] GetFractionString
'* [詳  細] 分数を約分し、文字列で取得する｡
'*
'* @param num 分子
'* @param den 分母
'* @return String 文字列の分数（num/den）
'******************************************************************************
Public Function GetFractionString(num As Long, den As Long) As String
    GetFractionString = Fraction2String(GetFraction(num, den))
End Function

'******************************************************************************
'* [概  要] Fraction2String
'* [詳  細] 分数を約分し、文字列で取得する｡
'*
'* @param num 分子
'* @param den 分母
'* @return String 文字列の分数（num/den）
'******************************************************************************
Private Function Fraction2String(f As Fraction) As String
    Fraction2String = f.numerator & "/" & f.denominator
End Function

'******************************************************************************
'* [概  要] GetFraction
'* [詳  細] 分数型を約分して取得する｡
'*
'* @param num 分子
'* @param den 分母
'* @return Fraction 分数型
'******************************************************************************
Private Function GetFraction(num As Long, den As Long) As Fraction
    ' aとbの最大公約数
    With New XlWsFunctions
        Dim fgcd As Long: fgcd = .Gcd(num, den)
    End With
    ' 約分
    GetFraction.numerator = num / fgcd
    GetFraction.denominator = den / fgcd
End Function

'******************************************************************************
'* [概  要] SumFraction
'* [詳  細] 分数型を加算する｡
'*
'* @param f1 分数型値１
'* @param f2 分数型値２
'* @return Fraction 分数型
'******************************************************************************
Private Function SumFraction(f1 As Fraction, f2 As Fraction) As Fraction
    Dim f3num As Long: f3num = f1.numerator * f2.denominator + f1.denominator * f2.numerator
    Dim f3den As Long: f3den = f1.denominator * f2.denominator

    With New XlWsFunctions
        Dim f3gcd As Long: f3gcd = .Gcd(f3num, f3den)
    End With

    SumFraction.numerator = f3num / f3gcd
    SumFraction.denominator = f3den / f3gcd
End Function

'******************************************************************************
'* [概  要] ProductFraction
'* [詳  細] 分数型を乗算する｡
'*
'* @param f1 分数型値１
'* @param f2 分数型値２
'* @return Fraction 分数型
'******************************************************************************
Private Function ProductFraction(f1 As Fraction, f2 As Fraction) As Fraction
    Dim f3num As Long: f3num = f1.numerator * f2.numerator
    Dim f3den As Long: f3den = f1.denominator * f2.denominator
    
    With New XlWsFunctions
        Dim f3gcd As Long: f3gcd = .Gcd(f3num, f3den)
    End With
    
    ProductFraction.numerator = f3num / f3gcd
    ProductFraction.denominator = f3den / f3gcd
End Function

'******************************************************************************
'* [概  要] DivFraction
'* [詳  細] 分数型を除算する｡
'*
'* @param f1 分数型値１
'* @param f2 分数型値２
'* @return Fraction 分数型
'******************************************************************************
Private Function DivFraction(f1 As Fraction, f2 As Fraction) As Fraction
    Dim f3num As Long: f3num = f1.numerator * f2.denominator
    Dim f3den As Long: f3den = f1.denominator * f2.numerator
    
    With New XlWsFunctions
        Dim f3gcd As Long: f3gcd = .Gcd(f3num, f3den)
    End With
    
    DivFraction.numerator = f3num / f3gcd
    DivFraction.denominator = f3den / f3gcd
End Function
