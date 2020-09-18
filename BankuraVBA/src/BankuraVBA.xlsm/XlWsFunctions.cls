VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XlWsFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Excel WorksheetFunctionラップクラス
'* [詳  細] Excel WorksheetFunctionの各メソッドをラップして提供する。
'*
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mXlWsFunc As Object

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Dim app As Object
    If CheckXlApplication = True Then
        Set app = Application
    Else
        Set app = CreateObject("Excel.Application")
    End If
    Set mXlWsFunc = app.WorksheetFunction
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************
Public Function AccrInt(arg1, arg2, arg3, arg4, arg5, arg6, Optional arg7) As Double
    AccrInt = mXlWsFunc.AccrInt(arg1, arg2, arg3, arg4, arg5, arg6, arg7)
End Function

Public Function AccrIntM(arg1, arg2, arg3, arg4, Optional arg5) As Double
    AccrIntM = mXlWsFunc.AccrIntM(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function Acos(arg1 As Double) As Double
    Acos = mXlWsFunc.Acos(arg1)
End Function

Public Function Acosh(arg1 As Double) As Double
    Acosh = mXlWsFunc.Acosh(arg1)
End Function

Public Function Acot(arg1 As Double) As Double
    Acot = mXlWsFunc.Acot(arg1)
End Function

Public Function Acoth(arg1 As Double) As Double
    Acoth = mXlWsFunc.Acoth(arg1)
End Function

Public Function Aggregate(arg1 As Double, arg2 As Double, arg3 As Object, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Aggregate = mXlWsFunc.Aggregate(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function AmorDegrc(arg1, arg2, arg3, arg4, arg5, arg6, Optional arg7) As Double
    AmorDegrc = mXlWsFunc.AmorDegrc(arg1, arg2, arg3, arg4, arg5, arg6, arg7)
End Function

Public Function AmorLinc(arg1, arg2, arg3, arg4, arg5, arg6, Optional arg7) As Double
    AmorLinc = mXlWsFunc.AmorLinc(arg1, arg2, arg3, arg4, arg5, arg6, arg7)
End Function

Public Function And_(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Boolean
    And_ = mXlWsFunc.And(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Arabic(arg1 As String) As Double
    Arabic = mXlWsFunc.Arabic(arg1)
End Function

Public Function Asc(arg1 As String) As String
    Asc = mXlWsFunc.Asc(arg1)
End Function

Public Function Asin(arg1 As Double) As Double
    Asin = mXlWsFunc.Asin(arg1)
End Function

Public Function Asinh(arg1 As Double) As Double
    Asinh = mXlWsFunc.Asinh(arg1)
End Function

Public Function Atan2(arg1 As Double, arg2 As Double) As Double
    Atan2 = mXlWsFunc.Atan2(arg1, arg2)
End Function

Public Function Atanh(arg1 As Double) As Double
    Atanh = mXlWsFunc.Atanh(arg1)
End Function

Public Function AveDev(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    AveDev = mXlWsFunc.AveDev(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Average(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Average = mXlWsFunc.Average(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function AverageIf(arg1 As Object, arg2, Optional arg3) As Double
    AverageIf = mXlWsFunc.AverageIf(arg1, arg2, arg3)
End Function

Public Function AverageIfs(arg1 As Object, arg2 As Object, arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29) As Double
    AverageIfs = mXlWsFunc.AverageIfs(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29)
End Function

Public Function BahtText(arg1 As Double) As String
    BahtText = mXlWsFunc.BahtText(arg1)
End Function

Public Function Base(arg1 As Double, arg2 As Double, Optional arg3) As String
    Base = mXlWsFunc.Base(arg1, arg2, arg3)
End Function

Public Function BesselI(arg1, arg2) As Double
    BesselI = mXlWsFunc.BesselI(arg1, arg2)
End Function

Public Function BesselJ(arg1, arg2) As Double
    BesselJ = mXlWsFunc.BesselJ(arg1, arg2)
End Function

Public Function BesselK(arg1, arg2) As Double
    BesselK = mXlWsFunc.BesselK(arg1, arg2)
End Function

Public Function BesselY(arg1, arg2) As Double
    BesselY = mXlWsFunc.BesselY(arg1, arg2)
End Function

Public Function Beta_Dist(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Boolean, Optional arg5, Optional arg6) As Double
    Beta_Dist = mXlWsFunc.Beta_Dist(arg1, arg2, arg3, arg4, arg5, arg6)
End Function

Public Function Beta_Inv(arg1 As Double, arg2 As Double, arg3 As Double, Optional arg4, Optional arg5) As Double
    Beta_Inv = mXlWsFunc.Beta_Inv(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function Bin2Dec(arg1) As String
    Bin2Dec = mXlWsFunc.Bin2Dec(arg1)
End Function

Public Function Bin2Hex(arg1, Optional arg2) As String
    Bin2Hex = mXlWsFunc.Bin2Hex(arg1, arg2)
End Function

Public Function Bin2Oct(arg1, Optional arg2) As String
    Bin2Oct = mXlWsFunc.Bin2Oct(arg1, arg2)
End Function

Public Function Binom_Dist(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Boolean) As Double
    Binom_Dist = mXlWsFunc.Binom_Dist(arg1, arg2, arg3, arg4)
End Function

Public Function Binom_Dist_Object(arg1 As Double, arg2 As Double, arg3 As Double, Optional arg4) As Double
    Binom_Dist_Object = mXlWsFunc.Binom_Dist_Object(arg1, arg2, arg3, arg4)
End Function

Public Function Binom_Inv(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    Binom_Inv = mXlWsFunc.Binom_Inv(arg1, arg2, arg3)
End Function

Public Function BinomDist(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Boolean) As Double
    BinomDist = mXlWsFunc.BinomDist(arg1, arg2, arg3, arg4)
End Function

Public Function Bitand(arg1 As Double, arg2 As Double) As Double
    Bitand = mXlWsFunc.Bitand(arg1, arg2)
End Function

Public Function Bitlshift(arg1 As Double, arg2 As Double) As Double
    Bitlshift = mXlWsFunc.Bitlshift(arg1, arg2)
End Function

Public Function Bitor(arg1 As Double, arg2 As Double) As Double
    Bitor = mXlWsFunc.Bitor(arg1, arg2)
End Function

Public Function Bitrshift(arg1 As Double, arg2 As Double) As Double
    Bitrshift = mXlWsFunc.Bitrshift(arg1, arg2)
End Function

Public Function Bitxor(arg1 As Double, arg2 As Double) As Double
    Bitxor = mXlWsFunc.Bitxor(arg1, arg2)
End Function

Public Function Ceiling(arg1 As Double, arg2 As Double) As Double
    Ceiling = mXlWsFunc.Ceiling(arg1, arg2)
End Function

Public Function Ceiling_Math(arg1 As Double, Optional arg2, Optional arg3) As Double
    Ceiling_Math = mXlWsFunc.Ceiling_Math(arg1, arg2, arg3)
End Function

Public Function Ceiling_Precise(arg1 As Double, Optional arg2) As Double
    Ceiling_Precise = mXlWsFunc.Ceiling_Precise(arg1, arg2)
End Function

Public Function ChiDist(arg1 As Double, arg2 As Double) As Double
    ChiDist = mXlWsFunc.ChiDist(arg1, arg2)
End Function

Public Function ChiInv(arg1 As Double, arg2 As Double) As Double
    ChiInv = mXlWsFunc.ChiInv(arg1, arg2)
End Function

Public Function ChiSq_Dist(arg1 As Double, arg2 As Double, arg3 As Boolean) As Double
    ChiSq_Dist = mXlWsFunc.ChiSq_Dist(arg1, arg2, arg3)
End Function

Public Function ChiSq_Dist_RT(arg1 As Double, arg2 As Double) As Double
    ChiSq_Dist_RT = mXlWsFunc.ChiSq_Dist_RT(arg1, arg2)
End Function

Public Function ChiSq_Inv(arg1 As Double, arg2 As Double) As Double
    ChiSq_Inv = mXlWsFunc.ChiSq_Inv(arg1, arg2)
End Function

Public Function ChiSq_Inv_RT(arg1 As Double, arg2 As Double) As Double
    ChiSq_Inv_RT = mXlWsFunc.ChiSq_Inv_RT(arg1, arg2)
End Function

Public Function ChiSq_Test(arg1, arg2) As Double
    ChiSq_Test = mXlWsFunc.ChiSq_Test(arg1, arg2)
End Function

Public Function ChiTest(arg1, arg2) As Double
    ChiTest = mXlWsFunc.ChiTest(arg1, arg2)
End Function

Public Function Choose(arg1, arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30)
    Choose = mXlWsFunc.Choose(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Clean(arg1 As String) As String
    Clean = mXlWsFunc.Clean(arg1)
End Function

Public Function Combin(arg1 As Double, arg2 As Double) As Double
    Combin = mXlWsFunc.Combin(arg1, arg2)
End Function

Public Function Combina(arg1 As Double, arg2 As Double) As Double
    Combina = mXlWsFunc.Combina(arg1, arg2)
End Function

Public Function Complex(arg1, arg2, Optional arg3) As String
    Complex = mXlWsFunc.Complex(arg1, arg2, arg3)
End Function

Public Function Concat(arg1 As String, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29) As String
    Concat = mXlWsFunc.Concat(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29)
End Function

Public Function Confidence(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    Confidence = mXlWsFunc.Confidence(arg1, arg2, arg3)
End Function

Public Function Confidence_Norm(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    Confidence_Norm = mXlWsFunc.Confidence_Norm(arg1, arg2, arg3)
End Function

Public Function Confidence_T(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    Confidence_T = mXlWsFunc.Confidence_T(arg1, arg2, arg3)
End Function

Public Function Convert(arg1, arg2, arg3) As Double
    Convert = mXlWsFunc.Convert(arg1, arg2, arg3)
End Function

Public Function Correl(arg1, arg2) As Double
    Correl = mXlWsFunc.Correl(arg1, arg2)
End Function

Public Function Cosh(arg1 As Double) As Double
    Cosh = mXlWsFunc.Cosh(arg1)
End Function

Public Function Cot(arg1 As Double) As Double
    Cot = mXlWsFunc.Cot(arg1)
End Function

Public Function Coth(arg1 As Double) As Double
    Coth = mXlWsFunc.Coth(arg1)
End Function

Public Function Count(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Count = mXlWsFunc.Count(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function CountA(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    CountA = mXlWsFunc.CountA(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function CountBlank(arg1 As Object) As Double
    CountBlank = mXlWsFunc.CountBlank(arg1)
End Function

Public Function CountIf(arg1 As Object, arg2) As Double
    CountIf = mXlWsFunc.CountIf(arg1, arg2)
End Function

Public Function CountIfs(arg1 As Object, arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    CountIfs = mXlWsFunc.CountIfs(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function CoupDayBs(arg1, arg2, arg3, Optional arg4) As Double
    CoupDayBs = mXlWsFunc.CoupDayBs(arg1, arg2, arg3, arg4)
End Function

Public Function CoupDays(arg1, arg2, arg3, Optional arg4) As Double
    CoupDays = mXlWsFunc.CoupDays(arg1, arg2, arg3, arg4)
End Function

Public Function CoupDaysNc(arg1, arg2, arg3, Optional arg4) As Double
    CoupDaysNc = mXlWsFunc.CoupDaysNc(arg1, arg2, arg3, arg4)
End Function

Public Function CoupNcd(arg1, arg2, arg3, Optional arg4) As Double
    CoupNcd = mXlWsFunc.CoupNcd(arg1, arg2, arg3, arg4)
End Function

Public Function CoupNum(arg1, arg2, arg3, Optional arg4) As Double
    CoupNum = mXlWsFunc.CoupNum(arg1, arg2, arg3, arg4)
End Function

Public Function CoupPcd(arg1, arg2, arg3, Optional arg4) As Double
    CoupPcd = mXlWsFunc.CoupPcd(arg1, arg2, arg3, arg4)
End Function

Public Function Covar(arg1, arg2) As Double
    Covar = mXlWsFunc.Covar(arg1, arg2)
End Function

Public Function Covariance_P(arg1, arg2) As Double
    Covariance_P = mXlWsFunc.Covariance_P(arg1, arg2)
End Function

Public Function Covariance_S(arg1, arg2) As Double
    Covariance_S = mXlWsFunc.Covariance_S(arg1, arg2)
End Function

Public Function CritBinom(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    CritBinom = mXlWsFunc.CritBinom(arg1, arg2, arg3)
End Function

Public Function Csc(arg1 As Double) As Double
    Csc = mXlWsFunc.Csc(arg1)
End Function

Public Function Csch(arg1 As Double) As Double
    Csch = mXlWsFunc.Csch(arg1)
End Function

Public Function CumIPmt(arg1, arg2, arg3, arg4, arg5, arg6) As Double
    CumIPmt = mXlWsFunc.CumIPmt(arg1, arg2, arg3, arg4, arg5, arg6)
End Function

Public Function CumPrinc(arg1, arg2, arg3, arg4, arg5, arg6) As Double
    CumPrinc = mXlWsFunc.CumPrinc(arg1, arg2, arg3, arg4, arg5, arg6)
End Function

Public Function DAverage(arg1 As Object, arg2, arg3) As Double
    DAverage = mXlWsFunc.DAverage(arg1, arg2, arg3)
End Function

Public Function days(arg1, arg2) As Double
    days = mXlWsFunc.days(arg1, arg2)
End Function

Public Function Days360(arg1, arg2, Optional arg3) As Double
    Days360 = mXlWsFunc.Days360(arg1, arg2, arg3)
End Function

Public Function Db(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Double, Optional arg5) As Double
    Db = mXlWsFunc.Db(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function Dbcs(arg1 As String) As String
    Dbcs = mXlWsFunc.Dbcs(arg1)
End Function

Public Function DCount(arg1 As Object, arg2, arg3) As Double
    DCount = mXlWsFunc.DCount(arg1, arg2, arg3)
End Function

Public Function DCountA(arg1 As Object, arg2, arg3) As Double
    DCountA = mXlWsFunc.DCountA(arg1, arg2, arg3)
End Function

Public Function Ddb(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Double, Optional arg5) As Double
    Ddb = mXlWsFunc.Ddb(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function Dec2Bin(arg1, Optional arg2) As String
    Dec2Bin = mXlWsFunc.Dec2Bin(arg1, arg2)
End Function

Public Function Dec2Hex(arg1, Optional arg2) As String
    Dec2Hex = mXlWsFunc.Dec2Hex(arg1, arg2)
End Function

Public Function Dec2Oct(arg1, Optional arg2) As String
    Dec2Oct = mXlWsFunc.Dec2Oct(arg1, arg2)
End Function

Public Function Decimal_(arg1 As String, arg2 As Double) As Double
    Decimal_ = mXlWsFunc.Decimal(arg1, arg2)
End Function

Public Function Degrees(arg1 As Double) As Double
    Degrees = mXlWsFunc.Degrees(arg1)
End Function

Public Function Delta(arg1, Optional arg2) As Double
    Delta = mXlWsFunc.Delta(arg1, arg2)
End Function

Public Function DevSq(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    DevSq = mXlWsFunc.DevSq(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function DGet(arg1 As Object, arg2, arg3)
    DGet = mXlWsFunc.DGet(arg1, arg2, arg3)
End Function

Public Function Disc(arg1, arg2, arg3, arg4, Optional arg5) As Double
    Disc = mXlWsFunc.Disc(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function DMax(arg1 As Object, arg2, arg3) As Double
    DMax = mXlWsFunc.DMax(arg1, arg2, arg3)
End Function

Public Function DMin(arg1 As Object, arg2, arg3) As Double
    DMin = mXlWsFunc.DMin(arg1, arg2, arg3)
End Function

Public Function Dollar(arg1 As Double, Optional arg2) As String
    Dollar = mXlWsFunc.Dollar(arg1, arg2)
End Function

Public Function DollarDe(arg1, arg2) As Double
    DollarDe = mXlWsFunc.DollarDe(arg1, arg2)
End Function

Public Function DollarFr(arg1, arg2) As Double
    DollarFr = mXlWsFunc.DollarFr(arg1, arg2)
End Function

Public Function DProduct(arg1 As Object, arg2, arg3) As Double
    DProduct = mXlWsFunc.DProduct(arg1, arg2, arg3)
End Function

Public Function DStDev(arg1 As Object, arg2, arg3) As Double
    DStDev = mXlWsFunc.DStDev(arg1, arg2, arg3)
End Function

Public Function DStDevP(arg1 As Object, arg2, arg3) As Double
    DStDevP = mXlWsFunc.DStDevP(arg1, arg2, arg3)
End Function

Public Function DSum(arg1 As Object, arg2, arg3) As Double
    DSum = mXlWsFunc.DSum(arg1, arg2, arg3)
End Function

Public Function Duration(arg1, arg2, arg3, arg4, arg5, Optional arg6) As Double
    Duration = mXlWsFunc.Duration(arg1, arg2, arg3, arg4, arg5, arg6)
End Function

Public Function DVar(arg1 As Object, arg2, arg3) As Double
    DVar = mXlWsFunc.DVar(arg1, arg2, arg3)
End Function

Public Function DVarP(arg1 As Object, arg2, arg3) As Double
    DVarP = mXlWsFunc.DVarP(arg1, arg2, arg3)
End Function

Public Function EDate(arg1, arg2) As Double
    EDate = mXlWsFunc.EDate(arg1, arg2)
End Function

Public Function Effect(arg1, arg2) As Double
    Effect = mXlWsFunc.Effect(arg1, arg2)
End Function

Public Function EncodeUrl(arg1 As String)
    EncodeUrl = mXlWsFunc.EncodeUrl(arg1)
End Function

Public Function EoMonth(arg1, arg2) As Double
    EoMonth = mXlWsFunc.EoMonth(arg1, arg2)
End Function

Public Function Erf(arg1, Optional arg2) As Double
    Erf = mXlWsFunc.Erf(arg1, arg2)
End Function

Public Function Erf_Precise(arg1) As Double
    Erf_Precise = mXlWsFunc.Erf_Precise(arg1)
End Function

Public Function ErfC(arg1) As Double
    ErfC = mXlWsFunc.ErfC(arg1)
End Function

Public Function ErfC_Precise(arg1) As Double
    ErfC_Precise = mXlWsFunc.ErfC_Precise(arg1)
End Function

Public Function Even(arg1 As Double) As Double
    Even = mXlWsFunc.Even(arg1)
End Function

Public Function Expon_Dist(arg1 As Double, arg2 As Double, arg3 As Boolean) As Double
    Expon_Dist = mXlWsFunc.Expon_Dist(arg1, arg2, arg3)
End Function

Public Function ExponDist(arg1 As Double, arg2 As Double, arg3 As Boolean) As Double
    ExponDist = mXlWsFunc.ExponDist(arg1, arg2, arg3)
End Function

Public Function F_Dist(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Boolean) As Double
    F_Dist = mXlWsFunc.F_Dist(arg1, arg2, arg3, arg4)
End Function

Public Function F_Dist_RT(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    F_Dist_RT = mXlWsFunc.F_Dist_RT(arg1, arg2, arg3)
End Function

Public Function F_Inv(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    F_Inv = mXlWsFunc.F_Inv(arg1, arg2, arg3)
End Function

Public Function F_Inv_RT(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    F_Inv_RT = mXlWsFunc.F_Inv_RT(arg1, arg2, arg3)
End Function

Public Function F_Test(arg1, arg2) As Double
    F_Test = mXlWsFunc.F_Test(arg1, arg2)
End Function

Public Function Fact(arg1 As Double) As Double
    Fact = mXlWsFunc.Fact(arg1)
End Function

Public Function FactDouble(arg1) As Double
    FactDouble = mXlWsFunc.FactDouble(arg1)
End Function

Public Function FDist(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    FDist = mXlWsFunc.FDist(arg1, arg2, arg3)
End Function

Public Function FieldValue(arg1, arg2 As String)
    FieldValue = mXlWsFunc.FieldValue(arg1, arg2)
End Function

Public Function Filter(arg1, arg2, Optional arg3)
    Filter = mXlWsFunc.Filter(arg1, arg2, arg3)
End Function

Public Function FilterXML(arg1 As String, arg2 As String)
    FilterXML = mXlWsFunc.FilterXML(arg1, arg2)
End Function

Public Function Find(arg1 As String, arg2 As String, Optional arg3) As Double
    Find = mXlWsFunc.Find(arg1, arg2, arg3)
End Function

Public Function FindB(arg1 As String, arg2 As String, Optional arg3) As Double
    FindB = mXlWsFunc.FindB(arg1, arg2, arg3)
End Function

Public Function FInv(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    FInv = mXlWsFunc.FInv(arg1, arg2, arg3)
End Function

Public Function Fisher(arg1 As Double) As Double
    Fisher = mXlWsFunc.Fisher(arg1)
End Function

Public Function FisherInv(arg1 As Double) As Double
    FisherInv = mXlWsFunc.FisherInv(arg1)
End Function

Public Function Fixed(arg1 As Double, Optional arg2, Optional arg3) As String
    Fixed = mXlWsFunc.Fixed(arg1, arg2, arg3)
End Function

Public Function Floor(arg1 As Double, arg2 As Double) As Double
    Floor = mXlWsFunc.Floor(arg1, arg2)
End Function

Public Function Floor_Math(arg1 As Double, Optional arg2, Optional arg3) As Double
    Floor_Math = mXlWsFunc.Floor_Math(arg1, arg2, arg3)
End Function

Public Function Floor_Precise(arg1 As Double, Optional arg2) As Double
    Floor_Precise = mXlWsFunc.Floor_Precise(arg1, arg2)
End Function

Public Function Forecast_ETS(arg1 As Double, arg2, arg3, Optional arg4, Optional arg5, Optional arg6) As Double
    Forecast_ETS = mXlWsFunc.Forecast_ETS(arg1, arg2, arg3, arg4, arg5, arg6)
End Function

Public Function Forecast_ETS_ConfInt(arg1 As Double, arg2, arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7) As Double
    Forecast_ETS_ConfInt = mXlWsFunc.Forecast_ETS_ConfInt(arg1, arg2, arg3, arg4, arg5, arg6, arg7)
End Function

Public Function Forecast_ETS_Seasonality(arg1, arg2, Optional arg3, Optional arg4) As Double
    Forecast_ETS_Seasonality = mXlWsFunc.Forecast_ETS_Seasonality(arg1, arg2, arg3, arg4)
End Function

Public Function Forecast_ETS_STAT(arg1, arg2, arg3 As Double, Optional arg4, Optional arg5, Optional arg6) As Double
    Forecast_ETS_STAT = mXlWsFunc.Forecast_ETS_STAT(arg1, arg2, arg3, arg4, arg5, arg6)
End Function

Public Function Forecast_Linear(arg1 As Double, arg2, arg3) As Double
    Forecast_Linear = mXlWsFunc.Forecast_Linear(arg1, arg2, arg3)
End Function

Public Function Frequency(arg1, arg2)
    Frequency = mXlWsFunc.Frequency(arg1, arg2)
End Function

Public Function FTest(arg1, arg2) As Double
    FTest = mXlWsFunc.FTest(arg1, arg2)
End Function

Public Function Fv(arg1 As Double, arg2 As Double, arg3 As Double, Optional arg4, Optional arg5) As Double
    Fv = mXlWsFunc.Fv(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function FVSchedule(arg1, arg2) As Double
    FVSchedule = mXlWsFunc.FVSchedule(arg1, arg2)
End Function

Public Function Gamma(arg1 As Double) As Double
    Gamma = mXlWsFunc.Gamma(arg1)
End Function

Public Function Gamma_Dist(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Boolean) As Double
    Gamma_Dist = mXlWsFunc.Gamma_Dist(arg1, arg2, arg3, arg4)
End Function

Public Function Gamma_Inv(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    Gamma_Inv = mXlWsFunc.Gamma_Inv(arg1, arg2, arg3)
End Function

Public Function GammaDist(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Boolean) As Double
    GammaDist = mXlWsFunc.GammaDist(arg1, arg2, arg3, arg4)
End Function

Public Function GammaInv(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    GammaInv = mXlWsFunc.GammaInv(arg1, arg2, arg3)
End Function

Public Function GammaLn(arg1 As Double) As Double
    GammaLn = mXlWsFunc.GammaLn(arg1)
End Function

Public Function GammaLn_Precise(arg1 As Double) As Double
    GammaLn_Precise = mXlWsFunc.GammaLn_Precise(arg1)
End Function

Public Function Gauss(arg1 As Double) As Double
    Gauss = mXlWsFunc.Gauss(arg1)
End Function

Public Function Gcd(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Gcd = mXlWsFunc.Gcd(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function GeoMean(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    GeoMean = mXlWsFunc.GeoMean(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function GeStep(arg1, Optional arg2) As Double
    GeStep = mXlWsFunc.GeStep(arg1, arg2)
End Function

Public Function Growth(arg1, Optional arg2, Optional arg3, Optional arg4)
    Growth = mXlWsFunc.Growth(arg1, arg2, arg3, arg4)
End Function

Public Function HarMean(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    HarMean = mXlWsFunc.HarMean(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Hex2Bin(arg1, Optional arg2) As String
    Hex2Bin = mXlWsFunc.Hex2Bin(arg1, arg2)
End Function

Public Function Hex2Dec(arg1) As String
    Hex2Dec = mXlWsFunc.Hex2Dec(arg1)
End Function

Public Function Hex2Oct(arg1, Optional arg2) As String
    Hex2Oct = mXlWsFunc.Hex2Oct(arg1, arg2)
End Function

Public Function HLookup(arg1, arg2, arg3, Optional arg4)
    HLookup = mXlWsFunc.HLookup(arg1, arg2, arg3, arg4)
End Function

Public Function HypGeom_Dist(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Double, arg5 As Boolean) As Double
    HypGeom_Dist = mXlWsFunc.HypGeom_Dist(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function HypGeomDist(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Double) As Double
    HypGeomDist = mXlWsFunc.HypGeomDist(arg1, arg2, arg3, arg4)
End Function

Public Function IfError(arg1, arg2)
    IfError = mXlWsFunc.IfError(arg1, arg2)
End Function

Public Function IfNa(arg1, arg2)
    IfNa = mXlWsFunc.IfNa(arg1, arg2)
End Function

Public Function ImAbs(arg1) As String
    ImAbs = mXlWsFunc.ImAbs(arg1)
End Function

Public Function Imaginary(arg1) As Double
    Imaginary = mXlWsFunc.Imaginary(arg1)
End Function

Public Function ImArgument(arg1) As String
    ImArgument = mXlWsFunc.ImArgument(arg1)
End Function

Public Function ImConjugate(arg1) As String
    ImConjugate = mXlWsFunc.ImConjugate(arg1)
End Function

Public Function ImCos(arg1) As String
    ImCos = mXlWsFunc.ImCos(arg1)
End Function

Public Function ImCosh(arg1) As String
    ImCosh = mXlWsFunc.ImCosh(arg1)
End Function

Public Function ImCot(arg1) As String
    ImCot = mXlWsFunc.ImCot(arg1)
End Function

Public Function ImCsc(arg1) As String
    ImCsc = mXlWsFunc.ImCsc(arg1)
End Function

Public Function ImCsch(arg1) As String
    ImCsch = mXlWsFunc.ImCsch(arg1)
End Function

Public Function ImDiv(arg1, arg2) As String
    ImDiv = mXlWsFunc.ImDiv(arg1, arg2)
End Function

Public Function ImExp(arg1) As String
    ImExp = mXlWsFunc.ImExp(arg1)
End Function

Public Function ImLn(arg1) As String
    ImLn = mXlWsFunc.ImLn(arg1)
End Function

Public Function ImLog10(arg1) As String
    ImLog10 = mXlWsFunc.ImLog10(arg1)
End Function

Public Function ImLog2(arg1) As String
    ImLog2 = mXlWsFunc.ImLog2(arg1)
End Function

Public Function ImPower(arg1, arg2) As String
    ImPower = mXlWsFunc.ImPower(arg1, arg2)
End Function

Public Function ImProduct(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As String
    ImProduct = mXlWsFunc.ImProduct(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function ImReal(arg1) As Double
    ImReal = mXlWsFunc.ImReal(arg1)
End Function

Public Function ImSec(arg1) As String
    ImSec = mXlWsFunc.ImSec(arg1)
End Function

Public Function ImSech(arg1) As String
    ImSech = mXlWsFunc.ImSech(arg1)
End Function

Public Function ImSin(arg1) As String
    ImSin = mXlWsFunc.ImSin(arg1)
End Function

Public Function ImSinh(arg1) As String
    ImSinh = mXlWsFunc.ImSinh(arg1)
End Function

Public Function ImSqrt(arg1) As String
    ImSqrt = mXlWsFunc.ImSqrt(arg1)
End Function

Public Function ImSub(arg1, arg2) As String
    ImSub = mXlWsFunc.ImSub(arg1, arg2)
End Function

Public Function ImSum(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As String
    ImSum = mXlWsFunc.ImSum(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function ImTan(arg1) As String
    ImTan = mXlWsFunc.ImTan(arg1)
End Function

Public Function Index(arg1, arg2 As Double, Optional arg3, Optional arg4)
    Index = mXlWsFunc.Index(arg1, arg2, arg3, arg4)
End Function

Public Function Intercept(arg1, arg2) As Double
    Intercept = mXlWsFunc.Intercept(arg1, arg2)
End Function

Public Function IntRate(arg1, arg2, arg3, arg4, Optional arg5) As Double
    IntRate = mXlWsFunc.IntRate(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function Ipmt(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Double, Optional arg5, Optional arg6) As Double
    Ipmt = mXlWsFunc.Ipmt(arg1, arg2, arg3, arg4, arg5, arg6)
End Function

Public Function Irr(arg1, Optional arg2) As Double
    Irr = mXlWsFunc.Irr(arg1, arg2)
End Function

Public Function IsErr(arg1) As Boolean
    IsErr = mXlWsFunc.IsErr(arg1)
End Function

Public Function IsError(arg1) As Boolean
    IsError = mXlWsFunc.IsError(arg1)
End Function

Public Function IsEven(arg1) As Boolean
    IsEven = mXlWsFunc.IsEven(arg1)
End Function

Public Function IsFormula(arg1 As Object) As Boolean
    IsFormula = mXlWsFunc.IsFormula(arg1)
End Function

Public Function IsLogical(arg1) As Boolean
    IsLogical = mXlWsFunc.IsLogical(arg1)
End Function

Public Function IsNA(arg1) As Boolean
    IsNA = mXlWsFunc.IsNA(arg1)
End Function

Public Function IsNonText(arg1) As Boolean
    IsNonText = mXlWsFunc.IsNonText(arg1)
End Function

Public Function IsNumber(arg1) As Boolean
    IsNumber = mXlWsFunc.IsNumber(arg1)
End Function

Public Function ISO_Ceiling(arg1 As Double, Optional arg2) As Double
    ISO_Ceiling = mXlWsFunc.ISO_Ceiling(arg1, arg2)
End Function

Public Function IsOdd(arg1) As Boolean
    IsOdd = mXlWsFunc.IsOdd(arg1)
End Function

Public Function IsoWeekNum(arg1 As Double, Optional arg2) As Double
    IsoWeekNum = mXlWsFunc.IsoWeekNum(arg1, arg2)
End Function

Public Function Ispmt(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Double) As Double
    Ispmt = mXlWsFunc.Ispmt(arg1, arg2, arg3, arg4)
End Function

Public Function IsText(arg1) As Boolean
    IsText = mXlWsFunc.IsText(arg1)
End Function

Public Function Kurt(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Kurt = mXlWsFunc.Kurt(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Large(arg1, arg2 As Double) As Double
    Large = mXlWsFunc.Large(arg1, arg2)
End Function

Public Function Lcm(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Lcm = mXlWsFunc.Lcm(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function LinEst(arg1, Optional arg2, Optional arg3, Optional arg4)
    LinEst = mXlWsFunc.LinEst(arg1, arg2, arg3, arg4)
End Function

Public Function Ln(arg1 As Double) As Double
    Ln = mXlWsFunc.Ln(arg1)
End Function

Public Function Log(arg1 As Double, Optional arg2) As Double
    Log = mXlWsFunc.Log(arg1, arg2)
End Function

Public Function Log10(arg1 As Double) As Double
    Log10 = mXlWsFunc.Log10(arg1)
End Function

Public Function LogEst(arg1, Optional arg2, Optional arg3, Optional arg4)
    LogEst = mXlWsFunc.LogEst(arg1, arg2, arg3, arg4)
End Function

Public Function LogInv(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    LogInv = mXlWsFunc.LogInv(arg1, arg2, arg3)
End Function

Public Function LogNorm_Dist(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Boolean) As Double
    LogNorm_Dist = mXlWsFunc.LogNorm_Dist(arg1, arg2, arg3, arg4)
End Function

Public Function LogNorm_Inv(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    LogNorm_Inv = mXlWsFunc.LogNorm_Inv(arg1, arg2, arg3)
End Function

Public Function LogNormDist(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    LogNormDist = mXlWsFunc.LogNormDist(arg1, arg2, arg3)
End Function

Public Function Lookup(arg1, arg2, Optional arg3)
    Lookup = mXlWsFunc.Lookup(arg1, arg2, arg3)
End Function

Public Function Match(arg1, arg2, Optional arg3) As Double
    Match = mXlWsFunc.Match(arg1, arg2, arg3)
End Function

Public Function Max(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Max = mXlWsFunc.Max(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function MaxIfs(arg1 As Object, arg2 As Object, arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29) As Double
    MaxIfs = mXlWsFunc.MaxIfs(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29)
End Function

Public Function MDeterm(arg1) As Double
    MDeterm = mXlWsFunc.MDeterm(arg1)
End Function

Public Function MDuration(arg1, arg2, arg3, arg4, arg5, Optional arg6) As Double
    MDuration = mXlWsFunc.MDuration(arg1, arg2, arg3, arg4, arg5, arg6)
End Function

Public Function Median(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Median = mXlWsFunc.Median(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Min(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Min = mXlWsFunc.Min(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function MinIfs(arg1 As Object, arg2 As Object, arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29) As Double
    MinIfs = mXlWsFunc.MinIfs(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29)
End Function

Public Function MInverse(arg1)
    MInverse = mXlWsFunc.MInverse(arg1)
End Function

Public Function MIrr(arg1, arg2 As Double, arg3 As Double) As Double
    MIrr = mXlWsFunc.MIrr(arg1, arg2, arg3)
End Function

Public Function MMult(arg1, arg2)
    MMult = mXlWsFunc.MMult(arg1, arg2)
End Function

Public Function mode(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    mode = mXlWsFunc.mode(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Mode_Mult(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30)
    Mode_Mult = mXlWsFunc.Mode_Mult(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Mode_Sngl(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Mode_Sngl = mXlWsFunc.Mode_Sngl(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function MRound(arg1, arg2) As Double
    MRound = mXlWsFunc.MRound(arg1, arg2)
End Function

Public Function MultiNomial(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    MultiNomial = mXlWsFunc.MultiNomial(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Munit(arg1 As Double)
    Munit = mXlWsFunc.Munit(arg1)
End Function

Public Function NegBinom_Dist(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Boolean) As Double
    NegBinom_Dist = mXlWsFunc.NegBinom_Dist(arg1, arg2, arg3, arg4)
End Function

Public Function NegBinomDist(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    NegBinomDist = mXlWsFunc.NegBinomDist(arg1, arg2, arg3)
End Function

Public Function NetworkDays(arg1, arg2, Optional arg3) As Double
    NetworkDays = mXlWsFunc.NetworkDays(arg1, arg2, arg3)
End Function

Public Function NetworkDays_Intl(arg1, arg2, Optional arg3, Optional arg4) As Double
    NetworkDays_Intl = mXlWsFunc.NetworkDays_Intl(arg1, arg2, arg3, arg4)
End Function

Public Function Nominal(arg1, arg2) As Double
    Nominal = mXlWsFunc.Nominal(arg1, arg2)
End Function

Public Function Norm_Dist(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Boolean) As Double
    Norm_Dist = mXlWsFunc.Norm_Dist(arg1, arg2, arg3, arg4)
End Function

Public Function Norm_Inv(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    Norm_Inv = mXlWsFunc.Norm_Inv(arg1, arg2, arg3)
End Function

Public Function Norm_S_Dist(arg1 As Double, arg2 As Boolean) As Double
    Norm_S_Dist = mXlWsFunc.Norm_S_Dist(arg1, arg2)
End Function

Public Function Norm_S_Inv(arg1 As Double) As Double
    Norm_S_Inv = mXlWsFunc.Norm_S_Inv(arg1)
End Function

Public Function NormDist(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Boolean) As Double
    NormDist = mXlWsFunc.NormDist(arg1, arg2, arg3, arg4)
End Function

Public Function NormInv(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    NormInv = mXlWsFunc.NormInv(arg1, arg2, arg3)
End Function

Public Function NormSDist(arg1 As Double) As Double
    NormSDist = mXlWsFunc.NormSDist(arg1)
End Function

Public Function NormSInv(arg1 As Double) As Double
    NormSInv = mXlWsFunc.NormSInv(arg1)
End Function

Public Function NPer(arg1 As Double, arg2 As Double, arg3 As Double, Optional arg4, Optional arg5) As Double
    NPer = mXlWsFunc.NPer(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function Npv(arg1 As Double, arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Npv = mXlWsFunc.Npv(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function NumberValue(arg1 As String, arg2 As String, arg3 As String) As Double
    NumberValue = mXlWsFunc.NumberValue(arg1, arg2, arg3)
End Function

Public Function Oct2Bin(arg1, Optional arg2) As String
    Oct2Bin = mXlWsFunc.Oct2Bin(arg1, arg2)
End Function

Public Function Oct2Dec(arg1) As String
    Oct2Dec = mXlWsFunc.Oct2Dec(arg1)
End Function

Public Function Oct2Hex(arg1, Optional arg2) As String
    Oct2Hex = mXlWsFunc.Oct2Hex(arg1, arg2)
End Function

Public Function Odd(arg1 As Double) As Double
    Odd = mXlWsFunc.Odd(arg1)
End Function

Public Function OddFPrice(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, Optional arg9) As Double
    OddFPrice = mXlWsFunc.OddFPrice(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9)
End Function

Public Function OddFYield(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, Optional arg9) As Double
    OddFYield = mXlWsFunc.OddFYield(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9)
End Function

Public Function OddLPrice(arg1, arg2, arg3, arg4, arg5, arg6, arg7, Optional arg8) As Double
    OddLPrice = mXlWsFunc.OddLPrice(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8)
End Function

Public Function OddLYield(arg1, arg2, arg3, arg4, arg5, arg6, arg7, Optional arg8) As Double
    OddLYield = mXlWsFunc.OddLYield(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8)
End Function

Public Function Or_(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Boolean
    Or_ = mXlWsFunc.Or(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function PDuration(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    PDuration = mXlWsFunc.PDuration(arg1, arg2, arg3)
End Function

Public Function Pearson(arg1, arg2) As Double
    Pearson = mXlWsFunc.Pearson(arg1, arg2)
End Function

Public Function Percentile(arg1, arg2 As Double) As Double
    Percentile = mXlWsFunc.Percentile(arg1, arg2)
End Function

Public Function Percentile_Exc(arg1, arg2 As Double) As Double
    Percentile_Exc = mXlWsFunc.Percentile_Exc(arg1, arg2)
End Function

Public Function Percentile_Inc(arg1, arg2 As Double) As Double
    Percentile_Inc = mXlWsFunc.Percentile_Inc(arg1, arg2)
End Function

Public Function PercentRank(arg1, arg2 As Double, Optional arg3) As Double
    PercentRank = mXlWsFunc.PercentRank(arg1, arg2, arg3)
End Function

Public Function PercentRank_Exc(arg1, arg2 As Double, Optional arg3) As Double
    PercentRank_Exc = mXlWsFunc.PercentRank_Exc(arg1, arg2, arg3)
End Function

Public Function PercentRank_Inc(arg1, arg2 As Double, Optional arg3) As Double
    PercentRank_Inc = mXlWsFunc.PercentRank_Inc(arg1, arg2, arg3)
End Function

Public Function Permut(arg1 As Double, arg2 As Double) As Double
    Permut = mXlWsFunc.Permut(arg1, arg2)
End Function

Public Function Permutationa(arg1 As Double, arg2 As Double) As Double
    Permutationa = mXlWsFunc.Permutationa(arg1, arg2)
End Function

Public Function Phi(arg1 As Double) As Double
    Phi = mXlWsFunc.Phi(arg1)
End Function

Public Function Phonetic(arg1 As Object) As String
    Phonetic = mXlWsFunc.Phonetic(arg1)
End Function

Public Function Pi() As Double
    Pi = mXlWsFunc.Pi()
End Function

Public Function Pmt(arg1 As Double, arg2 As Double, arg3 As Double, Optional arg4, Optional arg5) As Double
    Pmt = mXlWsFunc.Pmt(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function Poisson(arg1 As Double, arg2 As Double, arg3 As Boolean) As Double
    Poisson = mXlWsFunc.Poisson(arg1, arg2, arg3)
End Function

Public Function Poisson_Dist(arg1 As Double, arg2 As Double, arg3 As Boolean) As Double
    Poisson_Dist = mXlWsFunc.Poisson_Dist(arg1, arg2, arg3)
End Function

Public Function Power(arg1 As Double, arg2 As Double) As Double
    Power = mXlWsFunc.Power(arg1, arg2)
End Function

Public Function Ppmt(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Double, Optional arg5, Optional arg6) As Double
    Ppmt = mXlWsFunc.Ppmt(arg1, arg2, arg3, arg4, arg5, arg6)
End Function

Public Function Price(arg1, arg2, arg3, arg4, arg5, arg6, Optional arg7) As Double
    Price = mXlWsFunc.Price(arg1, arg2, arg3, arg4, arg5, arg6, arg7)
End Function

Public Function PriceDisc(arg1, arg2, arg3, arg4, Optional arg5) As Double
    PriceDisc = mXlWsFunc.PriceDisc(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function PriceMat(arg1, arg2, arg3, arg4, arg5, Optional arg6) As Double
    PriceMat = mXlWsFunc.PriceMat(arg1, arg2, arg3, arg4, arg5, arg6)
End Function

Public Function Prob(arg1, arg2, arg3 As Double, Optional arg4) As Double
    Prob = mXlWsFunc.Prob(arg1, arg2, arg3, arg4)
End Function

Public Function Product(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Product = mXlWsFunc.Product(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Proper(arg1 As String) As String
    Proper = mXlWsFunc.Proper(arg1)
End Function

Public Function Pv(arg1 As Double, arg2 As Double, arg3 As Double, Optional arg4, Optional arg5) As Double
    Pv = mXlWsFunc.Pv(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function Quartile(arg1, arg2 As Double) As Double
    Quartile = mXlWsFunc.Quartile(arg1, arg2)
End Function

Public Function Quartile_Exc(arg1, arg2 As Double) As Double
    Quartile_Exc = mXlWsFunc.Quartile_Exc(arg1, arg2)
End Function

Public Function Quartile_Inc(arg1, arg2 As Double) As Double
    Quartile_Inc = mXlWsFunc.Quartile_Inc(arg1, arg2)
End Function

Public Function Quotient(arg1, arg2) As Double
    Quotient = mXlWsFunc.Quotient(arg1, arg2)
End Function

Public Function Radians(arg1 As Double) As Double
    Radians = mXlWsFunc.Radians(arg1)
End Function

Public Function RandArray(Optional arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5)
    RandArray = mXlWsFunc.RandArray(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function RandBetween(arg1, arg2) As Double
    RandBetween = mXlWsFunc.RandBetween(arg1, arg2)
End Function

Public Function Rank(arg1 As Double, arg2 As Object, Optional arg3) As Double
    Rank = mXlWsFunc.Rank(arg1, arg2, arg3)
End Function

Public Function Rank_Avg(arg1 As Double, arg2 As Object, Optional arg3) As Double
    Rank_Avg = mXlWsFunc.Rank_Avg(arg1, arg2, arg3)
End Function

Public Function Rank_Eq(arg1 As Double, arg2 As Object, Optional arg3) As Double
    Rank_Eq = mXlWsFunc.Rank_Eq(arg1, arg2, arg3)
End Function

Public Function Rate(arg1 As Double, arg2 As Double, arg3 As Double, Optional arg4, Optional arg5, Optional arg6) As Double
    Rate = mXlWsFunc.Rate(arg1, arg2, arg3, arg4, arg5, arg6)
End Function

Public Function Received(arg1, arg2, arg3, arg4, Optional arg5) As Double
    Received = mXlWsFunc.Received(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function Replace(arg1 As String, arg2 As Double, arg3 As Double, arg4 As String) As String
    Replace = mXlWsFunc.Replace(arg1, arg2, arg3, arg4)
End Function

Public Function ReplaceB(arg1 As String, arg2 As Double, arg3 As Double, arg4 As String) As String
    ReplaceB = mXlWsFunc.ReplaceB(arg1, arg2, arg3, arg4)
End Function

Public Function Rept(arg1 As String, arg2 As Double) As String
    Rept = mXlWsFunc.Rept(arg1, arg2)
End Function

Public Function Roman(arg1 As Double, Optional arg2) As String
    Roman = mXlWsFunc.Roman(arg1, arg2)
End Function

Public Function Round(arg1 As Double, arg2 As Double) As Double
    Round = mXlWsFunc.Round(arg1, arg2)
End Function

Public Function RoundDown(arg1 As Double, arg2 As Double) As Double
    RoundDown = mXlWsFunc.RoundDown(arg1, arg2)
End Function

Public Function RoundUp(arg1 As Double, arg2 As Double) As Double
    RoundUp = mXlWsFunc.RoundUp(arg1, arg2)
End Function

Public Function Rri(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    Rri = mXlWsFunc.Rri(arg1, arg2, arg3)
End Function

Public Function RSq(arg1, arg2) As Double
    RSq = mXlWsFunc.RSq(arg1, arg2)
End Function

Public Function RTD(progID, Server, topic1, Optional topic2, Optional topic3, Optional topic4, Optional topic5, Optional topic6, Optional topic7, Optional topic8, Optional topic9, Optional topic10, Optional topic11, Optional topic12, Optional topic13, Optional topic14, Optional topic15, Optional topic16, Optional topic17, Optional topic18, Optional topic19, Optional topic20, Optional topic21, Optional topic22, Optional topic23, Optional topic24, Optional topic25, Optional topic26, Optional topic27, Optional topic28)
    RTD = mXlWsFunc.RTD(progID, Server, topic1, topic2, topic3, topic4, topic5, topic6, topic7, topic8, topic9, topic10, topic11, topic12, topic13, topic14, topic15, topic16, topic17, topic18, topic19, topic20, topic21, topic22, topic23, topic24, topic25, topic26, topic27, topic28)
End Function

Public Function Search(arg1 As String, arg2 As String, Optional arg3) As Double
    Search = mXlWsFunc.Search(arg1, arg2, arg3)
End Function

Public Function SearchB(arg1 As String, arg2 As String, Optional arg3) As Double
    SearchB = mXlWsFunc.SearchB(arg1, arg2, arg3)
End Function

Public Function Sec(arg1 As Double) As Double
    Sec = mXlWsFunc.Sec(arg1)
End Function

Public Function Sech(arg1 As Double) As Double
    Sech = mXlWsFunc.Sech(arg1)
End Function

Public Function Sequence(arg1, Optional arg2, Optional arg3, Optional arg4)
    Sequence = mXlWsFunc.Sequence(arg1, arg2, arg3, arg4)
End Function

Public Function SeriesSum(arg1, arg2, arg3, arg4) As Double
    SeriesSum = mXlWsFunc.SeriesSum(arg1, arg2, arg3, arg4)
End Function

Public Function Single_(arg1)
    Single_ = mXlWsFunc.Single(arg1)
End Function

Public Function Sinh(arg1 As Double) As Double
    Sinh = mXlWsFunc.Sinh(arg1)
End Function

Public Function Skew(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Skew = mXlWsFunc.Skew(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Skew_p(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Skew_p = mXlWsFunc.Skew_p(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Sln(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    Sln = mXlWsFunc.Sln(arg1, arg2, arg3)
End Function

Public Function Slope(arg1, arg2) As Double
    Slope = mXlWsFunc.Slope(arg1, arg2)
End Function

Public Function Small(arg1, arg2 As Double) As Double
    Small = mXlWsFunc.Small(arg1, arg2)
End Function

Public Function Sort(arg1, Optional arg2, Optional arg3, Optional arg4)
    Sort = mXlWsFunc.Sort(arg1, arg2, arg3, arg4)
End Function

Public Function SortBy(arg1, arg2, arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30)
    SortBy = mXlWsFunc.SortBy(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function SqrtPi(arg1) As Double
    SqrtPi = mXlWsFunc.SqrtPi(arg1)
End Function

Public Function Standardize(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    Standardize = mXlWsFunc.Standardize(arg1, arg2, arg3)
End Function

Public Function StDev(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    StDev = mXlWsFunc.StDev(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function StDev_P(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    StDev_P = mXlWsFunc.StDev_P(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function StDev_S(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    StDev_S = mXlWsFunc.StDev_S(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function StDevP(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    StDevP = mXlWsFunc.StDevP(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function StEyx(arg1, arg2) As Double
    StEyx = mXlWsFunc.StEyx(arg1, arg2)
End Function

Public Function StockHistory(arg1, arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29)
    StockHistory = mXlWsFunc.StockHistory(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29)
End Function

Public Function Substitute(arg1 As String, arg2 As String, arg3 As String, Optional arg4) As String
    Substitute = mXlWsFunc.Substitute(arg1, arg2, arg3, arg4)
End Function

Public Function Subtotal(arg1 As Double, arg2 As Object, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Subtotal = mXlWsFunc.Subtotal(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Sum(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Sum = mXlWsFunc.Sum(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function SumIf(arg1 As Object, arg2, Optional arg3) As Double
    SumIf = mXlWsFunc.SumIf(arg1, arg2, arg3)
End Function

Public Function SumIfs(arg1 As Object, arg2 As Object, arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29) As Double
    SumIfs = mXlWsFunc.SumIfs(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29)
End Function

Public Function SumProduct(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    SumProduct = mXlWsFunc.SumProduct(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function SumSq(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    SumSq = mXlWsFunc.SumSq(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function SumX2MY2(arg1, arg2) As Double
    SumX2MY2 = mXlWsFunc.SumX2MY2(arg1, arg2)
End Function

Public Function SumX2PY2(arg1, arg2) As Double
    SumX2PY2 = mXlWsFunc.SumX2PY2(arg1, arg2)
End Function

Public Function SumXMY2(arg1, arg2) As Double
    SumXMY2 = mXlWsFunc.SumXMY2(arg1, arg2)
End Function

Public Function Syd(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Double) As Double
    Syd = mXlWsFunc.Syd(arg1, arg2, arg3, arg4)
End Function

Public Function T_Dist(arg1 As Double, arg2 As Double, arg3 As Boolean) As Double
    T_Dist = mXlWsFunc.T_Dist(arg1, arg2, arg3)
End Function

Public Function T_Dist_2T(arg1 As Double, arg2 As Double) As Double
    T_Dist_2T = mXlWsFunc.T_Dist_2T(arg1, arg2)
End Function

Public Function T_Dist_RT(arg1 As Double, arg2 As Double) As Double
    T_Dist_RT = mXlWsFunc.T_Dist_RT(arg1, arg2)
End Function

Public Function T_Inv(arg1 As Double, arg2 As Double) As Double
    T_Inv = mXlWsFunc.T_Inv(arg1, arg2)
End Function

Public Function T_Inv_2T(arg1 As Double, arg2 As Double) As Double
    T_Inv_2T = mXlWsFunc.T_Inv_2T(arg1, arg2)
End Function

Public Function T_Test(arg1, arg2, arg3 As Double, arg4 As Double) As Double
    T_Test = mXlWsFunc.T_Test(arg1, arg2, arg3, arg4)
End Function

Public Function Tanh(arg1 As Double) As Double
    Tanh = mXlWsFunc.Tanh(arg1)
End Function

Public Function TBillEq(arg1, arg2, Optional arg3) As Double
    TBillEq = mXlWsFunc.TBillEq(arg1, arg2, arg3)
End Function

Public Function TBillPrice(arg1, arg2, Optional arg3) As Double
    TBillPrice = mXlWsFunc.TBillPrice(arg1, arg2, arg3)
End Function

Public Function TBillYield(arg1, arg2, Optional arg3) As Double
    TBillYield = mXlWsFunc.TBillYield(arg1, arg2, arg3)
End Function

Public Function TDist(arg1 As Double, arg2 As Double, arg3 As Double) As Double
    TDist = mXlWsFunc.TDist(arg1, arg2, arg3)
End Function

Public Function text(arg1, arg2 As String) As String
    text = mXlWsFunc.text(arg1, arg2)
End Function

Public Function TextJoin(arg1 As String, arg2 As Boolean, arg3 As String, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29) As String
    TextJoin = mXlWsFunc.TextJoin(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29)
End Function

Public Function TInv(arg1 As Double, arg2 As Double) As Double
    TInv = mXlWsFunc.TInv(arg1, arg2)
End Function

Public Function Transpose(arg1)
    Transpose = mXlWsFunc.Transpose(arg1)
End Function

Public Function Trend(arg1, Optional arg2, Optional arg3, Optional arg4)
    Trend = mXlWsFunc.Trend(arg1, arg2, arg3, arg4)
End Function

Public Function Trim(arg1 As String) As String
    Trim = mXlWsFunc.Trim(arg1)
End Function

Public Function TrimMean(arg1, arg2 As Double) As Double
    TrimMean = mXlWsFunc.TrimMean(arg1, arg2)
End Function

Public Function TTest(arg1, arg2, arg3 As Double, arg4 As Double) As Double
    TTest = mXlWsFunc.TTest(arg1, arg2, arg3, arg4)
End Function

Public Function Unichar(arg1 As Double) As String
    Unichar = mXlWsFunc.Unichar(arg1)
End Function

Public Function unicode(arg1 As String) As Double
    unicode = mXlWsFunc.unicode(arg1)
End Function

Public Function USDollar(arg1 As Double, arg2 As Double) As String
    USDollar = mXlWsFunc.USDollar(arg1, arg2)
End Function

Public Function var(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    var = mXlWsFunc.var(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Var_P(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Var_P = mXlWsFunc.Var_P(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Var_S(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    Var_S = mXlWsFunc.Var_S(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function VarP(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Double
    VarP = mXlWsFunc.VarP(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function Vdb(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Double, arg5 As Double, Optional arg6, Optional arg7) As Double
    Vdb = mXlWsFunc.Vdb(arg1, arg2, arg3, arg4, arg5, arg6, arg7)
End Function

Public Function VLookup(arg1, arg2, arg3, Optional arg4)
    VLookup = mXlWsFunc.VLookup(arg1, arg2, arg3, arg4)
End Function

Public Function WebService(arg1 As String)
    WebService = mXlWsFunc.WebService(arg1)
End Function

Public Function Weekday(arg1, Optional arg2) As Double
    Weekday = mXlWsFunc.Weekday(arg1, arg2)
End Function

Public Function WeekNum(arg1, Optional arg2) As Double
    WeekNum = mXlWsFunc.WeekNum(arg1, arg2)
End Function

Public Function Weibull(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Boolean) As Double
    Weibull = mXlWsFunc.Weibull(arg1, arg2, arg3, arg4)
End Function

Public Function Weibull_Dist(arg1 As Double, arg2 As Double, arg3 As Double, arg4 As Boolean) As Double
    Weibull_Dist = mXlWsFunc.Weibull_Dist(arg1, arg2, arg3, arg4)
End Function

Public Function WorkDay(arg1, arg2, Optional arg3) As Double
    WorkDay = mXlWsFunc.WorkDay(arg1, arg2, arg3)
End Function

Public Function WorkDay_Intl(arg1, arg2, Optional arg3, Optional arg4) As Double
    WorkDay_Intl = mXlWsFunc.WorkDay_Intl(arg1, arg2, arg3, arg4)
End Function

Public Function Xirr(arg1, arg2, Optional arg3) As Double
    Xirr = mXlWsFunc.Xirr(arg1, arg2, arg3)
End Function

Public Function XLookup(arg1, arg2, arg3, Optional arg4, Optional arg5, Optional arg6)
    XLookup = mXlWsFunc.XLookup(arg1, arg2, arg3, arg4, arg5, arg6)
End Function

Public Function XMatch(arg1, arg2, Optional arg3, Optional arg4) As Double
    XMatch = mXlWsFunc.XMatch(arg1, arg2, arg3, arg4)
End Function

Public Function Xnpv(arg1, arg2) As Double
    Xnpv = mXlWsFunc.Xnpv(arg1, arg2)
End Function

Public Function Xor_(arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10, Optional arg11, Optional arg12, Optional arg13, Optional arg14, Optional arg15, Optional arg16, Optional arg17, Optional arg18, Optional arg19, Optional arg20, Optional arg21, Optional arg22, Optional arg23, Optional arg24, Optional arg25, Optional arg26, Optional arg27, Optional arg28, Optional arg29, Optional arg30) As Boolean
    Xor_ = mXlWsFunc.Xor(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30)
End Function

Public Function YearFrac(arg1, arg2, Optional arg3) As Double
    YearFrac = mXlWsFunc.YearFrac(arg1, arg2, arg3)
End Function

Public Function YieldDisc(arg1, arg2, arg3, arg4, Optional arg5) As Double
    YieldDisc = mXlWsFunc.YieldDisc(arg1, arg2, arg3, arg4, arg5)
End Function

Public Function Z_Test(arg1, arg2 As Double, Optional arg3) As Double
    Z_Test = mXlWsFunc.Z_Test(arg1, arg2, arg3)
End Function

Public Function ZTest(arg1, arg2 As Double, Optional arg3) As Double
    ZTest = mXlWsFunc.ZTest(arg1, arg2, arg3)
End Function

