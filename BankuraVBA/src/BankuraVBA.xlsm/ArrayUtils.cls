VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] 配列操作ユーティリティクラス
'* [詳  細] 配列操作を行うユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'*
'* [依　存] <内部依存>
'*           Base, ValidateUtils
'*          <外部依存>
'*           Ariawase Core, Ext, ArrayEx
'* [参  考] https://qiita.com/nukie_53/items/9a7a1eb07eff50ae1e8b
'*          https://blog.sgnet.co.jp/2017/01/vba.html
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function SafeArrayAllocDescriptor Lib "oleaut32" (ByVal cDims As Long, ByRef ppsaOut() As Any) As Long

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
'* プロパティ定義
'******************************************************************************
Public Property Get EmptyLongArray() As Long()

    Dim apiResult As Long
    apiResult = SafeArrayAllocDescriptor(1, EmptyLongArray)
    If apiResult <> 0 Then Err.Raise 5
End Property

Public Property Get EmptyVariantArray() As Variant()
    Let EmptyVariantArray = VBA.Array()
End Property

Public Property Get EmptyVariant2DArray() As Variant()
    Dim apiResult As Long
    apiResult = SafeArrayAllocDescriptor(2, EmptyVariant2DArray)
    If apiResult <> 0 Then Err.Raise 5
End Property

Public Property Get EmptyStringArray() As String()
    Let EmptyStringArray = VBA.Split(VBA.vbNullString)
End Property

Public Property Get EmptyByteArray() As Byte()
    Let EmptyByteArray = VBA.vbNullString
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
'* [概  要] ToVArrayFlattened
'* [詳  細] Variant型配列内にVariant配列と値が混在する場合に、
'*          Variant配列の値を展開し値だけのVariant型配列に変換する｡
'*
'* @param vArray Variant型配列
'* @return Variant 変換後Variant型配列
'*
'******************************************************************************
Public Function ToVArrayFlattened(ByRef vArray As Variant) As Variant
    Dim arr As ArrayEx: Set arr = New ArrayEx
    
    Call AddAllFromHierarchicalArray(arr, vArray)
    ToVArrayFlattened = arr.ToArray
End Function

'******************************************************************************
'* [概  要] AddAllFromHierarchicalArray
'* [詳  細] Variant型配列内にVariant配列と値が混在する場合、
'*          また配列内の配列にさらに配列が格納される階層構造になっている場合、
'*          ArrayExオブジェクトに全ての要素を格納する（フラット化）｡
'*
'* @param arr ArrayEx
'* @param vArray Variant型配列
'*
'******************************************************************************
Public Sub AddAllFromHierarchicalArray(ByVal arr As ArrayEx, ByRef vArray As Variant)
    If ValidateUtils.IsPrimitive(vArray) Then
        arr.AddVal vArray
    ElseIf IsObject(vArray) Then
        arr.AddObj vArray
    ElseIf VBA.IsArray(vArray) Then
        Dim v
        For Each v In vArray
            Call AddAllFromHierarchicalArray(arr, v)
        Next
    End If
End Sub

'******************************************************************************
'* [概  要] GetDimension
'* [詳  細] 指定されたVariant配列の次元数を返却する｡
'*
'* @param vArray 検査する配列
'* @return Integer 次元数(配列ではない場合-1)
'*
'******************************************************************************
Public Function GetDimension(ByRef vArray As Variant) As Integer
    On Error GoTo ErrorHandler
    GetDimension = Core.ArrRank(vArray)
    Exit Function
ErrorHandler:
    GetDimension = -1
End Function

'******************************************************************************
'* [概  要] GetLength
'* [詳  細] 指定されたVariant配列の要素数を返却する｡
'*
'*
'* @param vArray 検査する配列
'* @param dimmention 次元
'* @return Long 要素数(配列ではない場合-1)
'*
'******************************************************************************
Public Function GetLength(ByRef vArray As Variant, Optional dimmention As Integer = 1) As Long
    On Error GoTo ErrorHandler
    GetLength = Core.ArrLen(vArray, dimmention)
    Exit Function
ErrorHandler:
    GetLength = -1
End Function

'******************************************************************************
'* [概  要] IsRedim
'* [詳  細] 指定された配列が初期化済みか判定する｡
'*
'* @param vArray 検査する配列
'* @return Boolean 判定結果(True:初期化済み)
'*
'******************************************************************************
Public Function IsRedim(ByRef vArray As Variant) As Boolean
    On Error Resume Next
    Err.Clear
    IsRedim = CBool(UBound(vArray))
    IsRedim = (Err.Number = 0)
End Function

'******************************************************************************
'* [概  要] IsEmptyArray
'* [詳  細] 指定された配列が未初期化・空配列か判定する｡
'*
'* @param vArray 検査する配列
'* @return Boolean 判定結果(True:未初期化・空)
'*
'******************************************************************************
Public Function IsEmptyArray(ByRef vArray As Variant) As Boolean
    IsEmptyArray = False
    If IsRedim(vArray) Then
        If UBound(vArray) = -1 Then
            IsEmptyArray = True
        End If
    Else
        IsEmptyArray = True
    End If
End Function

'******************************************************************************
'* [概  要] Is1DArray
'* [詳  細] 指定された配列が1次元配列か判定する｡
'*
'* @param vArray 検査する配列
'* @return Boolean 判定結果(True:1次元配列)
'*
'******************************************************************************
Public Function Is1DArray(ByRef vArray As Variant) As Boolean
    If GetDimension(vArray) = 1 Then
        Is1DArray = True
        Exit Function
    End If
    Is1DArray = False
End Function

'******************************************************************************
'* [概  要] Is2DArray
'* [詳  細] 指定された配列が2次元配列か判定する｡
'*
'* @param vArray 検査する配列
'* @return Boolean 判定結果(True:2次元配列)
'*
'******************************************************************************
Public Function Is2DArray(ByRef vArray As Variant) As Boolean
    If GetDimension(vArray) = 2 Then
        Is2DArray = True
        Exit Function
    End If
    Is2DArray = False
End Function

'******************************************************************************
'* [概  要] IsJagArray
'* [詳  細] 指定された配列がジャグ配列（2次元まで）か判定する｡
'*
'* @param vArray 検査する配列
'* @return Boolean 判定結果(True:ジャグ配列)
'*
'******************************************************************************
Public Function IsJagArray(ByRef vArray As Variant) As Boolean
    If Not VBA.IsArray(vArray) Then Exit Function
    If Me.IsEmptyArray(vArray) Then Exit Function
    If Not Me.Is1DArray(vArray) Then Exit Function

    Dim v
    For Each v In vArray
        If Not Me.Is1DArray(v) Then Exit Function
    Next
    IsJagArray = True
End Function

'******************************************************************************
'* [概  要] ArrayRank
'* [詳  細] 配列の次元を返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr As Variant(Of Array(Of T))
'* @return As Integer
'******************************************************************************
Public Function ArrayRank(ByVal arr As Variant) As Integer
    ArrayRank = Core.ArrRank(arr)
End Function

'******************************************************************************
'* [概  要] ArrayLen
'* [詳  細] 配列の長さ（要素数）を返却する。次元を指定可能。デフォルトは１次元。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr As Variant(Of Array(Of T))
'* @param dimen As Integer
'* @return As Long
'******************************************************************************
Public Function ArrayLen(ByVal arr As Variant, Optional ByVal dimen As Integer = 1) As Long
    ArrayLen = Core.ArrLen(arr, dimen)
End Function

'******************************************************************************
'* [概  要] Equals
'* [詳  細] ２つの一次元配列を比較し、イコールであればTrueを返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr1 As Variant(Of Array(Of T))
'* @param arr2 As Variant(Of Array(Of T))
'* @param swAllowNull As Boolean
'* @return As Variant(Of Boolean Or Null Or Empty)
'******************************************************************************
Public Function Equals( _
    ByVal arr1 As Variant, ByVal arr2 As Variant, _
    Optional ByVal swAllowNull As Boolean = False _
    ) As Variant
    Equals = Core.ArrEquals(arr1, arr2, swAllowNull)
End Function

'******************************************************************************
'* [概  要] Compare
'* [詳  細] ２つの一次元配列（x,y）を比較し、x=yであれば0、x<yであれば-1、x>y
'*          であれば1を返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr1 As Variant(Of Array(Of T))
'* @param arr2 As Variant(Of Array(Of T))
'* @return As Variant(Of Integer Or Null)
'******************************************************************************
Public Function Compare( _
    ByVal arr1 As Variant, ByVal arr2 As Variant, _
    Optional ByVal swAllowNull As Boolean = False _
    ) As Variant
    Compare = Core.ArrCompare(arr1, arr2, swAllowNull)
End Function

'******************************************************************************
'* [概  要] IndexOf
'* [詳  細] 一次元配列から指定した値と一致する要素のインデックスを返却する。
'*          値の捜査開始位置と要素数を任意指定可能。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr As Variant(Of Array(Of T))
'* @param val As Variant(Of T)
'* @param ixStart As Variant(Of Long)
'* @param cnt As Variant(Of Long)
'* @return As Long
'******************************************************************************
Public Function IndexOf( _
    ByVal arr As Variant, ByVal val As Variant, _
    Optional ByVal ixStart As Variant, Optional ByVal cnt As Variant _
    ) As Long
    IndexOf = Core.ArrIndexOf(arr, val, ixStart, cnt)
End Function

'******************************************************************************
'* [概  要] Reverse
'* [詳  細] 一次元配列の要素の並び順を反転させる。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr As Variant(Of Array(Of T))
'******************************************************************************
Public Sub Reverse(ByRef arr As Variant)
    Call Core.ArrRev(arr)
End Sub

'******************************************************************************
'* [概  要] Sort
'* [詳  細] 一次元配列のソートを行う。デフォルトは昇順。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr As Variant(Of Array(Of T))
'* @param orderAsc As Boolean
'******************************************************************************
Public Sub Sort(ByRef arr As Variant, Optional ByVal orderAsc As Boolean = True)
    Call Core.ArrSort(arr, orderAsc)
End Sub

'******************************************************************************
'* [概  要] ArrayUniq
'* [詳  細] 一次元配列の重複要素を排除する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr As Variant(Of Array(Of T))
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function Uniq(ByVal arr As Variant) As Variant
    Uniq = Core.ArrUniq(arr)
End Function

'******************************************************************************
'* [概  要] Concat
'* [詳  細] 一次元配列と一次元配列を合体（マージ）して返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr1 As Variant(Of Array(Of T))
'* @param arr2 As Variant(Of Array(Of T))
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function Concat(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    Concat = Core.ArrConcat(arr1, arr2)
End Function

'******************************************************************************
'* [概  要] ArraySlice
'* [詳  細] 一次元配列をスライス（一部を配列として切り出）して返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr As Variant(Of Array(Of T))
'* @param ixStart As Variant(Of Long)
'* @param ixEnd As Variant(Of Long)
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function Slice( _
    ByVal arr As Variant, _
    Optional ByVal ixStart As Variant, Optional ByVal ixEnd As Variant _
    ) As Variant
    Slice = Core.ArrSlice(arr, ixStart, ixEnd)
End Function

'******************************************************************************
'* [概  要] ArrayFlatten
'* [詳  細] ジャグ配列のフラット化。ジャグ配列を一次元配列に変換して返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param jagArray As Variant(Of Array(Of Array(Of T)))
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function Flatten(ByVal jagArr As Variant) As Variant
    Flatten = Core.ArrFlatten(jagArr)
End Function

'******************************************************************************
'* [概  要] ArrayToCollection
'* [詳  細] 一次元配列をCollectionオブジェクトに変換する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr As Variant(Of Array(Of T))
'* @return As Collection(Of T)
'******************************************************************************
Public Function ArrayToCollection(ByVal arr As Variant) As Collection
    Set ArrayToCollection = Core.ArrToClct(arr)
End Function

'******************************************************************************
'* [概  要] CollectionToArray
'* [詳  細] Collectionオブジェクトを一次元配列に変換する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param clct As Collection(Of T)
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function CollectionToArray(ByVal clct As Collection) As Variant
    CollectionToArray = Core.ClctToArr(clct)
End Function

'******************************************************************************
'* [概  要] JagArrayToArray2D
'* [詳  細] ジャグ配列を二次元配列に変換して返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param jagArr As Variant(Of Array(Of Array(Of T))
'* @return As Variant(Of Array(Of T, T))
'******************************************************************************
Public Function JagArrayToArray2D(ByVal jagArr As Variant) As Variant
    JagArrayToArray2D = Core.JagArrToArr2D(jagArr)
End Function

'******************************************************************************
'* [概  要] Array2DToJagArray
'* [詳  細] 二次元配列をジャグ配列に変換して返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr2D As Variant(Of Array(Of T, T))
'* @return As Variant(Of Array(Of Array(Of T))
'******************************************************************************
Public Function Array2DToJagArray(ByVal arr2d As Variant) As Variant
    Array2DToJagArray = Core.Arr2DToJagArr(arr2d)
End Function

'******************************************************************************
'* [概  要] AssocArrayToDictionary
'* [詳  細] 連想配列をDictionaryオブジェクトに変換する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param aarr Variant(Of Array(Of Tupple))
'* @return Object(Of Dictionary)
'******************************************************************************
Public Function AssocArrayToDictionary(ByVal aarr As Variant) As Object
    Set AssocArrayToDictionary = Ext.AssocArrToDict(aarr)
End Function

'******************************************************************************
'* [概  要] DictionaryToAssocArray
'* [詳  細] Dictionaryオブジェクトを連想配列に変換する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param dict
'* @return Variant(Of Array(Of Tupple))
'******************************************************************************
Public Function DictionaryToAssocArray(ByVal dict As Object) As Variant
    DictionaryToAssocArray = Ext.DictToAssocArr(dict)
End Function

'******************************************************************************
'* [概  要] EnumeratorToArray
'* [詳  細] 繰り返し処理（For Each）可能なオブジェクトを一次元配列に変換する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param enumr As Enumerator(Of T)
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function EnumeratorToArray(ByVal enumr As Object) As Variant
    EnumeratorToArray = Ext.EnumeratorToArr(enumr)
End Function

'******************************************************************************
'* [概  要] Range
'* [詳  細] 指定した範囲の数値の配列を作成する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fromVal As Variant(Of T)
'* @param toVal As Variant(Of T)
'* @param stepVal As Variant(Of T)
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function Range( _
    ByVal fromVal As Variant, ByVal toVal As Variant, Optional ByVal stepVal As Variant = 1 _
    ) As Variant
    Range = Ext.ArrRange(fromVal, toVal, stepVal)
End Function

'******************************************************************************
'* [概  要] Map
'* [詳  細] 配列の全要素に対して指定したファンクションの処理を適用する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of T, U)
'* @param arr As Variant(Of Array(Of T))
'* @return As Variant(Of Array(Of U))
'******************************************************************************
Public Function Map(ByVal fun As Func, ByVal arr As Variant) As Variant
    Map = Ext.ArrMap(fun, arr)
End Function

'******************************************************************************
'* [概  要] Zip
'* [詳  細] ２つの配列に対して指定したファンクションの処理を適用し
'*          １つの配列を返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of T, U, R)
'* @param arr1 As Variant(Of Array(Of T))
'* @param arr2 As Variant(Of Array(Of U))
'* @return As Variant(Of Array(Of R))
'******************************************************************************
Public Function Zip( _
    ByVal fun As Func, ByVal arr1 As Variant, ByVal arr2 As Variant _
    ) As Variant
    Zip = Ext.ArrZip(fun, arr1, arr2)
End Function

'******************************************************************************
'* [概  要] Filter
'* [詳  細] 配列に対して指定したフィルタ処理用のファンクションを適用し、
'*          条件に一致する要素のみの配列を返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of T, Boolean)
'* @param arr As Variant(Of Array(Of T))
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function Filter(ByVal fun As Func, ByVal arr As Variant) As Variant
    Filter = Ext.ArrFilter(fun, arr)
End Function

'******************************************************************************
'* [概  要] GroupBy
'* [詳  細] 配列に対して指定したグルーピング処理用のファンクションを適用し、
'*          結果を返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of T, K)
'* @param arr As Variant(Of Array(Of T))
'* @return As Variant(Of Array(Of Tuple`2(Of K, T)))
'******************************************************************************
Public Function GroupBy(ByVal fun As Func, ByVal arr As Variant) As Variant
    GroupBy = Ext.ArrGroupBy(fun, arr)
End Function

'******************************************************************************
'* [概  要] Fold
'* [詳  細] 畳み込み関数。配列内の要素に対して指定した演算用ファンクション
'*          を適用し、その結果と次の要素にまたファンクションを適用する
'*          といった処理を繰り返し最終的な結果を返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of U, T, U)
'* @param arr As Variant(Of Array(Of T))
'* @param seedv As Variant(Of U)
'* @return As Variant(Of U)
'******************************************************************************
Public Function Fold( _
    ByVal fun As Func, ByVal arr As Variant, Optional ByVal seedv As Variant _
    ) As Variant
    Fold = Ext.ArrFold(fun, arr, seedv)
End Function

'******************************************************************************
'* [概  要] Scan
'* [詳  細] 配列の各要素に対して繰り返し指定した演算用ファンクションを適用し、
'*          累積した計算結果のリストを返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of U, T, U)
'* @param arr As Variant(Of Array(Of T))
'* @param seedv As Variant(Of U)
'* @return As Variant(Of Array(Of U))
'******************************************************************************
Public Function Scan( _
    ByVal fun As Func, ByVal arr As Variant, Optional ByVal seedv As Variant _
    ) As Variant
    Scan = Ext.ArrScan(fun, arr, seedv)
End Function

'******************************************************************************
'* [概  要] Unfold
'* [詳  細] Foldの逆。初期値から指定した演算用ファンクションを適用しリストを
'*          生成する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func
'* @param seedv As Variant(Of T)
'* @return As Variant(Of Array(Of U))
'******************************************************************************
Public Function Unfold(ByVal fun As Func, ByVal seedv As Variant) As Variant
    Unfold = Ext.ArrUnfold(fun, seedv)
End Function

'******************************************************************************
'* [概  要] RedimPreserve2D
'* [詳  細] 2次元配列の1次元目の長さを変更する。
'*          添え字の開始が1始まりに変更される。
'*
'* @param orgArray 2次元配列
'* @param legthTo 1次元目の長さ
'* @return Variant 変更後2次元配列
'******************************************************************************
Public Function RedimPreserve2D(ByVal orgArray, ByVal lengthTo) As Variant
    Dim transedArray()
    transedArray = WorksheetFunction.Transpose(orgArray)
    ReDim Preserve transedArray(1 To UBound(transedArray, 1), 1 To lengthTo)
    RedimPreserve2D = WorksheetFunction.Transpose(transedArray)
End Function

'******************************************************************************
'* [概  要] SetDefaultTo2DArrayRow
'* [詳  細] 2次元配列の指定行に対して、全てのカラム列にデフォルト値を設定する。
'*
'* @param v2DArr 2次元配列
'* @param rowNum 指定行
'* @param defaultVal デフォルト値
'* @return Variant 変更後2次元配列
'******************************************************************************
Public Function SetDefaultTo2DArrayRow(ByRef v2DArr As Variant, ByVal rownum As Long, ByVal defaultVal As Variant) As Variant
    Dim i As Long
    For i = LBound(v2DArr, 2) To UBound(v2DArr, 2)
        v2DArr(rownum, i) = defaultVal
    Next
    SetDefaultTo2DArrayRow = v2DArr
End Function

'******************************************************************************
'* [概  要] Copy1DArrayTo2DArrayRow
'* [詳  細] 2次元配列の指定行に対して、1次元配列の値を設定する。
'*
'* @param v2DArr 2次元配列
'* @param rowNum 指定行
'* @param v1DArr 1次元配列
'* @return Variant 変更後2次元配列
'******************************************************************************
Public Function Copy1DArrayTo2DArrayRow(ByRef v2DArr As Variant, ByVal rownum As Long, ByRef v1DArr As Variant) As Variant
    Dim i As Long
    For i = LBound(v2DArr, 2) To UBound(v2DArr, 2)
        v2DArr(rownum, i) = v1DArr(i)
    Next
    Copy1DArrayTo2DArrayRow = v2DArr
End Function

'******************************************************************************
'* [概  要] Create2DArrayFromPreExisting
'* [詳  細] 既に存在する2次元配列と同じ要素数の新しい2次元配列を作成する。
'*
'* @param v2DArr 2次元配列
'* @return Variant 新規2次元配列
'******************************************************************************
Public Function Create2DArrayFromPreExisting(ByRef v2DArr As Variant) As Variant
    Create2DArrayFromPreExisting = Create2DArray(LBound(v2DArr, 1), UBound(v2DArr, 1), LBound(v2DArr, 2), UBound(v2DArr, 2))
End Function

'******************************************************************************
'* [概  要] Create2DArray
'* [詳  細] 新しい2次元配列を作成する。
'*
'* @param lower1 1次元開始インデックス
'* @param upper1 1次元終了インデックス
'* @param lower2 2次元開始インデックス
'* @param upper2 2次元終了インデックス
'* @return Variant 2次元配列
'******************************************************************************
Public Function Create2DArray(ByVal lower1 As Long, ByVal upper1 As Long, ByVal lower2 As Long, ByVal upper2 As Long) As Variant
    Dim vNew2DArr As Variant
    ReDim vNew2DArr(lower1 To upper1, lower2 To upper2) As Variant
    Create2DArray = vNew2DArr
End Function

'******************************************************************************
'* [概  要] Create2DArrayEasy
'* [詳  細] 新しい2次元配列を作成する。開始インデックスは0固定
'*
'* @param upper1 1次元終了インデックス
'* @param upper2 2次元終了インデックス
'* @return Variant 2次元配列
'******************************************************************************
Public Function Create2DArrayEasy(ByVal upper1 As Long, ByVal upper2 As Long) As Variant
    Dim vNew2DArr As Variant
    ReDim vNew2DArr(0 To upper1, 0 To upper2) As Variant
    Create2DArrayEasy = vNew2DArr
End Function

'******************************************************************************
'* [概  要] Create2DArrayWithValue
'* [詳  細] 新しい2次元配列を作成する。以下のように値を設定する。
'*            Create2DArrayWithValue(Array(1,2,3),
'*                                   Array(4,5,6),
'*                                   Array(7,8,9))
'*
'* @param lower1 1次元開始インデックス
'* @param upper1 1次元終了インデックス
'* @param lower2 2次元開始インデックス
'* @param upper2 2次元終了インデックス
'* @return Variant 2次元配列
'******************************************************************************
Public Function Create2DArrayWithValue(ParamArray arrList() As Variant) As Variant
    If IsMissing(arrList) Then
        GoTo InValidParam
    End If
    
    Dim v, c As Long, rownum As Long, colnum As Long
    c = 0: rownum = 0: colnum = 0
    
    For Each v In arrList
        If Not IsArray(v) Then
            GoTo InValidParam
        End If
        If colnum = 0 Then
            colnum = GetLength(v)
        Else
            colnum = GetLength(v)
            If colnum <> c Then GoTo InValidParam
        End If
        c = GetLength(v)
        rownum = rownum + 1
    Next
    
    Dim vNew2DArr: vNew2DArr = Create2DArrayEasy(rownum - 1, colnum - 1)
    Dim i As Long: i = LBound(vNew2DArr)

    For Each v In arrList
        Call Copy1DArrayTo2DArrayRow(vNew2DArr, i, v)
        i = i + 1
    Next
        
    Create2DArrayWithValue = vNew2DArr

    Exit Function
InValidParam:
    Create2DArrayWithValue = Empty
End Function

'******************************************************************************
'* [概  要] Copy2DArray
'* [詳  細] 2次元配列をコピーして新しい2次元配列を作成する。
'*          開始インデックスは0固定。
'*
'* @param source2DArr コピー元2次元配列
'* @param start1DIdx 1次元開始インデックス
'* @param end1DIdx   1次元終了インデックス
'* @param start2DIdx 2次元開始インデックス
'* @param end2DIdx   2次元終了インデックス
'* @return Variant 2次元配列
'******************************************************************************
Public Function Copy2DArray(ByRef source2DArr As Variant, _
                            Optional ByVal start1DIdx As Long = -1, Optional ByVal end1DIdx As Long = -1, _
                            Optional ByVal start2DIdx As Long = -1, Optional ByVal end2DIdx As Long = -1) As Variant
    If start1DIdx < 0 Then start1DIdx = LBound(source2DArr, 1)
    If end1DIdx < 0 Then end1DIdx = UBound(source2DArr, 1)
    If start2DIdx < 0 Then start2DIdx = LBound(source2DArr, 2)
    If end2DIdx < 0 Then end2DIdx = UBound(source2DArr, 2)

    Dim vNew2DArr As Variant: vNew2DArr = Create2DArray(0, end1DIdx - start1DIdx, 0, end2DIdx - start2DIdx)
    
    Dim i As Long, j As Long
    Dim x As Long, y As Long
    x = 0
    For i = start1DIdx To end1DIdx
        y = 0
        For j = start2DIdx To end2DIdx
            vNew2DArr(x, y) = source2DArr(i, j)
            y = y + 1
        Next j
        x = x + 1
    Next i
    Copy2DArray = vNew2DArr
End Function

'******************************************************************************
'* [概  要] Copy2DArrayRow
'* [詳  細] 2次元配列の行をコピーして新しい1次元配列を作成する。
'*          開始インデックスは0固定。
'*
'* @param source2DArr コピー元2次元配列
'* @param rowIdx 行インデックス
'* @return Variant 1次元配列
'******************************************************************************
Public Function Copy2DArrayRow(ByRef source2DArr As Variant, ByVal rowIdx As Long) As Variant
    Dim vNewArr As Variant: vNewArr = Create1DArray(0, UBound(source2DArr) - LBound(source2DArr))
    
    Dim colidx As Long
    Dim y As Long: y = 0
    For colidx = LBound(source2DArr) To UBound(source2DArr)
        vNewArr(y) = source2DArr(rowIdx, colidx)
        y = y + 1
    Next
    Copy2DArrayRow = vNewArr
End Function

'******************************************************************************
'* [概  要] Contains
'* [詳  細] 1次元配列に指定した項目が含まれるかチェックする。
'*
'* @param varr 1次元配列
'* @param itm 項目
'* @return Boolean チェック結果（True：含む）
'******************************************************************************
Public Function Contains(ByRef vArr As Variant, ByVal itm As Variant) As Boolean
    Dim v As Variant
    For Each v In vArr
        If Core.Eq(v, itm) Then
            Contains = True
            Exit Function
        End If
    Next
    Contains = False
End Function

'******************************************************************************
'* [概  要] Insert
'* [詳  細] 項目を1次元配列の指定位置に追加する。
'*
'* @param varr 1次元配列
'* @param itm 項目
'* @return Variant 追加後1次元配列
'******************************************************************************
Public Function Insert(ByRef vArr As Variant, ByVal idx As Long, ByVal itm As Variant) As Variant
    If idx < LBound(vArr) Or UBound(vArr) + 1 < idx Then
        Insert = vArr
        Exit Function
    ElseIf idx = LBound(vArr) Then
        If Me.IsEmptyArray(vArr) Then
            Dim ax As New ArrayEx
            Insert = ax.Add(itm).ToArray
            Exit Function
        End If
        Dim ax1 As New ArrayEx
        Insert = ax1.AddAll(itm, vArr).ToArray
        Exit Function
    ElseIf idx = UBound(vArr) + 1 Then
        Dim ax2 As New ArrayEx
        Insert = ax2.AddAll(vArr, itm).ToArray
        Exit Function
    End If

    Dim arr1: arr1 = Core.ArrSlice(vArr, LBound(vArr), idx - 1)
    Dim arr2: arr2 = Core.ArrSlice(vArr, idx, UBound(vArr))
    Dim ax3 As New ArrayEx
    Insert = ax3.AddAll(arr1, itm, arr2).ToArray
End Function

'******************************************************************************
'* [概  要] InsertAll
'* [詳  細] 項目（複数）を指定位置に追加する。
'*
'* @param varr 1次元配列
'* @param itms 項目
'* @return Variant 追加後1次元配列
'******************************************************************************
Public Function InsertAll(ByRef vArr As Variant, ByVal idx As Long, ParamArray itms() As Variant) As Variant
    If IsMissing(itms) Then
        InsertAll = vArr
        Exit Function
    End If
    Dim tmp: tmp = itms
    If idx < LBound(vArr) Or UBound(vArr) + 1 < idx Then
        InsertAll = vArr
        Exit Function
    ElseIf idx = LBound(vArr) Then
        If LBound(vArr) = UBound(vArr) Then
            Dim ax1 As New ArrayEx
            InsertAll = ax1.AddAll(tmp).ToArray
            Exit Function
        End If
        Dim ax2 As New ArrayEx
        InsertAll = ax2.AddAll(tmp, vArr).ToArray
        Exit Function
    ElseIf idx = UBound(vArr) + 1 Then
        Dim ax3 As New ArrayEx
        InsertAll = ax3.AddAll(vArr, tmp).ToArray
        Exit Function
    End If
    
    Dim arr1: arr1 = Core.ArrSlice(vArr, LBound(vArr), idx - 1)
    Dim arr2: arr2 = Core.ArrSlice(vArr, idx, UBound(vArr))
    Dim ax As New ArrayEx
    
    InsertAll = ax.AddAll(arr1, tmp, arr2).ToArray
End Function

'******************************************************************************
'* [概  要] LastIndexOf
'* [詳  細] 一次元配列から指定した値と一致する要素のインデックスを返却する。
'*          最後尾から検索して最初に一致したものを返却。
'*          値の捜査開始位置と要素数を任意指定可能。
'*
'* @param arr As Variant(Of Array(Of T))
'* @param val As Variant(Of T)
'* @param ixStart As Variant(Of Long)
'* @param cnt As Variant(Of Long)
'* @return As Long
'******************************************************************************
Public Function LastIndexOf(ByVal arr As Variant, ByVal val As Variant, _
    Optional ByVal ixStart As Variant, Optional ByVal cnt As Variant _
    ) As Long
    
    If Not IsArray(arr) Then Err.Raise 13
    
    Dim ixLast As Long:  ixLast = UBound(arr)
    Dim aLen As Long: aLen = Core.ArrLen(arr)
    If IsMissing(ixStart) Then ixStart = ixLast
    If IsNumeric(ixStart) Then ixStart = CLng(ixStart) Else Err.Raise 13
    If ixStart < LBound(arr) Or ixStart > ixLast Then Err.Raise 5
    If IsMissing(cnt) Then cnt = aLen
    If IsNumeric(cnt) Then cnt = CLng(cnt) Else Err.Raise 13
    cnt = Min(cnt, aLen)
    
    LastIndexOf = -1
    Dim ixEnd As Long: ixEnd = ixStart - cnt + 1
    If ixEnd < LBound(arr) Then
        ixEnd = LBound(arr)
    End If

    Dim i As Long
    For i = ixStart To ixEnd Step -1
        If Core.Equals(arr(i), val, True) Then
            LastIndexOf = i
            GoTo Escape
        End If
    Next
    
Escape:
End Function

'******************************************************************************
'* [概  要] IndexListOf
'* [詳  細] 1次元配列から指定した値と一致する要素のインデックスリストを返却する。
'*
'* @param arr 1次元配列
'* @param val 値
'* @param reverseOrder リストの並びを逆順にするか
'* @return Variant インデックスリスト（1次元配列）
'******************************************************************************
Public Function IndexListOf(ByVal arr As Variant, ByVal val As Variant, Optional ByVal reverseOrder As Boolean = False) As Variant
    Dim arrex As ArrayEx: Set arrex = New ArrayEx
    If Not Me.Is1DArray(arr) Or Me.IsEmptyArray(arr) Then
        IndexListOf = arrex.ToArray
        Exit Function
    End If
    Dim uIdx As Long, lIdx As Long
    uIdx = UBound(arr)
    lIdx = LBound(arr)

    Dim idx As Long, ret As Long
    If Not reverseOrder Then
        idx = lIdx
        Do
            ret = Me.IndexOf(arr, val, idx)
            If idx <= ret Then
                arrex.Add ret
                If ret = uIdx Then Exit Do
                idx = ret + 1
            Else
                Exit Do
            End If
        Loop
    Else
        idx = uIdx
        Do
            ret = Me.LastIndexOf(arr, val, idx)
            If ret >= 0 Then
                arrex.Add ret
                If ret = lIdx Then Exit Do
                idx = ret - 1
            Else
                Exit Do
            End If
        Loop
    End If
    IndexListOf = arrex.ToArray
End Function

'******************************************************************************
'* [概  要] Remove
'* [詳  細] 1次元配列の最初に見つかった特定の要素を削除する。
'*
'* @param varr 1次元配列
'* @param itm 項目
'* @return Variant 削除後1次元配列
'******************************************************************************
Public Function Remove(ByRef vArr As Variant, ByVal itm As Variant) As Variant
    Dim idx As Long: idx = IndexOf(vArr, itm)
    Remove = RemoveAt(vArr, idx)
End Function

'******************************************************************************
'* [概  要] RemoveAt
'* [詳  細] 1次元配列の指定したインデックスにある要素を削除する。
'*
'* @param varr 1次元配列
'* @param idx インデックス
'* @return Variant 削除後1次元配列
'******************************************************************************
Public Function RemoveAt(ByRef vArr As Variant, ByVal idx As Long) As Variant
    If idx < LBound(vArr) Or UBound(vArr) < idx Then
        RemoveAt = vArr
        Exit Function
    ElseIf idx = LBound(vArr) Then
        If LBound(vArr) = UBound(vArr) Then
            RemoveAt = Me.EmptyVariantArray
            Exit Function
        End If
        RemoveAt = Core.ArrSlice(vArr, idx + 1, UBound(vArr))
        Exit Function
    ElseIf idx = UBound(vArr) Then
        RemoveAt = Core.ArrSlice(vArr, LBound(vArr), idx - 1)
        Exit Function
    End If
    
    Dim arr1: arr1 = Core.ArrSlice(vArr, LBound(vArr), idx - 1)
    Dim arr2: arr2 = Core.ArrSlice(vArr, idx + 1, UBound(vArr))
    Dim ax As New ArrayEx
    RemoveAt = ax.AddAll(arr1, arr2).ToArray
End Function

'******************************************************************************
'* [概  要] RemoveRange
'* [詳  細] 1次元配列の要素の範囲を削除する。
'*
'* @param varr 1次元配列
'* @param idx1 開始インデックス
'* @param idx2 終了インデックス
'* @return Variant 削除後1次元配列
'******************************************************************************
Public Function RemoveRange(ByRef vArr As Variant, ByVal idx1 As Long, ByVal idx2 As Long) As Variant
    If idx1 < LBound(vArr) Or UBound(vArr) < idx1 Or idx2 < LBound(vArr) Or UBound(vArr) < idx2 Or idx1 > idx2 Then
        RemoveRange = vArr
        Exit Function
    ElseIf idx1 = LBound(vArr) Then
        If LBound(vArr) = UBound(vArr) Then
            RemoveRange = Me.EmptyVariantArray
            Exit Function
        End If
        If idx2 = UBound(vArr) Then
            RemoveRange = Me.EmptyVariantArray
            Exit Function
        End If
        RemoveRange = Core.ArrSlice(vArr, idx2 + 1, UBound(vArr))
        Exit Function
    ElseIf idx1 = UBound(vArr) Then
        RemoveRange = Core.ArrSlice(vArr, LBound(vArr), UBound(vArr) - 1)
        Exit Function
    ElseIf idx2 = UBound(vArr) Then
        RemoveRange = Core.ArrSlice(vArr, LBound(vArr), idx1 - 1)
        Exit Function
    End If
    
    Dim arr1: arr1 = Core.ArrSlice(vArr, LBound(vArr), idx1 - 1)
    Dim arr2: arr2 = Core.ArrSlice(vArr, idx2 + 1, UBound(vArr))
    Dim ax As New ArrayEx
    RemoveRange = ax.AddAll(arr1, arr2).ToArray
End Function

'******************************************************************************
'* [概  要] Repeat
'* [詳  細] 指定した要素を指定回数分追加した1次元配列を作成する。
'*
'* @param itm 項目
'* @param count 項目
'* @return Variant 1次元配列
'******************************************************************************
Public Function Repeat(ByVal itm As Variant, ByVal Count As Long) As Variant
    If Count <= 0 Then
        Repeat = EmptyVariantArray
        Exit Function
    End If
    Dim vArr
    ReDim vArr(0 To Count - 1)
    Dim i As Long
    For i = 0 To Count - 1
        If IsObject(itm) Then
            Set vArr(i) = itm
        Else
            vArr(i) = itm
        End If
    Next
    Repeat = vArr
End Function

'******************************************************************************
'* [概  要] ToString
'* [詳  細] JSON文字列に変換する。
'*
'* @param varr 1次元配列
'* @return String 1次元配列
'******************************************************************************
Public Function ToString(ByRef vArr As Variant) As String
    Dim strx As StringEx: Set strx = New StringEx

    strx.Append "["
    Dim i As Long
    For i = LBound(vArr) To UBound(vArr)
        strx.Append VariantUtils.ToString(vArr(i))
        If i < UBound(vArr) Then
            strx.Append ", "
        End If
    Next
    strx.Append "]"
    ToString = strx.ToString
End Function

'******************************************************************************
'* [概  要] GetFirst
'* [詳  細] 最初の要素を取得する。
'*
'* @param varr 1次元配列
'* @return Variant 最初の要素
'******************************************************************************
Public Function GetFirst(ByRef vArr As Variant) As Variant
    If Not IsArray(vArr) Then
        GetFirst = Empty
        Exit Function
    End If
    If IsObject(vArr(UBound(vArr))) Then
        Set GetFirst = vArr(LBound(vArr))
    Else
        GetFirst = vArr(LBound(vArr))
    End If
End Function

'******************************************************************************
'* [概  要] GetLast
'* [詳  細] 最後の要素を取得する。
'*
'* @param varr 1次元配列
'* @return Variant 最後の要素
'******************************************************************************
Public Function GetLast(ByRef vArr As Variant) As Variant
    If Not IsArray(vArr) Then
        GetLast = Empty
        Exit Function
    End If
    If IsObject(vArr(UBound(vArr))) Then
        Set GetLast = vArr(UBound(vArr))
    Else
        GetLast = vArr(UBound(vArr))
    End If
End Function

'******************************************************************************
'* [概  要] Join
'* [詳  細] 配列を連結して文字列にする。
'*
'* @param varr 1次元配列
'* @param delimiter デリミタ
'* @param wrapChar 囲み文字
'* @param rmLineSeparator 改行を除去するか
'* @return String 連結文字列
'******************************************************************************
Public Function Join(ByRef vArr As Variant, Optional ByVal Delimiter As String, Optional ByVal wrapChar As String, Optional ByVal rmLineSeparator As Boolean = True) As String
    If Not IsArray(vArr) Then
        Join = ""
        Exit Function
    End If
    If wrapChar = "" And rmLineSeparator = False Then
        Join = VBA.Join(vArr, Delimiter)
        Exit Function
    End If
    Dim tmpItm As String
    Dim v, sbuf As StringEx: Set sbuf = New StringEx
    Dim firstAct As Boolean: firstAct = True
    For Each v In vArr
        If firstAct Then
            sbuf.Append wrapChar
            sbuf.Append IIf(rmLineSeparator, StringUtils.RemoveLineSeparator(CStr(v)), CStr(v))
            sbuf.Append wrapChar
            firstAct = False
        Else
            sbuf.Append Delimiter
            sbuf.Append wrapChar
            sbuf.Append IIf(rmLineSeparator, StringUtils.RemoveLineSeparator(CStr(v)), CStr(v))
            sbuf.Append wrapChar
        End If
    Next
    Join = sbuf.ToString
End Function

'******************************************************************************
'* [概  要] Search
'* [詳  細] 配列（文字列）から指定した文字列と部分一致する項目を
'*          取得し、1次元配列で返却する。
'*
'* @param vArr 配列
'* @param searchVal 検索文字列（配列で複数指定可）
'* @return Variant 検索結果（1次元配列）
'******************************************************************************
Public Function Search(ByRef vArr As Variant, ByVal searchVal As Variant) As Variant
    Dim arrx As ArrayEx: Set arrx = New ArrayEx
    Dim v As Variant
    For Each v In vArr
        If ValidateUtils.IsString(v) Then
            If StringUtils.ContainsAll(v, searchVal) Then
                arrx.Add CStr(v)
            End If
        End If
    Next
    Search = arrx.ToArray
End Function

'******************************************************************************
'* [概  要] RegexSearch
'* [詳  細] 配列（文字列）から指定した文字列と正規表現が一致する項目を
'*          取得し、1次元配列で返却する。
'*
'* @param vArr 配列
'* @param searchStr 検索文字列（正規表現）
'* @param igCase 大文字小文字の区別
'* @return Variant 検索結果（1次元配列）
'******************************************************************************
Public Function RegexSearch(ByRef vArr As Variant, ByVal searchStr As String, Optional ByVal igCase As Boolean = False) As Variant
    Dim arrx As ArrayEx: Set arrx = New ArrayEx
    Dim v As Variant
    For Each v In vArr
        If ValidateUtils.IsString(v) Then
            If ValidateUtils.ValidatePattern(v, searchStr, igCase) Then
                arrx.Add CStr(v)
            End If
        End If
    Next
    RegexSearch = arrx.ToArray
End Function

'******************************************************************************
'* [概  要] Array2DExTo2DArray
'* [詳  細] Array2DExを2次元配列に変換する。
'*
'* @param arr2dex Array2DExオブジェクト
'* @return Variant 2次元配列
'******************************************************************************
Public Function Array2DExTo2DArray(ByVal arr2dex As Array2DEx) As Variant
    Dim arr2d As Variant: arr2d = ArrayUtils.EmptyVariant2DArray
    
    Dim ixOut As Long, ixInn As Long
    Dim lbOut As Long, lbInn As Long, lbInnFst As Long
    Dim ubOut As Long, ubInn As Long, ubInnFst As Long
    
    lbOut = 0
    ubOut = arr2dex.RowMaxIndex
    If ubOut < 0 Then GoTo Ending
    
    lbInnFst = 0
    ubInnFst = arr2dex.ColMaxIndex
    
    ReDim arr2d(lbOut To ubOut, lbInnFst To ubInnFst)
    If IsObject(arr2dex(lbOut, lbInnFst)) Then
        For ixOut = lbOut To ubOut
            lbInn = 0
            ubInn = arr2dex.ColMaxIndex
            For ixInn = lbInn To ubInn: Set arr2d(ixOut, ixInn) = arr2dex(ixOut, ixInn): Next
        Next
    Else
        For ixOut = lbOut To ubOut
            lbInn = 0
            ubInn = arr2dex.ColMaxIndex
            For ixInn = lbInn To ubInn: arr2d(ixOut, ixInn) = arr2dex(ixOut, ixInn): Next
        Next
    End If

Ending:
    Array2DExTo2DArray = arr2d
End Function

'******************************************************************************
'* [概  要] Array2DToArray2DEx
'* [詳  細] 2次元配列をArray2DExに変換する。
'*
'* @param arr2d 2次元配列
'* @return Array2DEx Array2DExオブジェクト
'******************************************************************************
Public Function Array2DToArray2DEx(ByVal arr2d As Variant) As Array2DEx
    Dim arr2dex As Array2DEx: Set arr2dex = New Array2DEx
    Set Array2DToArray2DEx = arr2dex.Init(arr2d)
End Function

'******************************************************************************
'* [概  要] CreateArrayEx
'* [詳  細] ArrayExオブジェクトを作成する。
'*
'* @param itemCount 項目数
'* @return ArrayEx オブジェクト
'******************************************************************************
Public Function CreateArrayEx(Optional ByVal itemCount) As ArrayEx
    Set CreateArrayEx = New ArrayEx
    If Not VBA.IsMissing(itemCount) Then
        CreateArrayEx.Expand itemCount
    End If
End Function

'******************************************************************************
'* [概  要] CreateArray2DEx
'* [詳  細] Array2DExオブジェクトを作成する。
'*
'* @param rowCount 行数
'* @param colCount 列数
'* @return Array2DEx オブジェクト
'******************************************************************************
Public Function CreateArray2DEx(Optional ByVal rowCount, Optional ByVal colCount) As Array2DEx
    Set CreateArray2DEx = New Array2DEx
    If Not VBA.IsMissing(rowCount) And Not VBA.IsMissing(colCount) Then
        CreateArray2DEx.Expand rowCount, colCount
    ElseIf Not VBA.IsMissing(rowCount) Then
        CreateArray2DEx.ExpandRows rowCount
    ElseIf Not VBA.IsMissing(colCount) Then
        CreateArray2DEx.ExpandColumns colCount
    End If
End Function

'******************************************************************************
'* [概  要] Create1DArray
'* [詳  細] 新しい1次元配列を作成する。
'*
'* @param lower 開始インデックス
'* @param upper 終了インデックス
'* @return Variant 1次元配列
'******************************************************************************
Public Function Create1DArray(ByVal lower As Long, ByVal upper As Long) As Variant
    Dim vNewArr As Variant
    ReDim vNewArr(lower To upper) As Variant
    Create1DArray = vNewArr
End Function

'******************************************************************************
'* [概  要] Create1DArrayEasy
'* [詳  細] 新しい1次元配列を作成する。開始インデックスは0固定
'*
'* @param upper 終了インデックス
'* @return Variant 1次元配列
'******************************************************************************
Public Function Create1DArrayEasy(ByVal upper As Long) As Variant
    Dim vNewArr As Variant
    ReDim vNewArr(0 To upper) As Variant
    Create1DArrayEasy = vNewArr
End Function


'******************************************************************************
'* [概  要] ReadCsvFile
'* [詳  細] CSVファイルを行単位で読込み、Variant配列に設定して返却する。
'*
'* @param filePath CSVファイルパス
'* @param headerExists ヘッダ行の有無
'* @param quotExists ""で括っているか
'* @param separator 改行コード
'* @return 読込テキストデータ（Variant2次元配列）
'******************************************************************************
Public Function CsvStringToArray(ByVal csvData As String, _
                                 Optional ByVal HeaderExists As Boolean = True, _
                                 Optional ByVal QuotExists As Boolean = False, _
                                 Optional ByVal separator As LineSeparatorsEnum = adCRLF) As Variant
    Dim vLine, lCount As Long, i As Long, vRow, vArr()
    Dim vData As Variant: vData = Split(csvData, IO.FromLineSeparatorsEnum(separator))
    Dim lSize As Long: lSize = GetLength(vData)
    
     ' 空データは空Variant配列で返却
    If (HeaderExists And lSize <= 1) Or (Not HeaderExists And lSize = 0) Then
        Exit Function
    End If

    For Each vLine In vData
        ' 1行目処理
        If lCount = 0 Then
            vRow = StringUtils.SplitCsvRow(CStr(vLine), QuotExists)
            If Not HeaderExists Then
                ReDim vArr(0 To lSize - 1, 0 To UBound(vRow))
                For i = LBound(vRow) To UBound(vRow)
                    vArr(lCount, i) = vRow(i)
                Next
                lCount = lCount + 1
            Else
                ReDim vArr(0 To lSize - 2, 0 To UBound(vRow))
            End If
            Dim colCount As Long: colCount = UBound(vRow)
        End If

        ' 2行目以降処理
        vRow = StringUtils.SplitCsvRow(CStr(vLine), QuotExists)
        If UBound(vRow) <> colCount Then
            CsvStringToArray = Empty
            Exit Function
        End If
        
        For i = LBound(vRow) To UBound(vRow)
            vArr(lCount, i) = vRow(i)
        Next
        lCount = lCount + 1
    Next
    CsvStringToArray = vArr
End Function


