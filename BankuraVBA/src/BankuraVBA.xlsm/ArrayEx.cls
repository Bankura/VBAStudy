VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''+----                                                                   --+
'''|                             Ariawase 0.9.0                              |
'''|                Ariawase is free library for VBA cowboys.                |
'''|          The Project Page: https://github.com/vbaidiot/Ariawase         |
'''+--                                                                   ----+
Option Explicit

Private xItems As Variant
Private xLength As Long
Private xIndex As Long

Private Sub Class_Initialize()
    xIndex = -1
    xLength = -1 + 32
    ReDim xItems(xLength - 1)
End Sub

Private Sub Extend()
    If xIndex < xLength Then GoTo Escape
    xLength = xLength + 1 'possible overflow (Err.Raise 6)
    xLength = -1 + xLength + xLength
    ReDim Preserve xItems(xLength - 1)
    
Escape:
End Sub

Public Function AddVal(ByVal val As Variant)
    xIndex = xIndex + 1
    Extend
    Let xItems(xIndex) = val
End Function

Public Function AddObj(ByVal obj As Variant)
    xIndex = xIndex + 1
    Extend
    Set xItems(xIndex) = obj
End Function

Public Function ToArray() As Variant
    Dim arr As Variant: arr = xItems
    If xIndex > -1 Then
        ReDim Preserve arr(xIndex)
    Else
        arr = Array()
    End If
    ToArray = arr
End Function


'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* 拡張メソッド
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'*-----------------------------------------------------------------------------
'* Item プロパティ
'*-----------------------------------------------------------------------------
Public Property Get Item(idx As Long) As Variant
    If IsInvalidIndex(idx) Then
        Item = Empty
        Exit Property
    End If

    If IsObject(xItems(idx)) Then
        Set Item = xItems(idx)
    Else
        Item = xItems(idx)
    End If
End Property
Public Property Let Item(idx As Long, arg As Variant)
    If IsInvalidIndex(idx) Then
        Exit Property
    End If
    
    If IsObject(arg) Then
        Set xItems(idx) = arg
    Else
        xItems(idx) = arg
    End If
End Property

Private Function IsInvalidIndex(idx As Long) As Boolean
    IsInvalidIndex = (idx < 0 Or xIndex < idx)
End Function


'******************************************************************************
'* [概  要] Init メソッド
'* [詳  細] 初期化用メソッド｡
'*
'* @param val 初期配列
'* @return ArrayExオブジェクト
'*
'******************************************************************************
Public Function Init(Optional ByVal val) As ArrayEx
    If IsArray(val) Then
        xIndex = UBound(val)
        xLength = xIndex + 32
        xItems = val
        ReDim xItems(xLength - 1)
    Else
        Call Clear
    End If
    Set Init = Me
End Function

'******************************************************************************
'* [概  要] Clear
'* [詳  細] 初期化する｡
'*
'* @return ArrayExオブジェクト
'*
'******************************************************************************
Public Function Clear() As ArrayEx
    xIndex = -1
    xLength = -1 + 32
    ReDim xItems(xLength - 1)
    Set Clear = Me
End Function

'******************************************************************************
'* [概  要] Length
'* [詳  細] 要素数を返却する｡
'*
'* @return Long 要素数
'*
'******************************************************************************
Property Get Length() As Long
    Length = xIndex + 1
End Property

'******************************************************************************
'* [概  要] LastIndex
'* [詳  細] 最終インデックスを返却する｡
'*
'* @return Long 最終インデックス
'*
'******************************************************************************
Property Get LastIndex() As Long
    LastIndex = xIndex
End Property

'******************************************************************************
'* [概  要] Add
'* [詳  細] 項目を追加する｡
'*
'* @param var 追加項目
'* @return ArrayEx
'*
'******************************************************************************
Public Function Add(ByVal var As Variant) As ArrayEx
    If IsObject(var) Then
        Call AddObj(var)
    Else
        Call AddVal(var)
    End If
    Set Add = Me
End Function

'******************************************************************************
'* [概  要] AddAll
'* [詳  細] 項目を複数追加する｡
'*
'* @param vars 追加項目（複数）
'* @return ArrayEx
'*
'******************************************************************************
Public Function AddAll(ParamArray vars() As Variant) As ArrayEx
    Dim tmp: tmp = vars
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        Add v
    Next
    Set AddAll = Me
End Function

'******************************************************************************
'* [概  要] IsEmptyArray
'* [詳  細] 配列が未初期化・空配列か判定する｡
'*
'* @return Boolean 判定結果(True:未初期化・空)
'*
'******************************************************************************
Public Function IsEmptyArray() As Boolean
    IsEmptyArray = (xIndex = -1)
End Function

'******************************************************************************
'* [概  要] Equals
'* [詳  細] 一次元配列を比較し、イコールであればTrueを返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr2 As Variant(Of Array(Of T))
'* @param swAllowNull As Boolean
'* @return As Variant(Of Boolean Or Null Or Empty)
'******************************************************************************
Public Function Equals(ByVal arr2 As Variant, Optional ByVal swAllowNull As Boolean = False) As Variant
    Equals = Core.ArrEquals(Me.ToArray, arr2, swAllowNull)
End Function

'******************************************************************************
'* [概  要] ArrayCompare
'* [詳  細] 一次元配列（x,y）を比較し、x=yであれば0、x<yであれば-1、x>y
'*          であれば1を返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr2 As Variant(Of Array(Of T))
'* @return As Variant(Of Integer Or Null)
'******************************************************************************
Public Function Compare(ByVal arr2 As Variant, Optional ByVal swAllowNull As Boolean = False) As Variant
    Compare = Core.ArrCompare(Me.ToArray, arr2, swAllowNull)
End Function

'******************************************************************************
'* [概  要] IndexOf
'* [詳  細] 一次元配列から指定した値と一致する要素のインデックスを返却する。
'*          値の捜査開始位置と要素数を任意指定可能。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param val As Variant(Of T)
'* @param ixStart As Variant(Of Long)
'* @param cnt As Variant(Of Long)
'* @return As Long
'******************************************************************************
Public Function IndexOf(ByVal val As Variant, Optional ByVal ixStart As Variant, Optional ByVal cnt As Variant) As Long
    IndexOf = Core.ArrIndexOf(Me.ToArray, val, ixStart, cnt)
End Function

'******************************************************************************
'* [概  要] Reverse
'* [詳  細] 一次元配列の要素の並び順を反転させる。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr As Variant(Of Array(Of T))
'******************************************************************************
Public Sub Reverse()
    Dim arr: arr = Me.ToArray
    Call Core.ArrRev(arr)
    Call Init(arr)
End Sub

'******************************************************************************
'* [概  要] Sort
'* [詳  細] 一次元配列のソートを行う。デフォルトは昇順。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param orderAsc As Boolean
'******************************************************************************
Public Sub Sort(Optional ByVal orderAsc As Boolean = True)
    Dim arr: arr = Me.ToArray
    Call Core.ArrSort(arr, orderAsc)
    Call Init(arr)
End Sub

'******************************************************************************
'* [概  要] Uniq
'* [詳  細] 一次元配列の重複要素を排除する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function Uniq() As Variant
    Uniq = Core.ArrUniq(Me.ToArray)
End Function

'******************************************************************************
'* [概  要] Concat
'* [詳  細] 一次元配列toと合体（マージ）して返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr1 As Variant(Of Array(Of T))
'* @param arr2 As Variant(Of Array(Of T))
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function Concat(ByVal arr2 As Variant) As Variant
    Concat = Core.ArrConcat(Me.ToArray, arr2)
End Function

'******************************************************************************
'* [概  要] Slice
'* [詳  細] 一次元配列をスライス（一部を配列として切り出）して返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param ixStart As Variant(Of Long)
'* @param ixEnd As Variant(Of Long)
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function Slice(Optional ByVal ixStart As Variant, Optional ByVal ixEnd As Variant) As Variant
    Slice = Core.ArrSlice(Me.ToArray, ixStart, ixEnd)
End Function

'******************************************************************************
'* [概  要] FromJagArray
'* [詳  細] ジャグ配列のフラット化。ジャグ配列を一次元配列(ArrayEx)に変換して
'*          返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param jagArray As Variant(Of Array(Of Array(Of T)))
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function FromJagArray(ByVal jagArr As Variant) As ArrayEx
    Dim arr: arr = Core.ArrFlatten(jagArr)
    Set FromJagArray = Init(arr)
End Function

'******************************************************************************
'* [概  要] ToCollection
'* [詳  細] 一次元配列をCollectionオブジェクトに変換する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @return As Collection(Of T)
'******************************************************************************
Public Function ToCollection() As Collection
    Set ToCollection = Core.ArrToClct(Me.ToArray)
End Function

'******************************************************************************
'* [概  要] FromCollection
'* [詳  細] Collectionオブジェクトを一次元配列(ArrayEx)に変換する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param clct As Collection(Of T)
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function FromCollection(ByVal clct As Collection) As ArrayEx
    Dim arr: arr = Core.ClctToArr(clct)
    Set FromCollection = Init(arr)
End Function

'******************************************************************************
'* [概  要] FromEnumerator
'* [詳  細] 繰り返し処理（For Each）可能なオブジェクトを一次元配列(ArrayEx)に
'*          変換する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param enumr As Enumerator(Of T)
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function FromEnumerator(ByVal enumr As Object) As ArrayEx
    Dim arr: arr = Ext.EnumeratorToArr(enumr)
    Set FromEnumerator = Init(arr)
End Function

'******************************************************************************
'* [概  要] Range
'* [詳  細] 指定した範囲の数値の配列を作成する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fromVal As Variant(Of T)
'* @param toVal As Variant(Of T)
'* @param stepVal As Variant(Of T)
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function Range(ByVal fromVal As Variant, ByVal toVal As Variant, Optional ByVal stepVal As Variant = 1) As ArrayEx
    Dim arr: arr = Ext.ArrRange(fromVal, toVal, stepVal)
    Set Range = Init(arr)
End Function

'******************************************************************************
'* [概  要] Map
'* [詳  細] 配列の全要素に対して指定したファンクションの処理を適用する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of T, U)
'* @return ArrayEx(Of Array(Of U))
'******************************************************************************
Public Function Map(ByVal fun As Func) As ArrayEx
    Dim arr: arr = Ext.ArrMap(fun, Me.ToArray)
    Set Map = Init(arr)
End Function

'******************************************************************************
'* [概  要] Zip
'* [詳  細] ２つの配列に対して指定したファンクションの処理を適用し
'*          １つの配列を返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of T, U, R)
'* @param arr2 As Variant(Of Array(Of U))
'* @return ArrayEx(Of Array(Of R))
'******************************************************************************
Public Function Zip(ByVal fun As Func, ByVal arr2 As Variant) As ArrayEx
    Dim arr: arr = Ext.ArrZip(fun, Me.ToArray, arr2)
    Set Zip = Init(arr)
End Function

'******************************************************************************
'* [概  要] Filter
'* [詳  細] 配列に対して指定したフィルタ処理用のファンクションを適用し、
'*          条件に一致する要素のみの配列を返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of T, Boolean)
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function Filter(ByVal fun As Func) As ArrayEx
    Dim arr: arr = Ext.ArrFilter(fun, Me.ToArray)
    Set Filter = Init(arr)
End Function

'******************************************************************************
'* [概  要] GroupBy
'* [詳  細] 配列に対して指定したグルーピング処理用のファンクションを適用し、
'*          結果を返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of T, K)
'* @return ArrayEx(Of Array(Of Tuple`2(Of K, T)))
'******************************************************************************
Public Function GroupBy(ByVal fun As Func) As ArrayEx
    Dim arr: arr = Ext.ArrGroupBy(fun, Me.ToArray)
    Set GroupBy = Init(arr)
End Function

'******************************************************************************
'* [概  要] Fold
'* [詳  細] 畳み込み関数。配列内の要素に対して指定した演算用ファンクション
'*          を適用し、その結果と次の要素にまたファンクションを適用する
'*          といった処理を繰り返し最終的な結果を返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of U, T, U)
'* @param seedv As Variant(Of U)
'* @return ArrayEx(Of Array(Of U))
'******************************************************************************
Public Function Fold(ByVal fun As Func, Optional ByVal seedv As Variant) As ArrayEx
    Dim arr: arr = Ext.ArrFold(fun, Me.ToArray, seedv)
    Set Fold = Init(arr)
End Function

'******************************************************************************
'* [概  要] Scan
'* [詳  細] 配列の各要素に対して繰り返し指定した演算用ファンクションを適用し、
'*          累積した計算結果のリストを返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of U, T, U)
'* @param seedv As Variant(Of U)
'* @return ArrayEx(Of Array(Of U))
'******************************************************************************
Public Function Scan(ByVal fun As Func, Optional ByVal seedv As Variant) As ArrayEx
    Dim arr: arr = Ext.ArrScan(fun, Me.ToArray, seedv)
    Set Scan = Init(arr)
End Function

'******************************************************************************
'* [概  要] Unfold
'* [詳  細] Foldの逆。初期値から指定した演算用ファンクションを適用しリストを
'*          生成する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func
'* @param seedv As Variant(Of T)
'* @return ArrayEx(Of Array(Of U))
'******************************************************************************
Public Function Unfold(ByVal fun As Func, ByVal seedv As Variant) As ArrayEx
    Dim arr: arr = Ext.ArrUnfold(fun, seedv)
    Set Unfold = Init(arr)
End Function

'******************************************************************************
'* [概  要] CopyTo2DArrayRow
'* [詳  細] 2次元配列の指定行に対して、1次元配列の値を設定する。
'*
'* @param v2DArr 2次元配列
'* @param rowNum 指定行
'* @return Variant 変更後2次元配列
'******************************************************************************
Public Function CopyTo2DArrayRow(ByRef v2DArr As Variant, ByVal rownum As Long) As Variant
    CopyTo2DArrayRow = ArrayUtils.Copy1DArrayTo2DArrayRow(v2DArr, rownum, Me.ToArray)
End Function

'******************************************************************************
'* [概  要] Contains
'* [詳  細] 指定した項目が含まれるかチェックする。
'*
'* @param itm 項目
'* @return Boolean チェック結果（True：含む）
'******************************************************************************
Public Function Contains(ByVal itm As Variant) As Boolean
    Contains = ArrayUtils.Contains(Me.ToArray, itm)
End Function

'******************************************************************************
'* [概  要] Clone
'* [詳  細] ArrayExオブジェクトを複製する。
'*
'* @return ArrayEx 新しいArrayEx
'******************************************************************************
Public Function Clone() As ArrayEx
    Dim ax As New ArrayEx
    Set Clone = ax.AddAll(Me.ToArray)
End Function

'******************************************************************************
'* [概  要] Insert
'* [詳  細] 項目を指定位置に追加する。
'*
'* @param itm 項目
'* @return ArrayEx
'******************************************************************************
Public Function Insert(ByVal idx As Long, ByVal itm As Variant) As ArrayEx
    Set Insert = Init(ArrayUtils.ArrayInsert(Me.ToArray, idx, itm))
End Function

'******************************************************************************
'* [概  要] InsertAll
'* [詳  細] 項目（複数）を指定位置に追加する。
'*
'* @param itms 項目
'* @return ArrayEx
'******************************************************************************
Public Function InsertAll(ByVal idx As Long, ParamArray itms() As Variant) As ArrayEx
    Dim tmp: tmp = itms
    Set InsertAll = Init(ArrayUtils.ArrayInsertAll(Me.ToArray, idx, tmp))
End Function

'******************************************************************************
'* [概  要] LastIndexOf
'* [詳  細] 指定した値と一致する要素のインデックスを返却する。
'*          最後尾から検索して最初に一致したものを返却。
'*          値の捜査開始位置と要素数を任意指定可能。
'*
'* @param val As Variant(Of T)
'* @param ixStart As Variant(Of Long)
'* @param cnt As Variant(Of Long)
'* @return As Long
'******************************************************************************
Public Function LastIndexOf(ByVal val As Variant, _
    Optional ByVal ixStart As Variant, Optional ByVal cnt As Variant _
    ) As Long
    LastIndexOf = ArrayUtils.ArrayLastIndexOf(Me.ToArray, val, ixStart, cnt)
End Function

'******************************************************************************
'* [概  要] Remove
'* [詳  細] 1次元配列の最初に見つかった特定の要素を削除する。
'*
'* @param itm 項目
'* @return ArrayEx 削除後1次元配列
'******************************************************************************
Public Function Remove(ByVal itm As Variant) As ArrayEx
    Set Remove = Init(ArrayUtils.Remove(Me.ToArray, itm))
End Function

'******************************************************************************
'* [概  要] RemoveAt
'* [詳  細] 指定したインデックスにある要素を削除する。
'*
'* @param idx インデックス
'* @return ArrayEx 削除後1次元配列
'******************************************************************************
Public Function RemoveAt(ByVal idx As Long) As ArrayEx
    Set RemoveAt = Init(ArrayUtils.RemoveAt(Me.ToArray, idx))
End Function

'******************************************************************************
'* [概  要] RemoveRange
'* [詳  細] 1次元配列の要素の範囲を削除する。
'*
'* @param idx1 開始インデックス
'* @param idx2 終了インデックス
'* @return ArrayEx 削除後1次元配列
'******************************************************************************
Public Function RemoveRange(ByVal idx1 As Long, ByVal idx2 As Long) As ArrayEx
    Set RemoveRange = Init(ArrayUtils.RemoveRange(Me.ToArray, idx1, idx2))
End Function

'******************************************************************************
'* [概  要] Repeat
'* [詳  細] 指定した要素を指定回数分追加した1次元配列を作成する。
'*
'* @param itm 項目
'* @param count 項目
'* @return Variant 1次元配列
'******************************************************************************
Public Function Repeat(ByVal itm As Variant, ByVal count As Long) As ArrayEx
    Set Repeat = Init(ArrayUtils.Repeat(itm, count))
End Function
