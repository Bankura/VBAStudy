VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] コレクションユーティリティクラス
'* [詳  細] コレクション関連のユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'*
'* [依　存] VBA-Web v4.1.6
'*          Ariawase 0.9.0
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2020-2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Sub CopyMemoryByRef Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************
#If Win64 Then
    Private Const FOWARD_SHIFT As Long = 40
    Private Const BACKWARD_SHIFT As Long = 48
    Private Const KEY_POS As Long = 24
#Else
    Private Const FOWARD_SHIFT As Long = 24
    Private Const BACKWARD_SHIFT As Long = 28
    Private Const KEY_POS As Long = 16
#End If

'******************************************************************************
'* 内部変数定義
'******************************************************************************

'******************************************************************************
'* プロパティ定義
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
'* [概  要] Push
'* [詳  細] Collectionに値をPushする。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param clct As Collection(Of T)
'* @param val As Variant(Of T)
'******************************************************************************
Public Sub Push(ByVal clct As Collection, ByVal val As Variant)
    Call Core.Push(clct, val)
End Sub

'******************************************************************************
'* [概  要] Pop
'* [詳  細] Collectionから値をPopする。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param clct As Collection(Of T)
'* @return As Variant(Of T)
'******************************************************************************
Public Function Pop(ByVal clct As Collection) As Variant
    Pop = Core.Pop(clct)
End Function

'******************************************************************************
'* [概  要] Shift
'* [詳  細] Collectionに値をShiftする。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param clct As Collection(Of T)
'* @param val As Variant(Of T)
'******************************************************************************
Public Sub shift(ByVal clct As Collection, ByVal val As Variant)
    Call Core.shift(clct, val)
End Sub

'******************************************************************************
'* [概  要] Unshift
'* [詳  細] Collectionから値をUnshiftする。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param clct As Collection(Of T)
'* @return As Variant(Of T)
'******************************************************************************
Public Function Unshift(ByVal clct As Collection) As Variant
    Unshift = Core.Unshift(clct)
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
'* [概  要] AssocArrayToDictionary
'* [詳  細] 連想配列をDictionaryオブジェクトに変換する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param aarr Variant(Of Array(Of Tupple))
'* @return Object(Of Dictionary)
'******************************************************************************
Public Function AssocArrayToDictionary(ByVal aarr As Variant) As Object
    Set AssocArrayToDictionary = ext.AssocArrToDict(aarr)
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
    DictionaryToAssocArray = ext.DictToAssocArr(dict)
End Function

'******************************************************************************
'* [概  要] HasKey
'* [詳  細] Collectionにキーに対応する値があるかチェック｡
'*
'* @param col 検査するCollection
'* @param keyName キー名
'* @return Boolean 判定結果(True:あり)
'*
'******************************************************************************
Public Function HasKey(ByRef col As Collection, ByVal keyName As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = col(keyName)
    HasKey = (Err.Number = 0)
    Err.Clear
End Function

'******************************************************************************
'* [概  要] IsEmpty
'* [詳  細] Collectionが空（Nothing含む）かチェック｡
'*
'* @param col 検査するCollection
'* @return Boolean 判定結果(True:あり)
'*
'******************************************************************************
Public Function IsEmpty(ByRef col As Collection) As Boolean
    IsEmpty = True
    If col Is Nothing Then
        Exit Function
    ElseIf col.Count = 0 Then
        Exit Function
    End If
    IsEmpty = False
End Function

'******************************************************************************
'* [概  要] GetCollectionKeyByIndex メソッド
'* [詳  細] コレクションのキーをIndexを指定して取得する。
'* [参  考] <https://foolexp.wordpress.com/2012/12/03/p2wpsu-5k/>
'*
'* @param index インデックス
'* @param col コレクション
'* @param colChange コレクションの変更有無
'*                 （コレクション生成後に値を入れ替えている場合は
'*                   前方からの検索しかできないのでTrueを指定）
'* @return キー名
'*
'******************************************************************************
Public Function GetCollectionKeyByIndex(ByVal ind As Long, col As Collection, Optional colChange As Boolean = False) As String
    Dim ptr As LongPtr

    Dim i As LongPtr
    Dim sKey As String
    If col Is Nothing Then
        Err.Raise 91
    Else
        If ind > col.Count Then
            Err.Raise 9
        ElseIf (ind <= (col.Count + 1) / 2) Or colChange Then
            Call CopyMemoryByRef(ptr, ByVal (ObjPtr(col) + FOWARD_SHIFT), LONGPTR_SIZE)
            For i = 2 To ind
                Call CopyMemoryByRef(ptr, ByVal (ptr + FOWARD_SHIFT), LONGPTR_SIZE)
            Next
        Else
            Call CopyMemoryByRef(ptr, ByVal (ObjPtr(col) + BACKWARD_SHIFT), LONGPTR_SIZE)
            For i = col.Count - 1 To ind Step -1
                Call CopyMemoryByRef(ptr, ByVal (ptr + BACKWARD_SHIFT), LONGPTR_SIZE)
            Next
        End If

        Dim tmp As LongPtr
        tmp = StrPtr(sKey)
        Call CopyMemoryByRef(ByVal VarPtr(sKey), ByVal (ptr + KEY_POS), LONGPTR_SIZE)
        
        GetCollectionKeyByIndex = sKey
        Call CopyMemoryByRef(ByVal VarPtr(sKey), tmp, LONGPTR_SIZE)
    End If
End Function

'******************************************************************************
'* [概  要] GetCollectionIndexByKey メソッド
'* [詳  細] コレクションのIndexをキーを指定して取得する。
'* [参  考] <https://foolexp.wordpress.com/2012/12/03/p2wpsu-5k/>
'*
'* @param key キー
'* @param col コレクション
'* @return Index
'*
'******************************************************************************
Public Function GetCollectionIndexByKey(ByVal Key As String, col As Collection, Optional ByVal comp As VbCompareMethod = vbTextCompare) As Long
    Dim ptr   As LongPtr
    Dim sKey  As String
    Dim aKey  As LongPtr
 
    If col Is Nothing Then
        Err.Raise 91
    Else
        If col.Count Then
            'save StrPtr
            aKey = StrPtr(sKey)
            'first Ptr
            Call CopyMemoryByRef(ptr, ByVal (ObjPtr(col) + FOWARD_SHIFT), LONGPTR_SIZE)
            GetCollectionIndexByKey = 1
            
            ' walk items upwards from first
            Do
                Call CopyMemoryByRef(ByVal VarPtr(sKey), ByVal (ptr + KEY_POS), LONGPTR_SIZE)
                If StrComp(Key, sKey, comp) = 0 Then
                    'found
                    Exit Do
                End If
                'next Index
                GetCollectionIndexByKey = GetCollectionIndexByKey + 1
                
                ' 無限ループ対策
                If GetCollectionIndexByKey > col.Count Then
                    GetCollectionIndexByKey = -1
                    Exit Function
                End If
                
                ' next Ptr
                Call CopyMemoryByRef(ptr, ByVal (ptr + FOWARD_SHIFT), LONGPTR_SIZE)
            Loop Until ptr = 0
            
            ' restore original StrPtr
            Call CopyMemoryByRef(ByVal VarPtr(sKey), aKey, LONGPTR_SIZE)
        End If
        If ptr = 0 Then
            'key not found
            GetCollectionIndexByKey = -1
        End If
    End If
End Function

'******************************************************************************
'* [概  要] GetCollectionKeys メソッド
'* [詳  細] コレクションのキーを配列で取得する。
'*
'* @param col コレクション
'* @return キー名
'*
'******************************************************************************
Public Function GetCollectionKeys(col As Collection) As Variant
    Dim ptr As LongPtr
    Dim tmpptr As LongPtr
    Dim sKey As String
    Dim arrex As ArrayEx: Set arrex = New ArrayEx
    
    If col Is Nothing Then
        Err.Raise 91
    Else
        Call CopyMemoryByRef(ptr, ByVal (ObjPtr(col) + FOWARD_SHIFT), LONGPTR_SIZE)
        tmpptr = StrPtr(sKey)
        Call CopyMemoryByRef(ByVal VarPtr(sKey), ByVal (ptr + KEY_POS), LONGPTR_SIZE)
        arrex.Add sKey
        Call CopyMemoryByRef(ByVal VarPtr(sKey), tmpptr, LONGPTR_SIZE)
        
        Dim i As Long
        For i = 2 To col.Count
            Call CopyMemoryByRef(ptr, ByVal (ptr + FOWARD_SHIFT), LONGPTR_SIZE)
            tmpptr = StrPtr(sKey)
            Call CopyMemoryByRef(ByVal VarPtr(sKey), ByVal (ptr + KEY_POS), LONGPTR_SIZE)
            arrex.Add sKey
            Call CopyMemoryByRef(ByVal VarPtr(sKey), tmpptr, LONGPTR_SIZE)
        Next
    End If
    GetCollectionKeys = arrex.ToArray
End Function

'*******************************************************************************
'* [概  要] CollectionToString
'* [詳  細] Collectionを文字列変換する。
'*          VBA-WEB内のVBA-JSONを利用。
'*
'* @param var Collection型データ
'* @return String 結果文字列
'*
'*******************************************************************************
Public Function CollectionToString(ByVal col As Collection) As String
    WebHelpers.JsonOptions.NonEncode = True
    CollectionToString = WebHelpers.ConvertToJson(col)
End Function

'*******************************************************************************
'* [概  要] DictionaryToString
'* [詳  細] Dictionaryを文字列変換する。
'*          VBA-WEB内のVBA-JSONを利用。
'*
'* @param dic Dictionary型データ
'* @return String 結果文字列
'*
'*******************************************************************************
Public Function DictionaryToString(ByVal dic As Object) As String
    WebHelpers.JsonOptions.NonEncode = True
    DictionaryToString = WebHelpers.ConvertToJson(dic)
End Function


'*******************************************************************************
'* [概  要] CollectionSwap
'* [詳  細] Collection内の値の交換を行う。
'* [参  考] https://thom.hateblo.jp/entry/2015/11/29/212934
'*
'* @param col 対象のCollection
'* @param index1 交換対象となるインデックス1
'* @param index2 交換対象となるインデックス2
'*
'*******************************************************************************
Public Sub CollectionSwap(ByVal col As Collection, ByVal index1 As Long, ByVal index2 As Long)
    If index1 = index2 Then Exit Sub

    Dim ind1 As Long, ind2 As Long
    If index1 > index2 Then
        ind1 = index2
        ind2 = index1
    Else
        ind1 = index1
        ind2 = index2
    End If

    Dim Item1 As Variant, Item2 As Variant
    Dim key1 As String, key2 As String

    Call VariantUtils.SetItem(Item1, col.Item(ind1))
    Call VariantUtils.SetItem(Item2, col.Item(ind2))

    key1 = GetCollectionKeyByIndex(ind1, col, True)
    key2 = GetCollectionKeyByIndex(ind2, col, True)
    
    If Abs(ind1 - ind2) = 1 Then
        col.Remove ind2
        col.Remove ind1
        
        If col.Count < ind1 Then
            Dim insNo As Long: insNo = col.Count
            If key1 <> "" Then
                col.Add Item1, key1, After:=insNo
            Else
                col.Add Item1, After:=insNo
            End If
            If key2 <> "" Then
                col.Add Item2, key2, After:=insNo
            Else
                col.Add Item2, After:=insNo
            End If
        Else
            If key1 <> "" Then
                col.Add Item1, key1, Before:=ind1
            Else
                col.Add Item1, Before:=ind1
            End If
            If key2 <> "" Then
                col.Add Item2, key2, Before:=ind1
            Else
                col.Add Item2, Before:=ind1
            End If
        End If
    Else
        col.Remove ind2
        If key2 <> "" Then
            col.Add Item2, key2, After:=ind1
        Else
            col.Add Item2, After:=ind1
        End If
        col.Remove ind1
        If key1 <> "" Then
            col.Add Item1, key1, After:=ind2 - 1
        Else
            col.Add Item1, After:=ind2 - 1
        End If
    End If
End Sub

'*******************************************************************************
'* [概  要] CollectionSort
'* [詳  細] Collectionのソートを行う。標準はバブルソート。
'* [参  考] https://thom.hateblo.jp/entry/2015/11/29/212934
'*
'* @param col 対象のCollection
'* @param fnc ソート比較用の関数名、またはFuncオブジェクト、または
'*        Compare（x1, x2）As Booleanメソッドを持つオブジェクト
'*
'*******************************************************************************
Public Sub CollectionSort(ByVal col As Collection, Optional ByVal fnc, Optional useQuickSort As Boolean = False)
    If useQuickSort Then
        Call QuickSort(col, 1, col.Count, fnc)
        Exit Sub
    End If
    
    Dim i As Long, j As Long
    For i = 1 To col.Count
        For j = col.Count To i Step -1
            If IsMissing(fnc) Then
                If col(i) > col(j) Then
                    CollectionSwap col, i, j
                End If
            ElseIf ValidateUtils.IsFunc(fnc) Then
                If fnc.Apply(col(i), col(j)) Then
                    CollectionSwap col, i, j
                End If
            ElseIf ValidateUtils.IsString(fnc) And fnc <> "" Then
                If Application.Run(fnc, col(i), col(j)) Then
                    CollectionSwap col, i, j
                End If
            ElseIf IsObject(fnc) Then
                If fnc.Compare(col(i), col(j)) Then
                    CollectionSwap col, i, j
                End If
            Else
                If col(i) > col(j) Then
                    CollectionSwap col, i, j
                End If
            End If
        Next
    Next
End Sub


'*******************************************************************************
'* [概  要] QuickSort
'* [詳  細] Collectionのクイックソートを行う。
'*
'* @param col      対象のCollection
'* @param startInd 開始インデックス
'* @param endInd   終了インデックス
'* @param fnc      ソート比較用の関数名、またはFuncオブジェクト、または
'*                 Compare（x1, x2, ByVal flg As Boolean）As Booleanメソッドを
'*                 持つオブジェクト
'*                 第3引数のflgで条件を反転させる。軸要素は反転前、反転後の条件
'*                 のどちらにも含まれないようにする必要がある。
'*
'* [fncの指定関数例]
'*   Function Compare(val1, val2, ByVal flg As Boolean) As Boolean
'*       If flg Then
'*           Compare = val1(2) > val2(2)
'*       Else
'*           Compare = val1(2) < val2(2)
'*       End If
'*   End Function
'*
'*******************************************************************************
Private Sub QuickSort(ByVal col As Collection, ByVal startInd As Long, ByVal endInd As Long, Optional ByVal fnc)
    If startInd >= endInd Then
        Exit Sub
    End If
    
    ' 軸要素の選択
    Dim pivot As Variant
    pivot = col(Int((startInd + endInd) / 2))
    
    ' パーティション分割
    Dim i As Long: i = startInd
    Dim j As Long: j = endInd
    Do
        If IsMissing(fnc) Then
            Do While i <= endInd And col(i) < pivot: i = i + 1: Loop
            Do While j >= startInd And col(j) > pivot: j = j - 1: Loop
            
        ElseIf ValidateUtils.IsFunc(fnc) Then
            Do While i <= endInd And fnc.Apply(col(i), pivot, True): i = i + 1: Loop
            Do While j >= startInd And fnc.Apply(col(j), pivot, False): j = j - 1: Loop
            
        ElseIf ValidateUtils.IsString(fnc) And fnc <> "" Then
            Do While i <= endInd And Application.Run(fnc, col(i), pivot, True): i = i + 1: Loop
            Do While j >= startInd And Application.Run(fnc, col(j), pivot, False): j = j - 1: Loop
            
        ElseIf IsObject(fnc) Then
            Do While i <= endInd And fnc.Compare(col(i), pivot, True): i = i + 1: Loop
            Do While j >= startInd And fnc.Compare(col(j), pivot, False): j = j - 1: Loop
            
        Else
            Do While i <= endInd And col(i) < pivot: i = i + 1: Loop
            Do While j >= startInd And col(j) > pivot: j = j - 1: Loop
            
        End If
        If i >= j Then Exit Do
        
        CollectionSwap col, i, j
        i = i + 1
        j = j - 1
    Loop
    
    ' 再帰実行
    Call QuickSort(col, startInd, i - 1, fnc)
    Call QuickSort(col, j + 1, endInd, fnc)
End Sub
