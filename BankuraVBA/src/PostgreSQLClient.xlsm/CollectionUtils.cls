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
'* Copyright (c) 2020 Bankura
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
Public Sub Shift(ByVal clct As Collection, ByVal val As Variant)
    Call Core.Shift(clct, val)
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
'* @return キー名
'*
'******************************************************************************
Public Function GetCollectionKeyByIndex(ByVal Index As Long, col As Collection) As String
    Dim ptr As LongPtr

    Dim i As LongPtr
    Dim sKey As String
    If col Is Nothing Then
        Err.Raise 91
    Else
        Select Case Index
            Case Is > col.Count
                Err.Raise 9
            Case Is <= col.Count / 2
                Call CopyMemoryByRef(ptr, ByVal (ObjPtr(col) + FOWARD_SHIFT), LONGPTR_SIZE)
                For i = 2 To Index
                    Call CopyMemoryByRef(ptr, ByVal (ptr + FOWARD_SHIFT), LONGPTR_SIZE)
                Next i
            Case Else
                Call CopyMemoryByRef(ptr, ByVal (ObjPtr(col) + BACKWARD_SHIFT), LONGPTR_SIZE)
                For i = col.Count - 1 To Index Step -1
                    Call CopyMemoryByRef(ptr, ByVal (ptr + BACKWARD_SHIFT), LONGPTR_SIZE)
                Next i
        End Select
        i = StrPtr(sKey)
        Call CopyMemoryByRef(ByVal VarPtr(sKey), ByVal (ptr + KEY_POS), LONGPTR_SIZE)
        
        GetCollectionKeyByIndex = sKey
        Call CopyMemoryByRef(ByVal VarPtr(sKey), i, LONGPTR_SIZE)
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
'* [概  要] ObjectToString
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
