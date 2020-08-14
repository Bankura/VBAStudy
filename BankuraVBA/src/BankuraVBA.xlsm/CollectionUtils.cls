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
Private mWinApi As WinAPI

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mWinApi = Base.GetWinAPI()
End Sub
Private Sub Class_Terminate()
    Set mWinApi = Nothing
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
    HasKey = (err.Number = 0)
    err.Clear
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
        err.Raise 91
    Else
        Select Case Index
            Case Is > col.count
                err.Raise 9
            Case Is <= col.count / 2
                Call mWinApi.CopyMemoryByRef(ptr, ObjPtr(col) + FOWARD_SHIFT, LONGPTR_SIZE, False, True)
                For i = 2 To Index
                    Call mWinApi.CopyMemoryByRef(ptr, ptr + FOWARD_SHIFT, LONGPTR_SIZE, False, True)
                Next i
            Case Else
                Call mWinApi.CopyMemoryByRef(ptr, ObjPtr(col) + BACKWARD_SHIFT, LONGPTR_SIZE, False, True)
                For i = col.count - 1 To Index Step -1
                    Call mWinApi.CopyMemoryByRef(ptr, ptr + BACKWARD_SHIFT, LONGPTR_SIZE, False, True)
                Next i
        End Select
        i = StrPtr(sKey)
        Call mWinApi.CopyMemoryByRef(VarPtr(sKey), ptr + KEY_POS, LONGPTR_SIZE, True, True)
        
        GetCollectionKeyByIndex = sKey
        Call mWinApi.CopyMemoryByRef(VarPtr(sKey), i, LONGPTR_SIZE, True, False)
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
        err.Raise 91
    Else
        If col.count Then
            'save StrPtr
            aKey = StrPtr(sKey)
            'first Ptr
            Call mWinApi.CopyMemoryByRef(ptr, ObjPtr(col) + FOWARD_SHIFT, LONGPTR_SIZE, False, True)
            GetCollectionIndexByKey = 1
            
            ' walk items upwards from first
            Do
                Call mWinApi.CopyMemoryByRef(VarPtr(sKey), ptr + KEY_POS, LONGPTR_SIZE, True, True)
                If StrComp(Key, sKey, comp) = 0 Then
                    'found
                    Exit Do
                End If
                'next Index
                GetCollectionIndexByKey = GetCollectionIndexByKey + 1
                
                ' 無限ループ対策
                If GetCollectionIndexByKey > col.count Then
                    GetCollectionIndexByKey = -1
                    Exit Function
                End If
                
                ' next Ptr
                Call mWinApi.CopyMemoryByRef(ptr, ptr + FOWARD_SHIFT, LONGPTR_SIZE, False, True)
            Loop Until ptr = 0
            
            ' restore original StrPtr
            Call mWinApi.CopyMemoryByRef(VarPtr(sKey), aKey, LONGPTR_SIZE, True, False)
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
    DictionaryToString = WebHelpers.ConvertToJson(dic)
End Function
