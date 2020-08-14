VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFilterReconstructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'clsFilterReconstructure
'init で渡したワークシートにあるオートフィルタのフィルタ条件を
'記憶して、同じ条件を再設定する。
'http://yamav102.cocolog-nifty.com/blog/2018/02/post-82ac.html

Private mws As Worksheet

'// それぞれの列のオートフィルタ条件を記憶する変数
Private mCriterias() As Collection

'初期化メソッド
Friend Sub Init(newWs As Worksheet)
    Set mws = newWs
End Sub

'オートフィルタの条件を記憶します。
Public Sub StoreCriterias()

    '// オートフィルタの無いシートなら何もしない
    If mws.AutoFilter Is Nothing Then Exit Sub
    
    Dim fltr As AutoFilter
    Dim Clms As Filters
    Dim operator As XlAutoFilterOperator
    
    Set fltr = mws.AutoFilter
    Set Clms = fltr.Filters
    
    '//　フィルタが掛かっていたら
    If fltr.FilterMode Then
        Dim i As Long
        
        ReDim mCriterias(1 To Clms.count)
        
        '// 全カラムのフィルタ条件を調査
        For i = 1 To Clms.count
        
            Set mCriterias(i) = New Collection
            
            '// フィルタ条件が設定されている
            If Clms.Item(i).On Then
            
                '// 演算子
                mCriterias(i).Add Clms.Item(i).operator, "operator"
                
                '// 第1条件
                If Clms.Item(i).count > 0 Then
                    mCriterias(i).Add Clms.Item(i).Criteria1, "Criteria1"
                End If
                
                '//operatorが、xlFilterValues(7) 以外で、条件が2つ以上の場合は第2条件があるので取得
                If 7 <> Clms.Item(i).operator And Clms.Item(i).count > 1 Then
                    mCriterias(i).Add Clms.Item(i).Criteria2, "Criteria2"
                End If
            
            End If '// If Clms.Item(i).On Then
        
        Next '// For i = 1 To Clms.Count
    
    End If '// If fltr.FilterMode Then
    
End Sub
    
'フィルタを解除して全行表示します
Public Sub ShowAllData()
    '// オートフィルタの無いシートなら何もしない
    If mws.AutoFilter Is Nothing Then Exit Sub
    
    If mws.FilterMode Then
        mws.ShowAllData
    End If
End Sub
    
'フィルタ条件の復元
Public Sub ReConstructFilter()
    
    '// オートフィルタの無いシートなら何もしない
    If mws.AutoFilter Is Nothing Then Exit Sub
    
    '// フィルタ条件が無ければ何もしない
    If IsEmptyArray(mCriterias) Then Exit Sub
    
    Dim i As Long
    Dim rng As Range
    
    Set rng = mws.AutoFilter.Range
    
    '// フィルタ領域の全ての列の条件を復元
    For i = 1 To mws.AutoFilter.Filters.count
        
        With mCriterias(i)
            ' 第2条件がある場合
            If HasKey(mCriterias(i), "Criteria2") Then
                rng.AutoFilter Field:=i, _
                Criteria1:=.Item("Criteria1"), _
                Criteria2:=.Item("Criteria2"), _
                operator:=.Item("operator")
            ' 第1条件のみの場合
            ElseIf HasKey(mCriterias(i), "Criteria1") Then
                rng.AutoFilter Field:=i, _
                Criteria1:=.Item("Criteria1"), _
                operator:=.Item("operator")
            End If
        End With
    Next
End Sub

'// 空の配列を渡すと True を返す
Private Function IsEmptyArray(arr) As Boolean
    Dim errNum As Long
    Dim itemCnt As Long
    
    On Error Resume Next
    
    itemCnt = UBound(arr)
    errNum = err.Number
    
    On Error GoTo 0
    
    'collection の空の配列の ubound はエラーにならない。-1　が返る。
    If errNum <> 0 Or itemCnt = -1 Then
        IsEmptyArray = True
    Else
        IsEmptyArray = False
    End If

End Function

'//辞書にキーがあるかチェック
Private Function HasKey(Coll As Collection, strKey As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = Coll(strKey)
    HasKey = (err.Number = 0)
    err.Clear
End Function

