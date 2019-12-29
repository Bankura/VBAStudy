Attribute VB_Name = "RegExpExTest"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] RegExpExテスト用モジュール
'* [詳  細] RegExpExテスト用モジュール
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* [概  要] RegExpExRegExp のTest。
'* [詳  細] RegExpExRegExp のTest用処理。
'* [参  考] <https://vbabeginner.net/vba%E3%81%A7%E6%AD%A3%E8%A6%8F%E8%A1%A8%E7%8F%BE%E3%82%92%E5%88%A9%E7%94%A8%E3%81%99%E3%82%8B/>
'*
'******************************************************************************
Sub RegExpExRegExpTest()
    Dim reg As RegExpExRegExp
    Set reg = New RegExpExRegExp
    Dim s As String

    Dim oMatches As RegExpExMatchCollection
    Dim oMatch As RegExpExMatch
    Dim i As Long
    Dim iCount
    Dim iFirstIndex
    Dim iLength
    Dim sValue
    Dim oSubMatches As RegExpExSubMatches
    Dim iSub As Long
    Dim iSubCount
    
    '// 検索対象文字列
    s = "abcd1234efgh5678ijkl9012"
    
    '// 検索条件設定
    reg.Global_ = True              ' 検索範囲（True：文字列の最後まで検索、False：最初の一致まで検索）
    reg.IgnoreCase = True           ' 大文字小文字の区別（True：区別しない、False：区別する）
    reg.Pattern = "([a-z]+)(\d+)"   ' 検索パターン（ここでは連続する数字を検索条件に設定）
    
    '// 検索実行
    Set oMatches = reg.Execute(s)
    
    '// 検索一致件数を取得
    iCount = oMatches.Count
    
    '// 検索一致件数だけループ
    For i = 0 To iCount - 1
        '// コレクションの現ループオブジェクトを取得
        Set oMatch = oMatches.Item(i)
        
        '// 最初の検索一致位置
        iFirstIndex = oMatch.FirstIndex
        '// 検索一致文字列の長さ
        iLength = oMatch.Length
        '// 検索一致文字列
        sValue = oMatch.Value
        
        Debug.Print "最初検索一致位置：" & iFirstIndex & " 長さ：" & iLength & " 文字列：" & sValue
        
        '// 検索一致
        Set oSubMatches = oMatch.SubMatches
        
        '// サブ表現（丸括弧で囲われている検索条件）されている数を取得
        iSubCount = oSubMatches.Count
        
        '// サブ表現の数だけループ
        For iSub = 0 To iSubCount - 1
            Debug.Print "サブ表現一致文字列：" & oSubMatches.Item(iSub)
        Next
    Next
End Sub


