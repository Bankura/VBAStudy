Attribute VB_Name = "GoogleSearchMain"
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Selenium使用Google検索モジュール
'* [詳  細] Google検索を行うモジュール。
'*
'* @author Bankura
'* Copyright (c) 2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'******************************************************************************
'* 変数定義
'******************************************************************************

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] GoogleSearchMain メソッド
'* [詳  細] テストメソッド。
'******************************************************************************
Sub GoogleSearchMain()
    Dim gSearch As New GoogleSearch
    
    ' 事前チェック実行
    If Not gSearch.WebDriver.ChromePreCheck() Then
        Debug.Print "事前チェック失敗"
        Exit Sub
    End If
    
    ' 検索設定
    gSearch.SearchWord = "本好きの下剋上～司書になるためには手段を選んでいられません～ 兵士の娘I"
    gSearch.GoogleSearchType = GSEARCH_BOOK
    gSearch.UseFilter = True
    gSearch.MaxSearchCount = 20
    
    ' Chrome起動
    gSearch.Go

    ' 検索と表示
    gSearch.SearchAndShow
    
    'gSearch.WebDriver.CloseWindow
End Sub
