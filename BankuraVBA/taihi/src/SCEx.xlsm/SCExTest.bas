Attribute VB_Name = "SCExTest"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] SCExテスト用モジュール
'* [詳  細] SCExテスト用モジュール
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* [概  要] SCEx のTest。
'* [詳  細] SCEx のTest用処理。
'*
'******************************************************************************
Sub SCExTest()
    'https://oe-roel.hatenadiary.org/entry/20080805/1217952672
    Dim sc As New SCExScriptControl
    
    
    Dim tmpStr As String: tmpStr = "abc 123"
    Debug.Print "EncodeURI(str):" & sc.EncodeURI(tmpStr)
    Debug.Print "DecodeURI(encodeURI(str)):" & sc.DecodeURI(sc.EncodeURI(tmpStr))
    
    sc.Reset
    sc.Language = "JScript"
    sc.AddCode "function hoge(){return (arguments.length);}"
    Debug.Print sc.Run("hoge", 1, 2, 3) ' Prints "3"

    sc.Reset
    sc.AddCode "function hoge2(x){return x.Name;}"
    Debug.Print sc.Run("hoge2", ThisWorkbook) ' Workbook.Nameが返却される
    
    Debug.Print sc.AllowUI
    Debug.Print sc.Language
    Debug.Print sc.SitehWnd
    Debug.Print sc.State
    Debug.Print sc.Timeout
    Debug.Print sc.UseSafeSubset
    Debug.Print sc.Modules.Count
    Debug.Print sc.Procedures.Count


    Set sc = Nothing
End Sub


