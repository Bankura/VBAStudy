Attribute VB_Name = "SampleAlertSelect"
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Selenium使用サンプル（アラート操作）
'* [詳  細] アラート操作のサンプル。
'*
'* @author Bankura
'* Copyright (c) 2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] SampleAlertSelect
'* [詳  細] アラート操作のテスト関数。
'* [参  考] https://rabbitfoot.xyz/selenium-alertpush/
'*
'******************************************************************************
Sub SampleAlertSelect()
    Dim driver As SeleniumExDriver
    Set driver = New SeleniumExDriver
    
    Call driver.StartChrome
    driver.GetPage "file:\" & Replace(ThisWorkbook.FullName, ThisWorkbook.name, "") & "\html\Alert_ok_cancel.html"
    Sleep 5000
    
    Dim sAlert As Selenium.Alert
    Set sAlert = driver.SwitchToAlert
    Debug.Print sAlert.text
    
    sAlert.Dismiss
    Sleep 5000
    
    Set sAlert = driver.SwitchToAlert
    Debug.Print sAlert.text
    sAlert.Accept
    
    driver.CloseWindow
    Set driver = Nothing
End Sub



