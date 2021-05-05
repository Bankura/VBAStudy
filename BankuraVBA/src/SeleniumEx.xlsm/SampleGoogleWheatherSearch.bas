Attribute VB_Name = "SampleGoogleWheatherSearch"
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Selenium使用サンプル（Google天気検索）
'* [詳  細] Google天気検索を行いスクリーンショットを撮るサンプル。
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
'* [概  要] SampleGoogleWheatherSearch
'* [詳  細] Google天気検索を行いスクリーンショットを撮るテストメソッド。
'*
'******************************************************************************
Sub SampleGoogleWheatherSearch()
    Dim driver As SeleniumExDriver
    Set driver = New SeleniumExDriver
    
    ' ドライバ/ブラウザ設定
    driver.SetArgumentMaximized
    driver.SetPreferenceForDownload
    driver.SetTimeoutsImplicitWait 10000

    ' Google検索ページ表示
    driver.StartChrome

    ' 自動操作されている表示を消したいがうまくいかない
    driver.SetCapability "excludeSwitches", Array("enable-automation")
    driver.SetCapability "useAutomationExtension", False
    'Call driver.Manage.Capabilities.Add("excludeSwitches", Array("enable-automation"))
    'Call driver.Manage.Capabilities.Add("useAutomationExtension", False)

    driver.ShowCapabilities
    
    driver.GetPage "https://www.google.co.jp/"
    'driver.StartChrome "https://www.google.co.jp/"
    
    ' 検索ワードに天気を指定し検索
    'Dim elm As Selenium.WebElement
    'Set elm = driver.FindElementByName("q")
    'Call driver.InputText(elm, "天気")
    driver.FindElementByName("q").SendKeys "天気"
    driver.FindElementByName("btnK").Submit
    
    ' 検索結果の件数表示
    Debug.Print driver.FindElementById("result-stats").text

    ' 末尾までスクロール
    'Call driver.ScrollEnd

    ' スクショの保存場所を指定
    Dim imagefolder As String: imagefolder = Replace(ThisWorkbook.FullName, ThisWorkbook.name, "") & "\tmp\"
    

    Dim ypos As Long: ypos = driver.GetWindowYPosition()
    Debug.Print ypos
        
    Dim i As Long: i = 1
    Do
        'スクショをフォルダに保存
        driver.TakeScreenshot.SaveAs imagefolder & "test" & i & ".png"
        
        Call driver.ScrollNextPageBySpace
        Sleep 100
        
        If ypos = driver.GetWindowYPosition() Then
            Exit Do
        End If
        
        ypos = driver.GetWindowYPosition()
        Debug.Print ypos
        i = i + 1
    Loop
    driver.CloseWindow
End Sub
