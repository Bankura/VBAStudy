Attribute VB_Name = "SamplePreCheck"
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Selenium使用サンプル（初期チェック）
'* [詳  細] 初期チェック処理のサンプル。
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
'* [概  要] SamplePreCheck1
'* [詳  細] 初期チェック処理のテスト関数１。
'* [参  考] https://dampgblog.hinohikari291.com/autodownloadchromedriver/
'*
'******************************************************************************
Sub SamplePreCheck1()
    Dim driver As SeleniumExDriver
    Set driver = New SeleniumExDriver
    
    Debug.Print driver.GetChromeDriverMainVersion
    
    Dim chromeVersion As String: chromeVersion = driver.GetChromeVersion
    Debug.Print chromeVersion
    If driver.IsChromeDriverMatching() Then
        Debug.Print "ChromeDriverは最新です"
    End If
    
    Dim driverLatestVersion As String: driverLatestVersion = driver.GetChromeDriverLatestVersion(chromeVersion)
    Debug.Print driverLatestVersion

    driver.Quit
End Sub

'******************************************************************************
'* [概  要] SamplePreCheck2
'* [詳  細] 初期チェック処理のテスト関数２。
'*
'******************************************************************************
Sub SamplePreCheck2()
    Dim driver As SeleniumExDriver
    Set driver = New SeleniumExDriver
    
    If Not driver.ChromePreCheck Then
        Debug.Print "Oh No..."
        Exit Sub
    End If

    driver.SetPreferenceForDownload
    driver.StartChrome
    driver.GetPage "https://www.google.co.jp/"
    
    driver.ShowPerformanceTiming
    'driver.ShowCapabilities
    'driver.ShowElements "Google 検索", SearchElemTypeEnum.SEARCH_ATTRIBUTE, 10, "value"
    driver.ShowElements "a", , 10, True
    
    driver.ShowTagCounts
    
    driver.CloseWindow
    driver.Quit

End Sub
