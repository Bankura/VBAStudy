Attribute VB_Name = "SamplePreCheck"
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] Selenium�g�p�T���v���i�����`�F�b�N�j
'* [��  ��] �����`�F�b�N�����̃T���v���B
'*
'* @author Bankura
'* Copyright (c) 2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI��`
'******************************************************************************
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] SamplePreCheck1
'* [��  ��] �����`�F�b�N�����̃e�X�g�֐��P�B
'* [�Q  �l] https://dampgblog.hinohikari291.com/autodownloadchromedriver/
'*
'******************************************************************************
Sub SamplePreCheck1()
    Dim driver As SeleniumExDriver
    Set driver = New SeleniumExDriver
    
    Debug.Print driver.GetChromeDriverMainVersion
    
    Dim chromeVersion As String: chromeVersion = driver.GetChromeVersion
    Debug.Print chromeVersion
    If driver.IsChromeDriverMatching() Then
        Debug.Print "ChromeDriver�͍ŐV�ł�"
    End If
    
    Dim driverLatestVersion As String: driverLatestVersion = driver.GetChromeDriverLatestVersion(chromeVersion)
    Debug.Print driverLatestVersion

    driver.Quit
End Sub

'******************************************************************************
'* [�T  �v] SamplePreCheck2
'* [��  ��] �����`�F�b�N�����̃e�X�g�֐��Q�B
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
    'driver.ShowElements "Google ����", SearchElemTypeEnum.SEARCH_ATTRIBUTE, 10, "value"
    driver.ShowElements "a", , 10, True
    
    driver.ShowTagCounts
    
    driver.CloseWindow
    driver.Quit

End Sub
