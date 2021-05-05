Attribute VB_Name = "SampleGoogleWheatherSearch"
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] Selenium�g�p�T���v���iGoogle�V�C�����j
'* [��  ��] Google�V�C�������s���X�N���[���V���b�g���B��T���v���B
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
'* [�T  �v] SampleGoogleWheatherSearch
'* [��  ��] Google�V�C�������s���X�N���[���V���b�g���B��e�X�g���\�b�h�B
'*
'******************************************************************************
Sub SampleGoogleWheatherSearch()
    Dim driver As SeleniumExDriver
    Set driver = New SeleniumExDriver
    
    ' �h���C�o/�u���E�U�ݒ�
    driver.SetArgumentMaximized
    driver.SetPreferenceForDownload
    driver.SetTimeoutsImplicitWait 10000

    ' Google�����y�[�W�\��
    driver.StartChrome

    ' �������삳��Ă���\�����������������܂������Ȃ�
    driver.SetCapability "excludeSwitches", Array("enable-automation")
    driver.SetCapability "useAutomationExtension", False
    'Call driver.Manage.Capabilities.Add("excludeSwitches", Array("enable-automation"))
    'Call driver.Manage.Capabilities.Add("useAutomationExtension", False)

    driver.ShowCapabilities
    
    driver.GetPage "https://www.google.co.jp/"
    'driver.StartChrome "https://www.google.co.jp/"
    
    ' �������[�h�ɓV�C���w�肵����
    'Dim elm As Selenium.WebElement
    'Set elm = driver.FindElementByName("q")
    'Call driver.InputText(elm, "�V�C")
    driver.FindElementByName("q").SendKeys "�V�C"
    driver.FindElementByName("btnK").Submit
    
    ' �������ʂ̌����\��
    Debug.Print driver.FindElementById("result-stats").text

    ' �����܂ŃX�N���[��
    'Call driver.ScrollEnd

    ' �X�N�V���̕ۑ��ꏊ���w��
    Dim imagefolder As String: imagefolder = Replace(ThisWorkbook.FullName, ThisWorkbook.name, "") & "\tmp\"
    

    Dim ypos As Long: ypos = driver.GetWindowYPosition()
    Debug.Print ypos
        
    Dim i As Long: i = 1
    Do
        '�X�N�V�����t�H���_�ɕۑ�
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
