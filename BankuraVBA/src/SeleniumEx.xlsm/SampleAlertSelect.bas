Attribute VB_Name = "SampleAlertSelect"
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] Selenium�g�p�T���v���i�A���[�g����j
'* [��  ��] �A���[�g����̃T���v���B
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
'* [�T  �v] SampleAlertSelect
'* [��  ��] �A���[�g����̃e�X�g�֐��B
'* [�Q  �l] https://rabbitfoot.xyz/selenium-alertpush/
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



