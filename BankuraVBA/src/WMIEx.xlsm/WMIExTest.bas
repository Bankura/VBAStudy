Attribute VB_Name = "WMIExTest"
Option Explicit
#Const USE_REFERENCE = False    '参照設定使用有無

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WMIExテスト用モジュール
'* [詳  細] WMIExテスト用モジュール
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* [概  要] WMIEx のTest。
'* [詳  細] WMIEx のTest用処理。
'* [参  考] <>
'*
'******************************************************************************
Sub WMIExTest()
    Dim objWMIService As New WMIExSWbemServicesEx
    Set objWMIService = objWMIService.CreateInstance()

End Sub

'******************************************************************************
'* [概  要] WMIExSWbemDateTime のTest。
'* [詳  細] WMIExSWbemDateTime のTest用処理。
'* [参  考] <https://selifelog.com/blog-entry-310.html>
'*
'******************************************************************************
Sub WMIExSWbemDateTimeTest()
    Dim objDateTime As WMIExSWbemDateTime
    Dim myStartDate As Date
    myStartDate = CDate("2015/07/31 09:00:00")
    
    Set objDateTime = New WMIExSWbemDateTime
    'objDateTime.UTC = 0
    objDateTime.SetVarDate myStartDate, True
    Debug.Print objDateTime

    Dim utcDate As String
    utcDate = objDateTime.GetUTCDate("2015/07/31 11:50")
    Debug.Print utcDate
    Debug.Print objDateTime.UTCtoJST(utcDate)
    
    utcDate = objDateTime.GetUTCDate(Now())
    Debug.Print utcDate
    
    Dim temp As Date
    temp = objDateTime.UTCtoJST("20150731055917.000000-000")
    Debug.Print temp
    temp = objDateTime.UTCtoJST(utcDate)
    Debug.Print temp
End Sub


'******************************************************************************
'* [概  要] WMIExSWbemLocator のTest。
'* [詳  細] WMIExSWbemLocator のTest用処理。
'* [参  考] <https://www.bnote.net/windows/wsh/wmi03_locator.shtml>
'*
'******************************************************************************
Sub WMIExSWbemLocatorTest()
    Dim objLocator As WMIExSWbemLocator
    Dim objService As WMIExSWbemServicesEx
    Dim objClasses As WMIExSWbemObjectSet
    Dim objClass As WMIExSWbemObject
    
    Set objLocator = New WMIExSWbemLocator
    Set objService = objLocator.ConnectServer()
    
    objService.Security_.ImpersonationLevel = wbemAuthenticationLevelCall '3
    
    Set objClasses = objService.InstancesOf("Win32_LogicalDisk")
    For Each objClass In objClasses
        Debug.Print objClass.GetObjectText_
        Debug.Print
            
        Dim objProps As WMIExSWbemPropertySet
        Dim objProp As WMIExSWbemProperty
        
        'Property表示
        Debug.Print "[Properties]"
        Set objProps = objClass.Properties_
        For Each objProp In objProps
            Debug.Print objProp.Name & " : " & objProp.Value
        Next
        Debug.Print
        
        Dim objMethods As WMIExSWbemMethodSet
        Dim objMethod As WMIExSWbemMethod
        
        'Method表示
        Debug.Print "[Methods]"
        Set objMethods = objClass.Methods_
        For Each objMethod In objMethods
            Debug.Print objMethod.Name & " : " & objMethod.Origin
        Next
        Debug.Print
    Next
End Sub

'******************************************************************************
'* [概  要] WMIExSWbemServicesEx の ExecQuery のTest。
'* [詳  細] WMIExSWbemServicesEx の ExecQuery  のTest用処理。
'* [参  考] <https://www.bnote.net/windows/wmi.html>
'*
'******************************************************************************
Sub WMIExSWbemServicesEx_ExecQueryTest()
    Dim objService As New WMIExSWbemServicesEx
    Dim objClasses As WMIExSWbemObjectSet
    Dim objClass As WMIExSWbemObject
    
    Set objService = objService.CreateInstance()
    Set objClasses = objService.ExecQuery("Select * from Win32_BIOS")

    For Each objClass In objClasses
        Debug.Print "Build Number         : " & objClass.Properties_("BuildNumber")
        Debug.Print "Current Language     : " & objClass.Properties_("CurrentLanguage")
        Debug.Print "Installable Languages: " & objClass.Properties_("InstallableLanguages")
        Debug.Print "Manufacturer         : " & objClass.Properties_("Manufacturer")
        Debug.Print "Name                 : " & objClass.Properties_("Name")
        Debug.Print "Primary BIOS         : " & objClass.Properties_("PrimaryBIOS")
        Debug.Print "Serial Number        : " & objClass.Properties_("SerialNumber")
        Debug.Print "SMBIOS Version       : " & objClass.Properties_("SMBIOSBIOSVersion")
        Debug.Print "SMBIOS Major Version : " & objClass.Properties_("SMBIOSMajorVersion")
        Debug.Print "SMBIOS Minor Version : " & objClass.Properties_("SMBIOSMinorVersion")
        Debug.Print "SMBIOS Present       : " & objClass.Properties_("SMBIOSPresent")
        Debug.Print "Status               : " & objClass.Properties_("Status")
        
        Dim objProps As WMIExSWbemPropertySet
        Dim objProp As WMIExSWbemProperty
        
        'Property表示
        Debug.Print
        Debug.Print "[Properties]"
        Set objProps = objClass.Properties_
        For Each objProp In objProps
            If IsObject(objProp.Value) Then
                Debug.Print objProp.Name & " : " & TypeName(objProp.Value)
            ElseIf IsArray(objProp.Value) Then
                Debug.Print objProp.Name & " : " & TypeName(objProp.Value)
            Else
                Debug.Print objProp.Name & " : " & objProp.Value
            End If

        Next
        Debug.Print
    Next

End Sub




'******************************************************************************
'* [概  要] WMIExSWbemSinkTest のTest。
'* [詳  細] WMIExSWbemSinkTest のTest用処理。
'* [参  考] <>
'*
'******************************************************************************
#If USE_REFERENCE Then
Sub WMIExSWbemSinkTest()
    Dim objSink As New WMIExSWbemSink
    Set objSink = objSink.CreateInstance()
    Call objSink.ExecNTLogEvent
End Sub
#End If
