Attribute VB_Name = "WScriptExTest"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WScriptEx�e�X�g�p���W���[��
'* [��  ��] WScriptEx�e�X�g�p���W���[��
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* [�T  �v] WScriptExWshCollection ��Test�B
'* [��  ��] WScriptExWshCollection ��Test�p�B
'*
'******************************************************************************
Sub WScriptExWshCollectionTest()
    Dim wsh As New WScriptExWshShell
    Dim col As WScriptExWshCollection

'    Dim i
'    Set col = wsh.SpecialFolders
'    For i = 0 To col.Count - 1
'        Debug.Print col.Item(i)
'    Next

    Dim var
    Set col = wsh.SpecialFolders
    For Each var In col
        Debug.Print var
    Next
End Sub

'******************************************************************************
'* [�T  �v] WScriptExWshEnvironment ��Test�B
'* [��  ��] WScriptExWshEnvironment ��Test�p�B
'*
'******************************************************************************
Sub WScriptExWshEnvironmentTest()
    Dim wsh As New WScriptExWshShell
    Dim obj As WScriptExWshEnvironment
   
    Dim var
    Set obj = wsh.Environment
    For Each var In obj
        Debug.Print var
    Next
End Sub

'******************************************************************************
'* [�T  �v] WScriptExWshExec ��Test�B
'* [��  ��] WScriptExWshExec ��Test�p�B
'*
'******************************************************************************
Sub WScriptExWshExecTest()
    Dim wsh As New WScriptExWshShell
    Dim obj As WScriptExWshExec
    'Dim obj As Object
    Set obj = wsh.Exec("ipconfig.exe")

    Do Until obj.Origin.StdOut.AtEndOfStream
        Dim strLine As String, iColon As String, strAddress As String
        strLine = obj.StdOut.ReadLine
        If InStr(strLine, "IPv4 �A�h���X") <> 0 Then
            iColon = InStr(strLine, ":")
            strAddress = Mid(strLine, iColon + 2)
            WScriptEx.Echo strAddress
        End If
    Loop
End Sub

'******************************************************************************
'* [�T  �v] WScriptExWshNetwork ��Test�B
'* [��  ��] WScriptExWshNetwork ��Test�p�B
'*
'******************************************************************************
Sub WScriptExWshNetworkTest()
    Dim wsh As New WScriptExWshShell
    Dim wshnet As New WScriptExWshNetwork
    
    Debug.Print wshnet.ComputerName
    Debug.Print wshnet.UserDomain
    Debug.Print wshnet.UserName
    'Debug.Print wshnet.Organization
    'Debug.Print wshnet.Site
    'Debug.Print wshnet.UserProfile

End Sub

'******************************************************************************
'* [�T  �v] WScriptExWshShell ��Test�B
'* [��  ��] WScriptExWshShell ��Test�p�B
'*
'******************************************************************************
Sub WScriptExWshShellTest()
    Dim wsh As New WScriptExWshShell
    Debug.Print wsh.Is32BitProcessorForApp
    Debug.Print wsh.Is32BitProcessor


End Sub

'******************************************************************************
'* [�T  �v] WScriptExWshShortcut ��Test�B
'* [��  ��] WScriptExWshShortcut ��Test�p�B
'*
'******************************************************************************
Sub WScriptExWshShortcutTest()
    Dim wsh As New WScriptExWshShell
    Dim obj As WScriptExWshShortcut
End Sub


'******************************************************************************
'* [�T  �v] WScriptExWshURLShortcut ��Test�B
'* [��  ��] WScriptExWshURLShortcut ��Test�p�B
'*
'******************************************************************************
Sub WScriptExWshURLShortcutTest()
    Dim wsh As New WScriptExWshShell
    Dim obj As WScriptExWshURLShortcut
End Sub

'******************************************************************************
'* [�T  �v] WScriptExWshShell �� SpecialFolders ���\�b�h��Test�B
'* [��  ��] WScriptExWshShell ��SpecialFolders ��Test�p�B
'*
'******************************************************************************
Sub WScriptExWshShell_SpecialFoldersTest()
'    Dim WshShell As WScriptExWshShell, strDesktop As String, oShellLink As Object
'    Set WshShell = New WScriptExWshShell
'    strDesktop = WshShell.SpecialFolders("Desktop")
'    Set oShellLink = WshShell.CreateShortcut(strDesktop & "\Shortcut Script.lnk")
'    oShellLink.TargetPath = "C:\develop\data\vbs\test1.vbs"
'    oShellLink.WindowStyle = 1
'    oShellLink.Hotkey = "CTRL+SHIFT+F"
'    oShellLink.IconLocation = "notepad.exe, 0"
'    oShellLink.Description = "Shortcut Script"
'    oShellLink.WorkingDirectory = strDesktop
'    oShellLink.Save
    Dim wsh As New WScriptExWshShell
    Dim col As WScriptExWshCollection
    Dim i
    Set col = wsh.SpecialFolders
    For i = 0 To col.Count - 1
        Debug.Print col.Item(i)
    Next
End Sub

'******************************************************************************
'* [�T  �v] WScriptExWshShell �� Run ���\�b�h��Test�B
'* [��  ��] WScriptExWshShell Run ��Test�p�B
'* [�Q  �l] <https://www.atmarkit.co.jp/ait/articles/0407/08/news101_2.html>
'*
'******************************************************************************
Sub WScriptExWshShell_RunTest()
    Dim wsh As New WScriptExWshShell
    'wsh.Run "C:\WINDOWS\system32\notepad.exe"
    'wsh.Run "%SystemRoot%\notepad.exe C:\develop\data\text\Shift_JIS.txt"
    'wsh.Run "notepad C:\develop\data\text\Shift_JIS.txt"
    
    Dim res As Integer
    res = wsh.Run("fc.exe C:\develop\data\text\Shift_JIS.txt C:\develop\data\text\UTF-8.txt", 2, True)

    If res = 0 Then
        WScriptEx.Echo "test.txt��test2.txt�͓������e�̃t�@�C���ł�"
    ElseIf res = 1 Then
        WScriptEx.Echo "test.txt��test2.txt�͈قȂ���e�̃t�@�C���ł�"
    Else
        WScriptEx.Echo "FC ������ " & res & " ��Ԃ��܂���"
    End If
End Sub

'******************************************************************************
'* [�T  �v] WScriptExWshShell �� Exec ���\�b�h��Test�B
'* [��  ��] WScriptExWshShell Exec ��Test�p�B
'* [�Q  �l] <https://www.atmarkit.co.jp/ait/articles/0407/08/news101_3.html>
'*
'******************************************************************************
Sub WScriptExWshShell_ExecTest()
    Dim wsh As New WScriptExWshShell
    Dim wshExec As WScriptExWshExec
    Set wshExec = wsh.Exec("fc.exe C:\develop\data\text\Shift_JIS.txt C:\develop\data\text\UTF-8.txt")
    
    Do While wshExec.Status = 0
        WScriptEx.Sleep 100
    Loop
    
    Dim res As Integer
    res = wshExec.ExitCode
    If res = 0 Then
        WScriptEx.Echo "test.txt��test2.txt�͓������e�̃t�@�C���ł�"
    ElseIf res = 1 Then
        WScriptEx.Echo "test.txt��test2.txt�͈قȂ���e�̃t�@�C���ł�"
    Else
        WScriptEx.Echo "FC ������ " & res & " ��Ԃ��܂���"
    End If
End Sub

'******************************************************************************
'* [�T  �v] WScriptExWshShell �� SendKeys ���\�b�h��Test�B
'* [��  ��] WScriptExWshShell SendKeys ��Test�p�B
'* [�Q  �l] <https://www.atmarkit.co.jp/ait/articles/0407/08/news101_3.html>
'*
'******************************************************************************
Sub WScriptExWshShell_SendKeysTest()
    Dim wsh As New WScriptExWshShell
    wsh.Run "notepad.exe"
    WScriptEx.Sleep 1000
    wsh.SendKeys "hello"
    WScriptEx.Sleep 1000
    wsh.SendKeys "tallo"
    WScriptEx.Sleep 1000
    wsh.SendKeys "gello"
    WScriptEx.Sleep 1000
    
    wsh.SendKeys "{HOME}"  ' �s���ɖ߂�
    WScriptEx.Sleep 1000
    wsh.SendKeys "^f"      ' Ctrl+F�Ō����_�C�A���O�̕\��
    WScriptEx.Sleep 1000
    
    wsh.SendKeys "tallo~"    ' hello�Ɠ��͂���Enter�L�[������
    wsh.SendKeys "{ESC}"   ' �_�C�A���O�����
    WScriptEx.Sleep 1000
    wsh.SendKeys "{HOME}"  ' �s���ɖ߂�
    WScriptEx.Sleep 1000
    wsh.SendKeys "+{END}"  ' Shift+End �ōs���܂őI��
    WScriptEx.Sleep 1000
    wsh.SendKeys "^c"      ' Ctrl+C�ŃR�s�[
    WScriptEx.Sleep 1000
    wsh.SendKeys "%{F4}"   ' Alt+F4�ŏI��
End Sub

