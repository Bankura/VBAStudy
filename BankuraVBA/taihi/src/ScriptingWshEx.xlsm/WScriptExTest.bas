Attribute VB_Name = "WScriptExTest"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WScriptExテスト用モジュール
'* [詳  細] WScriptExテスト用モジュール
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* [概  要] WScriptExWshCollection のTest。
'* [詳  細] WScriptExWshCollection のTest用。
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
'* [概  要] WScriptExWshEnvironment のTest。
'* [詳  細] WScriptExWshEnvironment のTest用。
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
'* [概  要] WScriptExWshExec のTest。
'* [詳  細] WScriptExWshExec のTest用。
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
        If InStr(strLine, "IPv4 アドレス") <> 0 Then
            iColon = InStr(strLine, ":")
            strAddress = Mid(strLine, iColon + 2)
            WScriptEx.Echo strAddress
        End If
    Loop
End Sub

'******************************************************************************
'* [概  要] WScriptExWshNetwork のTest。
'* [詳  細] WScriptExWshNetwork のTest用。
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
'* [概  要] WScriptExWshShell のTest。
'* [詳  細] WScriptExWshShell のTest用。
'*
'******************************************************************************
Sub WScriptExWshShellTest()
    Dim wsh As New WScriptExWshShell
    Debug.Print wsh.Is32BitProcessorForApp
    Debug.Print wsh.Is32BitProcessor


End Sub

'******************************************************************************
'* [概  要] WScriptExWshShortcut のTest。
'* [詳  細] WScriptExWshShortcut のTest用。
'*
'******************************************************************************
Sub WScriptExWshShortcutTest()
    Dim wsh As New WScriptExWshShell
    Dim obj As WScriptExWshShortcut
End Sub


'******************************************************************************
'* [概  要] WScriptExWshURLShortcut のTest。
'* [詳  細] WScriptExWshURLShortcut のTest用。
'*
'******************************************************************************
Sub WScriptExWshURLShortcutTest()
    Dim wsh As New WScriptExWshShell
    Dim obj As WScriptExWshURLShortcut
End Sub

'******************************************************************************
'* [概  要] WScriptExWshShell の SpecialFolders メソッドのTest。
'* [詳  細] WScriptExWshShell のSpecialFolders のTest用。
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
'* [概  要] WScriptExWshShell の Run メソッドのTest。
'* [詳  細] WScriptExWshShell Run のTest用。
'* [参  考] <https://www.atmarkit.co.jp/ait/articles/0407/08/news101_2.html>
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
        WScriptEx.Echo "test.txtとtest2.txtは同じ内容のファイルです"
    ElseIf res = 1 Then
        WScriptEx.Echo "test.txtとtest2.txtは異なる内容のファイルです"
    Else
        WScriptEx.Echo "FC が結果 " & res & " を返しました"
    End If
End Sub

'******************************************************************************
'* [概  要] WScriptExWshShell の Exec メソッドのTest。
'* [詳  細] WScriptExWshShell Exec のTest用。
'* [参  考] <https://www.atmarkit.co.jp/ait/articles/0407/08/news101_3.html>
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
        WScriptEx.Echo "test.txtとtest2.txtは同じ内容のファイルです"
    ElseIf res = 1 Then
        WScriptEx.Echo "test.txtとtest2.txtは異なる内容のファイルです"
    Else
        WScriptEx.Echo "FC が結果 " & res & " を返しました"
    End If
End Sub

'******************************************************************************
'* [概  要] WScriptExWshShell の SendKeys メソッドのTest。
'* [詳  細] WScriptExWshShell SendKeys のTest用。
'* [参  考] <https://www.atmarkit.co.jp/ait/articles/0407/08/news101_3.html>
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
    
    wsh.SendKeys "{HOME}"  ' 行頭に戻る
    WScriptEx.Sleep 1000
    wsh.SendKeys "^f"      ' Ctrl+Fで検索ダイアログの表示
    WScriptEx.Sleep 1000
    
    wsh.SendKeys "tallo~"    ' helloと入力してEnterキーを押す
    wsh.SendKeys "{ESC}"   ' ダイアログを閉じる
    WScriptEx.Sleep 1000
    wsh.SendKeys "{HOME}"  ' 行頭に戻る
    WScriptEx.Sleep 1000
    wsh.SendKeys "+{END}"  ' Shift+End で行末まで選択
    WScriptEx.Sleep 1000
    wsh.SendKeys "^c"      ' Ctrl+Cでコピー
    WScriptEx.Sleep 1000
    wsh.SendKeys "%{F4}"   ' Alt+F4で終了
End Sub

