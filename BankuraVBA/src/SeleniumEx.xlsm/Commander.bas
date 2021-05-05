Attribute VB_Name = "Commander"
Option Explicit
' https://thom.hateblo.jp/entry/2017/01/31/012913

Private fso_ As Object
Private shell_ As Object

Property Get SharedFSO() As Object
    If fso_ Is Nothing Then Set fso_ = CreateObject("Scripting.FileSystemObject")
    Set SharedFSO = fso_
End Property

Property Get SharedWshShell() As Object
    If shell_ Is Nothing Then Set shell_ = CreateObject("WScript.Shell")
    Set SharedWshShell = shell_
End Property

Function GetTempFilePath(Optional create_file As Boolean = False) As String
    Dim ret As String
    ret = Environ$("temp") & "\" & SharedFSO.GetTempName
    If create_file Then
        Call SharedFSO.CreateTextFile(ret)
    End If
    GetTempFilePath = ret
End Function

Function GetCommandResultAsTextStream(command_string, Optional temp_path) As Object
    Dim tempPath As String
    If IsMissing(temp_path) Then
        tempPath = GetTempFilePath
    Else
        tempPath = temp_path
    End If
    Const WshHide = 0
    Const ForReading = 1
    Call SharedWshShell.Run("cmd.exe /c " & command_string & " > " & tempPath, WshHide, True)
    Set GetCommandResultAsTextStream = SharedFSO.OpenTextFile(tempPath, ForReading)
End Function

Function GetCommandResult(command_string) As String
    Dim ret As String
    Dim ts As Object
    Dim tempPath As String: tempPath = GetTempFilePath
    Set ts = GetCommandResultAsTextStream(command_string, tempPath)
    If ts.AtEndOfStream Then
        ret = ""
    Else
        ret = ts.ReadAll
    End If
    ts.Close
    Call SharedFSO.DeleteFile(tempPath, True)
    GetCommandResult = ret
End Function

Function GetCommandResultAsArray(command_string) As String()
    Dim ret() As String
    ret = Split(GetCommandResult(command_string), vbNewLine)
    GetCommandResultAsArray = ret
End Function


Function GetPSCommandResultAsTextStream(command_string, Optional temp_path) As Object
    Dim tempPath As String
    If IsMissing(temp_path) Then
        tempPath = GetTempFilePath
    Else
        tempPath = temp_path
    End If

    Const WshHide = 0
    Const ForReading = 1

    Call SharedWshShell.Run("powershell -ExecutionPolicy RemoteSigned -Command Invoke-Expression """ & command_string & " | Out-File -filePath " & tempPath & " -encoding Default""", WshHide, True)
    Set GetPSCommandResultAsTextStream = SharedFSO.OpenTextFile(tempPath, ForReading)
End Function

Function GetPSCommandResult(command_string) As String
    Dim ret As String
    Dim ts As Object

    Dim tempPath As String: tempPath = GetTempFilePath
    Set ts = GetPSCommandResultAsTextStream(command_string, tempPath)
    If ts.AtEndOfStream Then
        ret = ""
    Else
        ret = ts.ReadAll
    End If
    ts.Close
    Call SharedFSO.DeleteFile(tempPath, True)
    GetPSCommandResult = ret
End Function

Function GetPSCommandResultAsArray(command_string) As String()
    Dim ret() As String
    ret = Split(GetPSCommandResult(command_string), vbNewLine)
    GetPSCommandResultAsArray = ret
End Function
