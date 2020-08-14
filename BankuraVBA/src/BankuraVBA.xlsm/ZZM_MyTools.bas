Attribute VB_Name = "ZZM_MyTools"
Option Explicit

Sub zzzz()
    Dim strReadFilePath As String: strReadFilePath = "C:\dev\vba\aaaaaaa.txt"
    Dim strWriteFilePath As String: strWriteFilePath = "C:\dev\vba\bbbbbbb.txt"
    Call aaaa(strReadFilePath, strWriteFilePath)
End Sub

Sub aaaa(strReadFilePath As String, strWriteFilePath As String)
    Dim buf As String, tmp As Variant

    Open strReadFilePath For Input As #1
    Open strWriteFilePath For Output As #2
        
    
    Do Until EOF(1)
        Line Input #1, buf
        If StringUtils.StartsWith(buf, "    ") Then
            Print #2, StringUtils.ReplaceEach(buf, Array("Optional ", " As Double", " As String", " As Range", " As Boolean"), Array("", "", "", "", ""))
        Else
            Print #2, buf
        End If
    Loop
    
    Close #2
    Close #1

End Sub
