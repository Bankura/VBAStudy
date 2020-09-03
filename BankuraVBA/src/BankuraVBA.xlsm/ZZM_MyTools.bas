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

Sub Shapes_Delete() 'アクティブシート上にある全ての
                    'オートシェイプや画像オブジェクトを消去する
    Dim objShape As Shape
    With Application
        .ScreenUpdating = False
        For Each objShape In ActiveSheet.Shapes
            objShape.Delete
        Next
        .ScreenUpdating = True
    End With
End Sub

Sub Images_Delete() 'アクティブシート上にある全ての画像オブジェクトを消去する
    Dim objShape As Shape
    With Application
        .ScreenUpdating = False
        For Each objShape In ActiveSheet.Shapes
            If objShape.Type = msoPicture Then objShape.Delete
        Next
        .ScreenUpdating = True
    End With
End Sub

Sub Add_Image_Name() 'アクティブシート上にある全ての画像に連番を振る
    Const conName As String = "Image_" '画像の名前。名前の変更はココ
    Dim c As Long
    Dim objShape As Object
    For Each objShape In ActiveSheet.Shapes
        If objShape.Type = msoPicture Then
            c = c + 1
            objShape.name = conName & c
        End If
    Next
End Sub


