Attribute VB_Name = "ScriptingExTest"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ScriptingExテスト用モジュール
'* [詳  細] ScriptingExテスト用モジュール
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* [概  要] ScriptingExDictionaryのTest。
'* [詳  細]  ScriptingExDictionaryのTest用。
'*
'******************************************************************************
Sub ScriptingExDictionaryTest()
    Dim dic As ScriptingExDictionary
    Set dic = New ScriptingExDictionary
    
    dic.Add "hoge", "fuga"
    Debug.Print dic.Count
    Debug.Print dic.CompareMode
    Debug.Print dic.Exists("hoge")
    Debug.Print dic.Item("hoge")
    
    dic.Key("hoge") = "uge"
    Debug.Print dic.Item("uge")
    
    Dim vItem
    Dim varItem: varItem = dic.Items
    For Each vItem In varItem
        Debug.Print "Item=" & vItem
    Next
    
    Dim vKey
    Dim varKey: varKey = dic.Keys
    For Each vKey In varKey
        Debug.Print "Key=" & vKey
    Next

    'dic.Remove "uge"
    dic.Remove 1

    Dim i As Integer
    dic.Add "キー1", "アイテム1"
    dic.Add "キー2", "アイテム2"
    dic.Add "キー3", "アイテム3"

    For i = 0 To dic.Count - 1
        Debug.Print dic.Items(i)
    Next i

    For i = 0 To dic.Count - 1
        Debug.Print dic.Keys(i)
    Next i

    Dim var As Variant
    For Each var In dic
        Debug.Print var & "," & dic.Item(var)
    Next var
    
    dic.RemoveAll
    Debug.Print "Last:"
    For Each var In dic
        Debug.Print var & "," & dic.Item(var)
    Next var
End Sub

'******************************************************************************
'* [概  要] ScriptingExDriveのTest。
'* [詳  細]  ScriptingExDriveのTest用。
'*
'******************************************************************************
Sub ScriptingExDriveTest()

    Dim fso As New ScriptingExFileSystemObject
    Dim myDrive As ScriptingExDrive
    
    Set myDrive = fso.Drives("C:")
    'Set myDrive = fso.Drives.Item("C:")
    
    Debug.Print "ドライブ名：" & myDrive.DriveLetter
    Debug.Print "ファイルシステム：" & myDrive.FileSystem
    Debug.Print "ルートフォルダ：" & myDrive.RootFolder.Name
    Debug.Print "ドライブのパス：" & myDrive.Path
End Sub
 
'******************************************************************************
'* [概  要] ScriptingExDrivesのTest。
'* [詳  細]  ScriptingExDrivesのTest用。
'*
'******************************************************************************
Sub ScriptingExDrivesTest()

    Dim fso As New ScriptingExFileSystemObject
    Dim myDrives As ScriptingExDrives
    Dim myDrive As ScriptingExDrive

    Dim i As Long
    Dim dtype As String
    
    Set myDrives = fso.Drives

    For Each myDrive In myDrives
        With myDrive
            i = i + 1
            Debug.Print "ドライブ名：" & .DriveLetter
        
            Select Case .DriveType
        
                Case Removable: dtype = "リムーバブルディスク"
                Case Fixed: dtype = "ハードディスク"
                Case Remote: dtype = "ネットワークドライブ"
                Case CDRom: dtype = "CD-ROM"
                Case RamDisk: dtype = "RAMディスク"
                Case Else: dtype = "不明"
        
            End Select
                
            Debug.Print "ドライブ種類：" & dtype
            
            Debug.Print "ドライブの準備：" & IIf(.IsReady, "準備完了", "準備出来ていません")
            
        End With
    
    Next myDrive
End Sub

'******************************************************************************
'* [概  要] ScriptingExEncoderのTest。
'* [詳  細] ScriptingExEncoderのTest用。
'* [参  考] <http://sammaya.jugem.jp/?eid=13>
'*
'******************************************************************************
Sub ScriptingExEncoderTest()
    Dim fso As ScriptingExFileSystemObject
    Dim enc As ScriptingExEncoder
    Dim file As ScriptingExFile
    Dim stream As ScriptingExTextStream
    Dim fileName As String
    Dim oEncFile As ScriptingExTextStream
    Dim oFilesToEncode(1)
    Dim sDest As String
    Dim sFileOut As String, i
    Dim sSourceFile As String
   
    'Set oFilesToEncode = WScript.Arguments
    oFilesToEncode(0) = "C:\develop\data\vbs\test1.vbs"
    oFilesToEncode(1) = "C:\develop\data\vbs\test2.vbs"
    
    
    Set enc = New ScriptingExEncoder
    'For i = 0 To oFilesToEncode.Count - 1
    For i = LBound(oFilesToEncode) To UBound(oFilesToEncode)
        Set fso = New ScriptingExFileSystemObject
        fileName = oFilesToEncode(i)
        Set file = fso.GetFile(fileName)
        Set stream = file.OpenAsTextStream(1)
        sSourceFile = stream.ReadAll
        stream.CloseStream
        
        sDest = enc.EncodeScriptFile(".vbs", sSourceFile, 0, "")
        sFileOut = Left(fileName, Len(fileName) - 3) & "vbe"
        
        Set oEncFile = fso.CreateTextFile(sFileOut)
        oEncFile.WriteText sDest
        oEncFile.CloseStream
    Next
End Sub

'******************************************************************************
'* [概  要] ScriptingExFileのTest。
'* [詳  細] ScriptingExFileのTest用。
'*
'******************************************************************************
Sub ScriptingExFileTest()
    Dim path1 As String: path1 = "C:\develop\data\text\Shift_JIS.txt"
    Dim path2 As String: path2 = "C:\develop\data\text\Shift_JIS2.txt"
    Dim path3 As String: path3 = "C:\develop\data\text\Shift_JIS3.txt"
    Dim file As ScriptingExFile
    Dim file2 As ScriptingExFile
    Dim file3 As ScriptingExFile
    
    With New ScriptingExFileSystemObject
        Set file = .GetFile(path1)
        Debug.Print "Attributes：" & file.Attributes
        Debug.Print "DateCreated：" & file.DateCreated
        Debug.Print "DateLastAccessed：" & file.DateLastAccessed
        Debug.Print "DateLastModified：" & file.DateLastModified
        Debug.Print "Drive：" & file.Drive.DriveLetter
        
        Debug.Print "Name：" & file.Name
        Debug.Print "ParentFolder：" & file.ParentFolder.Path
        Debug.Print "ShortName：" & file.ShortName
        Debug.Print "ShortPath：" & file.ShortPath
        Debug.Print "Size：" & file.Size
        Debug.Print "Type：" & file.Type_
        file.Copy path2
        Set file2 = .GetFile(path2)
        Debug.Print "file2 Name：" & file2.Name
        file2.Move path3
        Set file3 = .GetFile(path3)
        Debug.Print "file3 Name：" & file3.Name
        file3.Delete
    End With

End Sub

'******************************************************************************
'* [概  要] ScriptingExFilesのTest。
'* [詳  細] ScriptingExFilesのTest用。
'*
'******************************************************************************
Sub ScriptingExFilesTest()
    Dim path1 As String: path1 = "C:\develop\data\vbs"
    Dim arr() As String, i As Long
    
    With New ScriptingExFileSystemObject
        Dim files As ScriptingExFiles
        Set files = .GetFolder(path1).files
    
        ReDim arr(1 To files.Count) As String
        Dim file As ScriptingExFile
        For Each file In files
            i = i + 1
            arr(i) = file.Name
        Next file
    End With
    
    For i = LBound(arr) To UBound(arr)
        Debug.Print arr(i)
    Next i
End Sub


'******************************************************************************
'* [概  要] ScriptingExDictionaryTestのCountメソッドのTest。
'* [詳  細] CountメソッドのTest用。
'*
'******************************************************************************
Sub ScriptingExDictionaryTest_Count()
    Dim dic As ScriptingExDictionary
    Set dic = New ScriptingExDictionary
    dic.Add "hoge", "fuga"
    dic.Add "uge", "key"
    Debug.Print dic.Count
End Sub



