Attribute VB_Name = "ScriptingExTest"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] ScriptingEx�e�X�g�p���W���[��
'* [��  ��] ScriptingEx�e�X�g�p���W���[��
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* [�T  �v] ScriptingExDictionary��Test�B
'* [��  ��]  ScriptingExDictionary��Test�p�B
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
    dic.Add "�L�[1", "�A�C�e��1"
    dic.Add "�L�[2", "�A�C�e��2"
    dic.Add "�L�[3", "�A�C�e��3"

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
'* [�T  �v] ScriptingExDrive��Test�B
'* [��  ��]  ScriptingExDrive��Test�p�B
'*
'******************************************************************************
Sub ScriptingExDriveTest()

    Dim fso As New ScriptingExFileSystemObject
    Dim myDrive As ScriptingExDrive
    
    Set myDrive = fso.Drives("C:")
    'Set myDrive = fso.Drives.Item("C:")
    
    Debug.Print "�h���C�u���F" & myDrive.DriveLetter
    Debug.Print "�t�@�C���V�X�e���F" & myDrive.FileSystem
    Debug.Print "���[�g�t�H���_�F" & myDrive.RootFolder.Name
    Debug.Print "�h���C�u�̃p�X�F" & myDrive.Path
End Sub
 
'******************************************************************************
'* [�T  �v] ScriptingExDrives��Test�B
'* [��  ��]  ScriptingExDrives��Test�p�B
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
            Debug.Print "�h���C�u���F" & .DriveLetter
        
            Select Case .DriveType
        
                Case Removable: dtype = "�����[�o�u���f�B�X�N"
                Case Fixed: dtype = "�n�[�h�f�B�X�N"
                Case Remote: dtype = "�l�b�g���[�N�h���C�u"
                Case CDRom: dtype = "CD-ROM"
                Case RamDisk: dtype = "RAM�f�B�X�N"
                Case Else: dtype = "�s��"
        
            End Select
                
            Debug.Print "�h���C�u��ށF" & dtype
            
            Debug.Print "�h���C�u�̏����F" & IIf(.IsReady, "��������", "�����o���Ă��܂���")
            
        End With
    
    Next myDrive
End Sub

'******************************************************************************
'* [�T  �v] ScriptingExEncoder��Test�B
'* [��  ��] ScriptingExEncoder��Test�p�B
'* [�Q  �l] <http://sammaya.jugem.jp/?eid=13>
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
'* [�T  �v] ScriptingExFile��Test�B
'* [��  ��] ScriptingExFile��Test�p�B
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
        Debug.Print "Attributes�F" & file.Attributes
        Debug.Print "DateCreated�F" & file.DateCreated
        Debug.Print "DateLastAccessed�F" & file.DateLastAccessed
        Debug.Print "DateLastModified�F" & file.DateLastModified
        Debug.Print "Drive�F" & file.Drive.DriveLetter
        
        Debug.Print "Name�F" & file.Name
        Debug.Print "ParentFolder�F" & file.ParentFolder.Path
        Debug.Print "ShortName�F" & file.ShortName
        Debug.Print "ShortPath�F" & file.ShortPath
        Debug.Print "Size�F" & file.Size
        Debug.Print "Type�F" & file.Type_
        file.Copy path2
        Set file2 = .GetFile(path2)
        Debug.Print "file2 Name�F" & file2.Name
        file2.Move path3
        Set file3 = .GetFile(path3)
        Debug.Print "file3 Name�F" & file3.Name
        file3.Delete
    End With

End Sub

'******************************************************************************
'* [�T  �v] ScriptingExFiles��Test�B
'* [��  ��] ScriptingExFiles��Test�p�B
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
'* [�T  �v] ScriptingExDictionaryTest��Count���\�b�h��Test�B
'* [��  ��] Count���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ScriptingExDictionaryTest_Count()
    Dim dic As ScriptingExDictionary
    Set dic = New ScriptingExDictionary
    dic.Add "hoge", "fuga"
    dic.Add "uge", "key"
    Debug.Print dic.Count
End Sub



