Attribute VB_Name = "ADODBExTest"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] ADODBEx�e�X�g�p���W���[��
'* [��  ��] ADODBEx�e�X�g�p���W���[��
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* [�T  �v] ADODBExConnection��Test�B
'* [��  ��] ADODBExConnection��Test�p�����������B
'*
'******************************************************************************
Sub test()

'    Dim adoCn As ADODB.Connection
'    Dim adoErr As ADODB.Error
'    Dim adoErrs As ADODB.Errors
'    Dim adoPrm As ADODB.Parameter
'    Dim adoPrms As ADODB.Parameters
'    Dim adoPrp As ADODB.Property
'    Dim adoPrps As ADODB.Properties
'    Dim adoRd As ADODB.Record
'    Dim adoRs As ADODB.Recordset
'    Dim adoSt As ADODB.Stream

  '�O����Access�t�@�C�����w�肵�Đڑ�����ꍇ
  Dim adoCn As ADODBExConnection
  Set adoCn = New ADODBExConnection 'ADO�R�l�N�V�����̃C���X�^���X�쐬
  adoCn.OpenCn "Provider=Microsoft.ACE.OLEDB.12.0;" & _
             "Data Source=C:\develop\mydb.accdb;" 'Access�t�@�C�����w��
             
  Dim strSQL As String
  strSQL = "select * from ����ŗ��}�X�^"
  
  '�ǉ��E�X�V�E�폜�̏ꍇ----------------------------------
  'adoCn.Execute strSQL 'SQL�����s
  '--------------------------�ǉ��E�X�V�E�폜�̏ꍇ�����܂�
  
  '�Ǎ��̏ꍇ----------------------------------------------
  Dim adoRs As ADODBExRecordset
  Set adoRs = New ADODBExRecordset 'ADO���R�[�h�Z�b�g�̃C���X�^���X�쐬
  adoRs.OpenRs strSQL, adoCn '���R�[�h���o
  Do Until adoRs.EOF '���o�������R�[�h���I������܂ŏ������J��Ԃ�
   Debug.Print adoRs!�K�p�J�n�� & " " & adoRs!����ŗ� '�t�B�[���h�����o��
   ' Debug.Print adoRs.Fields.Item(0).Name & adoRs.Fields.Item(0).Value
   
   
    adoRs.MoveNext '���̃��R�[�h�Ɉړ�����
  Loop
  adoRs.CloseRs: Set adoRs = Nothing '���R�[�h�Z�b�g�̔j��
  '--------------------------------------�Ǎ��̏ꍇ�����܂�
  
  adoCn.CloseCn: Set adoCn = Nothing '�R�l�N�V�����̔j��


End Sub

'******************************************************************************
'* [�T  �v] ADODBExStream��EncodeUrl���\�b�h��Test�B
'* [��  ��] EncodeUrl���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExStreamTest_EncodeUrl()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    Debug.Print adoSt.EncodeUrl("-._~��1234�񂱂���ABCD%")
End Sub

'******************************************************************************
'* [�T  �v] ADODBExStream��DecodeUrl���\�b�h��Test�B
'* [��  ��] DecodeUrl���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExStreamTest_DecodeUrl()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    Debug.Print adoSt.DecodeUrl("-._~%E3%81%881234%E3%82%93%E3%81%93%E3%81%8A%E3%81%A9ABCD%25")
End Sub

'******************************************************************************
'* [�T  �v] ADODBExStream��EncodeBase64���\�b�h��Test�B
'* [��  ��] EncodeBase64���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExStreamTest_EncodeBase64()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    Debug.Print adoSt.EncodeBase64("���ꂪBase64�G���R�[�h")
End Sub

'******************************************************************************
'* [�T  �v] ADODBExStream��DecodeBase64���\�b�h��Test�B
'* [��  ��] DecodeBase64���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExStreamTest_DecodeBase64()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    Debug.Print adoSt.DecodeBase64("44GT44KM44GMQmFzZTY044Ko44Oz44Kz44O844OJ")
End Sub

'******************************************************************************
'* [�T  �v] ADODBExStream��ReadUTF8TextFile���\�b�h��Test�B
'* [��  ��] ReadUTF8TextFile���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExStreamTest_ReadUTF8TextFile()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    
    Debug.Print adoSt.ReadUTF8TextFile("C:\develop\data\text\UTF-8.txt")
    'Debug.Print adoSt.ReadUTF8TextFile("C:\develop\data\text\UTF-8_Bom.txt")
End Sub

'******************************************************************************
'* [�T  �v] ADODBExStream��ReadTextFile���\�b�h��Test�B
'* [��  ��] ReadTextFile���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExStreamTest_ReadTextFile()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    
    'Debug.Print adoSt.ReadTextFile("C:\develop\data\text\UTF-8.txt", "UTF-8")
    'Debug.Print adoSt.ReadTextFile("C:\develop\data\text\Shift_JIS.txt", "Shift_JIS")
    Debug.Print adoSt.ReadTextFile("C:\develop\data\text\UTF-8_Bom.txt", "UTF-8")
    'Debug.Print adoSt.ReadTextFile("C:\develop\data\text\Shift_JIS.txt")
    
End Sub

'******************************************************************************
'* [�T  �v] ADODBExStream��ReadTextFileToVArray���\�b�h��Test�B
'* [��  ��] ReadTextFileToVArray���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExStreamTest_ReadTextFileToVArray()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    
    Dim vArr As Variant, i As Long
    Debug.Print "�s��: " & adoSt.ReadTextFileLineCount("C:\develop\data\text\UTF-8_Bom.txt", "UTF-8")
    vArr = adoSt.ReadTextFileToVArray("C:\develop\data\text\UTF-8_Bom.txt", "UTF-8")
    
    For i = LBound(vArr) To UBound(vArr)
        Debug.Print i & ": " & vArr(i)
    Next i
End Sub

'******************************************************************************
'* [�T  �v] ADODBExStream��WriteTextFile���\�b�h��Test�B
'* [��  ��] WriteTextFile���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExStreamTest_WriteTextFile()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    Dim data As String
    'Debug.Print adoSt.ReadTextFile("C:\develop\data\text\UTF-8.txt", "UTF-8")
    data = adoSt.ReadTextFile("C:\develop\data\text\Shift_JIS.txt", "Shift_JIS")
    Call adoSt.WriteTextFile("C:\develop\data\text\Write_UTF-8.txt", data, "UTF-8")
    
End Sub

'******************************************************************************
'* [�T  �v] ADODBExStream��WriteTextFileFromVArray���\�b�h��Test�B
'* [��  ��] WriteTextFileFromVArray���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExStreamTest_WriteTextFileFromVArray()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    Dim vArr As Variant, i As Long
    Debug.Print "�s��: " & adoSt.ReadTextFileLineCount("C:\develop\data\text\UTF-8_Bom.txt", "UTF-8")
    vArr = adoSt.ReadTextFileToVArray("C:\develop\data\text\UTF-8_Bom.txt", "UTF-8")
    Call adoSt.WriteTextFileFromVArray("C:\develop\data\text\Write_SJIS.txt", vArr, "Shift_JIS", , True)
    
End Sub


'******************************************************************************
'* [�T  �v] ADODBExStreamWriteUTF8TextFile���\�b�h��Test�
'* [��  ��] WriteUTF8TextFile���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExStreamTest_WriteUTF8TextFile()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    Dim data As String
    'Debug.Print adoSt.ReadTextFile("C:\develop\data\text\UTF-8.txt", "UTF-8")
    data = adoSt.ReadTextFile("C:\develop\data\text\Shift_JIS.txt", "Shift_JIS")
    Call adoSt.WriteUTF8TextFile("C:\develop\data\text\Write_UTF-8NoBom.txt", data)
    Call adoSt.WriteUTF8TextFile("C:\develop\data\text\Write_UTF-8Bom.txt", data, True, True)
End Sub

'******************************************************************************
'* [�T  �v] ADODBExStream��IsBomIncluded���\�b�h��Test�
'* [��  ��] IsBomIncluded���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExStreamTest_IsBomIncluded()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream

    Debug.Print "BOM:" & adoSt.IsBomIncluded("C:\develop\data\text\UTF-8.txt")
    Debug.Print "BOM:" & adoSt.IsBomIncluded("C:\develop\data\text\UTF-8_Bom.txt")
    
End Sub

'******************************************************************************
'* [�T  �v] ADODBExStream��ReadFileToDump���\�b�h��Test�
'* [��  ��] ReadFileToDump���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExStreamTest_��ReadFileToDump()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream

    Debug.Print adoSt.ReadFileToDump("C:\develop\data\text\UTF-8.txt")
    
End Sub


'******************************************************************************
'* [�T  �v] ADODBEx��ChangeFileEncode���\�b�h��Test�
'* [��  ��] ChangeFileEncode���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExTest_ChangeFileEncode()
    'setup:
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    
    Dim fPath As String: fPath = "C:\develop\data\text\work\UTF-8.txt"
    Dim vArr As Variant, i As Long
    
    'when:
    Call ChangeFileEncode(fPath, "UTF-8", "shift_jis")
    
    'check:
    Debug.Print "��Shift_JIS�œǍ�"
    vArr = adoSt.ReadTextFileToVArray(fPath, "shift_jis")
    For i = LBound(vArr) To UBound(vArr)
        Debug.Print i & ": " & vArr(i)
    Next i
    
    'when:
    Call ChangeFileEncode(fPath, "shift_jis", "UTF-8")
  
    'check:
    Debug.Print "��UTF-8�œǍ�"
    vArr = adoSt.ReadTextFileToVArray(fPath, "UTF-8")
    For i = LBound(vArr) To UBound(vArr)
        Debug.Print i & ": " & vArr(i)
    Next i
End Sub
'******************************************************************************
'* [�T  �v] ADODBEx��ChangeFilesEncode���\�b�h��Test�
'* [��  ��] ChangeFileEncode���\�b�h��Test�p�B
'*
'******************************************************************************
Sub ADODBExTest_ChangeFilesEncode()
    Dim fPath As String: fPath = "C:\develop\data\text\work2"
    
    'when:
    Call ChangeFilesEncode(fPath, "UTF-8", "shift_jis")
End Sub

'******************************************************************************
'* [�T  �v] ��e�ʃt�@�C���쐬�p���\�b�h�
'* [��  ��] ��e�ʃt�@�C�����쐬����B
'*
'******************************************************************************
Public Sub CreateTextFileBigData()
                                   
    Dim filePath As String
    filePath = "C:\develop\data\text\bigdata3.txt"
    Dim tmp, i As Long
    With New ADODBExStream
        .Mode = adModeReadWrite
        .Type_ = adTypeText
        .CharSet = "UTF-8"
        .LineSeparator = adCRLF
        .OpenStream
        For i = 0 To 20000
            .WriteText "����͑�e�ʃt�@�C�����쐬���镶�́B,���ꂾ���J��Ԃ��Α�e�ʂƂ����Ă��ߌ��ł͂���܂��B�ԂԂB", adWriteLine
        Next
        .ExcludeBom
        .SaveToFile filePath, adSaveCreateOverWrite '�t�@�C���㏑�w��
        .CloseStream
    End With
End Sub

'******************************************************************************
'* [�T  �v] ADODBEx��ReadAndWrite���\�b�h��Test�
'* [��  ��] ReadAndWrite���\�b�h��Test�p�B
'*
'******************************************************************************
Public Sub ADODBExTest_ReadAndWrite()
    Dim filePath As String
    filePath = "C:\develop\data\text\bigdata3.txt"
    Dim filePath2 As String
    filePath2 = "C:\develop\data\text\bigdata4.txt"
    
    Call ReadAndWrite(filePath, "UTF-8", adCRLF, filePath2, "UTF-8", adLF, "SampleFunc", 4096, False)
    
End Sub

'******************************************************************************
'* [�T  �v] �s�ҏW�p���\�b�h�B
'* [��  ��] ADODBEx��ReadAndWrite���\�b�h��Test�Ŏg�p����s�ҏW�p���\�b�h�
'*
'******************************************************************************
Public Function SampleFunc(rowData As String) As String
    Dim cols, colData, ret As String
    cols = Split(rowData, ",")
    For Each colData In cols
        If ret = "" Then
            ret = colData & "�ŏ�,"
        Else
            ret = ret & colData & "�W�G���h"
        End If
    Next
    ret = Replace(ret, "�ԂԂB", "�ڂ�ڂ�B")
    SampleFunc = ret
End Function

