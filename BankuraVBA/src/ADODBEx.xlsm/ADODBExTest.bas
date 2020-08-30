Attribute VB_Name = "ADODBExTest"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ADODBExテスト用モジュール
'* [詳  細] ADODBExテスト用モジュール
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* [概  要] ADODBExConnectionのTest。
'* [詳  細] ADODBExConnectionのTest用お試し処理。
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

  '外部のAccessファイルを指定して接続する場合
  Dim adoCn As ADODBExConnection
  Set adoCn = New ADODBExConnection 'ADOコネクションのインスタンス作成
  adoCn.OpenCn "Provider=Microsoft.ACE.OLEDB.12.0;" & _
             "Data Source=C:\develop\mydb.accdb;" 'Accessファイルを指定
             
  Dim strSQL As String
  strSQL = "select * from 消費税率マスタ"
  
  '追加・更新・削除の場合----------------------------------
  'adoCn.Execute strSQL 'SQLを実行
  '--------------------------追加・更新・削除の場合ここまで
  
  '読込の場合----------------------------------------------
  Dim adoRs As ADODBExRecordset
  Set adoRs = New ADODBExRecordset 'ADOレコードセットのインスタンス作成
  adoRs.OpenRs strSQL, adoCn 'レコード抽出
  Do Until adoRs.EOF '抽出したレコードが終了するまで処理を繰り返す
   Debug.Print adoRs!適用開始日 & " " & adoRs!消費税率 'フィールドを取り出す
   ' Debug.Print adoRs.Fields.Item(0).Name & adoRs.Fields.Item(0).Value
   
   
    adoRs.MoveNext '次のレコードに移動する
  Loop
  adoRs.CloseRs: Set adoRs = Nothing 'レコードセットの破棄
  '--------------------------------------読込の場合ここまで
  
  adoCn.CloseCn: Set adoCn = Nothing 'コネクションの破棄


End Sub

'******************************************************************************
'* [概  要] ADODBExStreamのEncodeUrlメソッドのTest。
'* [詳  細] EncodeUrlメソッドのTest用。
'*
'******************************************************************************
Sub ADODBExStreamTest_EncodeUrl()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    Debug.Print adoSt.EncodeUrl("-._~え1234んこおどABCD%")
End Sub

'******************************************************************************
'* [概  要] ADODBExStreamのDecodeUrlメソッドのTest。
'* [詳  細] DecodeUrlメソッドのTest用。
'*
'******************************************************************************
Sub ADODBExStreamTest_DecodeUrl()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    Debug.Print adoSt.DecodeUrl("-._~%E3%81%881234%E3%82%93%E3%81%93%E3%81%8A%E3%81%A9ABCD%25")
End Sub

'******************************************************************************
'* [概  要] ADODBExStreamのEncodeBase64メソッドのTest。
'* [詳  細] EncodeBase64メソッドのTest用。
'*
'******************************************************************************
Sub ADODBExStreamTest_EncodeBase64()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    Debug.Print adoSt.EncodeBase64("これがBase64エンコード")
End Sub

'******************************************************************************
'* [概  要] ADODBExStreamのDecodeBase64メソッドのTest。
'* [詳  細] DecodeBase64メソッドのTest用。
'*
'******************************************************************************
Sub ADODBExStreamTest_DecodeBase64()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    Debug.Print adoSt.DecodeBase64("44GT44KM44GMQmFzZTY044Ko44Oz44Kz44O844OJ")
End Sub

'******************************************************************************
'* [概  要] ADODBExStreamのReadUTF8TextFileメソッドのTest。
'* [詳  細] ReadUTF8TextFileメソッドのTest用。
'*
'******************************************************************************
Sub ADODBExStreamTest_ReadUTF8TextFile()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    
    Debug.Print adoSt.ReadUTF8TextFile("C:\develop\data\text\UTF-8.txt")
    'Debug.Print adoSt.ReadUTF8TextFile("C:\develop\data\text\UTF-8_Bom.txt")
End Sub

'******************************************************************************
'* [概  要] ADODBExStreamのReadTextFileメソッドのTest。
'* [詳  細] ReadTextFileメソッドのTest用。
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
'* [概  要] ADODBExStreamのReadTextFileToVArrayメソッドのTest。
'* [詳  細] ReadTextFileToVArrayメソッドのTest用。
'*
'******************************************************************************
Sub ADODBExStreamTest_ReadTextFileToVArray()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    
    Dim vArr As Variant, i As Long
    Debug.Print "行数: " & adoSt.ReadTextFileLineCount("C:\develop\data\text\UTF-8_Bom.txt", "UTF-8")
    vArr = adoSt.ReadTextFileToVArray("C:\develop\data\text\UTF-8_Bom.txt", "UTF-8")
    
    For i = LBound(vArr) To UBound(vArr)
        Debug.Print i & ": " & vArr(i)
    Next i
End Sub

'******************************************************************************
'* [概  要] ADODBExStreamのWriteTextFileメソッドのTest。
'* [詳  細] WriteTextFileメソッドのTest用。
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
'* [概  要] ADODBExStreamのWriteTextFileFromVArrayメソッドのTest。
'* [詳  細] WriteTextFileFromVArrayメソッドのTest用。
'*
'******************************************************************************
Sub ADODBExStreamTest_WriteTextFileFromVArray()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream
    Dim vArr As Variant, i As Long
    Debug.Print "行数: " & adoSt.ReadTextFileLineCount("C:\develop\data\text\UTF-8_Bom.txt", "UTF-8")
    vArr = adoSt.ReadTextFileToVArray("C:\develop\data\text\UTF-8_Bom.txt", "UTF-8")
    Call adoSt.WriteTextFileFromVArray("C:\develop\data\text\Write_SJIS.txt", vArr, "Shift_JIS", , True)
    
End Sub


'******************************************************************************
'* [概  要] ADODBExStreamWriteUTF8TextFileメソッドのTest｡
'* [詳  細] WriteUTF8TextFileメソッドのTest用。
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
'* [概  要] ADODBExStreamのIsBomIncludedメソッドのTest｡
'* [詳  細] IsBomIncludedメソッドのTest用。
'*
'******************************************************************************
Sub ADODBExStreamTest_IsBomIncluded()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream

    Debug.Print "BOM:" & adoSt.IsBomIncluded("C:\develop\data\text\UTF-8.txt")
    Debug.Print "BOM:" & adoSt.IsBomIncluded("C:\develop\data\text\UTF-8_Bom.txt")
    
End Sub

'******************************************************************************
'* [概  要] ADODBExStreamのReadFileToDumpメソッドのTest｡
'* [詳  細] ReadFileToDumpメソッドのTest用。
'*
'******************************************************************************
Sub ADODBExStreamTest_のReadFileToDump()
    Dim adoSt As ADODBExStream
    Set adoSt = New ADODBExStream

    Debug.Print adoSt.ReadFileToDump("C:\develop\data\text\UTF-8.txt")
    
End Sub


'******************************************************************************
'* [概  要] ADODBExのChangeFileEncodeメソッドのTest｡
'* [詳  細] ChangeFileEncodeメソッドのTest用。
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
    Debug.Print "◆Shift_JISで読込"
    vArr = adoSt.ReadTextFileToVArray(fPath, "shift_jis")
    For i = LBound(vArr) To UBound(vArr)
        Debug.Print i & ": " & vArr(i)
    Next i
    
    'when:
    Call ChangeFileEncode(fPath, "shift_jis", "UTF-8")
  
    'check:
    Debug.Print "◆UTF-8で読込"
    vArr = adoSt.ReadTextFileToVArray(fPath, "UTF-8")
    For i = LBound(vArr) To UBound(vArr)
        Debug.Print i & ": " & vArr(i)
    Next i
End Sub
'******************************************************************************
'* [概  要] ADODBExのChangeFilesEncodeメソッドのTest｡
'* [詳  細] ChangeFileEncodeメソッドのTest用。
'*
'******************************************************************************
Sub ADODBExTest_ChangeFilesEncode()
    Dim fPath As String: fPath = "C:\develop\data\text\work2"
    
    'when:
    Call ChangeFilesEncode(fPath, "UTF-8", "shift_jis")
End Sub

'******************************************************************************
'* [概  要] 大容量ファイル作成用メソッド｡
'* [詳  細] 大容量ファイルを作成する。
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
            .WriteText "これは大容量ファイルを作成する文章。,これだけ繰り返せば大容量といっても過言ではあるまい。ぶつぶつ。", adWriteLine
        Next
        .ExcludeBom
        .SaveToFile filePath, adSaveCreateOverWrite 'ファイル上書指定
        .CloseStream
    End With
End Sub

'******************************************************************************
'* [概  要] ADODBExのReadAndWriteメソッドのTest｡
'* [詳  細] ReadAndWriteメソッドのTest用。
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
'* [概  要] 行編集用メソッド。
'* [詳  細] ADODBExのReadAndWriteメソッドのTestで使用する行編集用メソッド｡
'*
'******************************************************************************
Public Function SampleFunc(rowData As String) As String
    Dim cols, colData, ret As String
    cols = Split(rowData, ",")
    For Each colData In cols
        If ret = "" Then
            ret = colData & "最初,"
        Else
            ret = ret & colData & "ジエンド"
        End If
    Next
    ret = Replace(ret, "ぶつぶつ。", "ぼやぼや。")
    SampleFunc = ret
End Function

