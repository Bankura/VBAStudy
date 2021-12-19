VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EncodeDecodeUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Encode/Decodeユーティリティクラス
'* [詳  細] Encode/Decodeに関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'* [参  考] https://defun.exblog.jp/5309470/
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* WindowsAPI定義
'******************************************************************************

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************

'******************************************************************************
'* 内部変数定義
'******************************************************************************

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] URLエンコード処理。
'* [詳  細] URLエンコードを行う。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Public Function EncodeUrl(Target As String, Optional cSet As String = "UTF-8") As String
    Dim buf() As Byte, s As String, i As Long
    
    buf = BinaryUtils.TxtString2ByteArray(Target, cSet)
    
    For i = 0 To UBound(buf)
        Dim flg As Boolean: flg = False
        Select Case buf(i)
            Case 45, 46, 95, 126 '-._~
                flg = True
            Case 48 To 57 '0-9
                flg = True
            Case 65 To 90 'A-Z
                flg = True
            Case 97 To 122 'a-z
                flg = True
        End Select
        If flg Then
            s = s & Chr(buf(i))
        Else
            s = s & "%" & Hex(buf(i))
        End If
    Next
    EncodeUrl = s
End Function

'******************************************************************************
'* [概  要] URLデコード処理。
'* [詳  細] URLデコードを行う。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Public Function DecodeUrl(Target As String, Optional cSet As String = "UTF-8") As String
    Dim bin() As Byte, var, i As Long, j As Long, cnt As Long
    Dim buf() As Byte
    Dim tmp As String
    
    ' %で区切って処理
    var = Split(Target, "%")
    
    ' Byte配列の要素数計算
    For i = LBound(var) To UBound(var)
        tmp = CStr(var(i))
        If i = LBound(var) And Left(Target, 1) <> "%" Then
            For j = 1 To Len(tmp)
               cnt = cnt + 1
            Next j
        ElseIf Len(tmp) = 2 Then
            cnt = cnt + 1
        Else
            cnt = cnt + 1
            For j = 3 To Len(tmp)
               cnt = cnt + 1
            Next j
        End If
    Next i
    ReDim bin(0 To cnt - 1) As Byte
    cnt = 0
    
    ' Byteデータに変換
    For i = LBound(var) To UBound(var)
        tmp = CStr(var(i))
        
        If i = LBound(var) And Left(Target, 1) <> "%" Then
            For j = 1 To Len(tmp)
               bin(cnt) = Asc(Mid(tmp, j, 1))
               cnt = cnt + 1
            Next j
        ElseIf Len(tmp) = 2 Then
            bin(cnt) = CByte(val("&H" & tmp))
            cnt = cnt + 1
        Else
            bin(cnt) = CByte(val("&H" & Mid(tmp, 1, 2)))
            cnt = cnt + 1
            For j = 3 To Len(tmp)
               bin(cnt) = Asc(Mid(tmp, j, 1))
               cnt = cnt + 1
            Next j
        End If
    Next i
    
    ' Byteから文字列へ変換
    DecodeUrl = BinaryUtils.ByteArray2TxtString(bin, cSet)
End Function

'******************************************************************************
'* [概  要] Base64エンコード処理。
'* [詳  細] Base64エンコードを行う。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Public Function EncodeBase64(Target As String, Optional cSet As String = "UTF-8") As String
    Dim byteData() As Byte
    byteData = BinaryUtils.TxtString2ByteArray(Target, cSet)
    
    EncodeBase64 = BytesToBase64(byteData)
End Function

'******************************************************************************
'* [概  要] BytesToBase64
'* [詳  細] Byte配列→base64文字列への変換。
'*
'* @param varBytes Byte配列
'* @return base64文字列
'******************************************************************************
Public Function BytesToBase64(varBytes() As Byte) As String
    With CreateObject("Msxml2.DOMDocument.6.0").createElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = Replace(.text, vbLf, "")
    End With
End Function

'******************************************************************************
'* [概  要] Base64デコード処理。
'* [詳  細] Base64デコードを行う。
'* [参　考] <https://defun.exblog.jp/5309470/>
'*
'* @param target 対象となる文字列
'* @return デコード後文字列
'******************************************************************************
Public Function DecodeBase64(Target As String, Optional cSet As String = "UTF-8") As String
    Dim byteData() As Byte
    byteData = Base64toBytes(Target)
    
    DecodeBase64 = BinaryUtils.ByteArray2TxtString(byteData, cSet)
End Function

'******************************************************************************
'* [概  要] Base64toBytes
'* [詳  細] base64文字列→Byte配列への変換。
'*
'* @param varStr base64文字列
'* @return Byte配列
'******************************************************************************
 Public Function Base64toBytes(varStr As String) As Byte()
    With CreateObject("Msxml2.DOMDocument.6.0").createElement("b64")
         .DataType = "bin.base64"
         .text = varStr
         Base64toBytes = .nodeTypedValue
    End With
 End Function



