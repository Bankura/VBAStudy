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
'* [�@�\��] Encode/Decode���[�e�B���e�B�N���X
'* [��  ��] Encode/Decode�Ɋւ��郆�[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'* [�Q  �l] https://defun.exblog.jp/5309470/
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* WindowsAPI��`
'******************************************************************************

'******************************************************************************
'* Enum��`
'******************************************************************************

'******************************************************************************
'* �萔��`
'******************************************************************************

'******************************************************************************
'* �����ϐ���`
'******************************************************************************

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] URL�G���R�[�h�����B
'* [��  ��] URL�G���R�[�h���s���B
'*
'* @param target �ΏۂƂȂ镶����
'* @return �G���R�[�h�㕶����
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
'* [�T  �v] URL�f�R�[�h�����B
'* [��  ��] URL�f�R�[�h���s���B
'*
'* @param target �ΏۂƂȂ镶����
'* @return �G���R�[�h�㕶����
'******************************************************************************
Public Function DecodeUrl(Target As String, Optional cSet As String = "UTF-8") As String
    Dim bin() As Byte, var, i As Long, j As Long, cnt As Long
    Dim buf() As Byte
    Dim tmp As String
    
    ' %�ŋ�؂��ď���
    var = Split(Target, "%")
    
    ' Byte�z��̗v�f���v�Z
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
    
    ' Byte�f�[�^�ɕϊ�
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
    
    ' Byte���當����֕ϊ�
    DecodeUrl = BinaryUtils.ByteArray2TxtString(bin, cSet)
End Function

'******************************************************************************
'* [�T  �v] Base64�G���R�[�h�����B
'* [��  ��] Base64�G���R�[�h���s���B
'*
'* @param target �ΏۂƂȂ镶����
'* @return �G���R�[�h�㕶����
'******************************************************************************
Public Function EncodeBase64(Target As String, Optional cSet As String = "UTF-8") As String
    Dim byteData() As Byte
    byteData = BinaryUtils.TxtString2ByteArray(Target, cSet)
    
    EncodeBase64 = BytesToBase64(byteData)
End Function

'******************************************************************************
'* [�T  �v] BytesToBase64
'* [��  ��] Byte�z��base64������ւ̕ϊ��B
'*
'* @param varBytes Byte�z��
'* @return base64������
'******************************************************************************
Public Function BytesToBase64(varBytes() As Byte) As String
    With CreateObject("Msxml2.DOMDocument.6.0").createElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = Replace(.text, vbLf, "")
    End With
End Function

'******************************************************************************
'* [�T  �v] Base64�f�R�[�h�����B
'* [��  ��] Base64�f�R�[�h���s���B
'* [�Q�@�l] <https://defun.exblog.jp/5309470/>
'*
'* @param target �ΏۂƂȂ镶����
'* @return �f�R�[�h�㕶����
'******************************************************************************
Public Function DecodeBase64(Target As String, Optional cSet As String = "UTF-8") As String
    Dim byteData() As Byte
    byteData = Base64toBytes(Target)
    
    DecodeBase64 = BinaryUtils.ByteArray2TxtString(byteData, cSet)
End Function

'******************************************************************************
'* [�T  �v] Base64toBytes
'* [��  ��] base64������Byte�z��ւ̕ϊ��B
'*
'* @param varStr base64������
'* @return Byte�z��
'******************************************************************************
 Public Function Base64toBytes(varStr As String) As Byte()
    With CreateObject("Msxml2.DOMDocument.6.0").createElement("b64")
         .DataType = "bin.base64"
         .text = varStr
         Base64toBytes = .nodeTypedValue
    End With
 End Function


