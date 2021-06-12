VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinaryUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] �o�C�i�����샆�[�e�B���e�B�N���X
'* [��  ��] �o�C�i������Ɋւ��郆�[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'* [�Q  �l]
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* WindowsAPI��`
'******************************************************************************
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, ByVal Source As Any, ByVal Length As LongPtr)

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
'* [�T  �v] CopyMemory2ByteArray
'* [��  ��] �w�肵���������A�h���X����w��o�C�g�T�C�Y���̃o�C�g�����擾
'*          ���A�o�C�g�z��Ƃ��ĕԋp����B
'*
'* @param addr �A�h���X
'*             �iVarPtr, StrPtr, ObjPtr�̒l���x�[�X�Ɏw�肷��z��j
'* @param size �擾����o�C�g�T�C�Y
'* @return Byte() ����������擾�����o�C�g���
'******************************************************************************
Public Function CopyMemory2ByteArray(ByVal addr As LongPtr, ByVal size As Long) As Byte()
    If addr <= 0 Or size <= 0 Then
        CopyMemory2ByteArray = ArrayUtils.EmptyByteArray
        Exit Function
    End If
  
    ReDim bData(1 To size) As Byte
    Call CopyMemory(VarPtr(bData(1)), addr, size)
    CopyMemory2ByteArray = bData
End Function

'******************************************************************************
'* [�T  �v] ByteArray2HexStrings
'* [��  ��] �o�C�g�z���16�i������ɕϊ�����B
'*
'* @param bData() �o�C�g�z��
'* @param delimiter ��؂蕶���i�f�t�H���g�F���p�X�y�[�X�j
'* @return String 16�i������
'******************************************************************************
Public Function ByteArray2HexStrings(bData() As Byte, Optional ByVal Delimiter As String = " ") As String
    Dim size As Long
    size = ArrayUtils.GetLength(bData)
    If size <= 0 Then
        ByteArray2HexStrings = ""
        Exit Function
    End If
    
    Dim i As Long, hVal As String
    Dim pStr As StringEx: Set pStr = New StringEx
    
    For i = LBound(bData) To UBound(bData)
        hVal = Byte2HexString(bData(i))
        If i > LBound(bData) Then
            Call pStr.Append(Delimiter)
        End If
        Call pStr.Append(hVal)
    Next i
    ByteArray2HexStrings = pStr.ToString
End Function

'******************************************************************************
'* [�T  �v] HexStrings2ByteArray
'* [��  ��] 16�i��������o�C�g�z���ϊ�����B
'*
'* @param hstr 16�i������
'* @param delimiter ��؂蕶���i�f�t�H���g�F���p�X�y�[�X�j
'* @return String 16�i������
'******************************************************************************
Public Function HexStrings2ByteArray(ByVal hstr As String, Optional ByVal Delimiter As String = " ") As Byte()
    If hstr = "" Then
        HexStrings2ByteArray = ArrayUtils.EmptyByteArray
        Exit Function
    End If
    
    Dim vArr
    If Delimiter = "" Then
        If Len(hstr) Mod 2 <> 0 Then
            HexStrings2ByteArray = ArrayUtils.EmptyByteArray
            Exit Function
        End If
        vArr = StringUtils.SplitByLength(hstr, 2)
    Else
        vArr = Split(hstr, Delimiter)
    End If
    
    Dim bData() As Byte, i As Long
    ReDim bData(LBound(vArr) To UBound(vArr)) As Byte
    
    For i = LBound(vArr) To UBound(vArr)
        bData(i) = HexString2Byte(vArr(i))
    Next i
    HexStrings2ByteArray = bData
End Function

'******************************************************************************
'* [�T  �v] ByteArray2BinaryStrings
'* [��  ��] �o�C�g�z����o�C�i��������ɕϊ�����B
'*
'* @param bData() �o�C�g�z��
'* @return String �o�C�i��������
'******************************************************************************
Public Function ByteArray2BinaryStrings(bData() As Byte) As String
    Dim size As Long
    size = ArrayUtils.GetLength(bData)
    If size <= 0 Then
        ByteArray2BinaryStrings = ""
        Exit Function
    End If
    
    Dim i As Long, bVal As String, hVal As String
    For i = LBound(bData) To UBound(bData)
        hVal = Byte2HexString(bData(i))
        bVal = Hex2Bin(hVal) & bVal
    Next i
    ByteArray2BinaryStrings = bVal
End Function

'******************************************************************************
'* [�T  �v] ByteArray2LongPtr
'* [��  ��] �o�C�g�z��𐮐��l�ɕϊ�����B
'*
'* @param bData() �o�C�g�z��
'* @return LongPtr �����l
'******************************************************************************
Public Function ByteArray2LongPtr(bData() As Byte) As LongPtr
    Dim size As Long
    size = ArrayUtils.GetLength(bData)
    If size <= 0 Or 8 < size Then
        ByteArray2LongPtr = 0
        Exit Function
    End If
    
    Dim i As Long, dVal As LongPtr, hVal As String
    For i = LBound(bData) To UBound(bData)
        hVal = Byte2HexString(bData(i))
        dVal = dVal + Hex2Dec(hVal, i, LBound(bData))
    Next i
    
    ByteArray2LongPtr = dVal
End Function

'******************************************************************************
'* [�T  �v] Hex2dec
'* [��  ��] 1�o�C�g����16�i��������o�C�g�����l�ɕϊ�����B
'*          �����o�C�g����16�i������̈ʒu���w�肷�邱�Ƃŏd�݂Â����\�B
'*
'* @param hVal 1�o�C�g����16�i������
'* @param pos �ʒu
'* @param startPos �J�n�ʒu
'* @return LongPtr �o�C�g�����l
'******************************************************************************
Public Function Hex2Dec(ByVal hVal As String, Optional ByVal pos As LongPtr = 0, Optional ByVal startPos As LongPtr = 0) As LongPtr
    Hex2Dec = HexString2LongPtr(hVal) * (256 ^ (pos - startPos))
End Function

'******************************************************************************
'* [�T  �v] Hex2bin
'* [��  ��] 1�o�C�g����16�i��������o�C�i��������ɕϊ�����B
'*
'* @param hVal 1�o�C�g����16�i������
'* @return String �o�C�i��������
'******************************************************************************
Public Function Hex2Bin(ByVal hVal As String) As String
    Dim dVal As LongPtr
    dVal = HexString2LongPtr(hVal)
    Dim power2 As LongPtr: power2 = 1
    Do Until dVal < power2
        If (dVal And power2) <> 0 Then
            Hex2Bin = "1" & Hex2Bin
        Else
            Hex2Bin = "0" & Hex2Bin
        End If
        power2 = power2 * 2
    Loop
    Hex2Bin = StringUtils.LeftPad0(Hex2Bin, 8)
End Function

'******************************************************************************
'* [�T  �v] Bin2Hex
'* [��  ��] �o�C�i���������16�i������ɕϊ�����B
'*
'* @param bVal �o�C�i��������
'* @return LongPtr 16�i������
'******************************************************************************
Public Function Bin2Hex(ByVal bVal As String) As String
    If Not IsBinaryString(bVal) Then
        Bin2Hex = ""
        Exit Function
    End If
    
    Dim str As StringEx: Set str = New StringEx
    Dim v, tmp As Long: tmp = 0
    For Each v In StringUtils.SplitByLength(bVal, 4)
        Dim i As Long
        tmp = 0
        For i = Len(v) To 1 Step -1
            If Mid(v, Len(v) - i + 1, 1) = "1" Then
               tmp = tmp + (2 ^ (i - 1))
            End If
        Next
        str.Append Hex$(tmp)
    Next

    Bin2Hex = str.ToString
End Function

'******************************************************************************
'* [�T  �v] Byte2HexString
'* [��  ��] 1�o�C�g��16�i������ɕϊ�����B
'*
'* @param b 1�o�C�g
'* @param prependSuffix �擪��"&H"��t�^���邩�i�f�t�H���g�FFalse�j
'* @return String 16�i������
'******************************************************************************
Public Function Byte2HexString(ByVal b As Byte, Optional ByVal prependSuffix As Boolean = False) As String
    Byte2HexString = StringUtils.LeftPad0(Hex$(b), 2)
    If prependSuffix Then Byte2HexString = "&H" & Byte2HexString
End Function

'******************************************************************************
'* [�T  �v] HexString2LongPtr
'* [��  ��] 1�o�C�g����16�i��������o�C�g�����l�ɕϊ�����B
'*
'* @param hVal 1�o�C�g����16�i������
'* @return LongPtr �o�C�g�����l
'******************************************************************************
Public Function HexString2LongPtr(ByVal hstr As String) As LongPtr
    HexString2LongPtr = CLngPtr(StringUtils.PrependIfMissingIgnoreCase(hstr, "&H"))
End Function

'******************************************************************************
'* [�T  �v] HexString2Byte
'* [��  ��] 1�o�C�g����16�i��������o�C�g�ɕϊ�����B
'*
'* @param hVal 1�o�C�g����16�i������
'* @return Byte �o�C�g
'******************************************************************************
Public Function HexString2Byte(ByVal hstr As String) As Byte
    HexString2Byte = CByte(StringUtils.PrependIfMissingIgnoreCase(hstr, "&H"))
End Function

'******************************************************************************
'* [�T  �v] IsHexString
'* [��  ��] 16�i�����񂩂ǂ����`�F�b�N�B
'*
'* @param val �`�F�b�N�Ώە�����
'* @return Boolean �`�F�b�N���ʁiTrue�FOK�AFalse�FNG�j
'******************************************************************************
Public Function IsHexString(ByVal val As String) As Boolean
    Dim i As Long

    For i = 1 To Len(val)
        If Not Mid(val, i, 1) Like "[0-9a-fA-F]" Then
            IsHexString = False
            Exit Function
        End If
    Next
    IsHexString = True
End Function

'******************************************************************************
'* [�T  �v] IsBinaryString
'* [��  ��] 2�i�����񂩂ǂ����`�F�b�N�B
'*
'* @param val �`�F�b�N�Ώە�����
'* @return Boolean �`�F�b�N���ʁiTrue�FOK�AFalse�FNG�j
'******************************************************************************
Public Function IsBinaryString(ByVal val As String) As Boolean
    If Len(val) Mod 4 <> 0 Then
        IsBinaryString = False
        Exit Function
    End If
    Dim i As Long

    For i = 1 To Len(val)
        If Not Mid(val, i, 1) Like "[0-1]" Then
            IsBinaryString = False
            Exit Function
        End If
    Next
    IsBinaryString = True
End Function

'******************************************************************************
'* [�T  �v] �����񁨃o�C�g�f�[�^�ϊ������B
'* [��  ��] �����񂩂�o�C�g�f�[�^�֕ϊ����s���B
'*
'* @param text �ΏۂƂȂ镶����
'* @param cSet �L�����Z�b�g�i�����R�[�h�w��j�B�ȗ����́uUTF-8�v�B
'*               ���g�p�\��Charset��F
'*                 utf-8�Ashift_jis�Aeuc-jp�AISO-2022-JP�Aunicode�i����j
'*                 �uHKEY_CLASSES_ROOT\MIME\Database\Charset�v�Q�ƁB
'* @return �ϊ���o�C�g�f�[�^
'******************************************************************************
Public Function TxtString2ByteArray(ByVal text As String, Optional ByVal cSet As String = "UTF-8") As Byte()
    TxtString2ByteArray = StringUtils.GetBytes(text, cSet)
End Function

'******************************************************************************
'* [�T  �v] �o�C�g�f�[�^��������ϊ������B
'* [��  ��] �o�C�g�f�[�^���當����֕ϊ����s���B
'*
'* @param text �ΏۂƂȂ�o�C�g�f�[�^
'* @param cSet �L�����Z�b�g�i�����R�[�h�w��j�B�ȗ����́uUTF-8�v�B
'*               ���g�p�\��Charset��F
'*                 utf-8�Ashift_jis�Aeuc-jp�AISO-2022-JP�Aunicode�i����j
'*                 �uHKEY_CLASSES_ROOT\MIME\Database\Charset�v�Q�ƁB
'* @return �ϊ��㕶����
'******************************************************************************
Public Function ByteArray2TxtString(bin() As Byte, Optional ByVal cSet As String = "UTF-8") As String
    ByteArray2TxtString = StringUtils.ToEncodedString(bin, cSet)
End Function