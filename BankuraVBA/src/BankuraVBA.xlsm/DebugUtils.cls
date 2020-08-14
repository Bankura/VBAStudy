VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] �f�o�b�O���[�e�B���e�B�N���X
'* [��  ��] �f�o�b�O�Ɋւ��郆�[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'*
'* [�Q  �l] https://excel-ubara.com/excelvba4/EXCEL_VBA_422.html
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
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
Private mWinApi As WinAPI

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************


'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mWinApi = Base.GetWinAPI()
End Sub
Private Sub Class_Terminate()
    Set mWinApi = Nothing
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] Dump
'* [��  ��] �w�肵���ϐ��̓��e���_���v������ɕϊ�����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param x �Ώۂ̕ϐ�
'* @return String �_���v������
'*
'******************************************************************************
Public Function Dump(ByVal x As Variant) As String
    Dump = Core.Dump(x)
End Function

'******************************************************************************
'* [�T  �v] PrintDump
'* [��  ��] �w�肵���ϐ��̓��e���_���v������ɕϊ����A�C�~�f�B�G�C�g�E�B���h�E
'*          �ɏo�͂���B
'*
'* @param x �Ώۂ̕ϐ�
'*
'******************************************************************************
Public Sub PrintDump(ByVal x As Variant)
    Debug.Print Dump(x)
End Sub

'******************************************************************************
'* [�T  �v] Variant�z��f�o�b�O�o�͏���
'* [��  ��] Variant�z��̓��e���C�~�f�B�G�C�g�E�B���h�E�ɏo�͂���B
'*          2�����z��܂őΉ��B
'*
'* @param vArr Variant�z��
'******************************************************************************
Public Sub PrintVariantArray(varr)
    Dim i As Long, j As Long, tmp As String
    For i = LBound(varr, 1) To UBound(varr, 1)
        If ArrayUtils.Is2DArray(varr) Then
            For j = LBound(varr, 2) To UBound(varr, 2)
                tmp = tmp & varr(i, j) & " | "
            Next
        Else
            tmp = varr(i)
        End If
        Debug.Print tmp
        tmp = ""
    Next
End Sub

'******************************************************************************
'* [�T  �v] Print2DArrayRow
'* [��  ��] Variant2�����z��̎w�肵���s�i1�����C���f�b�N�X�j�̓��e��
'*          �C�~�f�B�G�C�g�E�B���h�E�ɏo�͂���B
'*
'* @param vArr Variant�z��
'* @param row �s�i1�����C���f�b�N�X�j
'******************************************************************************
Public Sub Print2DArrayRow(varr, ByVal row As Long)
    Dim j As Long, tmp As String
    For j = LBound(varr, 2) To UBound(varr, 2)
        tmp = tmp & varr(row, j) & " | "
    Next
    Debug.Print tmp
    tmp = ""

End Sub

'******************************************************************************
'* [�T  �v] RecordFormat�f�o�b�O�o�͏���
'* [��  ��] RecordFormat�̓��e���C�~�f�B�G�C�g�E�B���h�E�ɏo�͂���B
'*
'* @param rf RecordFormat
'******************************************************************************
Public Sub PrintRecordSet(rf As RecordFormat)
    Dim record As Collection, itm As Item, tmp As String
    For Each record In rf.RecordSet
        For Each itm In record
            tmp = tmp & itm.value & " | "
        Next
        Debug.Print tmp
        tmp = ""
    Next
End Sub

'******************************************************************************
'* [�T  �v] DumpMemoryFromVariant
'* [��  ��] �ϐ��̃A�h���X(VarPtr)�ƃ������̃_���v���C�~�f�B�G�C�g�E�B���h�E
'*          �ɏo�͂���B
'*
'* @param var �������̃_���v���擾����ϐ�
'* @return LongPtr ����������擾�����o�C�g�����l
'******************************************************************************
Public Function DumpMemoryFromVariant(ByRef var) As LongPtr
    Dim addr As LongPtr, size As Long, opt As String
    Dim vType As Long
    vType = CLng(BinaryUtils.ByteArray2LongPtr(BinaryUtils.CopyMemory2ByteArray(VarPtr(var), 2)))
    Select Case vType
        Case VT_I2, VT_BYREF + VT_I2     ' Integer
            size = 2
            opt = "dec"
        Case VT_I4, VT_DATE, VT_BYREF + VT_I4, VT_BYREF + VT_DATE   ' Long, Date
            size = 4
            opt = "dec"
        Case VT_R4, VT_BYREF + VT_R4     ' Single
            size = 4
            opt = "bin"
        Case VT_R8, VT_BYREF + VT_R8     ' Double
            size = 8
            opt = "bin"
        Case VT_BOOL, VT_UI1, VT_BYREF + VT_BOOL, VT_BYREF + VT_UI1  ' Boolean, Byte
            size = 1
            opt = "bin"
        Case VT_BSTR, VT_DISPATCH, VT_BYREF + VT_BSTR, VT_BYREF + VT_DISPATCH ' String, Object
            size = 8
            opt = "dec"
        Case Else ' LongLong, Currency
            size = 8
            opt = "bin"
    End Select
    
    If (vType And VT_BYREF) = VT_BYREF Then
        addr = BinaryUtils.ByteArray2LongPtr(BinaryUtils.CopyMemory2ByteArray(VarPtr(var) + LONGPTR_SIZE, LONGPTR_SIZE))
    Else
        addr = VarPtr(var) + LONGPTR_SIZE
    End If
    
    DumpMemoryFromVariant = DumpMemory(addr, size, opt)
End Function

'******************************************************************************
'* [�T  �v] DumpMemoryFromString
'* [��  ��] �ϐ��̃A�h���X(StrPtr)�ƃ������̃_���v���C�~�f�B�G�C�g�E�B���h�E
'*          �ɏo�͂���B
'*
'* @param var �������̃_���v���擾����ϐ�
'* @return LongPtr ����������擾�����o�C�g�����l
'******************************************************************************
Public Function DumpMemoryFromString(ByRef var) As LongPtr
    DumpMemoryFromString = DumpMemory(StrPtr(var), LenB(var), "str")
End Function

'******************************************************************************
'* [�T  �v] DumpMemoryFromObject
'* [��  ��] �ϐ��̃A�h���X(ObjPtr)�ƃ������̃_���v���C�~�f�B�G�C�g�E�B���h�E
'*          �ɏo�͂���B
'*
'* @param var �������̃_���v���擾����Object�ϐ�
'* @return LongPtr ����������擾�����o�C�g�����l
'******************************************************************************
Public Function DumpMemoryFromObject(ByRef obj As Object) As LongPtr
    DumpMemoryFromObject = DumpMemory(ObjPtr(obj), LONGPTR_SIZE, "dec")
End Function

'******************************************************************************
'* [�T  �v] DumpMemory
'* [��  ��] �ϐ��̃A�h���X(StrPtr)�ƃ������̃_���v���C�~�f�B�G�C�g�E�B���h�E
'*          �ɏo�͂���B
'*
'* @param addr �������_���v���擾����A�h���X
'*             �iVarPtr, StrPtr, ObjPtr�̒l���x�[�X�Ɏw�肷��z��j
'* @param size �������̃_���v���擾����o�C�g�T�C�Y
'* @param opt �������_���v�̓��e�ɉ����ĕ\������l�̕ϊ��`��
'*              dec     �F �o�C�g�����l
'*              str     �F ������
'*              bin     �F �o�C�i���i2�i�j
'*              �w��Ȃ��F �\�����Ȃ�
'* @return LongPtr ����������擾�����o�C�g�����l
'******************************************************************************
Public Function DumpMemory(ByVal addr As LongPtr, _
                    ByVal size As Long, _
                    Optional ByVal opt As String) As LongPtr

    If addr <= 0 Or size <= 0 Then
        Debug.Print addr & " :"
        Exit Function
    End If

    ReDim bData(1 To size) As Byte
    bData = BinaryUtils.CopyMemory2ByteArray(addr, size)
  
    Select Case LCase(opt)
        Case "str"
            Call DumpByteArrayAsString(bData, addr)
        Case "dec"
            Call DumpByteArrayAsDecimal(bData, addr)
        Case "bin"
            Call DumpByteArrayAsBinary(bData, addr)
        Case Else
            Call DumpByteArray(bData, addr)
    End Select

    DumpMemory = BinaryUtils.ByteArray2LongPtr(bData)
End Function

'******************************************************************************
'* [�T  �v] DumpByteArray
'* [��  ��] �o�C�g����16�i�_���v������ɕϊ����ăC�~�f�B�G�C�g�E�B���h�E��
'*          �ɏo�͂���B
'*
'* @param bData() �o�C�g���i�z��j
'* @param addr �A�h���X�i�ǉ��\���p�j�B�o�C�g���̎擾�����Ƃ��Ďw��B
'******************************************************************************
Public Sub DumpByteArray(bData() As Byte, Optional ByVal addr)
    Dim pStr As String
    If ValidateUtils.IsNotMissing(addr) Then
        pStr = addr & " : "
    End If
    pStr = pStr & BinaryUtils.ByteArray2HexStrings(bData)
    Debug.Print pStr
End Sub

'******************************************************************************
'* [�T  �v] DumpByteArrayAsString
'* [��  ��] �o�C�g����16�i�_���v������ɕϊ����ăC�~�f�B�G�C�g�E�B���h�E��
'*          �ɏo�͂���B
'*          �������𕶎���ϊ��������e��ǉ����ďo�͂���B
'*
'* @param bData() �o�C�g���i�z��j
'* @param addr �A�h���X�i�ǉ��\���p�j�B�o�C�g���̎擾�����Ƃ��Ďw��B
'******************************************************************************
Public Sub DumpByteArrayAsString(bData() As Byte, Optional ByVal addr)
    Dim pStr As String
    If ValidateUtils.IsNotMissing(addr) Then
        pStr = addr & " : "
    End If
    pStr = pStr & BinaryUtils.ByteArray2HexStrings(bData) & " �� " & CStr(bData)
    Debug.Print pStr
End Sub

'******************************************************************************
'* [�T  �v] DumpByteArrayAsDecimal
'* [��  ��] �o�C�g����16�i�_���v������ɕϊ����ăC�~�f�B�G�C�g�E�B���h�E��
'*          �ɏo�͂���B
'*          �������𐮐��l�ɕϊ��������e��ǉ����ďo�͂���B
'*
'* @param bData() �o�C�g���i�z��j
'* @param addr �A�h���X�i�ǉ��\���p�j�B�o�C�g���̎擾�����Ƃ��Ďw��B
'******************************************************************************
Public Sub DumpByteArrayAsDecimal(bData() As Byte, Optional ByVal addr)
    Dim pStr As String
    If ValidateUtils.IsNotMissing(addr) Then
        pStr = addr & " : "
    End If
    pStr = pStr & BinaryUtils.ByteArray2HexStrings(bData) & " �� " & BinaryUtils.ByteArray2LongPtr(bData)
    Debug.Print pStr
End Sub

'******************************************************************************
'* [�T  �v] DumpByteArrayAsBinary
'* [��  ��] �o�C�g����16�i�_���v������ɕϊ����ăC�~�f�B�G�C�g�E�B���h�E��
'*          �ɏo�͂���B
'*          ���������o�C�i���ϊ��������e��ǉ����ďo�͂���B
'*
'* @param bData() �o�C�g���i�z��j
'* @param addr �A�h���X�i�ǉ��\���p�j�B�o�C�g���̎擾�����Ƃ��Ďw��B
'******************************************************************************
Public Sub DumpByteArrayAsBinary(bData() As Byte, Optional ByVal addr)
    Dim pStr As String
    If ValidateUtils.IsNotMissing(addr) Then
        pStr = addr & " : "
    End If
    pStr = pStr & BinaryUtils.ByteArray2HexStrings(bData) & " �� " & BinaryUtils.ByteArray2BinaryStrings(bData)
    Debug.Print pStr
End Sub


