VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''+----                                                                   --+
'''|                             Ariawase 0.9.0                              |
'''|                Ariawase is free library for VBA cowboys.                |
'''|          The Project Page: https://github.com/vbaidiot/Ariawase         |
'''+--                                                                   ----+
Option Explicit

Private xItems As Variant
Private xLength As Long
Private xIndex As Long

Private Sub Class_Initialize()
    xIndex = -1
    xLength = -1 + 32
    ReDim xItems(xLength - 1)
End Sub

Private Sub Extend()
    If xIndex < xLength Then GoTo Escape
    xLength = xLength + 1 'possible overflow (Err.Raise 6)
    xLength = -1 + xLength + xLength
    ReDim Preserve xItems(xLength - 1)
    
Escape:
End Sub

Public Function AddVal(ByVal val As Variant)
    xIndex = xIndex + 1
    Extend
    Let xItems(xIndex) = val
End Function

Public Function AddObj(ByVal obj As Variant)
    xIndex = xIndex + 1
    Extend
    Set xItems(xIndex) = obj
End Function

Public Function ToArray() As Variant
    Dim arr As Variant: arr = xItems
    If xIndex > -1 Then
        ReDim Preserve arr(xIndex)
    Else
        arr = Array()
    End If
    ToArray = arr
End Function


'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* �g�����\�b�h
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'*-----------------------------------------------------------------------------
'* Item �v���p�e�B
'*-----------------------------------------------------------------------------
Public Property Get Item(ByVal idx As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If IsInvalidIndex(idx) Then
        Item = Empty
        Exit Property
    End If

    If IsObject(xItems(idx)) Then
        Set Item = xItems(idx)
    Else
        Item = xItems(idx)
    End If
End Property
Public Property Let Item(ByVal idx As Long, ByVal arg As Variant)
    If IsInvalidIndex(idx) Then
        Exit Property
    End If
    
    If IsObject(arg) Then
        Set xItems(idx) = arg
    Else
        xItems(idx) = arg
    End If
End Property

Private Function IsInvalidIndex(ByVal idx As Long) As Boolean
    IsInvalidIndex = (idx < 0 Or xIndex < idx)
End Function


'******************************************************************************
'* [�T  �v] Init ���\�b�h
'* [��  ��] �������p���\�b�h�
'*
'* @param val �����z��
'* @return ArrayEx�I�u�W�F�N�g
'*
'******************************************************************************
Public Function Init(Optional ByVal val) As ArrayEx
    If VBA.IsMissing(val) Then
        Call Clear
        Set Init = Me
        Exit Function
    End If
    If IsArray(val) Then
        xIndex = UBound(val)
        xLength = xIndex + 32
        xItems = val
        ReDim Preserve xItems(xLength - 1)
    Else
        Call Clear
        Call Add(val)
    End If
    Set Init = Me
End Function

'******************************************************************************
'* [�T  �v] Clear
'* [��  ��] ����������
'*
'* @return ArrayEx�I�u�W�F�N�g
'*
'******************************************************************************
Public Function Clear() As ArrayEx
    xIndex = -1
    xLength = -1 + 32
    ReDim xItems(xLength - 1)
    Set Clear = Me
End Function

'******************************************************************************
'* [�T  �v] Length
'* [��  ��] �v�f����ԋp����
'*
'* @return Long �v�f��
'*
'******************************************************************************
Property Get Length() As Long
    Length = xIndex + 1
End Property

'******************************************************************************
'* [�T  �v] LastIndex
'* [��  ��] �ŏI�C���f�b�N�X��ԋp����
'*
'* @return Long �ŏI�C���f�b�N�X
'*
'******************************************************************************
Property Get LastIndex() As Long
    LastIndex = xIndex
End Property

'******************************************************************************
'* [�T  �v] Add
'* [��  ��] ���ڂ�ǉ�����
'*
'* @param var �ǉ�����
'* @return ArrayEx
'*
'******************************************************************************
Public Function Add(ByVal var As Variant) As ArrayEx
    If IsObject(var) Then
        Call AddObj(var)
    Else
        Call AddVal(var)
    End If
    Set Add = Me
End Function

'******************************************************************************
'* [�T  �v] AddAll
'* [��  ��] ���ڂ𕡐��ǉ�����
'*
'* @param vars �ǉ����ځi�����j
'* @return ArrayEx
'*
'******************************************************************************
Public Function AddAll(ParamArray vars() As Variant) As ArrayEx
    Dim tmp: tmp = vars
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        Add v
    Next
    Set AddAll = Me
End Function

'******************************************************************************
'* [�T  �v] Expand
'* [��  ��] �z��̈���g������
'*
'* @param xAddCount �ǉ����ڐ�
'* @return ArrayEx
'*
'******************************************************************************
Public Function Expand(ByVal xAddCount As Long) As ArrayEx
    xIndex = xIndex + xAddCount
    Call Extend
    Set Expand = Me
End Function

'******************************************************************************
'* [�T  �v] IsEmptyArray
'* [��  ��] �z�񂪖��������E��z�񂩔��肷��
'*
'* @return Boolean ���茋��(True:���������E��)
'*
'******************************************************************************
Public Function IsEmptyArray() As Boolean
    IsEmptyArray = (xIndex = -1)
End Function

'******************************************************************************
'* [�T  �v] Equals
'* [��  ��] �ꎟ���z����r���A�C�R�[���ł����True��ԋp����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param arr2 As Variant(Of Array(Of T))
'* @param swAllowNull As Boolean
'* @return As Variant(Of Boolean Or Null Or Empty)
'******************************************************************************
Public Function Equals(ByVal arr2 As Variant, Optional ByVal swAllowNull As Boolean = False) As Variant
    Equals = Core.ArrEquals(Me.ToArray, arr2, swAllowNull)
End Function

'******************************************************************************
'* [�T  �v] Compare
'* [��  ��] �ꎟ���z��ix,y�j���r���Ax=y�ł����0�Ax<y�ł����-1�Ax>y
'*          �ł����1��ԋp����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param arr2 As Variant(Of Array(Of T))
'* @return As Variant(Of Integer Or Null)
'******************************************************************************
Public Function Compare(ByVal arr2 As Variant, Optional ByVal swAllowNull As Boolean = False) As Variant
    Compare = Core.ArrCompare(Me.ToArray, arr2, swAllowNull)
End Function

'******************************************************************************
'* [�T  �v] IndexOf
'* [��  ��] �ꎟ���z�񂩂�w�肵���l�ƈ�v����v�f�̃C���f�b�N�X��ԋp����B
'*          �l�̑{���J�n�ʒu�Ɨv�f����C�ӎw��\�B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param val As Variant(Of T)
'* @param ixStart As Variant(Of Long)
'* @param cnt As Variant(Of Long)
'* @return As Long
'******************************************************************************
Public Function IndexOf(ByVal val As Variant, Optional ByVal ixStart As Variant, Optional ByVal cnt As Variant) As Long
    IndexOf = Core.ArrIndexOf(Me.ToArray, val, ixStart, cnt)
End Function

'******************************************************************************
'* [�T  �v] Reverse
'* [��  ��] �ꎟ���z��̗v�f�̕��я��𔽓]������B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param arr As Variant(Of Array(Of T))
'******************************************************************************
Public Sub Reverse()
    Dim arr: arr = Me.ToArray
    Call Core.ArrRev(arr)
    Call Init(arr)
End Sub

'******************************************************************************
'* [�T  �v] Sort
'* [��  ��] �ꎟ���z��̃\�[�g���s���B�f�t�H���g�͏����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param orderAsc As Boolean
'******************************************************************************
Public Sub Sort(Optional ByVal orderAsc As Boolean = True)
    Dim arr: arr = Me.ToArray
    Call Core.ArrSort(arr, orderAsc)
    Call Init(arr)
End Sub

'******************************************************************************
'* [�T  �v] Uniq
'* [��  ��] �ꎟ���z��̏d���v�f��r������B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function Uniq() As Variant
    Uniq = Core.ArrUniq(Me.ToArray)
End Function

'******************************************************************************
'* [�T  �v] Concat
'* [��  ��] �ꎟ���z��ƍ��́i�}�[�W�j���ĕԋp����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param arr1 As Variant(Of Array(Of T))
'* @param arr2 As Variant(Of Array(Of T))
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function Concat(ByVal arr2 As Variant) As Variant
    Concat = Core.ArrConcat(Me.ToArray, arr2)
End Function

'******************************************************************************
'* [�T  �v] Slice
'* [��  ��] �ꎟ���z����X���C�X�i�ꕔ��z��Ƃ��Đ؂�o�j���ĕԋp����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param ixStart As Variant(Of Long)
'* @param ixEnd As Variant(Of Long)
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function Slice(Optional ByVal ixStart As Variant, Optional ByVal ixEnd As Variant) As Variant
    Slice = Core.ArrSlice(Me.ToArray, ixStart, ixEnd)
End Function

'******************************************************************************
'* [�T  �v] FromJagArray
'* [��  ��] �W���O�z��̃t���b�g���B�W���O�z����ꎟ���z��(ArrayEx)�ɕϊ�����
'*          �ԋp����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param jagArray As Variant(Of Array(Of Array(Of T)))
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function FromJagArray(ByVal jagArr As Variant) As ArrayEx
    Dim arr: arr = Core.ArrFlatten(jagArr)
    Set FromJagArray = Init(arr)
End Function

'******************************************************************************
'* [�T  �v] ToCollection
'* [��  ��] �ꎟ���z���Collection�I�u�W�F�N�g�ɕϊ�����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @return As Collection(Of T)
'******************************************************************************
Public Function ToCollection() As Collection
    Set ToCollection = Core.ArrToClct(Me.ToArray)
End Function

'******************************************************************************
'* [�T  �v] FromCollection
'* [��  ��] Collection�I�u�W�F�N�g���ꎟ���z��(ArrayEx)�ɕϊ�����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param clct As Collection(Of T)
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function FromCollection(ByVal clct As Collection) As ArrayEx
    Dim arr: arr = Core.ClctToArr(clct)
    Set FromCollection = Init(arr)
End Function

'******************************************************************************
'* [�T  �v] FromEnumerator
'* [��  ��] �J��Ԃ������iFor Each�j�\�ȃI�u�W�F�N�g���ꎟ���z��(ArrayEx)��
'*          �ϊ�����B
'*          Ariawase��Ext���W���[�����b�p�[���\�b�h�B
'*
'* @param enumr As Enumerator(Of T)
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function FromEnumerator(ByVal enumr As Object) As ArrayEx
    Dim arr: arr = Ext.EnumeratorToArr(enumr)
    Set FromEnumerator = Init(arr)
End Function

'******************************************************************************
'* [�T  �v] Range
'* [��  ��] �w�肵���͈͂̐��l�̔z����쐬����B
'*          Ariawase��Ext���W���[�����b�p�[���\�b�h�B
'*
'* @param fromVal As Variant(Of T)
'* @param toVal As Variant(Of T)
'* @param stepVal As Variant(Of T)
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function Range(ByVal fromVal As Variant, ByVal toVal As Variant, Optional ByVal stepVal As Variant = 1) As ArrayEx
    Dim arr: arr = Ext.ArrRange(fromVal, toVal, stepVal)
    Set Range = Init(arr)
End Function

'******************************************************************************
'* [�T  �v] Map
'* [��  ��] �z��̑S�v�f�ɑ΂��Ďw�肵���t�@���N�V�����̏�����K�p����B
'*          Ariawase��Ext���W���[�����b�p�[���\�b�h�B
'*
'* @param fun As Func(Of T, U)
'* @return ArrayEx(Of Array(Of U))
'******************************************************************************
Public Function Map(ByVal fun As Func) As ArrayEx
    Dim arr: arr = Ext.ArrMap(fun, Me.ToArray)
    Set Map = Init(arr)
End Function

'******************************************************************************
'* [�T  �v] Zip
'* [��  ��] �Q�̔z��ɑ΂��Ďw�肵���t�@���N�V�����̏�����K�p��
'*          �P�̔z���Ԃ��B
'*          Ariawase��Ext���W���[�����b�p�[���\�b�h�B
'*
'* @param fun As Func(Of T, U, R)
'* @param arr2 As Variant(Of Array(Of U))
'* @return ArrayEx(Of Array(Of R))
'******************************************************************************
Public Function Zip(ByVal fun As Func, ByVal arr2 As Variant) As ArrayEx
    Dim arr: arr = Ext.ArrZip(fun, Me.ToArray, arr2)
    Set Zip = Init(arr)
End Function

'******************************************************************************
'* [�T  �v] Filter
'* [��  ��] �z��ɑ΂��Ďw�肵���t�B���^�����p�̃t�@���N�V������K�p���A
'*          �����Ɉ�v����v�f�݂̂̔z���Ԃ��B
'*          Ariawase��Ext���W���[�����b�p�[���\�b�h�B
'*
'* @param fun As Func(Of T, Boolean)
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function Filter(ByVal fun As Func) As ArrayEx
    Dim arr: arr = Ext.ArrFilter(fun, Me.ToArray)
    Set Filter = Init(arr)
End Function

'******************************************************************************
'* [�T  �v] GroupBy
'* [��  ��] �z��ɑ΂��Ďw�肵���O���[�s���O�����p�̃t�@���N�V������K�p���A
'*          ���ʂ�Ԃ��B
'*          Ariawase��Ext���W���[�����b�p�[���\�b�h�B
'*
'* @param fun As Func(Of T, K)
'* @return ArrayEx(Of Array(Of Tuple`2(Of K, T)))
'******************************************************************************
Public Function GroupBy(ByVal fun As Func) As ArrayEx
    Dim arr: arr = Ext.ArrGroupBy(fun, Me.ToArray)
    Set GroupBy = Init(arr)
End Function

'******************************************************************************
'* [�T  �v] Fold
'* [��  ��] ��ݍ��݊֐��B�z����̗v�f�ɑ΂��Ďw�肵�����Z�p�t�@���N�V����
'*          ��K�p���A���̌��ʂƎ��̗v�f�ɂ܂��t�@���N�V������K�p����
'*          �Ƃ������������J��Ԃ��ŏI�I�Ȍ��ʂ�Ԃ��B
'*          Ariawase��Ext���W���[�����b�p�[���\�b�h�B
'*
'* @param fun As Func(Of U, T, U)
'* @param seedv As Variant(Of U)
'* @return Variant(Of U)
'******************************************************************************
Public Function Fold(ByVal fun As Func, Optional ByVal seedv As Variant) As Variant
    Fold = Ext.ArrFold(fun, Me.ToArray, seedv)
End Function

'******************************************************************************
'* [�T  �v] Scan
'* [��  ��] �z��̊e�v�f�ɑ΂��ČJ��Ԃ��w�肵�����Z�p�t�@���N�V������K�p���A
'*          �ݐς����v�Z���ʂ̃��X�g��Ԃ��B
'*          Ariawase��Ext���W���[�����b�p�[���\�b�h�B
'*
'* @param fun As Func(Of U, T, U)
'* @param seedv As Variant(Of U)
'* @return ArrayEx(Of Array(Of U))
'******************************************************************************
Public Function Scan(ByVal fun As Func, Optional ByVal seedv As Variant) As ArrayEx
    Dim arr: arr = Ext.ArrScan(fun, Me.ToArray, seedv)
    Set Scan = Init(arr)
End Function

'******************************************************************************
'* [�T  �v] Unfold
'* [��  ��] Fold�̋t�B�����l����w�肵�����Z�p�t�@���N�V������K�p�����X�g��
'*          ��������B
'*          Ariawase��Ext���W���[�����b�p�[���\�b�h�B
'*
'* @param fun As Func
'* @param seedv As Variant(Of T)
'* @return ArrayEx(Of Array(Of U))
'******************************************************************************
Public Function Unfold(ByVal fun As Func, ByVal seedv As Variant) As ArrayEx
    Dim arr: arr = Ext.ArrUnfold(fun, seedv)
    Set Unfold = Init(arr)
End Function

'******************************************************************************
'* [�T  �v] CopyTo2DArrayRow
'* [��  ��] 2�����z��̎w��s�ɑ΂��āA1�����z��̒l��ݒ肷��B
'*
'* @param v2DArr 2�����z��
'* @param rowNum �w��s
'* @return Variant �ύX��2�����z��
'******************************************************************************
Public Function CopyTo2DArrayRow(ByRef v2DArr As Variant, ByVal rownum As Long) As Variant
    CopyTo2DArrayRow = ArrayUtils.Copy1DArrayTo2DArrayRow(v2DArr, rownum, Me.ToArray)
End Function

'******************************************************************************
'* [�T  �v] Contains
'* [��  ��] �w�肵�����ڂ��܂܂�邩�`�F�b�N����B
'*
'* @param itm ����
'* @return Boolean �`�F�b�N���ʁiTrue�F�܂ށj
'******************************************************************************
Public Function Contains(ByVal itm As Variant) As Boolean
    Contains = ArrayUtils.Contains(Me.ToArray, itm)
End Function

'******************************************************************************
'* [�T  �v] Clone
'* [��  ��] ArrayEx�I�u�W�F�N�g�𕡐�����B
'*
'* @return ArrayEx �V����ArrayEx
'******************************************************************************
Public Function Clone() As ArrayEx
    Dim ax As New ArrayEx
    Set Clone = ax.AddAll(Me.ToArray)
End Function

'******************************************************************************
'* [�T  �v] Insert
'* [��  ��] ���ڂ��w��ʒu�ɒǉ�����B
'*
'* @param itm ����
'* @return ArrayEx
'******************************************************************************
Public Function Insert(ByVal idx As Long, ByVal itm As Variant) As ArrayEx
    Set Insert = Init(ArrayUtils.Insert(Me.ToArray, idx, itm))
End Function

'******************************************************************************
'* [�T  �v] InsertAll
'* [��  ��] ���ځi�����j���w��ʒu�ɒǉ�����B
'*
'* @param itms ����
'* @return ArrayEx
'******************************************************************************
Public Function InsertAll(ByVal idx As Long, ParamArray itms() As Variant) As ArrayEx
    Dim tmp: tmp = itms
    Set InsertAll = Init(ArrayUtils.InsertAll(Me.ToArray, idx, tmp))
End Function

'******************************************************************************
'* [�T  �v] LastIndexOf
'* [��  ��] �w�肵���l�ƈ�v����v�f�̃C���f�b�N�X��ԋp����B
'*          �Ō�����猟�����čŏ��Ɉ�v�������̂�ԋp�B
'*          �l�̑{���J�n�ʒu�Ɨv�f����C�ӎw��\�B
'*
'* @param val As Variant(Of T)
'* @param ixStart As Variant(Of Long)
'* @param cnt As Variant(Of Long)
'* @return As Long
'******************************************************************************
Public Function LastIndexOf(ByVal val As Variant, _
    Optional ByVal ixStart As Variant, Optional ByVal cnt As Variant _
    ) As Long
    LastIndexOf = ArrayUtils.LastIndexOf(Me.ToArray, val, ixStart, cnt)
End Function

'******************************************************************************
'* [�T  �v] IndexListOf
'* [��  ��] 1�����z�񂩂�w�肵���l�ƈ�v����v�f�̃C���f�b�N�X���X�g��ԋp����B
'*
'* @param val �l
'* @param reverseOrder ���X�g�̕��т��t���ɂ��邩
'* @return Variant �C���f�b�N�X���X�g�i1�����z��j
'******************************************************************************
Public Function IndexListOf(ByVal val As Variant, Optional ByVal reverseOrder As Boolean = False) As Variant
    IndexListOf = ArrayUtils.IndexListOf(Me.ToArray, val, reverseOrder)
End Function

'******************************************************************************
'* [�T  �v] Remove
'* [��  ��] 1�����z��̍ŏ��Ɍ�����������̗v�f���폜����B
'*
'* @param itm ����
'* @return ArrayEx �폜��1�����z��
'******************************************************************************
Public Function Remove(ByVal itm As Variant) As ArrayEx
    Set Remove = Init(ArrayUtils.Remove(Me.ToArray, itm))
End Function

'******************************************************************************
'* [�T  �v] RemoveAt
'* [��  ��] �w�肵���C���f�b�N�X�ɂ���v�f���폜����B
'*
'* @param idx �C���f�b�N�X
'* @return ArrayEx �폜��1�����z��
'******************************************************************************
Public Function RemoveAt(ByVal idx As Long) As ArrayEx
    Set RemoveAt = Init(ArrayUtils.RemoveAt(Me.ToArray, idx))
End Function

'******************************************************************************
'* [�T  �v] RemoveRange
'* [��  ��] 1�����z��̗v�f�͈̔͂��폜����B
'*
'* @param idx1 �J�n�C���f�b�N�X
'* @param idx2 �I���C���f�b�N�X
'* @return ArrayEx �폜��1�����z��
'******************************************************************************
Public Function RemoveRange(ByVal idx1 As Long, ByVal idx2 As Long) As ArrayEx
    Set RemoveRange = Init(ArrayUtils.RemoveRange(Me.ToArray, idx1, idx2))
End Function

'******************************************************************************
'* [�T  �v] Repeat
'* [��  ��] �w�肵���v�f���w��񐔕��ǉ�����1�����z����쐬����B
'*
'* @param itm ����
'* @param count ����
'* @return Variant 1�����z��
'******************************************************************************
Public Function Repeat(ByVal itm As Variant, ByVal Count As Long) As ArrayEx
    Set Repeat = Init(ArrayUtils.Repeat(itm, Count))
End Function

'******************************************************************************
'* [�T  �v] GetFirst
'* [��  ��] �ŏ��̗v�f���擾����B
'*
'* @return Variant �ŏ��̗v�f
'******************************************************************************
Public Function GetFirst() As Variant
    If IsEmptyArray Then Exit Function
    If IsObject(xItems(0)) Then
        Set GetFirst = xItems(0)
    Else
        GetFirst = xItems(0)
    End If
End Function

'******************************************************************************
'* [�T  �v] GetLast
'* [��  ��] �Ō�̗v�f���擾����B
'*
'* @return Variant �Ō�̗v�f
'******************************************************************************
Public Function GetLast(ByRef vArr As Variant) As Variant
    If IsEmptyArray Then Exit Function
    If IsObject(xItems(xIndex)) Then
        Set GetLast = xItems(xIndex)
    Else
        GetLast = xItems(xIndex)
    End If
End Function

'******************************************************************************
'* [�T  �v] Join
'* [��  ��] �z���A�����ĕ�����ɂ���B
'*
'* @param delimiter �f���~�^
'* @param wrapChar �͂ݕ���
'* @param rmLineSeparator ���s���������邩
'* @return String �A��������
'******************************************************************************
Public Function Join(Optional ByVal Delimiter As String, Optional ByVal wrapChar As String, Optional ByVal rmLineSeparator As Boolean = True) As String
    Join = ArrayUtils.Join(Me.ToArray, Delimiter, wrapChar, rmLineSeparator)
End Function

'******************************************************************************
'* [�T  �v] Search
'* [��  ��] �z��i������j����w�肵��������ƕ�����v���鍀�ڂ�
'*          �擾���A1�����z��ŕԋp����B
'*
'* @param searchVal ����������i�z��ŕ����w��j
'* @return Variant �������ʁi1�����z��j
'******************************************************************************
Public Function Search(ByVal searchVal As Variant) As ArrayEx
    Set Search = Core.Init(New ArrayEx, ArrayUtils.Search(Me.ToArray, searchVal))
End Function

'******************************************************************************
'* [�T  �v] RegexSearch
'* [��  ��] �z��i������j����w�肵��������Ɛ��K�\������v���鍀�ڂ�
'*          �擾���A1�����z��ŕԋp����B
'*
'* @param searchStr ����������i���K�\���j
'* @param igCase �啶���������̋��
'* @return Variant �������ʁi1�����z��j
'******************************************************************************
Public Function RegexSearch(ByVal searchStr As String, Optional ByVal igCase As Boolean = False) As ArrayEx
    Set RegexSearch = Core.Init(New ArrayEx, ArrayUtils.RegexSearch(Me.ToArray, searchStr, igCase))
End Function

'******************************************************************************
'* [�T  �v] ToString
'* [��  ��] JSON������ɕϊ�����B
'*
'* @return String JSON������
'******************************************************************************
Public Function ToString() As String
    ToString = ArrayUtils.ToString(Me.ToArray)
End Function