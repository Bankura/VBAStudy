VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] �R���N�V�������[�e�B���e�B�N���X
'* [��  ��] �R���N�V�����֘A�̃��[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'*
'* [�ˁ@��] VBA-Web v4.1.6
'*          Ariawase 0.9.0
'* [�Q  �l]
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum��`
'******************************************************************************

'******************************************************************************
'* �萔��`
'******************************************************************************
#If Win64 Then
    Private Const FOWARD_SHIFT As Long = 40
    Private Const BACKWARD_SHIFT As Long = 48
    Private Const KEY_POS As Long = 24
#Else
    Private Const FOWARD_SHIFT As Long = 24
    Private Const BACKWARD_SHIFT As Long = 28
    Private Const KEY_POS As Long = 16
#End If

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
'* [�T  �v] Push
'* [��  ��] Collection�ɒl��Push����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param clct As Collection(Of T)
'* @param val As Variant(Of T)
'******************************************************************************
Public Sub Push(ByVal clct As Collection, ByVal val As Variant)
    Call Core.Push(clct, val)
End Sub

'******************************************************************************
'* [�T  �v] Pop
'* [��  ��] Collection����l��Pop����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param clct As Collection(Of T)
'* @return As Variant(Of T)
'******************************************************************************
Public Function Pop(ByVal clct As Collection) As Variant
    Pop = Core.Pop(clct)
End Function

'******************************************************************************
'* [�T  �v] Shift
'* [��  ��] Collection�ɒl��Shift����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param clct As Collection(Of T)
'* @param val As Variant(Of T)
'******************************************************************************
Public Sub Shift(ByVal clct As Collection, ByVal val As Variant)
    Call Core.Shift(clct, val)
End Sub

'******************************************************************************
'* [�T  �v] Unshift
'* [��  ��] Collection����l��Unshift����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param clct As Collection(Of T)
'* @return As Variant(Of T)
'******************************************************************************
Public Function Unshift(ByVal clct As Collection) As Variant
    Unshift = Core.Unshift(clct)
End Function

'******************************************************************************
'* [�T  �v] ArrayToCollection
'* [��  ��] �ꎟ���z���Collection�I�u�W�F�N�g�ɕϊ�����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param arr As Variant(Of Array(Of T))
'* @return As Collection(Of T)
'******************************************************************************
Public Function ArrayToCollection(ByVal arr As Variant) As Collection
    Set ArrayToCollection = Core.ArrToClct(arr)
End Function

'******************************************************************************
'* [�T  �v] CollectionToArray
'* [��  ��] Collection�I�u�W�F�N�g���ꎟ���z��ɕϊ�����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param clct As Collection(Of T)
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function CollectionToArray(ByVal clct As Collection) As Variant
    CollectionToArray = Core.ClctToArr(clct)
End Function

'******************************************************************************
'* [�T  �v] AssocArrayToDictionary
'* [��  ��] �A�z�z���Dictionary�I�u�W�F�N�g�ɕϊ�����B
'*          Ariawase��Ext���W���[�����b�p�[���\�b�h�B
'*
'* @param aarr Variant(Of Array(Of Tupple))
'* @return Object(Of Dictionary)
'******************************************************************************
Public Function AssocArrayToDictionary(ByVal aarr As Variant) As Object
    Set AssocArrayToDictionary = Ext.AssocArrToDict(aarr)
End Function

'******************************************************************************
'* [�T  �v] DictionaryToAssocArray
'* [��  ��] Dictionary�I�u�W�F�N�g��A�z�z��ɕϊ�����B
'*          Ariawase��Ext���W���[�����b�p�[���\�b�h�B
'*
'* @param dict
'* @return Variant(Of Array(Of Tupple))
'******************************************************************************
Public Function DictionaryToAssocArray(ByVal dict As Object) As Variant
    DictionaryToAssocArray = Ext.DictToAssocArr(dict)
End Function

'******************************************************************************
'* [�T  �v] HasKey
'* [��  ��] Collection�ɃL�[�ɑΉ�����l�����邩�`�F�b�N�
'*
'* @param col ��������Collection
'* @param keyName �L�[��
'* @return Boolean ���茋��(True:����)
'*
'******************************************************************************
Public Function HasKey(ByRef col As Collection, ByVal keyName As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = col(keyName)
    HasKey = (err.Number = 0)
    err.Clear
End Function

'******************************************************************************
'* [�T  �v] GetCollectionKeyByIndex ���\�b�h
'* [��  ��] �R���N�V�����̃L�[��Index���w�肵�Ď擾����B
'* [�Q  �l] <https://foolexp.wordpress.com/2012/12/03/p2wpsu-5k/>
'*
'* @param index �C���f�b�N�X
'* @param col �R���N�V����
'* @return �L�[��
'*
'******************************************************************************
Public Function GetCollectionKeyByIndex(ByVal Index As Long, col As Collection) As String
    Dim ptr As LongPtr

    Dim i As LongPtr
    Dim sKey As String
    If col Is Nothing Then
        err.Raise 91
    Else
        Select Case Index
            Case Is > col.count
                err.Raise 9
            Case Is <= col.count / 2
                Call mWinApi.CopyMemoryByRef(ptr, ObjPtr(col) + FOWARD_SHIFT, LONGPTR_SIZE, False, True)
                For i = 2 To Index
                    Call mWinApi.CopyMemoryByRef(ptr, ptr + FOWARD_SHIFT, LONGPTR_SIZE, False, True)
                Next i
            Case Else
                Call mWinApi.CopyMemoryByRef(ptr, ObjPtr(col) + BACKWARD_SHIFT, LONGPTR_SIZE, False, True)
                For i = col.count - 1 To Index Step -1
                    Call mWinApi.CopyMemoryByRef(ptr, ptr + BACKWARD_SHIFT, LONGPTR_SIZE, False, True)
                Next i
        End Select
        i = StrPtr(sKey)
        Call mWinApi.CopyMemoryByRef(VarPtr(sKey), ptr + KEY_POS, LONGPTR_SIZE, True, True)
        
        GetCollectionKeyByIndex = sKey
        Call mWinApi.CopyMemoryByRef(VarPtr(sKey), i, LONGPTR_SIZE, True, False)
    End If
End Function

'******************************************************************************
'* [�T  �v] GetCollectionIndexByKey ���\�b�h
'* [��  ��] �R���N�V������Index���L�[���w�肵�Ď擾����B
'* [�Q  �l] <https://foolexp.wordpress.com/2012/12/03/p2wpsu-5k/>
'*
'* @param key �L�[
'* @param col �R���N�V����
'* @return Index
'*
'******************************************************************************
Public Function GetCollectionIndexByKey(ByVal Key As String, col As Collection, Optional ByVal comp As VbCompareMethod = vbTextCompare) As Long
    Dim ptr   As LongPtr
    Dim sKey  As String
    Dim aKey  As LongPtr
 
    If col Is Nothing Then
        err.Raise 91
    Else
        If col.count Then
            'save StrPtr
            aKey = StrPtr(sKey)
            'first Ptr
            Call mWinApi.CopyMemoryByRef(ptr, ObjPtr(col) + FOWARD_SHIFT, LONGPTR_SIZE, False, True)
            GetCollectionIndexByKey = 1
            
            ' walk items upwards from first
            Do
                Call mWinApi.CopyMemoryByRef(VarPtr(sKey), ptr + KEY_POS, LONGPTR_SIZE, True, True)
                If StrComp(Key, sKey, comp) = 0 Then
                    'found
                    Exit Do
                End If
                'next Index
                GetCollectionIndexByKey = GetCollectionIndexByKey + 1
                
                ' �������[�v�΍�
                If GetCollectionIndexByKey > col.count Then
                    GetCollectionIndexByKey = -1
                    Exit Function
                End If
                
                ' next Ptr
                Call mWinApi.CopyMemoryByRef(ptr, ptr + FOWARD_SHIFT, LONGPTR_SIZE, False, True)
            Loop Until ptr = 0
            
            ' restore original StrPtr
            Call mWinApi.CopyMemoryByRef(VarPtr(sKey), aKey, LONGPTR_SIZE, True, False)
        End If
        If ptr = 0 Then
            'key not found
            GetCollectionIndexByKey = -1
        End If
    End If
End Function

'*******************************************************************************
'* [�T  �v] CollectionToString
'* [��  ��] Collection�𕶎���ϊ�����B
'*          VBA-WEB����VBA-JSON�𗘗p�B
'*
'* @param var Collection�^�f�[�^
'* @return String ���ʕ�����
'*
'*******************************************************************************
Public Function CollectionToString(ByVal col As Collection) As String
    CollectionToString = WebHelpers.ConvertToJson(col)
End Function

'*******************************************************************************
'* [�T  �v] ObjectToString
'* [��  ��] Dictionary�𕶎���ϊ�����B
'*          VBA-WEB����VBA-JSON�𗘗p�B
'*
'* @param dic Dictionary�^�f�[�^
'* @return String ���ʕ�����
'*
'*******************************************************************************
Public Function DictionaryToString(ByVal dic As Object) As String
    DictionaryToString = WebHelpers.ConvertToJson(dic)
End Function