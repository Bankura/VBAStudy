VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VariantUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] Variant�֘A���[�e�B���e�B�N���X
'* [��  ��] Variant������s�����[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'* [�Q  �l]
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
'* �ϐ���`
'******************************************************************************

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************
'******************************************************************************
'* [�T  �v] EmptyStringIfMissing
'* [��  ��] Optional�������ȗ�����Ă���ꍇ�A�󕶎���ɕϊ�����
'*
'* @param arg �����iOptional������n���j
'* @return String �ϊ��㕶����
'*
'******************************************************************************
Public Function EmptyStringIfMissing(ByVal arg) As String
    If IsMissing(arg) Then
        EmptyStringIfMissing = ""
    End If
    EmptyStringIfMissing = arg
End Function

'******************************************************************************
'* [�T  �v] DefaultIfMissing
'* [��  ��] Optional�������ȗ�����Ă���ꍇ�A�f�t�H���g�l�ɕϊ�����
'*
'* @param arg �����iOptional������n���j
'* @param defaultValue �f�t�H���g�l
'* @return Variant �ϊ���Variant
'*
'******************************************************************************
Public Function DefaultIfMissing(ByVal arg, ByVal defaultValue) As Variant
    If IsMissing(arg) Then
        DefaultIfMissing = defaultValue
    End If
    DefaultIfMissing = arg
End Function


'*******************************************************************************
'* [�T  �v] ToJsonString
'* [��  ��] �I�u�W�F�N�g�A�v���~�e�B�u�^�𕶎���ϊ�����B
'*          JSON�`�����ӎ�����������ɕϊ����邪�A���S��JSON�`��
'*          �ɏ����͂��Ă��Ȃ��B
'*
'* @param txt Variant�^�f�[�^
'* @return String ���ʕ�����
'*
'*******************************************************************************
Public Function ToString(ByVal var As Variant) As String
    If IsObject(var) Then
        ToString = ObjectToString(var)
        
    ElseIf IsArray(var) Then
        ToString = ArrayUtils.ToString(var)
        
    ElseIf ValidateUtils.IsString(var) Then
        ToString = """" & CStr(var) & """"
        
    ElseIf IsEmpty(var) Then
        ToString = "Empty"
        
    ElseIf IsNull(var) Then
        ToString = "Null"
        
    ElseIf IsError(var) Then
        ToString = "Error"
        
    Else
        ToString = CStr(var)
        
    End If
End Function

'*******************************************************************************
'* [�T  �v] ObjectToString
'* [��  ��] �I�u�W�F�N�g�𕶎���ϊ�����B
'*          JSON�`�����ӎ�����������ɕϊ����邪�A���S��JSON�`��
'*          �ɏ����͂��Ă��Ȃ��B
'*
'* @param var Variant�^�f�[�^
'* @return String ���ʕ�����
'*
'*******************************************************************************
Public Function ObjectToString(ByVal var As Variant) As String
    If IsObject(var) Then
        If ValidateUtils.IsNothing(var) Then
            ObjectToString = "Nothing"
        
        ElseIf ValidateUtils.IsArrayEx(var) Then
            ObjectToString = ArrayUtils.ToString(var.ToArray)
        
        ElseIf ValidateUtils.IsStringEx(var) Then
            ObjectToString = """" & var.ToString & """"
        
        ElseIf ValidateUtils.IsCollection(var) Then
            ObjectToString = CollectionUtils.CollectionToString(var)
    
        ElseIf ValidateUtils.IsDictionary(var) Or ValidateUtils.IsDictionaryEx(var) Then
            ObjectToString = CollectionUtils.DictionaryToString(var)
        
        Else
            ObjectToString = "(" & TypeName(var) & ")"
        End If
    Else
        ObjectToString = "(no obj)"
    End If
End Function

