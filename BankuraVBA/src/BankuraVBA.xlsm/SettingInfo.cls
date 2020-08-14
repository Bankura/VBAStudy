VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] �ݒ���N���X
'* [��  ��] �ݒ����ێ����邽�߂̃N���X�B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* �����萔��`
'******************************************************************************
Private Const SETTING_SHEET_NAME As String = "setting"
Private Const SETTING_SH_START_ROW As Long = 4
Private Const SETTING_SH_START_COL As Long = 4

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mMasters As Object 'Dictionary
Private mSetting As Object 'Dictionary

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************
'�Ȃ�

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mSetting = CreateObject("Scripting.Dictionary")
    Call GetSettingInfo
    Set mMasters = CreateObject("Scripting.Dictionary")
    Call GetMasterInfo
End Sub
Private Sub Class_Terminate()
    Set mMasters = Nothing
    Set mSetting = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] �ݒ���擾�����B
'* [��  ��] �ݒ����ݒ�V�[�g���擾���A�ێ����܂��
'*
'******************************************************************************
Private Sub GetSettingInfo()
    Dim wb As Workbook, ws As Worksheet
    Set wb = Application.ThisWorkbook
    Set ws = wb.Sheets(SETTING_SHEET_NAME)
    Set mSetting = GetMasterDataFromSheet(ws, SETTING_SH_START_ROW, SETTING_SH_START_COL, 2)
End Sub

'******************************************************************************
'* [�T  �v] �}�X�^���ꊇ�擾�����B
'* [��  ��] �}�X�^�����ꊇ�擾���A�ێ����܂��
'*
'******************************************************************************
Private Sub GetMasterInfo()
    Dim mstHeadLetters As String: mstHeadLetters = GetSettingValue("MasterHeadLetters")
    Dim StartRow As Long: StartRow = GetSettingValue("MasterStartRowNo")
    Dim StartCol As Long: StartCol = GetSettingValue("MasterStartColNo")
    
    Dim reg As Object
    Set reg = GetRegExp()
    With reg
        .Pattern = "^" & mstHeadLetters & ".*"
        .ignorecase = True
        .Global = True
    End With
    
    Dim wb As Workbook, ws As Worksheet
    Set wb = Application.ThisWorkbook
    For Each ws In wb.Worksheets
        If reg.Test(ws.Name) Then
            Dim master As Object
            Set master = GetMasterDataFromSheet(ws, StartRow, StartCol, 2)
            Call mMasters.Add(ws.Name, master)
        End If
    Next
End Sub

'******************************************************************************
'* [�T  �v] �}�X�^���擾�����B
'* [��  ��] �}�X�^�����擾���A�����f�B�N�V���i���ɕێ����܂��
'*
'* @param dataSheet ���[�N�V�[�g
'* @param lStartRow �f�[�^�J�n�s�ԍ�
'* @param lStartCol �f�[�^�J�n��ԍ�
'* @param itemCount ���ڗ�
'*
'******************************************************************************
Private Function GetMasterDataFromSheet(dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, Optional colCount As Long) As Object
    Dim varr: varr = XlWorkSheetUtils.GetVariantDataFromSheet(dataSheet, lStartRow, lStartCol, colCount)
    If IsEmpty(varr) Then
        Exit Function
    End If
    Dim master As Object: Set master = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(varr, 1) To UBound(varr, 1)
        Dim code As String: code = varr(i, 1)
        Dim val As String: val = varr(i, 2)
        Call master.Add(code, val)
    Next
    Set GetMasterDataFromSheet = master
End Function

'******************************************************************************
'* [�T  �v] �ݒ�l�擾�����B
'* [��  ��] �ݒ��񂩂�ݒ�l���擾���܂��
'*
'* @param key �L�[�l
'* @return �ݒ�l
'*
'******************************************************************************
Public Function GetSettingValue(ByVal Key As String) As String
    If mSetting.Exists(Key) Then
        GetSettingValue = mSetting.Item(Key)
    Else
        GetSettingValue = ""
    End If
End Function

'******************************************************************************
'* [�T  �v] �}�X�^�擾�����B
'* [��  ��] �}�X�^��񂩂�}�X�^�I�u�W�F�N�g���擾���܂��
'*
'* @param masterName �L�[�l�i�}�X�^�V�[�g���j
'* @return �}�X�^�I�u�W�F�N�g
'*
'******************************************************************************
Public Function GetMaster(ByVal masterName As String) As Object
    If mMasters.Exists(masterName) Then
        Set GetMaster = mMasters.Item(masterName)
    Else
        Set GetMaster = Nothing
    End If
End Function

'******************************************************************************
'* [�T  �v] �}�X�^�l�擾�����B
'* [��  ��] �}�X�^��񂩂�}�X�^���E�R�[�h�l�i�L�[�j���w�肵�Ēl���擾���܂��
'*
'* @param masterName �}�X�^�V�[�g��
'* @param code �L�[�l�i�R�[�h�j
'* @return �l
'*
'******************************************************************************
Public Function GetMasterValueByCode(ByVal masterName As String, ByVal code As String) As String
    Dim master As Object
    If mMasters.Exists(masterName) Then
        Set master = mMasters.Item(masterName)
        If master.Exists(code) Then
            GetMasterValueByCode = master.Item(code)
            Exit Function
        End If
    End If
    GetMasterValueByCode = ""
End Function

'******************************************************************************
'* [�T  �v] �}�X�^�R�[�h�l�擾�����B
'* [��  ��] �}�X�^��񂩂�}�X�^���E�l���w�肵�ăR�[�h�l���擾���܂��
'*          �O��Ƃ��āA�l����Ӂi�d���l�Ȃ��j�ł���K�v������܂��B
'*
'* @param masterName �}�X�^�V�[�g��
'* @param val �L�[�l�i�R�[�h�j
'* @return �R�[�h�l
'*
'******************************************************************************
Public Function GetMasterCodeByValue(ByVal masterName As String, ByVal val As String) As String
    Dim master As Object, vKey
    If mMasters.Exists(masterName) Then
        Set master = mMasters.Item(masterName)
        For Each vKey In master.keys
            If master.Item(vKey) = val Then
                GetMasterCodeByValue = CStr(vKey)
                Exit Function
            End If
        Next
    End If
    GetMasterCodeByValue = ""
End Function