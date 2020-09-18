Attribute VB_Name = "PsqlExecuter"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] psql���s���C���@�\
'* [��  ��] Http���N�G�X�g���M���C����������������B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* [�T  �v] psql�R�}���h�����SQL�����s����B
'* [��  ��] psql�R�}���h����āASQL�����s���A���s���ʂ�ԋp����B
'*          ���[�J������psql���C���X�g�[������Ă���K�v������B
'*          psql��postgreSQL�C���X�g�[�����ɍ��킹�ăC���X�g�[�������B
'*
'* @param psql     psql�R�}���h�̃t���p�X
'* @param host     �A�N�Z�X����DB�̃z�X�g�i��Flocalhost�j
'* @param port     �A�N�Z�X����DB�̃|�[�g�ԍ�
'* @param dbName   DB��
'* @param userName DB���[�U��
'* @param password DB�p�X���[�h
'* @param sql      SQL��
'* @param clEncode Client�G���R�[�h�@���C��
'* @return Result���
'*
'******************************************************************************
Public Function ExecPsql(psql As String, host As String, port As String, _
                         dbName As String, userName As String, password As String, _
                         sql As String, Optional clEncode As String = "SJIS") As Result
    Dim cmd As String, oExec
    
    ' �R�}���h�g����
    cmd = psql & " -h " & host & " -p " & port & " -d " & dbName & " -U " & userName & " -c """ & sql & """ -A"
    Debug.Print cmd
    
    ' Wscript.Shell�I�u�W�F�N�g����
    With CreateObject("Wscript.Shell")
        ' �p�X���[�h�����ϐ��ɐݒ�
        .Environment("Process").Item("PGPASSWORD") = password
    
        ' Client(psql)�̃G���R�[�h�ݒ�F�f�t�H���g�uSJIS�v
        .Environment("Process").Item("PGCLIENTENCODING") = clEncode
    
        ' �R�}���h���s
        Set oExec = .Exec(cmd)
    End With
    

    ' ���s���ʐݒ�
    Dim res As Result: Set res = New Result
    res.ExitCd = oExec.ExitCode
    If oExec.ExitCode <> 0 Then
        ' �R�}���h���s������
        Dim errTxt As String: errTxt = oExec.StdErr.ReadAll
        res.StdErrTxt = errTxt
        Set ExecPsql = res
        Exit Function
    End If

    ' ���펞����
    ' �f�[�^�ʌv��
    Dim lRowMax As Long: lRowMax = 0
    Dim lColMax As Long: lColMax = 0
    
    Dim stdOutTxt As String
    While Not oExec.stdOut.AtEndOfStream
        Dim strTmp As String: strTmp = oExec.stdOut.ReadLine
        If strTmp <> "" Then
            If lRowMax = 0 Then
                Dim vTmpCols As Variant: vTmpCols = Split(strTmp, "|")
                lColMax = UBound(vTmpCols) - LBound(vTmpCols) + 1
                stdOutTxt = strTmp
            Else
                stdOutTxt = stdOutTxt & vbCrLf & strTmp
            End If
            lRowMax = lRowMax + 1
        End If
    Wend
    '�f�[�^�ݒ�
    Dim vArray()
    ReDim vArray(0 To lRowMax - 1, 0 To lColMax - 1)
    Dim i As Long, j As Long, stdOut As Variant, cols As Variant
    stdOut = Split(stdOutTxt, vbCrLf)
    For i = LBound(vArray, 1) To UBound(vArray, 1)
        Debug.Print "Line: " & stdOut(i)
        If stdOut(i) <> "" Then
            If InStr(stdOut(i), "|") = 0 Then
                vArray(i, 0) = stdOut(i)
                For j = LBound(vArray, 2) + 1 To UBound(vArray, 2)
                    vArray(i, j) = ""
                Next j
            Else
                cols = Split(stdOut(i), "|")
                For j = LBound(cols) To UBound(cols)
                    vArray(i, j) = cols(j)
                Next j
            End If

        End If
    Next i
    res.StdOutList = vArray
    res.RowMax = lRowMax
    res.ColMax = lColMax
    
    Set oExec = Nothing
    Set ExecPsql = res

End Function

'******************************************************************************
'* [�T  �v] �J�n����
'* [��  ��] ��ʕ`��̒�~���A�������\�ɉe���̂���ݒ��ύX����B
'*
'******************************************************************************
Public Sub CommonStr()
    '�e��ݒ�ύX
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
End Sub

'******************************************************************************
'* [�T  �v] �I������
'* [��  ��] �J�n�����ōs�����ݒ����������B
'*
'******************************************************************************
Public Sub CommonEnd()
    '�e��ݒ�ύX
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub

'******************************************************************************
'* [�T  �v] �V�[�g���݃`�F�b�N
'* [��  ��] �V�[�g�����݂��邩�𔻒肷��B
'*
'* @param strName �u�b�N��
'* @param wb �u�b�N�I�u�W�F�N�g
'* @return �������ʁiTrue:���݂��� False�F���݂��Ȃ��j
'******************************************************************************
Public Function CheckSheet(strName As String, Optional wb As Workbook) As Boolean
    Dim ws As Worksheet
    Dim flg As Boolean

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    '�V�[�g�̌���
    For Each ws In wb.Worksheets
        If ws.Name = strName Then
            flg = True
            Exit For
        End If
    Next ws
    
    CheckSheet = flg
End Function
