Attribute VB_Name = "Main"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] Http���N�G�X�g���M���C���@�\
'* [��  ��] Http���N�G�X�g���M���C����������������B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* [�T  �v] Http���N�G�X�g�𑗐M���A���X�|���X���擾�E�ԋp����B
'* [��  ��] �uWinHttp.WinHttpRequest.5.1�v���g�p���āAHttp���N�G�X�g�𑗐M
'*          ����B
'*
'* @param method         ���\�b�h�iGET,POST,PUT,DELETE,HEAD,OPTIONS,PATCH�j
'* @param url            �A�N�Z�X����URL�i��Fhttp://localhost:8080/api/v2/test�j
'* @param reqHdrParams() Header�p�����[�^(2�����z��)
'*                       ��2�C���f�b�N�X �c 0:name, 1:value
'* @param reqBody        Body
'* @param lTimeout       �^�C���A�E�g���ԁi�~���b�j�@���C��
'* @param resEncode      ���X�|���X�G���R�[�h�@���C��
'* @return Response���
'*
'******************************************************************************
Public Function SendRequest(method As String, url As String, reqHdrParams() As String, reqBody As String, _
                            Optional lTimeout As Long = 0, _
                            Optional resEncode As String = "utf-8") As Response
    Dim winHttp As Variant
    Set winHttp = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' HTTP ���N�G�X�g
    winHttp.Open method, url, False
    
    ' ���N�G�X�g�w�b�_�ݒ�
    Dim i As Long
    If GetDimension(reqHdrParams) <> 0 Then
        For i = 0 To UBound(reqHdrParams, 1)
            winHttp.setRequestHeader reqHdrParams(i, 0), reqHdrParams(i, 1)
        Next i
    End If

    ' �^�C���A�E�g�ݒ�
    If lTimeout > 0 Then
        winHttp.SetTimeouts lTimeout, lTimeout, lTimeout, lTimeout
    End If
    
    ' ���N�G�X�g���M
    If method = "GET" Or method = "DELETE" Or method = "HEAD" Or method = "OPTIONS" Then
        winHttp.send
    Else
        winHttp.send reqBody
    End If
    
    ' �X�e�[�^�X�ƃ��X�|���X�{�f�B�̎擾
    Dim status As String:     status = winHttp.status
    Dim statusTxt As String:  statusTxt = winHttp.StatusText
    Dim resHeaders As String: resHeaders = winHttp.GetAllResponseHeaders
    Dim resBody As String
    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adReadAll = -1
    With CreateObject("ADODB.Stream")
        .Type = adTypeBinary
        .Open
        .Write winHttp.ResponseBody
        .Position = 0
        .Type = adTypeText
        .Charset = resEncode
        resBody = .ReadText(adReadAll)
        .Close
    End With
  
    ' ���X�|���X�ݒ�
    Dim res As Response: Set res = New Response
    res.StatusCd = status
    res.StatusText = statusTxt
    res.Body = resBody
    res.Headers = resHeaders
    
    Set SendRequest = res

End Function

'******************************************************************************
'* [�T  �v] �����擾�����B
'* [��  ��] �w�肵���z�񂪉������̔z�񂩔��肷��B
'*
'* @param targetArray �ΏۂƂȂ�z��
'* @return ������
'******************************************************************************
Private Function GetDimension(targetArray As Variant) As Long
    Dim dimention As Long
    Dim tmp As Long
    
    dimention = 1
    On Error Resume Next
    While Err.Number = 0
        tmp = UBound(targetArray, dimention)
        dimention = dimention + 1
    Wend
    GetDimension = dimention - 2
End Function

'******************************************************************************
'* [�T  �v] URL�G���R�[�h�����iUTF-8�j�B
'* [��  ��] URL�G���R�[�h���s���B
'*
'* @param target �ΏۂƂȂ镶����
'* @return �G���R�[�h�㕶����
'******************************************************************************
Public Function EncodeURL(target As String) As String
    Dim xlBit As String
    xlBit = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    '32Bit��Excel�̏ꍇ��ScriptControl�Ŏg�p
    If xlBit = "x86" Then
        EncodeURL = EncodeUrlFor32(target)
    Else
        ' 64Bit�ł�Excel2010�ȑO�̏ꍇ���O�����ł��g�p
        Dim ver As Integer: ver = CInt(Application.Version)
        If ver <= 14 Then
            EncodeURL = EncodeUrlMyP(target)
            'PowerShell�ł͎g�p���Ȃ�
            'EncodeURL = EncodeURLByPs(target)
        ' 64Bit�ł�Excel2013�ȍ~�̏ꍇ��Excel�֐��g�p
        Else
            EncodeURL = EncodeUrlForXl2013OrLater(target)
        End If
    End If

End Function

'******************************************************************************
'* [�T  �v] URL�G���R�[�h�����B
'* [��  ��] URL�G���R�[�h���s���BExcel2010�ȑO��64bit�p�BPowerShell���g�p����B
'*
'* @param target �ΏۂƂȂ镶����
'* @param pEncode encode�w��ishift-jis���j
'* @return �G���R�[�h�㕶����
'******************************************************************************
Private Function EncodeURLByPs(target As String, Optional pEncode As String = "utf-8") As String
    Dim cmd
    Dim i As Long, c As Long
    Dim tmp As String, flg As Boolean
    flg = False
    
    For i = 1 To Len(target)
        tmp = Mid(target, i, 1)
        c = Asc(tmp)
        Select Case tmp
            ' �\�񕶎��̏ꍇ
            Case ":", "/", "?", "#", "[", "]", "@", "!", "$", "&", "'", "(", ")", "*", "+", ",", ";", "="
                flg = True
                Exit For
        End Select
        If c < 33 Or c > 127 Then
                flg = True
                Exit For
        End If
    Next
    ' TODO�F�����������Ɠr���Ő؂��
    If flg Then
        target = Replace(target, "`", "``")
        target = Replace(target, """", "`""""""")
        cmd = "PowerShell -Command ""[void]([Reflection.Assembly]::LoadWithPartialName(""""""System.Web""""""));[Web.HttpUtility]::UrlEncode(""""""" & target & """"""", [Text.Encoding]::GetEncoding(""""""" & pEncode & """""""))"""
        With CreateObject("WScript.Shell").Exec(cmd)
            Do While (.status = 0)
                DoEvents
            Loop
            EncodeURLByPs = .StdOut.ReadLine
        End With
        Exit Function
    End If
    EncodeURLByPs = target
End Function

'******************************************************************************
'* [�T  �v] URL�G���R�[�h�����B
'* [��  ��] URL�G���R�[�h���s���B
'*          ����ScriptControl���֐����g�p�����A���O�ŃG���R�[�h���Ă���̂�
'*          �ǂ�Excel�o�[�W�����E32bit/64bit�ł��g�p�\�B
'*
'* @param target �ΏۂƂȂ镶����
'* @return �G���R�[�h�㕶����
'******************************************************************************
Private Function EncodeUrlMyP(target As String) As String
    Dim buf() As Byte, s As String, i As Long
    
    With CreateObject("ADODB.Stream")
        .Mode = 3 'adModeReadWrite
        .Open
        .Type = 2 'adTypeText
        .Charset = "UTF-8"
        .WriteText target
        
        .Position = 0
        .Type = 1 'adTypeBinary
        .Position = 3 'BOM��΂�
        buf = .Read
        .Close
    End With

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
    EncodeUrlMyP = s
End Function

'******************************************************************************
'* [�T  �v] URL�G���R�[�h�����B
'* [��  ��] URL�G���R�[�h���s���BExcel2013�ȍ~�Ŏg�p�\�B
'*
'* @param target �ΏۂƂȂ镶����
'* @return �G���R�[�h�㕶����
'******************************************************************************
Private Function EncodeUrlForXl2013OrLater(target As String) As String
    EncodeUrlForXl2013OrLater = Application.WorksheetFunction.EncodeURL(target)
End Function

'******************************************************************************
'* [�T  �v] URL�G���R�[�h�����B
'* [��  ��] URL�G���R�[�h���s���B64bit�ł͎g���Ȃ��B
'*
'* @param target �ΏۂƂȂ镶����
'* @return �G���R�[�h�㕶����
'******************************************************************************
Private Function EncodeUrlFor32(target As String) As String
    Dim sc As Object

    Set sc = CreateObject("ScriptControl")
    sc.Language = "Jscript"
    EncodeUrlFor32 = sc.CodeObject.encodeURIComponent(target)
    Set sc = Nothing

End Function

