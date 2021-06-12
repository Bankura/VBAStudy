VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] Http�ʐM���[�e�B���e�B�N���X
'* [��  ��] Http�ʐM�Ɋւ��郆�[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'* [�Q  �l]
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* WindowsAPI��`
'******************************************************************************
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As LongPtr) As Long

'******************************************************************************
'* Enum��`
'******************************************************************************

'******************************************************************************
'* �萔��`
'******************************************************************************

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mWinHttp As Object

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
End Sub
Private Sub Class_Terminate()
    Set mWinHttp = Nothing
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************


'******************************************************************************
' [�֐���] DownloadFileByUrl
' [���@��] �w�肵��URL����t�@�C�����_�E�����[�h���A�w�肵���ۑ���i�t�@�C���j
'          �ɕۑ�����B
' [���@��] strFileUrl �_�E�����[�h�Ώۂ�URL
'          strSavePath �ۑ���t�@�C�����i�t���p�X�j
' [�߂�l] Boolean �������ʁiTrue:���� False�F�ُ�j
'******************************************************************************
Public Function DownloadFileByUrl(strFileUrl As String, strSavePath As String) As Boolean

    Dim lngCacheDelRs As Long
    Dim lngDlRs As Long

    ' �L���b�V���N���A
    lngCacheDelRs = DeleteUrlCacheEntry(strFileUrl)
    
    ' URL����t�@�C�����_�E�����[�h���ĕۑ�
    lngDlRs = URLDownloadToFile(0, strFileUrl, strSavePath, 0, 0)

    If lngDlRs <> 0 Then
        DownloadFileByUrl = False
        Exit Function
    End If
    DownloadFileByUrl = True
End Function

'******************************************************************************
'* [�T  �v] ReNewWinHttp ���\�b�h
'* [��  ��] WinHttpRequest�I�u�W�F�N�g���Đ�������B
'*
'******************************************************************************
Private Sub ReNewWinHttp()
    Set mWinHttp = Nothing
    Set mWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
End Sub

'******************************************************************************
'* [�T  �v] Post ���\�b�h
'* [��  ��] Post���M���s���B
'*
'* @param HttpRequest Request���
'* @return Response���
'*
'******************************************************************************
Public Function Post(req As HttpRequest) As HttpResponse
    req.method = "POST"
    Set Post = Transmit(req)
End Function

'******************************************************************************
'* [�T  �v] PutReq ���\�b�h
'* [��  ��] Put���M���s���B
'*
'* @param HttpRequest Request���
'* @return Response���
'*
'******************************************************************************
Public Function PutReq(req As HttpRequest) As HttpResponse
    req.method = "PUT"
    Set PutReq = Transmit(req)
End Function

'******************************************************************************
'* [�T  �v] Patch ���\�b�h
'* [��  ��] Patch���M���s���B
'*
'* @param HttpRequest Request���
'* @return Response���
'*
'******************************************************************************
Public Function Patch(req As HttpRequest) As HttpResponse
    req.method = "PATCH"
    Set Patch = Transmit(req)
End Function

'******************************************************************************
'* [�T  �v] GetReq ���\�b�h
'* [��  ��] Get���M���s���B
'*
'* @param HttpRequest Request���
'* @return Response���
'*
'******************************************************************************
Public Function GetReq(req As HttpRequest) As HttpResponse
    req.method = "GET"
    Set GetReq = Transmit(req)
End Function

'******************************************************************************
'* [�T  �v] Delete ���\�b�h
'* [��  ��] Delete���M���s���B
'*
'* @param ResponseDto Request���
'* @return Response���
'*
'******************************************************************************
Public Function Delete(req As HttpRequest) As HttpResponse
    req.method = "DELETE"
    Set Delete = Transmit(req)
End Function

'******************************************************************************
'* [�T  �v] Head ���\�b�h
'* [��  ��] Head���M���s���B
'*
'* @param HttpRequest Request���
'* @return Response���
'*
'******************************************************************************
Public Function Head(req As HttpRequest) As HttpResponse
    req.method = "HEAD"
    Set Head = Transmit(req)
End Function

'******************************************************************************
'* [�T  �v] OptionsReq ���\�b�h
'* [��  ��] Options���M���s���B
'*
'* @param HttpRequest Request���
'* @return Response���
'*
'******************************************************************************
Public Function OptionsReq(req As HttpRequest) As HttpResponse
    req.method = "OPTIONS"
    Set OptionsReq = Transmit(req)
End Function

'******************************************************************************
'* [�T  �v] Transmit ���\�b�h
'* [��  ��] �ȒP��Http�ʐM���s���B
'*
'* @param HttpRequest Request���
'* @return Response���
'*
'******************************************************************************
Public Function Transmit(req As HttpRequest, Optional saveOriginalBody As Boolean = False) As HttpResponse
    Call ReNewWinHttp
    mWinHttp.Open req.method, req.Url, False
    Call SetRequestHeaders(req)
    If req.method = "GET" Or req.method = "DELETE" Or req.method = "HEAD" Or req.method = "OPTIONS" Then
        mWinHttp.Send
    Else
        mWinHttp.Send req.Body
    End If
    Set Transmit = SetHttpResponse(, saveOriginalBody)
End Function

'******************************************************************************
'* [�T  �v] SetRequestHeaders ���\�b�h
'* [��  ��] HttpRequest��Header����ݒ肷��B
'*
'* @param Request���
'*
'******************************************************************************
Public Sub SetRequestHeaders(req As HttpRequest)
    Dim v
    For Each v In req.Headers.keys
        mWinHttp.SetRequestHeader CStr(v), req.Headers(v)
    Next
End Sub

'******************************************************************************
'* [�T  �v] SetHttpResponse ���\�b�h
'* [��  ��] Request���M���ʂ� HttpResponse �ɐݒ肷��B
'*
'* @param  encode �ȗ��\�BResponse�̕����R�[�h�B
'* @param  saveOriginalBody �ȗ��\�BBody��Text���������̂܂ܕێ����邩�w��B
'* @return Response���
'*
'******************************************************************************
Public Function SetHttpResponse(Optional encode As String = "utf-8", Optional saveOriginalBody As Boolean = False) As HttpResponse
    Dim Status As String:     Status = mWinHttp.Status
    Dim statusTxt As String:  statusTxt = mWinHttp.StatusText
    Dim resHeaders As String: resHeaders = mWinHttp.GetAllResponseHeaders
    
    Dim resBody As Variant
    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adReadAll = -1
    
    If saveOriginalBody Then
        resBody = mWinHttp.ResponseBody
    Else
        With CreateObject("ADODB.Stream")
            .Type = adTypeBinary
            .Open
            .Write mWinHttp.ResponseBody
            .Position = 0
            .Type = adTypeText
            .CharSet = encode
            resBody = .ReadText(adReadAll)
            .Close
        End With
    End If
  
    ' ���X�|���X�ݒ�
    Dim res As HttpResponse: Set res = New HttpResponse
    res.StatusCd = Status
    res.StatusText = statusTxt
    res.Body = resBody
    res.Headers = resHeaders
    
    Set SetHttpResponse = res
End Function

'******************************************************************************
'* [�T  �v] SendRequest ���\�b�h
'* [��  ��] Http���N�G�X�g�𑗐M����B
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
Public Function SendRequest(method As String, Url As String, reqHdrParams() As String, reqBody As String, _
                            Optional lTimeout As Long = 0, _
                            Optional resEncode As String = "utf-8") As HttpResponse
    Call ReNewWinHttp
    
    ' HTTP ���N�G�X�g
    mWinHttp.Open method, Url, False

    ' ���N�G�X�g�w�b�_�ݒ�
    Dim i As Long
    If ArrayUtils.GetDimension(reqHdrParams) <> 0 Then
        For i = 0 To UBound(reqHdrParams, 1)
            mWinHttp.SetRequestHeader reqHdrParams(i, 0), reqHdrParams(i, 1)
        Next i
    End If

    ' �^�C���A�E�g�ݒ�
    If lTimeout > 0 Then
        mWinHttp.SetTimeouts lTimeout, lTimeout, lTimeout, lTimeout
    End If
    
    ' ���N�G�X�g���M
    If method = "GET" Or method = "DELETE" Or method = "HEAD" Or method = "OPTIONS" Then
        mWinHttp.Send
    Else
        mWinHttp.Send reqBody
    End If

    Set SendRequest = SetHttpResponse(resEncode)
End Function

'******************************************************************************
'* [�T  �v] DownloadFile ���\�b�h
'* [��  ��] �t�@�C�����_�E�����[�h����B
'*
'* @param  req Request���
'* @param  saveFilePath �_�E�����[�h�t�@�C���ۑ���t���p�X
'* @return Response���
'*
'******************************************************************************
Public Function DownloadFile(req As HttpRequest, saveFilePath As String) As HttpResponse
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2

    Set DownloadFile = Transmit(req, True)
    
    If mWinHttp.Status >= 200 And mWinHttp.Status < 400 Then
        With CreateObject("ADODB.Stream")
            .Type = adTypeBinary
            .Open
            .Write mWinHttp.ResponseBody
            .SaveToFile saveFilePath, adSaveCreateOverWrite
            .Close
        End With
    End If
End Function
