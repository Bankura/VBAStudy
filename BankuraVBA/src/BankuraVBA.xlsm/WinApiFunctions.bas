Attribute VB_Name = "WinApiFunctions"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WindowsAPI�֘A�֐����W���[��
'* [��  ��] WindowsAPI�ɓn���R�[���o�b�N�֐�����WindowsAPI�֘A�̏������`����B
'*
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
'* �\���̒�`
'******************************************************************************
Private Type WbkDtl
    phwnd   As LongPtr
    hWnd    As LongPtr
    wkb     As Excel.Workbook
End Type

'******************************************************************************
'* �萔��`
'******************************************************************************
Private Const OBJID_NATIVEOM = &HFFFFFFF0
Private Const OBJID_CLIENT = &HFFFFFFFC
Private Const IID_IMdcList = "{8BD21D23-EC42-11CE-9E0D-00AA006002F3}"
Private Const IID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
Private Const IID_IDispatch = "{00020400-0000-0000-C000-000000000046}"

'******************************************************************************
'* �ϐ���`
'******************************************************************************
Private tmpProcessId As Long
Private tmpHwnd As LongPtr
Private tmpParentHwnd As LongPtr
Public wD() As WbkDtl

'******************************************************************************
' [�֐���] EnumerateWindow
' [���@��] �E�B���h�E��񋓂��邽�߂̃R�[���o�b�N�֐��B
' [���@��] hwnd �E�B���h�E�n���h��
'          lParam lParam
' [�߂�l] Boolean
'******************************************************************************
Public Function EnumerateWindow(ByVal hWnd As LongPtr, lParam As Long) As Boolean

    If GetWinAPI.IsWindowVisible(hWnd) Then
        Call PrintCaptionAndProcess(hWnd)
    End If
    EnumerateWindow = True
End Function

'******************************************************************************
' [�֐���] PrintCaptionAndProcess
' [���@��] �E�B���h�E�̃L���v�V�����ƃv���Z�X����\������B
' [���@��] hwnd �E�B���h�E�n���h��
'          lParam lParam
' [�߂�l] Boolean
'******************************************************************************
Private Sub PrintCaptionAndProcess(ByVal hWnd As LongPtr)
    Dim strClassBuff As String * 128
    Dim strTextBuff  As String * 516
    Dim strClass     As String
    Dim strText      As String
    Dim lngRtnCode   As Long
    Dim lngProcesID  As Long
    
    ' �N���X���擾
    lngRtnCode = GetWinAPI.GetClassName(hWnd, strClassBuff, Len(strClassBuff))
    strClass = Left(strClassBuff, InStr(strClassBuff, vbNullChar) - 1)
    
    ' �E�B���h�E�̃L���v�V�������擾
    lngRtnCode = GetWinAPI.GetWindowText(hWnd, strTextBuff, Len(strTextBuff))
    strText = Left(strTextBuff, InStr(strTextBuff, vbNullChar) - 1)
    
    ' �v���Z�XID���擾
    lngRtnCode = GetWinAPI.GetWindowThreadProcessId(hWnd, lngProcesID)

    Debug.Print "PID:" & lngProcesID, "hwnd:" & hWnd, "�N���X��:" & strClass, "Caption:" & strText
End Sub


'******************************************************************************
' [�֐���] GetEnumWindowProcessId
' [���@��] �E�B���h�E�̃v���Z�XID���擾���邽�߂̃R�[���o�b�N�֐��B
' [���@��] hwnd �E�B���h�E�n���h��
'          lParam lParam
' [�߂�l] Boolean
'******************************************************************************
Public Function GetEnumWindowProcessId(ByVal hWnd As LongPtr, lParam As Long) As Boolean
    Dim lngRtnCode  As Long
    Dim lngProcesID  As Long
    If GetWinAPI.IsWindowVisible(hWnd) Then
      '  �v���Z�XID���擾
        lngRtnCode = GetWinAPI.GetWindowThreadProcessId(hWnd, lngProcesID)
        ' �w�肵���v���Z�XID�ƈ�v����ꍇ�̓E�B���h�E�n���h����ݒ�
        If lngProcesID = tmpProcessId Then
            tmpHwnd = hWnd
        End If
    End If
    GetEnumWindowProcessId = True
End Function

'******************************************************************************
' [�֐���] PrintCaptionAndProcessMain
' [���@��] �E�B���h�E�̃L���v�V�����ƃv���Z�X����\������B
' [���@��] �Ȃ�
' [�߂�l] �Ȃ�
'******************************************************************************
Public Sub PrintCaptionAndProcessMain()
    Dim lngRtnCode  As Long

    lngRtnCode = GetWinAPI.EnumWindows(AddressOf EnumerateWindow, 0&)
End Sub


' �g�p�s�F���܂������Ȃ��B�����𖾂̂��ߎc���Ă����B
'******************************************************************************
' [�֐���] GetHwndByPid
' [���@��] �w�肵���v���Z�XID�ɑΉ������E�B���h�E�n���h�����擾����B
' [���@��] pId �v���Z�XID
' [�߂�l] LongPtr �E�B���h�E�n���h��
'******************************************************************************
Public Function GetHwndByPid(pid As Long) As LongPtr

    Dim hWnd As LongPtr
    Dim pIdLast As Long
    GetHwndByPid = 0
    
    '�f�X�N�g�b�vWindow�̎qWindow�̂����g�b�v���x����Window���擾
    hWnd = GetWinAPI.GetDesktopWindow()
    hWnd = GetWinAPI.GetWindow(hWnd, 5)

    Do While (0 <> hWnd)
        pIdLast = 0
        Call GetWinAPI.GetWindowThreadProcessId(hWnd, pIdLast)
'        Debug.Print "hwnd:" & hwnd & " pIdLast:" & pIdLast & "  pid:" & pid
        If (pid = pIdLast) Then
            GetHwndByPid = hWnd
            Exit Do
        End If
        '����Window�n���h�����擾
        hWnd = GetWinAPI.GetWindow(hWnd, 2)
    Loop
End Function

'******************************************************************************
' [�֐���] GetHwndByPid
' [���@��] �\������Ă���E�B���h�E�̃n���h�����v���Z�XID����擾����B
'          �iEnumWindows�𗘗p�j
' [���@��] pid
' [�߂�l] LongPtr
'******************************************************************************
Public Function GetHwndByPid2(pid As Long) As LongPtr
    Dim lngRtnCode  As Long
    tmpProcessId = 0
    tmpHwnd = 0
    
    tmpProcessId = pid
    lngRtnCode = GetWinAPI.EnumWindows(AddressOf GetEnumWindowProcessId, 0)
    GetHwndByPid2 = tmpHwnd
End Function


'******************************************************************************
'* [�T  �v] EnumChildSubProc
'* [��  ��] EnumChildWindows API�R�[���o�b�N�֐�
'*
'* @param hwndChild �q�E�B���h�E�n���h��
'* @param lParam LPARAM
'* @return Long ��������
'******************************************************************************
Public Function EnumChildSubProc(ByVal hWndChild As LongPtr, _
                                ByVal lParam As Long) As Long

    Dim strClassBuff As String * 128
    Dim strClass     As String
    Dim strTextBuff  As String * 516
    Dim strText      As String
    Dim lngRtnCode   As Long
    
    ' �N���X���擾
    lngRtnCode = GetWinAPI.GetClassName(hWndChild, strClassBuff, Len(strClassBuff))
    strClass = Left(strClassBuff, InStr(strClassBuff, vbNullChar) - 1)
    If strClass = "EXCEL7" Then
        ' �e�L�X�g���o�b�t�@�ɂ���
        lngRtnCode = GetWinAPI.GetWindowText(hWndChild, strTextBuff, Len(strTextBuff))
        strText = Left(strTextBuff, InStr(strTextBuff, vbNullChar) - 1)
        If InStr(1, strText, ".xla") = 0 Then
            If Sgn(wD) = 0 Then
                ReDim wD(0)
                wD(0).phwnd = tmpParentHwnd
                wD(0).hWnd = hWndChild
            Else
                ReDim Preserve wD(UBound(wD) + 1)
                wD(UBound(wD)).phwnd = tmpParentHwnd
                wD(UBound(wD)).hWnd = hWndChild
            End If
        End If

    End If
EnumChildPass:
    EnumChildSubProc = True
End Function

'******************************************************************************
'* [�T  �v] EnumWindowsProc
'* [��  ��] EnumWindows API�R�[���o�b�N�֐�
'*
'* @param hwnd �n���h��
'* @param lParam LPARAM
'* @return Long ��������
'******************************************************************************
Public Function EnumWindowsProc(ByVal hWnd As LongPtr, _
                                ByVal lParam As LongPtr) As Long

    Dim strClassBuff As String * 128
    Dim strClass     As String
    Dim lngRtnCode   As Long
    Dim lngThreadID  As Long
    Dim lngProcesID  As Long
    
    ' �N���X���擾
    lngRtnCode = GetWinAPI.GetClassName(hWnd, strClassBuff, Len(strClassBuff))
    strClass = Left(strClassBuff, InStr(strClassBuff, vbNullChar) - 1)
    
    If strClass = "XLMAIN" Then
        tmpParentHwnd = hWnd
        ' �q�E�B���h�E���
        lngRtnCode = GetWinAPI.EnumChildWindows(hWnd, _
                         AddressOf EnumChildSubProc, lParam)
                        
    End If
EnumPass:
    EnumWindowsProc = True
End Function

'******************************************************************************
'* [�T  �v] GetExcelBookProc
'* [��  ��] �ʃv���Z�XExcel�u�b�N�I�u�W�F�N�g�擾�֐����C��
'*
'* @return Boolean �������ʁiTrue:���� False�F�ُ�j
'******************************************************************************
Public Function GetExcelBookProc() As Boolean
    Dim lngRtnCode  As Long
    Dim i           As Long
    Dim iArr()      As Integer
    Dim wD2()       As WbkDtl
    Dim cnt         As Long

    Erase wD
    ' ���[�N�u�b�N�̃E�B���h�E�n���h�����擾
    lngRtnCode = GetWinAPI.EnumWindows(AddressOf EnumWindowsProc, 0)
    If Sgn(wD) <> 0 Then
        cnt = 0
        For i = 0 To UBound(wD)
            If GetExcelBook(wD(i)) Then
                ReDim Preserve iArr(cnt)
                iArr(cnt) = i
                cnt = cnt + 1
            End If
        Next
        
        ReDim wD2(0 To UBound(iArr)) As WbkDtl
        For i = 0 To UBound(iArr)
            wD2(i).hWnd = wD(iArr(i)).hWnd
            wD2(i).phwnd = wD(iArr(i)).phwnd
            Set wD2(i).wkb = wD(iArr(i)).wkb
        Next

        Erase wD
        ReDim wD(0 To UBound(wD2)) As WbkDtl
        For i = 0 To UBound(wD2)
            wD(i).hWnd = wD2(i).hWnd
            wD(i).phwnd = wD2(i).phwnd
            Set wD(i).wkb = wD2(i).wkb
        Next
        Erase wD2
        GetExcelBookProc = True
    Else
        GetExcelBookProc = False
    End If

End Function

'******************************************************************************
'* [�T  �v] GetExcelBook
'* [��  ��] �ʃv���Z�XExcel�u�b�N�I�u�W�F�N�g�擾�֐�
'*
'* @param wDl �n���h���{Excel�u�b�N�\����
'*
'******************************************************************************
Private Function GetExcelBook(wDl As WbkDtl) As Boolean
    Dim bytID()     As Byte
    Dim IID(0 To 3) As LongPtr
    Dim lngResult   As LongPtr
    Dim lngRtnCode  As LongPtr

    Dim wbw         As Excel.Window
    GetExcelBook = False
    
    If GetWinAPI.IsWindow(wDl.hWnd) = 0 Then Exit Function
    lngResult = GetWinAPI.SendMessage(wDl.hWnd, WM_GETOBJECT, 0, OBJID_NATIVEOM)
    If lngResult Then
        bytID = IID_IDispatch & vbNullChar

        GetWinAPI.IIDFromString bytID(0), IID(0)
        lngRtnCode = GetWinAPI.ObjectFromLresult(lngResult, IID(0), 0, wbw)
        If Not wbw Is Nothing Then Set wDl.wkb = wbw.Parent
        GetExcelBook = True
    End If
    
End Function
