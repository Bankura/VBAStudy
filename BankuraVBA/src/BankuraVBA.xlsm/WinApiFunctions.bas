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
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (lpsz As Any, lpiid As Any) As Long
Private Declare PtrSafe Function ObjectFromLresult Lib "oleacc" (ByVal lResult As LongPtr, riid As Any, ByVal wParam As LongPtr, ppvObject As Any) As LongPtr

'******************************************************************************
'* Enum��`
'******************************************************************************

'******************************************************************************
'* �\���̒�`
'******************************************************************************
Private Type WbkDtl
    phwnd   As LongPtr
    hwnd    As LongPtr
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
Public wd() As WbkDtl

'******************************************************************************
'* [�T  �v] PrintCaptionAndProcessMain
'* [��  ��] �E�B���h�E�̃L���v�V�����ƃv���Z�X����\������B
'*
'******************************************************************************
Public Sub PrintCaptionAndProcessMain()
    Dim lngRtnCode  As Long

    lngRtnCode = EnumWindows(AddressOf EnumerateWindow, 0&)
End Sub

'******************************************************************************
'* [�T  �v] EnumerateWindow
'* [��  ��] �E�B���h�E��񋓂��邽�߂̃R�[���o�b�N�֐��B
'*
'* @param hwnd �E�B���h�E�n���h��
'* @param lParam lParam
'* @return Boolean
'******************************************************************************
Public Function EnumerateWindow(ByVal hwnd As LongPtr, lParam As Long) As Boolean

    If IsWindowVisible(hwnd) Then
        Call PrintCaptionAndProcess(hwnd)
    End If
    EnumerateWindow = True
End Function

'******************************************************************************
'* [�T  �v] PrintCaptionAndProcess
'* [��  ��] �E�B���h�E�̃L���v�V�����ƃv���Z�X����\������B
'*
'* @param hwnd �E�B���h�E�n���h��
'******************************************************************************
Private Sub PrintCaptionAndProcess(ByVal hwnd As LongPtr)
    Dim strClassBuff As String * 128
    Dim strTextBuff  As String * 516
    Dim strClass     As String
    Dim strText      As String
    Dim lngRtnCode   As Long
    Dim lngProcesID  As Long
    
    ' �N���X���擾
    lngRtnCode = GetClassName(hwnd, strClassBuff, Len(strClassBuff))
    strClass = Left(strClassBuff, InStr(strClassBuff, vbNullChar) - 1)
    
    ' �E�B���h�E�̃L���v�V�������擾
    lngRtnCode = GetWindowText(hwnd, strTextBuff, Len(strTextBuff))
    strText = Left(strTextBuff, InStr(strTextBuff, vbNullChar) - 1)
    
    ' �v���Z�XID���擾
    lngRtnCode = GetWindowThreadProcessId(hwnd, lngProcesID)

    Debug.Print "PID:" & lngProcesID, "hwnd:" & hwnd, "�N���X��:" & strClass, "Caption:" & strText
End Sub

'******************************************************************************
'* [�T  �v] GetExcelBookProc
'* [��  ��] �ʃv���Z�XExcel�u�b�N�I�u�W�F�N�g�擾�֐����C���B
'*          �擾����Excel�v���Z�X���́A�O���[�o���ϐ��� wd �\���̂Ɋi�[�����B
'*
'* @return Boolean �������ʁiTrue:���� False�F�ُ�j
'******************************************************************************
Public Function GetExcelBookProc() As Boolean
    Dim lngRtnCode  As Long
    Dim i           As Long
    Dim iArr()      As Integer
    Dim wD2()       As WbkDtl
    Dim cnt         As Long

    Erase wd
    ' ���[�N�u�b�N�̃E�B���h�E�n���h�����擾
    lngRtnCode = EnumWindows(AddressOf EnumWindowsProc, 0)
    
    On Error Resume Next
    Dim cktemp As Long
    cktemp = UBound(wd)
    
    If Err.Number = 0 Then
        On Error GoTo 0
        cnt = 0
        For i = 0 To UBound(wd)
            If GetExcelBook(wd(i)) Then
                ReDim Preserve iArr(cnt)
                iArr(cnt) = i
                cnt = cnt + 1
            End If
        Next
        
        ReDim wD2(0 To UBound(iArr)) As WbkDtl
        For i = 0 To UBound(iArr)
            wD2(i).hwnd = wd(iArr(i)).hwnd
            wD2(i).phwnd = wd(iArr(i)).phwnd
            Set wD2(i).wkb = wd(iArr(i)).wkb
        Next

        Erase wd
        ReDim wd(0 To UBound(wD2)) As WbkDtl
        For i = 0 To UBound(wD2)
            wd(i).hwnd = wD2(i).hwnd
            wd(i).phwnd = wD2(i).phwnd
            Set wd(i).wkb = wD2(i).wkb
        Next
        Erase wD2
        GetExcelBookProc = True
    Else
        Err.Clear
        On Error GoTo 0
        GetExcelBookProc = False
    End If
End Function

'******************************************************************************
'* [�T  �v] EnumWindowsProc
'* [��  ��] EnumWindows API�R�[���o�b�N�֐�
'*
'* @param hwnd �n���h��
'* @param lParam LPARAM
'* @return Long ��������
'******************************************************************************
Public Function EnumWindowsProc(ByVal hwnd As LongPtr, _
                                ByVal lParam As LongPtr) As Long

    Dim strClassBuff As String * 128
    Dim strClass     As String
    Dim lngRtnCode   As Long
    Dim lngThreadID  As Long
    Dim lngProcesID  As Long
    
    ' �N���X���擾
    lngRtnCode = GetClassName(hwnd, strClassBuff, Len(strClassBuff))
    strClass = Left(strClassBuff, InStr(strClassBuff, vbNullChar) - 1)
    
    If strClass = "XLMAIN" Then
        tmpParentHwnd = hwnd
        ' �q�E�B���h�E���
        lngRtnCode = EnumChildWindows(hwnd, _
                         AddressOf EnumChildSubProc, lParam)
                        
    End If
EnumPass:
    EnumWindowsProc = True
End Function

'******************************************************************************
'* [�T  �v] EnumChildSubProc
'* [��  ��] EnumChildWindows API�R�[���o�b�N�֐�
'*
'* @param hwndChild �q�E�B���h�E�n���h��
'* @param lParam LPARAM
'* @return Long ��������
'******************************************************************************
Public Function EnumChildSubProc(ByVal hwndChild As LongPtr, _
                                ByVal lParam As Long) As Long

    Dim strClassBuff As String * 128
    Dim strClass     As String
    Dim strTextBuff  As String * 516
    Dim strText      As String
    Dim lngRtnCode   As Long
    
    ' �N���X���擾
    lngRtnCode = GetClassName(hwndChild, strClassBuff, Len(strClassBuff))
    strClass = Left(strClassBuff, InStr(strClassBuff, vbNullChar) - 1)
    If strClass = "EXCEL7" Then
        ' �e�L�X�g���o�b�t�@�ɂ���
        lngRtnCode = GetWindowText(hwndChild, strTextBuff, Len(strTextBuff))
        strText = Left(strTextBuff, InStr(strTextBuff, vbNullChar) - 1)
        If InStr(1, strText, ".xla") = 0 Then
            On Error Resume Next
            Dim cktemp As LongPtr
            cktemp = UBound(wd)
        
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                ReDim wd(0)
                wd(0).phwnd = tmpParentHwnd
                wd(0).hwnd = hwndChild
            Else
                On Error GoTo 0
                ReDim Preserve wd(UBound(wd) + 1)
                wd(UBound(wd)).phwnd = tmpParentHwnd
                wd(UBound(wd)).hwnd = hwndChild
            End If
        End If

    End If
EnumChildPass:
    EnumChildSubProc = True
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
    
    If IsWindow(wDl.hwnd) = 0 Then Exit Function
    lngResult = SendMessage(wDl.hwnd, WM_GETOBJECT, 0, ByVal OBJID_NATIVEOM)
    If lngResult Then
        bytID = IID_IDispatch & vbNullChar

        IIDFromString bytID(0), IID(0)
        lngRtnCode = ObjectFromLresult(lngResult, IID(0), 0, wbw)
        If Not wbw Is Nothing Then Set wDl.wkb = wbw.Parent
        GetExcelBook = True
    End If
    
End Function

'******************************************************************************
'* [�T  �v] GetEnumWindowProcessId
'* [��  ��] �E�B���h�E�̃v���Z�XID���擾���邽�߂̃R�[���o�b�N�֐��B
'*
'* @param hwnd �E�B���h�E�n���h��
'* @param lParam lParam
'* @return Boolean
'******************************************************************************
Public Function GetEnumWindowProcessId(ByVal hwnd As LongPtr, lParam As Long) As Boolean
    Dim lngRtnCode  As Long
    Dim lngProcesID  As Long
    If IsWindowVisible(hwnd) Then
      '  �v���Z�XID���擾
        lngRtnCode = GetWindowThreadProcessId(hwnd, lngProcesID)
        ' �w�肵���v���Z�XID�ƈ�v����ꍇ�̓E�B���h�E�n���h����ݒ�
        If lngProcesID = tmpProcessId Then
            tmpHwnd = hwnd
        End If
    End If
    GetEnumWindowProcessId = True
End Function

'******************************************************************************
'* [�T  �v] GetHwndByPid
'* [��  ��] �w�肵���v���Z�XID�ɑΉ������E�B���h�E�n���h�����擾����B
'*          �g�p�s�F���܂������Ȃ��B�����𖾂̂��ߎc���Ă����B
'*
'* @param pId �v���Z�XID
'* @return LongPtr �E�B���h�E�n���h��
'******************************************************************************
Public Function GetHwndByPid(pid As Long) As LongPtr

    Dim hwnd As LongPtr
    Dim pIdLast As Long
    GetHwndByPid = 0
    
    '�f�X�N�g�b�vWindow�̎qWindow�̂����g�b�v���x����Window���擾
    hwnd = GetDesktopWindow()
    hwnd = GetWindow(hwnd, 5)

    Do While (0 <> hwnd)
        pIdLast = 0
        Call GetWindowThreadProcessId(hwnd, pIdLast)
'        Debug.Print "hwnd:" & hwnd & " pIdLast:" & pIdLast & "  pid:" & pid
        If (pid = pIdLast) Then
            GetHwndByPid = hwnd
            Exit Do
        End If
        '����Window�n���h�����擾
        hwnd = GetWindow(hwnd, 2)
    Loop
End Function

'******************************************************************************
'* [�T  �v] GetHwndByPid
'* [��  ��] �\������Ă���E�B���h�E�̃n���h�����v���Z�XID����擾����B
'*         �iEnumWindows�𗘗p�j
'*
'* @param pid �v���Z�XID
'* @return LongPtr �E�B���h�E�n���h��
'******************************************************************************
Public Function GetHwndByPid2(pid As Long) As LongPtr
    Dim lngRtnCode  As Long
    tmpProcessId = 0
    tmpHwnd = 0
    
    tmpProcessId = pid
    lngRtnCode = EnumWindows(AddressOf GetEnumWindowProcessId, 0)
    GetHwndByPid2 = tmpHwnd
End Function
