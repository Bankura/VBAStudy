Attribute VB_Name = "Base"
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] BankuraVBA���ʊ�Ճ��W���[��
'* [��  ��] ���ʂŎg�p����v���V�[�W����񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019-2021 Bankura
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
Public Type Array2DIndex
    x As Long
    y As Long
End Type

'******************************************************************************
'* �萔��`
'******************************************************************************
Public Const MAX_INT As Integer = 32767
Public Const MAX_LONG As Long = 2147483647
Public Const TWIP As Long = 567

#If Win64 Then
    Public Const LONGPTR_SIZE = 8
#Else
    Public Const LONGPTR_SIZE = 4
#End If

' For Wsh
Public Const WshHide = 0
Public Const ForReading = 1
    
'******************************************************************************
'* �ϐ���`
'******************************************************************************
Private mDisplayAlerts As Boolean
Private mScreenUpdating As Boolean
Private mCalculation As Long
Private mEnableEvents As Boolean
Private mRegExp As Object
Private mShell As Object
Private mWshNetwork As Object
Private mSc As Object
Private mSettingInfo As SettingInfo
Private mSettingSheetName As String
Private mSettingSheetStartRow As Long
Private mSettingSheetStartCol As Long
Private mLogger As Logger
Private mCallbackObjCol As Collection
Private mCallbackParamCol As Collection
Private mCallbackResultCol As Collection

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************
'*-----------------------------------------------------------------------------
'* Logger �v���p�e�B
'*-----------------------------------------------------------------------------
Public Property Get Logger(Optional logFName As String = "bankuravba.log") As Logger
    If mLogger Is Nothing Then
        Set mLogger = New Logger
        Call mLogger.Init(LogLevelEnum.lvTrace, True, IO.ExecPath, logFName)
    End If
    Set Logger = mLogger
End Property

'*-----------------------------------------------------------------------------
'* SettingInfo �v���p�e�B
'*-----------------------------------------------------------------------------
Public Property Get SettingInfo() As SettingInfo
    Set SettingInfo = mSettingInfo
End Property
Public Property Set SettingInfo(ByVal arg As SettingInfo)
    Set mSettingInfo = arg
End Property

'*-----------------------------------------------------------------------------
'* SettingSheetName �v���p�e�B
'*-----------------------------------------------------------------------------
Public Property Get SettingSheetName() As String
    SettingSheetName = mSettingSheetName
End Property
Public Property Let SettingSheetName(ByVal arg As String)
    mSettingSheetName = arg
End Property

'*-----------------------------------------------------------------------------
'* SettingSheetStartRow �v���p�e�B
'*-----------------------------------------------------------------------------
Public Property Get SettingSheetStartRow() As Long
    SettingSheetStartRow = mSettingSheetStartRow
End Property
Public Property Let SettingSheetStartRow(ByVal arg As Long)
    mSettingSheetStartRow = arg
End Property

'*-----------------------------------------------------------------------------
'* SettingSheetStartCol �v���p�e�B
'*-----------------------------------------------------------------------------
Public Property Get SettingSheetStartCol() As Long
    SettingSheetStartCol = mSettingSheetStartCol
End Property
Public Property Let SettingSheetStartCol(ByVal arg As Long)
    mSettingSheetStartCol = arg
End Property

'*-----------------------------------------------------------------------------
'* ActiveSheetEx �v���p�e�B
'*-----------------------------------------------------------------------------
Public Property Get ActiveSheetEx() As WorkSheetEx
    Set ActiveSheetEx = Core.Init(New WorkSheetEx, Application.ActiveSheet.Name)
End Property

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] �G���[�����B
'* [��  ��] �G���[�������̏������s���B
'*
'******************************************************************************
Public Sub ErrorProcess()
    Debug.Print "�G���[���� Number: " & Err.Number & " Source: " & Err.Source & " Description: " & Err.Description
End Sub

'******************************************************************************
'* [�T  �v] �J�n�����B
'* [��  ��] �����̃X�s�[�h����̂��߁AExcel�̐ݒ��ύX����B
'*
'******************************************************************************
Public Sub StartProcess()
    Call SaveApplicationProperties
    With Application
        .Cursor = xlWait
        .DisplayAlerts = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
End Sub

'******************************************************************************
'* [�T  �v] �I�������B
'* [��  ��] �����̃X�s�[�h����̂��ߕύX����Excel�̐ݒ�����ɖ߂��B
'*
'******************************************************************************
Public Sub EndProcess()
    With Application
        .Cursor = xlDefault
        .DisplayAlerts = mDisplayAlerts
        .ScreenUpdating = mScreenUpdating
        .Calculation = mCalculation
        .EnableEvents = mEnableEvents
        .StatusBar = False
    End With
End Sub

'******************************************************************************
'* [�T  �v] Application�ݒ�ޔ������B
'* [��  ��] Application�̐ݒ�������o�ϐ��ɑޔ�����B
'*
'******************************************************************************
Public Sub SaveApplicationProperties()
    With Application
        mDisplayAlerts = .DisplayAlerts
        mScreenUpdating = .ScreenUpdating
        mCalculation = .Calculation
        mEnableEvents = .EnableEvents
    End With
End Sub

'******************************************************************************
'* [�T  �v] ���K�\���I�u�W�F�N�g�擾�����B
'* [��  ��] ���K�\���I�u�W�F�N�g���擾����B�������̏ꍇ��������B
'*
'******************************************************************************
Public Function GetRegExp() As Object
    If mRegExp Is Nothing Then
        Set mRegExp = CreateObject("VBScript.RegExp")
    End If
    Set GetRegExp = mRegExp
End Function

'******************************************************************************
'* [�T  �v] Shell�I�u�W�F�N�g�擾�����B
'* [��  ��] Shell�I�u�W�F�N�g���擾����B�������̏ꍇ��������B
'*
'******************************************************************************
Public Function GetShell() As Object
    If mShell Is Nothing Then
        Set mShell = CreateObject("Shell.Application")
    End If
    Set GetShell = mShell
End Function

'******************************************************************************
'* [�T  �v] WScript.Network�I�u�W�F�N�g�擾�����B
'* [��  ��] WScript.Network�I�u�W�F�N�g���擾����B�������̏ꍇ��������B
'*
'******************************************************************************
Public Function GetWshNetwork() As Object
    If mWshNetwork Is Nothing Then
        Set mWshNetwork = CreateObject("WScript.Network")
    End If
    Set GetWshNetwork = mWshNetwork
End Function

'******************************************************************************
'* [�T  �v] ScriptControl�I�u�W�F�N�g�擾�����B
'* [��  ��] ScriptControl�I�u�W�F�N�g���擾����B�������̏ꍇ��������B
'*
'******************************************************************************
Public Function GetScriptControl() As Object
    If mSc Is Nothing Then
        Set mSc = CreateObject32bit("MSScriptControl.ScriptControl")
    End If
    Set GetScriptControl = mSc
End Function

'******************************************************************************
'* [�T  �v] CDO.Message�I�u�W�F�N�g���������B
'* [��  ��] CDO.Message�I�u�W�F�N�g�𐶐�����B
'*
'******************************************************************************
Public Function CreateCDOMessage() As Object
    Set CreateCDOMessage = CreateObject("CDO.Message")
End Function

'******************************************************************************
'* [�T  �v] CreateWmiServices ���\�b�h
'* [��  ��] GetObject ���\�b�h���g�p����SWbemServices�I�u�W�F�N�g�𐶐�����B
'*
'* [��  ��] WMI�̖��O��Ԃ́A�ȉ��̂悤�Ɂu�R���s���[�^�̊Ǘ��v����m�F�\�B
'*          �E�R���g���[���p�l�����Ǘ��c�[�����R���s���[�^�̊Ǘ� ��I��
'*            �ȉ��̃R�}���h�ł��N���\
'*            %windir%\system32\compmgmt.msc /s
'*          �E�T�[�r�X�ƃA�v���P�[�V������WMI�R���g���[����I�����A
'*            �E�N���b�N���v���p�e�B��I��
'*          �E�\�����ꂽ�uWMI�R���g���[���̃v���p�e�B�v�̃Z�L�����e�B�^�u��I��
'*
'* [�Q  �l] http://dodonpa.la.coocan.jp/windows_service_wmi_1.htm
'*
'* @param strComputer �ȗ��B�R���s���[�^���B
'* @param ns          �ȗ��B���O��ԁB
'* @return SWbemServices�I�u�W�F�N�g�B
'*
'******************************************************************************
Function CreateSWbemServices(Optional strComputer As String = ".", Optional ns As String = "\root\cimv2", Optional userId As String, Optional passwd As String) As Object
    If userId = "" Then
        Set CreateSWbemServices = GetObject("winmgmts:\\" & strComputer & ns)
    Else
        Set CreateSWbemServices = Core.wmi.ConnectServer(strComputer, ns, userId, passwd)
    End If
End Function

'******************************************************************************
'* [�T  �v] �ݒ���I�u�W�F�N�g�擾�����B
'* [��  ��] �ݒ���I�u�W�F�N�g���擾����B�������̏ꍇ��������B
'*
'******************************************************************************
Public Function GetSettingInfo() As SettingInfo
    If mSettingInfo Is Nothing Then
        Set mSettingInfo = Core.Init(New SettingInfo, mSettingSheetName, mSettingSheetStartRow, mSettingSheetStartCol)
    End If
    Set GetSettingInfo = mSettingInfo
End Function

Public Function GetMasterValueByCode(masterName As String, code As String) As String
    If mSettingInfo Is Nothing Then
        Set mSettingInfo = GetSettingInfo()
    End If
    GetMasterValueByCode = mSettingInfo.GetMasterValueByCode(masterName, code)
End Function
Public Function GetMasterCodeByValue(masterName As String, val As String) As String
    If mSettingInfo Is Nothing Then
        Set mSettingInfo = GetSettingInfo()
    End If
    GetMasterCodeByValue = mSettingInfo.GetMasterCodeByValue(masterName, val)
End Function

'*******************************************************************************
'* [�T  �v] �R���s���[�^���ݒ菈��
'* [��  ��] ���s�[���̃R���s���[�^�����擾�B
'*
'* @param String �R���s���[�^��
'*
'*******************************************************************************
Public Function GetComputerName() As String
    GetComputerName = Core.Wsh.ComputerName
End Function

'******************************************************************************
'* [�T  �v] ���s�A�v���P�[�V�������菈��
'* [��  ��] ���s�A�v���P�[�V������Excel�����肷��B
'*
'* @param Boolean �������ʁiTrue:Excel False�FExcel�ȊO�j
'*
'******************************************************************************
Public Function CheckXlApplication() As Boolean
    CheckXlApplication = InStr(Application.Name, "Excel") > 0
End Function

'******************************************************************************
'* [�T  �v] Is32BitProcessorForApp
'* [��  ��] �g�p����A�v���P�[�V������32�r�b�g�����`�F�b�N����B
'*
'* @return �`�F�b�N���ʁiTrue: 32Bit�AFalse: 64bit�j
'*
'******************************************************************************
Public Function Is32BitProcessorForApp() As Boolean
    Dim proc As String: proc = Wsh.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    If proc = "x86" Then
       Is32BitProcessorForApp = True
    Else
       Is32BitProcessorForApp = False
    End If
End Function

'******************************************************************************
'* [�T  �v] Is32BitProcessor
'* [��  ��] �g�p����[���̃v���Z�b�T��32�r�b�g�����`�F�b�N����B
'*
'* @return �`�F�b�N���ʁiTrue: 32Bit�AFalse: 64bit�j
'*
'******************************************************************************
Public Function Is32BitProcessor() As Boolean
    Dim proc As String: proc = Wsh.ExpandEnvironmentStrings("%PROCESSOR_ARCHITEW6432%")
    If proc = "x86" Then
       Is32BitProcessor = True
    Else
       Is32BitProcessor = False
    End If
End Function

'******************************************************************************
'* [�T  �v] CreateObject32bit
'* [��  ��] 32�r�b�g����Object�𐶐�����B
'* [�Q  �l] <https://github.com/vocho/vbs/blob/a5c3ee608103638678c983da00ec290c4b8ab90c/CreateObject32bit.vbs>
'*
'* @param strClassName �����Ώۂ̃N���X���B"Shell.Application"���B
'* @return 32�r�b�g��Object
'*
'******************************************************************************
Public Function CreateObject32bit(ByVal strClassName As String) As Variant
    If Is32BitProcessorForApp Then
     Set CreateObject32bit = CreateObject(strClassName)
     Exit Function
    End If
    
    Base.GetShell.Windows().Item(0).PutProperty strClassName, Nothing
    
    ' �ꎞ�X�N���v�g�R�}���h�e�L�X�g����
    Dim strScriptCodes As String
    strScriptCodes = "CreateObject(""Shell.Application"").Windows().Item(0).PutProperty """ & strClassName & """, CreateObject(""" & strClassName & """)" & vbNewLine & _
                     "Set objExec = CreateObject(""WScript.Shell"").Exec(""MSHTA.EXE -"")" & vbNewLine & _
                     "Set objWMIService = GetObject(""winmgmts:"")" & vbNewLine & _
                     "lngCurrentPID = objWMIService.Get(""Win32_Process.Handle="" & objExec.ProcessID).ParentProcessID" & vbNewLine & _
                     "objExec.Terminate" & vbNewLine & _
                     "lngParentPID = objWMIService.Get(""Win32_Process.Handle="" & lngCurrentPID).ParentProcessID" & vbNewLine & _
                     "Do While objWMIService.ExecQuery(""SELECT * FROM Win32_Process WHERE ProcessID="" & lngParentPID).Count<>0" & vbNewLine & _
                     "    WScript.Sleep 1000" & vbNewLine & _
                     "Loop" & vbNewLine & _
                     "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbNewLine & _
                     "If objFSO.FileExists(WScript.ScriptFullName) Then objFSO.DeleteFile WScript.ScriptFullName" & vbNewLine & _
                     ""

    ' �ꎞ�X�N���v�g�t�@�C���쐬
    Dim strTempFile As String: strTempFile = FileUtils.GetTempFilePath(, ".vbs")
    Call FileUtils.WriteUTF8TextFile(strTempFile, strScriptCodes)
    
    ' �ꎞ�X�N���v�g�t�@�C�����s(32bit)
    With Core.Wsh.Environment("Process")
        .Item("SysWOW64") = IO.fso.BuildPath(.Item("SystemRoot"), "SysWOW64")
        .Item("WScriptName") = IO.fso.GetFileName("C:\WINDOWS\SysWOW64\cscript.exe")
        .Item("WScriptWOW64") = IO.fso.BuildPath(.Item("SysWOW64"), .Item("WScriptName"))
        .Item("Run") = .Item("WScriptWOW64") & " """ & strTempFile & """"
         Core.Wsh.Run .Item("Run"), True
    End With
    
    ' �I�u�W�F�N�g�󂯎��
    Do
        Set CreateObject32bit = Base.GetShell.Windows().Item(0).GetProperty(strClassName)
    Loop While CreateObject32bit Is Nothing
End Function

'******************************************************************************
'* [�T  �v] SetAppSettingsNormal
'* [��  ��] �A�v���P�[�V�����̐ݒ��ʏ�̐ݒ�ɂ���B
'*
'******************************************************************************
Public Sub SetAppSettingsNormal()
    With Application
        .Cursor = xlDefault
        .DisplayAlerts = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub

'******************************************************************************
'* [�T  �v] JudgeCond
'* [��  ��] 2�l���w�肵����r���Z�q�Ŕ�r�E���肷��B
'*
'* @param val1 ��r����l1
'* @param val2 ��r����l2
'* @param cond ��r���Z�q�i"<", ">", "<=", ">=", "="�j
'* @param flg ���茋�ʂ̔��]�p�t���O
'*            "<", ">"�̏ꍇ�ɔ��]��̏�����"="���܂܂Ȃ��B
'*            �N�C�b�N�\�[�g�̔���Ŏg�p�B
'*
'* @return Boolean ���茋��
'*
'******************************************************************************
Public Function JudgeCond(val1, val2, cond As String, Optional flg As Boolean = True) As Boolean
    If flg Then
        Select Case cond
            Case ">"
                JudgeCond = val1 > val2
            Case "<"
                JudgeCond = val1 < val2
            Case ">="
                JudgeCond = val1 >= val2
            Case "<="
                JudgeCond = val1 <= val2
            Case "="
                JudgeCond = val1 = val2
            Case Else
                Err.Raise 9999, "JudgeCond", "Bad Condition."
        End Select
    Else
        Select Case cond
            Case ">"
                JudgeCond = val1 < val2
            Case "<"
                JudgeCond = val1 > val2
            Case ">="
                JudgeCond = val1 <= val2
            Case "<="
                JudgeCond = val1 >= val2
            Case "="
                JudgeCond = val1 = val2
            Case Else
                Err.Raise 9999, "JudgeCond", "Bad Condition."
        End Select
    End If
End Function

'******************************************************************************
'* [�T  �v] Swap
'* [��  ��] �l�����ւ���B
'*
'* @param val1 �l1
'* @param val2 �l2
'******************************************************************************
Public Sub Swap(val1, val2)
    Dim tmp: tmp = val1
    val1 = val2
    val2 = tmp
End Sub

'******************************************************************************
'* [�T  �v] CreateUUID
'* [��  ��] UUID(GUID)�𐶐�����B
'* [�Q  �l] https://stackoverflow.com/a/46474125/918626
'*
'* @return String UUID
'*
'******************************************************************************
Public Function CreateUUID() As String
    Dim myUuid As String
    Randomize Timer() + Application.hWnd
    Do While Len(myUuid) < 32
        If Len(myUuid) = 16 Then
            myUuid = myUuid & Hex$(8 + CInt(Rnd * 3))
        End If
        myUuid = myUuid & Hex$(CInt(Rnd * 15))
    Loop
    CreateUUID = Mid(myUuid, 1, 8) & "-" & Mid(myUuid, 9, 4) & "-" & Mid(myUuid, 13, 4) & "-" & Mid(myUuid, 17, 4) & "-" & Mid(myUuid, 21, 12)
End Function

'******************************************************************************
'* [�T  �v] AppendEnvItem
'* [��  ��] ���ϐ���ǉ�����B
'*
'* @param itemName  ���ږ�
'* @param itemValue �ݒ�l
'* @param envType   ���ϐ��̎�ށi�f�t�H���g��"Process"�j
'*                    "System"  : �V�X�e�����ϐ��B�S���[�U�[�ɓK�p�����B
'*                    "User"    : ���[�U�[���ϐ��B���[�U�[�ɓK�p�����B
'*                    "Volatile": ���������ϐ��B���O�I�t���ɔj�������B
'*                    "Process" : �v���Z�X���ϐ��B�v���Z�X�I�����ɔj���B
'* @param appendHead �擪�ɉ����邩�ǂ����B
'*
'******************************************************************************
Public Sub AppendEnvItem(itemName As String, itemValue, Optional envType As String = "Process", Optional appendHead As Boolean = True)
    With Core.Wsh
        Dim destEnvValue: destEnvValue = .Environment(envType).Item(itemName)
        Dim sep As String: sep = IIf(destEnvValue <> "", ";", "")
        
        If Not StringUtils.Contains(destEnvValue, itemValue) Then
            If appendHead Then
                .Environment(envType).Item(itemName) = itemValue & sep & destEnvValue
            Else
                .Environment(envType).Item(itemName) = destEnvValue & sep & itemValue
            End If
        End If
    End With
End Sub

'******************************************************************************
'* [�T  �v] EditEnvItem
'* [��  ��] ���ϐ���ҏW����B
'*
'* @param itemName  ���ږ�
'* @param itemValue �ݒ�l
'* @param envType   ���ϐ��̎�ށi�f�t�H���g��"Process"�j
'*                    "System"  : �V�X�e�����ϐ��B�S���[�U�[�ɓK�p�����B
'*                    "User"    : ���[�U�[���ϐ��B���[�U�[�ɓK�p�����B
'*                    "Volatile": ���������ϐ��B���O�I�t���ɔj�������B
'*                    "Process" : �v���Z�X���ϐ��B�v���Z�X�I�����ɔj���B
'*
'******************************************************************************
Public Sub EditEnvItem(itemName As String, itemValue, Optional envType As String = "Process")
    With Core.Wsh
        .Environment(envType).Item(itemName) = itemValue
    End With
End Sub

'******************************************************************************
'* [�T  �v] ForEach
'* [��  ��] �uFor Each�v�ŌJ��Ԃ��\�ȃI�u�W�F�N�g�ɑ΂��āA������K�p����B
'*
'* @param obj  �uFor Each�v�ŌJ��Ԃ��\�ȃI�u�W�F�N�g
'* @param proc �֐����A�܂���Func�I�u�W�F�N�g�A�܂���
'*                   Exec�ix As Object�j���\�b�h��
'*                   ���I�u�W�F�N�g�B
'*
'******************************************************************************
Public Sub ForEach(ByVal obj As Object, ByVal proc As Variant)
    Dim o
    For Each o In obj
        If ValidateUtils.IsFunc(proc) Then
            Call proc.Apply(o)
        ElseIf ValidateUtils.IsString(proc) And proc <> "" Then
            Call Application.Run(proc, o)
        ElseIf IsObject(proc) Then
            Call proc.Exec(o)
        End If
    Next
End Sub

'******************************************************************************
'* [�T  �v] GetRandom
'* [��  ��] �����_���l�𐶐�����B
'*
'* @param willRandomize Randomize�����s���邩�ǂ���
'* @return Single �����_���l
'*
'******************************************************************************
Public Function GetRandom(Optional willRandomize As Boolean = False) As Single
    If willRandomize Then
        Randomize
    End If
    GetRandom = Rnd
End Function

'******************************************************************************
'* [�T  �v] GetRandomInt
'* [��  ��] �����_���Ȑ����𐶐�����B
'*
'* @param minVal �ŏ��l
'* @param maxVal �ő�l
'* @param willRandomize Randomize�����s���邩�ǂ���
'* @return Long �����_���Ȑ���
'*
'******************************************************************************
Public Function GetRandomInt(Optional minVal As Long = 0, Optional maxVal As Long = MAX_LONG - 1, Optional willRandomize As Boolean = False) As Long
    Dim randomVal As Single: randomVal = GetRandom(willRandomize)
    If minVal > maxVal Then
        Swap minVal, maxVal
    End If
    GetRandomInt = Int((maxVal * randomVal) + (0 - (minVal * randomVal) + minVal) + randomVal)
End Function

'******************************************************************************
'* [�T  �v] GetRandomDate
'* [��  ��] �����_���ȓ��t�𐶐�����B
'*
'* @param minVal �ŏ��l
'* @param maxVal �ő�l
'* @param noTimeVal �������܂܂Ȃ����iTrue�F�܂܂Ȃ��j
'* @param willRandomize Randomize�����s���邩�ǂ���
'* @return Date �����_���ȓ��t�i�����b�܂ށj
'*
'******************************************************************************
Public Function GetRandomDate(Optional minVal As Single = 0, Optional maxVal As Single = 2958465, Optional noTimeVal As Boolean = False, Optional willRandomize As Boolean = False) As Date
    Dim randomVal As Single: randomVal = GetRandom(willRandomize)
    If minVal = 0 Then
        minVal = Now
    End If
    If minVal > maxVal Then
        Swap minVal, maxVal
    End If
    If noTimeVal Then
        GetRandomDate = CDate(Int((maxVal * randomVal) + (0 - (minVal * randomVal) + minVal) + randomVal))
    Else
        GetRandomDate = CDate((maxVal * randomVal) + (0 - (minVal * randomVal) + minVal) + randomVal)
    End If
End Function


'******************************************************************************
'* [�T  �v] GetArrayDataRandom
'* [��  ��] �z��̃f�[�^�������_���Ɏ擾����B
'*
'* @param arr �Ώۂ̔z��
'* @param willRandomize Randomize�����s���邩�ǂ���
'* @return Variant �����_���ȗv�f
'*
'******************************************************************************
Public Function GetArrayDataRandom(Arr As Variant, Optional willRandomize As Boolean = False) As Variant
    GetArrayDataRandom = Arr(Int((UBound(Arr) - LBound(Arr) + 1) * GetRandom(willRandomize) + LBound(Arr)))
End Function

'******************************************************************************
'* [�T  �v] GetRandomIntArray
'* [��  ��] �����_���Ȑ������i�[�����z��̃f�[�^�𐶐�����B
'*
'* @param numOfItems �z��̌�
'* @param minVal �ŏ��l
'* @param maxVal �ő�l
'* @param noOverLap �d���l�����e���Ȃ����iTrue:�d���Ȃ��j
'* @param willRandomize Randomize�����s���邩�ǂ���
'* @return Variant �z��f�[�^
'*
'******************************************************************************
Public Function GetRandomIntArray(numOfItems As Long, _
                                  Optional minVal As Long = 0, _
                                  Optional maxVal As Long = MAX_LONG - 1, _
                                  Optional noOverLap As Boolean = True, _
                                  Optional willRandomize As Boolean = False) As Variant
    If willRandomize Then
        Randomize
    End If
    
    If minVal > maxVal Then
        Swap minVal, maxVal
    End If
    
    If numOfItems > (maxVal - minVal + 1) Then
        noOverLap = False
    End If
    
    Dim Arr() As Long: ReDim Arr(0 To numOfItems - 1)
    Dim i As Long, j As Long, numVal As Long
    For i = 0 To numOfItems - 1
        If Not noOverLap Then
            Arr(i) = GetRandomInt(minVal, maxVal)
        Else
            Do
                Dim flg As Boolean: flg = True
                numVal = GetRandomInt(minVal, maxVal)
                
                If i = 0 Then
                    Arr(i) = numVal
                    Exit Do
                End If
                
                For j = 0 To i - 1
                    If Arr(j) = numVal Then
                       flg = False
                       Exit For
                    End If
                Next
                
                If flg Then
                    Arr(i) = numVal
                    Exit Do
                End If
            Loop
        End If
    Next
    GetRandomIntArray = Arr
End Function

'******************************************************************************
'* [�T  �v] GetRandomString
'* [��  ��] �����_���ȕ�����𐶐�����B
'*
'* @param textLength   ������
'* @param useableChars �g�p���镶��
'* @param willRandomize Randomize�����s���邩�ǂ���
'* @return String �����_���ȕ�����
'*
'******************************************************************************
Public Function GetRandomString(ByVal textLength As Long, _
                                Optional useableChars As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", _
                                Optional willRandomize As Boolean = False) As String
    Dim tmpStr As StringEx: Set tmpStr = New StringEx
    Dim maxLen As Long: maxLen = StringUtils.CheckLength(useableChars)
    
    Dim v As Variant
    For Each v In GetRandomIntArray(textLength, 1, maxLen, , willRandomize)
        tmpStr.Append Mid$(useableChars, CLng(v), 1)
    Next
    GetRandomString = tmpStr.ToString
End Function

'******************************************************************************
'* [�T  �v] GetRandomHiragana
'* [��  ��] �����_���ȂЂ炪�ȕ�����𐶐�����B
'*
'* @param textLength   ������
'* @param willRandomize Randomize�����s���邩�ǂ���
'* @return String �����_���ȂЂ炪�ȕ�����
'*
'******************************************************************************
Public Function GetRandomHiragana(ByVal textLength As Long, _
                                  Optional willRandomize As Boolean = False) As String
    GetRandomHiragana = GetRandomString(textLength, _
        "�����������������������������������ĂƂȂɂʂ˂�" & _
        "�͂Ђӂւق܂݂ނ߂��������������" & _
        "�������������������������Âłǂ΂тԂׂڂς҂Ղ؂�", willRandomize)
End Function

'******************************************************************************
'* [�T  �v] GetRandomKatakana
'* [��  ��] �����_���ȃJ�^�J�i������𐶐�����B
'*
'* @param textLength   ������
'* @param willRandomize Randomize�����s���邩�ǂ���
'* @return String �����_���ȃJ�^�J�i������
'*
'******************************************************************************
Public Function GetRandomKatakana(ByVal textLength As Long, _
                                  Optional willRandomize As Boolean = False) As String
    GetRandomKatakana = StrConv(GetRandomHiragana(textLength, willRandomize), vbKatakana)
End Function

'******************************************************************************
'* [�T  �v] GetRandomNumString
'* [��  ��] �����_���Ȑ���������𐶐�����B
'*
'* @param textLength   ������
'* @param willRandomize Randomize�����s���邩�ǂ���
'* @return String �����_���Ȑ���������
'*
'******************************************************************************
Public Function GetRandomNumString(ByVal textLength As Long, _
                                  Optional willRandomize As Boolean = False) As String
    GetRandomNumString = GetRandomString(textLength, "0123456789", willRandomize)
End Function

'******************************************************************************
'* [�T  �v] GetRandomHalfAlphaNumeric
'* [��  ��] �����_���Ȕ��p�p��������𐶐�����B
'*
'* @param textLength   ������
'* @param willRandomize Randomize�����s���邩�ǂ���
'* @return String �����_���Ȕ��p�p��������
'*
'******************************************************************************
Public Function GetRandomHalfAlphaNumeric(ByVal textLength As Long, _
                                  Optional willRandomize As Boolean = False) As String
    GetRandomHalfAlphaNumeric = GetRandomString(textLength, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", willRandomize)
End Function

'******************************************************************************
'* [�T  �v] GetRandomHalfAlphaNumericSymbol
'* [��  ��] �����_���Ȕ��p�p��������𐶐�����B
'*
'* @param textLength   ������
'* @param willRandomize Randomize�����s���邩�ǂ���
'* @return String �����_���Ȕ��p�p��������
'*
'******************************************************************************
Public Function GetRandomHalfAlphaNumericSymbol(ByVal textLength As Long, _
                                                Optional allowSpace As Boolean = True, _
                                                Optional willRandomize As Boolean = False) As String
    GetRandomHalfAlphaNumericSymbol = GetRandomString(textLength, IIf(allowSpace, " ", "") & "!""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~", willRandomize)
End Function

'******************************************************************************
'* [�T  �v] ParallelExec
'* [��  ��] ���񏈗����s���B
'* [��  ��] �����͂Ȃ��B
'* [�Q  �l] https://www.excel-chunchun.com/entry/2019/03/27/005233
'*
'* @param fncName ���s����֐���
'*                �Ăяo���֐����ŁA�I�����Ɉȉ��̏������s�����ƁB
'*                  Application.DisplayAlerts = False
'*                  ThisWorkbook.Close False
'*                �������s��Ȃ��ꍇ�A�q�v���Z�X�̎��s�����m�ł����A
'*                �������[�v�ƂȂ�B
'* @param procNum ������s����v���Z�X��
'* @param paParams �֐��ɓn���p�����[�^
'*
'******************************************************************************
Public Sub ParallelExec(ByVal fncName As String, ByVal procNum As Integer, ParamArray paParams())
    Dim params As Variant: params = VariantUtils.EmptyArrayIfParamArrayMissing(paParams)
    Dim apps As Collection: Set apps = New Collection
    Dim app As Excel.Application
    Dim wb As Workbook
    Dim i As Long
    For i = 1 To procNum
        Set app = New Application
        apps.Add app
        
        ' ��WorkBook��ʂ̃C���X�^���X��Open�i�ǎ���p�j
        Set wb = app.Workbooks.Open(ThisWorkbook.FullName, _
                                    UpdateLinks:=False, _
                                    ReadOnly:=True)
        ' �q�v���Z�X���s
        app.Run "'" & wb.Name & "'!ParallelSubExec", i, fncName, params
        
        DoEvents
    Next
    Set app = Nothing: Set wb = Nothing
    
    ' �q�v���Z�X�I���҂�
    For i = 1 To apps.Count
        Do While apps(i).Workbooks.Count > 0
            Application.Wait [Now() + "00:00:00.2"]
            DoEvents
        Loop
    Next
    
    ' �qExcel�̃C���X�^���X�̔j��
    On Error Resume Next
    For i = 1 To apps.Count
        apps(1).Quit
        apps.Remove 1
    Next
    On Error GoTo 0
End Sub

'******************************************************************************
'* [�T  �v] ParallelSubExec
'* [��  ��] ParallelExec�̃T�u�����B
'* [�Q  �l] https://www.excel-chunchun.com/entry/2019/03/27/005233
'*
'* @param n       ���s�ԍ�
'* @param fncName ���s����֐���
'* @param params  �֐��ɓn���p�����[�^
'*
'******************************************************************************
Private Sub ParallelSubExec(n As Long, fncName As String, params As Variant)
    Dim param
    Dim sb As StringEx: Set sb = New StringEx
    Call sb.Append("'").Append(fncName).Append(" """).Append(n - 1).Append("""")
    For Each param In params
        sb.Append ", """
        sb.Append CStr(param)
        sb.Append """"
    Next
    sb.Append "'"
    Application.OnTime [Now() + "00:00:00.2"], sb.ToString
End Sub

'******************************************************************************
'* [�T  �v] OnTimeForClass
'* [��  ��] �N���X�̃��\�b�h�ɑ΂���Ontime���������s����B
'* [��  ��] �EOnTime���l�A�ʂ̏��������s���̏ꍇ�A�ʏ������I������܂őҋ@����B
'*            �i����Ŏ��s�͂���Ȃ��j
'*          �E�������s��A���s���ʂ�mCallbackResultCol�ɒ~�ς���邽�߁A
'*            �s�v�ɂȂ������s���ʂ́AClearResultOnTimeForClass ���Ăяo���āA
'*            �N���A���邱�ƁB
'*
'* @param startSec    ���s�J�n�܂ł̑ҋ@���ԁi�b�j
'* @param callbackObj ���s����N���X�̃I�u�W�F�N�g
'* @param fncName     ���s���郁�\�b�h��
'* @param paParams    �֐��ɓn���p�����[�^
'* @return String     ���s�\��L�[�i���s���ʂ��m�F����ۂɎg�p�j
'*
'******************************************************************************
Public Function OnTimeForClass(startSec As Long, callbackObj As Object, fncName As String, ParamArray paParams()) As String
    If mCallbackObjCol Is Nothing Then
        Set mCallbackObjCol = New Collection
    End If
    If mCallbackParamCol Is Nothing Then
        Set mCallbackParamCol = New Collection
    End If
    If mCallbackResultCol Is Nothing Then
        Set mCallbackResultCol = New Collection
    End If
    
    Dim keyStr As String: keyStr = Base.CreateUUID()
    Dim params As Variant: params = VariantUtils.EmptyArrayIfParamArrayMissing(paParams)
    mCallbackObjCol.Add callbackObj, keyStr
    mCallbackParamCol.Add params, keyStr
    
    Application.OnTime Now + TimeSerial(0, 0, startSec), "'OnTimeForClassSubExec """ & fncName & """, """ & keyStr & """'"
    OnTimeForClass = keyStr
End Function

'******************************************************************************
'* [�T  �v] OnTimeForClassSubExec
'* [��  ��] OnTimeForClass�̃T�u�����B
'*
'* @param fncName ���s���郁�\�b�h��
'* @param keyStr  ���s�\��L�[�i���s���ʂ��m�F����ۂɎg�p�j
'*
'******************************************************************************
Public Sub OnTimeForClassSubExec(fncName As String, keyStr As String)
    Dim ret, p: p = mCallbackParamCol(keyStr)
    Dim callbackObj As Object: Set callbackObj = mCallbackObjCol(keyStr)

    Select Case ArrayUtils.GetLength(p)
        Case 1
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0))), keyStr)
        Case 2
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1))), keyStr)
        Case 3
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2))), keyStr)
        Case 4
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3))), keyStr)
        Case 5
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3)), CVar(p(4))), keyStr)
        Case 6
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3)), CVar(p(4)), CVar(p(5))), keyStr)
        Case 7
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3)), CVar(p(4)), CVar(p(5)), CVar(p(6))), keyStr)
        Case 8
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3)), CVar(p(4)), CVar(p(5)), CVar(p(6)), CVar(p(7))), keyStr)
        Case 9
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3)), CVar(p(4)), CVar(p(5)), CVar(p(6)), CVar(p(7)), CVar(p(8))), keyStr)
        Case 10
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3)), CVar(p(4)), CVar(p(5)), CVar(p(6)), CVar(p(7)), CVar(p(8)), CVar(p(9))), keyStr)
        Case Else
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod), keyStr)
    End Select
    mCallbackObjCol.Remove keyStr
    mCallbackParamCol.Remove keyStr
End Sub

'******************************************************************************
'* [�T  �v] GetResultOnTimeForClass
'* [��  ��] OnTimeForClass�Ŏ��s���������̎��s���ʂ��擾����B
'*
'* @param keyStr  ���s�\��L�[�iOnTimeForClass�̖߂�l�j
'* @@return Variant ���s���ʁi�߂�l���Ȃ������̏ꍇEmpty�j
'*
'******************************************************************************
Public Function GetResultOnTimeForClass(keyStr As String) As Variant
    GetResultOnTimeForClass = mCallbackResultCol(keyStr)
End Function

'******************************************************************************
'* [�T  �v] GetResultOnTimeForClass
'* [��  ��] OnTimeForClass�Ŏ��s���������̎��s���ʂ��N���A����B
'*
'******************************************************************************
Public Sub ClearResultOnTimeForClass()
    Set mCallbackResultCol = New Collection
End Sub
