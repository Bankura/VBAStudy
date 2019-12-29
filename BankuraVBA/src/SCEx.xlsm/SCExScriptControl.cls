VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SCExScriptControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] MSScriptControl.ScriptControl クラスのラップ・拡張クラス
'* [詳  細] MSScriptControl.ScriptControl クラスのWrapperとしての機能を提供する他、
'*          MSScriptControl.ScriptControl クラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' MSScriptControl.ScriptControl クラスについて。
' Control to host scripting engines that understand the ActiveX Scripting interface.
'
' [イベント]
' Event Error()
'     Event fired when any error occurs in the scripting engine
' Event Timeout()
'     Event fired when a script is aborted because of a timeout
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As MSScriptControl.ScriptControl
Private mOrigin As Object

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* オリジナルObject プロパティ
'*-----------------------------------------------------------------------------
Property Get Origin() As Object
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Object)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* AllowUI プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Enable or disable display of the UI.
'*-----------------------------------------------------------------------------
Property Get AllowUI() As Boolean
    AllowUI = mOrigin.AllowUI
End Property

'*-----------------------------------------------------------------------------
'* CodeObject プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Object exposed by the scripting engine that contains methods and properties defined in the code added to the global module.
'*-----------------------------------------------------------------------------
Property Get CodeObject() As Object
    Set CodeObject = mOrigin.CodeObject
End Property
    
'*-----------------------------------------------------------------------------
'* Error プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The last error reported by the scripting engine.
'*-----------------------------------------------------------------------------
Property Get Error() As SCExError
    Dim wrap As SCExError: Set wrap = New SCExError
    Set wrap.Origin = mOrigin.Error
    Set Error = wrap
End Property

'*-----------------------------------------------------------------------------
'* Language プロパティ
'*
'* [補  足]
'* ・Language engine to use.
'*   "JScript",  "VBScript", "JavaScript"
'*-----------------------------------------------------------------------------
Property Get Language() As String
    Language = mOrigin.Language
End Property
Property Let Language(arg As String)
    mOrigin.Language = arg
End Property

'*-----------------------------------------------------------------------------
'* Modules プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Collection of modules for the ScriptControl.
'*-----------------------------------------------------------------------------
Property Get Modules() As SCExModules
    Dim wrap As SCExModules: Set wrap = New SCExModules
    Set wrap.Origin = mOrigin.Modules
    Set Modules = wrap
End Property

'*-----------------------------------------------------------------------------
'* Procedures プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Collection of procedures that are defined in the global module.
'*-----------------------------------------------------------------------------
Property Get Procedures() As SCExProcedures
    Dim wrap As SCExProcedures: Set wrap = New SCExProcedures
    Set wrap.Origin = mOrigin.Procedures
    Set Procedures = wrap
End Property

'*-----------------------------------------------------------------------------
'* SitehWnd プロパティ
'*
'* [補  足]
'* ・hWnd used as a parent for displaying UI.
'*-----------------------------------------------------------------------------
Property Get SitehWnd() As Long
    SitehWnd = mOrigin.SitehWnd
End Property
Property Let SitehWnd(arg As Long)
    mOrigin.SitehWnd = arg
End Property
    
'*-----------------------------------------------------------------------------
'* State プロパティ
'*
'* [補  足]
'* ・State of the control.
'*-----------------------------------------------------------------------------
Property Get State() As ScriptControlStates
    State = mOrigin.State
End Property
Property Let State(arg As ScriptControlStates)
    mOrigin.State = arg
End Property

'*-----------------------------------------------------------------------------
'* Timeout プロパティ
'*
'* [補  足]
'* ・Length of time in milliseconds that a script can execute before being considered hung.
'*-----------------------------------------------------------------------------
Property Get Timeout() As Long
    Timeout = mOrigin.Timeout
End Property
Property Let Timeout(arg As Long)
    mOrigin.Timeout = arg
End Property

'*-----------------------------------------------------------------------------
'* UseSafeSubset プロパティ
'*
'* [補  足]
'* ・Force script to execute in safe mode and disallow potentially harmful actions.
'*-----------------------------------------------------------------------------
Property Get UseSafeSubset() As Boolean
    UseSafeSubset = mOrigin.UseSafeSubset
End Property
Property Let UseSafeSubset(arg As Boolean)
    mOrigin.UseSafeSubset = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject32bit("MSScriptControl.ScriptControl")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] AddCode メソッド
'* [詳  細] AddCode のラッパーメソッド。
'*          Add code to the global module.
'*
'* @param Code xx
'*
'******************************************************************************
Sub AddCode(Code As String)
    Call mOrigin.AddCode(Code)
End Sub

'******************************************************************************
'* [概  要] AddObject メソッド
'* [詳  細] AddObject のラッパーメソッド。
'*          Add an object to the global namespace of the scripting engine.
'*
'* @param Name xx
'* @param Object xx
'* @param AddMembers xx
'*
'******************************************************************************
Sub AddObject(Name As String, Object As Object, Optional AddMembers As Boolean = False)
    Call mOrigin.AddObject(Name, Object, AddMembers)
End Sub
  
'******************************************************************************
'* [概  要] Eval メソッド
'* [詳  細] Eval のラッパーメソッド。
'*          Evaluate an expression within the context of the global module.
'*
'* @param Expression xx
'* @return Variant
'*
'******************************************************************************
Function Eval(Expression As String)
    Eval = mOrigin.Eval(Expression)
End Function

'******************************************************************************
'* [概  要] ExecuteStatement メソッド
'* [詳  細] ExecuteStatement のラッパーメソッド。
'*          Execute a statement within the context of the global module.
'*
'* @param Statement xx
'*
'******************************************************************************
Sub ExecuteStatement(Statement As String)
    Call mOrigin.ExecuteStatement(Statement)
End Sub

'******************************************************************************
'* [概  要] Reset メソッド
'* [詳  細] Reset のラッパーメソッド。
'*          Reset the scripting engine to a newly created state.
'*
'******************************************************************************
Sub Reset()
    Call mOrigin.Reset
End Sub

'******************************************************************************
'* [概  要] Run メソッド
'* [詳  細] Run のラッパーメソッド。
'*          Call a procedure defined in the global module.
'*
'* @param ProcedureName xx
'* @param Parameters() xx
'* @return Variant
'*
'******************************************************************************
Function Run(ProcedureName As String, ParamArray Parameters() As Variant)
    Select Case UBound(Parameters)
        Case LBound(Parameters)
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)))
        Case LBound(Parameters) + 1
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1))
        Case LBound(Parameters) + 2
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2))
        Case LBound(Parameters) + 3
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3))
        Case LBound(Parameters) + 4
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3), _
                                             Parameters(LBound(Parameters) + 4))
        Case LBound(Parameters) + 5
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3), _
                                             Parameters(LBound(Parameters) + 4), _
                                             Parameters(LBound(Parameters) + 5))
        Case LBound(Parameters) + 6
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3), _
                                             Parameters(LBound(Parameters) + 4), _
                                             Parameters(LBound(Parameters) + 5), _
                                             Parameters(LBound(Parameters) + 6))
        Case LBound(Parameters) + 7
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3), _
                                             Parameters(LBound(Parameters) + 4), _
                                             Parameters(LBound(Parameters) + 5), _
                                             Parameters(LBound(Parameters) + 6), _
                                             Parameters(LBound(Parameters) + 7))
        Case LBound(Parameters) + 8
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3), _
                                             Parameters(LBound(Parameters) + 4), _
                                             Parameters(LBound(Parameters) + 5), _
                                             Parameters(LBound(Parameters) + 6), _
                                             Parameters(LBound(Parameters) + 7), _
                                             Parameters(LBound(Parameters) + 8))
        Case LBound(Parameters) + 9
            Run = mOrigin.Run(ProcedureName, Parameters(LBound(Parameters)), _
                                             Parameters(LBound(Parameters) + 1), _
                                             Parameters(LBound(Parameters) + 2), _
                                             Parameters(LBound(Parameters) + 3), _
                                             Parameters(LBound(Parameters) + 4), _
                                             Parameters(LBound(Parameters) + 5), _
                                             Parameters(LBound(Parameters) + 6), _
                                             Parameters(LBound(Parameters) + 7), _
                                             Parameters(LBound(Parameters) + 8), _
                                             Parameters(LBound(Parameters) + 9))
    End Select
End Function


'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] EncodeURI メソッド
'* [詳  細] URIエンコード処理を行います。
'* [参  考] <https://github.com/vocho/vbs/blob/a5c3ee608103638678c983da00ec290c4b8ab90c/CreateObject32bit.vbs>
'*          <https://srad.jp/~black-hole/journal/516080/>
'*
'* @param strURI 未エンコードURI
'* @return エンコード済みURIt
'*
'******************************************************************************
Function EncodeURI(ByVal strURI As String) As String
    With Me
        .Language = "JScript"
        EncodeURI = .CodeObject.encodeURIComponent(strURI)
    End With
End Function

'******************************************************************************
'* [概  要] DecodeURI メソッド
'* [詳  細] URIデコード処理を行います。
'* [参  考] <https://github.com/vocho/vbs/blob/a5c3ee608103638678c983da00ec290c4b8ab90c/CreateObject32bit.vbs>
'*          <https://srad.jp/~black-hole/journal/516080/>
'*
'* @param strURI 未デコードURI
'* @return デコード済みURI
'*
'******************************************************************************
Function DecodeURI(ByVal strURI As String) As String
    With Me
        .Language = "JScript"
        DecodeURI = .CodeObject.decodeURIComponent(strURI)
    End With
End Function


'******************************************************************************
'* [概  要] CreateObject32bit メソッド
'* [詳  細] 32ビット環境のObjectを生成します。
'* [参  考] <https://github.com/vocho/vbs/blob/a5c3ee608103638678c983da00ec290c4b8ab90c/CreateObject32bit.vbs>
'*
'* @param strClassName 生成対象のクラス名。"Shell.Application"等。
'* @return 32ビット環境Object
'*
'******************************************************************************
Private Function CreateObject32bit(strClassName As String)
    If Is32BitProcessorForApp Then
     Set CreateObject32bit = CreateObject(strClassName)
     Exit Function
    End If
    
    Dim shell As Object: Set shell = CreateObject("Shell.Application")
    shell.Windows().Item(0).PutProperty strClassName, Nothing
    ' 一時スクリプトコマンドテキスト生成
    Dim strScript As String
    strScript = "CreateObject(""Shell.Application"").Windows().Item(0).PutProperty """ & strClassName & """, CreateObject(""" & strClassName & """)" & vbNewLine & _
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

    ' 一時スクリプトファイル作成
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim strTempFile
    With fso
        Do
            strTempFile = .BuildPath(.GetSpecialFolder(2), .GetTempName() & ".vbs") ' Const TemporaryFolder = 2
        Loop While .FileExists(strTempFile)
        With .OpenTextFile(strTempFile, 2, True) ' ForWriting = 2, ForAppending = 8
            .WriteLine strScript
            .Close
        End With
    End With
    
    ' 一時スクリプトファイル実行(32bit)
    With CreateObject("WScript.Shell").Environment("Process")
        .Item("SysWOW64") = fso.BuildPath(.Item("SystemRoot"), "SysWOW64")
        .Item("WScriptName") = fso.GetFileName("C:\WINDOWS\SysWOW64\cscript.exe")
        .Item("WScriptWOW64") = fso.BuildPath(.Item("SysWOW64"), .Item("WScriptName"))
        .Item("Run") = .Item("WScriptWOW64") & " """ & strTempFile & """"
         CreateObject("WScript.Shell").Run .Item("Run"), True
    End With
    
    ' オブジェクト受け取り
    Do
        Set CreateObject32bit = shell.Windows().Item(0).GetProperty(strClassName)
    Loop While CreateObject32bit Is Nothing
End Function

'******************************************************************************
'* [概  要] Is32BitProcessorForApp メソッド
'* [詳  細] 使用するアプリケーションが32ビットかをチェックします。
'*
'* @return チェック結果（True: 32Bit、False: 64bit）
'*
'******************************************************************************
Private Function Is32BitProcessorForApp() As Boolean
    Dim proc As String: proc = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    If proc = "x86" Then
       Is32BitProcessorForApp = True
    Else
       Is32BitProcessorForApp = False
    End If
End Function
