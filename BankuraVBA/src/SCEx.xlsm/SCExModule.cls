VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SCExModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] MSScriptControl.Module クラスのラップ・拡張クラス
'* [詳  細] MSScriptControl.Module クラスのWrapperとしての機能を提供する他、
'*          MSScriptControl.Module クラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' MSScriptControl.Module クラスについて。
' Context in which functions can be defined and expressions can be evaluated
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As MSScriptControl.Module
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
'* CodeObject プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Object exposed by the scripting engine that contains methods and properties defined in the code added to the module.
'*-----------------------------------------------------------------------------
Property Get CodeObject() As Object
    Set CodeObject = mOrigin.CodeObject
End Property

'*-----------------------------------------------------------------------------
'* Name プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・Name of the module.
'*-----------------------------------------------------------------------------
Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = mOrigin.Name
End Property

'*-----------------------------------------------------------------------------
'* Procedures プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Collection of procedures that are defined in the module.
'*-----------------------------------------------------------------------------
Property Get Procedures() As SCExProcedures
    Dim wrap As SCExProcedures: Set wrap = New SCExProcedures
    Set wrap.Origin = mOrigin.Procedures
    Set Procedures = wrap
End Property


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("MSScriptControl.Module")
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
'*          Add code to the module.
'*
'* @param Code xx
'*
'******************************************************************************
Sub AddCode(Code As String)
    Call mOrigin.AddCode(Code)
End Sub

'******************************************************************************
'* [概  要] Eval メソッド
'* [詳  細] Eval のラッパーメソッド。
'*          Evaluate an expression within the context of the module.
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
'*          Execute a statement within the context of the module.
'*
'* @param Statement xx
'*
'******************************************************************************
Sub ExecuteStatement(Statement As String)
    Call mOrigin.ExecuteStatement(Statement)
End Sub
    
'******************************************************************************
'* [概  要] Run メソッド
'* [詳  細] Run のラッパーメソッド。
'*          Call a procedure defined in the module.
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
        Case Else
            '未対応
            Run = mOrigin.Run(ProcedureName, Parameters())
    End Select
End Function
