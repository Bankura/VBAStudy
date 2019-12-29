VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#Const USE_REFERENCE = False    '参照設定使用有無

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemSinkクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemSinkクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemSinkクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemSinkクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393877(v=vs.80)>
' A sink for events arising from asynchronous operations
'
' [イベント情報]
'Event OnCompleted(iHResult As WbemErrorEnum, objWbemErrorObject As SWbemObject, objWbemAsyncContext As SWbemNamedValueSet)
'    Event triggered when an asynchronous operation is completed
'Event OnObjectPut(objWbemObjectPath As SWbemObjectPath, objWbemAsyncContext As SWbemNamedValueSet)
'    Event triggered when an object path is available following a Put operation
'Event OnObjectReady(objWbemObject As SWbemObject, objWbemAsyncContext As SWbemNamedValueSet)
'    Event triggered when an Object is available
'Event OnProgress(iUpperBound As Long, iCurrent As Long, strMessage As String, objWbemAsyncContext As SWbemNamedValueSet
'    Event triggered to report the progress of an asynchronous operation
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
#If USE_REFERENCE Then
    Private WithEvents mOrigin As WbemScripting.SWbemSink
Attribute mOrigin.VB_VarHelpID = -1
#Else
    Private mOrigin As Object
#End If
    
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

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemSink")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Cancel メソッド
'* [詳  細] Cancel のラッパーメソッド。
'*          Cancel an asynchronous operation.
'*
'******************************************************************************
Sub Cancel()
    Call mOrigin.Cancel
End Sub
    
'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] CreateInstance メソッド
'* [詳  細] タイプ名を指定してSWbemSinkオブジェクトを作成する。
'*
'* @return WMIExSWbemSink SWbemSinkラップオブジェクト。
'*
'******************************************************************************
Function CreateInstance() As WMIExSWbemSink
#If USE_REFERENCE Then
    Set mOrigin = New WbemScripting.SWbemSink
#Else
    Set mOrigin = CreateObject("WbemScripting.SWbemSink")
#End If
    Set CreateInstance = Me
End Function

'******************************************************************************
'* [概  要] ExecNTLogEvent メソッド
'* [詳  細] Win32_NTLogEventを非同期実行してイベントを発生させる。
'* [備  考] 実験用。参照設定が必要。
'*
'******************************************************************************
#If USE_REFERENCE Then
Sub ExecNTLogEvent()
    Dim objWMIService As New WMIExSWbemServicesEx
    Set objWMIService = objWMIService.CreateInstance()
    objWMIService.InstancesOfAsync mOrigin, "Win32_NTLogEvent"
    Dim i As Long
    For i = 0 To 10000
        DoEvents
    Next i
End Sub
#End If

'------------------------------------------------------------------------------
' [NOTE]
' ・イベント処理を行う場合、参照設定が必要
' ・mOriginの定義を「Private WithEvents mOrigin As WbemScripting.SWbemSink」とする
' ・オブジェクトは「Set mOrigin = New WbemScripting.SWbemSink」のようにNewする
'   必要あり（CreateObjectではダメ）
'------------------------------------------------------------------------------
#If USE_REFERENCE Then
' Event triggered when an asynchronous operation is completed
Sub mOrigin_OnCompleted(ByVal iHResult As WbemScripting.WbemErrorEnum, ByVal objWbemErrorObject As WbemScripting.SWbemObject, ByVal objWbemAsyncContext As WbemScripting.SWbemNamedValueSet)
    Debug.Print "Asynchronous operation is done."
End Sub

 
' Event triggered when an object path is available following a Put operation
Sub mOrigin_OnObjectPut(ByVal objWbemObjectPath As SWbemObjectPath, ByVal objWbemAsyncContext As SWbemNamedValueSet)
    Debug.Print "An object path is available following a Put operation."
End Sub

' Event triggered when an Object is available
Sub mOrigin_OnObjectReady(ByVal objEvent As SWbemObject, ByVal objAsyncContext As SWbemNamedValueSet)
    Debug.Print "Category: " & objEvent.Category
    Debug.Print "Computer Name: " & objEvent.ComputerName
    Debug.Print "Event Code: " & objEvent.EventCode
    Debug.Print "Message: " & objEvent.message
    Debug.Print "Record Number: " & objEvent.RecordNumber
    Debug.Print "Source Name: " & objEvent.SourceName
    Debug.Print "Time Written: " & objEvent.TimeWritten
    Debug.Print "Event Type: " & objEvent.Type
    Debug.Print "User: " & objEvent.User
End Sub

' Event triggered to report the progress of an asynchronous operation
Sub mOrigin_OnProgress(ByVal iUpperBound As Long, ByVal iCurrent As Long, ByVal strMessage As String, ByVal objWbemAsyncContext As SWbemNamedValueSet)
    Debug.Print "report the progress of an asynchronous operation."
End Sub
#End If
