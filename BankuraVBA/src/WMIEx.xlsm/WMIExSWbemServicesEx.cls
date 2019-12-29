VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemServicesEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemServicesExクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemServicesExクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemServicesExクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemServicesExクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393855(v=vs.80)>
' A connection to a Namespace
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemServicesEx
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
'* Security_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The Security Configurator for this Object.
'*-----------------------------------------------------------------------------
Property Get Security_() As WMIExSWbemSecurity
    Dim wrap As WMIExSWbemSecurity: Set wrap = New WMIExSWbemSecurity
    Set wrap.Origin = mOrigin.Security_()
    Set Security_ = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemServices")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] AssociatorsOf メソッド
'* [詳  細] AssociatorsOf のラッパーメソッド。
'*          Get the Associators of a class or instance.
'*
'* @param strObjectPath xx
'* @param strAssocClass xx
'* @param strResultClass xx
'* @param strResultRole xx
'* @param strRole xx
'* @param bClassesOnly xx
'* @param bSchemaOnly xx
'* @param strRequiredAssocQualifier xx
'* @param strRequiredQualifier xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemObjectSet
'*
'******************************************************************************
Function AssociatorsOf(strObjectPath As String, _
                       Optional strAssocClass As String, _
                       Optional strResultClass As String, _
                       Optional strResultRole As String, _
                       Optional strRole As String, _
                       Optional bClassesOnly As Boolean = False, _
                       Optional bSchemaOnly As Boolean = False, _
                       Optional strRequiredAssocQualifier As String, _
                       Optional strRequiredQualifier As String, _
                       Optional iFlags As Long = 16, _
                       Optional objWbemNamedValueSet As Object) As WMIExSWbemObjectSet
    Dim wrap As WMIExSWbemObjectSet: Set wrap = New WMIExSWbemObjectSet
    Set wrap.Origin = mOrigin.AssociatorsOf(strObjectPath, strAssocClass, strResultClass, strResultRole, strRole, bClassesOnly, bSchemaOnly, strRequiredAssocQualifier, strRequiredQualifier, iFlags, objWbemNamedValueSet)
    Set AssociatorsOf = wrap
End Function

'******************************************************************************
'* [概  要] AssociatorsOfAsync メソッド
'* [詳  細] AssociatorsOfAsync のラッパーメソッド。
'*          Get the Associators of a class or instance asynchronously.
'*
'* @param objWbemSink xx
'* @param strObjectPath xx
'* @param strAssocClass xx
'* @param strResultClass xx
'* @param strResultRole xx
'* @param strRole xx
'* @param bClassesOnly xx
'* @param bSchemaOnly xx
'* @param strRequiredAssocQualifier xx
'* @param strRequiredQualifier xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'*
'******************************************************************************
Sub AssociatorsOfAsync(objWbemSink As Object, _
                       strObjectPath As String, _
                       Optional strAssocClass As String, _
                       Optional strResultClass As String, _
                       Optional strResultRole As String, _
                       Optional strRole As String, _
                       Optional bClassesOnly As Boolean = False, _
                       Optional bSchemaOnly As Boolean = False, _
                       Optional strRequiredAssocQualifier As String, _
                       Optional strRequiredQualifier As String, _
                       Optional iFlags As Long, _
                       Optional objWbemNamedValueSet As Object, _
                       Optional objWbemAsyncContext As Object)
    Call mOrigin.AssociatorsOfAsync(objWbemSink, strObjectPath, strAssocClass, strResultClass, strResultRole, strRole, bClassesOnly, bSchemaOnly, strRequiredAssocQualifier, strRequiredQualifier, iFlags, objWbemNamedValueSet)
End Sub

'******************************************************************************
'* [概  要] Delete メソッド
'* [詳  細] Delete のラッパーメソッド。
'*          Delete a Class or Instance.
'*
'* @param strObjectPath xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'*
'******************************************************************************
Sub Delete(strObjectPath As String, Optional iFlags As Long, Optional objWbemNamedValueSet As Object)
    'TODO:Optional taiou
    Call mOrigin.Delete(strObjectPath, iFlags, objWbemNamedValueSet)
End Sub

'******************************************************************************
'* [概  要] DeleteAsync メソッド
'* [詳  細] DeleteAsync のラッパーメソッド。
'*          Delete a Class or Instance asynchronously.
'*
'* @param objWbemSink xx
'* @param strObjectPath xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub DeleteAsync(objWbemSink As Object, strObjectPath As String, Optional iFlags As Long, Optional objWbemNamedValueSet As Object, Optional objWbemAsyncContext As Object)
    Call mOrigin.DeleteAsync(objWbemSink, strObjectPath, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub

'******************************************************************************
'* [概  要] ExecMethod メソッド
'* [詳  細] ExecMethod のラッパーメソッド。
'*          Execute a Method.
'*
'* @param strObjectPath xx
'* @param strMethodName xx
'* @param objWbemInParameters xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemObject
'*
'******************************************************************************
Function ExecMethod(strObjectPath As String, strMethodName As String, Optional objWbemInParameters As Object, Optional iFlags As Long, Optional objWbemNamedValueSet As Object) As WMIExSWbemObject
    Dim wrap As WMIExSWbemObject: Set wrap = New WMIExSWbemObject
    Set wrap.Origin = mOrigin.ExecMethod(strObjectPath, strMethodName, objWbemInParameters, iFlags, objWbemNamedValueSet)
    Set ExecMethod = wrap
End Function

'******************************************************************************
'* [概  要] ExecMethodAsync メソッド
'* [詳  細] ExecMethodAsync のラッパーメソッド。
'*          Execute a Method asynchronously.
'*
'* @param objWbemSink xx
'* @param strObjectPath xx
'* @param strMethodName xx
'* @param objWbemInParameters xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub ExecMethodAsync(objWbemSink As Object, _
                    strObjectPath As String, _
                    strMethodName As String, _
                    Optional objWbemInParameters As Object, _
                    Optional iFlags As Long, _
                    Optional objWbemNamedValueSet As Object, _
                    Optional objWbemAsyncContext As Object)
    Call mOrigin.ExecMethodAsync(objWbemSink, strObjectPath, strMethodName, objWbemInParameters, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub

'******************************************************************************
'* [概  要] ExecNotificationQuery メソッド
'* [詳  細] ExecNotificationQuery のラッパーメソッド。
'*          Execute a Query to receive Notifications.
'*
'* @param strQuery xx
'* @param strQueryLanguage xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemEventSource
'*
'******************************************************************************
Function ExecNotificationQuery(strQuery As String, _
                               Optional strQueryLanguage As String = "WQL", _
                               Optional iFlags As Long = 48, _
                               Optional objWbemNamedValueSet As Object) As WMIExSWbemEventSource
    Dim wrap As WMIExSWbemEventSource: Set wrap = New WMIExSWbemEventSource
    Set wrap.Origin = mOrigin.ExecNotificationQuery(strQuery, strQueryLanguage, iFlags, objWbemNamedValueSet)
    Set ExecNotificationQuery = wrap
End Function

'******************************************************************************
'* [概  要] ExecNotificationQueryAsync メソッド
'* [詳  細] ExecNotificationQueryAsync のラッパーメソッド。
'*          Execute an asynchronous Query to receive Notifications.
'*
'* @param objWbemSink xx
'* @param strQuery xx
'* @param strQueryLanguage xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub ExecNotificationQueryAsync(objWbemSink As Object, _
                               strQuery As String, _
                               Optional strQueryLanguage As String = "WQL", _
                               Optional iFlags As Long, _
                               Optional objWbemNamedValueSet As Object, _
                               Optional objWbemAsyncContext As Object)
    Call mOrigin.ExecNotificationQueryAsync(objWbemSink, strQuery, strQueryLanguage, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub

'******************************************************************************
'* [概  要] ExecQuery メソッド
'* [詳  細] ExecQuery のラッパーメソッド。
'*          Execute a Query.
'*
'* @param strQuery xx
'* @param strQueryLanguage xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemEventSource
'*
'******************************************************************************
Function ExecQuery(strQuery As String, _
                   Optional strQueryLanguage As String = "WQL", _
                   Optional iFlags As Long = 16, _
                   Optional objWbemNamedValueSet As Object) As WMIExSWbemObjectSet
    Dim wrap As WMIExSWbemObjectSet: Set wrap = New WMIExSWbemObjectSet
    Set wrap.Origin = mOrigin.ExecQuery(strQuery, strQueryLanguage, iFlags, objWbemNamedValueSet)
    Set ExecQuery = wrap
End Function

'******************************************************************************
'* [概  要] ExecQueryAsync メソッド
'* [詳  細] ExecQueryAsync のラッパーメソッド。
'*          Execute an asynchronous Query.
'*
'* @param objWbemSink xx
'* @param strQuery xx
'* @param strQueryLanguage xx
'* @param lFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub ExecQueryAsync(objWbemSink As Object, _
                   strQuery As String, _
                   Optional strQueryLanguage As String = "WQL", _
                   Optional lFlags As Long, _
                   Optional objWbemNamedValueSet As Object, _
                   Optional objWbemAsyncContext As Object)
    Call mOrigin.ExecQueryAsync(objWbemSink, strQuery, strQueryLanguage, lFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub

'******************************************************************************
'* [概  要] Get_ メソッド
'* [詳  細] Get のラッパーメソッド。
'*          Get a single Class or Instance.
'*
'* @param strObjectPath xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemObject
'*
'******************************************************************************
Function Get_(Optional strObjectPath As String, Optional iFlags As Long, Optional objWbemNamedValueSet As Object) As WMIExSWbemObject
    Dim wrap As WMIExSWbemObject: Set wrap = New WMIExSWbemObject
    Set wrap.Origin = mOrigin.Get(strObjectPath, iFlags, objWbemNamedValueSet)
    Set Get_ = wrap
End Function

'******************************************************************************
'* [概  要] GetAsync メソッド
'* [詳  細] GetAsync のラッパーメソッド。
'*          Get a single Class or Instance asynchronously.
'*
'* @param objWbemSink xx
'* @param strObjectPath xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub GetAsync(objWbemSink As Object, _
             Optional strObjectPath As String, _
             Optional iFlags As Long, _
             Optional objWbemNamedValueSet As Object, _
             Optional objWbemAsyncContext As Object)
    Call mOrigin.GetAsync(objWbemSink, strObjectPath, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub

'******************************************************************************
'* [概  要] InstancesOf メソッド
'* [詳  細] InstancesOf のラッパーメソッド。
'*          Enumerate the Instances of a Class.
'*
'* @param strClass xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemObjectSet
'*
'******************************************************************************
Function InstancesOf(strClass As String, Optional iFlags As Long = 16, Optional objWbemNamedValueSet As Object) As WMIExSWbemObjectSet
    Dim wrap As WMIExSWbemObjectSet: Set wrap = New WMIExSWbemObjectSet
    Set wrap.Origin = mOrigin.InstancesOf(strClass, iFlags, objWbemNamedValueSet)
    Set InstancesOf = wrap
End Function

'******************************************************************************
'* [概  要] InstancesOfAsync メソッド
'* [詳  細] InstancesOfAsync のラッパーメソッド。
'*          Enumerate the Instances of a Class asynchronously.
'*
'* @param objWbemSink xx
'* @param strClass xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub InstancesOfAsync(objWbemSink As Object, strClass As String, Optional iFlags As Long, Optional objWbemNamedValueSet As Object, Optional objWbemAsyncContext As Object)
    Call mOrigin.InstancesOfAsync(objWbemSink, strClass, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub

'******************************************************************************
'* [概  要] InstancesOf メソッド
'* [詳  細] InstancesOf のラッパーメソッド。
'*          Save the Object to this Namespace.
'*
'* @param objWbemObject xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemObjectPath
'*
'******************************************************************************
Function Put_(objWbemObject As WMIExSWbemObjectEx, Optional iFlags As Long, Optional objWbemNamedValueSet As Object) As WMIExSWbemObjectPath
    Dim wrap As WMIExSWbemObjectPath: Set wrap = New WMIExSWbemObjectPath
    Set wrap.Origin = mOrigin.Put(objWbemObject.Origin, iFlags, objWbemNamedValueSet)
    Set Put_ = wrap
End Function

'******************************************************************************
'* [概  要] PutAsync メソッド
'* [詳  細] PutAsync のラッパーメソッド。
'*          Save the Object to this Namespace asynchronously.
'*
'* @param objWbemSink xx
'* @param objWbemObject xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub PutAsync(objWbemSink As WMIExSWbemSink, objWbemObject As WMIExSWbemObjectEx, Optional iFlags As Long, Optional objWbemNamedValueSet As Object, Optional objWbemAsyncContext As Object)
    Call mOrigin.PutAsync(objWbemSink.Origin, objWbemObject.Origin, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub
    
'******************************************************************************
'* [概  要] ReferencesTo メソッド
'* [詳  細] ReferencesTo のラッパーメソッド。
'*          Get the References to a class or instance.
'*
'* @param strObjectPath xx
'* @param strResultClass xx
'* @param strRole xx
'* @param bClassesOnly xx
'* @param bSchemaOnly xx
'* @param strRequiredQualifier xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemObjectSet
'*
'******************************************************************************
Function ReferencesTo(strObjectPath As String, _
                      Optional strResultClass As String, _
                      Optional strRole As String, _
                      Optional bClassesOnly As Boolean = False, _
                      Optional bSchemaOnly As Boolean = False, _
                      Optional strRequiredQualifier As String, _
                      Optional iFlags As Long = 16, _
                      Optional objWbemNamedValueSet As Object) As WMIExSWbemObjectSet
    Dim wrap As WMIExSWbemObjectSet: Set wrap = New WMIExSWbemObjectSet
    Set wrap.Origin = mOrigin.ReferencesTo(strObjectPath, strResultClass, strRole, bClassesOnly, bSchemaOnly, strRequiredQualifier, iFlags, objWbemNamedValueSet)
    Set ReferencesTo = wrap
End Function

'******************************************************************************
'* [概  要] ReferencesToAsync メソッド
'* [詳  細] ReferencesToAsync のラッパーメソッド。
'*          Get the References to a class or instance asynchronously.
'*
'* @param objWbemSink xx
'* @param strObjectPath xx
'* @param strResultClass xx
'* @param strRole xx
'* @param bClassesOnly xx
'* @param bSchemaOnly xx
'* @param strRequiredQualifier xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub ReferencesToAsync(objWbemSink As Object, _
                      strObjectPath As String, _
                      Optional strResultClass As String, _
                      Optional strRole As String, _
                      Optional bClassesOnly As Boolean = False, _
                      Optional bSchemaOnly As Boolean = False, _
                      Optional strRequiredQualifier As String, _
                      Optional iFlags As Long, _
                      Optional objWbemNamedValueSet As Object, _
                      Optional objWbemAsyncContext As Object)
    Call mOrigin.ReferencesToAsync(objWbemSink, strObjectPath, strResultClass, strRole, bClassesOnly, bSchemaOnly, strRequiredQualifier, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub

'******************************************************************************
'* [概  要] SubclassesOf メソッド
'* [詳  細] SubclassesOf のラッパーメソッド。
'*          Enumerate subclasses of this Class.
'*
'* @param strSuperclass xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemObjectSet
'*
'******************************************************************************
Function SubclassesOf(Optional strSuperclass As String, Optional iFlags As Long = 16, Optional objWbemNamedValueSet As Object) As WMIExSWbemObjectSet
    Dim wrap As WMIExSWbemObjectSet: Set wrap = New WMIExSWbemObjectSet
    Set wrap.Origin = mOrigin.SubclassesOf(strSuperclass, iFlags, objWbemNamedValueSet)
    Set SubclassesOf = wrap
End Function

'******************************************************************************
'* [概  要] SubclassesAsync_ メソッド
'* [詳  細] SubclassesAsync_ のラッパーメソッド。
'*          Enumerate the subclasses of a Class asynchronously.
'*
'* @param objWbemSink xx
'* @param strSuperclass xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub SubclassesOfAsync(objWbemSink As Object, Optional strSuperclass As String, Optional iFlags As Long, Optional objWbemNamedValueSet As Object, Optional objWbemAsyncContext As Object)
    Call mOrigin.SubclassesOfAsync(objWbemSink, strSuperclass, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub


'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] CreateInstance メソッド
'* [詳  細] GetObject メソッドを使用してSWbemServicesオブジェクトを生成する。
'*
'* @param strComputer 省略可。コンピュータ名。
'* @return WMIExSWbemServices SWbemServicesラップオブジェクト。
'*
'******************************************************************************
Function CreateInstance(Optional strComputer As String = ".") As WMIExSWbemServicesEx
    Set Me.Origin = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set CreateInstance = Me
End Function
