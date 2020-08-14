VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemObjectEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemObjectExクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemObjectExクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemObjectExクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemObjectExクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393742(v=vs.80)>
' A Class or Instance
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemObjectEx
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
'* Derivation_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・An array of strings describing the class derivation heirarchy,
'*   in most-derived-from order (the first element in the array defines
'*   the superclass and the last element defines the dynasty class).
'*-----------------------------------------------------------------------------
Property Get Derivation_() As Variant
    If IsObject(mOrigin.Derivation_) Then
        Set Derivation_ = mOrigin.Derivation_
    Else
        Derivation_ = mOrigin.Derivation_
    End If
End Property

'*-----------------------------------------------------------------------------
'* Methods_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The collection of Methods of this Object.
'*-----------------------------------------------------------------------------
Property Get Methods_() As WMIExSWbemMethodSet
    Dim wrap As WMIExSWbemMethodSet: Set wrap = New WMIExSWbemMethodSet
    Set wrap.Origin = mOrigin.Methods_()
    Set Methods_ = wrap
End Property

'*-----------------------------------------------------------------------------
'* Path_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The path of this Object.
'*-----------------------------------------------------------------------------
Property Get Path_() As WMIExSWbemObjectPath
    Dim wrap As WMIExSWbemObjectPath: Set wrap = New WMIExSWbemObjectPath
    Set wrap.Origin = mOrigin.Path_()
    Set Path_ = wrap
End Property

'*-----------------------------------------------------------------------------
'* Properties_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The collection of Properties of this Object.
'*-----------------------------------------------------------------------------
Property Get Properties_() As WMIExSWbemPropertySet
    Dim wrap As WMIExSWbemPropertySet: Set wrap = New WMIExSWbemPropertySet
    Set wrap.Origin = mOrigin.Properties_()
    Set Properties_ = wrap
End Property

'*-----------------------------------------------------------------------------
'* Qualifiers_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The collection of Qualifiers of this Object.
'*-----------------------------------------------------------------------------
Property Get Qualifiers_() As WMIExSWbemQualifierSet
    Dim wrap As WMIExSWbemQualifierSet: Set wrap = New WMIExSWbemQualifierSet
    Set wrap.Origin = mOrigin.Qualifiers_()
    Set Qualifiers_ = wrap
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

'*-----------------------------------------------------------------------------
'* SystemProperties_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The collection of System Properties of this Object.
'*-----------------------------------------------------------------------------
Property Get SystemProperties_() As WMIExSWbemPropertySet
    Dim wrap As WMIExSWbemPropertySet: Set wrap = New WMIExSWbemPropertySet
    Set wrap.Origin = mOrigin.SystemProperties_()
    Set SystemProperties_ = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemObjectEx")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Associators_ メソッド
'* [詳  細] Associators_ のラッパーメソッド。
'*          Get the Associators of this Object.
'*
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
Function Associators_(Optional strAssocClass As String, _
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
    Set wrap.Origin = mOrigin.Associators_(strAssocClass, strResultClass, strResultRole, strRole, bClassesOnly, bSchemaOnly, strRequiredAssocQualifier, strRequiredQualifier, iFlags, objWbemNamedValueSet)
    Set Associators_ = wrap
End Function

'******************************************************************************
'* [概  要] AssociatorsAsync_ メソッド
'* [詳  細] AssociatorsAsync_ のラッパーメソッド。
'*          Get the Associators of this Object asynchronously.
'*
'* @param objWbemSink xx
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
Sub AssociatorsAsync_(objWbemSink As Object, _
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
    Call mOrigin.AssociatorsAsync_(objWbemSink, strAssocClass, strResultClass, strResultRole, strRole, bClassesOnly, bSchemaOnly, strRequiredAssocQualifier, strRequiredQualifier, iFlags, objWbemNamedValueSet)
End Sub

'******************************************************************************
'* [概  要] Clone_ メソッド
'* [詳  細] Clone_ のラッパーメソッド。
'*          Clone this Object.
'*
'* @return WMIExSWbemObjectSet
'*
'******************************************************************************
Function Clone_() As WMIExSWbemObject
    Dim wrap As WMIExSWbemObjectSet: Set wrap = New WMIExSWbemObjectSet
    Set wrap.Origin = mOrigin.Clone_()
    Set Clone_ = wrap
End Function

'******************************************************************************
'* [概  要] CompareTo_ メソッド
'* [詳  細] CompareTo_ のラッパーメソッド。
'*          Compare this Object with another.
'*
'* @param objWbemObject xx
'* @param iFlags xx
'* @return Boolean
'*
'******************************************************************************
Function CompareTo_(objWbemObject As Object, Optional iFlags As Long) As Boolean
    CompareTo_ = mOrigin.CompareTo_(objWbemObject, iFlags)
End Function

'******************************************************************************
'* [概  要] Delete_ メソッド
'* [詳  細] Delete_ のラッパーメソッド。
'*          Delete this Object.
'*
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'*
'******************************************************************************
Sub Delete_(Optional iFlags As Long, Optional objWbemNamedValueSet As Object)
    Call mOrigin.Delete_(iFlags, objWbemNamedValueSet)
End Sub

'******************************************************************************
'* [概  要] DeleteAsync_ メソッド
'* [詳  細] DeleteAsync_ のラッパーメソッド。
'*          Delete this Object asynchronously.
'*
'* @param objWbemSink xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub DeleteAsync_(objWbemSink As Object, Optional iFlags As Long, Optional objWbemNamedValueSet As Object, Optional objWbemAsyncContext As Object)
    Call mOrigin.DeleteAsync_(objWbemSink, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub

'******************************************************************************
'* [概  要] ExecMethod_ メソッド
'* [詳  細] ExecMethod_ のラッパーメソッド。
'*          Execute a Method of this Object.
'*
'* @param strMethodName xx
'* @param objWbemInParameters xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemObject
'*
'******************************************************************************
Function ExecMethod_(strMethodName As String, Optional objWbemInParameters As Object, Optional iFlags As Long, Optional objWbemNamedValueSet As Object) As WMIExSWbemObject
    Dim wrap As WMIExSWbemObject: Set wrap = New WMIExSWbemObject
    Set wrap.Origin = mOrigin.ExecMethod_(strMethodName, objWbemInParameters, iFlags, objWbemNamedValueSet)
    Set ExecMethod_ = wrap
End Function

'******************************************************************************
'* [概  要] ExecMethodAsync_ メソッド
'* [詳  細] ExecMethodAsync_ のラッパーメソッド。
'*          Execute a Method of this Object asynchronously.
'*
'* @param objWbemSink xx
'* @param strMethodName xx
'* @param objWbemInParameters xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub ExecMethodAsync_(objWbemSink As Object, _
                     strMethodName As String, _
                     Optional objWbemInParameters As Object, _
                     Optional iFlags As Long, _
                     Optional objWbemNamedValueSet As Object, _
                     Optional objWbemAsyncContext As Object)
    Call mOrigin.ExecMethodAsync_(objWbemSink, strMethodName, objWbemInParameters, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub

'******************************************************************************
'* [概  要] GetObjectText_ メソッド
'* [詳  細] GetObjectText_ のラッパーメソッド。
'*          Get the MOF text of this Object.
'*
'* @param iFlags xx
'* @return String
'*
'******************************************************************************
Function GetObjectText_(Optional iFlags As Long) As String
    GetObjectText_ = mOrigin.GetObjectText_(iFlags)
End Function

'******************************************************************************
'* [概  要] GetText_ メソッド
'* [詳  細] GetText_ のラッパーメソッド。
'*          Retrieve a textual representation of this Object.
'*
'* @param iObjectTextFormat xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return String
'*
'******************************************************************************
Function GetText_(iObjectTextFormat As WbemObjectTextFormatEnum, Optional iFlags As Long, Optional objWbemNamedValueSet As Object) As String
    GetText_ = mOrigin.GetText_(iObjectTextFormat, iFlags, objWbemNamedValueSet)
End Function
    
'******************************************************************************
'* [概  要] Instances_ メソッド
'* [詳  細] Instances_ のラッパーメソッド。
'*          Return all instances of this Class.
'*
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemObjectSet
'*
'******************************************************************************
Function Instances_(Optional iFlags As Long = 16, Optional objWbemNamedValueSet As Object) As WMIExSWbemObjectSet
    Dim wrap As WMIExSWbemObjectSet: Set wrap = New WMIExSWbemObjectSet
    Set wrap.Origin = mOrigin.Instances_(iFlags, objWbemNamedValueSet)
    Set Instances_ = wrap
End Function

'******************************************************************************
'* [概  要] InstancesAsync_ メソッド
'* [詳  細] InstancesAsync_ のラッパーメソッド。
'*          Return all instances of this Class asynchronously.
'*
'* @param objWbemSink xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub InstancesAsync_(objWbemSink As Object, Optional iFlags As Long, Optional objWbemNamedValueSet As Object, Optional objWbemAsyncContext As Object)
    Call mOrigin.InstancesAsync_(objWbemSink, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub

'******************************************************************************
'* [概  要] Put_ メソッド
'* [詳  細] Put_ のラッパーメソッド。
'*          Save this Object.
'*
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemObjectPath
'*
'******************************************************************************
Function Put_(Optional iFlags As Long, Optional objWbemNamedValueSet As Object) As WMIExSWbemObjectPath
    Dim wrap As WMIExSWbemObjectPath: Set wrap = New WMIExSWbemObjectPath
    Set wrap.Origin = mOrigin.Put_(iFlags, objWbemNamedValueSet)
    Set Put_ = wrap
End Function

'******************************************************************************
'* [概  要] PutAsync_ メソッド
'* [詳  細] PutAsync_ のラッパーメソッド。
'*          Save this Object asynchronously.
'*
'* @param objWbemSink xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub PutAsync_(objWbemSink As Object, Optional iFlags As Long, Optional objWbemNamedValueSet As Object, Optional objWbemAsyncContext As Object)
    Call mOrigin.PutAsync_(objWbemSink, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub

'******************************************************************************
'* [概  要] References_ メソッド
'* [詳  細] References_ のラッパーメソッド。
'*          Get the References to this Object.
'*
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
Function References_(Optional strResultClass As String, _
                     Optional strRole As String, _
                     Optional bClassesOnly As Boolean = False, _
                     Optional bSchemaOnly As Boolean = False, _
                     Optional strRequiredQualifier As String, _
                     Optional iFlags As Long = 16, _
                     Optional objWbemNamedValueSet As Object) As WMIExSWbemObjectSet
    Dim wrap As WMIExSWbemObjectSet: Set wrap = New WMIExSWbemObjectSet
    Set wrap.Origin = mOrigin.References_(strResultClass, strRole, bClassesOnly, bSchemaOnly, strRequiredQualifier, iFlags, objWbemNamedValueSet)
    Set References_ = wrap
End Function

'******************************************************************************
'* [概  要] ReferencesAsync_ メソッド
'* [詳  細] ReferencesAsync_ のラッパーメソッド。
'*          Get the References to this Object asynchronously.
'*
'* @param objWbemSink xx
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
Sub ReferencesAsync_(objWbemSink As Object, _
                     Optional strResultClass As String, _
                     Optional strRole As String, _
                     Optional bClassesOnly As Boolean = False, _
                     Optional bSchemaOnly As Boolean = False, _
                     Optional strRequiredQualifier As String, _
                     Optional iFlags As Long, _
                     Optional objWbemNamedValueSet As Object, _
                     Optional objWbemAsyncContext As Object)
    Call mOrigin.ReferencesAsync_(objWbemSink, strResultClass, strRole, bClassesOnly, bSchemaOnly, strRequiredQualifier, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub

'******************************************************************************
'* [概  要] Refresh_ メソッド
'* [詳  細] Refresh_ のラッパーメソッド。
'*          Refresh this Object.
'*
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'*
'******************************************************************************
Sub Refresh_(Optional iFlags As Long, Optional objWbemNamedValueSet As Object)
    Call mOrigin.Refresh_(iFlags, objWbemNamedValueSet)
End Sub

'******************************************************************************
'* [概  要] SetFromText_ メソッド
'* [詳  細] SetFromText_ のラッパーメソッド。
'*          Set this Object using the supplied textual representation.
'*
'* @param bsText xx
'* @param iObjectTextFormat xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'*
'******************************************************************************
Sub SetFromText_(bsText As String, iObjectTextFormat As WbemObjectTextFormatEnum, Optional iFlags As Long, Optional objWbemNamedValueSet As Object)
    Call mOrigin.SetFromText_(bsText, iObjectTextFormat, iFlags, objWbemNamedValueSet)
End Sub

    
'******************************************************************************
'* [概  要] SpawnDerivedClass_ メソッド
'* [詳  細] SpawnDerivedClass_ のラッパーメソッド。
'*          Create a subclass of this Object.
'*
'* @param iFlags xx
'* @return WMIExSWbemObject
'*
'******************************************************************************
Function SpawnDerivedClass_(Optional iFlags As Long) As WMIExSWbemObject
    Dim wrap As WMIExSWbemObject: Set wrap = New WMIExSWbemObject
    Set wrap.Origin = mOrigin.SpawnDerivedClass_(iFlags)
    Set SpawnDerivedClass_ = wrap
End Function

'******************************************************************************
'* [概  要] SpawnInstance_ メソッド
'* [詳  細] SpawnInstance_ のラッパーメソッド。
'*          Create an Instance of this Object.
'*
'* @param iFlags xx
'* @return WMIExSWbemObject
'*
'******************************************************************************
Function SpawnInstance_(Optional iFlags As Long) As WMIExSWbemObject
    Dim wrap As WMIExSWbemObject: Set wrap = New WMIExSWbemObject
    Set wrap.Origin = mOrigin.SpawnInstance_(iFlags)
    Set SpawnInstance_ = wrap
End Function

'******************************************************************************
'* [概  要] Subclasses_ メソッド
'* [詳  細] Subclasses_ のラッパーメソッド。
'*          Enumerate subclasses of this Class.
'*
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemObjectSet
'*
'******************************************************************************
Function Subclasses_(Optional iFlags As Long = 16, Optional objWbemNamedValueSet As Object) As WMIExSWbemObjectSet
    Dim wrap As WMIExSWbemObjectSet: Set wrap = New WMIExSWbemObjectSet
    Set wrap.Origin = mOrigin.Subclasses_(iFlags, objWbemNamedValueSet)
    Set Subclasses_ = wrap
End Function

'******************************************************************************
'* [概  要] SubclassesAsync_ メソッド
'* [詳  細] SubclassesAsync_ のラッパーメソッド。
'*          Enumerate subclasses of this Class asynchronously.
'*
'* @param objWbemSink xx
'* @param iFlags xx
'* @param objWbemNamedValueSet xx
'* @param objWbemAsyncContext xx
'*
'******************************************************************************
Sub SubclassesAsync_(objWbemSink As Object, Optional iFlags As Long, Optional objWbemNamedValueSet As Object, Optional objWbemAsyncContext As Object)
    Call mOrigin.SubclassesAsync_(objWbemSink, iFlags, objWbemNamedValueSet, objWbemAsyncContext)
End Sub

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------
