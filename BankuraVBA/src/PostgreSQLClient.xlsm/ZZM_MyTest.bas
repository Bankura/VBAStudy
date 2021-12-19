Attribute VB_Name = "ZZM_MyTest"
Option Explicit

Sub PsqlCommander_Test001()

    Dim cmder As PsqlCommander
    Set cmder = New PsqlCommander
    cmder.DbHost = "localhost"
    cmder.DbPort = 5433
    cmder.dbName = "ban"
    cmder.DbUserName = "ban"
    cmder.DbPassword = "ban"
    'cmder.TuplesOnly = True

    Dim strSQL As String: strSQL = "select id as ""ÉÜÅ[ÉUID"", name as ""ñºëO"" from table1"
    Dim vArr
    
    vArr = cmder.Exec(strSQL)
    
    DebugUtils.PrintArray vArr
End Sub
