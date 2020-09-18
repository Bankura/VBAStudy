Attribute VB_Name = "ZZM_MyTest"
Option Explicit

Function FizzBuzz(ByVal n As Long) As String
    Select Case BitFlag(n Mod 5 = 0, n Mod 3 = 0)
        Case 0: FizzBuzz = CStr(n)
        Case 1: FizzBuzz = "Fizz"
        Case 2: FizzBuzz = "Buzz"
        Case 3: FizzBuzz = "FizzBuzz"
        Case Else: Err.Raise 51 'UNREACHABLE
    End Select
End Function

''' EntryPoint
Sub Main()
    Debug.Print Join(ArrMap(Init(New Func, vbString, AddressOf FizzBuzz), ArrRange(1&, 100&)))
End Sub

Sub TestWorksheetEx001()
    Dim Wsh As WorkSheetEx
    Set Wsh = New WorkSheetEx
    Set Wsh.Origin = ThisWorkbook.ActiveSheet
    
    Dim sx As StringEx
    Set sx = Init(New StringEx, "aaa")
    Debug.Print sx


End Sub

Sub TestValidateUtils001()
    Dim dic As Object
    Set dic = Core.CreateDictionary()
   
    
    Dim wshObj As Object
    Set wshObj = Core.Wsh
    Dim wmiObj As Object
    Set wmiObj = Core.wmi
    Dim regObj As Object
    Set regObj = Core.CreateRegExp("[a-Z]")
    Set regObj = Base.GetRegExp
    Dim srpObj As Object
    Set srpObj = Core.CreateStdRegProv
    
    Debug.Print TypeName(dic)
    Debug.Print TypeName(wshObj)
    Debug.Print TypeName(wmiObj)
    Debug.Print TypeName(regObj)
    Debug.Print TypeName(srpObj)
    
    'Debug.Print ValidateUtils.
    
    
End Sub

Sub TestStringUtils001_AppendIfMissing()
    Debug.Print StringUtils.AppendIfMissing("", "xyz") = "xyz"
    Debug.Print StringUtils.AppendIfMissing("abc", "xyz") = "abcxyz"
    Debug.Print StringUtils.AppendIfMissing("abcxyz", "xyz") = "abcxyz"
    Debug.Print StringUtils.AppendIfMissing("abcXYZ", "xyz") = "abcXYZxyz"
    Debug.Print StringUtils.AppendIfMissing("abc", "xyz", "") = "abc"
    Debug.Print StringUtils.AppendIfMissing("abc", "xyz", "mno") = "abcxyz"
    Debug.Print StringUtils.AppendIfMissing("abcxyz", "xyz", "mno") = "abcxyz"
    Debug.Print StringUtils.AppendIfMissing("abcmno", "xyz", "mno") = "abcmno"
    Debug.Print StringUtils.AppendIfMissing("abcXYZ", "xyz", "mno") = "abcXYZxyz"
    Debug.Print StringUtils.AppendIfMissing("abcMNO", "xyz", "mno") = "abcMNOxyz"
End Sub
Sub TestStringUtils002_AppendIfMissingIgnoreCase()

    Debug.Print StringUtils.AppendIfMissingIgnoreCase("", "xyz") = "xyz"
    Debug.Print StringUtils.AppendIfMissingIgnoreCase("abc", "xyz") = "abcxyz"
    Debug.Print StringUtils.AppendIfMissingIgnoreCase("abcxyz", "xyz") = "abcxyz"
    Debug.Print StringUtils.AppendIfMissingIgnoreCase("abcXYZ", "xyz") = "abcXYZ"
    Debug.Print StringUtils.AppendIfMissingIgnoreCase("abc", "xyz", "") = "abc"
    Debug.Print StringUtils.AppendIfMissingIgnoreCase("abc", "xyz", "mno") = "abcxyz"
    Debug.Print StringUtils.AppendIfMissingIgnoreCase("abcxyz", "xyz", "mno") = "abcxyz"
    Debug.Print StringUtils.AppendIfMissingIgnoreCase("abcmno", "xyz", "mno") = "abcmno"
    Debug.Print StringUtils.AppendIfMissingIgnoreCase("abcXYZ", "xyz", "mno") = "abcXYZ"
    Debug.Print StringUtils.AppendIfMissingIgnoreCase("abcMNO", "xyz", "mno") = "abcMNO"
End Sub

Sub TestStringUtils003_Capitalize()
    Debug.Print StringUtils.Capitalize("") = ""
    Debug.Print StringUtils.Capitalize("cat") = "Cat"
    Debug.Print StringUtils.Capitalize("cAt") = "CAt"
    Debug.Print StringUtils.Capitalize("'cat'") = "'cat'"
End Sub

Sub TestStringUtils004_ToCharArray()
    Dim v: v = StringUtils.ToCharArray("abcde")
    Debug.Print Join(v, ",")
    v = StringUtils.ToCharArrayReverse("abcde")
    Debug.Print Join(v, ",")

End Sub


Sub TestHoge()
'    Call PrintCaptionAndProcessMain
'
'    Debug.Print "Hei"
'    Base.GetWinAPI.Sleep 1000
'    Debug.Print "Hei"
'    Base.GetWinAPI.Sleep 1000
'    Debug.Print "Hoo"
    
    Debug.Print GetExcelBookProc
    Dim i As Long
    For i = 0 To UBound(WinApiFunctions.wD)
        Debug.Print WinApiFunctions.wD(i).wkb.name
    Next
    
End Sub

Sub dumpVariant()
    Dim v
    Debug.Print "VarPtr : " & VarPtr(v)
  
    Debug.Print "文字列"
    v = "2いB": Debug.Print "Let : " & v
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemory(VarPtr(v), 8, "dec") 'データ型
    Call DebugUtils.DumpMemory(VarPtr(v) + 8, 8, "dec") 'StrPtrのアドレスがここ
    Call DebugUtils.DumpMemory(StrPtr(v), LenB(v), "str") '実際の文字列
    Call DebugUtils.DumpMemory(StrPtr(v) + LenB(v), 2) '終端2バイト
    Call DebugUtils.DumpMemoryFromVariant(v)

    Debug.Print "整数"
    v = CInt(12345): Debug.Print "Let : " & v
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemory(VarPtr(v), 8, "dec") 'データ型
    Call DebugUtils.DumpMemory(VarPtr(v) + 8, 2, "dec") '実際の数値
    Call DebugUtils.DumpMemory(StrPtr(v), LenB(v), "str") '数値が文字列で
    Call DebugUtils.DumpMemoryFromVariant(v)
    
    Debug.Print "長整数"
    v = CLng(12345678): Debug.Print "Let : " & v
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemory(VarPtr(v), 8, "dec") 'データ型
    Call DebugUtils.DumpMemory(VarPtr(v) + 8, 4, "dec") '実際の数値
    Call DebugUtils.DumpMemory(StrPtr(v), LenB(v), "str") '数値が文字列で
    Call DebugUtils.DumpMemoryFromVariant(v)
  
    Debug.Print "単精度小数点"
    v = CSng(1234.5): Debug.Print "Let : " & v
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemory(VarPtr(v), 8, "dec") 'データ型
    Call DebugUtils.DumpMemory(VarPtr(v) + 8, 4, "bin") '実際の数値
    Call DebugUtils.DumpMemory(StrPtr(v), LenB(v), "str") '数値が文字列で
    Call DebugUtils.DumpMemoryFromVariant(v)

    Debug.Print "倍精度小数点"
    v = CDbl(1234.5678): Debug.Print "Let : " & v
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemory(VarPtr(v), 8, "dec") 'データ型
    Call DebugUtils.DumpMemory(VarPtr(v) + 8, 8, "bin") '実際の数値
    Call DebugUtils.DumpMemory(StrPtr(v), LenB(v), "str") '数値が文字列で
    Call DebugUtils.DumpMemoryFromVariant(v)
    
      
    Debug.Print "Byte"
    v = CByte(2): Debug.Print "Let : " & v
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemory(VarPtr(v), 8, "dec") 'データ型
    Call DebugUtils.DumpMemory(VarPtr(v) + 8, 1, "bin") '実際の数値
    Call DebugUtils.DumpMemory(StrPtr(v), LenB(v), "str") '数値が文字列で
    Call DebugUtils.DumpMemoryFromVariant(v)
    
    Debug.Print "Boolean"
    v = False: Debug.Print "Let : " & v
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemory(VarPtr(v), 8, "dec") 'データ型
    Call DebugUtils.DumpMemory(VarPtr(v) + 8, 1, "bin") '実際の数値
    Call DebugUtils.DumpMemory(StrPtr(v), LenB(v), "str") '数値が文字列で
    Call DebugUtils.DumpMemoryFromVariant(v)
    
    Debug.Print "Date"
    v = Now: Debug.Print "Let : " & v
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemory(VarPtr(v), 8, "dec") 'データ型
    Call DebugUtils.DumpMemory(VarPtr(v) + 8, 4, "dec") '実際の数値
    Call DebugUtils.DumpMemory(StrPtr(v), LenB(v), "str") '数値が文字列で
    Call DebugUtils.DumpMemoryFromVariant(v)
    
    Debug.Print "Currency"
    v = CCur(3331234.5678): Debug.Print "Let : " & v
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemory(VarPtr(v), 8, "dec") 'データ型
    Call DebugUtils.DumpMemory(VarPtr(v) + 8, 8, "bin") '実際の数値
    Call DebugUtils.DumpMemory(StrPtr(v), LenB(v), "str") '数値が文字列で
    Call DebugUtils.DumpMemoryFromVariant(v)
    
    Debug.Print "LongLong"
    v = 33312345678^: Debug.Print "Let : " & v
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemory(VarPtr(v), 8, "dec") 'データ型
    Call DebugUtils.DumpMemory(VarPtr(v) + 8, 8, "bin") '実際の数値
    Call DebugUtils.DumpMemory(StrPtr(v), LenB(v), "str") '数値が文字列で
    Call DebugUtils.DumpMemoryFromVariant(v)
    
    Debug.Print "コレクション"
    Dim v3 As Variant
    Debug.Print "VarPtr : " & VarPtr(v3)
    Set v3 = New Collection: Debug.Print "Set : " & TypeName(v3)
    Debug.Print "ObjPtr : " & ObjPtr(v3)
    Call DebugUtils.DumpMemory(VarPtr(v3), 8, "dec") 'ObjPtrのアドレス
    Call DebugUtils.DumpMemory(ObjPtr(v3), 8, "dec") 'オブジェクトのアドレス
    Call DebugUtils.DumpMemoryFromVariant(v3)
End Sub

Sub dumpObject()
    Debug.Print "ワークシート"
    Dim v1 As Worksheet
    Debug.Print "VarPtr : " & VarPtr(v1)
    Debug.Print "ObjPtr : " & ObjPtr(v1)
    Call DebugUtils.DumpMemory(VarPtr(v1), 8, "dec") 'ObjPtrのアドレス
    Call DebugUtils.DumpMemory(ObjPtr(v1), 8, "dec") 'オブジェクトのアドレス
    Set v1 = Worksheets(1): Debug.Print "Set : " & TypeName(v1)
    Debug.Print "ObjPtr : " & ObjPtr(v1)
    Call DebugUtils.DumpMemory(VarPtr(v1), 8, "dec") 'ObjPtrのアドレス
    Call DebugUtils.DumpMemory(ObjPtr(v1), 8, "dec") 'オブジェクトのアドレス
    Call DebugUtils.DumpMemoryFromObject(v1)
  
'    Debug.Print "クラス"
'    Dim v2 As Class1
'    Debug.Print "VarPtr : " & VarPtr(v2)
'    Debug.Print "ObjPtr : " & ObjPtr(v2)
'    Set v2 = New Class1: Debug.Print "Set : " & TypeName(v2)
'    Debug.Print "ObjPtr : " & ObjPtr(v2)
'    Call DumpMemory(VarPtr(v2), 8, "dec") 'ObjPtrのアドレス
'    Call DumpMemory(ObjPtr(v2), 8, "dec") 'オブジェクトのアドレス
  
    Debug.Print "コレクション"
    Dim v3 As Collection
    Debug.Print "VarPtr : " & VarPtr(v3)
    Debug.Print "ObjPtr : " & ObjPtr(v3)
    Set v3 = New Collection: Debug.Print "Set : " & TypeName(v3)
    Debug.Print "ObjPtr : " & ObjPtr(v3)
    Call DebugUtils.DumpMemory(VarPtr(v3), 8, "dec") 'ObjPtrのアドレス
    Call DebugUtils.DumpMemory(ObjPtr(v3), 8, "dec") 'オブジェクトのアドレス
    Call DebugUtils.DumpMemoryFromObject(v3)
  
    Debug.Print "ファイルシステムオブジェクト"
    Dim v4 As Object
    Debug.Print "VarPtr : " & VarPtr(v4)
    Debug.Print "ObjPtr : " & ObjPtr(v4)
    Set v4 = CreateObject("Scripting.FileSystemObject"): Debug.Print "Set : " & TypeName(v4)
    Debug.Print "ObjPtr : " & ObjPtr(v4)
    Call DebugUtils.DumpMemory(VarPtr(v4), 8, "dec") 'ObjPtrのアドレス
    Call DebugUtils.DumpMemory(ObjPtr(v4), 8, "dec") 'オブジェクトのアドレス
    Call DebugUtils.DumpMemoryFromObject(v4)
End Sub

Sub dumpArray()
    Dim v(2) As Long
    v(0) = 1: v(1) = 2: v(2) = 3

    Debug.Print "VarPtr(0) : " & VarPtr(v(0))
    Debug.Print "VarPtr(1) : " & VarPtr(v(1))
    Debug.Print "VarPtr(2) : " & VarPtr(v(2))

    Call DebugUtils.DumpMemory(VarPtr(v(0)), LenB(v(0)), "dec")
    Call DebugUtils.DumpMemory(VarPtr(v(1)), LenB(v(1)), "dec")
    Call DebugUtils.DumpMemory(VarPtr(v(2)), LenB(v(2)), "dec")
    
    Dim vv: vv = v
    Debug.Print "VarPtr(vv) : " & VarPtr(vv)
    Debug.Print "VarPtr(v0) : " & VarPtr(vv(0))
    Debug.Print "VarPtr(v1) : " & VarPtr(vv(1))
    Debug.Print "VarPtr(v2) : " & VarPtr(vv(2))
    Call DebugUtils.DumpMemory(VarPtr(vv), 8, "dec")
    Call DebugUtils.DumpMemory(VarPtr(vv(0)), LenB(vv(0)), "dec")
    Call DebugUtils.DumpMemory(VarPtr(vv(1)), LenB(vv(1)), "dec")
    Call DebugUtils.DumpMemory(VarPtr(vv(2)), LenB(vv(2)), "dec")
End Sub

Sub dumpString()
    Dim v As String
    Debug.Print "VarPtr : " & VarPtr(v)
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemory(VarPtr(v), LONGPTR_SIZE, "dec")
    v = "1あA"
    Debug.Print "Let : " & v
    Debug.Print "VarPtr : " & VarPtr(v)
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemory(VarPtr(v), LONGPTR_SIZE, "dec")
    Call DebugUtils.DumpMemory(StrPtr(v) - 4, 4, "dec") '文字列長は-4位置
    Call DebugUtils.DumpMemory(StrPtr(v), LenB(v), "str") '実際の文字列
    Call DebugUtils.DumpMemory(StrPtr(v) + LenB(v), 2) '終端2バイト
    v = "1あA"
    Debug.Print "Let : " & v
    Debug.Print "VarPtr : " & VarPtr(v)
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemory(VarPtr(v), LONGPTR_SIZE, "dec")
    Call DebugUtils.DumpMemory(StrPtr(v) - 4, 4, "dec") '文字列長は-4位置
    Call DebugUtils.DumpMemory(StrPtr(v), LenB(v), "str") '実際の文字列
    Call DebugUtils.DumpMemory(StrPtr(v) + LenB(v), 2) '終端2バイト
    
    Debug.Print "■■■■■■■■"
    Debug.Print "VarPtr : " & VarPtr(v)
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemoryFromVariant(v)
    v = "1あA"
    Debug.Print "Let : " & v
    Debug.Print "VarPtr : " & VarPtr(v)
    Debug.Print "StrPtr : " & StrPtr(v)
    Call DebugUtils.DumpMemoryFromVariant(v)
    Call DebugUtils.DumpMemory(StrPtr(v) - 4, 4, "dec") '文字列長は-4位置
    Call DebugUtils.DumpMemoryFromString(v) '実際の文字列
    Call DebugUtils.DumpMemory(StrPtr(v) + LenB(v), 2) '終端2バイト
    
    Debug.Print "コレクション"
    Dim v3 As Collection
    Debug.Print "VarPtr : " & VarPtr(v3)
    Debug.Print "ObjPtr : " & ObjPtr(v3)
    Set v3 = New Collection: Debug.Print "Set : " & TypeName(v3)
    Debug.Print "ObjPtr : " & ObjPtr(v3)
    Call DebugUtils.DumpMemory(VarPtr(v3), 8, "dec") 'ObjPtrのアドレス
    Call DebugUtils.DumpMemory(ObjPtr(v3), 8, "dec") 'オブジェクトのアドレス
    Call DebugUtils.DumpMemoryFromVariant(v3) 'ObjPtrのアドレス
    Call DebugUtils.DumpMemoryFromObject(v3)
    
End Sub

Sub dumpnanka()
    Debug.Print "コレクション"
    Dim v3 As Collection
    Debug.Print "VarPtr : " & VarPtr(v3)
    Debug.Print "ObjPtr : " & ObjPtr(v3)
    Set v3 = New Collection: Debug.Print "Set : " & TypeName(v3)
    v3.Add "aaa", "AAAAVV"
    v3.Add "bbb", "B"
    v3.Add "ccc", "C"
    v3.Add "ddd", "D"
    v3.Add "e", "E"
    
    Debug.Print "ObjPtr : " & ObjPtr(v3)
    
    Dim a As LongPtr, i As Long
    Dim b As LongPtr
    Dim c As LongPtr
    Dim d As LongPtr
    
    Debug.Print "-----------"
    For i = 0 To 16
        Call DebugUtils.DumpMemory(ObjPtr(v3) + 8 * i, 8, "dec") 'オブジェクトのアドレス
    Next
    

    a = DebugUtils.DumpMemory(ObjPtr(v3) + 40, 8, "dec")
'    Debug.Print a
'    For i = 0 To 4
'        Call DebugUtils.DumpMemory(a + 8 * i, 8, "dec") 'オブジェクトのアドレス
'    Next
    Debug.Print "-----------"
    Debug.Print "■First Value"
    b = DebugUtils.DumpMemory(a + 8 * 1, 8, "dec")
    Call DebugUtils.DumpMemory(b, 8, "dec")
    Call DebugUtils.DumpMemory(b, 8, "str")
   
    
    Debug.Print "-----------"
    Debug.Print "■First Key"
    b = DebugUtils.DumpMemory(a + 8 * 3, 8, "dec")
    Call DebugUtils.DumpMemory(b, 8, "dec")
    Call DebugUtils.DumpMemory(b, 8, "str")
    Debug.Print "-----------"
    
    Debug.Print a
    For i = 0 To 6
        Call DebugUtils.DumpMemory(a + 8 * i, 8, "dec") 'オブジェクトのアドレス
    Next
    Debug.Print "==========="
    c = DebugUtils.DumpMemory(a + 40, 8, "dec")
    For i = 0 To 6
        Call DebugUtils.DumpMemory(c + 8 * i, 8, "dec") 'オブジェクトのアドレス
    Next
    
    Debug.Print "-----------"
    Debug.Print "■Snd Value"
    d = DebugUtils.DumpMemory(c + 8 * 1, 8, "dec")
    Call DebugUtils.DumpMemory(d, 8, "dec")
    Call DebugUtils.DumpMemory(d, 8, "str")
    
    Debug.Print "-----------"
    Debug.Print "■Snd Key"
    d = DebugUtils.DumpMemory(c + 8 * 3, 8, "dec")
    Call DebugUtils.DumpMemory(d, 8, "dec")
    Call DebugUtils.DumpMemory(d, 8, "str")
    
    
    Debug.Print "■Last Value"
    a = DebugUtils.DumpMemory(ObjPtr(v3) + 48, 8, "dec")
'    Debug.Print a
'    For i = 0 To 4
'        Call DebugUtils.DumpMemory(a + 8 * i, 8, "dec") 'オブジェクトのアドレス
'    Next
    Debug.Print "-----------"
    b = DebugUtils.DumpMemory(a + 8 * 1, 8, "dec")
    Call DebugUtils.DumpMemory(b, 8, "dec")
    Call DebugUtils.DumpMemory(b, 8, "str")

    Debug.Print "-----------"
    Debug.Print "■Last Key"
    b = DebugUtils.DumpMemory(a + 8 * 3, 8, "dec")
    Call DebugUtils.DumpMemory(b, 8, "dec")
    Call DebugUtils.DumpMemory(b, 8, "str")
    Debug.Print "-----------"

    Debug.Print a
    For i = 0 To 6
        Call DebugUtils.DumpMemory(a + 8 * i, 8, "dec") 'オブジェクトのアドレス
    Next
    Debug.Print "==========="
    c = DebugUtils.DumpMemory(a + 48, 8, "dec")
    For i = 0 To 6
        Call DebugUtils.DumpMemory(c + 8 * i, 8, "dec") 'オブジェクトのアドレス
    Next

    Debug.Print "-----------"
    Debug.Print "■Fourth Value"
    d = DebugUtils.DumpMemory(c + 8 * 1, 8, "dec")
    Call DebugUtils.DumpMemory(d, 8, "dec")
    Call DebugUtils.DumpMemory(d, 8, "str")
    
End Sub

Sub dumpnanka2()
    Debug.Print "コレクション"
    Dim v3 As Collection
    Debug.Print "VarPtr : " & VarPtr(v3)
    Debug.Print "ObjPtr : " & ObjPtr(v3)
    Set v3 = New Collection: Debug.Print "Set : " & TypeName(v3)
    v3.Add "aaa", "AAAAVV"
    v3.Add "bbb", "B"
    v3.Add "ccc", "C"
    v3.Add "ddd", "D"
    v3.Add "e", "E"
    
    Debug.Print "ObjPtr : " & ObjPtr(v3)
    
    Dim x As LongPtr
    Dim y As LongPtr
    Dim a As LongPtr
    Dim b As LongPtr
    Dim c As LongPtr
    Dim d As LongPtr
    Dim e As LongPtr
    Dim f As LongPtr
    Dim g As LongPtr
    Dim h As LongPtr
    Dim i As LongPtr
    Dim j As LongPtr
    Dim k As LongPtr
    Dim l As LongPtr
    Dim m As LongPtr
    Dim n As LongPtr
    Dim o As LongPtr
    Dim p As LongPtr
    Dim q As LongPtr
    Dim r As LongPtr
    Dim s As LongPtr
    Dim t As LongPtr
    
    '★
    x = ObjPtr(v3) + 40
    a = BinaryUtils.ByteArray2LongPtr(BinaryUtils.CopyMemory2ByteArray(x, LONGPTR_SIZE))
    b = a + 8

    Debug.Print "■First Value"
    c = DebugUtils.DumpMemory(b, 8, "dec")
    Call DebugUtils.DumpMemory(c, 8, "str")

    d = a + 24
    
    Debug.Print "■First Key"
    e = DebugUtils.DumpMemory(d, 8, "dec")
    Call DebugUtils.DumpMemory(e, 8, "str")

    Debug.Print "==========="

    '★
    f = a + 40
    g = BinaryUtils.ByteArray2LongPtr(BinaryUtils.CopyMemory2ByteArray(f, LONGPTR_SIZE))
    h = g + 8

    Debug.Print "■Snd Value"
    i = DebugUtils.DumpMemory(h, 8, "dec")
    Call DebugUtils.DumpMemory(i, 8, "str")
    
    j = g + 24
        
    Debug.Print "-----------"
    Debug.Print "■Snd Key"
    k = DebugUtils.DumpMemory(j, 8, "dec")
    Call DebugUtils.DumpMemory(k, 8, "str")
    
    
    '★
    y = ObjPtr(v3) + 48
    l = BinaryUtils.ByteArray2LongPtr(BinaryUtils.CopyMemory2ByteArray(y, LONGPTR_SIZE))
    m = l + 8
    
    Debug.Print "■Last Value"
    n = DebugUtils.DumpMemory(m, 8, "dec")
    Call DebugUtils.DumpMemory(n, 8, "str")

    o = l + 24
    
    Debug.Print "■Last Key"
    p = DebugUtils.DumpMemory(o, 8, "dec")
    Call DebugUtils.DumpMemory(p, 8, "str")
    Debug.Print "-----------"


    Debug.Print "==========="
    '★
    q = l + 48
    r = BinaryUtils.ByteArray2LongPtr(BinaryUtils.CopyMemory2ByteArray(q, LONGPTR_SIZE))
    s = r + 8
    
    Debug.Print "■Fourth Value"
    t = DebugUtils.DumpMemory(s * 1, 8, "dec")
    Call DebugUtils.DumpMemory(t, 8, "str")
    
    Debug.Print "==========="
    Dim str As String
    Dim strp As LongPtr
    
    Debug.Print VarPtr(str)
    Debug.Print StrPtr(str)
    strp = StrPtr(str)
    Debug.Print "str: " + str
    Call GetWinAPI.CopyMemory(VarPtr(str), d, LONGPTR_SIZE)
    Debug.Print "str: " + str
    Debug.Print VarPtr(str)
    Debug.Print StrPtr(str)
    Call GetWinAPI.CopyMemoryByRef(VarPtr(str), strp, LONGPTR_SIZE, True, False)
    Debug.Print "str: " + str
    Debug.Print VarPtr(str)
    Debug.Print StrPtr(str)

End Sub

Sub ColUtils_Test()
    Dim col As New Collection

    col.Add "aaa", "A"
    col.Add "bbb", "B"
    col.Add "ccc", "C"
    col.Add "ddd", "D"
    col.Add "e", "E"

    Debug.Print col.Count
    Dim v
    For Each v In col
        Debug.Print v
    Next
    Debug.Print CollectionUtils.GetCollectionKeyByIndex(4, col)
    
    Debug.Print CollectionUtils.GetCollectionIndexByKey("E", col)
    
End Sub

Public Sub Sample()
  Dim s1 As String, s2 As String

  s1 = "中国語テスト：" & vbNewLine & _
       ChrW(&H94F6) & ChrW(&H884C) & ChrW(&H6682) & ChrW(&H505C) & _
       ChrW(&H65B0) & ChrW(&H589E) & ChrW(&H4F4F) & ChrW(&H623F) & _
       ChrW(&H8D37) & ChrW(&H6B3E)
       
  Dim cb As New ClipBoardUtils
  cb.SetClipBoard s1
  s2 = cb.GetClipBoard
  CreateObject("WScript.Shell").Popup s2
End Sub

'暗号化
Sub CryptoUtils_EncryptStringTripleDES_Test()

    CryptoUtils.InitializationVector = "12345678"
    CryptoUtils.TripleDesKey = "bankurarakusitai"
    Debug.Print "password -> " & CryptoUtils.EncryptStringTripleDES("password")
    Debug.Print "password -> " & CryptoUtils.EncryptStringTripleDES("password", HexString)
End Sub

'復号化
Sub CryptoUtils_DecryptStringTripleDES_Test()
    CryptoUtils.InitializationVector = "12345678"
    CryptoUtils.TripleDesKey = "bankurarakusitai"
    
    Debug.Print "0p3WWy330MxJnyyD1xzMlQ== -> " & CryptoUtils.DecryptStringTripleDES("0p3WWy330MxJnyyD1xzMlQ==")
    Debug.Print "D29DD65B2DF7D0CC499F2C83D71CCC95 -> " & CryptoUtils.DecryptStringTripleDES("D29DD65B2DF7D0CC499F2C83D71CCC95", HexString)
    
End Sub

'暗号化
Sub CryptoUtils_EncryptStringSHA256_Test()
    Debug.Print "password -> " & CryptoUtils.EncryptStringSHA256("password")
    Debug.Print "password -> " & CryptoUtils.EncryptStringSHA256("password", HexString)
End Sub

Sub Commander_Test001()
    Dim cmdr As New Commander
End Sub

Sub DosCommander_Test001()
    Dim cmdr As New DosCommander
    Dim v
'    For Each v In cmdr.CmdDir("C:\dev\vba", "/R", "/A")
'        Debug.Print v
'    Next
    For Each v In cmdr.GetFilePathsRecursive("C:\dev\vba")
        Debug.Print v
    Next
End Sub

Sub PowershellCommander_Test001()
    Dim cmdr As New PowerShellCommander
    Dim v
    For Each v In cmdr.GetCommandResultAsArray("Get-ChildItem | Select-Object CreationTime,LastWriteTime,LastAccessTime,Name")
        Debug.Print v
    Next

End Sub

Sub RowEnumerator_Test001()
    Dim myTest As New ZZC_MyTest
    myTest.Main

End Sub

Sub DebugUtils_Test001()
    'DebugUtils.PrintVariantArray Array(1, 2, 3)
    DebugUtils.PrintVariantArray ArrayUtils.Copy2DArray(ArrayUtils.Create2DArrayWithValue(Array(1, 2, 3), Array(2, 2, 3), Array(3, 2, 3)), 1, 1, 1, 1)
End Sub

Sub CollectionUtils_Test001()
    Dim col As New Collection
    col.Add "1"
    col.Add 2
    col.Add "3"
    col.Add "4"
    col.Add Array("a", "b", "c")
    Debug.Print CollectionUtils.CollectionToString(col)

End Sub

Sub CollectionUtils_Test002()
'    Dim dic As New DictionaryEx
    Dim dic As Object: Set dic = Core.CreateDictionary()
    dic.Add "k1", "1"
    dic.Add "k2", 2
    dic.Add "k3", "3"
    dic.Add "k4", "4"
    dic.Add "k5", Array("a", "b", "c")
    Debug.Print CollectionUtils.DictionaryToString(dic)

End Sub

Sub ArrayUtils_Test001()
    Dim col As New Collection
    col.Add "1"
    col.Add 2
    col.Add "3"
    col.Add "4"
    col.Add Array("a", "b", "c")
    
    Dim dic As Object: Set dic = Core.CreateDictionary()
    dic.Add "k1", "1"
    dic.Add "k2", 2
    dic.Add "k3", "3"
    dic.Add "k4", "4"
    dic.Add "k5", Array("a", "b", "c")
    
    Dim arr
    arr = Array("ara", "yada", "cyo", 2020, col, dic, Core.Wsh)
    
    Debug.Print ArrayUtils.ToString(arr)
End Sub


Sub ObjectCreate_Test001()
    Dim obj As Object
    'Set obj = CreateObject("AMOVIE.ActiveMovieControl.2")
    'Set obj = CreateObject("new:{05589FA1-C356-11CE-BF01-00AA0055595A}")
    
    
    Debug.Print TypeName(obj)

    
End Sub

Sub OleDBProviderExists_Test001()

    Dim vArr, v
    Dim stdRegProv As Object: Set stdRegProv = CreateStdRegProv()
    ' 32bit
    stdRegProv.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\Classes", vArr
    
    For Each v In ArrayUtils.Search(vArr, "PostgreSQL")
        Debug.Print v
    Next

    ' 64bit
    stdRegProv.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Classes", vArr
    
    For Each v In ArrayUtils.Search(vArr, "PostgreSQL")
        Debug.Print v
    Next
End Sub

Sub OdbcDriverList_Test001()

    Dim vArr, v
    Dim stdRegProv As Object: Set stdRegProv = CreateStdRegProv()
    ' 32bit
'    stdRegProv.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI", vArr
'
'    For Each v In ArrayUtils.Search(vArr, "Ora")
'        Debug.Print v
'    Next

    ' 64bit
    stdRegProv.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBCINST.INI", vArr
    
    For Each v In ArrayUtils.Search(vArr, "Ora")
        Debug.Print v
    Next
    
End Sub

Sub AccessDBConnect_Test001()
    Call DatabaseUtils.Disconnect
    'Call DatabaseUtils.OpenAccess("C:\develop\mydb_pass.accdb", "pass")
    Call DatabaseUtils.OpenAccess("C:\develop\mydb_pass.accdb", "pass", "odbc")
    Dim strSQL As String: strSQL = "select * from 消費税率マスタ"
    
    strSQL = "select 適用開始日 from 消費税率マスタ"
    Dim vArr, v
    
    vArr = DatabaseUtils.SelectList(strSQL)
    
    For Each v In vArr
        Debug.Print v
    Next
    Call DatabaseUtils.Disconnect
End Sub

Sub SQLServerDBConnect_Test001()
    Call DatabaseUtils.Disconnect
    ' Call DatabaseUtils.OpenSqlServer("localhost\SQLExpress", "TestDB2", 1433, "sa", "xxxx")
    Call DatabaseUtils.OpenSqlServer("localhost\SQLExpress", "TestDB2", 1433, "sa", "xxxx", "oledb")
    Dim strSQL As String: strSQL = "select * from LocationMst"
    
    strSQL = "select LocationCd, LocationName from LocationMst"
    Dim vArr, v
    
    vArr = DatabaseUtils.SelectList(strSQL)
    
    For Each v In vArr
        Debug.Print v
    Next
    Call DatabaseUtils.Disconnect
End Sub

Sub SQLServerLocalDBConnect_Test001()
    Call DatabaseUtils.Disconnect
    Call DatabaseUtils.OpenSqlServerLocalDB("C:\Users\bankura\TestDB.mdf", "TestDB")
    Dim strSQL As String: strSQL = "select * from table1"
    
    strSQL = "select id, name from table1"
    Dim vArr, v
    
    vArr = DatabaseUtils.SelectList(strSQL)
    
    For Each v In vArr
        Debug.Print v
    Next
    Call DatabaseUtils.Disconnect
End Sub

Sub SQLiteConnect_Test001()
    Call DatabaseUtils.Disconnect
    Call DatabaseUtils.OpenSQLite("C:\SQLite\test.db")
    Dim strSQL As String: strSQL = "select * from table1"
    
    Dim vArr, v
    
    vArr = DatabaseUtils.SelectList(strSQL)
    
    For Each v In vArr
        Debug.Print v
    Next
    Call DatabaseUtils.Disconnect
End Sub

Sub MySqlConnect_Test001()
    Call DatabaseUtils.Disconnect
    Call DatabaseUtils.OpenMySql("localhost", 3306, "sakila", "root", "bankura")
    Dim strSQL As String: strSQL = "select * from country where country like 'United%'"

    Dim vArr, v
    
    vArr = DatabaseUtils.SelectList(strSQL)
    
    For Each v In vArr
        Debug.Print v
    Next
    Call DatabaseUtils.Disconnect
End Sub


Sub PostgreDBConnect_Test001()
    Call DatabaseUtils.Disconnect
    Call DatabaseUtils.OpenPostgreSql("localhost", 5433, "ban", "ban", "ban", "odbc")
    Dim strSQL As String: strSQL = "select * from table1"
    
    strSQL = "select name from table1"
    Dim vArr, v
    
    vArr = DatabaseUtils.SelectList(strSQL)
    
    For Each v In vArr
        Debug.Print v
    Next
    Call DatabaseUtils.Disconnect
End Sub

Sub OracleDBConnect_Test001()
    Call DatabaseUtils.Disconnect
    Call DatabaseUtils.OpenOracle("localhost", 1521, "XEPDB1", "bankura", "bankura", "oledb")
    Dim strSQL As String: strSQL = "select * from table1"

    Dim vArr, v
    
    vArr = DatabaseUtils.SelectList(strSQL)
    
    For Each v In vArr
        Debug.Print v
    Next
    Call DatabaseUtils.Disconnect
End Sub


Sub ExcelDBConnect_Test001()
    Call DatabaseUtils.Disconnect

    Call DatabaseUtils.OpenExcel("C:\develop\myxldb.xlsb", True, 2, "oledb")
    'Call DatabaseUtils.OpenExcel("C:\develop\myxldb.xlsb", , , "odbc")
    Dim strSQL As String: strSQL = "select * from [table1$]"
    
    'strSQL = "select name from [table1$]"
    Dim vArr, v
    
    vArr = DatabaseUtils.SelectList(strSQL)
    
    For Each v In vArr
        Debug.Print v
    Next
    Call DatabaseUtils.Disconnect
End Sub

Sub CsvDBConnect_Test001()
    Call DatabaseUtils.Disconnect

    Call DatabaseUtils.OpenCsv("C:\develop\", False, "oledb")
    'Call DatabaseUtils.OpenCsv("C:\develop\", , "odbc")
    Dim strSQL As String: strSQL = "select * from [myxldb.csv]"
    
    'strSQL = "select name from [myxldb.csv]"
    Dim vArr, v
    
    vArr = DatabaseUtils.SelectList(strSQL)
    
    For Each v In vArr
        Debug.Print v
    Next
    Call DatabaseUtils.Disconnect
End Sub

Sub TextDBConnect_Test001()
    Call DatabaseUtils.Disconnect

    Call DatabaseUtils.OpenText("C:\develop\", True, "oledb")
    'Call DatabaseUtils.OpenText("C:\develop\", , "odbc")
    Dim strSQL As String: strSQL = "select * from [myxldb.csv]"
    

    Dim vArr, v
    
    vArr = DatabaseUtils.SelectList(strSQL)
    
    For Each v In vArr
        Debug.Print v
    Next
    Call DatabaseUtils.Disconnect
End Sub
