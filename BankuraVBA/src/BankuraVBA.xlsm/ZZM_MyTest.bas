Attribute VB_Name = "ZZM_MyTest"
Option Explicit

' Use Ariawase -------------------------------------------------------------------------------------
''' EntryPoint
Sub FizzBuzzMain()
    Debug.Print Join(ArrMap(Init(New Func, vbString, AddressOf FizzBuzz), ArrRange(1&, 100&)))
End Sub
Public Function FizzBuzz(ByVal n As Long) As String
    Select Case BitFlag(n Mod 5 = 0, n Mod 3 = 0)
        Case 0: FizzBuzz = CStr(n)
        Case 1: FizzBuzz = "Fizz"
        Case 2: FizzBuzz = "Buzz"
        Case 3: FizzBuzz = "FizzBuzz"
        Case Else: Err.Raise 51 'UNREACHABLE
    End Select
End Function
' --------------------------------------------------------------------------------------------------

' Test Array2DEx -----------------------------------------------------------------------------------

Sub TestArray2DEx001_Init()
    Dim arr2d
    arr2d = ArrayUtils.Create2DArrayWithValue(Array(1, 2, 3), Array(4, 5, 6), Array(7, 8, 9))

    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx, arr2d)
    Debug.Print "2次元配列から初期化"
    DebugUtils.PrintArray arr2dex.To2DArray
    
    Dim arr2dex2 As Array2DEx
    Set arr2dex2 = Core.Init(New Array2DEx, arr2dex)
    Debug.Print "Array2DExから初期化"
    DebugUtils.PrintArray arr2dex.To2DArray

    Dim arr2dex3 As Array2DEx
    Set arr2dex3 = Core.Init(New Array2DEx, Array(1, 2, 3))
    Debug.Print "1次元配列から初期化"
    DebugUtils.PrintArray arr2dex3.To2DArray
    
    Dim arr2dex4 As Array2DEx
    Set arr2dex4 = Core.Init(New Array2DEx, arr2dex2.ToArrayExOfArrayEx)
    Debug.Print "ArrayExOfArrayExから初期化"
    DebugUtils.PrintArray arr2dex4.To2DArray
    
    Dim arr2dex5 As Array2DEx
    Set arr2dex5 = Core.Init(New Array2DEx, Core.Init(New ArrayEx, Array(1, 2, 3)))
    Debug.Print "ArrayExから初期化"
    DebugUtils.PrintArray arr2dex5.To2DArray
    
    Dim arr2dex6 As Array2DEx
    Set arr2dex6 = Core.Init(New Array2DEx, "test")
    Debug.Print "文字列から初期化"
    DebugUtils.PrintArray arr2dex6.To2DArray
End Sub

Sub TestArray2DEx002_Add()
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx)
    arr2dex.DefaultInitValue = 0
    arr2dex.AddColumn Array(1, 2, 3)
    arr2dex.AddColumn Array(1, 2, 3)
    arr2dex.AddColumn Array(1, 2, 3)
    arr2dex.AddColumn Array(1, 2, 3)
    arr2dex.AddRow Array(12, 2, 3, 4, 5)
    
    Dim arr2d
    arr2d = ArrayUtils.Create2DArrayWithValue(Array(1, 2, 3), Array(4, 5, 6), Array(7, 8, 9))
    Dim arr2dex2 As Array2DEx
    Set arr2dex2 = Core.Init(New Array2DEx, arr2d)
    arr2dex.AddRows arr2dex2
    arr2dex(1, 2) = 5
    arr2dex.AddColumns arr2d
    arr2dex.Sort 0, False
    DebugUtils.PrintArray arr2dex.To2DArray
End Sub

Sub TestArray2DEx003_Add()
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx)
    arr2dex.DefaultInitValue = 0
    arr2dex.AddColumns Array(1, 2, 3, 4, 5), Array(1, 2, 3, 4, 5), ArrayUtils.Create2DArrayWithValue(Array(1, 2, 3), Array(4, 5, 6), Array(7, 8, 9))
    DebugUtils.PrintArray arr2dex.To2DArray
End Sub

Sub TestArray2DEx004_Expand()
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx)
    arr2dex.DefaultInitValue = 0
    arr2dex.Expand 5, 5
    DebugUtils.PrintArray arr2dex.To2DArray
End Sub

Sub TestArray2DEx005()
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx).Range(10, 5, 5)
    DebugUtils.PrintArray arr2dex.To2DArray
    Debug.Print arr2dex.Contains(9)
    Debug.Print arr2dex.Contains(10)
    Debug.Print arr2dex.Contains(11)
    Debug.Print arr2dex.Contains(45)
    Debug.Print arr2dex.Contains(46)
    Debug.Print arr2dex.RowContains(0, 9)
    Debug.Print arr2dex.RowContains(0, 16)
    Debug.Print arr2dex.ColContains(0, 10)
    Debug.Print arr2dex.ColContains(0, 9)
    
    Dim arr2dex2 As Array2DEx
    Set arr2dex2 = Core.Init(New Array2DEx).Range(10, 5, 5)
    Debug.Print arr2dex.Equals(arr2dex2)
    arr2dex(0, 0) = 1
    Debug.Print arr2dex.Equals(arr2dex2)
    Debug.Print arr2dex.ColIndexOf(1, 35)
    Debug.Print arr2dex.RowIndexOf(4, 35)
    
    Dim xy As Array2DIndex
    xy = arr2dex.IndexOf(35)
    Debug.Print xy.x, xy.y
    
    Debug.Print arr2dex.EntireRowIndexOf(ArrayUtils.Range(34, 39))
End Sub

Sub TestArray2DEx006()
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx)
    arr2dex.DefaultInitValue = 0
    arr2dex.AddRows Array(1, 2, 3), Array(1, 2, 3), Array(1, 2, 3), Array(1, 4, 6)
    DebugUtils.PrintArray arr2dex.Uniq.To2DArray
    
    Dim arr2dex2 As Array2DEx, arr2dex3 As Array2DEx
    Set arr2dex2 = Core.Init(New Array2DEx).Range(10, 2, 2)
    Set arr2dex3 = arr2dex.Concat(arr2dex2)
    Debug.Print "Concat"
    DebugUtils.PrintArray arr2dex3.To2DArray
    
    Set arr2dex3 = arr2dex.RowSlice(3, 3)
    Debug.Print "RowSlice"
    DebugUtils.PrintArray arr2dex3.To2DArray
    
    Set arr2dex3 = arr2dex.ColSlice(2, 2)
    Debug.Print "ColSlice"
    DebugUtils.PrintArray arr2dex3.To2DArray
    
    Debug.Print "Range 0"
    DebugUtils.PrintArray Core.Init(New Array2DEx).Range(10, 0, 0).To2DArray
    
    Debug.Print "Map"
    Dim fun As Func
    Set fun = Core.Init(New Func, vbLong, AddressOf TestFuctionSquere)
    DebugUtils.PrintArray arr2dex.Map(fun).To2DArray
    
    Debug.Print "Zip"
    Set arr2dex3 = arr2dex3.Range(1, 3, 2)
    Set fun = Core.Init(New Func, vbLong, AddressOf TestFuctionSumUp)
    DebugUtils.PrintArray arr2dex.Zip(fun, arr2dex3.To2DArray).To2DArray
    
    Debug.Print "RowFilter"
    Set fun = Core.Init(New Func, vbBoolean, AddressOf TestFuctionMyFilter)
    DebugUtils.PrintArray arr2dex.RowFilter(fun).To2DArray
    
    Debug.Print "ColFilter"
    DebugUtils.PrintArray arr2dex.ColFilter(fun).To2DArray

    Debug.Print "RowFold"
    Set fun = Core.Init(New Func, vbLong, AddressOf TestFuctionMyFold)
    DebugUtils.PrintArray arr2dex.RowFold(fun).ToArray
    
    Debug.Print "ColFold"
    DebugUtils.PrintArray arr2dex.ColFold(fun).ToArray

    Debug.Print "RowScan"
    DebugUtils.PrintArray arr2dex.RowScan(fun).To2DArray
    
    Debug.Print "ColScan"
    DebugUtils.PrintArray arr2dex.ColScan(fun).To2DArray
End Sub
Public Function TestFuctionSquere(ByVal Source As Long) As Long
    TestFuctionSquere = Source * Source
End Function
Public Function TestFuctionSumUp(ByVal source1 As Long, ByVal source2 As Long) As Long
    TestFuctionSumUp = source1 + source2
End Function
Public Function TestFuctionMyFilter(ByVal Source) As Boolean
    TestFuctionMyFilter = ArrayUtils.Contains(Source, 2)
End Function
Public Function TestFuctionMyFold(ByVal basedata As Long, ByVal elem As Long) As Long
    TestFuctionMyFold = basedata + elem
End Function

Sub TestArray2DEx007()
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx)
    arr2dex.DefaultInitValue = 0
    arr2dex.AddRows Array(1, 2, 3), Array(1, 2, 3), Array(1, 2, 3), Array(1, 4, 6)

    Call arr2dex.RowInsert(1, Array(2, 2, 3))
    Call arr2dex.ColInsert(1, Array(9, 8, 7, 6, 5))
    Call arr2dex.RowsInsert(2, Array(3, 2, 3, 2), Array(5, 5, 5, 5), Array(5, 5, 5, 5))
    Call arr2dex.ColsInsert(2, ArrayUtils.Create2DArrayWithValue(Array(1, 2), Array(3, 4), Array(5, 6), Array(7, 8), Array(1, 2), Array(3, 4), Array(5, 6), Array(7, 8)))
    Call arr2dex.RowRemove(0, 5)
    Call arr2dex.ColRemove(2, 5)
    Call arr2dex.RowRemoveAt(1)
    Call arr2dex.ColRemoveAt(1)
    DebugUtils.Show arr2dex.RowLastIndexOf(1, 2)
    DebugUtils.Show arr2dex.ColLastIndexOf(2, 2)
    Dim p As Array2DIndex: p = arr2dex.LastIndexOf(2)
    DebugUtils.Show p.x & " " & p.y
    DebugUtils.PrintArray arr2dex.To2DArray
    
    Debug.Print "Join"
    Debug.Print arr2dex.Join(",")
    
    Call arr2dex.RowRemoveRange(2, 3)
    Call arr2dex.ColRemoveRange(1, 2)
    DebugUtils.Show "RemoveRange"
    Debug.Print arr2dex.Join(",")
End Sub

Sub TestArray2DEx008()
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx)
    arr2dex.DefaultInitValue = "none"
    arr2dex.AddRows Array("山田太郎", "鈴木史郎", "竹田伸二"), Array("舞山記理子", "三原静子", "前山裕次郎"), Array("徳田綾子", "霧島司", "松尾三太夫"), Array("モリソン", "ローゼマイン", "フェルディナンド")
    DebugUtils.PrintArray arr2dex.To2DArray
    
    DebugUtils.PrintArray arr2dex.RowSearch(1, "郎").ToArray
    DebugUtils.PrintArray arr2dex.ColSearch(0, "子").ToArray
    DebugUtils.PrintArray arr2dex.Search("郎").ToArray
    
    DebugUtils.Show "[正規表現]"
    DebugUtils.PrintArray arr2dex.RowRegexSearch(0, "^.*田.*$").ToArray
    DebugUtils.PrintArray arr2dex.ColRegexSearch(0, ".*子$").ToArray
    DebugUtils.PrintArray arr2dex.RegexSearch(".*子$").ToArray
End Sub

Sub TestArray2DEx009()
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx)
    arr2dex.DefaultInitValue = 0
    arr2dex.AddRows Array(1, 2, 3), Array(1, 2, 3), Array(1, 2, 3), Array(1, 4, 6)

    DebugUtils.Show arr2dex
End Sub


' Test ArrayEx -----------------------------------------------------------------------------------

Sub TestArrayEx001_Init()

End Sub

' Test ValidateUtils -------------------------------------------------------------------------------

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
    For i = 0 To UBound(WinApiFunctions.wd)
        Debug.Print WinApiFunctions.wd(i).wkb.Name
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
    Dim C As LongPtr
    Dim d As LongPtr
    
    Debug.Print "-----------"
    For i = 0 To 64
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
    C = DebugUtils.DumpMemory(a + 40, 8, "dec")
    For i = 0 To 6
        Call DebugUtils.DumpMemory(C + 8 * i, 8, "dec") 'オブジェクトのアドレス
    Next
    
    Debug.Print "-----------"
    Debug.Print "■Snd Value"
    d = DebugUtils.DumpMemory(C + 8 * 1, 8, "dec")
    Call DebugUtils.DumpMemory(d, 8, "dec")
    Call DebugUtils.DumpMemory(d, 8, "str")
    
    Debug.Print "-----------"
    Debug.Print "■Snd Key"
    d = DebugUtils.DumpMemory(C + 8 * 3, 8, "dec")
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
    C = DebugUtils.DumpMemory(a + 48, 8, "dec")
    For i = 0 To 6
        Call DebugUtils.DumpMemory(C + 8 * i, 8, "dec") 'オブジェクトのアドレス
    Next

    Debug.Print "-----------"
    Debug.Print "■Fourth Value"
    d = DebugUtils.DumpMemory(C + 8 * 1, 8, "dec")
    Call DebugUtils.DumpMemory(d, 8, "dec")
    Call DebugUtils.DumpMemory(d, 8, "str")
    
End Sub


Sub ColUtils_Test()
    Dim col As New Collection

    col.Add "aaa", "A"
    col.Add "bbb", "B"
    col.Add "ccc", "C"
    col.Add "ddd", "D"
    col.Add "e", "E"

    Debug.Print col.count
    Dim v
    For Each v In col
        Debug.Print v
    Next
    Debug.Print CollectionUtils.GetCollectionKeyByIndex(3, col)
    
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
    For Each v In cmdr.GetFilePathsRecursive("C:\develop\data")
        Debug.Print v
    Next
End Sub

Sub PowershellCommander_Test001()
    Dim cmdr As New PowerShellCommander
    Dim v
    For Each v In cmdr.Exec("Get-ChildItem | Select-Object CreationTime,LastWriteTime,LastAccessTime,Name", False)
        Debug.Print v
    Next

End Sub

Sub RowEnumerator_Test001()
    Dim myTest As New ZZC_MyTest
    myTest.Main

End Sub

Sub DebugUtils_Test001()
    'DebugUtils.PrintVariantArray Array(1, 2, 3)
    DebugUtils.PrintArray ArrayUtils.Copy2DArray(ArrayUtils.Create2DArrayWithValue(Array(1, 2, 3), Array(2, 2, 3), Array(3, 2, 3)), 1, 1, 1, 1)
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
    
    Dim Arr
    Arr = Array("ara", "yada", "cyo", 2020, col, dic, Core.Wsh)
    
    Debug.Print ArrayUtils.ToString(Arr)
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


Sub TextCollectionSort_Test001()
    Dim col As Collection
    Set col = New Collection

    Call col.Add(Array(1, "Kimetu", 200), "aaa")
    Call col.Add(Array(2, "ゴブリン", 300), "bbb")
    Call col.Add(Array(3, "スライム", 10), "ccc")
    Call col.Add(Array(4, "キングだむ", 37499900), "ddd")
    Call col.Add(Array(5, "フリー連", 9910), "eee")
    Call col.Add(Array(6, "おおい涼真", 320), "fff")

    Dim v
    For Each v In CollectionUtils.GetCollectionKeys(col)
        DebugUtils.Show v
    Next

    Dim fnc As Func
    Set fnc = Core.Init(New Func, vbBoolean, AddressOf CompareTesTesQuick)
    
    'Call CollectionUtils.CollectionSort(col, fnc, True)
    Call CollectionUtils.CollectionSort(col, "CompareTesTesQuick", True)

    For Each v In col
        DebugUtils.Show v
    Next
    
    For Each v In CollectionUtils.GetCollectionKeys(col)
        DebugUtils.Show v
    Next

End Sub


Function CompareTesTes(val1, val2) As Boolean
    CompareTesTes = val1(2) > val2(2)
End Function

Function CompareTesTesQuick(val1, val2, ByVal flg As Boolean) As Boolean
'    If flg Then
'        CompareTesTesQuick = val1(1) > val2(1)
'    Else
'        CompareTesTesQuick = val1(1) < val2(1)
'    End If
    
    If flg Then
        If val1(1) = val2(1) Then
            CompareTesTesQuick = val1(2) < val2(2)
            Exit Function
        End If
        CompareTesTesQuick = val1(1) < val2(1)
    Else
        If val1(1) = val2(1) Then
            CompareTesTesQuick = val1(2) > val2(2)
            Exit Function
        End If
        CompareTesTesQuick = val1(1) > val2(1)
    End If
End Function


Sub Array2DSortTest001()
    
    Dim wsex As WorkSheetEx: Set wsex = Core.Init(New WorkSheetEx, "Sheet1")
    'Dim var As Variant: var = wsex.ExportArray(1, 1, , True)
    Dim arr2dex As Array2DEx
    Set arr2dex = wsex.ExportArray2DEx(1, 3)

    'Dim fnc As Func
    'Set fnc = Core.Init(New Func, vbBoolean, AddressOf CompareTesTesQuick)
    
    'Call ArrayUtils.Array2DSort(var, fnc)
    
    'Call ArrayUtils.Array2DSort(var, Array("0 asc", "1 desc", "2"))
    'DebugUtils.PrintArray var
    
    'DebugUtils.PrintArray arr2dex.SortBy(fnc)
    'arr2dex.SortBy fnc
    
    DebugUtils.PrintArray arr2dex.SortBy(Array("0 desc", "1 desc", "2 desc")).RowSlice(0, 4)

End Sub

Sub MesureExcuteTime_Test()
    DebugUtils.MesureExcuteTime Core.Init(New Func, VbVarType.vbVariant, AddressOf PythonCommander_Test002)
End Sub

Sub CodeModuleUtils_Test001()
    Debug.Print ArrayUtils.GetLength(Split(VBCodeModuleUtils.GetComponentAllCodes("ClipBoardUtils"), vbNewLine))

    Debug.Print "行数：" & VBCodeModuleUtils.CountOfLines("ClipBoardUtils") & _
                " 空行：" & VBCodeModuleUtils.CountOfEmptyLines("ClipBoardUtils") & _
                " コメント行：" & VBCodeModuleUtils.CountOfCommentLines("ClipBoardUtils") & _
                " 実行数：" & VBCodeModuleUtils.CountOfLogicalLines("ClipBoardUtils")
                
                
    DebugUtils.Show VBCodeModuleUtils.GetProcNames("Item")
    
    DebugUtils.Show VBCodeModuleUtils.CountOfProc("Item")
    
    'DebugUtils.PrintVariantArray VBCodeModuleUtils.GetComponentsInfo().SortBy(Array("1 asc", "0 asc"))
    Dim arr2d As Variant: arr2d = VBCodeModuleUtils.GetComponentsInfo().SortBy(Array("1 asc", "0 asc")).To2DArray
    
    DebugUtils.PrintArray arr2d
End Sub


Sub CodeModuleUtils_Test002()
    DebugUtils.WashImmediateWindow
    DebugUtils.Show VBCodeModuleUtils.GetProcKindName("Item", "Name")
    DebugUtils.Show VBCodeModuleUtils.GetProcStartLine("Item", "Name")
    DebugUtils.Show VBCodeModuleUtils.GetProcCountLines("Item", "Name")
    DebugUtils.Show VBCodeModuleUtils.GetProcNameOfLine("Item", 92)
    
End Sub

Sub CodeModuleUtils_Test003()
    VBCodeModuleUtils.ClearImmediateWindow
End Sub


Sub PythonCommander_Test001()

    Dim cmder As PyCommander
    Set cmder = New PyCommander
    Debug.Print cmder.ExecCommand("print('aaa')")
    
    cmder.ScriptMode
    'Debug.Print cmder.ExecScript("C:\develop\python\csvread_pandas_test.py")
    cmder.ExecScript ("C:\develop\python\createForm.py")
End Sub

Sub PythonCommander_Test002()

    Dim cmder As PyCommander
    Set cmder = New PyCommander
    
    Dim arrex As ArrayEx
    Set arrex = New ArrayEx
    arrex.Add "import pandas as pd"
    arrex.Add "csv_input = pd.read_csv(filepath_or_buffer='C:/develop/data/csv/test_stock.csv', encoding='utf_8', sep=',')"
    arrex.Add "print(csv_input[['コード', '銘柄名']]) "
    
    Debug.Print cmder.WriteScriptAndRun(arrex.ToArray)
    
End Sub



Sub powershellTest001()
    Dim cmder As PowerShellCommander
    Set cmder = New PowerShellCommander
    Dim v
    'For Each v In cmder.Exec("gsv -c localhost|?{$_.status -like 'r*'}", False)
    '    Debug.Print v
    'Next
    DebugUtils.PrintArray cmder.Exec("gsv -c localhost|?{$_.status -like 'r*'}", False)
End Sub

Sub powershellTest002()

    Dim cmder As PowerShellCommander
    Set cmder = New PowerShellCommander
    
    Dim arrex As ArrayEx
    Set arrex = New ArrayEx
    arrex.Add "cd C:\develop\powershell"
    arrex.Add "$cnt = (Get-ChildItem -Recurse).count"
    arrex.Add "$allsize = 0"
    arrex.Add "Get-ChildItem -Recurse | ForEach-Object { $allsize += $_.Length }"
    arrex.Add "$average = $allsize / $cnt"
    arrex.Add "Write-Output ""ファイルサイズの平均は ${average} byteです。"""
    Debug.Print cmder.WriteScriptAndRun(arrex.ToArray)
    
End Sub

Sub powershellTest003()

    Dim cmder As PowerShellCommander
    Set cmder = New PowerShellCommander
    
    Dim dic As DictionaryEx
    Set dic = cmder.PSVersionTable

    DebugUtils.Show dic

    Debug.Print dic("PSVersion")("Major") & "." & dic("PSVersion")("Minor") & "." & dic("PSVersion")("Build") & "." & dic("PSVersion")("Revision")

    Debug.Print cmder.GetExecutionPolicy
    
    
End Sub


Sub PsqlCommander_Test001()

    Dim cmder As PsqlCommander
    Set cmder = New PsqlCommander
    cmder.DbHost = "localhost"
    cmder.DbPort = 5433
    cmder.dbName = "ban"
    cmder.DbUserName = "ban"
    cmder.DbPassword = "ban"
    'cmder.TuplesOnly = True

    Dim strSQL As String: strSQL = "select id as ""ユーザID"", name as ""名前"" from table1"
    Dim vArr
    
    vArr = cmder.Exec(strSQL)
    
    DebugUtils.PrintArray vArr
End Sub

Sub DosCommander_Test901()
    Dim cmder As DosCommander
    Set cmder = New DosCommander

    Debug.Print cmder.Exec("C:\develop\powershell\fuga.bat")
End Sub

Sub DosCommander_Test902()

    Dim cmder As DosCommander
    Set cmder = New DosCommander
    
    Dim arrex As ArrayEx
    Set arrex = New ArrayEx
    arrex.Add "@echo off"
    arrex.Add "echo ふがあ"

    Debug.Print cmder.WriteBatchAndRun(arrex.ToArray)
End Sub

Sub foooooooo()
    Dim mysheet As New WorkSheetEx
    Call mysheet.Init("Sheet1")
    
    Dim v
    Set v = mysheet.GetRowToArrayEx(1)
    DebugUtils.Show v
End Sub


Sub ApplyFunc2ExcelFiles_Test001()
    Dim fnc As Func
    Set fnc = Core.Init(New Func, vbBoolean, AddressOf BookA1Print)
    'Call XlBookUtils.ApplyProc2Books("C:\develop\data\xls", fnc, True)
    Call XlBookUtils.ApplyProc2Books("C:\develop\data\xls", "BookA1Print", True)
End Sub

Sub BookA1Print(ByVal bookObj As Workbook)
    Debug.Print bookObj.Worksheets(1).Cells(1, 1).Value
    bookObj.Worksheets(1).Cells(1, 1).Value = bookObj.Worksheets(1).Cells(1, 1).Value & "X"
    bookObj.Save

End Sub

Sub ForEach_Test001()
    Dim fnc As Func
    Set fnc = Core.Init(New Func, vbBoolean, AddressOf PrintNameTes)
    Call ForEach(ThisWorkbook.Worksheets, fnc)

End Sub
Sub PrintNameTes(ByVal obj As Object)
    Debug.Print obj.Name
End Sub

Sub munuuu()
    Dim v
    
'    For Each v In SystemUtils.ExecWmiQuery("Select * from Win32_BIOS")
'        Debug.Print "Build Number         : " & v.Properties_("BuildNumber")
'        Debug.Print "Current Language     : " & v.Properties_("CurrentLanguage")
'        Debug.Print "Installable Languages: " & v.Properties_("InstallableLanguages")
'        Debug.Print "Manufacturer         : " & v.Properties_("Manufacturer")
'        Debug.Print "Name                 : " & v.Properties_("Name")
'        Debug.Print "Primary BIOS         : " & v.Properties_("PrimaryBIOS")
'        Debug.Print "Serial Number        : " & v.Properties_("SerialNumber")
'        Debug.Print "SMBIOS Version       : " & v.Properties_("SMBIOSBIOSVersion")
'        Debug.Print "SMBIOS Major Version : " & v.Properties_("SMBIOSMajorVersion")
'        Debug.Print "SMBIOS Minor Version : " & v.Properties_("SMBIOSMinorVersion")
'        Debug.Print "SMBIOS Present       : " & v.Properties_("SMBIOSPresent")
'        Debug.Print "Status               : " & v.Properties_("Status")
'    Next
    For Each v In SystemUtils.ExecWmiQuery("Select * From Win32_NTLogEvent Where Logfile='System' And TimeGenerated > '2021/11/01'" & _
                                           " And (Eventcode = '6005' Or Eventcode = '6006' Or Eventcode = '7001' Or Eventcode = '7002')")
        Debug.Print Format(Mid(v.TimeGenerated, 1, 14), "####/##/## ##:##:##"), v.EventCode, v.Message
    Next
End Sub

Public Sub UXUtils_IconTest001()
    Call UXUtils.AddIcon(Application.hwnd, "アイコンてすと") 'システムトレイにアイコンを登録
End Sub

Public Sub UXUtils_IconTest002()
    Call UXUtils.ShowBalloon("バルーンメッセージてすと", "バルーンタイトル", 1, 10)       'バルーンチップの表示
End Sub

Public Sub UXUtils_IconTest003()
    Call UXUtils.ModifyIcon("C:\windows\system32\notepad.exe") 'アイコンの変更
End Sub

Public Sub UXUtils_IconTest004()
    Call UXUtils.DeleteIcon 'アイコンの削除
End Sub
Public Sub UXUtils_IconTest005()
    Call UXUtils.NotifyToast("Excelからの通知", "私がやってきた！", , 30)
End Sub

Public Sub WorkSheetEx_Test001()
    Dim ws As WorkSheetEx: Set ws = Core.Init(New WorkSheetEx, "Sheet1")
    Call ws.SetAutoFilter(1, 3, 1, 6).FilterOn(1, "4")
End Sub
Public Sub WorkSheetEx_Test002()
    Dim ws As WorkSheetEx: Set ws = Core.Init(New WorkSheetEx, "Sheet1")
    Call ws.FilterOff
End Sub
Public Sub WorkSheetEx_Test003()
    Dim ws As WorkSheetEx: Set ws = Core.Init(New WorkSheetEx, "Sheet1")
    Call ws.RemoveAutoFilter
End Sub

Public Sub BaseParallel_Test001()
    Call Base.ParallelExec("testFunc", 3, "C:\develop\tmp\test0.txt", "C:\develop\tmp\test1.txt", "C:\develop\tmp\test2.txt")
End Sub

Public Sub testFunc(n As Long, ParamArray paParams())
    Call FileUtils.CreateNullCharFile(CStr(paParams(n)), 500000000)
    Application.DisplayAlerts = False
    ThisWorkbook.Close False
End Sub

Public Sub OnTimeForClass_Test001()
    Dim tesObj As ZZC_MyTest
    Set tesObj = New ZZC_MyTest
    
    Call Base.OnTimeForClass(1, tesObj, "DebugPrintLong", 1)
    Call Base.OnTimeForClass(2, tesObj, "DebugPrintLong", 2)
    Call Base.OnTimeForClass(3, tesObj, "DebugPrintLong", 3)
    Call Base.OnTimeForClass(4, tesObj, "DebugPrintLong", 4)
    Call Base.OnTimeForClass(5, tesObj, "DebugPrintLong", 5)
     
End Sub
Public Sub OnTimeForClass_Test002()
    Dim tesObj As ZZC_MyTest
    Set tesObj = New ZZC_MyTest
    
    Dim k1 As String: k1 = Base.OnTimeForClass(1, tesObj, "DebugPrintLongR", 1)
    Dim k2 As String: k2 = Base.OnTimeForClass(2, tesObj, "DebugPrintLongR", 2)
    Dim k3 As String: k3 = Base.OnTimeForClass(3, tesObj, "DebugPrintLongR", 3)
    Dim k4 As String: k4 = Base.OnTimeForClass(4, tesObj, "DebugPrintLongR", 4)
    Dim k5 As String: k5 = Base.OnTimeForClass(5, tesObj, "DebugPrintLongR", 5)

    Debug.Print "ResultKey1: " & k1
    Debug.Print "ResultKey2: " & k2
    Debug.Print "ResultKey3: " & k3
    Debug.Print "ResultKey4: " & k4
    Debug.Print "ResultKey5: " & k5
End Sub
Public Sub OnTimeForClass_Test003()

    Call Base.OnTimeForClass(1, DebugUtils, "Show", 111)
End Sub
