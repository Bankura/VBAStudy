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
    DebugUtils.PrintVariantArray arr2dex.To2DArray
    
    Dim arr2dex2 As Array2DEx
    Set arr2dex2 = Core.Init(New Array2DEx, arr2dex)
    Debug.Print "Array2DExから初期化"
    DebugUtils.PrintVariantArray arr2dex.To2DArray

    Dim arr2dex3 As Array2DEx
    Set arr2dex3 = Core.Init(New Array2DEx, Array(1, 2, 3))
    Debug.Print "1次元配列から初期化"
    DebugUtils.PrintVariantArray arr2dex3.To2DArray
    
    Dim arr2dex4 As Array2DEx
    Set arr2dex4 = Core.Init(New Array2DEx, arr2dex2.ToArrayExOfArrayEx)
    Debug.Print "ArrayExOfArrayExから初期化"
    DebugUtils.PrintVariantArray arr2dex4.To2DArray
    
    Dim arr2dex5 As Array2DEx
    Set arr2dex5 = Core.Init(New Array2DEx, Core.Init(New ArrayEx, Array(1, 2, 3)))
    Debug.Print "ArrayExから初期化"
    DebugUtils.PrintVariantArray arr2dex5.To2DArray
    
    Dim arr2dex6 As Array2DEx
    Set arr2dex6 = Core.Init(New Array2DEx, "test")
    Debug.Print "文字列から初期化"
    DebugUtils.PrintVariantArray arr2dex6.To2DArray
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
    DebugUtils.PrintVariantArray arr2dex.To2DArray
End Sub

Sub TestArray2DEx003_Add()
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx)
    arr2dex.DefaultInitValue = 0
    arr2dex.AddColumns Array(1, 2, 3, 4, 5), Array(1, 2, 3, 4, 5), ArrayUtils.Create2DArrayWithValue(Array(1, 2, 3), Array(4, 5, 6), Array(7, 8, 9))
    DebugUtils.PrintVariantArray arr2dex.To2DArray
End Sub

Sub TestArray2DEx004_Expand()
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx)
    arr2dex.DefaultInitValue = 0
    arr2dex.Expand 5, 5
    DebugUtils.PrintVariantArray arr2dex.To2DArray
End Sub

Sub TestArray2DEx005()
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx).Range(10, 5, 5)
    DebugUtils.PrintVariantArray arr2dex.To2DArray
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
    DebugUtils.PrintVariantArray arr2dex.Uniq.To2DArray
    
    Dim arr2dex2 As Array2DEx, arr2dex3 As Array2DEx
    Set arr2dex2 = Core.Init(New Array2DEx).Range(10, 2, 2)
    Set arr2dex3 = arr2dex.Concat(arr2dex2)
    Debug.Print "Concat"
    DebugUtils.PrintVariantArray arr2dex3.To2DArray
    
    Set arr2dex3 = arr2dex.RowSlice(3, 3)
    Debug.Print "RowSlice"
    DebugUtils.PrintVariantArray arr2dex3.To2DArray
    
    Set arr2dex3 = arr2dex.ColSlice(2, 2)
    Debug.Print "ColSlice"
    DebugUtils.PrintVariantArray arr2dex3.To2DArray
    
    Debug.Print "Range 0"
    DebugUtils.PrintVariantArray Core.Init(New Array2DEx).Range(10, 0, 0).To2DArray
    
    Debug.Print "Map"
    Dim fun As Func
    Set fun = Core.Init(New Func, vbLong, AddressOf TestFuctionSquere)
    DebugUtils.PrintVariantArray arr2dex.Map(fun).To2DArray
    
    Debug.Print "Zip"
    Set arr2dex3 = arr2dex3.Range(1, 3, 2)
    Set fun = Core.Init(New Func, vbLong, AddressOf TestFuctionSumUp)
    DebugUtils.PrintVariantArray arr2dex.Zip(fun, arr2dex3.To2DArray).To2DArray
    
    Debug.Print "RowFilter"
    Set fun = Core.Init(New Func, vbBoolean, AddressOf TestFuctionMyFilter)
    DebugUtils.PrintVariantArray arr2dex.RowFilter(fun).To2DArray
    
    Debug.Print "ColFilter"
    DebugUtils.PrintVariantArray arr2dex.ColFilter(fun).To2DArray

    Debug.Print "RowFold"
    Set fun = Core.Init(New Func, vbLong, AddressOf TestFuctionMyFold)
    DebugUtils.PrintVariantArray arr2dex.RowFold(fun).ToArray
    
    Debug.Print "ColFold"
    DebugUtils.PrintVariantArray arr2dex.ColFold(fun).ToArray

    Debug.Print "RowScan"
    DebugUtils.PrintVariantArray arr2dex.RowScan(fun).To2DArray
    
    Debug.Print "ColScan"
    DebugUtils.PrintVariantArray arr2dex.ColScan(fun).To2DArray
End Sub
Public Function TestFuctionSquere(ByVal source As Long) As Long
    TestFuctionSquere = source * source
End Function
Public Function TestFuctionSumUp(ByVal source1 As Long, ByVal source2 As Long) As Long
    TestFuctionSumUp = source1 + source2
End Function
Public Function TestFuctionMyFilter(ByVal source) As Boolean
    TestFuctionMyFilter = ArrayUtils.Contains(source, 2)
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
    DebugUtils.PrintVariantArray arr2dex.To2DArray
    
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
    DebugUtils.PrintVariantArray arr2dex.To2DArray
    
    DebugUtils.PrintVariantArray arr2dex.RowSearch(1, "郎").ToArray
    DebugUtils.PrintVariantArray arr2dex.ColSearch(0, "子").ToArray
    DebugUtils.PrintVariantArray arr2dex.Search("郎").ToArray
    
    DebugUtils.Show "[正規表現]"
    DebugUtils.PrintVariantArray arr2dex.RowRegexSearch(0, "^.*田.*$").ToArray
    DebugUtils.PrintVariantArray arr2dex.ColRegexSearch(0, ".*子$").ToArray
    DebugUtils.PrintVariantArray arr2dex.RegexSearch(".*子$").ToArray
End Sub

Sub TestArray2DEx009()
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx)
    arr2dex.DefaultInitValue = 0
    arr2dex.AddRows Array(1, 2, 3), Array(1, 2, 3), Array(1, 2, 3), Array(1, 4, 6)

    DebugUtils.Show arr2dex.ToString
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
    For i = 0 To UBound(WinApiFunctions.wD)
        Debug.Print WinApiFunctions.wD(i).wkb.Name
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



