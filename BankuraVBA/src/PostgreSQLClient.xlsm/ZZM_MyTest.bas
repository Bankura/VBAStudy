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





