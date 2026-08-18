VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NamedFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IFunction

Private FunctionName As String
Private Bound()      As IFunction

Friend Function Bind(ByRef n_Bound() As IFunction) As IFunction
    Bound = n_Bound
    Set Bind = Me
End Function

Friend Sub Initialize(ByVal n_Name As String)
    FunctionName = n_Name
End Sub

Public Function Create(ByVal FunctionName As String) As NamedFunction
    Set Create = New NamedFunction
    Call Create.Initialize(FunctionName)
End Function

Private Function IFunction_Evaluate(ParamArray Arguments() As Variant) As Variant
    Dim Size As Long
    Size = UBound(Arguments)

    Dim Args() As IFunction

    If Size > -1 Then
        ReDim Args(Size)
        
        Dim i As Long
        For i = 0 To Size
            Set Args(i) = Arguments(i)
        Next i
    End If

    Call VBFPAssign(IFunction_Evaluate, Run(Args).Evaluate())
End Function

Private Function IFunction_Run(ParamArray Arguments() As Variant) As IFunction
    Dim Size As Long
    Size = UBound(Arguments)

    Dim Args() As IFunction

    If Size > -1 Then
        ReDim Args(Size)
        
        Dim i As Long
        For i = 0 To Size
            Set Args(i) = Arguments(i)
        Next i
    End If

    Set IFunction_Run = Run(Args)
End Function

Public Function IFunction_Bind(ParamArray Arguments() As Variant) As IFunction
    Dim Result As NamedFunction
    Dim i      As Long
    Dim Count  As Long
    Dim Args() As IFunction

    Set Result = New NamedFunction
    Call Result.Initialize(FunctionName)
    Count = UBound(Arguments)

    ReDim Args(Count)
    For i = 0 To Count
        Set Args(i) = Arguments(i)
    Next i
    Set IFunction_Bind = Result.Bind(Args)
End Function


Private Function Run(ByRef Arguments() As IFunction) As IFunction
    Dim A()    As Variant
    Dim Count  As Long
    Dim Result As Variant

    A = BuildArguments(Arguments, Count)
    Call VBFPAssign(Run, RunFunction(A))
End Function

Private Function BuildArguments(ByRef Arguments() As IFunction, ByRef Count As Long) As Variant()
    Dim PassedCount As Long   : PassedCount = USize(Arguments)
    Dim BoundCount  As Long   : BoundCount = USize(Bound)
    Dim TotalCount  As Long   : TotalCount = (BoundCount + 1 + PassedCount + 1) - 1
    Dim Result()    As Variant: ReDim Result(TotalCount)

    Dim i As Long
    For i = 0 To BoundCount
        Set Result(i) = Bound(i)
    Next i
    Dim Current As Long
    Current = i
    For i = 0 To PassedCount
        Set Result(Current + i) = Arguments(i)
    Next i
    BuildArguments = Result
End Function

Private Function RunFunction(ByRef A() As Variant) As Variant
    Dim Count As Long
    Count = UBound(A)
    Select Case Count + 1
        Case 00: Call VBFPAssign(RunFunction, Application.Run(FunctionName))
        Case 01: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0)))
        Case 02: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1)))
        Case 03: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2)))
        Case 04: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3)))
        Case 05: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4)))
        Case 06: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5)))
        Case 07: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6)))
        Case 08: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7)))
        Case 09: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8)))
        Case 10: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9)))
        Case 11: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10)))
        Case 12: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11)))
        Case 13: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12)))
        Case 14: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13)))
        Case 15: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14)))
        Case 16: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15)))
        Case 17: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16)))
        Case 18: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17)))
        Case 19: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18)))
        Case 20: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19)))
        Case 21: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20)))
        Case 22: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21)))
        Case 23: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22)))
        Case 24: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23)))
        Case 25: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23), A(24)))
        Case 26: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23), A(24), A(25)))
        Case 27: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23), A(24), A(25), A(26)))
        Case 28: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23), A(24), A(25), A(26), A(27)))
        Case 29: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23), A(24), A(25), A(26), A(27), A(28)))
        Case 30: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23), A(24), A(25), A(26), A(27), A(28), A(29)))
        Case Else: Err.Raise 5, "NamedFunction", "Unsupported number of arguments: " & Count
    End Select
End Function

Private Function USize(ByRef X As Variant) As Long
    On Error Resume Next
    USize = -1
    USize = UBound(X)
End Function