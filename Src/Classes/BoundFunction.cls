VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "BoundFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IFunction

Private p_Function As IFunction
Private p_Bound()  As IFunction

Friend Sub Initialize(ByVal n_Function As IFunction, ByRef n_Bound() As IFunction)
    Set p_Function = n_Function
    p_Bound        = n_Bound
End Sub

Public Function Create(ByVal n_Function As IFunction, ParamArray Arguments() As Variant) As BoundFunction
    Dim Temp() As Variant
    Dim Bound() As IFunction

    Temp = Arguments
    Bound = ArgumentsToFunctions(Temp)
    
    Set Create = CreateArr(n_Function, Bound)
End Function

Public Function CreateArr(ByVal n_Function As IFunction, ByRef Arguments() As IFunction) As BoundFunction
    Set CreateArr = New BoundFunction
    Call CreateArr.Initialize(n_Function, Arguments)
End Function

Private Function IFunction_Evaluate(ParamArray Arguments() As Variant) As Variant
    Dim Temp() As Variant
    Dim Passed() As IFunction

    Temp = Arguments
    Passed = ArgumentsToFunctions(Temp)
    Call VBFPAssign(IFunction_Evaluate, Run(Passed).Evaluate())
End Function

Private Function IFunction_EvaluateArr(ByRef Arguments() As IFunction) As Variant
    Call VBFPAssign(IFunction_EvaluateArr, Run(Arguments).Evaluate())
End Function

Private Function IFunction_Run(ParamArray Arguments() As Variant) As IFunction
    Dim Temp() As Variant
    Dim Passed() As IFunction

    Temp = Arguments
    Passed = ArgumentsToFunctions(Temp)
    Call VBFPAssign(IFunction_Run, Run(Passed))
End Function

Private Function IFunction_RunArr(ByRef Arguments() As IFunction) As IFunction
    Call VBFPAssign(IFunction_RunArr, Run(Arguments))
End Function

Private Function Run(ByRef Arguments() As IFunction) As IFunction
    Dim Combined() As IFunction
    Combined = Combine(p_Bound, Arguments)

    Set Run = p_Function.RunArr(Combined)
End Function

Private Function Combine(ByRef First() As IFunction, ByRef Second() As IFunction) As IFunction()
    Dim FirstCount  As Long
    Dim SecondCount As Long
    Dim TotalCount  As Long

    FirstCount  = USize(First)
    SecondCount = USize(Second)

    If FirstCount < 0 And SecondCount < 0 Then
        Dim Emptyy() As IFunction
        Combine = Emptyy
        Exit Function
    End If

    TotalCount = (FirstCount + 1) + (SecondCount + 1)

    Dim Result() As IFunction
    ReDim Result(TotalCount - 1)

    Dim i As Long
    For i = 0 To FirstCount
        Set Result(i) = First(i)
    Next i

    Dim Position As Long
    Position = i
    For i = 0 To SecondCount
        Set Result(Position + i) = Second(i)
    Next i

    Combine = Result
End Function