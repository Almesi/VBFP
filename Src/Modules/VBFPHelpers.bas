Attribute VB_Name = "VBFPHelpers"

Option Explicit

Public Function VBFPValueF(ByVal Value As Variant) As IFunction
    Set VBFPValueF = VBFPValue(Value)
End Function
Public Function VBFPValue(ByVal Value As Variant) As ValueFunction
    Set VBFPValue = ValueFunction.Create(Value)
End Function

Public Function VBFPNamedF(ByVal Name As String) As IFunction
    Set VBFPNamedF = VBFPNamed(Name)
End Function
Public Function VBFPNamed(ByVal Name As String) As NamedFunction
    Set VBFPNamed = NamedFunction.Create(Name)
End Function

Public Sub VBFPPrint(ByVal Value As IFunction)
    Dim Evaluated As Variant
    Dim Element   As Variant

    Evaluated = Value.Evaluate()

    If IsObject(Evaluated) Then
        If TypeOf Evaluated Is IFunction Then
            Call VBFPPrint(Evaluated)
            Exit Sub
        End If
    End If

    If IsArray(Evaluated) Then
        For Each Element In Evaluated
            If TypeOf Element Is IFunction Then
                Call VBFPPrint(Element)
            Else
                Debug.Print Element
            End If
        Next Element
    Else
        Debug.Print Evaluated
    End If
End Sub

Public Sub VBFPPrintExtra(ByVal ExtraText As String, ByVal Value As IFunction)
    Debug.Print ExtraText
    Call VBFPPrint(Value)
End Sub

Public Sub VBFPAssign(ByRef Goal As Variant, ByVal Value As Variant)
    If IsObject(Value) Then
        Set Goal = Value
    Else
        Let Goal = Value
    End If
End Sub