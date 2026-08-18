Attribute VB_Name = "VBFPExamples"

Option Explicit

Public Sub ExampleArithmetic()
    Debug.Print "=== Arithmetic ==="

    Call VBFPPrintExtra("10 + 5 = ", VBFPNamedF("Add").Bind(VBFPValue(10)).Run(VBFPValue(5)))
    Call VBFPPrintExtra("10 - 5 = ", VBFPNamedF("Subtract").Bind(VBFPValue(10)).Run(VBFPValue(5)))
    Call VBFPPrintExtra("10 * 5 = ", VBFPNamedF("Multiply").Bind(VBFPValue(10)).Run(VBFPValue(5)))
    Call VBFPPrintExtra("10 / 5 = ", VBFPNamedF("Divide").Bind(VBFPValue(10)).Run(VBFPValue(5)))

End Sub

Public Sub ExampleBinding()
    Debug.Print "=== Binding ==="

    Dim Add10 As IFunction

    Set Add10 = VBFPNamedF("Add").Bind(VBFPValue(10))

    Call VBFPPrintExtra("10 + 1 = " , Add10.Run(VBFPValue(1)))
    Call VBFPPrintExtra("10 + 5 = " , Add10.Run(VBFPValue(5)))
    Call VBFPPrintExtra("10 + 20 = ", Add10.Run(VBFPValue(20)))

End Sub

Public Sub ExampleMap()
    Debug.Print "=== Map ==="

    Dim Data(4) As IFunction

    Set Data(0) = VBFPValue(1)
    Set Data(1) = VBFPValue(2)
    Set Data(2) = VBFPValue(3)
    Set Data(3) = VBFPValue(4)
    Set Data(4) = VBFPValue(5)

    Dim Elements As IFunction
    Set Elements = VBFPValue(Data)

    Dim Add10 As IFunction
    Set Add10 = VBFPNamedF("Add").Bind(VBFPValue(10))

    Dim Result As IFunction
    Set Result = VBFPNamedF("Map") _
        .Bind(Elements) _
        .Run(Add10)

    Call VBFPPrint(Result)

End Sub

Public Sub ExampleMapArithmetic()
    Debug.Print "=== Map Arithmetic ==="

    Dim Data(4) As IFunction

    Set Data(0) = VBFPValue(1)
    Set Data(1) = VBFPValue(2)
    Set Data(2) = VBFPValue(3)
    Set Data(3) = VBFPValue(4)
    Set Data(4) = VBFPValue(5)

    Dim Elements As IFunction
    Set Elements = VBFPValue(Data)

    Dim Func As IFunction
    Dim Result As IFunction

    Set Func = VBFPNamedF("Add").Bind(VBFPValue(2))
    Set Result = VBFPNamedF("Map").Bind(Elements).Run(Func)

    Debug.Print "Add 2:"
    Call VBFPPrint(Result)

    Set Func = VBFPNamedF("Multiply").Bind(VBFPValue(10))
    Set Result = VBFPNamedF("Map").Bind(Elements).Run(Func)

    Debug.Print "Multiply 10:"
    Call VBFPPrint(Result)

    Set Func = VBFPNamedF("Subtract").Bind(VBFPValue(1))
    Set Result = VBFPNamedF("Map").Bind(Elements).Run(Func)

    Debug.Print "Subtract 1:"
    Call VBFPPrint(Result)

End Sub

Public Sub ExamplePredicates()
    Debug.Print "=== Predicates ==="

    Call VBFPPrintExtra("5 > 2 = ",VBFPNamedF("GreaterThan").Bind(VBFPValue(5)).Run(VBFPValue(2)))
    Call VBFPPrintExtra("5 < 2 = ",VBFPNamedF("LessThan").Bind(VBFPValue(5)).Run(VBFPValue(2)))
    Call VBFPPrintExtra("5 = 5 = ",VBFPNamedF("Equal").Bind(VBFPValue(5)).Run(VBFPValue(5)))

End Sub

Public Sub ExampleStrings()
    Debug.Print "=== Strings ==="

    Dim Hello As IFunction
    Dim World As IFunction

    Set Hello = VBFPValue("Hello, ")
    Set World = VBFPValue("World!")

    Call VBFPPrintExtra("Concatenate", VBFPNamedF("Concatenate").Bind(Hello).Run(World))
    Call VBFPPrintExtra("UpperCase"  , VBFPNamedF("UpperCase").Run(VBFPValue("hello world")))

End Sub

Public Sub ExampleArrayFunctions()
    Debug.Print "=== Array Functions ==="

    Dim Data(4) As IFunction

    Set Data(0) = VBFPValue(10)
    Set Data(1) = VBFPValue(20)
    Set Data(2) = VBFPValue(30)
    Set Data(3) = VBFPValue(40)
    Set Data(4) = VBFPValue(50)

    Dim Elements As IFunction
    Set Elements = VBFPValue(Data)

    Call VBFPPrintExtra("First  = ", VBFPNamedF("First").Run(Elements))
    Call VBFPPrintExtra("Last   = ", VBFPNamedF("Last").Run(Elements))
    Call VBFPPrintExtra("Length = ", VBFPNamedF("ArrayLength").Run(Elements))

End Sub

Public Sub ExampleFold()
    Debug.Print "=== Fold ==="

    Dim Data(4) As IFunction

    Set Data(0) = VBFPValue(1)
    Set Data(1) = VBFPValue(2)
    Set Data(2) = VBFPValue(3)
    Set Data(3) = VBFPValue(4)
    Set Data(4) = VBFPValue(5)

    Dim Elements As IFunction
    Set Elements = VBFPValue(Data)

    Dim Add As IFunction
    Set Add = VBFPNamedF("Add")

    Dim Initial As IFunction
    Set Initial = VBFPValue(0)

    Dim Result As IFunction

    Set Result = VBFPNamedF("Fold") _
        .Bind(Elements, Add) _
        .Run(Initial)

    Call VBFPPrintExtra("1 + 2 + 3 + 4 + 5 = ", Result)

End Sub

Public Sub ExampleFactorial()
    Debug.Print "=== Fold / Factorial ==="

    Dim Data(4) As IFunction

    Set Data(0) = VBFPValue(1)
    Set Data(1) = VBFPValue(2)
    Set Data(2) = VBFPValue(3)
    Set Data(3) = VBFPValue(4)
    Set Data(4) = VBFPValue(5)

    Dim Elements As IFunction
    Set Elements = VBFPValue(Data)

    Dim Multiply As IFunction
    Set Multiply = VBFPNamedF("Multiply")

    Dim Initial As IFunction
    Set Initial = VBFPValue(1)

    Dim Result As IFunction

    Set Result = VBFPNamedF("Fold") _
        .Bind(Elements, Multiply) _
        .Run(Initial)

    Call VBFPPrintExtra("1 * 2 * 3 * 4 * 5 = ", Result)

End Sub

Public Sub ExampleIfThenElse()
    Debug.Print "=== IfThenElse ==="

    Dim Condition As IFunction
    Set Condition = VBFPValue(True)

    Dim WhenTrue As IFunction
    Set WhenTrue = VBFPValue("Condition was true")

    Dim WhenFalse As IFunction
    Set WhenFalse = VBFPValue("Condition was false")

    Dim Result As IFunction

    Set Result = VBFPNamedF("IfThenElse") _
        .Bind(Condition, WhenTrue) _
        .Run(WhenFalse)

    Call VBFPPrint(Result)

End Sub

Public Sub ExamplePipeline()
    Debug.Print "=== Pipeline ==="

    Dim Value As IFunction
    Set Value = VBFPValue(5)

    Dim Add10 As IFunction
    Set Add10 = VBFPNamedF("Add").Bind(VBFPValue(10))

    Dim Multiply2 As IFunction
    Set Multiply2 = VBFPNamedF("Multiply").Bind(VBFPValue(2))

    Dim Result As IFunction

    Set Result = Add10.Run(Value)
    Set Result = Multiply2.Run(Result)

    Call VBFPPrintExtra("(5 + 10) * 2 = ", Result)

End Sub

Public Sub RunExamples()
    ExampleArithmetic
    ExampleBinding
    ExampleMap
    ExampleMapArithmetic
    ExamplePredicates
    ExampleStrings
    ExampleArrayFunctions
    ExampleFold
    ExampleFactorial
    ExampleIfThenElse
    ExamplePipeline
End Sub