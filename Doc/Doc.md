IFunction is the basis for everything.
With IFunction it is possible to store anything, as long as it is able to retrieve some Data.
IFunction is an Interface that currently has 3 Virtuals:
1. Bind     --> Tries to bind Data (IFunction) to the Object for the next Retrieval
2. Run      --> Retrieves the underlying Data, dependant on Implementation
3. Evaluate --> Returns the underlying Data of the created IFunctin from Run

The reasoning behind these Virtuals is the following:
Since we are dealing purely with IFunction with this Library, everything will try to take it as Input.
But because we are still in VBA and we need actual Values sometimes outside or inside the Library the Evaluate provides access to that Data.
Bind is the Interesting part and the main reason this Library was created, to provide Functions with bound Arguments.

We have 2 Implementations:
1. ValueFunction
    * Just a Wrapper to some Value
    * Can be anything, Numbers, String, Variant or another IFunction
2. NamedFunction
    * Holds the Name to a Function
    * Holds bound Arguments
    * Will call the function together with bound and passed arguments on Run


# Workflow
The Idea is, to always use IFunction when working with this Library.
Lets take a simple example:

```vb
Public Function Add(ByVal Num1 As IFunction, ByVal Num2 As IFunction) As IFunction
    Set Add = VBFPValue(Num1.Evaluate() + Num2.Evaluate())
End Function
```

Why so complicated? Why not just passing in Doubles?
Because with functional programming we unlock Bound Function which make this very interesting

This way we not only pass in Values but we can call another Function too.

Lets do for example:
```vb
Num1 = NamedFunction("CircleArea").Bind(ValueFunction(10))
Func = NamedFunction("Add").Bind(Num1)


Debug.Print Func.Evaluate(ValueFunction(1)) ' 10^2 * Pi + 1
Debug.Print Func.Evaluate(ValueFunction(2)) ' 10^2 * Pi + 2
Debug.Print Func.Evaluate(ValueFunction(3)) ' 10^2 * Pi + 3
Debug.Print Func.Evaluate(ValueFunction(4)) ' 10^2 * Pi + 4
```

We bound a Function to Add and then just called with one Extra argument, the Function figured it out itself.
This might seem like a bad example, so lets show one of the more powerful usecases:

```vb
Public Function Map(ByVal Elements As IFunction, ByVal Func As IFunction) As IFunction
    Dim Source() As IFunction : Source = Elements.Evaluate()
    Dim Size     As Long      : Size   = UBound(Source)
    Dim Result() As IFunction : ReDim Result(Size)

    Dim i As Long
    For i = 0 To Size
        Set Result(i) = Func.Run(Source(i))
    Next i

    Set Map = VBFPValue(Result)
End Function

Public Sub Test()
    Set Range   = ValueFunction(Array(1, 2, 3, 4, 5, 6))
    Set Func    = NamedFunction("CircleArea")
    Set MapFunc = NamedFunction("Map")

    Set Result  = MapFunc.Evaluate(Range, Func)
    ' 1^2*Pi
    ' 2^2*Pi
    ' 3^2*Pi
    ' 4^2*Pi
    ' 5^2*Pi
    ' 6^2*Pi
    Set Func2   = NamedFunction("Multiply").Bind(ValueFunction(2))
    Set Result2 = MapFunc.Evaluate(Range, Func2)
    ' 2*1
    ' 2*2
    ' 2*3
    ' 2*4
    ' 2*5
    ' 2*6

End Sub
```



# IFunction

IFunction is the basis for everything.

With IFunction it is possible to store anything, as long as it is able to retrieve some Data.

IFunction is an Interface that currently has 3 Virtuals:

1. **Bind**
   --> Tries to bind Data (IFunction) to the Object for the next Retrieval

2. **Run**
   --> Retrieves the underlying Data, dependant on Implementation

3. **Evaluate**
   --> Returns the underlying Data of the created IFunction from Run

### The reasoning behind these Virtuals is the following:

Since we are dealing purely with IFunction with this Library, everything will try to take it as Input.

But because we are still in VBA and we need actual Values sometimes outside or inside the Library, the Evaluate provides access to that Data.

Bind is the Interesting part and the main reason this Library was created: to provide Functions with bound Arguments.

---

# Implementations

We have 2 Implementations:

### 1. ValueFunction

* Just a Wrapper to some Value
* Can be anything: Numbers, String, Variant or another IFunction

### 2. NamedFunction

* Holds the Name to a Function
* Holds bound Arguments
* Will call the Function together with bound and passed arguments on Run

---

# Workflow

The Idea is, to always use IFunction when working with this Library.

Lets take a simple example:

```vb
Public Function Add(ByVal Num1 As IFunction, ByVal Num2 As IFunction) As IFunction
    Set Add = VBFPValue(Num1.Evaluate() + Num2.Evaluate())
End Function
```

Why so complicated? Why not just passing in Doubles?

Because with functional programming we unlock Bound Functions which make this very interesting.

This way we not only pass in Values but we can call another Function too.

Lets do for example:

```vb
Set Num1 = NamedFunction("CircleArea").Bind(ValueFunction(10))
Set Func = NamedFunction("Add").Bind(Num1)

Debug.Print Func.Evaluate(ValueFunction(1)) ' 10^2 * Pi + 1
Debug.Print Func.Evaluate(ValueFunction(2)) ' 10^2 * Pi + 2
Debug.Print Func.Evaluate(ValueFunction(3)) ' 10^2 * Pi + 3
Debug.Print Func.Evaluate(ValueFunction(4)) ' 10^2 * Pi + 4
```

We bound a Function to Add and then just called with one Extra argument, the Function figured it out itself.

This might seem like a bad example, so lets show one of the more powerful usecases:

```vb
Public Function Map(ByVal Elements As IFunction, ByVal Func As IFunction) As IFunction
    Dim Source() As IFunction : Source = Elements.Evaluate()
    Dim Size     As Long      : Size = UBound(Source)
    Dim Result() As IFunction : ReDim Result(Size)

    Dim i As Long
    For i = 0 To Size
        Set Result(i) = Func.Run(Source(i))
    Next i

    Set Map = VBFPValue(Result)
End Function

Public Sub Test()
    Set Range   = ValueFunction(Array(1, 2, 3, 4, 5, 6))
    Set Func    = NamedFunction("CircleArea")
    Set MapFunc = NamedFunction("Map")

    Set Result = MapFunc.Evaluate(Range, Func)
    ' 1^2 * Pi
    ' 2^2 * Pi
    ' 3^2 * Pi
    ' 4^2 * Pi
    ' 5^2 * Pi
    ' 6^2 * Pi

    Set Func2   = NamedFunction("Multiply").Bind(ValueFunction(2))
    Set Result2 = MapFunc.Evaluate(Range, Func2)
    ' 2 * 1
    ' 2 * 2
    ' 2 * 3
    ' 2 * 4
    ' 2 * 5
    ' 2 * 6
End Sub
```

This way we can easily change what the map function does just by changing which function it takes as Input.
You might ask yourself why, when excel already provides a MAP function with lambda as input.
Well, once this Library is finished, setting those things up in excel will be much easier.
You wouldnt need to create more functions and it is more powerful.