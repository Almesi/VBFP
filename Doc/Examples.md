# Examples

The following examples demonstrate the basic usage patterns of the functional
abstractions provided by the library.

---

## Wrapping a VBA Function

This is NOT possible, as it does not work solely with IFunction:

```vb
Public Function Add(A As Long, B As Long) As Long
    Add = A + B
End Function
```

This here however IS possible:


```vb
Public Function Add(ByVal Num1 As IFunction, ByVal Num2 As IFunction) As IFunction
    Set Add = VBFPValue(Num1.Evaluate() + Num2.Evaluate())
End Function

Public Sub Test()
    Dim Func As IFunction
    Set Func = NamedFunction.Create("Add")

    Dim Result As IFunction
    Set Result = Func.Run(VBFPValue(2), VBFPValue(3)) '5
End Sub
```

---

## 2. `Run` vs `Evaluate`

`Run` returns `IFunction`.
`Evaluate` returns the underlying VBA value.

For example:

```vb
Dim Func   As IFunction
Set Func = NamedFunction.Create("Add")

Dim Result As IFunction
Set Result = Func.Run(VBFPValue(6), VBFPValue(7)) ' Container holding 13 as Variant

Debug.Print Result.Evaluate() ' Returns 13
```

---

## 3. Passing an IFunction as an Argument

Because functions and values are both represented by `IFunction`, they can
be passed around using the same interface.

For example:

```vb
Dim A    As IFunction
Dim B    As IFunction
Dim Func As IFunction

Set A = ValueFunction.Create(10)
Set B = BoundFunction.Create("Square", 2)
Set Func = NamedFunction.Create("Add")

Debug.Print Func.Evaluate(A, B) ' 10 + (2*2) = 14
```

---

## 5. Binding Arguments

Create the underlying function:

```vb
Dim Func As IFunction
Set Func = NamedFunction.Create("Add")
```

Bind the first argument:

```vb
Dim AddSix As IFunction
Set AddSix = BoundFunction.Create(Func, 6)
```

`AddSix` now represents:

```text
Add(6, X)
```

The remaining argument can be supplied later:

```vb
Debug.Print AddSix.Evaluate(7) ' 6 + 7
```

---

## 6. Binding Multiple Arguments

`BoundFunction` can bind multiple arguments at once.

For example:

```vb
Dim Func       As IFunction
Dim AddSeveral As IFunction

Set Func       = NamedFunction.Create("Add")
Set AddSeveral = BoundFunction.Create(Func, 6, 7)
```

The resulting function represents:

```text
Add(6, 7)
```

It can then be evaluated without supplying additional arguments:

---

## 7. Binding Functions Instead of Values

Because arguments are represented as `IFunction`s, a bound argument does not
necessarily have to be a plain VBA value.

For example:

```vb
Dim A          As IFunction
Dim B          As IFunction
Dim Func       As IFunction
Dim AddValues  As IFunction

Set A         = ValueFunction.Create(10)
Set B         = BoundFunction.Create("Square", 2)
Set Func      = NamedFunction.Create("Add")
Set AddValues = BoundFunction.Create(Func, A, B)
```

The computation can then be evaluated:

```vb
Debug.Print AddValues.Evaluate()
```

Result:

```text
10 + (2*2) = 14
```

This allows computations to be built from other computations rather than
only from immediate values.

---

# Function Pipelines

A typical functional pattern is to process a value through several
computations.

For example:

```text
X -->
Add(5) -->
Multiply(2) -->
Result
```

For:

```text
X = 10
```

the computation becomes:

```text
10 -->
Add(5) -->
15 -->
Multiply(2) -->
30
```

The important concept is that the intermediate operations can themselves
be represented by `IFunction`s.

---

# Currying

Currying is different from ordinary argument binding.

A bound function can bind several arguments at once:

```text
Add(1, 2, 3)
```

A curried function instead accepts exactly one argument at a time.

For a hypothetical five-argument function:

```text
Function(A, B, C, D, E)
```

currying transforms it conceptually into:

```text
Function(A) -->
Function(B) -->
Function(C) -->
Function(D) -->
Function(E) -->
Result
```

---

# RunArr and EvaluateArr

The `Arr` variants exist to avoid the limitations of `ParamArray`.

Instead of:

```vb
Function.Evaluate(1, 2, 3)
```

an array of `IFunction`s can be supplied:

```vb
Dim Arguments(2) As IFunction

Set Arguments(0) = ValueFunction.Create(1)
Set Arguments(1) = ValueFunction.Create(2)
Set Arguments(2) = ValueFunction.Create(3)

Debug.Print Function.EvaluateArr(Arguments)
```

Likewise, `RunArr` returns an `IFunction`:

```vb
Dim Result As IFunction
Set Result = Function.RunArr(Arguments)
Debug.Print Result.Evaluate()
```

This provides an array-based alternative for situations where arguments are
already stored or generated dynamically.

---

# Summary

The most important usage principle is:
Do not think of an `IFunction` as the value itself.
Think of it as something that can produce a value.