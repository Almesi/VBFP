
---

# 1. `Map` — transform every element
```vb
Public Sub ExampleMap()
    Dim SquareFunc As IFunction
    Set SquareFunc = NamedFunction.Create("Square")

    Dim Values(5) As IFunction
    Values(0) = VBFPValue(1)
    Values(1) = VBFPValue(2)
    Values(2) = VBFPValue(3)
    Values(3) = VBFPValue(4)
    Values(4) = VBFPValue(5)

    Dim Result As IFunction
    Result = Map(Values, SquareFunc)

    Call VBFPPrint(Result) ' 1, 4, 9, 16, 25
End Sub
```

---

# 2. `Filter` — keep values satisfying a function
```vb
Public Sub ExampleFilter()
    Dim IsEvenFunc As IFunction
    Set IsEvenFunc = NamedFunction.Create("IsEven")

    Dim Values(6) As IFunction
    Values(0) = VBFPValue(1)
    Values(1) = VBFPValue(2)
    Values(2) = VBFPValue(3)
    Values(3) = VBFPValue(4)
    Values(4) = VBFPValue(5)
    Values(5) = VBFPValue(6)

    Dim Result As IFunction

    Result = Filter(Values, IsEvenFunc, )

    Call VBFPPrint(Result) ' 2, 4, 6
End Sub
```

---

# 3. `ForEach` — execute a function for every value
```vb
Public Sub ExampleForEach()
    Dim PrintFunc As IFunction
    Set PrintFunc = NamedFunction.Create("PrintFunction")

    Dim Values(5) As IFunction
    Values(0) = VBFPValue(1)
    Values(1) = VBFPValue(2)
    Values(2) = VBFPValue(3)
    Values(3) = VBFPValue(4)
    Values(4) = VBFPValue(5)

    Call ForEach(Values, PrintFunc)
End Sub
```

---

# 4. Combining `Map` and `Filter`
```vb
Public Sub ExampleMapFilter()
    Dim IsEvenFunc As IFunction
    Dim SquareFunc As IFunction

    Set IsEvenFunc = NamedFunction.Create("IsEven")
    Set SquareFunc = NamedFunction.Create("Square")

    Dim Values(5) As IFunction
    Values(0) = VBFPValue(1)
    Values(1) = VBFPValue(2)
    Values(2) = VBFPValue(3)
    Values(3) = VBFPValue(4)
    Values(4) = VBFPValue(5)
    Values(5) = VBFPValue(6)

    Dim EvenNumbers As IFunction
    Set EvenNumbers = Filter(Values, IsEvenFunc)

    Dim Result As Variant
    Result = Map(EvenNumbers, SquareFunc)

End Sub
```

---

# 5. `Pipe`
```vb
Dim Value        As IFunction
Dim AddTaxFunc   As IFunction
Dim DiscountFunc As IFunction
Dim Result       As IFunction

Set Value        = ValueFunction.Create(10)
Set AddTaxFunc   = NamedFunction.Create("AddTax")
Set DiscountFunc = NamedFunction.Create("ApplyDiscount")
Set Result = Pipe(DiscountFunc, AddTaxFunc)
```

# 6. Memoization with Fibonacci
```vb
Dim Fib As IFunction
Set Fib = MemoizedFunction.Create(NamedFunction.Create("Fibonacci"))

Debug.Print Fib.Evaluate(40)
Debug.Print Fib.Evaluate(40) ' recall
Debug.Print Fib.Evaluate(40) ' recall
```

# 7. A complete FP pipeline

Putting everything together:

```text
Orders
   ||
   \/
Map(GetPrice)
   ||
   \/
[100, 200, 50, 300]
   ||
   \/
Filter(IsExpensive)
   ||
   \/
[100, 200, 300]
   ||
   \/
Map(AddTax)
   ||
   \/
[119, 238, 357]
   ||
   \/
Reduce(Add)
   ||
   \/
714
```

The important thing is that every operation can be expressed in terms of functions:

```vb
Dim GetPrice        As IFunction : Set GetPrice        = NamedFunction.Create("GetPrice")
Dim IsExpensive     As IFunction : Set IsExpensive     = NamedFunction.Create("IsExpensive")
Dim AddTax          As IFunction : Set AddTax          = BoundFunction.Create("Multiply", 1.19)
Dim Add             As IFunction : Set Add             = NamedFunction.Create("Add")

Dim Prices          As IFunction : Set Prices          = Map(Orders, GetPrice)
Dim ExpensivePrices As IFunction : Set ExpensivePrices = Filter(Prices, IsExpensive)
Dim TaxedPrices     As IFunction : Set TaxedPrices     = Map(ExpensivePrices, AddTax)
Dim Total           As IFunction : Set Total           = Reduce(TaxedPrices, Add, 0)
```