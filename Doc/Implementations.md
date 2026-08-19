# ValueFunction
This is an important Implementation, since this represents a variable in normal VBA Terms.
It is needed when you want to exchange Data, but is by far the least interesting Implementation when using FP.

* Holds a Variant Value
* Just a Container to hold a Value
* Expected to be immutable after Creation
* Passed Arguments have no influence on the return Value
    * Imagine a function with no Input Arguments
    * By definition of FP it must always produce the same Output given the same Input
    * This means no arguments --> always same Output

---

# NamedFunction
This Implementation is what unlocks Functional Programming for VBA with its further following Implementations.
To keep the Library small and consistent this Impl. only takes the Name of the function to call as Input.
Compared to the previous Impl. it does not just return a Value, but runs a function and returns its Value.

* Holds a (non-safe) reference to a function
* Runs the function when returning its Value
* Expected to be immutable after Creation
* Passed Arguments will be forwarded to function

---

# BoundFunction
This is the strongest selling point in my opinion for Functional Programming and the most interesting for practical use.
It allows to bind an IFunction and some Arguments and later when running it combine them in order to pass them to said function.
Imagine: `Create(Add, 6)` then afterwards: `x.Run(7)  --> x.Run([6], 7)`

* Holds a reference to an IFunction
* Holds bound Arguments
* Expected to be immutable after Creation
* Passed Arguments will be combined with bound ones and forwarded to function

---

# ComposedFunction
Just combines an outer and inner function call into a single callabe function.
The goto if you want to chain functions, together with the Pipe function

* Holds a reference to an inner and outer IFunction
* Expected to be immutable after Creation
* Passed Arguments forwarded to inner

---

# MemoizedFunction
Recursively saves up the result of a function call according its input arguments.
If the value was already run with the same inputs as before the call will short circuit and return the Value that was already saved.

* Holds a reference to a Function
* Expected to be immutable after Creation
* Passed Arguments forwarded to Function
* Saves the Result
* Uses Result when the same Input was already calculated once

---

# CurriedFunction
This one is less interesting in its properties and is basically just here for completeness rather than usefulness.
Currying will produce a chain of Function calls with one bound Argument for the total count of Arguments.
In theoretical Terms it will produce a function chain that when called will transform it into:
```cpp
    CurriedAdd(1) --> Forward2
    Forward2(2)   --> Forward3
    Forward3(3)   --> Forward4
    Forward4(4)   --> Add5
    Add5(5)       --> 15
```

Each intermediate result is therefore itself a function. The first call does not produce the final value and does not simply collect all arguments into a list. Instead, each call binds exactly one argument and returns the next function in the chain.

This is the key difference between BoundFunction and CurriedFunction:
* BoundFunction binds any number of arguments and waits for the remaining arguments
* CurriedFunction binds exactly one argument per call and returns another function
* Once the required number of arguments has been supplied, the final function call produces the actual result

At least, as of right now, in **theory**, as the current underlying code will just combine the single arguments into a list and just create a BoundFunction.
Basically it is useless right now. But a current update will try to fix it, so that a call like `Forward1(1)(2)(3)(4)(5)` is possible