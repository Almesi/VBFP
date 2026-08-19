Functional programming (FP) is a programming paradigm built around the idea of **computation through functions**. Instead of primarily thinking in terms of objects that change state, procedures that perform actions, or variables that are repeatedly modified, FP treats functions as the fundamental building blocks of a program.

In a functional system, a function is ideally understood as a mathematical mapping:

**Input --> Output**
Given the same input, the function should produce the same output. This simple idea leads to several important properties that make functional programming distinctive.


### The basic pillars of functional programming
**1. Functions are first-class values**

Functions are treated like data. They can be stored, passed as arguments, returned from other functions, and combined to create new functions.

This is particularly important for VBA because VBA does not naturally treat functions as values in the same way that languages designed around FP do. An abstraction such as `IFunction` therefore provides a way to represent both **data and computation through a common interface**.

Instead of having the library distinguish between:

```
a value
a function
a function with some arguments
a function that produces another function
```

everything can be represented as an `IFunction`.
This gives the library a single fundamental abstraction:
**An IFunction is something that can produce a value.**

---

**2. Immutability**
Functional programming generally prefers immutable data. Once a value or function has been created, it should not be modified.
Instead of changing an existing value:

```
x = 5
x = 10
```

functional programming favors creating another value:

```
x = 5
y = 10
```

This makes programs easier to reason about because the meaning of an existing value does not change somewhere else in the program.

---

**3. Pure functions**
A pure function depends only on its inputs and produces an output without observable side effects.
Conceptually:

```
f(x) = y
```

Calling `f(x)` again should produce the same `y`.
Pure functions are powerful because they can be composed, reused, tested, cached, and reasoned about independently of the rest of the program.
This is one of the most important goals of functional programming in VBA: **to make computation behave as predictably as possible despite VBA itself being heavily oriented around mutable state, procedures, and side effects.**
Sadly, this is also the part that is the hardest to control via this library.
Esentially, the only way to ensure this is you, dear reader, by writing in a purely functional way.
But if you stick to just using IFunction and avoid global and public state, you should be good to go.

---

**4. Function composition**
Small functions become significantly more useful when they can be combined into larger functions.

For example:
```
A --> B --> C
```

means that the output of one function becomes the input of another.

Rather than writing one large procedure containing every operation, functional programming encourages constructing a larger computation from smaller, reusable functions.
With an `IFunction` abstraction, composition can conceptually operate entirely on `IFunction`s:

```
IFunction --> IFunction --> IFunction
```

The library therefore does not need to constantly leave the functional world to manipulate ordinary VBA values. Values can remain wrapped in `ValueFunction`s, while computations remain represented by other `IFunction` implementations until evaluation is actually required.

Best practical example for this would be a generic `ForEach` Function, that takes as input a list of elements and a function to call it on each of its elements
This function can then be used everywhere as a higher abstraction. 

---

### Why `IFunction` is important
The central idea behind this library is therefore more than simply creating a collection of functional-programming utilities.
It establishes a **functional universe inside VBA**.
Everything is an `IFunction`.

A value is an `IFunction`.
A named function is an `IFunction`.
A function with arguments already attached is an `IFunction`.
A curried function is an `IFunction`.
And the result of running a function is expected to remain an `IFunction`.

The actual underlying VBA value is only retrieved when `Evaluate` or `EvaluateArr` is explicitly required.
This creates a separation between **computation** and **data**

That distinction is fundamental. The library can continue manipulating functions without constantly evaluating them prematurely.

---


### Functional programming in VBA

The larger goal of the library is not necessarily to make VBA behave exactly like a language such as Haskell. That would fight against the nature of the language.
Instead, the goal is to implement the fundamental ideas of functional programming as faithfully as VBA allows.

VBA has several characteristics that work against FP:
* mutable variables are normal
* procedures and functions are primarily invoked rather than treated as values
* references are not inherently safe or immutable
* side effects are easy to introduce
* the language has limited native support for higher-order functions
* `Variant` and runtime dispatch make strong guarantees difficult.

The library therefore provides an abstraction layer that compensates for these limitations.
`IFunction` becomes the common language through which functional operations are expressed.

The objective is to make concepts such as:
* first-class functions
* higher-order functions
* immutability
* pure computation
* partial application
* currying
* composition
* lazy or deferred evaluation
* and functional data transformation

possible within VBA while maintaining a consistent API.