# IFunction


IFunction is the basis for everything.
With IFunction it is possible to store anything, as long as it is able to retrieve some Data.

Imagine IFunction as a Factory, that produces some Value.
It does not matter how it produces this Value, as long as it does.

This entire Library expects to only work with IFunction.
That means, that every function is expected to take in IFunction and produce IFunction.
Everything else is by definition an Error.


In practice IFunction is an Interface that currently has 4 Virtuals:
1. Run         --> Executes the inner workings of its "Factory", dependant on Implementation, returns IFunction (Expected)
2. Evaluate    --> Returns the underlying Data of the created IFunction from Run, when you need to work with the actual Data
3. RunArr      --> Previous one had ParamArray for convenience, but working with ParamArray is a pain, so array it is
4. EvaluateArr --> Previous one had ParamArray for convenience, but working with ParamArray is a pain, so array it is

The reasoning behind these Virtuals is the following:
Since we are dealing purely with IFunction with this Library, everything will try to take it as Input.
But because we are still in VBA and we need actual Values sometimes outside or inside the Library the Evaluate provides access to that Data.

We have several [Implementations](Implementations.md):