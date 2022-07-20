```
  _______   ______    ________   ______   ___   ___   ______   _________  ______
/_______/\ /_____/\  /_______/\ /_____/\ /___/\/__/\ /_____/\ /________/\/_____/\
\::: _  \ \\:::_ \ \ \::: _  \ \\:::__\/ \::.\ \\ \ \\::::_\/_\__.::.__\/\::::_\/_
 \::(_)  \/_\:(_) ) )_\::(_)  \ \\:\ \  __\:: \/_) \ \\:\/___/\  \::\ \   \:\/___/\
  \::  _  \ \\: __ `\ \\:: __  \ \\:\ \/_/\\:. __  ( ( \::___\/_  \::\ \   \_::._\:\
   \::(_)  \ \\ \ `\ \ \\:.\ \  \ \\:\_\ \ \\: \ )  \ \ \:\____/\  \::\ \    /____\:\
    \_______\/ \_\/ \_\/ \__\/\__\/ \_____\/ \__\/\__\/  \_____\/   \__\/    \_____\/
```

A powerful, elegent & lightweight Functional Programming library for *VBScript*, overturn your understanding of *VBScript* programming.

The first anonymous function support for *VBScript*, add your environment on demand.

Dozens of common, general and easy-to-use methods are also provided.

Through the encapsulation and nesting of functions, complex problems are stripped of their cocoons to show the essence of transformation.

It's just like magic! Programming and mathematics can be so similar!

Enjoy the charm of functional programming!


# View introduction in

- [Chinese](Readme_zh.md)
- [English](Readme.md)

# Getting Started

## Requirements

- Microsoft Windows Operating System

## Installation

Run following commands as **administrator**:

```
git clone https://github.com/OldLiu001/Brackets.vbs.git
cd Brackets.vbs
regsvr32 Brackets.wsc
```

**WARN: DO NOT REGISTER *Brackets.wsc* BY RIGHT CLICKING ON IT.**

Then to create a instance of class, use following code:

```
Set [] = CreateObject("Brackets")
```

## Portability

A Portable version can help you publish your script to others.

Copy script file *Brackets.vbs* & *Brackets.js* to your script's parent folder.

Assume that your script's file name is *MyScript.vbs*, use following template code:


*Template.wsf*

```
<job id="MyScript">
	<script language="JScript" src="Brackets.js"/>
	<script language="VBScript" src="Brackets.vbs"/>
	<script language="VBScript" src="MyScript.vbs"/>
</job>
```

Save *Template.wsf* to folder where your script also in.

In another way, you can embedding script & library into a single *WSF*:

*Template_Embedded.wsf*

```
<job id="MyScript">
	<script language="JScript">
		// contents of "Brackets.js"
	</script>
	<script language="VBScript">
		' contents of "Brackets.vbs"
	</script>
	<script language="VBScript">
		' contents  of "MyScript.vbs"
	</script>
</job>
```

Of course, you can only embed necessary part(s) of script into a *WSF*, we will talk about it no more.

To create a instance of the class:

```
Set [] = New Brackets
```

# Usage

## Methods for Function

*[].Function(strParameters, strBody) -> varFunction*

A restricted anonymous function generator.

The function it generates can only refer to the arguments & built-in functions in VBScript.

```
' Save Function to a Variable
Set IsBigger = [].Function("i, j", "Return i > j")
Msgbox IsBigger(2001, 1228)

' Recursion
Set Factorial = [].Function("i", "If i > 1 Then : Return i * Callee(i - 1) : Else : Return 1 : End If")
Msgbox Factorial(6)

' Variable-length Argument
Msgbox [].Function("", "Return ArgumentsCount")(2001, 1228)
Msgbox [].Function("", "Return Arguments(0) + Arguments(1)")(2001, 1228)
Msgbox [].Function("", "Return Join(Arguments, "" "")")(2001, 1228)
```

Argument "strParameters" doesn't support prefix "ByRef" & "ByVal". You can think of it as always "ByVal".

Keyword "Return" means save the return value, It will not really return.

```
Call [].Function("", "Return Empty : Msgbox ""Fake Return!""")()
```

*[].Lambda(strParameters, strBody, strBindings, arrBindings) -> varFunction*

Similar to *[].Function*, but with bindings.

```
' Save variable to Lambda's environment
Set Echo = [].Lambda("strArg", "WSH.Echo strArg", "WSH", Array(WScript))
Echo "Hello, world!"
' Modify binded variables

Set Counter = [].Lambda("", "Return i : i = i + 1", "i", Array(0))
Msgbox Counter() & " " & Counter() & " " & Counter()

Set Fibonacci = [].Lambda("", "Return i + j : i = i + j : j = i - j", "i, j", Array(0,1))
Fibonacci() : Fibonacci() : Fibonacci() ' -> 1 1 2
Msgbox Fibonacci() & " " & Fibonacci() & " " & Fibonacci()
```

*[].Times varSubprogram, lngTimes*

Run function many times.

```
[].Times [].Function("", "Msgbox ""Say something important three times."""), 3
```

*[].Once(varFunction) -> varDisposableFunction*

Run function just one time.

```
Set varTest = [].Once([].Function("", "Msgbox 1"))
varTest() : varTest() : varTest()
```

*[].Curry(varFunction, lngArgumentsCount)*

```
' varFunction(a,b,c), 3 -> varFunction(a)(b)(c)

Set varAdd = [].Curry([].Function("a, b, c","Return a + b + c"), 3)
Msgbox varAdd(1,2,3)
Msgbox varAdd(1,2)(3)
Msgbox varAdd(1)(2)(3)
Msgbox varAdd(1)(2,3)
```

*[].Partial(varFunction, arrArguments)*

```
' varFunction(a,b,c,d) , Array(1,2) -> varFunction(1,2,c,d)

Set varAdd = [].Function("a, b, c","Return a + b + c")
Set varAdd3 = [].Partial(varAdd, Array(1, 2))
Msgbox varAdd3(3)
```

*[].Compose(varFunc1, varFunc2, ...) -> varPipelineFunction*

```
' varF3(varF2(varF1(1))) -> varPipeline(1)

Set varF1=[].Function("x", "Return x + 10")
Set varF2=[].Function("x", "Return x * 10")
Set varF3=[].Function("x", "Return x - 10")
Msgbox [].Compose(varF1,varF2,varF3)(1)
```

## Methods for Array

*[].Range(numStart, numStop, numStep) -> arrNumber*

```
' Range(1,3,1) -> Array(1,2,3)
' Range(1,3,2) -> Array(1,3)
' Range(1,3,3) -> Array(1)
' Range(1,3,0) -> Error
' Range(1,3,-1) -> Array()

Msgbox [].Reduce([].Function("i, j", "Return i * j"), [].Range(1,6,1), 1)
```

*[].CArray(varSet) -> Array(...)*

Turn Set (like FSO.Drives) to Array.

```
Msgbox Join([].CArray(CreateObject("Scripting.FileSystemObject").Drives), " ")
```

*[].Append(varSet1, varSet2) -> arrAppended*

```
' Array(a, b), Array(c, d) -> Array(a, b, c, d)

Msgbox Join([].Append(Array(1, 2), Array(3, 4)))
```

*[].Flatten(arrNested) -> arrFlattened*

```
' Array(a, Array(b), Array(Array(c))) -> Array(a, b, c)

Msgbox Join([].Flatten(Array(1, 2, Array(3, Array(4)), Array(Array(Array(Array(5)))))), " ")
```

*Zip(varLeft, varRight) -> arrZipped*

```
' Array(a, b, c), Array(d, e, f) ->
' Array(Array(a, d), Array(b, e), Array(c, f))

[].ForEach [].Function("arrArg", "Msgbox arrArg(0) + arrArg(1)"),
	[].Zip(Array(1, 0, -1), Array(-1, 0, 1))
```

*[].Reverse(varSet) -> arrReversed*

```
Msgbox Join([].Reverse(Array(1,2,3)), " ") ' -> 3 2 1
```

## Methods for Function & Array

*[].Map(varFunction, varSet) -> arrMapped*

```
'Func, Array(item1,item2,...) -> Array(Func(item1),Func(item2),...)

Msgbox Join([].Map([].Function("i", "Return i^2"), [].Range(0,9,1)), " ")
```

*[].ForEach varSubprogram, varSet*

Similar to *[].Map*, but without return value.

```
[].ForEach [].Function("strArg", "Msgbox strArg"), [].Range(1,5,1)
```

*[].Apply(varFunction, varArguments) -> varReturn*

```
' varFunciton, Array(a, b, c, ...) -> varFunction(a, b, c, ...)

' Msgbox [].Function("i, j", "Return i+j")(12, 28)
Msgbox [].Apply([].Function("i, j", "Return i+j"), Array(12, 28))
```

*[].SpreadArguments(varFunction, varArguments) -> varReturn*

Same as *[].Apply*.

*[].GatherArguments(varFunction, varArguments) -> varReturn*

```
' varFunction(a, b, c, ...) -> varFunciton(Array(a, b, c, ...))

Msgbox [].GatherArguments([].Function("arrArg", "Return arrArg(1)"))(333, 444, 555)
```

*[].Filter(varFunction, varSet) -> arrFiltered*

Leave those items which pass the test.

```
Msgbox Join([].Filter([].Function("i", "Return i > 4"), Array(1,3,5,7)), " ")
```

*[].Reduce(varFunction, varSet, varInitialValue) -> varReduced*

Use binary arguments function to reduce an array to a single variable.

```
Msgbox [].Reduce([].Function("i, j", "Return i * j"), [].Range(1,6,1), 1)
Msgbox [].Reduce([].Function("i, j", "Return i Or j"), Array(True, True, False), False)
Msgbox [].Reduce([].Function("i, j", "Return i And j"), Array(True, True, False), True)
```

*[].Accumulate(varFunction, varSet, varInitialValue) -> varAccumulated*

Same as *[].Reduce*.

*[].Every(arrArguments, varFunction) -> boolTested*

All items meet the requirements.

```
Msgbox [].Every(Array(1, 2, 3), [].Function("i", "Return i > 0"))
Msgbox [].Every(Array(1, -1), [].Function("i", "Return i > 0"))
```

*[].Some(arrArguments, varFunction) -> boolTested*

Some items meet the requirements.

```
Msgbox [].Some(Array(1, -1), [].Function("i", "Return i > 0"))
Msgbox [].Some(Array(0, -1), [].Function("i", "Return i > 0"))
```

## Other Methods

*[].Set varVariable, varValue*

Assign *varValue* to *varVariable* whether *varValue* is an object or not.

```
[].Set objFS, CreateObject("Scripting.FileSystemObject")
[].Set PI, Atn(1) * 4
```

*[].If(boolCondition, varTrue, varFalse) -> varRet*

Just like ternary operator in other languages.

But no short-circuit, all arguments will be evaluated.

```	
Msgbox [].If(2000 > 3000, "2000гд > 3000$", "2000гд <= 3000$")
Msgbox [].If(0.1 + 0.2 = 0.3, "0.1 + 0.2 = 0.3", "0.1 + 0.2 <> 0.3")
```

*[].Assert boolCondition, strSource, strDescription*

```
[].Assert WScript.Arguments.Count = 1, _
	"WScript.Arguments", "Need a command-line argument."
```

*[].GetObject(strProgID) -> objCOM*

If strProgID available, get it directly, else create & get it.

```
[].Set objWord, [].GetObject("Word.Application")
```

*[].Unless boolPredicate, varSubprogram*

```
[].Unless True, [].Function("", "Msgbox 1")
[].Unless False, [].Function("", "Msgbox 2")
```

*Min(numA, numB) -> numMinimum*

Return the minimum value of two arguments.

```
Msgbox [].Min(-100, 100)

arrTest =  Array(1, 0, -1, -100)
Msgbox [].Reduce([].Lambda("i, j", "Return [].Min(i, j)", "[]", Array([])), arrTest, arrTest(0))
```

*Max(numA, numB) -> numMaximum*

Return the maximum value of two arguments.

```
Msgbox [].Max(-100, 100)

arrTest =  Array(10, 0, -1, -100)
Msgbox [].Reduce([].Lambda("i, j", "Return [].Max(i, j)", "[]", Array([])), arrTest, arrTest(0))
```

# References

Hungarian notation: *lng* **Long**, *str* **String**, *obj* **Object**, *arr* **Array**, *var* **Variable**, *num* **Number**.

# Examples(TODO)

# Contribute

Welcome!

Remember obey the Hungarian naming rule. Keep your code clear & meaningful.

Then simply open your pull request!

# See Also

- [SICP](https://mitpress.mit.edu/sites/default/files/sicp/index.html)

- [Python](https://www.python.org/)

- [underscore.js](https://github.com/jashkenas/underscore)

- [lodash.js](https://github.com/lodash/lodash)

- [lazy.js](https://github.com/dtao/lazy.js)

- [ramda.js](https://ramdajs.com/)
