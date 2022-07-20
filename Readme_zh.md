```
  _______   ______    ________   ______   ___   ___   ______   _________  ______
/_______/\ /_____/\  /_______/\ /_____/\ /___/\/__/\ /_____/\ /________/\/_____/\
\::: _  \ \\:::_ \ \ \::: _  \ \\:::__\/ \::.\ \\ \ \\::::_\/_\__.::.__\/\::::_\/_
 \::(_)  \/_\:(_) ) )_\::(_)  \ \\:\ \  __\:: \/_) \ \\:\/___/\  \::\ \   \:\/___/\
  \::  _  \ \\: __ `\ \\:: __  \ \\:\ \/_/\\:. __  ( ( \::___\/_  \::\ \   \_::._\:\
   \::(_)  \ \\ \ `\ \ \\:.\ \  \ \\:\_\ \ \\: \ )  \ \ \:\____/\  \::\ \    /____\:\
    \_______\/ \_\/ \_\/ \__\/\__\/ \_____\/ \__\/\__\/  \_____\/   \__\/    \_____\/
```

强大、优雅、简洁的 *VBS* 函数式编程类库，颠覆你对 *VBS* 编程的理解！

首创 *VBS* 匿名函数支持，按需添加环境。此外还一并封装了几十个常用、通用、易用的方法。

通过函数的封装和嵌套，将复杂的问题抽丝剥茧，展现其中的变换本质。原来，编程和数学竟能如此相像！

感受函数式编程的魅力吧！

# 浏览

- [中文](Readme_zh.md)
- [英文](Readme.md)

# 开始

## 环境要求

- 视窗操作系统

## 安装

以**管理员权限**运行以下命令：

```
git clone https://github.com/OldLiu001/Brackets.vbs.git
cd Brackets.vbs
regsvr32 Brackets.wsc
```

**警告：不要使用右键菜单注册 *Brackets.wsc* 。**

使用下列代码创建类的实例：

```
Set [] = CreateObject("Brackets")
```

## 便携

制作便携版本后，其他用户无需进行上述的安装操作即可使用您的脚本。

复制脚本 *Brackets.vbs* 和 *Brackets.js* 到您脚本所在的文件夹下。

假设您的脚本的文件名为 *MyScript.vbs* ，使用如下的代码模板：

*Template.wsf*

```
<job id="MyScript">
	<script language="JScript" src="Brackets.js"/>
	<script language="VBScript" src="Brackets.vbs"/>
	<script language="VBScript" src="MyScript.vbs"/>
</job>
```

将其放置到您脚本所在的文件夹下。

或将脚本和类库都嵌入单个 *WSF* 中：

*Template_Embedded.wsf*

```
<job id="MyScript">
	<script language="JScript">
		// 此处写 "Brackets.js" 的内容
	</script>
	<script language="VBScript">
		' 此处写 "Brackets.vbs" 的内容
	</script>
	<script language="VBScript">
		' 此处写 "MyScript.vbs" 的内容
	</script>
</job>
```

当然，您可以只将必要的部分嵌入 *WSF* 中，此处不再赘述。

创建类的实例：

```
Set [] = New Brackets
```

# 用法

## 函数相关方法

`[].Function(strParameters, strBody) -> varFunction`

受限的匿名函数生成器。

它生成的函数只能引用参数和 VBScript 中的内置函数。

```
' 赋值给变量
Set IsBigger = [].Function("i, j", "Return i > j")
Msgbox IsBigger(2001, 1228)

' 递归
Set Factorial = [].Function("i", "If i > 1 Then : Return i * Callee(i - 1) : Else : Return 1 : End If")
Msgbox Factorial(6)

' 变长参数
Msgbox [].Function("", "Return ArgumentsCount")(2001, 1228)
Msgbox [].Function("", "Return Arguments(0) + Arguments(1)")(2001, 1228)
Msgbox [].Function("", "Return Join(Arguments, "" "")")(2001, 1228)
```

参数 `strParameters` 不支持前缀 `ByRef` 以及 `ByVal`。你可以把它想象成永远 `ByVal`。

关键字 `Return` 表示保存返回值，它不会真正返回。

```
Call [].Function("", "Return Empty : Msgbox ""Fake Return!""")()
```

---

`[].Lambda(strParameters, strBody, strBindings, arrBindings) -> varFunction`

与 `[].Function` 类似，但可以绑定环境。

```
' 给Lambda绑定环境
Set Echo = [].Lambda("strArg", "WSH.Echo strArg", "WSH", Array(WScript))
Echo "Hello, world!"

' 修改绑定的变量

Set Counter = [].Lambda("", "Return i : i = i + 1", "i", Array(0))
Msgbox Counter() & " " & Counter() & " " & Counter()

Set Fibonacci = [].Lambda("", "Return i + j : i = i + j : j = i - j", "i, j", Array(0,1))
Fibonacci() : Fibonacci() : Fibonacci() ' -> 1 1 2
Msgbox Fibonacci() & " " & Fibonacci() & " " & Fibonacci()
```

---

`[].Times varSubprogram, lngTimes`

运行过程若干次。

```
[].Times [].Function("", "Msgbox ""Say something important three times."""), 3
```

---

`[].Once(varFunction) -> varDisposableFunction`

将函数包裹为只可运行一次。

```
Set varTest = [].Once([].Function("", "Msgbox 1"))
varTest() : varTest() : varTest()
```

---

`[].Curry(varFunction, lngArgumentsCount)`

打散参数以支持携带部分参数。

这个过程一般被称作“柯里化”。

```
' varFunction(a,b,c), 3 -> varFunction(a)(b)(c)

Set varAdd = [].Curry([].Function("a, b, c","Return a + b + c"), 3)
Msgbox varAdd(1,2,3)
Msgbox varAdd(1,2)(3)
Msgbox varAdd(1)(2)(3)
Msgbox varAdd(1)(2,3)
```

---

`[].Partial(varFunction, arrArguments)`

返回携带左半部分参数的函数。

这个过程一般被称作“偏应用”。

```
' varFunction(a,b,c,d) , Array(1,2) -> varFunction(1,2,c,d)

Set varAdd = [].Function("a, b, c","Return a + b + c")
Set varAdd3 = [].Partial(varAdd, Array(1, 2))
Msgbox varAdd3(3)
```

---

`[].Compose(varFunc1, varFunc2, ...) -> varPipelineFunction`

将若干单参函数打包成流水线函数。

```
' varF3(varF2(varF1(1))) -> varPipeline(1)

Set varF1=[].Function("x", "Return x + 10")
Set varF2=[].Function("x", "Return x * 10")
Set varF3=[].Function("x", "Return x - 10")
Msgbox [].Compose(varF1,varF2,varF3)(1)
```

## Methods for Array

`[].Range(numStart, numStop, numStep) -> arrNumber`

返回由若干有序、离散、等距的数组成的数组。

```
' Range(1,3,1) -> Array(1,2,3)
' Range(1,3,2) -> Array(1,3)
' Range(1,3,3) -> Array(1)
' Range(1,3,0) -> Error
' Range(1,3,-1) -> Array()

Msgbox [].Reduce([].Function("i, j", "Return i * j"), [].Range(1,6,1), 1)
```

---

`[].CArray(varSet) -> Array(...)`

将类似 `FSO.Drives` 的集合转换成数组。

```
Msgbox Join([].CArray(CreateObject("Scripting.FileSystemObject").Drives), " ")
```

---

`[].Append(varSet1, varSet2) -> arrAppended`

合并两个数组/集合。

```
' Array(a, b), Array(c, d) -> Array(a, b, c, d)

Msgbox Join([].Append(Array(1, 2), Array(3, 4)))
```

---

`[].Flatten(arrNested) -> arrFlattened`

将数组打平。

```
' Array(a, Array(b), Array(Array(c))) -> Array(a, b, c)

Msgbox Join([].Flatten(Array(1, 2, Array(3, Array(4)), Array(Array(Array(Array(5)))))), " ")
```

---

`[].Zip(varLeft, varRight) -> arrZipped`

将两个数组/集合内对应元素打包。

```
' Array(a, b, c), Array(d, e, f) ->
' Array(Array(a, d), Array(b, e), Array(c, f))

[].ForEach [].Function("arrArg", "Msgbox arrArg(0) + arrArg(1)"),
	[].Zip(Array(1, 0, -1), Array(-1, 0, 1))
```

---

`[].Reverse(varSet) -> arrReversed`

返回逆序后的数组。

```
Msgbox Join([].Reverse(Array(1,2,3)), " ") ' -> 3 2 1
```

## Methods for Function & Array

`[].Map(varFunction, varSet) -> arrMapped`

将象按顺序打包。

```
'Func, Array(item1,item2,...) -> Array(Func(item1),Func(item2),...)

Msgbox Join([].Map([].Function("i", "Return i^2"), [].Range(0,9,1)), " ")
```

---

`[].ForEach varSubprogram, varSet`

与 `[].Map` 类似，但不返回值。

```
[].ForEach [].Function("strArg", "Msgbox strArg"), [].Range(1,5,1)
```

---

`[].Apply(varFunction, varArguments) -> varReturn`

数组元素作为函数参数调用函数。

```
' varFunciton, Array(a, b, c, ...) -> varFunction(a, b, c, ...)

' Msgbox [].Function("i, j", "Return i+j")(12, 28)
Msgbox [].Apply([].Function("i, j", "Return i+j"), Array(12, 28))
```

---

`[].SpreadArguments(varFunction, varArguments) -> varReturn`

与 `[].Apply` 同。

---

`[].GatherArguments(varFunction, varArguments) -> varReturn`

将外层函数参数打包成数组传递给内层函数。

```
' varFunction(a, b, c, ...) -> varFunciton(Array(a, b, c, ...))

Msgbox [].GatherArguments([].Function("arrArg", "Return arrArg(1)"))(333, 444, 555)
```

---

`[].Filter(varFunction, varSet) -> arrFiltered`

将满足条件的数组内元素留下。

```
Msgbox Join([].Filter([].Function("i", "Return i > 4"), Array(1,3,5,7)), " ")
```

---

`[].Reduce(varFunction, varSet, varInitialValue) -> varReduced`

使用二元函数归约/累积数组至一个值。

```
Msgbox [].Reduce([].Function("i, j", "Return i * j"), [].Range(1,6,1), 1)
Msgbox [].Reduce([].Function("i, j", "Return i Or j"), Array(True, True, False), False)
Msgbox [].Reduce([].Function("i, j", "Return i And j"), Array(True, True, False), True)
```

---

`[].Accumulate(varFunction, varSet, varInitialValue) -> varAccumulated`

与 `[].Reduce` 同。

---

`[].Every(arrArguments, varFunction) -> boolTested`

判断是否所有宿主内的元素满足要求。

```
Msgbox [].Every(Array(1, 2, 3), [].Function("i", "Return i > 0"))
Msgbox [].Every(Array(1, -1), [].Function("i", "Return i > 0"))
```

---

`[].Some(arrArguments, varFunction) -> boolTested`

判断是否有宿主内的元素满足要求。

```
Msgbox [].Some(Array(1, -1), [].Function("i", "Return i > 0"))
Msgbox [].Some(Array(0, -1), [].Function("i", "Return i > 0"))
```

## Other Methods

`[].Set varVariable, varValue`

统一 *VBS* 内的赋值方式。

```
[].Set objFS, CreateObject("Scripting.FileSystemObject")
[].Set PI, Atn(1) * 4
```

---

`[].If(boolCondition, varTrue, varFalse) -> varRet`

类似其它语言中的三目运算符，但没有短路求值。

```	
Msgbox [].If(2000 > 3000, "2000￥ > 3000$", "2000￥ <= 3000$")
Msgbox [].If(0.1 + 0.2 = 0.3, "0.1 + 0.2 = 0.3", "0.1 + 0.2 <> 0.3")
```

---

`[].Assert boolCondition, strSource, strDescription`

断言条件满足，否则报错。

```
[].Assert WScript.Arguments.Count = 1, _
	"WScript.Arguments", "Need a command-line argument."
```

---

`[].GetObject(strProgID) -> objCOM`

若已存在 *COM* 对象，则直接得到，否则创建一个。

```
[].Set objWord, [].GetObject("Word.Application")
```

---

`[].Unless boolPredicate, varSubprogram`

若谓词不满足，则执行子程序。

```
[].Unless True, [].Function("", "Msgbox 1")
[].Unless False, [].Function("", "Msgbox 2")
```

---

`[].Min(numA, numB) -> numMinimum`

取双参中的最小值。

```
Msgbox [].Min(-100, 100)

arrTest =  Array(1, 0, -1, -100)
Msgbox [].Reduce([].Lambda("i, j", "Return [].Min(i, j)", "[]", Array([])), arrTest, arrTest(0))
```

---

`[].Max(numA, numB) -> numMaximum`

取双参中的最大值。

```
Msgbox [].Max(-100, 100)

arrTest =  Array(10, 0, -1, -100)
Msgbox [].Reduce([].Lambda("i, j", "Return [].Max(i, j)", "[]", Array([])), arrTest, arrTest(0))
```

# 参考

匈牙利命名：*lng* **Long**, *str* **String**, *obj* **Object**, *arr* **Array**, *var* **Variable**, *num* **Number**.

# 示例(TODO)

# 贡献

如果您发现任何 *BUG* ，请提交一个 *Issue* 。

欢迎贡献代码！

记得遵守匈牙利命名规则，然后提交 *Pull Request* 即可！

# 参照

- [SICP](https://mitpress.mit.edu/sites/default/files/sicp/index.html)

- [Python](https://www.python.org/)

- [underscore.js](https://github.com/jashkenas/underscore)

- [lodash.js](https://github.com/lodash/lodash)

- [lazy.js](https://github.com/dtao/lazy.js)

- [ramda.js](https://ramdajs.com/)
