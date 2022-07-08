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

## 用法(TODO)

注意：在VBS中没有短路求值原则，所有的参数都将被求值
apply只接收匿名函数。

# 参考(TODO)

匈牙利命名：*lng* **Long**, *str* **String**, *obj* **Object**, *arr* **Array**, *var* **Variable**, *num* **Number**.

# 示例(TODO)

# 参照

- [underscore.js](https://github.com/jashkenas/underscore)

- [lodash.js](https://github.com/lodash/lodash)

- [lazy.js](https://github.com/dtao/lazy.js)

- [ramda.js](https://ramdajs.com/)

