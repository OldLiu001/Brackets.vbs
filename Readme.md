```
  _______   ______    ________   ______   ___   ___   ______   _________  ______
/_______/\ /_____/\  /_______/\ /_____/\ /___/\/__/\ /_____/\ /________/\/_____/\
\::: _  \ \\:::_ \ \ \::: _  \ \\:::__\/ \::.\ \\ \ \\::::_\/_\__.::.__\/\::::_\/_
 \::(_)  \/_\:(_) ) )_\::(_)  \ \\:\ \  __\:: \/_) \ \\:\/___/\  \::\ \   \:\/___/\
  \::  _  \ \\: __ `\ \\:: __  \ \\:\ \/_/\\:. __  ( ( \::___\/_  \::\ \   \_::._\:\
   \::(_)  \ \\ \ `\ \ \\:.\ \  \ \\:\_\ \ \\: \ )  \ \ \:\____/\  \::\ \    /____\:\
    \_______\/ \_\/ \_\/ \__\/\__\/ \_____\/ \__\/\__\/  \_____\/   \__\/    \_____\/
```

A powerful, elegent & lightweight Functional Programming library for *VBScript*, overturn your understanding of programming in *VBScript*.

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

## Usage(TODO)

# References(TODO)

Hungarian notation: *lng* **Long**, *str* **String**, *obj* **Object**, *arr* **Array**, *var* **Variable**, *num* **Number**.

# Examples(TODO)

# See Also

- [underscore.js](https://github.com/jashkenas/underscore)

- [lodash.js](https://github.com/lodash/lodash)

- [lazy.js](https://github.com/dtao/lazy.js)

- [ramda.js](https://ramdajs.com/)
