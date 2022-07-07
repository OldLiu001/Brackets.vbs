''Option Explicit
set fso=createobject("Scripting.FileSystemObject")
''msgbox typename(fso.Drives)
''msgbox isarray(fso.drives)

Dim []
Set [] = CreateObject("Brackets")
''dim a
''[].Set a,-3
''msgbox [].If(a>0,"pos", "neg")
''call [].Function("wsh,z","wsh.echo z")(wsh,1)
''[].Assert False, "ez", "edes"
[].foreach [].function("x","msgbox x"),array(1,2,3)
''[].foreach [].function("x","msgbox x"),array(1,2,3)
''wsh.echo join([].range(1,10,1),",")
''call [].function("x,wsh,[]","[].foreach [].function(""i,wsh"", ""wsh.echo i"").echo x(2)")([].range(1,10,0.5),wsh,[])
''wsh.echo join([].range(1,10,-.5),",")
''wsh.echo join([].range(1,10,0),",")
msgbox join([].map([].function("i","return i^2"), array(2,3)), " ")

'sub a(s)
'	msgbox "sub"
'end sub

'msgbox typename(getref("a")(1))
