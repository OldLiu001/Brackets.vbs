msgbox eval("12")


' Dim []
' Set [] = CreateObject("Brackets")
' ' varFunction(a,b,c), 3 -> varFunction(a)(b)(c)

' Set varAdd = [].Curry([].Function("a, b, c","Return a + b + c"), 3)
' Msgbox varAdd(1,2,3)
' Msgbox varAdd(1,2)(3)
' Msgbox varAdd(1)(2)(3)
' Msgbox varAdd(1)(2,3)
