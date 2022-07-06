Option Explicit

Public []
Set [] = New Brackets

Class Brackets
	Public Sub [Set](ByRef varVariable, ByRef varValue)
		If IsObject(varValue) Then
			Set varVariable = varValue
		Else
			varVariable = varValue
		End If
	End Sub

	Public Function [If](ByVal boolCondition, ByRef varTrue, ByRef varFalse)
		If boolCondition Then
			[Set] [If], varTrue
		Else
			[Set] [If], varFalse
		End If
	End Function

	Public Function [Function](ByVal strParameters, ByVal strBody)
		Set [Function] = New Lambda
		[Function].Init strParameters, strBody
		Set [Function] = WrapArguments(WrapFunction([Function]))
	End Function

	Public Sub ForEach(varSubprogram, varCollection)
		Dim varItem
		For Each varItem in varCollection
			Call varSubprogram(varItem)
		Next
	End Sub

	Public Function Range(lngStart, lngStop, lngStep)
		Assert (lngStop - lngStart) / lngStep >= 0, _
			"<Function Range>: Invaild parameter(s)."
		
		Dim arrRange(), lngCounter, lngPointer
		ReDim arrRange(Fix((lngStop - lngStart) / lngStep))
		
		lngPointer = 0
		For lngCounter = lngStart To lngStop Step lngStep
			arrRange(lngPointer) = lngCounter
			lngPointer = lngPointer + 1
		Next
		Range = arrRange
	End Function

	Public Function Map(varFunction, varCollection)
		Dim arrMap(), varItem, lngPointer
		ReDim arrMap(1)
		lngPointer = 0
		For Each varItem in varCollection
			If UBound(arrMap) < lngPointer Then
				ReDim Preserve arrMap(UBound(arrMap) * 2)
			End If
			arrMap(lngPointer) = varFunction(varItem)
			lngPointer = lngPointer + 1
		Next
		ReDim Preserve arrMap(lngPointer)
		Map = arrMap
	End Function

	Public Sub Assert(ByVal boolCondition, ByRef strMessage)
		If Not boolCondition Then
			Err.Raise vbObjectError, "", strMessage
		End If
	End Sub

End Class

Class Lambda
	Private varReturnValue
	Private Sub Return(varValue)
		If IsObject(varValue) Then
			Set varReturnValue = varValue
		Else
			varReturnValue = varValue
		End If
	End Sub
	Public Property Get ReturnValue()
		If IsObject(varReturnValue) Then
			Set ReturnValue = varReturnValue
		Else
			ReturnValue = varReturnValue
		End If
	End Property

	Private strCodeBody
	Public Sub Init(strParameters, strBody)
		Dim lngParameterCounter, strParameter
		strCodeBody = ""
		lngParameterCounter = 0
		For Each strParameter in Split(Replace(strParameters, " ", ""), ",")
			strCodeBody = strCodeBody & _
				"Dim " & strParameter & vbNewLine & _
				"If IsObject(arrArguments.[" & lngParameterCounter & "]) Then" & vbNewLine & _
				"	Set " & strParameter & " = arrArguments.[" & lngParameterCounter & "]" & vbNewLine & _
				"Else" & vbNewLine & _
				"	" & strParameter & " = arrArguments.[" & lngParameterCounter & "]" & vbNewLine & _
				"End If" & vbNewLine
			lngParameterCounter = lngParameterCounter + 1
		Next
		strCodeBody = strCodeBody & strBody
	End Sub

	Public Sub Apply(arrArguments)
		Execute strCodeBody
	End Sub
End Class

''Set lam = [].Function("a,b", "msgbox a" & vbnewline & "msgbox b: return a+b")
''lam.apply(array(3,4))
''msgbox "ret:" & lam.ReturnValue
' function z2()
' msgbox 1
' end function
'vbscript ������д
' function [sub](z)
' 	msgbox z
' end function
' sub 1
' end sub
' msgbox [].If(true,2,1)
' msgbox fix(1)
' function [fix](z)
' [fix] =fix(z) + 1
' end function
' msgbox [string]( 10,"*")

' function z2
' 	msgbox 20
' end function
' function [z2] '�Ḳ���ϱߵ�
' 	msgbox 10
' end function

''class z
''sub main
''dim s
''s = 3
''msgbox 1
''executeglobal "public sub show(x):wsh.echo s:end sub" '����global
''show "z"
''msgbox 1
''msgbox 2
''getref("z2")()
''msgbox typename(getref("z2"))
''msgbox typename(eval("getref(""show"")"))
''[].ForEach getref("show"),CreateObject("Scripting.FileSystemObject").GetFolder(".").Files
''[].ForEach getref("show"),[].Map(getref("m"), CreateObject("Scripting.FileSystemObject").GetFolder(".").Files)
''end sub

''function m(i)
''	m = i.name
''end function
''end class
''dim s 
''set s = new z
''s.main
' Function ShowFolderList(folderspec)
'	Dim fso, f, f1, fc, s
'	Set fso = CreateObject("Scripting.FileSystemObject")
'	Set f = fso.GetFolder(folderspec)
'	Set fc = f.Files
'	msgbox len(fc)
'	For Each f1 in fc
'	   s = s & f1.name 
'	   s = s & "<BR>"
'	Next
'	ShowFolderList = s
' End Function

' msgbox ShowFolderList(".")



' VBScript��̬�������Զ���Ĺ�����(DynamicObject)
' ��������ʵ�ִ��빩��Ҳο��� 
' '
' ' ASP/VBScript Dynamic Object Generator
' ' Author: WangYe
' ' For more information please visit
' '	 http://wangye.org/
' ' This code is distributed under the BSD license
' '
' Const PROPERTY_ACCESS_READONLY = 1
' Const PROPERTY_ACCESS_WRITEONLY = -1
' Const PROPERTY_ACCESS_ALL = 0
' ?
' Class DynamicObject
'	 Private m_objProperties
'	 Private m_strName
' ?
'	 Private Sub Class_Initialize()
'	 Set m_objProperties = CreateObject("Scripting.Dictionary")
'	 m_strName = "AnonymousObject"
'	 End Sub
' ?
'	 Private Sub Class_Terminate()
'	 If Not IsObject(m_objProperties) Then
'		 m_objProperties.RemoveAll
'	 End If
'	 Set m_objProperties = Nothing
'	 End Sub
' ?
'	 Public Sub setClassName(strName)
'	 m_strName = strName
'	 End Sub
' ?
'	 Public Sub add(key, value, access)
'	 m_objProperties.Add key, Array(value, access)
'	 End Sub
' ?
'	 Public Sub setValue(key, value, access)
'	 If m_objProperties.Exists(key) Then
'		 m_objProperties.Item(key)(0) = value
'		 m_objProperties.Item(key)(1) = access
'	 Else
'		 add key,value,access
'	 End If
'	 End Sub
' ?
'	 Private Function getReadOnlyCode(strKey)
'	 Dim strPrivateName, strPublicGetName
'	 strPrivateName = "m_var" & strKey
'	 strPublicGetName = "get" & strKey
'	 getReadOnlyCode = _
'		 "Public Function " & strPublicGetName & "() :" & _
'		 strPublicGetName & "=" & strPrivateName & " : " & _
'		 "End Function : Public Property Get " & strKey & _
'		 " : " & strKey & "=" & strPrivateName & " : End Property : "
'	 End Function
' ?
'	 Private Function getWriteOnlyCode(strKey)
'	 Dim pstr
'	 Dim strPrivateName, strPublicSetName, strParamName
'	 strPrivateName = "m_var" & strKey
'	 strPublicSetName = "set" & strKey
'	 strParamName = "param" & strKey
'	 getWriteOnlyCode = _
'		 "Public Sub " & strPublicSetName & "(" & strParamName & ") :" & _
'		 strPrivateName & "=" & strParamName & " : " & _
'		 "End Sub : Public Property Let " & strKey & "(" & strParamName & ")" & _
'		 " : " & strPrivateName & "=" & strParamName & " : End Property : "
'	 End Function
' ?
'	 Private Function parse()
'	 Dim i, Keys, Items
'	 Keys = m_objProperties.Keys
'	 Items = m_objProperties.Items
' ?
'	 Dim init, pstr
'	 init = ""
'	 pstr = ""
'	 parse = "Class " & m_strName & " :" & _
'		 "Private Sub Class_Initialize() : "
' ?
'	 Dim strPrivateName
'	 For i = 0 To m_objProperties.Count - 1
'		 strPrivateName = "m_var" & Keys(i)
'		 init = init & strPrivateName & "=""" & _
'		 Replace(CStr(Items(i)(0)), """", """""") & """:"
'		 pstr = pstr & "Private " & strPrivateName & " : "
'		 If CInt(Items(i)(1)) > 0 Then ' ReadOnly
'		 pstr = pstr & getReadOnlyCode(Keys(i))
'		 ElseIf CInt(Items(i)(1)) < 0 Then ' WriteOnly
'		 pstr = pstr & getWriteOnlyCode(Keys(i))
'		 Else ' AccessAll
'		 pstr = pstr & getReadOnlyCode(Keys(i)) & _
'			 getWriteOnlyCode(Keys(i))
'		 End If
'	 Next
'	 parse = parse & init & "End Sub : " &  pstr & "End Class"
'	 End Function
' ?
'	 Public Function getObject()
'	 Call Execute(parse)
'	 Set getObject = Eval("New " & m_strName)
'	 End Function
' ?
'	 Public Sub invokeObject(ByRef obj)
'	 Call Execute(parse)
'	 Set obj = Eval("New " & m_strName)
'	 End Sub
' End Class
' �������Զ���ֱ��ṩ��Propertyֱ�ӷ���ģʽ��set����get��������ģʽ����Ȼ�һ��ṩ������Ȩ�޿��ƣ���add������ʹ�ã��ֱ���PROPERTY_ACCESS_READONLY������ֻ������PROPERTY_ACCESS_WRITEONLY������ֻд����PROPERTY_ACCESS_ALL�����Զ�д�������������������ʹ�ã�һ�����ӣ���
' Dim DynObj
' Set DynObj = New DynamicObject
'	 DynObj.add "Name", "WangYe", PROPERTY_ACCESS_READONLY
'	 DynObj.add "HomePage", "http://wangye.org", PROPERTY_ACCESS_READONLY
'	 DynObj.add "Job", "Programmer", PROPERTY_ACCESS_ALL
'	 '
'	 ' ���û��setClassName��
'	 ' �´����Ķ��󽫻��Զ�����ΪAnonymousObject
'	 ' �����������������󣬾ͱ���ָ������
'	 ' ����Ϳ�������������ظ����쳣
'	 DynObj.setClassName "User"
' ?
'	 Dim User
'	 Set User = DynObj.GetObject()
'	 ' ���� DynObj.invokeObject User
'	objfile.Write User.Name
'	 ' objfile.Write User.getName()
' objfile.Write User.HomePage
'	 ' objfile.Write User.getHomePage()
' objfile.Write User.Job
'	 ' objfile.Write User.getJob()

'	 ' �ı�����ֵ
'	 User.Job = "Engineer"
'	 ' User.setJob "Engineer"

'	 Response.Write User.getJob()
'	 Set User = Nothing
' ?
' Set DynObj = Nothing
' ��ԭ��ܼ򵥣�����ͨ��������Key-Value��̬����VBS Class�ű����룬Ȼ�����Executeִ���Ա��ڽ���δ�����뵽�������������У������ͨ��Eval�½��������
' ���ˣ��ͽ��ܵ��������ҿ��ܻ���½������һЩClassic ASP����ؼ��ɴ��롣
