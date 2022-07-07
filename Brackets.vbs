Option Explicit

Class Brackets
	Public Sub [Set](ByRef varVariable, ByRef varValue)
		' Unify the way of assignment in VBScript.
		If IsObject(varValue) Then
			Set varVariable = varValue
		Else
			varVariable = varValue
		End If
	End Sub

	Public Function [If](ByVal boolCondition, ByRef varTrue, ByRef varFalse)
		' Just like ternary operator in other languages.
		' But no short-circuit, all arguments will be evaluated.
		If boolCondition Then
			[Set] [If], varTrue
		Else
			[Set] [If], varFalse
		End If
	End Function

	Public Function [Function](ByVal strParameters, ByVal strBody)
		' A restricted anonymous function generator.
		' The function it generates can only refer to the arguments & built-in functions in VBScript.

		' Argument "strParameters" doesn't support prefix like "ByRef" & "ByVal".
		' Keyword "Return" will help you save return value, but It will not halt the running of function.
		Set [Function] = New AnonymousFunction
		[Function].Init strParameters, strBody
		Set [Function] = [_].WrapArguments([_].WrapFunction([Function]))
	End Function

	Public Function Lambda(ByVal strParameters, ByVal strBody) 'alias of Function
		Set Lambda = [Function](strParameters, strBody)
	End Function

	Public Sub Assert(ByVal boolCondition, ByVal strSource, ByVal strDescription)
		If Not boolCondition Then
			Err.Raise vbObjectError, strSource, strDescription
		End If
	End Sub

	Public Function Range(ByVal numStart, ByVal numStop, ByVal numStep)
		' Range(1,2,1) -> Array(1,2)
		' Range(1,2,2) -> Array(1)
		' Range(1,2,0) -> Error
		' Range(1,2,-1) -> Array()

		Assert numStep <> 0, "<Function> Range", "Step length must not be zero."
		
		If (numStop - numStart) / numStep >= 0 Then
			Dim arrRange(), numCounter, lngPointer
			ReDim arrRange(Fix((numStop - numStart) / numStep))
		
			lngPointer = 0
			For numCounter = numStart To numStop Step numStep
				arrRange(lngPointer) = numCounter
				lngPointer = lngPointer + 1
			Next
			Range = arrRange
		Else
			Range = Array()
		End If
	End Function

	Public Function Map(ByVal varFunction, ByRef varSet)
		' Func, Array(item1,item2,...) -> Array(Func(item1),Func(item2),...)

		Dim arrMap(), varItem, lngPointer
		ReDim arrMap(1)
		lngPointer = -1
		For Each varItem In varSet
			lngPointer = lngPointer + 1
			If UBound(arrMap) < lngPointer Then
				ReDim Preserve arrMap(UBound(arrMap) * 2)
			End If
			[Set] arrMap(lngPointer), varFunction(varItem)
		Next
		ReDim Preserve arrMap(lngPointer)
		Map = arrMap
	End Function

	Public Sub ForEach(ByVal varSubprogram, ByRef varSet)
		' A special Map which don't need return value.
		' So there has a simpler but slower implement:
		' Map varSubprogram, varSet

		Dim varItem
		For Each varItem In varSet
			Call varSubprogram(varItem)
		Next
	End Sub

	Public Function Apply(ByVal varFunction, ByRef varArguments)
		' Support only Anonymous Function
		[Set] Apply, [_].Apply(varFunction, CArray(varArguments))
	End Function

	Public Function CArray(ByRef varSet) 'Set & Array -> Array
		' A simpler & slower implement:
		' CArray = Map([Function]("x", "Return x"), varSet)

		If IsArray(varSet) Then 'just for efficiency
			CArray = varSet
		Else ' Deal with sets, e.g. "FSO.Drives"
			' You can expand this for higher efficiency.
			CArray = Map([Function]("x", "Return x"), varSet)
		End If
	End Function

	Public Function Filter(ByVal varFunction, ByRef varSet)
		Dim lngPointer, arrFiltered(), varItem
		ReDim arrFiltered(1)
		lngPointer = -1
		For Each varItem In varSet
			If varFunction(varItem) Then
				lngPointer = lngPointer + 1
				ReDim Preserve arrFiltered( _
					[If](lngPointer > UBound(arrFiltered), _
						UBound(arrFiltered) * 2, _
						UBound(arrFiltered)))
				[Set] arrFiltered(lngPointer), varItem
			End If
		Next
		ReDim Preserve arrFiltered(lngPointer)
		Filter = arrFiltered
	End Function

	Public Function Accumulate(varFunction, varSet, varInitialValue)

	End Function

	Public Function Reduce()
	End Function

	Public Function [GetObject](strProgID)
		' If strProgID available, get it directly, else create & get it.
		On Error Resume Next
		Set objCOM = GetObject(, strProgID)
		If Err.Number <> 0 Then
			Err.Clear
			Set objCOM = CreateObject(strProgID)
			Assume Err.Number = 0, "<Function> GetObject", _
				"Create COM object """ & strProgID & """ failed."
		End If
		On Error Goto 0
	End Function
End Class

'lazy 
' current next

Class AnonymousFunction
	Public Sub [Set](ByRef varVariable, ByRef varValue)
		If IsObject(varValue) Then
			Set varVariable = varValue
		Else
			varVariable = varValue
		End If
	End Sub

	Private varReturnValue
	Private Sub Return(varValue)
		[Set] varReturnValue, varValue
	End Sub
	Public Property Get ReturnValue()
		[Set] ReturnValue, varReturnValue
	End Property

	Private strCodeBody
	Public Sub Init(strParameters, strBody)
		Dim lngParameterCounter, strParameter
		strCodeBody = ""
		lngParameterCounter = 0
		For Each strParameter in Split(Replace(strParameters, " ", ""), ",")
			strCodeBody = strCodeBody & _
				"Dim " & strParameter & vbNewLine & _
				"[Set] " & strParameter & ", " & _
				"objArguments.[" & lngParameterCounter & "]" & vbNewLine
			lngParameterCounter = lngParameterCounter + 1
		Next
		strCodeBody = strCodeBody & strBody
	End Sub

	Public Sub Apply(objArguments)
		Execute strCodeBody
	End Sub
End Class

