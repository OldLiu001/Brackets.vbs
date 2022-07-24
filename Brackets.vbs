Option Explicit

Class Brackets
	Private []

	Public Compose, GatherArguments
	Private Sub Class_Initialize
		Set [] = New Brackets

		' varFunction(a, b, c, ...) -> varFunciton(Array(a, b, c, ...))
		[Set] GatherArguments, Lambda("", _
			"If IsEmpty(varFunction) Then" & vbNewLine & _
			"	Set varFunction = Arguments(0)" & vbNewLine & _
			"	Return Callee" & vbNewLine & _
			"Else" & vbNewLine & _
			"	Return varFunction(Arguments)" & vbNewLine & _
			"End If", _
			"varFunction", Array(Empty))

		''TODO FIX
		[Set] Compose, Lambda("", _
			"If IsEmpty(arrFunctions) Then" & vbNewLine & _
			"	arrFunctions = Arguments" & vbNewLine & _
			"	Return Callee" & vbNewLine & _
			"Else" & vbNewLine & _
			"	Return [].Reduce(" & _
			"		[].Function(" & _
			"			""varData, varFunction"", " & _
			"			""Return varFunction(varData)""), " & _
			"		[].Append(Arguments, arrFunctions))"& vbNewLine & _
			"End If", _
			"arrFunctions", Array(Empty))
	End Sub

	Public Sub [Set](ByRef varVariable, varValue)
		' Unify the way of assignment in VBScript.

		If IsObject(varValue) Then
			Set varVariable = varValue
		Else
			varVariable = varValue
		End If
	End Sub

	Public Function [If](boolCondition, varTrue, varFalse)
		' Just like ternary operator in other languages.
		' But no short-circuit, all arguments will be evaluated.

		[Set] [If], Array(varTrue,varFalse)((boolCondition) + 1)
	End Function

	Public Function Lambda(strParameters, strBody, strBindings, arrBindings)
		' Argument "strParameters" & "strBindings" doesn't support prefix "ByRef" & "ByVal".
		' You can think of it as always "ByVal".
		' Keyword "Return" means save the return value, It will not really return.

		Set Lambda = New AnonymousFunction
		Lambda.Init strParameters, strBody, strBindings, arrBindings
		Set Lambda = [_].GatherArguments(Lambda)
	End Function

	Public Function [Function](strParameters, strBody)
		' A restricted anonymous function generator.
		' The function it generates can only refer to the arguments & built-in functions in VBScript.

		Set [Function] = Lambda(strParameters, strBody, "", Empty)
	End Function

	Public Sub Assert(boolCondition, strSource, strDescription)
		' !boolCondition -> Error

		If Not boolCondition Then
			Err.Raise vbObjectError, strSource, strDescription
		End If
	End Sub

	Public Function Range(numStart, numStop, numStep)
		' Range(1,3,1) -> Array(1,2,3)
		' Range(1,3,9) -> Array(1)
		' Range(1,3,0) -> Error
		' Range(1,3,-1) -> Array()

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

	Public Function Map(varFunction, varSet)
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

	Public Sub ForEach(varSubprogram, varSet)
		' A special Map which don't need return value.
		' So there has a simpler but slower implement:
		' Map varSubprogram, varSet

		Dim varItem
		For Each varItem In varSet
			Call varSubprogram(varItem)
		Next
	End Sub

	Public Function Apply(varFunction, varArguments)
		' varFunciton, Array(a, b, c, ...) -> varFunction(a, b, c, ...)
		' Support only Anonymous Function

		[Set] Apply, [_].SpreadArguments(varFunction, CArray(varArguments))
	End Function

	Public Function SpreadArguments(varFunction, varArguments)
		[Set] SpreadArguments, Apply(varFunction, varArguments)
	End Function

	Public Function CArray(varSet) 'Set & Array -> Array
		' A simpler & slower implement:
		' CArray = Map([Function]("x", "Return x"), varSet)

		If IsArray(varSet) Then 'just for efficiency
			CArray = varSet
		Else ' Deal with sets, e.g. "FSO.Drives"
			' You can expand this for higher efficiency.
			CArray = Map([Function]("x", "Return x"), varSet)
		End If
	End Function

	Public Function Filter(varFunction, varSet)
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

	Public Function Accumulate(varFunction, varSet)
		Dim varItem
		Dim boolFirst
		boolFirst = True
		For Each varItem In varSet
			If boolFirst Then
				[Set] Accumulate, varItem
				boolFirst = False
			Else
				[Set] Accumulate, varFunction(Accumulate, varItem)
			End If
		Next
	End Function

	Public Function Reduce(varFunction, varSet)
		' Same as Accumulate(), just a alias.

		[Set] Reduce, Accumulate(varFunction, varSet)
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

	Public Function Append(varSet1, varSet2)
		' Array(1,2), Array(3) -> Array(1,2,3)
		Dim arrCombined(), lngPtr, varItem, varSet
		ReDim arrCombined(1)
		
		lngPtr = -1
		For Each varSet In Array(varSet1, varSet2)
			For Each varItem In varSet
				lngPtr = lngPtr + 1
				ReDim Preserve arrCombined( _
					[If](lngPtr > UBound(arrCombined), _
						UBound(arrCombined) * 2, _
						UBound(arrCombined)))
				[Set] arrCombined(lngPtr), varItem
			Next
		Next

		ReDim Preserve arrCombined(lngPtr)
		Append = arrCombined
	End Function

	Public Function Flatten(arrNested)
		If IsArray(arrNested) Then
			Flatten = Array()
			Dim varItem
			For Each varItem In arrNested
				Flatten = Append(Flatten, Flatten(varItem))
			Next
		Else
			Flatten = Array(arrNested)
		End If
	End Function

	Public Sub Unless(boolPredicate, varSubprogram)
		If Not boolPredicate Then
			Call varSubprogram()
		End If
	End Sub

	Public Sub Times(varSubprogram, lngTimes)
		Dim lngCounter
		For lngCounter = 1 To lngTimes
			Call varSubprogram()
		Next
	End Sub

	Public Function Every(arrArguments, varFunction)
		Every = Accumulate( _
			[Function]("boolLeft, boolRight", "Return boolLeft And boolRight"), _
			Map(varFunction, arrArguments))
	End Function

	Public Function Some(arrArguments, varFunction)
		Some = Accumulate( _
			[Function]("boolLeft, boolRight", "Return boolLeft Or boolRight"), _
			Map(varFunction, arrArguments))
	End Function

	Public Function Once(varFunction)
		Set Once = Lambda( _
			"", _
			"If boolFirst Then [].Apply varFunction, Arguments : boolFirst = False", _
			"boolFirst, [], varFunction", _
			Array(True, [], varFunction))
	End Function

	Public Function Min(numA, numB)
		Min = [If](numA < numB, numA, numB)
	End Function

	Public Function Max(numA, numB)
		Max = [If](numA > numB, numA, numB)
	End Function

	Public Function Zip(varLeft, varRight)
		' Array(a, b, c), Array(d, e, f) ->
		' Array(Array(a, d), Array(b, e), Array(c, f))

		Dim arrLeft, arrRight
		arrLeft = CArray(varLeft)
		arrRight = CArray(varRight)

		Dim lngPtr
		Dim arrZipped()
		ReDim arrZipped(1)
		For lngPtr = _
			0 To Min(UBound(varLeft), UBound(varRight))
			ReDim Preserve arrZipped( _
				[If](lngPtr > UBound(arrZipped), _
					UBound(arrZipped) * 2, _
					UBound(arrZipped)))
			arrZipped(lngPtr) = Append( _
				Array(varLeft(lngPtr)), _
				Array(varRight(lngPtr)))
		Next

		ReDim Preserve arrZipped( _
			Min(UBound(varLeft), UBound(varRight)))
		Zip = arrZipped
	End Function

	''TODO Fix
	Public Function Curry(varFunction, lngArgumentsCount)
		[Set] Curry, Lambda("", _
			"arrSavedArguments = [].Append(arrSavedArguments, Arguments)" & vbNewLine & _
			"If UBound(arrSavedArguments) = lngArgumentsCount - 1 Then" & vbNewLine & _
			"	Return [].Apply(varFunction, arrSavedArguments)" & vbNewLine & _
			"Else" & vbNewLine & _
			"	Return Callee" & vbNewLine & _
			"End If", _
			"[], varFunction, lngArgumentsCount, arrSavedArguments", _
			Array([], varFunction, lngArgumentsCount, Array()))
	End Function

	Public Function Partial(varFunction, arrArguments)
		[Set] Partial, Lambda("", _
			"Return [].Apply(varFunction, [].Append(arrArguments, Arguments))", _
			"[], varFunction, arrArguments", _
			Array([], varFunction, arrArguments))
	End Function

	Public Function Reverse(varSet)
		Dim arrSet, lngPtr, lngUpperBound
		[Set] arrSet, CArray(varSet)
		[Set] lngUpperBound, UBound(arrSet)
		Dim arrReversed()
		ReDim arrReversed(lngUpperBound)
		For lngPtr = 0 To lngUpperBound
			[Set] arrReversed(lngUpperBound - lngPtr), arrSet(lngPtr)
		Next
		[Set] Reverse, arrReversed
	End Function

	Public Function Chain
	End Function

	Public Function Value
	End Function
End Class

Class Lazy
End Class

Class Stream
	Public Current
	Public [Next]
End Class

Class Pair
	Public Left, Right
End Class

Class AnonymousFunction
	Private Sub [Set](ByRef varVariable, varValue)
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

	Private strCodeBody, arrSavedBindings
	Public Sub Init(strParameters, strBody, strBindings, arrBindings)
		Dim lngCounter, varItem

		strCodeBody = ""
		lngCounter = 0
		For Each varItem In Split(strParameters, ",")
			strCodeBody = strCodeBody & _
				"Dim " & varItem & vbNewLine & _
				"[Set] " & varItem & ", " & _
				"objArguments.[" & CStr(lngCounter) & "]" & vbNewLine
			lngCounter = lngCounter + 1
		Next

		strCodeBody = strCodeBody & _
			"Dim Arguments()" & vbNewLine & _
			"ReDim Arguments(objArguments.length - 1)" & vbNewLine & _
			"Dim ArgumentsCount" & vbNewLine & _
			"For ArgumentsCount = 0 To UBound(Arguments)" & vbNewLine & _
			"	[Set] Arguments(ArgumentsCount), _" & vbNewLine & _
			"		EVal(""objArguments.["" & CStr(ArgumentsCount) & ""]"")" & vbNewLine & _
			"Next" & vbNewLine & _
			"ArgumentsCount = objArguments.length" & vbNewLine

		arrSavedBindings = arrBindings
		lngCounter = -1
		For Each varItem In Split(strBindings, ",")
			lngCounter = lngCounter + 1
			strCodeBody = strCodeBody & _
				"Dim " & varItem & vbNewLine & _
				"[Set] " & varItem & ", " & _
				"arrSavedBindings(" & CStr(lngCounter) & ")" & vbNewLine
		Next
		strCodeBody = strCodeBody & strBody & vbNewLine

		lngCounter = -1
		For Each varItem In Split(strBindings, ",")
			lngCounter = lngCounter + 1
			strCodeBody = strCodeBody & _
				"[Set] arrSavedBindings(" & CStr(lngCounter) & "), " & _
				varItem & vbNewLine
		Next
	End Sub

	Public Sub Apply(objArguments, Callee)
		' Useful Keywords:
		' ArgumentsCount, Arguments,
		' Callee, Return

		' Other Binded Keywords:
		' objArguments, strCodeBody, arrSavedBindings,
		' varReturnValue, Set, ReturnValue, Apply
		
		' Don't Use: Exit Function

		Execute strCodeBody
	End Sub
End Class

