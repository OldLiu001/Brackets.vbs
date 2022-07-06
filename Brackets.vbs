Option Explicit

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
		Set [Function] = New AnonymousFunction
		[Function].Init strParameters, strBody
		Set [Function] = WrapArguments(WrapFunction([Function]))
	End Function

	Public Sub Assert(ByVal boolCondition, ByVal strSource, ByVal strDescription)
		If Not boolCondition Then
			Err.Raise vbObjectError, strSource, strDescription
		End If
	End Sub

	Public Sub ForEach(ByVal varSubprogram, ByRef varCollection)
		Dim varItem
		For Each varItem in varCollection
			Call varSubprogram(varItem)
		Next
	End Sub

	Public Function Range(ByVal lngStart, ByVal lngStop, ByVal lngStep)
		Assert (lngStop - lngStart) / lngStep >= 0, _
		"<Function> Range", "Invaild parameter(s)."
		
		Dim arrRange(), lngCounter, lngPointer
		ReDim arrRange(Fix((lngStop - lngStart) / lngStep))
		
		lngPointer = 0
		For lngCounter = lngStart To lngStop Step lngStep
			arrRange(lngPointer) = lngCounter
			lngPointer = lngPointer + 1
		Next
		Range = arrRange
	End Function

	Public Function Map(ByVal varFunction, ByRef varCollection)
		Dim arrMap(), varItem, lngPointer
		ReDim arrMap(1)
		lngPointer = 0
		For Each varItem in varCollection
			If UBound(arrMap) < lngPointer Then
				ReDim Preserve arrMap(UBound(arrMap) * 2)
			End If
			[Set] arrMap(lngPointer), varFunction(varItem)
			lngPointer = lngPointer + 1
		Next
		ReDim Preserve arrMap(lngPointer)
		Map = arrMap
	End Function

End Class

Class AnonymousFunction
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

