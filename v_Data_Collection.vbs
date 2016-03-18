Class v_Data_Collection
	Private pCollection

	Private Sub Class_Initialize()

	End Sub


	' Properties

	
	Public Property Get Count()
		If IsArrayAllocated(pCollection) Then Count = UBound(pCollection) + 1
	End Property

	Public Default Property Get Item(intIndex)
		If IsArrayAllocated(pCollection) Then
    			If IsObject(pCollection(intIndex)) Then
        			Set Item = pCollection(intIndex)
    			Else
        			Item = pCollection(intIndex)
    			End If
		End If
	End Property


	' Methods


	Public Sub Add(varItem)
		Extend 1

		If IsObject(varItem) Then
			Set pCollection(UBound(pCollection)) = varItem
		Else
			pCollection(UBound(pCollection)) = varItem
		End If
	End Sub

	Public Sub Clear()
		If IsArrayAllocated(pCollection) Then Erase pCollection
	End Sub

	Public Function Contains(varItem)
		Dim blnContains, _
			i

		blnContains = False

		For i = 0 To UBound(pCollection)
			If IsObject(varItem) And IsObject(pCollection(i)) Then
				If pCollection(i) Is varItem Then blnContains = True
			ElseIf Not IsObject(Collection(i)) And Not IsObject(varItem) Then
				If pCollection(i) = varItem Then blnContains = True
			End If
			If blnContains Then Exit For
		Next

		Contains = blnContains
	End Function

	Public Sub Remove(intIndex)
		If IsArrayAllocated(pCollection) Then
			Dim objCollection, _
				i

			Set objCollection = New v_Data_Collection

			For i = 0 To UBound(pCollection)
  				Do
    					If i = intIndex Then Exit Do
        				
					objCollection.Add pCollection(i)
  				Loop While False
			Next

			Clear()

			For i = 0 To objCollection.Count - 1
				Me.Add objCollection(i)
			Next
		End If
	End Sub

	
	' Helper Methods


	Private Sub Extend(intSize)
		If IsArrayAllocated(pCollection) Then
			ReDim Preserve pCollection(UBound(pCollection) + intSize)
		Else
			pCollection = Array()
			ReDim pCollection(intSize - 1)
		End If
	End Sub

	Private Function IsArrayAllocated(arr)
		IsArrayAllocated = False
		If IsArray(arr) Then
			On Error Resume Next
			Dim ub : ub = UBound(arr)
			If (Err.Number = 0) And (ub >= 0) Then IsArrayAllocated = True
		End If  
	End Function
	
	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "v_Data_Collection.vbs" Then

End If