Option Explicit

Sub Include(file)
	On Error Resume Next

	Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
	ExecuteGlobal FSO.OpenTextFile(file & ".vbs", 1).ReadAll()
	Set FSO = Nothing

	If Err.Number <> 0 Then
		If Err.Number = 1041 Then
			Err.Clear
		Else
			WScript.Quit 1
		End If
	End If
End Sub

If WScript.ScriptName = "main_template.vbs" Then
	Include "v_Data"

	Dim stack, _
		queue, _
		collection

	Set stack = New v_Data_Stack

	stack.Push "Apple"
	stack.Push "Orange"
	stack.Push "Banana"
	stack.Push "Strawberry"

	WScript.Echo stack.Pop()

	Set queue = New v_Data_Queue

	queue.Enqueue "Dog"
	queue.Enqueue "Cat"
	queue.Enqueue "Bird"
	queue.Enqueue "Lizard"

	WScript.Echo queue.Dequeue()

	Set collection = New v_Data_Collection

	collection.Add "Car"
	collection.Add True
	collection.Add 342

	WScript.Echo collection(2)
End If