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
		collection, _
		dictionary, _
		list, _
		arraylist, _
		arr, _
		hash

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

	collection.Add "Tree"
	collection.Add True
	collection.Add 342

	WScript.Echo collection(2)

	Set dictionary = New v_Data_Dictionary

	dictionary.Add "Key 1", "Item 1"
	dictionary.Add "Key 2", "Item 2"
	dictionary.Add "Key 3", "Item 3"
	dictionary.Add "Key 4", "Item 4"

	WScript.Echo dictionary("Key 3")

	Set list = New v_Data_List

	list.Add "Point", 1
	list.Add "Point Cloud", 2
	list.Add "Curve", 4
	list.Add "Surface", 8
	list.Add "Polysurface", 16
	list.Add "Mesh", 32

	WScript.Echo list.GetByIndex(1)

	Set arraylist = New v_Data_ArrayList

	arraylist.Add "Train"
	arraylist.Add "Bus"
	arraylist.Add "Car"
	arraylist.Add "Bicycle"
	arraylist.Add "Boat"

	WScript.Echo arraylist(2)

	Set arr = New v_Data_Array

	arr.FromArray Array("New York", "Chicago", "Los Angelos", "Miami", "Toronto")

	WScript.Echo arr(3)

	Set hash = New v_Data_HashTable

	hash.Add "FirstName", "Sam"
	hash.Add "LastName", "Smith"
	hash.Add "Title", "Supervisor"
	hash.Add "EmployeeCode", 1457345

	WScript.Echo hash("Title")
End If