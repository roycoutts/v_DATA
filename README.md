# v_Data
A library of data structures for VBScript.

The v_Data class libraries provide a set of wrapper objects around native VBScript and Microsoft .NET data structures:

VBScript:
```
Array()
Scripting.Dictionary
```
.NET:
```
System.Collections.ArrayList
System.Collections.Hashtable
System.Collections.Queue
System.Collections.SortedList
System.Collections.Stack
```
The OLE Programmatic Identifiers (ProgIDs) for these data structures are tedious to remember and sometimes difficult to implement (due to lack of documentation). v_Data provides a common interface for the ProgIDs to quickly access and use these data structures in VBScript.


v_Data_Array:
```
Set arr = New v_Data_Array

arr.FromArray Array("New York", "Chicago", "Los Angelos", "Miami", "Toronto")

WScript.Echo arr(3)
```

Output:
```
Miami
```


v_Data_ArrayList:
```
Set arraylist = New v_Data_ArrayList

arraylist.Add "Train"
arraylist.Add "Bus"
arraylist.Add "Car"
arraylist.Add "Bicycle"
arraylist.Add "Boat"

WScript.Echo arraylist(2)
```

Output:
```
Car
```


v_Data_Collection:
```
Set collection = New v_Data_Collection

collection.Add "Tree"
collection.Add True
collection.Add 342

WScript.Echo collection(2)
```

Output:
```
342
```


v_Data_Dictionary:
```
Set dictionary = New v_Data_Dictionary

dictionary.Add "Key 1", "Item 1"
dictionary.Add "Key 2", "Item 2"
dictionary.Add "Key 3", "Item 3"
dictionary.Add "Key 4", "Item 4"

WScript.Echo dictionary("Key 3")
```

Output:
```
Item 3
```


v_Data_List:
```
Set list = New v_Data_List

list.Add "Point", 1
list.Add "Point Cloud", 2
list.Add "Curve", 4
list.Add "Surface", 8
list.Add "Polysurface", 16
list.Add "Mesh", 32

WScript.Echo list.GetByIndex(1)
```

Output:
```
32
```


v_Data_Queue:
```
Set queue = New v_Data_Queue

queue.Enqueue "Dog"
queue.Enqueue "Cat"
queue.Enqueue "Bird"
queue.Enqueue "Lizard"

WScript.Echo queue.Dequeue()
```

Output:
```
Dog
```


v_Data_Stack:
```
Set stack = New v_Data_Stack

stack.Push "Apple"
stack.Push "Orange"
stack.Push "Banana"
stack.Push "Strawberry"

WScript.Echo stack.Pop()
```

Output:
```
Strawberry
```
