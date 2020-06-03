# CollectionExt.Take Method

Returns a specified number of contiguous elements from the start of a sequence.

```vb
Public Function Take(ByVal Source As Collection, ByVal Count As Long) As Collection
```

### Parameters

**Source** `Collection` <br>
The sequence to return elements from.

**Count** `Long` <br>
A projection function to apply to each element.

### Returns

`Collection` <br>
A Collection that contains the specified number of elements from the start of the input sequence.

### Errors

`OnArgumentNull` <br>
`Source` is `Nothing`

## Examples

The following code example demonstrates how to use Take to return elements from the start of a sequence.

```vb
' Start Module
Option Explicit

Public Sub Start()

    Dim Grades As Collection
    Set Grades = CollectionExt.Make(59, 82, 70, 56, 92, 98, 85)
    
    Dim BottomThreeGrades As Collection
    Set BottomThreeGrades = CollectionExt.Take( _
        CollectionExt.Sort(Grades, Lapis.Factory.GetLongComparer), 3)
        
    Debug.Print "The bottom three grades are:"
        
    Dim Grade As Variant
    For Each Grade In BottomThreeGrades
        Debug.Print Grade
    Next Grade

End Sub


' This code produces the following output:
' 56 
' 59 
' 70 
```

