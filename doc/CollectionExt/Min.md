# CollectionExt.Min Method

Invokes a Comparer on each element of a sequence and returns the minimum resulting value.

```vb
Public Function Min(ByVal Source As Collection, ByVal Comparer As IComparer) As Variant
```

### Parameters

**Source** `Collection` <br>
A sequence of values to determine the minimum value of `Source`.

**Comparer** `IComparer` <br>
A Comparer to apply to each element.

### Returns

`Variant` <br>
The minimum value in the sequence.

### Errors

`OnArgumentNull` <br>
`Source` is `Nothing`.

-or-

`Comparer` is `Nothing`.

## Examples

The following code example demonstrates how to use `Min` to determine the minimum value in a sequence of projected values.

```vb
' Pet class module
Option Explicit

Public Name As String
Public Age As Long
```

```vb
' PetByAgeComparer class module
Option Explicit

Implements Lapis.IComparer

Public Function Compare(ByVal x As Pet, ByVal y As Pet) As Long

    Compare = IIf(x.Age < y.Age, -1, _
              IIf(x.Age = y.Age, 0, 1))
              
End Function


Private Function IComparer_Compare(ByRef x As Variant, ByRef y As Variant) As Long
    IComparer_Compare = Me.Compare(x, y)
End Function
```

```vb
' Main module
Option Explicit

Public Sub Start()

    Dim Pets As Collection
    Set Pets = CollectionExt.Make(MakePet("Barley", 8), _
                                  MakePet("Boots", 4), _
                                  MakePet("Whiskers", 1))
    
    Dim Min As Pet
    Set Min = CollectionExt.Min(Pets, New PetByAgeComparer)
    
    Debug.Print "The youngest animal is age " & Min.Age
    
End Sub


Private Function MakePet(ByVal Name As String, ByVal Age As Long) As Pet
    
    Dim Output As New Pet
    Output.Name = Name
    Output.Age = Age
    Set MakePet = Output
    
End Function

' The example displays the following output:
' The youngest animal is age 1
```

