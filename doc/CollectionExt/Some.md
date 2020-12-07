# CollectionExt.Some Method

Determines whether any element of a sequence satisfies a condition.

```vb
Public Function Some(ByVal Source As Collection, Optional ByVal Predicate As Predicate) As Boolean
```

### Parameters

**Source** `Collection` <br>
A Collection whose elements to apply the predicate to.

**Predicate** `Predicate` <br>
A function to test each element for a condition.

### Returns

`Boolean` <br>
`True` if any elements in the source sequence pass the test in the specified predicate; otherwise, `False`.

### Errors

`OnArgumentNull` <br>
`Source` is `Nothing`

## Examples

The following code example demonstrates how to use All to determine whether all the elements in a sequence satisfy a condition. Variable AllOlderThan5Years is true if all the pet names start with "B" or if the pets array is empty.

```vb
' Pet Class Module
Option Explicit

Public Name As String
Public Age As Long
```

```vb
' Start Module
Option Explicit

Public Sub Start()

    Dim Pets As Collection
    Set Pets = CollectionExt.Make(MakePet("Barley", 10), _
                                  MakePet("Boots", 4), _
                                  MakePet("Whiskers", 6))
    Dim Pred As Lapis.Predicate
    Set Pred = Lapis.Factory.GetPredicate
    Pred.ComparisonValue = 5
    Pred.Operator = Lapis.ComparisonOperator.GreaterThanOrEqualTo

    Dim AnyOlderThan5Years As Boolean
    AnyOlderThan5Years = CollectionExt.Some(Pets, Pred)
    
    If AnyOlderThan5Years Then
        Debug.Print "Some pets are older than 5 years."
    Else
        Debug.Print "Not pet is older than 5 years."
    End If

End Sub

' This code produces the following output:
'
' Some pets are older than 5 years.


Private Function MakePet(ByVal Name As String, ByVal Age As Long) As Pet
    
    Dim Output As New Pet
    Output.Name = Name
    Output.Age = Age
    Set MakePet = Output
    
End Function
```

