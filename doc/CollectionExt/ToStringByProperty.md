# CollectionExt.ToStringByProperty Method

Returns string based on the given name of the object's property.

```vb
Public Function ToStringByProperty(ByVal Source As Collection, ByVal PropertyName As String, Optional ByVal Delimiter As String = ",") As String
```

### Parameters

**Source** `Collection` <br>
Set of items which will be converted.

**PropertyName** `String` <br>
Name of the object's property which will be used for conversion.

**Delimiter** `String` <br>
Optional. Defines delimiter which will be used between each item.

### Returns

`String` <br>
Text which will represent all items in the source.

### Errors

`OnArgumentOutOfRange` <br>
When value of the given property is not accessible.

## Examples

The following code example demonstrates how to use ToStringByProperty(Collection, String, String) to covert set to string separated by comma.

```vb
Option Explicit

Public Sub Start()

    Dim Pets As Collection
    Set Pets = Lapis.CollectionExt.Make(MakePet("Alfie", 2), _
                                        MakePet("Ricky", 5), _
                                        MakePet("Mikey", 6))
    Dim Result As String
    Result = Lapis.CollectionExt.ToStringByProperty(Pets, "Name", Delimiter:=", ")
    Debug.Print Result

End Sub


Private Function MakePet(ByVal Name As String, ByVal Age As Long) As Pet

    Dim Output As New Pet
    Output.Name = Name
    Output.Age = Age
    Set MakePet = Output

End Function

' This code produces the following output:
'
' Alfie, Ricky, Mikey

' Class Module: Pet
Option Explicit

Public Name As String
Public Age As Long
```

