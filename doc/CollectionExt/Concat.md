# CollectionExt.Concat Method

Concatenates two sequences.

```vb
Public Function Concat(ByVal First As Collection, ByVal Second As Collection) As Collection
```

### Parameters

**First** `Collection` <br>
The first sequence to concatenate.

**Second** `Collection` <br>
The sequence to concatenate to the first sequence.

### Returns

`Collection` <br>
A Collection that contains the concatenated elements of the two input sequences.

### Errors

`OnArgumentNull` <br>
`First` or `Second` is `Nothing`

## Examples

The following code example demonstrates how to use Concat(Collection, Collection) to concatenate two sequences.

```vb
' Pet Class
Option Explicit

Public Name as String
Public Age as String

```

```vb
' Main module
Option Explicit

Public Sub Start()

    Dim Cats As Collection
    Set Cats = Lapis.CollectionExt.Make(MakePet("Alfie", 2), _
                                        MakePet("Ricky", 5), _
                                        MakePet("Mikey", 6))

    Dim Dogs As Collection
    Set Dogs = Lapis.CollectionExt.Make(MakePet("Rooney", 7), _
                                        MakePet("Bull", 4), _
                                        MakePet("Fredro", 3))

    Dim Result As Collection
    Set Result = Lapis.CollectionExt.Concat(Cats, Dogs)
    
    Dim Item As Pet
    For Each Item In Result
        Debug.Print Item.Name
    Next Item

End Sub


Private Function MakePet(ByVal Name As String, ByVal Age As Long) As Pet

    Dim Output As New Pet
    Output.Name = Name
    Output.Age = Age
    Set MakePet = Output

End Function

' This code produces the following output:
'
' Alfie
' Ricky
' Mikey
' Rooney
' Bull
' Fredro
```

