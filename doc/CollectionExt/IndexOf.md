# CollectionExt.IndexOf Method

Searches for the specified object and returns the one-based index of the first occurrence within the entire Collection

```vb
Public Function IndexOf(ByVal Source As Collection, ByVal Value As Variant, Optional ByVal Comparer As IEqualityComparer) As Long
```

### Parameters

**Source** `Collection` <br>
Collection which will be scanned.

**Value** `Variant` <br>
The item to locate in the Collection.

**Comparer** `IEqualityComparer` <br>
Optional. A comparer to compare items between each other.

### Returns

`Long` <br>
The following example shows how IndexOf method can be used for strings.

```vb
Option Explicit
' Main module

Public Sub Start()

    Dim Dinosaurs as Collection
    Set Dinosaurs = Lapis.CollectionExt.Make("Tyrannosaurus", "Amargasaurus", "Mamenchisaurus", _
                                             "Brachiosaurus", "Deinonychus", "Tyrannosaurus", _
                                             "Compsognathus")

    Debug.Print "IndexOf Tyrannosaurus: " & Lapis.CollectionExt.IndexOf(Dinosaurs, "Tyrannosaurus")
    Debug.Print "IndexOf Hydronexus: " & Lapis.CollectionExt.IndexOf(Dinosaurs, "Hydronexus")

End Sub

' This code example produces the following output:
'
' IndexOf Tyrannosaurus: 1
' IndexOf Hydronexus: -1
```

