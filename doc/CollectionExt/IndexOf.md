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
Just a comparer

### Returns

`Long` <br>
The index of value if found in the collection; otherwise, 0.

