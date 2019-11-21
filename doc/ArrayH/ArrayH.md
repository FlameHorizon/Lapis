# ArrayH module
Represents a collection of helper method to work with an array type.

## Method
|Name|Description|
|-|-|
|[Exists (Variant, Variant)](./Exists.md)|Checks if given elements exists in the array.|
|[IsInitialized (Variant())](./IsInitialized.md)|Indicates if array is initialized.|
|[ToCollection (Variant())](./ToCollection.md)|Converts array into a collection.|
|[Copy (Variant(), Long, Variant(), Long, Long)](./Copy.md)|Copies elements from an Array starting at SourceIndex and pastes them to another Array starting at DestinationIndex. Number of elements which will be copied is specified in Length parameter.|
|[Rank (Variant()))](./Rank.md)|Returns the number of dimensions of an array.|
|[Length (Variant())](./Length.md)|Returns the number of elements in single dimension of array.|
|[NumElements (Variant(), Long)](./NumElements.md)|Returns the number of elements in the specified dimension (Dimension) of the array in Arr. If you omit Dimension, the first dimension is used.|
|[Clear (Variant(), Long, Long)](./Clear.md)|Clears an range of items in Array starting at Index.|
|[ToString (Variant())](./ToString.md)|Returns a String which contains every element in an Array recursively.|
|[BinarySearch (Variant(), Long, Long, Variant, IComparer)](./BinarySearch.md)|Searches a section of an array for a given element using a binary search algorithm.|
|[GetLowerBound (Variant(), Long)](./GetLowerBound.md)|Return the index of the first element of the specified dimension in the array.|
|[IndexOf (Variant, Variant, Long, Long)](./IndexOf.md)|Returns the index of the first occurrence of a given value in a range of an array.|
|[SetValue (Variant(), Variant, Long)](./SetValue.md)|Sets a value in the given array using element and index of within range of array.|
|[ToArrayIList (IList)](./ToArrayIList.md)|Converts an IList object into array.|
|[ToVariantArray (Variant)](./ToVariantArray.md)|Converts an array into variant array.|
|[StringArray (ParamArray Variant())](./StringArray.md)|Converts ParamArray into strongly typed array of strings.|