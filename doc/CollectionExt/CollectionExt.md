# CollectionExt module

The CollectionExt module contains methods that enables users to use collection in a more robust way.

# Methods

|Name|Description|
|-|-|
|[ToString (Collection, IToString, String)](./ToString.md)|Returns a collection of property values based on the items in the collection.|
|[ToStringByProperty (Collection, String)](./ToStringByProperty.md)||
|[GroupBy (Collection, String)](./GroupBy.md)|Returns a dictionary with grouped values where key is a unique value and item is a collection of items which matches key.|
|[Concat (Collection, Collection)](./Concat.md)|Joins two collections together.|
|[ToArray (Collection)](./ToArray.md)|Converts collection to an array.|
|[Distinct (Collection, IEqualityComparer)](./Distinct.md)|Returns a collection of items which unique property value.|
|[Contains (Collection, Variant, IEqualityComparer)](./Contains.md)|Checks if item exists in the collection using custom comparer.|
|[DistinctValues (Collection, IEqualityComparer)](./DistinctValues.md)|Returns a collection which contains distinct values from the Collection.|
|[IndexOf (Collection, Variant, IEqualityComparer)](./IndexOf.md)|Searches for the specified object and returns the one-based index of the first occurrence within the entire Collection|
|[AddRange (Collection, Collection)](./AddRange.md)|Adds a collection of items to the container.|
|[Sort (Collection, Lapis.IComparer)](./Sort.md)|Sorts given collection using merge sort according to defined comparer.|
|[Make (ParamArray Variant)](./Make.md)|Creates a new collection based on to list of arguments.|
|[Except (Collection, Collection, IEqualityComparer)](./Except.md)|Produces the set difference of two sequences by using the specified IEqualityComparer to compare values.|
|[Intersect (Collection, Collection, IEqualityComparer)](./Intersect.md)|Produces the set intersection of two sequences by using the specified IEqualityComparer to compare values.|
|[Min (Collection, IComparer)](./Min.md)||
|[Max (Collection, IComparer)](./Max.md)||
|[Range (Long, Long)](./Range.md)||
|[Repeat (Variant, Long)](./Repeat.md)||
|[Reverse (Collection)](./Reverse.md)||
|[Sum (Collection, Lapis.IConverter)](./Sum.md)||
|[Average (Collection, Lapis.IConverter)](./Average.md)||
|[Take (Collection, Long)](./Take.md)||
|[All (Collection, Predicate)](./All.md)||
|[Some (Collection, Predicate)](./Some.md)||
|[Skip (Collection, Long)](./Skip.md)||
|[SequenceEqual (Collection, Collection, IEqualityComparer)](./SequenceEqual.md)||
|[First (Collection, Predicate)](./First.md)||
|[Last (Collection, Predicate)](./Last.md)||
|[SelectOne (Collection, Predicate)](./SelectOne.md)||
|[Count (Collection, Predicate)](./Count.md)||
|[Where (Collection, Predicate)](./Where.md)||
