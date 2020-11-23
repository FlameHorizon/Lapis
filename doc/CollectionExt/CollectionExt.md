# CollectionExt module

The CollectionExt module contains methods that enables users to use collection in a more robust way.

# Methods

|Name|Description|
|-|-|
|[GroupBy (Collection, String)](./GroupBy.md)|Returns a dictionary with grouped values where key is a unique value and item is a collection of items which matches key.|
|[Concat (Collection, Collection)](./Concat.md)|Concatenates two sequences.|
|[ToArray (Collection)](./ToArray.md)|Creates an array from a collection.|
|[Distinct (Collection, IEqualityComparer)](./Distinct.md)|Returns distinct elements from a sequence by using a specified IEqualityComparer to compare values.|
|[Contains (Collection, Variant, IEqualityComparer)](./Contains.md)|Determines whether a sequence contains a specified element.|
|[IndexOf (Collection, Variant, IEqualityComparer)](./IndexOf.md)|Searches for the specified object and returns the one-based index of the first occurrence within the entire Collection|
|[AddRange (Collection, Collection)](./AddRange.md)|Adds the elements of the specified collection to the end of the set.|
|[Sort (Collection, Lapis.IComparer)](./Sort.md)|Sorts given collection using merge sort according to defined comparer.|
|[Make (ParamArray Variant)](./Make.md)|Creates a new collection based on to list of arguments.|
|[Except (Collection, Collection, IEqualityComparer)](./Except.md)|Produces the set difference of two sequences by using the specified IEqualityComparer to compare values.|
|[Intersect (Collection, Collection, IEqualityComparer)](./Intersect.md)|Produces the set intersection of two sequences by using the specified IEqualityComparer to compare values.|
|[Min (Collection, IComparer)](./Min.md)|Invokes a Comparer on each element of a sequence and returns the minimum resulting value.|
|[Max (Collection, IComparer)](./Max.md)|Invokes a Comparer on each element of a sequence and returns the maximum resulting value.|
|[Range (Long, Long)](./Range.md)|Generates a sequence of integral numbers within a specified range.|
|[Repeat (Variant, Long)](./Repeat.md)|Generates a sequence that contains one repeated value.|
|[Reverse (Collection)](./Reverse.md)|Inverts the order of the elements in a sequence.|
|[Sum (Collection, Lapis.IConverter)](./Sum.md)|Computes the sum of a sequence of Int32 values.|
|[Average (Collection, Lapis.IConverter)](./Average.md)|Computes the average of a sequence of values that is obtained by invoking a projection function on each element of the input sequence.|
|[Take (Collection, Long)](./Take.md)|Returns a specified number of contiguous elements from the start of a sequence.|
|[All (Collection, Predicate)](./All.md)|Determines whether all elements of a sequence satisfy a condition.|
|[Some (Collection, Predicate)](./Some.md)|Determines whether any element of a sequence satisfies a condition.|
|[Skip (Collection, Long)](./Skip.md)|Bypasses a specified number of elements in a sequence and then returns the remaining elements.|
|[SequenceEqual (Collection, Collection, IEqualityComparer)](./SequenceEqual.md)||
|[First (Collection, Predicate)](./First.md)|Returns the first element in a sequence that satisfies a specified condition.|
|[Last (Collection, Predicate)](./Last.md)|Returns the last element of a sequence.|
|[SelectOne (Collection, Predicate)](./SelectOne.md)|Returns the only element of a sequence that satisfies a specified condition, and throws an exception if more than one such element exists.|
|[Count (Collection, Predicate)](./Count.md)|Returns a number that represents how many elements in the specified sequence satisfy a condition.|
|[Where (Collection, Predicate)](./Where.md)|Filters a sequence of values based on a predicate|
|[Convert (Collection, Lapis.IConverter)](./Convert.md)|Projects each element of a sequence into a new form.|
