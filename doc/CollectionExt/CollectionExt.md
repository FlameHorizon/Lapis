# CollectionExt module

The CollectionExt module contains methods that enables users to use collection in a more robust way.

# Methods

|Name|Description|
|---|---|
|[AddRange (Collection, Collection)](./AddRange.md)|Adds the elements of the specified collection to the end of the set.|
|[All (Collection, ICallable)](./All.md)|Determines whether all elements of a sequence satisfy a condition.|
|[Average (Collection, ICallable)](./Average.md)|Computes the average of a sequence of values that is obtained by invoking a projection function on each element of the input sequence. If `Selector` is not defined,computes the average of a sequence values.|
|[Concat (Collection, Collection)](./Concat.md)|Concatenates two sequences.|
|[Contains (Collection, Variant, IEqualityComparer)](./Contains.md)|Determines whether a sequence contains a specified element.|
|[Convert (Collection, ICallable)](./Convert.md)|Projects each element of a sequence into a new form.|
|[Count (Collection, ICallable)](./Count.md)|Returns a number that represents how many elements in the specified sequence satisfy a condition. If `Predicate` is not defined, returns the number of elements in a sequence.|
|[Distinct (Collection, IEqualityComparer)](./Distinct.md)|Returns distinct elements from a sequence by using a specified IEqualityComparer to compare values.|
|[Except (Collection, Collection, IEqualityComparer)](./Except.md)|Produces the set difference of two sequences by using the specified IEqualityComparer to compare values.|
|[First (Collection, ICallable)](./First.md)|Returns the first element in a sequence that satisfies a specified condition. If `Predicate` is not defined then returns the first element of a sequence.|
|[GroupBy (Collection, String)](./GroupBy.md)|Returns a dictionary with grouped values where key is a unique value and item is a collection of items which matches key.|
|[IndexOf (Collection, Variant, IEqualityComparer)](./IndexOf.md)|Searches for the specified object and returns the one-based index of the first occurrence within the entire Collection.|
|[Intersect (Collection, Collection, IEqualityComparer)](./Intersect.md)|Produces the set intersection of two sequences by using the specified IEqualityComparer to compare values.|
|[Last (Collection, ICallable)](./Last.md)|Returns the last element of a sequence that satisfies a specified condition. If `Predicate` is not specified, returns the last element of a sequence.|
|[Make (ParamArray Variant)](./Make.md)|Creates a new collection based on to list of arguments.|
|[Max (Collection, ICallable)](./Max.md)|Invokes a transform function on each element of a sequence and returns the maximum value.|
|[Min (Collection, ICallable)](./Min.md)|Invokes a transform function on each element of a generic sequence and returns the minimum resulting value.|
|[Range (Long, Long)](./Range.md)|Generates a sequence of integral numbers within a specified range.|
|[Repeat (Variant, Long)](./Repeat.md)|Generates a sequence that contains one repeated value.|
|[Reverse (Collection)](./Reverse.md)|Inverts the order of the elements in a sequence.|
|[SelectOne (Collection, ICallable)](./SelectOne.md)|Returns the only element of a sequence that satisfies a specified condition, and throws an exception if more than one such element exists. If `Predicate` is not specified, returns the only element of a sequence, and throws an exception if there is not exactly one element in the sequence.|
|[SequenceEqual (Collection, Collection, IEqualityComparer)](./SequenceEqual.md)|Determines whether two sequences are equal by comparing their elements by using a specified IEqualityComparer.|
|[Skip (Collection, Long)](./Skip.md)|Bypasses a specified number of elements in a sequence and then returns the remaining elements.|
|[Some (Collection, ICallable)](./Some.md)|Determines whether any element of a sequence satisfies a condition. If `Predicate` is not given, then determines whether a sequence contains any elements.|
|[Sort (Collection, Lapis.IComparer)](./Sort.md)|Sorts given collection using merge sort according to defined comparer.|
|[Sum (Collection, ICallable)](./Sum.md)|Computes the sum of a sequence of values. If `Selector` is not defined computes the sum of a sequence values.|
|[Take (Collection, Long)](./Take.md)|Returns a specified number of contiguous elements from the start of a sequence.|
|[ToArray (Collection)](./ToArray.md)|Creates an array from a collection.|
|[Where (Collection, ICallable)](./Where.md)|Filters a sequence of values based on a predicate|
