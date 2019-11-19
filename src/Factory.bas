Attribute VB_Name = "Factory"
Option Explicit
'@Folder("Lapis")


Public Function GetStringEqualityComparer() As StringEqualityComparer
    Set GetStringEqualityComparer = New StringEqualityComparer
End Function


Public Function GetLongEqualityComparer() As LongEqualityComparer
    Set GetLongEqualityComparer = New LongEqualityComparer
End Function


Public Function GetSortedList() As SortedList
    Set GetSortedList = New SortedList
End Function


Public Function GetStack() As Stack
    Set GetStack = New Stack
End Function


Public Function GetTreeNode() As TreeNode
    Set GetTreeNode = New TreeNode
End Function


