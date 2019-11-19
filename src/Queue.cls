VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder("Collection")

Implements IVBACollection

Private Const ClassName As String = "Queue"
Private Const MinimumGrow As Long = 4

Private pArr() As Variant
Private pHead As Long
Private pTail As Long
Private pSize As Long
Private pGrowFactor As Long


' Default grow factor is 1.5
Public Sub Init1(ByVal Capacity As Long, Optional ByVal GrowFactor As Single = 1.5)
    
    Const MethodName = "Init1"
    
    If (Capacity < 0) Then
        Exception.ArgumentOutOfRangeException "Capacity", ClassName & "." & MethodName
    End If
    
    If (Not (GrowFactor >= 1 And GrowFactor <= 10)) Then
        Exception.ArgumentOutOfRangeException "GrowFactor", ClassName & "." & MethodName
    End If

    ReDim pArr(0 To Capacity - 1)
    pHead = 0
    pTail = 0
    pSize = 0
    pGrowFactor = CLng(GrowFactor * 100)
    
End Sub


Public Property Get Count() As Long
    Count = pSize
End Property


Private Property Get IVBACollection_Count() As Long
    IVBACollection_Count = Count
End Property


Public Sub CopyTo(ByRef Arr() As Variant, ByVal Index As Long)
    
    Const MethodName = "CopyTo"

    If Not ArrayH.IsInitialized(Arr) Then
        Exception.ArgumentNullException "Arr", ClassName & "." & MethodName
    End If

    If ArrayH.Rank(Arr) <> 1 Then
        Exception.ArgumentException "Arr", "Multi-dimension array are not supported." & ClassName & "." & MethodName
    End If

    If Index < 0 Then
        Exception.ArgumentOutOfRangeException "Index", ClassName & "." & MethodName
    End If

    Dim NumToCopy As Long: NumToCopy = pSize
    If NumToCopy = 0 Then
        Exit Sub
    End If

    Dim FirstPart As Long
    FirstPart = IIf(ArrayH.Length(pArr) - pHead < NumToCopy, ArrayH.Length(pArr) - pHead, NumToCopy)
    ArrayH.Copy pArr, pHead, Arr, Index, FirstPart

    NumToCopy = NumToCopy - FirstPart
    If NumToCopy > 0 Then
        ArrayH.Copy pArr, 0, Arr, Index + ArrayH.Length(pArr) - pHead, NumToCopy
    End If

End Sub


Private Sub IVBACollection_CopyTo(ByRef Arr() As Variant, ByVal Index As Long)
    CopyTo Arr, Index
End Sub


' Removes all objects from the queue.
Public Sub Clear()

    If pHead < pTail Then
        ArrayH.Clear pArr, pHead, pSize
    Else
        ArrayH.Clear pArr, pHead, ArrayH.Length(pArr) - pHead
        ArrayH.Clear pArr, 0, pTail
    End If

    pHead = 0
    pTail = 0
    pSize = 0

End Sub


' Adds element to the tail of the queue.
Public Sub Enqueue(ByVal Obj As Variant)
    
    ' Make the Array bigger using growth factor.
    If (pSize = ArrayH.Length(pArr)) Then
        Dim NewCapacity As Long
        NewCapacity = ArrayH.Length(pArr) * pGrowFactor / 100
        
        If NewCapacity < ArrayH.Length(pArr) + MinimumGrow Then
            NewCapacity = ArrayH.Length(pArr) + MinimumGrow
        End If
        
        SetCapacity NewCapacity
    End If

    pArr(pTail) = Obj
    pTail = (pTail + 1) Mod ArrayH.Length(pArr)
    pSize = pSize + 1

End Sub


' Grows or shrinks the buffer to hold capacity objects. Capacity must be >= pSize.
Private Sub SetCapacity(ByVal Capacity As Long)

    Dim NewArray() As Variant
    ReDim NewArray(0 To Capacity)

    If pSize > 0 Then
        If pHead < pTail Then
            ArrayH.Copy pArr, pHead, NewArray, 0, pSize
        Else
            ArrayH.Copy pArr, pHead, NewArray, 0, ArrayH.Length(pArr) - pHead
            ArrayH.Copy pArr, 0, NewArray, ArrayH.Length(pArr) - pHead, pTail
        End If
    End If

    pArr = NewArray
    pHead = 0
    pTail = IIf(pSize = Capacity, 0, pSize)

End Sub


' Determines whether an element is in the Queue.
Public Function Contains(ByRef Obj As Variant, ByRef Comparer As IVBAEqualityComparer) As Boolean
    
    Const MethodName = "Contains"

    If Comparer Is Nothing Then
        Exception.ArgumentNullException "Comparer", ClassName & "." & MethodName
    End If

    Dim Index As Long: Index = pHead
    Dim Count As Long: Count = pSize

    Do While (Count > 0)
        If Comparer.Equals(Obj, Nothing) Then
            If Comparer.Equals(pArr(Index), Nothing) Then
                Contains = True
                Exit Function
            End If
        ElseIf (Comparer.Equals(pArr(Index), Nothing) = False) And (Comparer.Equals(pArr(Index), Obj)) Then
            Contains = True
            Exit Function
        End If

        Index = (Index + 1) Mod ArrayH.Length(pArr)
        Count = Count - 1
    Loop
    Contains = False

End Function


' Removes and returns the object at the beginning of the Queue.
Public Function Dequeue() As Variant
    
    Const MethodName = "Dequeue"

    If Count = 0 Then
        Exception.InvalidOperationException vbNullString, "Queue is empty. " & ClassName & "." & MethodName
    End If

    Dim Removed As Variant
    Removed = pArr(pHead)
    pArr(pHead) = 0
    pHead = (pHead + 1) Mod ArrayH.Length(pArr)
    pSize = pSize - 1
    
    Dequeue = Removed

End Function


' Returns the object at the beginning of the Queue without removing it.
Public Function Peek() As Variant
    
    If Count = 0 Then
        Exception.InvalidOperationException vbNullString, "Queue is empty. " & ClassName & ".Peek"
    End If

    Peek = pArr(pHead)

End Function


' Copies the Queue elements to a new array.
Public Function ToArray() As Variant

    Dim Arr() As Variant
    
    If pSize = 0 Then
        ToArray = Arr
        Exit Function
    End If
    
    ReDim Arr(0 To pSize - 1)

    If (pHead < pTail) Then
        ArrayH.Copy pArr, pHead, Arr, 0, pSize
    Else
        ArrayH.Copy pArr, pHead, Arr, 0, ArrayH.Length(pArr) - pHead
        ArrayH.Copy pArr, 0, Arr, ArrayH.Length(pArr) - pHead, pTail
    End If

    ToArray = Arr

End Function


' Sets the capacity of the actual number of elements in the Queue.
Public Function TrimToSize() As Variant
    SetCapacity pSize
End Function


