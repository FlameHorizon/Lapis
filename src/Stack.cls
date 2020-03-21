VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Exposed
Option Explicit
'@Folder("Collection")

Private Const ClassName = "Stack"
Private Const DefaultCapacity As Long = 10

Private pArr() As Variant
Private pSize As Long


Private Sub Class_Initialize()

    ReDim pArr(0 To DefaultCapacity)
    pSize = 0
    
End Sub


Public Sub Init1(ByVal InitialCapacity As Long)
    
    If InitialCapacity < 0 Then
        Errors.OnArgumentError "InitialCapacity", _
                               "Initial capacity value must be a non negetive number. " & ClassName & ".Init1"
    End If
    
    Dim InitCap As Long
    InitCap = InitialCapacity
    If InitialCapacity < DefaultCapacity Then
        InitCap = DefaultCapacity
    End If

    ReDim pArr(0 To InitCap)
    pSize = 0
    
End Sub


Public Property Get Count() As Long
    Count = pSize
End Property


Private Property Get ICollection_Count() As Long
    ICollection_Count = Count
End Property


Public Sub Clear()
    
    ReDim pArr(0 To DefaultCapacity)
    pSize = 0

End Sub


Public Function Contains(ByRef Obj As Variant, ByRef Comparer As IEqualityComparer) As Boolean

    Dim Count As Long: Count = pSize

    Do While (Count >= 0)
        If Comparer.Equals(Obj, Nothing) Then
            If Comparer.Equals(pArr(Count), Nothing) Then
                Contains = True
                Exit Function
            End If
        ElseIf (Not Comparer.Equals(pArr(Count), Nothing)) And Comparer.Equals(pArr(Count), Obj) Then
            Contains = True
            Exit Function
        End If

        Count = Count - 1
    Loop

    Contains = False

End Function


' Pops an item from the top of the stack. If the stack is empty, Pop
' throws an InvalidOperationException.
Public Function Pop() As Variant

    If pSize = 0 Then
        Errors.OnInvalidOperation vbNullString, "Stack is empty." & ClassName & ".Pop"
    End If
    
    If IsObject(pArr(pSize - 1)) Then
        Set Pop = pArr(pSize - 1)
    Else
        Pop = pArr(pSize - 1)
    End If
    
    pArr(pSize) = Empty
    pSize = pSize - 1
    
End Function


' Pushes an item to the top of stack
Public Sub Push(ByRef Obj As Variant)
    
    If (pSize = UBound(pArr) + 1) Then
        Dim NewArray() As Variant
        ReDim NewArray(0 To 2 * (UBound(pArr)))
        ArrayExt.Copy pArr, 0, NewArray, 0, pSize
        pArr = NewArray
    End If
    
    If IsObject(Obj) Then
        Set pArr(pSize) = Obj
    Else
        pArr(pSize) = Obj
    End If
    pSize = pSize + 1
    
End Sub


' Returns the top object on the stack without removing it. If the stack
' is empty, Peek throws an InvalidOperationException.
Public Function Peek() As Variant

    If pSize = 0 Then
        Errors.OnInvalidOperation vbNullString, "Stack is empty." & ClassName & ".Peek"
    End If

    If IsObject(pArr(pSize - 1)) Then
        Set Peek = pArr(pSize - 1)
    Else
        Peek = pArr(pSize - 1)
    End If

End Function


' Copies the Stack to an array, in the same order Pop would return the items.
Public Function ToArray() As Variant()

    Dim Output() As Variant
    ReDim Output(0 To pSize - 1)

    Dim i As Long

    Do While (i < pSize)
        Output(i) = pArr(pSize - i - 1)
        i = i + 1
    Loop

    ToArray = Output

End Function


Public Sub CopyTo(ByRef Arr() As Variant, ByVal Index As Long)

    Const MethodName = ".CopyTo"
    
    If ArrayExt.Rank(Arr) < 1 Then
        Errors.OnArgumentError "Arr", _
                               "Array is not initalized. " & ClassName & MethodName
    End If
    
    If ArrayExt.Rank(Arr) > 1 Then
        Errors.OnArgumentError "Arr", _
                               "Multi dimensions arrray are not supported. " & ClassName & MethodName
    End If

    If Index < 0 Then
        Errors.OnArgumentError "Index", _
                               "Given argument is out of available range. " & ClassName & MethodName
    End If

    If ArrayExt.Length(Arr) - Index < pSize - 1 Then
        Errors.OnArgumentError vbNullString, "Invalid offset length." & ClassName & MethodName
    End If

    Dim i As Long: i = 0

    Do While (i < pSize)
        Arr(i + Index) = pArr(pSize - i - 1)
        i = i + 1
    Loop

End Sub


Private Sub ICollection_CopyTo(ByRef Arr() As Variant, ByVal Index As Long)
    CopyTo Arr, Index
End Sub



