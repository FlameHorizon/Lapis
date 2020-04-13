VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Predicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Lapis")
Option Explicit

Private Const ClassName As String = "ClassName"

Private pOperator As ComparisonOperator
Private pComparisonValue As Variant
Private pComparer As IComparer


Public Function Eval(ByVal Item As Variant) As Boolean
    
    If pComparer Is Nothing Then
        Set pComparer = ComparerStatic.Default
    End If
    
    Select Case pOperator
        Case ComparisonOperator.EqualTo
            Eval = (pComparer.Compare(pComparisonValue, Item) = 0)
        
        Case Else
            Errors.OnArgumentOutOfRange "Operator", ClassName & ".Eval"
            
    End Select
    
End Function


Public Property Get Operator() As ComparisonOperator
    Operator = pOperator
End Property


Public Property Let Operator(ByVal Value As ComparisonOperator)
    pOperator = Value
End Property


Public Property Get ComparisonValue() As Variant
    ComparisonValue = pComparisonValue
End Property


Public Property Let ComparisonValue(ByVal Value As Variant)
    pComparisonValue = Value
End Property


Public Property Get Comparer() As IComparer
    Set Comparer = pComparer
End Property


Public Property Let Comparer(ByVal Value As IComparer)
    Set pComparer = Value
End Property
