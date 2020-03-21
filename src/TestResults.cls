VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ExUnit")
Option Explicit

Private Const ClassName As String = "TestResults"

Private pItems As New Dictionary '<String, TestResult>


Public Property Get Count()
    Count = pItems.Count
End Property


Public Property Get Item(ByVal Key As Variant) As Variant
    
    If IsObject(pItems.Item(Key)) Then
        Set Item = pItems.Item(Key)
    Else
        Item = pItems.Item(Key)
    End If
    
End Property


Public Sub Add(ByVal Item As TestResult)
    
    Const MethodName = "Add"
    
    If Item Is Nothing Then
        Errors.OnArgumentNull "Item", GetSig(MethodName)
    End If
    
    If Item.Source = vbNullString Then
        Errors.OnArgumentError "Item", "Name of the Source can't be empty. " & GetSig(MethodName)
    End If
    
    If pItems.Exists(Item.Source) Then
        If ShouldUpdateTestResult(Item) Then
            pItems.Remove Item.Source
            pItems.Add Item.Source, Item
        End If
    Else
        pItems.Add Item.Source, Item
    End If
    
End Sub


Private Function GetSig(ByVal MethodName As String) As String
    GetSig = ClassName & "." & MethodName
End Function


Private Function ShouldUpdateTestResult(ByVal Result As TestResult) As Boolean

    ShouldUpdateTestResult = (pItems.Item(Result.Source).Outcome = TestOutcome.Passed _
                              And Result.Outcome = TestOutcome.Failed)
                              
End Function


Public Sub Remove(ByVal Key As Variant)
    pItems.Remove Key
End Sub


Public Property Get NewEnum() As IUnknown
    '@Ignore SetAssignmentWithIncompatibleObjectType
    Set NewEnum = pItems
End Property

