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


Public Sub Add(ByVal Item As TestResult)
    
    Const MethodName = "Add"
    
    If Item Is Nothing Then
        Exception.ArgumentNullException "Item", ClassName & "." & MethodName
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


Private Function ShouldUpdateTestResult(ByVal Result As TestResult) As Boolean

    ShouldUpdateTestResult = (pItems.Item(Result.Source).Result = TestOutcome.Passed _
                              And Result.Result = TestOutcome.Failed)
                              
End Function


Public Property Get Count()
    Count = pItems.Count
End Property


Public Property Get Enumerator() As Dictionary
    Set Enumerator = pItems
End Property
