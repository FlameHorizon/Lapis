VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestResultImmediatePrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests")
Option Explicit

Implements ITestResultPrinter

Private Const ClassName As String = "TestResultImmediatePrinter"


Public Sub PrintMany(ByVal Items As TestResults)
    
    Dim Item As Variant '<TestResult>
    For Each Item In Items.Enumerator.Items
        PrintSingle Item
    Next Item
    
End Sub


Public Sub PrintSingle(ByVal Item As TestResult)
    
    If Item Is Nothing Then
        Exception.ArgumentNullException "Item", ClassName & ".PrintSingle"
    End If
    
    If Item.Result = TestOutcome.Passed Then
        Debug.Print TestOutcomeEnum.ToString(Item.Result) & "; " & Item.Source
    Else
        Debug.Print TestOutcomeEnum.ToString(Item.Result) & "; " & Item.Description & "; " & Item.Source
    End If
    
End Sub


Private Sub ITestResultPrinter_PrintMany(ByVal Items As TestResults)
    PrintMany Items
End Sub


Private Sub ITestResultPrinter_PrintSingle(ByVal Item As TestResult)
    PrintSingle Item
End Sub

