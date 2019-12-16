VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestResultImmediatePrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ExUnit")
Option Explicit

Implements ITestResultPrinter

Private Const ClassName As String = "TestResultImmediatePrinter"


Public Sub PrintMany(ByVal Items As TestResults)
    
    Dim Key As Variant '<TestResult>
    For Each Key In Items.NewEnum
        PrintSingle Items.Item(Key)
    Next Key
    
End Sub


Public Sub PrintSingle(ByVal Item As TestResult)
    
    If Item Is Nothing Then
        Errors.OnArgumentNull "Item", ClassName & ".PrintSingle"
    End If
    Dim Formatter As New ImmediatePrinterFormatter
    Debug.Print Formatter.Format(Item)
    
End Sub


Private Sub ITestResultPrinter_PrintMany(ByVal Items As TestResults)
    PrintMany Items
End Sub


Private Sub ITestResultPrinter_PrintSingle(ByVal Item As TestResult)
    PrintSingle Item
End Sub


