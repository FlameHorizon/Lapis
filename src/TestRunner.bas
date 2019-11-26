Attribute VB_Name = "TestRunner"
Option Explicit
'@Folder("Tests")


Public Sub Start()

    ExUnit.ShowPassingTest = False
    ExUnit.Setup
    
    StackTests.Start
    StringHelperTests.Start
    TreeNodeTests.Start
    TestResultsTests.Start

    ExUnit.PrintTestResults New TestResultImmediatePrinter
    ExUnit.Teardown
    
End Sub


