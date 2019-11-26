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
    ImmediatePrinterFormatterTests.Start

    ' By defualt, method will print test results in the Immediate window.
    ExUnit.PrintTestResults
    ExUnit.Teardown
    
End Sub


