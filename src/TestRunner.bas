Attribute VB_Name = "TestRunner"
Option Explicit
'@Folder("Tests")


Public Sub Start()

    ExUnit.ShowPassingTest = False
    ExUnit.Setup
    Debug.Print "Setting up a test runner..."
    
    StringHelperTests.Start
    TestResultsTests.Start
    ImmediatePrinterFormatterTests.Start
    ExUnitTests.Start
    TreeNodeTests.Start
    CollectionTests.Start

    ComparersTests.Start
    LambdaTests.Start

    ' By defualt, method will print test results in the Immediate window.
    ExUnit.PrintTestResults
    ExUnit.Teardown
    Debug.Print "Test runner finished..."
    
End Sub


