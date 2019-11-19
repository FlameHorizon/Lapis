Attribute VB_Name = "TestRunner"
Option Explicit
'@Folder("Tests")


Public Sub Start()

    ExUnit.ShowPassingTest = False
    StackTests.Start
    StringHelperTests.Start
    TreeNodeTests.Start
    
End Sub


