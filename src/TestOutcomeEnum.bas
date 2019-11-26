Attribute VB_Name = "TestOutcomeEnum"
'@Folder("ExUnit")
Option Explicit

Public Enum TestOutcome
    
    Failed = 0
    Passed = 1

End Enum


Public Function ToString(ByVal Value As TestOutcome) As String
    
    Select Case Value
        Case TestOutcome.Failed
            ToString = "Failed"
        
        Case TestOutcome.Passed
            ToString = "Passed"
            
    End Select
    
End Function
