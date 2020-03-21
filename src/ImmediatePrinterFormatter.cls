VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImmediatePrinterFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ExUnit")
Option Explicit

Private Const ClassName As String = "ImmediatePrinterFormatter"


Public Function Format(ByVal Result As TestResult) As String
    
    If Result Is Nothing Then
        Errors.OnArgumentNull "Result", ClassName & ".PrintSingle"
    End If
    
    If Result.Outcome = TestOutcome.Passed Then
        Format = TestOutcomeEnum.ToString(Result.Outcome) & "; " & Result.Source
    Else
        Format = TestOutcomeEnum.ToString(Result.Outcome) & "; " & Result.Description & "; " & Result.Source
    End If
    
End Function
