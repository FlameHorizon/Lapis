VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ExUnit")
Option Explicit

Private pSource As String
Private pOutcome As TestOutcome
Private pDescription As String


Public Property Get Source() As String
    Source = pSource
End Property


Public Property Let Source(ByVal Value As String)
    pSource = Value
End Property


Public Property Get Outcome() As TestOutcome
    Outcome = pOutcome
End Property


Public Property Let Outcome(ByVal Value As TestOutcome)
    pOutcome = Value
End Property


Public Property Get Description() As String
    Description = pDescription
End Property


Public Property Let Description(ByVal Value As String)
    pDescription = Value
End Property


