Attribute VB_Name = "DateExt"
Option Explicit
'@Folder("Helper")


Public Function FirstDayOfMonth(ByVal Dt As Date) As Date
    FirstDayOfMonth = DateSerial(Year(Dt), Month(Dt), 1)
End Function


Public Function LastDayOfMonth(ByVal Dt As Date) As Date
    LastDayOfMonth = DateSerial(Year(Dt), Month(Dt), DaysInMonth(Dt))
End Function


Public Function DaysInMonth(ByVal Dt As Date) As Long
    
    Dim NewDt As Date: NewDt = VBA.DateAdd("m", 1, Dt)
    DaysInMonth = DateSerial(Year(NewDt), Month(NewDt), 1) - DateSerial(Year(Dt), Month(Dt), 1)
    
End Function


