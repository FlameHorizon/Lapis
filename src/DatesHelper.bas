Attribute VB_Name = "DatesHelper"
Option Explicit


Public Function FirstDayOfMonth(ByVal Dt As Date) As Date
    FirstDayOfMonth = DateSerial(Year(Dt), Month(Dt), 1)
End Function


Public Function LastDayOfMonth(ByVal Dt As Date) As Date
    LastDayOfMonth = DateSerial(Year(Dt), Month(Dt), DaysInMonth(Dt))
End Function


Public Function DaysInMonth(Dt As Date) As Long
    DaysInMonth = DateSerial(Year(Dt), Month(Dt) + 1, 1) - DateSerial(Year(Dt), Month(Dt), 1)
End Function


