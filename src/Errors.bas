Attribute VB_Name = "Errors"
'@Folder("Lapis")
Option Private Module
Option Explicit


' The error that is thrown when a null reference (Nothing in Visual Basic)
' is passed to a method that does not accept it as a valid argument.
'@Ignore ProcedureNotUsed
Public Sub OnArgumentNull(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 513, , "Value cannot be null." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
            
End Sub


' The error that is thrown when the value of an argument is outside
' the allowable range of values as defined by the invoked method.
'@Ignore ProcedureNotUsed
Public Sub OnArgumentOutOfRange(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 514, , "Specified argument was out of the range of valid values." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
            
End Sub


' The error that is thrown when a method call is invalid for the
' object's current state.
'@Ignore ProcedureNotUsed
Public Sub OnInvalidOperation(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 515, , "An error of type InvalidOpertionException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
            
End Sub


' Represents errors that occur during application execution.
'@Ignore ProcedureNotUsed
Public Sub OnBaseError(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 516, , "An error of type error was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
            
End Sub


' Occurs when an error is not caught.
'@Ignore ProcedureNotUsed
Public Sub OnUnhandledError(ByVal Message As String)

    Err.Raise 517, , "An error of type UnhandledException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine
            
End Sub


' The error that is thrown when one of the arguments provided to a method is not valid.
'@Ignore ProcedureNotUsed
Public Sub OnArgumentError(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 518, , "An error of type ArgumentException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
        
End Sub


' The error that is thrown when a requested method or operation is not implemented.
'@Ignore ProcedureNotUsed
Public Sub OnNotImplemented(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 519, , "An error of type NotImplementedException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
        
End Sub


' The error that is thrown when a file or directory cannot be found.
'@Ignore ProcedureNotUsed
Public Sub OnDirectoryNotFound(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 520, , "An error of type DirectoryNotFoundException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
        
End Sub


' The error that is thrown when Timeout is encountered
'@Ignore ProcedureNotUsed
Public Sub OnTimeout(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 523, , "An error of type Timeout was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
        
End Sub


' The error that is thrown when an invoked method is not supported, or when
' there is an attempt to read, seek, or write to a stream that does not support
' the invoked functionality.
'@Ignore ProcedureNotUsed
Public Sub OnNotSupported(ByVal ParamName As String, ByVal Message As String)
    
    Err.Raise 524, , "An error of type NotSupportedException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
        
End Sub


' The error that is thrown when an arithmetic, casting, or conversion
' operation in a checked context results in an overflow.
'@Ignore ProcedureNotUsed
Public Sub OnOverflow(ByVal ParamName As String, ByVal Message As String)
    
    Err.Raise 525, , "An error of type OverflowException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
                    
End Sub


' The error that is thrown when required arguments are missing
'@Ignore ProcedureNotUsed
Public Sub OnMissingConfiguration(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 526, , "An error of type MissingConfigurationException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
End Sub



