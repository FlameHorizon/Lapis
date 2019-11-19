VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Exception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder("Lapis")
'@PredeclaredId


' The exception that is thrown when a null reference (Nothing in Visual Basic)
' is passed to a method that does not accept it as a valid argument.
Public Sub ArgumentNullException(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 513, , "Value cannot be null." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
            
End Sub


' The exception that is thrown when the value of an argument is outside
' the allowable range of values as defined by the invoked method.
Public Sub ArgumentOutOfRangeException(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 514, , "Specified argument was out of the range of valid values." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
            
End Sub


' The exception that is thrown when a method call is invalid for the
' object's current state.
Public Sub InvalidOperationException(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 515, , "An exception of type InvalidOpertionException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
            
End Sub


' Represents errors that occur during application execution.
Public Sub BaseException(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 516, , "An exception of type Exception was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
            
End Sub


' Occurs when an exception is not caught.
Public Sub UnhandledException(ByVal Message As String)

    Err.Raise 517, , "An exception of type UnhandledException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine
            
End Sub


' The exception that is thrown when one of the arguments provided to a method is not valid.
Public Sub ArgumentException(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 518, , "An exception of type ArgumentException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
        
End Sub


' The exception that is thrown when a requested method or operation is not implemented.
Public Sub NotImplementedException(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 519, , "An exception of type NotImplementedException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
        
End Sub


' The exception that is thrown when a file or directory cannot be found.
Public Sub DirectoryNotFoundException(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 520, , "An exception of type DirectoryNotFoundException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
        
End Sub


' The exception that is thrown when Timeout is encountered
Public Sub TimeoutException(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 523, , "An exception of type Timeout was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
        
End Sub


' The exception that is thrown when an invoked method is not supported, or when
' there is an attempt to read, seek, or write to a stream that does not support
' the invoked functionality.
Public Sub NotSupportedException(ByVal ParamName As String, ByVal Message As String)
    
    Err.Raise 524, , "An exception of type NotSupportedException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
        
End Sub


' The exception that is thrown when an arithmetic, casting, or conversion
' operation in a checked context results in an overflow.

Public Sub OverflowException(ByVal ParamName As String, ByVal Message As String)
    
    Err.Raise 525, , "An exception of type OverflowException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
                    
End Sub


' The exception that is thrown when required arguments are missing
Public Sub MissingConfigurationException(ByVal ParamName As String, ByVal Message As String)

    Err.Raise 526, , "An exception of type MissingConfigurationException was thrown." & vbNewLine & vbNewLine & _
                    "Additional information: " & Message & vbNewLine & vbNewLine & _
                    "Parameter: " & ParamName
End Sub



