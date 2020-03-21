Attribute VB_Name = "ErrorCodeEnum"
'@Folder("Lapis")
Option Explicit

Private Const ModuleName As String = "ExceptionCodeEnum"

Public Enum ExceptionCode
    
    ArgumentNull = 513
    ArgumentOutOfRange = 514
    InvalidOperation = 515
    BaseException = 516
    Unhandled = 517
    ArgumentException = 518
    NotImplemented = 519
    DirectoryNotFound = 520
    Timeout = 523
    NotSupported = 524
    Overflow = 525
    MissingConfiguration = 526

End Enum


Public Function ToString(ByVal Value As ExceptionCode) As String

    Select Case Value
        Case ArgumentNull
            ToString = "ArgumentNull"
        
        Case ArgumentOutOfRange
            ToString = "ArgumentOutOfRange"
            
        Case InvalidOperation
            ToString = "InvalidOperation"
            
        Case BaseException
            ToString = "BaseException"
            
        Case Unhandled
            ToString = "Unhandled"
            
        Case ArgumentException
            ToString = "ArgumentException"
            
        Case NotImplemented
            ToString = "NotImplemented"
            
        Case DirectoryNotFound
            ToString = "DirectoryNotFound"
            
        Case Timeout
            ToString = "Timeout"
            
        Case NotSupported
            ToString = "NotSupported"
            
        Case Overflow
            ToString = "Overflow"
            
        Case MissingConfiguration
            ToString = "MissingConfiguration"
        
        Case Else
            Errors.OnArgumentOutOfRange "Value", ModuleName & ".ToString"
            
    End Select

End Function


Public Function TryToString(ByVal Value As ExceptionCode, ByRef Out As String) As Boolean
    
    On Error GoTo ErrHandler
    Out = ErrorCodeEnum.ToString(Value)
    TryToString = True
    
    Exit Function
ErrHandler:
    Select Case Err.Number
        Case ExceptionCode.ArgumentOutOfRange
            TryToString = False
        
        Case Else
            Errors.OnUnhandledError ModuleName & ".TryToString"
            
    End Select
    
End Function


