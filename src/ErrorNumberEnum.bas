Attribute VB_Name = "ErrorNumberEnum"
'@Folder("Lapis")
Option Explicit

' Error codes based on the following source:
' http://www.vba-market.com/list-of-vba-error-codes/

Private Const ModuleName As String = "ErrorNumberEnum"

Public Enum ErrorNumber
    
    ProcedureCallOrArgumentIsNotValid = 5
    Overflow = 6
    OutOfMemory = 7
    SubscriptOutOfRange = 9
    ThisArrayIsFixedOrTemporarilyLocked = 10
    DivisionByZero = 11
    TypeMismatch = 13
    OutOfStringSpace = 14
    ExpressionTooComplex = 16
    CantPerformRequestedOperation = 17
    UserInterruptOccurred = 18
    ResumeWithoutError = 20
    OutOfStackSpace = 28
    SubOrFunctionNotDefined = 35
    TooManyDllApplicationClients = 47
    ErrorLoadingDll = 48
    BadDllCallingConvention = 49
    InternalError = 51
    BadFileNameOrNumber = 52
    FileNotFound = 53
    BadFileMode = 54
    FileAlreadyOpen = 55
    DeviceIOError = 57
    FileAlreadyExists = 58
    BadRecordLength = 59
    DiskFull = 61
    InputPastEndOfFile = 62
    BadRecordNumber = 63
    TooManyFiles = 67
    DeviceUnavailable = 68
    PermissionDenied = 70
    DiskNotReady = 71
    CannotRenameWithDifferentDrive = 74
    PathFileAccessError = 75
    PathNotFound = 76
    ObjectVariableOrWithBlockVariableNotSet = 91
    ForLoopNotInitialized = 92
    InvalidPatternString = 93
    InvalidUseOfNull = 94
    ApplicationDefinedOrObjectDefinedError = 95
    UnableToSinkEventsOfObjectBecauseTheObjectIsAlreadyFiringEventsToTheMaximumNumberOfEventReceiversItSupports = 96
    CannotCallFriendFunctionOnObjectWhichIsNotAnInstanceOfDefiningClass = 97
    APropertyOrMethodCallCannotIncludeAReferenceToAPrivateObjectEitherAsAnArgumentOrAsAReturnValue = 98
    ClassDoesNotImplementTheSystemCollectionsICollectionInterface = 100
    SystemResourceOrDllCouldNotBeLoaded = 298
    CannotUseCharacterDeviceNamesInSpecifiedFileNames = 320
    FileFormatNotValid = 321
    CantCreateNecessaryTemporaryFile = 322
    FormatNotValidInResourceFile = 325
    DataValueNamedNotFound = 327
    IllegalParameterCantWriteArrays = 328
    CouldNotAccessSystemRegistry = 335
    ActiveXComponentNotCorrectlyRegistered = 336
    ActiveXComponentNotFound = 337
    ActiveXComponentDidNotRunCorrectly = 338
    ObjectAlreadyLoaded = 360
    CantLoadOrUnloadThisObject = 361
    ActiveXControlSpecifiedNotFound = 363
    ObjectWasUnloaded = 364
    UnableToUnloadWithinThisContext = 365
    TheSpecifiedFileIsOutOfDate = 368
    TheSpecifiedObjectCantBeUsedAsAnOwnerFormForShow = 371
    PropertyValueIsNotValid = 380
    PropertyArrayIndexIsNotValid = 381
    SetNotSupportedAtRunTime = 382
    SetNotSupported = 383
    NeedPropertyArrayIndex = 385
    SetNotPermitted = 387
    GetNotSupportedAtRunTime = 393
    GetNotSupported = 394
    FormAlreadyDisplayedCantShowModally = 400
    CodeMustCloseTopmostModalFormFirst = 402
    PropertyNotFound = 422
    PropertyOrMethodNotFound = 423
    ObjectRequired = 424
    CannotCreateActiveXComponent = 429
    ClassDoesNotSupportAutomationOrDoesNotSupportExpectedInterface = 430
    FileNameOrClassNameNotFoundDuringAutomationOperation = 432
    ObjectDoesntSupportThisPropertyOrMethod = 438
    AutomationError = 440
    ConnectionToTypeLibraryOrObjectLibraryForRemoteProcessHasBeenLost = 442
    AutomationObjectDoesNotHaveADefaultValue = 443
    ObjectDoesntSupportThisAction = 445
    ObjectDoesntSupportNamedArguments = 446
    ObjectDoesntSupportCurrentLocaleSetting = 447
    NamedArgumentNotFound = 448
    ArgumentNotOptional = 449
    WrongNumberOfArgumentsOrPropertyAssignmentNotValid = 450
    PropertyLetProcedureNotDefinedAndPropertyGetProcedureDidNotReturnAnObject = 451
    OrdinalIsNotValid = 452
    SpecifiedDllFunctionNotFound = 453
    CodeResourceNotFound = 454
    CodeResourceLockError = 455
    ThisKeyIsAlreadyAssociatedWithAnElementOfThisCollection = 457
    VariableUsesAnAutomationTypeNotSupportedInVisualBasic = 458
    ObjectOrClassDoesNotSupportTheSetOfEvents = 459
    ClipboardFormatIsNotValid = 460
    MethodOrDataMemberNotFound = 461
    TheRemoteServerMachineDoesNotExistOrIsUnavailable = 462
    ClassNotRegisteredOnLocalMachine = 463
    PictureIsNotValid = 481
    PrinterError = 482
    CannotSaveFileToTemp = 735
    SearchTextNotFound = 744
    ReplacementsTooLong = 746
    StopStatementEncountered = 999
    FeatureNotYetImplemented = 32768
    
End Enum


Public Function ToString(ByVal Value As ErrorNumber) As String
    
    Select Case Value
        Case ProcedureCallOrArgumentIsNotValid
            ToString = "Procedure call or argument is not valid (Visual Basic)"
        
        Case Overflow
            ToString = "Overflow (Visual Basic Run-Time Error)"
            
        Case OutOfMemory
            ToString = "Out of memory (Visual Basic Run-Time Error)"
            
        Case SubscriptOutOfRange
            ToString = "Subscript out of range (Visual Basic)"
            
        Case ThisArrayIsFixedOrTemporarilyLocked
            ToString = "This array is fixed or temporarily locked (Visual Basic)"
            
        Case DivisionByZero
            ToString = "Division by zero (Visual Basic Run-Time Error)"
            
        Case TypeMismatch
            ToString = "Type mismatch (Visual Basic)"
            
        Case OutOfStringSpace
            ToString = "Out of string space (Visual Basic)"
            
        Case ExpressionTooComplex
            ToString = "Expression too complex"
            
        Case CantPerformRequestedOperation
            ToString = "Can't perform requested operation (Visual Basic)"
            
        Case UserInterruptOccurred
            ToString = "User interrupt occurred"
            
        Case ResumeWithoutError
            ToString = "Resume without error"
            
        Case OutOfStackSpace
            ToString = "Out of stack space"
            
        Case SubOrFunctionNotDefined
            ToString = "Sub or function not defined"
            
        Case TooManyDllApplicationClients
            ToString = "Too many DLL application clients"
            
        Case ErrorLoadingDll
            ToString = "Error loading DLL"
            
        Case BadDllCallingConvention
            ToString = "Bad DLL calling convention"
            
        Case InternalError
            ToString = "Internal error (Visual Basic)"
            
        Case BadFileNameOrNumber
            ToString = "Bad file name or number"
            
        Case FileNotFound
            ToString = "File not found (Visual Basic Run-Time Error)"
            
        Case BadFileMode
            ToString = "Bad file mode"
            
        Case FileAlreadyOpen
            ToString = "File already open"
            
        Case DeviceIOError
            ToString = "Device I/O error"
            
        Case FileAlreadyExists
            ToString = "File already exists"
            
        Case BadRecordLength
            ToString = "Bad record length"
            
        Case DiskFull
            ToString = "Disk full (Visual Basic)"
            
        Case InputPastEndOfFile
            ToString = "Input past end of file"
            
        Case BadRecordNumber
            ToString = "Bad record number"
            
        Case TooManyFiles
            ToString = "Too many files"
            
        Case DeviceUnavailable
            ToString = "Device unavailable"
            
        Case PermissionDenied
            ToString = "Permission denied (Visual Basic)"
            
        Case DiskNotReady
            ToString = "Disk not ready"
            
        Case CannotRenameWithDifferentDrive
            ToString = "Cannot rename with different drive"
            
        Case PathFileAccessError
            ToString = "Path/File access error"
            
        Case PathNotFound
            ToString = "Path not found"
            
        Case ObjectVariableOrWithBlockVariableNotSet
            ToString = "Object variable or With block variable not set"
            
        Case ForLoopNotInitialized
            ToString = "For loop not initialized (Visual Basic)"
            
        Case InvalidPatternString
            ToString = "Invalid pattern string"
            
        Case InvalidUseOfNull
            ToString = "Invalid use of Null (obsolete, no longer used)"
                
        Case ApplicationDefinedOrObjectDefinedError
            ToString = "Application-defined or object-defined error"
            
        Case UnableToSinkEventsOfObjectBecauseTheObjectIsAlreadyFiringEventsToTheMaximumNumberOfEventReceiversItSupports
            ToString = "Unable to sink events of object because the object is already firing events to the maximum number of event receivers it supports"
            
        Case CannotCallFriendFunctionOnObjectWhichIsNotAnInstanceOfDefiningClass
            ToString = "Cannot call friend function on object which is not aninstance of defining class"
            
        Case APropertyOrMethodCallCannotIncludeAReferenceToAPrivateObjectEitherAsAnArgumentOrAsAReturnValue
            ToString = "A property or method call cannot include a reference to a private object, either as an argument or as a return value"
            
        Case ClassDoesNotImplementTheSystemCollectionsICollectionInterface
            ToString = "Class does not implement the System.Collections.ICollection interface"
            
        Case SystemResourceOrDllCouldNotBeLoaded
            ToString = "System resource or DLL could not be loaded"
            
        Case CannotUseCharacterDeviceNamesInSpecifiedFileNames
            ToString = "Cannot use character device names in specified file names"
            
        Case FileFormatNotValid
            ToString = "File format not valid"
            
        Case CantCreateNecessaryTemporaryFile
            ToString = "Can't create necessary temporary file"
            
        Case FormatNotValidInResourceFile
            ToString = "Format not valid in resource file"
            
        Case DataValueNamedNotFound
            ToString = "Data value named not found"
            
        Case IllegalParameterCantWriteArrays
            ToString = "Illegal parameter; can't write arrays"
            
        Case CouldNotAccessSystemRegistry
            ToString = "Could not access system registry"
            
        Case ActiveXComponentNotCorrectlyRegistered
            ToString = "ActiveX component not correctly registered"
            
        Case ActiveXComponentNotFound
            ToString = "ActiveX Component not found"
            
        Case ActiveXComponentDidNotRunCorrectly
            ToString = "ActiveX Component did not run correctly"
            
        Case ObjectAlreadyLoaded
            ToString = "Object already loaded"
            
        Case CantLoadOrUnloadThisObject
            ToString = "Can't load or unload this object"
            
        Case ActiveXControlSpecifiedNotFound
            ToString = "ActiveX Control specified not found"
            
        Case ObjectWasUnloaded
            ToString = "Object was unloaded"
            
        Case UnableToUnloadWithinThisContext
            ToString = "Unable to unload within this context"
            
        Case TheSpecifiedFileIsOutOfDate
            ToString = "The specified file is out of date"
            
        Case TheSpecifiedObjectCantBeUsedAsAnOwnerFormForShow
            ToString = "The specified object can't be used as an owner form for Show"
            
        Case PropertyValueIsNotValid
            ToString = "Property value is not valid"
            
        Case PropertyArrayIndexIsNotValid
            ToString = "Property array index is not valid"
            
        Case SetNotSupportedAtRunTime
            ToString = "Set not supported at run time"
            
        Case SetNotSupported
            ToString = "Set not supported (read-only property)"
            
        Case NeedPropertyArrayIndex
            ToString = "Need property array index"
            
        Case SetNotPermitted
            ToString = "Set not permitted"
            
        Case GetNotSupportedAtRunTime
            ToString = "Get not supported at run time"
            
        Case GetNotSupported
            ToString = "Get not supported (write-only property)"
            
        Case FormAlreadyDisplayedCantShowModally
            ToString = "Form already displayed; can't show modally"
            
        Case CodeMustCloseTopmostModalFormFirst
            ToString = "Code must close topmost modal form first"
            
        Case PropertyNotFound
            ToString = "Property not found"
            
        Case PropertyOrMethodNotFound
            ToString = "Property or method not found"
            
        Case ObjectRequired
            ToString = "Object required (Visual Basic)"
            
        Case CannotCreateActiveXComponent
            ToString = "Cannot create ActiveX Component"
            
        Case ClassDoesNotSupportAutomationOrDoesNotSupportExpectedInterface
            ToString = "Class does not support Automation or does not support expected interface"
            
        Case FileNameOrClassNameNotFoundDuringAutomationOperation
            ToString = "File name or class name not found during Automation operation (Visual Basic)"
            
        Case ObjectDoesntSupportThisPropertyOrMethod
            ToString = "Object doesn't support this property or method (Visual Basic)"
            
        Case AutomationError
            ToString = "Automation error"
            
        Case ConnectionToTypeLibraryOrObjectLibraryForRemoteProcessHasBeenLost
            ToString = "Connection to type library or object library for remote process has been lost"
            
        Case AutomationObjectDoesNotHaveADefaultValue
            ToString = "Automation object does not have a default value"
            
        Case ObjectDoesntSupportThisAction
            ToString = "Object doesn't support this action (Visual Basic)"
            
        Case ObjectDoesntSupportNamedArguments
            ToString = "Object doesn't support named arguments"
            
        Case ObjectDoesntSupportCurrentLocaleSetting
            ToString = "Object doesn't support current locale setting (Visual Basic)"
            
        Case NamedArgumentNotFound
            ToString = "Named argument not found (Visual Basic)"
            
        Case ArgumentNotOptional
            ToString = "Argument not optional (Visual Basic)"
            
        Case WrongNumberOfArgumentsOrPropertyAssignmentNotValid
            ToString = "Wrong number of arguments or property assignment not valid"
            
        Case PropertyLetProcedureNotDefinedAndPropertyGetProcedureDidNotReturnAnObject
            ToString = "Property let procedure not defined and property get procedure did not return an object"
            
        Case OrdinalIsNotValid
            ToString = "Ordinal is not valid"
            
        Case SpecifiedDllFunctionNotFound
            ToString = "Specified DLL function not found"
            
        Case CodeResourceNotFound
            ToString = "Code resource not found"
            
        Case CodeResourceLockError
            ToString = "Code resource lock error"
            
        Case ThisKeyIsAlreadyAssociatedWithAnElementOfThisCollection
            ToString = "This key is already associated with an element of this collection"
            
        Case VariableUsesAnAutomationTypeNotSupportedInVisualBasic
            ToString = "Variable uses an Automation type not supported in Visual Basic"
            
        Case ObjectOrClassDoesNotSupportTheSetOfEvents
            ToString = "Object or class does not support the set of events"
            
        Case ClipboardFormatIsNotValid
            ToString = "Clipboard format is not valid"
            
        Case MethodOrDataMemberNotFound
            ToString = "Method or data member not found"
            
        Case TheRemoteServerMachineDoesNotExistOrIsUnavailable
            ToString = "The remote server machine does not exist or is unavailable (Visual Basic)"
            
        Case ClassNotRegisteredOnLocalMachine
            ToString = "Class not registered on local machine"
            
        Case PictureIsNotValid
            ToString = "Picture is not valid"
            
        Case PrinterError
            ToString = "Printer error"
            
        Case CannotSaveFileToTemp
            ToString = "Cannot save file to TEMP"
            
        Case SearchTextNotFound
            ToString = "Search text not found"
            
        Case ReplacementsTooLong
            ToString = "Replacements too long"
            
        Case StopStatementEncountered
            ToString = "Stop statement encountered"
            
        Case FeatureNotYetImplemented
            ToString = "Feature not yet implemented"
            
        Case Else
            Errors.OnArgumentOutOfRange vbNullString, ModuleName & ".ToString"
        
    End Select
    
End Function


Public Function TryToString(ByVal Value As ErrorNumber, ByRef Out As String) As Boolean
    
    On Error GoTo ErrHandler
    Out = ErrorNumberEnum.ToString(Value)
    TryToString = True
    
    Exit Function
ErrHandler:
    Select Case Err.Number
        Case ErrorCode.ArgumentOutOfRange
            TryToString = False
        
        Case Else
            Errors.OnUnhandledError ModuleName & ".TryToString"
            
    End Select
    
End Function


