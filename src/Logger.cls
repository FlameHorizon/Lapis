VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Requires reference: Microsoft Scripting Runtime
' Requires reference: Microsoft ActiveX Data Objects X.X Library <- only if you want to log informations to database

Private Const ClassName     As String = "Logger"
Private Const FilePrefix    As String = "Log_"
Private Const LogsSheetName As String = "Logs"
Private Const DefaultLogsTableName As String = "Logs"
Private Const DefaultSharePointListName As String = "RobotLogs"
Private Const RobotColumnName   As String = "ROBOT_NAME"
Private Const MessageColumnName As String = "MESSAGE"
Private Const DateColumnName    As String = "CREATED_ON"
Private Const LevelColumnName   As String = "LOG_LEVEL"
Private Const ModuleColumnName  As String = "MODULE_NAME"
Private Const MethodColumnName  As String = "METHOD_NAME"
Private Const ProcessColumnName As String = "PROCESS_ID"

Public Enum Level
    Normal = 0
    Verbose = 1
    Error = 2
    Warning = 3
End Enum


Private pLogToExcelSheetFlag As Boolean
Private pLogToImmediateWindowFlag As Boolean
Private pLogToFileFlag       As Boolean
Private pLogToDatabaseFlag   As Boolean
Private pLogToSharePointFlag As Boolean
Private pInitiatedFlag       As Boolean
Private pVerboseFlag         As Boolean

Private pLogsWorkbook  As Excel.Workbook
Private pLogsWorksheet As Excel.Worksheet
Private pFso           As FileSystemObject
Private pLogsFilePath  As String
Private pProcessID     As String
Private pFileStream    As TextStream

' Database logging variables
Private pDbConnection  As ADODB.Connection
Private pLogsTable     As ADODB.Recordset
Private pMessageColumn As ADODB.field
Private pDateColumn    As ADODB.field
Private pLevelColumn   As ADODB.field
Private pModuleColumn  As ADODB.field
Private pMethodColumn  As ADODB.field
Private pProcessColumn As ADODB.field
Private pLogsTableName As String


Private Sub Class_Terminate()

    pInitiatedFlag = False
    
    If Not pFileStream Is Nothing Then
        pFileStream.Close
        Set pFileStream = Nothing
    End If
    
    If Not pLogsTable Is Nothing Then
        pLogsTable.Close
        Set pLogsTable = Nothing
    End If
End Sub


' This method should be called when Workbook is opened, or automation main macro is starting
' It initiates all necessary connections, based on selected log type
Public Sub Start()

    Const MethodName = "Start"
    On Error GoTo ErrHandler

    If pInitiatedFlag = False Then
        If pLogToExcelSheetFlag = True Then
            Set pLogsWorksheet = InitiateExcelSheet(pLogsWorkbook)
        End If
        
        If pLogToImmediateWindowFlag = True Then
            ' Nothing to do, this is a simple Debug.Print
        End If
        
        If pLogToFileFlag = True Then
            InitiateLogFile
        End If
        
        If pLogToDatabaseFlag = True Then
            If pLogsTableName = vbNullString Then pLogsTableName = DefaultLogsTableName
            Set pLogsTable = InitiateRecordset(pDbConnection)
        End If
        
        pInitiatedFlag = True
        Me.Create "Logger initiated.", ClassName, MethodName, Verbose
    End If
    
CleanUp:
    Exit Sub

ErrHandler:
    Debug.Print ArrangeMessage(Err.Description, ClassName, MethodName, Level.Error)
    Debug.Assert False
    GoTo CleanUp
End Sub


Public Sub Create(Message As String, ModuleName As String, SourceMethodName As String, _
                  Optional LogLevel As Level = Level.Normal)
                  
    Const MethodName = "Create"
    On Error GoTo ErrHandler

    If pInitiatedFlag = True Then
        If LogLevel = Verbose And pVerboseFlag = True _
           Or LogLevel <> Verbose Then
        
            Dim messageToLog As String
            messageToLog = ArrangeMessage(Message, ModuleName, SourceMethodName, LogLevel)
            
            If pLogToImmediateWindowFlag = True Then
                Debug.Print messageToLog
            End If
            
            If pLogToExcelSheetFlag = True Then
                LogToExcelSheet Message, ModuleName, SourceMethodName, LogLevel
            End If
            
            If pLogToFileFlag = True Then
                LogToFile messageToLog
            End If
            
            If pLogToDatabaseFlag = True Then
                LogToDatabase Message, ModuleName, SourceMethodName, LogLevel
            End If
            
        End If
        
    Else
        Debug.Print ArrangeMessage("Logger is not initiated. Use Start method", _
                                   ClassName, MethodName, Level.Error)
    End If

CleanUp:
    Exit Sub

ErrHandler:
    Debug.Print ArrangeMessage(Err.Description, ClassName, MethodName, Level.Error)
    Debug.Assert False
    GoTo CleanUp
End Sub


Private Function ArrangeMessage(Message As String, ModuleName As String, _
                                SourceMethodName As String, LogLevel As Level) As String

    ArrangeMessage = "[" & LogLevelName(LogLevel) & "]" _
                                                & "[" & Format$(Now(), "yyyy.mm.dd hh:nn") & "]"

    If pProcessID <> vbNullString Then
        ArrangeMessage = ArrangeMessage & "[" & pProcessID & "]"
    End If

    ArrangeMessage = ArrangeMessage & "[" & ModuleName & "." & SourceMethodName & "] " & _
                     vbTab & Message
End Function


' Get the name of used log level
Private Function LogLevelName(PassedLevel As Level) As String

    Select Case PassedLevel
        Case Level.Normal
            LogLevelName = "Log"
        Case Level.Verbose
            LogLevelName = "Deb"
        Case Level.Error
            LogLevelName = "Err"
        Case Level.Warning
            LogLevelName = "Warn"
        Case Else
            LogLevelName = "UNKNOWN"
    End Select
End Function


Private Sub InitiateLogFile()

    Const MethodName = "InitiateLogFile"

    If pLogsFilePath = vbNullString Then
        Exception.DirectoryNotFoundException "pLogsFilePath", pLogsFilePath
    End If
    
    Set pFso = New FileSystemObject
    Set pFileStream = CreateTextStream()
End Sub


Private Function CreateLogPath() As String

    Dim FileName As String
    FileName = FilePrefix & Format$(Now(), "yyyy_mm_dd_hhnn") & ".txt"
    
    CreateLogPath = pFso.BuildPath(pLogsFilePath, FileName)
End Function


Private Function CreateTextStream() As TextStream

    Dim LogPathWithName As String
    LogPathWithName = CreateLogPath()
    
    Const CreateFlag As Boolean = True
    Set CreateTextStream = pFso.OpenTextFile(LogPathWithName, ForWriting, CreateFlag, TristateTrue) ' TristateTrue <- Unicode coding
End Function


Private Sub LogToFile(Message As String)

    Const MethodName = "LogToFile"
    On Error GoTo ErrHandler
    
    pFileStream.WriteLine Message

CleanUp:
    Exit Sub

ErrHandler:
    Debug.Print ArrangeMessage(Err.Description, ClassName, MethodName, Level.Error)
    Debug.Assert False
    GoTo CleanUp
End Sub


Private Function InitiateExcelSheet(PassedWorkbook As Excel.Workbook) As Excel.Worksheet

    Dim LogSheetExist As Boolean: LogSheetExist = False
    Dim CurrentSheet As Excel.Worksheet
    
    For Each CurrentSheet In PassedWorkbook.Sheets
        If CurrentSheet.Name = LogsSheetName Then
            LogSheetExist = True
            Set InitiateExcelSheet = CurrentSheet
        End If
    Next
    
    ' Create new logs worksheet if none exists
    If LogSheetExist = False Then
        Set InitiateExcelSheet = PassedWorkbook.Sheets.Add
        InitiateExcelSheet.Name = LogsSheetName
    End If
End Function


Private Sub LogToExcelSheet(Message As String, ModuleName As String, _
                            SourceMethodName As String, LogLevel As Level)
                          
    Const MethodName = "logToExcelSheet"
    On Error GoTo ErrHandler

    Dim lastRow As Long
    ' Ctrl + Shift + End
    lastRow = pLogsWorksheet.Cells(pLogsWorksheet.Rows.Count, "A").End(xlUp).Row
    With pLogsWorksheet
        .Cells(lastRow + 1, 1).Value = LogLevelName(LogLevel)
        .Cells(lastRow + 1, 2).Value = ModuleName
        .Cells(lastRow + 1, 3).Value = SourceMethodName
        .Cells(lastRow + 1, 4).Value = Message
        .Cells(lastRow + 1, 5).Value = Now()
        .Cells(lastRow + 1, 6).Value = pProcessID
    End With
    
CleanUp:
    Exit Sub

ErrHandler:
    Debug.Print ArrangeMessage(Err.Description, ClassName, MethodName, Level.Error)
    Debug.Assert False
    GoTo CleanUp
End Sub


Private Function InitiateRecordset(ByRef Conn As ADODB.Connection) As ADODB.Recordset

    Const MethodName = "InitiateRecordset"
    On Error GoTo ErrHandler

    Set pLogsTable = New ADODB.Recordset
    With pLogsTable
        Set .ActiveConnection = Conn
        .LockType = adLockOptimistic
        .CursorType = adOpenKeyset
        .CursorLocation = adUseServer
        ' .Index = "ID" ' <- If more speed is required + seek functionality
        .Open pLogsTableName, Options:=adCmdTableDirect
        
        Set pMessageColumn = .Fields.Item(MessageColumnName)
        Set pDateColumn = .Fields.Item(DateColumnName)
        Set pLevelColumn = .Fields.Item(LevelColumnName)
        Set pModuleColumn = .Fields.Item(ModuleColumnName)
        Set pMethodColumn = .Fields.Item(MethodColumnName)
        Set pProcessColumn = .Fields.Item(ProcessColumnName)
    End With
    
    Set InitiateRecordset = pLogsTable
    
CleanUp:
    Exit Function

ErrHandler:
    Select Case Err.Number
        Case -2147217865
            ' This means selected database doesn't have logs table
            If MsAccessCreateLogsTable(pLogsTableName) = True Then
                Resume 0
            End If
        Case 3265 ' Item cannot be found in the collection corresponding to the requested name or ordinal.
            Const ErrorMessage As String = "You have probably changed log table columns. " & _
            "Either delete old logs table so automation can recreate it, " & _
            "or revert database column changes."

            Err.Raise Err.Number, MethodName, ArrangeMessage(ErrorMessage, ClassName, MethodName, Level.Error)
        Case Else
            Debug.Print Err.Number
            ' This type of construction is used when we don't want to continue work if any problems occur
            ' Additionally we get infomration, which method returned this error.
            Err.Raise Err.Number, MethodName, ArrangeMessage(Err.Description, ClassName, MethodName, Level.Error)
    End Select
End Function


Private Function MsAccessCreateLogsTable(TableName As String) As Boolean
    
    Const MethodName = "MsAccessCreateLogsTable"
    On Error GoTo ErrHandler

    Dim CreateSql As String: CreateSql = _
        "CREATE TABLE " & TableName & " (" _
      & "ID Autoincrement, " _
      & MessageColumnName & " LONGTEXT, " _
      & DateColumnName & " DATETIME, " _
      & LevelColumnName & " VARCHAR, " _
      & ModuleColumnName & " VARCHAR, " _
      & MethodColumnName & " VARCHAR, " _
      & ProcessColumnName & " VARCHAR " _
      & ");"
        
    With New ADODB.Command
        Set .ActiveConnection = pDbConnection
        .CommandText = CreateSql
        .Execute
    End With
    MsAccessCreateLogsTable = True
    
CleanUp:
    Exit Function

ErrHandler:
    ' This type of construction is used when we don't want to continue work if any problems occur
    ' Additionally we get infomration, which method returned this error.
    Err.Raise Err.Number, MethodName, ArrangeMessage(Err.Description, ClassName, MethodName, Level.Error)
End Function


Private Sub LogToDatabase(Message As String, ModuleName As String, _
                          SourceMethodName As String, LogLevel As Level)
                          
    Const MethodName = "LogToDatabase"
    On Error GoTo ErrHandler
    
    With pLogsTable
        .AddNew
        pMessageColumn.Value = Message
        pDateColumn.Value = Now()
        pLevelColumn.Value = LogLevelName(LogLevel)
        pModuleColumn.Value = ModuleName
        pMethodColumn.Value = SourceMethodName
        pProcessColumn.Value = pProcessID
        .Update
    End With

CleanUp:
    Exit Sub

ErrHandler:
    Debug.Print ArrangeMessage(Err.Description, ClassName, MethodName, Level.Error)
    Debug.Assert False
    GoTo CleanUp
End Sub


Public Property Let WriteToExcel(Value As Boolean)

    pLogToExcelSheetFlag = Value
End Property


Public Property Set WorkbookWithLogs(Value As Excel.Workbook)

    Set pLogsWorkbook = Value
End Property


Public Property Let WriteToFile(Value As Boolean)

    pLogToFileFlag = Value
End Property


Public Property Let FilePath(Value As String)

    pLogsFilePath = Value
End Property


Public Property Let WriteToDatabase(Value As Boolean)

    pLogToDatabaseFlag = Value
End Property


Public Property Set DatabaseConnection(Value As ADODB.Connection)
    
    Set pDbConnection = Value
End Property


Public Property Let WriteToImmediate(Value As Boolean)

    pLogToImmediateWindowFlag = Value
End Property


Public Property Let VerboseLogging(Value As Boolean)

    pVerboseFlag = Value
End Property


Public Property Let WriteToSharePoint(Value As Boolean)

    pLogToSharePointFlag = Value
End Property


Public Property Let CustomLogsTableName(Value As String)

    pLogsTableName = Value
End Property


Public Property Let CurrentProcessID(Value As String)
    
    pProcessID = Value
End Property



