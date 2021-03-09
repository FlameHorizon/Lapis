Attribute VB_Name = "SourceControl"
Option Explicit
'@Folder("Helper")

' Make sure you have selected option "Trust access to the VBA project object model"
' in the Developer Macro Settings (File - Options - Trust Center - Trust Center Settings).

' References which has to be included:
' - Microsoft Visual Basic for Application Extensibility 5.3

Private Const ModuleName As String = "SourceControl"

' Path to the folder where components will be saved.
Private pExportFolderPath As String

' Path to the folder form which components will be imported into project.
Private pImportFolderPath As String

' Indicates if empty components should be exported or not.
Private pExportEmptyComponents As Boolean

' Stores names of the modules which will not be imported to or updated in the VBProject
' while calling ExportProjectComponents or ImportProjectComponents methods.
Private pIgnoreImport As Dictionary '<Key: String, Value: String>


Public Property Get ExportEmptyComponents() As Boolean
    ExportEmptyComponents = pExportEmptyComponents
End Property


Public Property Let ExportEmptyComponents(ByVal Value As Boolean)
    pExportEmptyComponents = Value
End Property


' Exports and saves project's components, from Source workbook
' to the location which is specified in Path argument.
' If Source.VBProject is protected, throw an InvalidOperationErrors.
' If target path does not exists or if path does not points to a folder,
' throw an DirectoryNotFoundErrors.
Public Sub ExportProjectComponents(ByVal Source As VBProject, ByVal Path As String)

    Const MethodName = "ExportProjectComponents"

    If Source.Protection = vbext_pp_locked Then
        Errors.OnInvalidOperation "Source.Protection", _
                                  "The VBA project, in this workbook is protected. " & _
                                  "Therefore, it is not possible to export the components. " & _
                                  "Unlock your VBA project and try again. " & ModuleName & "." & MethodName
    End If
    
    With New Scripting.FileSystemObject
        If .FolderExists(Path) = False Then
            Errors.OnDirectoryNotFound "Path", ModuleName & "." & MethodName
        End If
    End With

    pExportFolderPath = NormalizePath(Path)
    
    Dim Cmp As VBComponent
    For Each Cmp In GetExportableComponents(Source.VBComponents)
        ExportComponent Cmp
    Next Cmp
    
End Sub


Private Property Get ExportableComponentsTypes() As Variant
    ExportableComponentsTypes = Array(vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_StdModule, vbext_ct_Document)
End Property


' Contains filtering logic for components.
Private Function GetExportableComponents(ByVal Source As VBIDE.VBComponents) As Collection '<VbComponents>
    
    Dim Output As New Collection
    Dim Cmp As VBIDE.VBComponent
    For Each Cmp In Source
        If IsExportable(Cmp) Then
            Output.Add Cmp
        End If
    Next Cmp
    
    Set GetExportableComponents = Output
    Set Cmp = Nothing
    Set Output = Nothing
    
End Function


Private Function IsExportable(ByVal Component As VBIDE.VBComponent) As Boolean
    
    ' Check if component is on the list of exportable components.
    If ArrayExt.Exists(Component.Type, ExportableComponentsTypes) = False Then
        IsExportable = False
        Exit Function
    End If
    
    ' Check if component does contain code.
    If IsComponentEmpty(Component) = False Then
        IsExportable = True
        Exit Function
    End If
    
    ' Check if we allow to export empty components.
    If pExportEmptyComponents = True Then
        IsExportable = True
        Exit Function
    End If
    
    IsExportable = False
    
End Function


' Indicates if component is empty by checking number of code lines.
' Files, which contains just Option Explicit will be counted as empty.
Private Function IsComponentEmpty(ByVal Source As VBIDE.VBComponent) As Boolean
    
    If Source.CodeModule.CountOfLines < 2 Then
        IsComponentEmpty = True
        
    ElseIf Source.CodeModule.CountOfLines = 2 Then
        Dim Ln1 As String: Ln1 = Source.CodeModule.Lines(1, 1)
        Dim Ln2 As String: Ln2 = Source.CodeModule.Lines(2, 1)
    
        IsComponentEmpty = (VBA.LCase$(Ln1) = "option explicit" And Ln2 = vbNullString)
    
    Else
        IsComponentEmpty = False
    End If
    
End Function


Private Sub ExportComponent(ByVal Component As VBIDE.VBComponent)
    
    Dim FileName As String: FileName = GetComponentFileName(Component)
    Dim ExportPath As String: ExportPath = pExportFolderPath & FileName
    Component.Export ExportPath
    
End Sub


' To avoid problems with saving components, add backslash
' at the end of folder path.
Private Function NormalizePath(ByVal Path As String) As String
    NormalizePath = Path & IIf(Path Like "*\", vbNullString, "\")
End Function


' Mapping between component type and the extension.
Private Property Get ComponentTypeToExtension() As Dictionary

    Dim Output As New Dictionary
    With Output
        .Add vbext_ct_ClassModule, "cls"
        .Add vbext_ct_MSForm, "frm"
        .Add vbext_ct_StdModule, "bas"
        .Add vbext_ct_Document, "doccls"
        .Add vbext_ct_ActiveXDesigner, "ocx"
    End With
    
    Set ComponentTypeToExtension = Output
    
End Property


' Imports modules to the specified Target workbook from the specified folder.
' Params
' Target - Workbook into which VBComponents will be imported to.
' Path - Path to the folder which contains modules to import.
' Throws an InvalidOperationException when Target's VBProject is locked.
' Throws an InvalidOperationException when IsImportEnabled is False.
' Throws an DirectoryNotFoundException when path to folder is not valid or folder doesn't exists.
Public Sub ImportProjectComponents(ByVal Target As Workbook, ByVal Path As String)

    Const MethodName = "ImportProjectComponents"

    If Target.VBProject.Protection = vbext_pp_locked Then
        Errors.OnInvalidOperation "Target.VBProject.Protection", _
                                  "The VBA project, in this workbook is protected " & _
                                  "therefor, it is not possible to import the components. " & _
                                  "Unlock your VBA project and try again. " & ModuleName & "." & MethodName
    End If

    If Tools.Fso.FolderExists(Path) = False Then
        Errors.OnDirectoryNotFound "Path", ModuleName & "." & MethodName
    End If
    
    ' This module can't be imported for the folder because this very module
    ' is directly responsible for importing components.
    Set pIgnoreImport = New Dictionary
    pIgnoreImport.Add ModuleName, ModuleName
    
    pImportFolderPath = NormalizePath(Path)

    ImportComponents Target, GetImportableComponents(Target)
    UpdateComponents Target, GetUpdatableComponents(Target)

End Sub


' Returns a Collection of paths to modules which can be imported into the VBProject
' based on the specified Path to a folder and specified Target's VBProject.
Private Function GetImportableComponents(ByVal Target As Workbook) As Collection '<String>

    Dim Output As New Collection '<String>
    Dim File As Scripting.File
    For Each File In GetImportableFiles(pImportFolderPath)
        If ComponentExists(Target, GetComponentName(File)) = False Then
            Output.Add File.Path
        End If
    Next File

    Set GetImportableComponents = Output
    Set File = Nothing
    Set Output = Nothing

End Function


Private Function GetImportableFiles(ByVal Path As String) As Collection '<File>

    Dim Output As New Collection '<File>
    Dim File As Scripting.File
    For Each File In Fso.GetFolder(Path).Files
        If IsImportable(File) Then
            Output.Add File
        End If
    Next File

    Set GetImportableFiles = Output

End Function


' Indicates if file should be imported into VBProject or not.
Private Function IsImportable(ByVal File As Scripting.File) As Boolean
    
    If pIgnoreImport.Exists(GetComponentName(File)) Then
        IsImportable = False
        Exit Function
    End If
    
    Dim Ext As String
    Ext = Fso.GetExtensionName(File.Path)
    
    If ArrayExt.Exists(Ext, SupportedImportExtensions) = False Then
        IsImportable = False
        Exit Function
    End If
    
    Dim Stream As New ADODB.Stream
    With Stream
        .Type = adTypeText
        .Charset = "ascii"
        .Open
        .LoadFromFile File.Path
    End With

    IsImportable = (Tools.LinesCount(Stream) > 0)

    Stream.Flush
    Stream.Close
    Set Stream = Nothing

End Function


Private Function SupportedImportExtensions() As Variant 'Array<String>
    
    SupportedImportExtensions = Array(ComponentTypeToExtension.Item(vbext_ct_ClassModule), _
                                      ComponentTypeToExtension.Item(vbext_ct_StdModule))
                                      
End Function


'Private Function ImportableComponentsTypes() As Variant
'    ImportableComponentsTypes = Array(vbext_ct_ClassModule, vbext_ct_StdModule)
'End Function


' Imports components based on to the specified Workbook using paths to components specified in Cmps collection.
' Param
' Paths: Collection <String>
Private Sub ImportComponents(ByVal Target As Workbook, ByVal Paths As Collection)

    Const MethodName = "ImportComponents"
    
    Dim Path As Variant '<String>
    Dim CmpName As String

    For Each Path In Paths
        CmpName = GetComponentName(Fso.GetFile(Path))
        Target.VBProject.VBComponents.Import Path
        Debug.Print ModuleName & "." & MethodName & ": Module imported - " & CmpName
    Next Path

End Sub


' Updates VBComponents based on the collection of paths to modules in the specified Target workbook.
' Params
' Target: Workbook. Place, where components will be updated.
' Paths: Collection <String>. List paths to source files.
Private Sub UpdateComponents(ByVal Target As Workbook, ByVal Paths As Collection)
    
    Const MethodName = "UpdateComponents"
    
    Dim Path As Variant '<String>
    Dim CmpName As String

    For Each Path In Paths
        CmpName = GetComponentName(Fso.GetFile(Path))
        RemoveComponent Target, CmpName
        Target.VBProject.VBComponents.Import Path
        Debug.Print ModuleName & "." & MethodName & ": Module updated - " & CmpName
    Next Path
    
End Sub


' Returns the name of the VB Component using VB_Name attribute based on the specified File.
Private Function GetComponentName(ByVal File As Scripting.File) As String
    
    Const MethodName As String = "GetComponentName"
    
    Dim Ln As Variant '<String>
    For Each Ln In Split(File.OpenAsTextStream.ReadAll, vbNewLine)
        If InStr(Ln, "Attribute VB_Name") Then
            GetComponentName = GetVbNameAttributeValueFrom(Ln)
            Exit Function
        End If
    Next Ln

    ' If, after walking through the file
    ' Attribute VB_Name was not found, thrown an Errors.
    Errors.OnArgumentError "File", _
                           "Attribute VB_Name was not found in the file. " _
                           & ModuleName & "." & MethodName

End Function


Private Function GetVbNameAttributeValueFrom(ByVal Line As String) As String
    
    Dim DblQuStart As Long
    Dim DblQuFinish As Long

    DblQuStart = InStr(1, Line, """", vbTextCompare)
    DblQuFinish = InStr(DblQuStart + 1, Line, """", vbTextCompare)
    GetVbNameAttributeValueFrom = Mid$(Line, DblQuStart + 1, DblQuFinish - DblQuStart - 1)
    
End Function


Private Sub RemoveComponent(ByVal Target As Workbook, ByVal Name As String)
    
    Dim Cmp As VBComponent
    Set Cmp = GetComponentByName(Target, Name)
        
    ' If name of the component is not changed before it is removed,
    ' new component will be imported with the "1" at the end and,
    ' we want to avoid that.
    Cmp.Name = GetReplacementName
    Target.VBProject.VBComponents.Remove Cmp

End Sub


' Returns a VBComponent object based on the name of the component and Source workbook.
Private Function GetComponentByName(ByVal Source As Workbook, ByVal Name As String) As VBComponent
    
    Dim Cmp As VBComponent
    For Each Cmp In Source.VBProject.VBComponents
        If Name = Cmp.Name Then
            Set GetComponentByName = Cmp
            Exit Function
        End If
    Next Cmp
    
    Errors.OnArgumentOutOfRange "Name", _
                                "Couldn't find the name component name [" & Name & "] in the Source workbook. " _
                                & ModuleName & ".GetComponentByName"

End Function


Private Function GetReplacementName() As String
    GetReplacementName = "A" & GetRandom10Digit & "_OLD"
End Function


Private Function GetRandom10Digit() As String
    
    Randomize
    GetRandom10Digit = Int(Rnd * 10000000000#)
    
End Function


' Contains logic when component should be updated based on the specified path to a folder.
' Returns a collection of paths to files which should be use to update modules.
Private Function GetUpdatableComponents(ByVal Target As Workbook) As Collection ' <String>

    Dim Output As New Collection '<String>
    Dim File As Scripting.File
    For Each File In GetImportableFiles(pImportFolderPath)
        If ComponentExists(Target, GetComponentName(File)) Then
            Output.Add File.Path
        End If
    Next File

    Set GetUpdatableComponents = Output
    Set File = Nothing
    Set Output = Nothing

End Function


Private Function ComponentExists(ByVal Source As Workbook, ByVal ComponentName As String) As Boolean
    
    On Error GoTo ErrHandler
    ' If this method didn't caused an error, this means compoment exists.
    GetComponentByName Source, ComponentName
    ComponentExists = True
    Exit Function
    
ErrHandler:
    If Err.Number = ErrorCode.ArgumentOutOfRange Then
        ComponentExists = False
        On Error GoTo -1
    Else
        ' Re-thorw error higher.
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
End Function


' Full name means - name of the component with an extension.
Private Function GetComponentFileName(ByVal Component As VBIDE.VBComponent) As String
    GetComponentFileName = Component.Name & "." & ComponentTypeToExtension.Item(Component.Type)
End Function


Public Sub UpdateProjectComponenets(ByVal Target As Workbook, ByVal Path As String)

    Const MethodName = "UpdateProjectComponenets"

    If Target.VBProject.Protection = vbext_pp_locked Then
        Errors.OnInvalidOperation "Target.VBProject.Protection", _
                                  "The VBA project, in this workbook is protected " & _
                                  "therefor, it is not possible to import the components. " & _
                                  "Unlock your VBA project and try again. " & ModuleName & "." & MethodName
    End If

    If Tools.Fso.FolderExists(Path) = False Then
        Errors.OnDirectoryNotFound "Path", ModuleName & "." & MethodName
    End If
    
    ' This module can't be imported for the folder because this very module
    ' is directly responsible for importing components.
    Set pIgnoreImport = New Dictionary
    pIgnoreImport.Add ModuleName, ModuleName
    
    pImportFolderPath = NormalizePath(Path)
    UpdateComponents Target, GetUpdatableComponents(Target)

End Sub
