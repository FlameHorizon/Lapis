Attribute VB_Name = "SourceControlH"
Option Explicit
'@Folder("Helper")

' Make sure you have selected option "Trust access to the VBA project object model"
' in the Developer Macro Settings (File - Options - Trust Center - Trust Center Settings).

' References which has to be included:
' - Microsoft Visual Basic for Application Extensibility 5.3

' How to hook up exporting project's components to the
' AfterSave/BeforeSave?

' Place this snippet inside ThisWorkbook component.
'   SourceControlHelper.ExportProjectComponents ExportLocation

Private Const ModuleName As String = "SourceControlHelper"

' Returns a path to the folder where every component
' of this VBA project will be saved.
Private pOutputFolder As String
Private pExportEmptyComponents As Boolean

' Stores names of the modules which will nor be imported to or updated in the VBProject
' while calling ExportProjectComponents or ImportProjectComponents methods.
Private pIgnoreImport As Dictionary '<Key: String, Value: String>

' Exports and saves project's components, from this workbook
' to the location which is specified in Path argument.


' If VBProject is protected, throw an InvalidOperationException.
' If target path does not exists or if path does not points to a folder,
' throw an DirectoryNotFoundException.
' If method is called when exporting is disabled, throw an InvalidOperationException.
Public Sub ExportProjectComponents(ByRef Source As Workbook, ByVal Path As String)

    Const MethodName = "ExportProjectComponents"

    If Source.VBProject.Protection = vbext_pp_locked Then
        Exception.InvalidOperationException "Source.VBProject.Protection", _
                                            "The VBA project, in this workbook is protected " & _
                                            "therefor, it is not possible to export the components. " & _
                                            "Unlock your VBA project and try again. " & ModuleName & "." & MethodName
    End If

    If Tools.Fso.FolderExists(Path) = False Then
        Exception.DirectoryNotFoundException "Path", ModuleName & "." & MethodName
    End If

    pOutputFolder = NormalizePath(Path)
    ExportComponents GetExportableComponents(Source.VBProject.VBComponents)
    
End Sub


Private Property Get ExportableComponentsTypes() As Variant
    ExportableComponentsTypes = Array(vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_StdModule, vbext_ct_Document)
End Property


' Contains filtering logic for components.
Private Function GetExportableComponents(ByRef Source As VBIDE.VBComponents) As Collection
    
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


Private Function IsExportable(ByRef Cmp As VBIDE.VBComponent) As Boolean
    
    ' Check if component is on the list of exportable components.
    If ArrayH.Exists(Cmp.Type, ExportableComponentsTypes) = False Then
        IsExportable = False
        Exit Function
    End If
    
    ' Check if component does contain code.
    If IsComponentEmpty(Cmp) = False Then
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
Private Function IsComponentEmpty(ByRef Source As VBIDE.VBComponent) As Boolean
    
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


Private Sub ExportComponents(ByRef Components As Collection)
    
    Dim Cmp As VBIDE.VBComponent
    For Each Cmp In Components
        ExportComponent Cmp
    Next Cmp
    
    Set Cmp = Nothing
    
End Sub


Private Sub ExportComponent(ByRef Component As VBIDE.VBComponent)
    
    Dim CmpFullName As String: CmpFullName = GetFullName(Component)
    Component.Export pOutputFolder & CmpFullName
    
End Sub


' To avoid problems with saving components, add backslash
' at the end of folder path.
Private Function NormalizePath(ByVal Path As String) As String
    NormalizePath = Path & IIf(Path Like "*\", vbNullString, "\")
End Function


' Mapping between component type and the extension.
Private Property Get ComponentTypeToExtension() As Dictionary

    Dim Output As New Dictionary
    Output.Add vbext_ct_ClassModule, "cls"
    Output.Add vbext_ct_MSForm, "frm"
    Output.Add vbext_ct_StdModule, "bas"
    Output.Add vbext_ct_Document, "doccls"
    Output.Add vbext_ct_ActiveXDesigner, "ocx"
    
    Set ComponentTypeToExtension = Output
    Set Output = Nothing
    
End Property


' Imports modules to the specified Target workbook from the specified folder.
' Params
' Target - Workbook into which VBComponents will be imported to.
' Path - Path to the folder which contains modules to import.
' Throws an InvalidOperationException when Target's VBProject is locked.
' Throws an InvalidOperationException when IsImportEnabled is False.
' Throws an DirectoryNotFoundException when path to folder is not valid or folder doesn't exists.
Public Sub ImportProjectComponents(ByRef Target As Workbook, ByVal Path As String)

    Const MethodName = "ImportProjectComponents"

    If Target.VBProject.Protection = vbext_pp_locked Then
        Exception.InvalidOperationException "Target.VBProject.Protection", _
                                            "The VBA project, in this workbook is protected " & _
                                            "therefor, it is not possible to import the components. " & _
                                            "Unlock your VBA project and try again. " & ModuleName & "." & MethodName
    End If

    If Tools.Fso.FolderExists(Path) = False Then
        Exception.DirectoryNotFoundException "Path", ModuleName & "." & MethodName
    End If
    
    ' This module can't be imported for the folder because this very module
    ' is directly responsible for importing components.
    Set pIgnoreImport = New Dictionary
    pIgnoreImport.Add ModuleName, ModuleName

    pOutputFolder = NormalizePath(Path)
    ImportComponentsTo Target, GetImportableComponents(Path, Target)
    UpdateComponents GetUpdatableComponents(Path, Target), Target

End Sub


' Returns a Collection of paths to modules which can be imported into the VBProject
' based on the specified Path to a folder and specified Target's VBProject.
Private Function GetImportableComponents(ByVal Path As String, ByRef Target As Workbook) As Collection '<String>

    Dim Output As New Collection '<String>
    
    Dim File As Scripting.File
    Dim CmpName As String
    For Each File In Fso.GetFolder(Path).Files
        If IsImportable(File) Then
            CmpName = GetComponentNameFrom(File)
            
            If ComponentExists(CmpName, Target) = False Then
                Output.Add File.Path
            End If
        End If
    Next File

    Set GetImportableComponents = Output
    Set File = Nothing
    Set Output = Nothing

End Function


' Indicates if file should be imported into VBProject or not.
Private Function IsImportable(ByRef File As Scripting.File) As Boolean
    
    If pIgnoreImport.Exists(GetComponentNameFrom(File)) Then
        IsImportable = False
        Exit Function
    End If
    
    Dim Ext As String
    Ext = Fso.GetExtensionName(File.Path)
    If ExtensionToComponentType.Exists(Ext) = False Then
        IsImportable = False
        Exit Function
    End If

    Dim CmpType As vbext_ComponentType
    CmpType = ExtensionToComponentType(Ext)

    If ArrayH.Exists(CmpType, ImportableComponentsTypes) = False Then
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


' Map between extension and VBComponentType enumerator.
Private Function ExtensionToComponentType() As Dictionary '<Key: String, Value: vbext_ComponentType>

    Dim Output As New Dictionary
    Output.Add "cls", vbext_ct_ClassModule
    Output.Add "frm", vbext_ct_MSForm
    Output.Add "bas", vbext_ct_StdModule
    Output.Add "doccls", vbext_ct_Document
    Output.Add "ocx", vbext_ct_ActiveXDesigner
    
    Set ExtensionToComponentType = Output
    Set Output = Nothing

End Function


Private Function ImportableComponentsTypes() As Variant
    ImportableComponentsTypes = Array(vbext_ct_ClassModule, vbext_ct_StdModule)
End Function


' Imports components based on to the specified Workbook using paths to components specified in Cmps collection.
' Param
' Paths: Collection <String>
Private Sub ImportComponentsTo(ByRef Target As Workbook, ByRef Paths As Collection)

    Const MethodName = "ImportComponents"
    
    Dim Path As Variant '<String>
    Dim Cmp As VBComponent

    For Each Path In Paths
        Set Cmp = Target.VBProject.VBComponents.Import(Path)
        Debug.Print ModuleName & "." & MethodName & ": Module imported - " & Cmp.Name
    Next Path
    
    Set Cmp = Nothing

End Sub


' Updates VBComponents based on the collection of paths to modules in the specified Target workbook.
' Params
' Source: Collection <String>
' Target: Workbook
Private Sub UpdateComponents(ByRef Source As Collection, ByRef Target As Workbook)
    
    Const MethodName = "UpdateComponents"
    
    Dim Path As Variant '<String>
    Dim File As Scripting.File
    Dim Cmp As VBComponent
    
    For Each Path In Source
        Set File = Fso.GetFile(Path)
        RemoveComponent GetComponentNameFrom(File), Target
        Set Cmp = Target.VBProject.VBComponents.Import(File.Path)
        Debug.Print ModuleName & "." & MethodName & ": Module Updated - " & Cmp.Name
    Next Path
    
    Set File = Nothing
    Set Cmp = Nothing
    
End Sub


' Returns the name of the VB Component using VB_Name attribute based on the specified File.
Private Function GetComponentNameFrom(ByRef File As Scripting.File) As String
    
    Dim St As Scripting.TextStream
    Set St = File.OpenAsTextStream
    Dim Ln As String
    
    Do Until St.AtEndOfStream = True
        Ln = St.ReadLine
        
        If InStr(Ln, "Attribute VB_Name") Then
            Dim DblQuStart As Long
            Dim DblQuFinish As Long
        
            DblQuStart = InStr(1, Ln, """", vbTextCompare)
            DblQuFinish = InStr(DblQuStart + 1, Ln, """", vbTextCompare)
            GetComponentNameFrom = Mid$(Ln, DblQuStart + 1, DblQuFinish - DblQuStart - 1)
            Exit Do
        End If
    Loop
    
    St.Close
    Set St = Nothing

End Function


Private Sub RemoveComponent(ByVal Name As String, ByRef Target As Workbook)
    
    Dim Cmp As VBComponent
    Set Cmp = GetComponentByName(Name, Target)
        
    ' If name of the component is not changed before it is removed,
    ' new component will be imported with the "1" at the end.
    Cmp.Name = "A" & GetRandom10Digit & "_OLD"
    Target.VBProject.VBComponents.Remove Cmp

    Set Cmp = Nothing

End Sub


' Returns a VBComponent object based on the name of the component and Source workbook.
Private Function GetComponentByName(ByVal CmpName As String, ByRef Source As Workbook) As VBComponent
    
    Dim Cmp As VBComponent
    For Each Cmp In Source.VBProject.VBComponents
        If CmpName = Cmp.Name Then
            Set GetComponentByName = Cmp
            Exit Function
        End If
    Next Cmp

End Function


Private Function GetRandom10Digit() As String
    
    Randomize
    GetRandom10Digit = Int(Rnd * 10000000000#)
    
End Function


' Contains logic when component should be updated based on the specified path to a folder.
' Returns a collection of paths to files which should be use to update modules.
Private Function GetUpdatableComponents(ByVal Path As String, ByRef Target As Workbook) As Collection ' <String>
    
    Dim Output As New Collection
    Dim File As Scripting.File
    Dim CmpName As String
    
    For Each File In Fso.GetFolder(Path).Files
        If IsImportable(File) Then
            CmpName = GetComponentNameFrom(File)
            If ComponentExists(CmpName, Target) Then
                Output.Add File.Path
            End If
        End If
    Next File
    
    Set GetUpdatableComponents = Output
    Set Output = Nothing
    Set File = Nothing

End Function


Private Function ComponentExists(ByVal CmpName As String, ByRef Source As Workbook) As Boolean
    ComponentExists = IIf(GetComponentByName(CmpName, Source) Is Nothing, False, True)
End Function


' Full name means - name of the component with an extension.
Private Function GetFullName(ByRef Component As VBIDE.VBComponent) As String
    GetFullName = Component.Name & "." & ComponentTypeToExtension(Component.Type)
End Function


Public Property Get ExportEmptyComponents() As Boolean
    ExportEmptyComponents = pExportEmptyComponents
End Property


Public Property Let ExportEmptyComponents(ByVal Value As Boolean)
    pExportEmptyComponents = Value
End Property



