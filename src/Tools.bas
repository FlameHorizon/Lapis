Attribute VB_Name = "Tools"
'@Folder("Lapis")
Option Explicit

Private Const ModuleName As String = "Tools"
'@Ignore EncapsulatePublicField
Public Fso As New FileSystemObject


Public Function SelectFolder() As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Select a folder"
        .Show
        If .SelectedItems.Count > 0 Then
            SelectFolder = .SelectedItems.Item(1)
        Else
            MsgBox "Folder is not selected."
        End If
    End With
    
End Function


' Returns number of text lines based on the specified Stream.
' Throws an ArgumentNullException when Stream is set to nothing.
Public Function LinesCount(ByVal Stream As ADODB.Stream) As Long

    Const MethodName = "LinesCount"
    
    If Stream Is Nothing Then
        Errors.OnArgumentNull "Stream", ModuleName & "." & MethodName
    End If
    
    If IsStreamClosed(Stream) Then
        Errors.OnInvalidOperation "Stream", "Stream is closed. " & ModuleName & "." & MethodName
    End If
    
    Dim Ln As Long
    Stream.Position = 0
    Do While Stream.EOS <> True
        Stream.SkipLine
        Ln = Ln + 1
    Loop
    
    LinesCount = Ln
    
End Function


' Determines if TextStream is closed.
' There is no property of TextStream object (like object.Closed) to know directly if Stream is close
' or open. To know TextStream state, method attempt to read cursor position. If it fails
' (throws an error), that would mean Stream is not readable (closed).
' Throws an ArgumentNullException when Stream is set to nothing.
Public Function IsStreamClosed(ByRef Stream As ADODB.Stream) As Boolean
    
    Const MethodName = "IsStreamClosed"
    
    If Stream Is Nothing Then
        Errors.OnArgumentNull "Stream", ModuleName & "." & MethodName
    End If
    
    On Error Resume Next
    '@Ignore VariableNotUsed
    Dim Pos As Long: Pos = Stream.Position
    
    If Err.Number = 91 Then
        IsStreamClosed = True
        On Error GoTo 0
        
    ElseIf Err.Number = 0 Then
        IsStreamClosed = False
        
    Else
        ' Other, unexpected error occured. This error has to be moved upward.
        Errors.OnUnhandledError Err.Number & ", " & Err.Description & " " & ModuleName & "." & MethodName
    End If
    
End Function


