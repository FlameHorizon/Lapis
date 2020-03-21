VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Exposed
Option Explicit
'@Folder("Collection")

Private Const ClassName As String = "TreeNode"

Private pInit As Boolean
Private pData As Variant
Private pChildren As New Collection '<TreeNode>
Private pParent As TreeNode


Public Sub Init1(ByVal Data As Variant)

    Const MethodName = "Init1"
    
    If pInit = True Then
        Errors.OnInvalidOperation vbNullString, _
                                  "Object is already initialzed. " _
                                  & ClassName & "." & MethodName
    End If
    
    If IsObject(Data) Then
        Set pData = Data
    Else
        pData = Data
    End If
    pInit = True
    
End Sub


Public Property Get Children() As Collection '<TreeNode>
    Set Children = pChildren
End Property


Public Property Get Parent() As TreeNode
    Set Parent = pParent
End Property


Public Property Let Parent(ByRef Value As TreeNode)
    Set pParent = Value
End Property


Public Property Get Data() As Variant
    
    If IsObject(pData) Then
        Set Data = pData
    Else
        Data = pData
    End If
    
End Property


Public Property Let Data(ByRef Value As Variant)
    
    If IsObject(Value) Then
        Set pData = Value
    Else
        pData = Value
    End If
    
End Property


' Info: Adds new child to the current node.
Public Sub AddChild(ByRef Child As TreeNode)

    CheckInit
    Child.Parent = Me
    Me.Children.Add Child

End Sub


Private Sub CheckInit()
    
    If pInit = False Then
        Errors.OnInvalidOperation vbNullString, _
                                  "Object was not initialized as intended. Use one of " _
                                  & "available Init methods. " & ClassName
    End If
    
End Sub


Public Sub AddChildData(ByVal Data As Variant)
    
    CheckInit
    Dim NewNode As New TreeNode
    NewNode.Init1 Data
    Me.AddChild NewNode
    
End Sub


' Info: Adds a collection of TreeNode objects as children to current node.
Public Sub AddChildren(ByRef Children As Collection) '<TreeNode>

    CheckInit
    Dim Node As TreeNode
    For Each Node In Children
        Node.Parent = Me
    Next Node
    
    CollectionExt.AddRange Me.Children, Children

End Sub



