Attribute VB_Name = "TreeNodeTests"
Option Explicit
'@Folder("Tests")

Private Const ModuleName As String = "TreeTests"


Public Sub Start()

    AddChildNodeTest
    AddChildNodeAsDataTest
    AddRangeOfChildNodesTest
    SetParentNodeTest

End Sub


Private Sub AddChildNodeTest()

    On Error GoTo ErrHandler
    Const MethodName = "AddChildNodeTest"
    
    Dim Trunk As New TreeNode
    Trunk.Init1 "1"
    
    Dim NewNode As New TreeNode
    NewNode.Init1 "2"
    Trunk.AddChild NewNode
    
    ExUnit.AreSame NewNode, Trunk.Children(1), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Function GetSig(ByVal MethodName As String) As String
    GetSig = ModuleName & "." & MethodName
End Function


Private Sub AddChildNodeAsDataTest()

    On Error GoTo ErrHandler
    Const MethodName = "AddChildNodeAsDataTest"
    
    Dim Trunk As New TreeNode
    Trunk.Init1 "1"
    Trunk.AddChildData "2"
    
    ExUnit.AreEqual "2", Trunk.Children(1).Data, GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub AddRangeOfChildNodesTest()

    On Error GoTo ErrHandler
    Const MethodName = "AddRangeOfChildNodesTest"
    
    Dim Trunk As New TreeNode
    Trunk.Init1 "1"
    
    Dim Node1 As New TreeNode: Node1.Init1 "2"
    Dim Node2 As New TreeNode: Node2.Init1 "3"
    
    Dim Nodes As New Collection '<TreeNode>
    Nodes.Add Node1
    Nodes.Add Node2
    
    Trunk.AddChildren Nodes
    
    ExUnit.AreEqual 2, Trunk.Children.Count, GetSig(MethodName)
    ExUnit.AreEqual "2", Trunk.Children(1).Data, GetSig(MethodName)
    ExUnit.AreEqual "3", Trunk.Children(2).Data, GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SetParentNodeTest()

    On Error GoTo ErrHandler
    Const MethodName = "SetParentNodeTest"
    
    Dim Trunk As New TreeNode: Trunk.Init1 "1"
    Dim Node1 As New TreeNode: Node1.Init1 "2"
    Dim Node2 As New TreeNode: Node2.Init1 "3"
    
    Node1.Parent = Trunk
    Node2.Parent = Node1
    
    ExUnit.AreSame Trunk, Node1.Parent, GetSig(MethodName)
    ExUnit.AreSame Node1, Node2.Parent, GetSig(MethodName)
    ExUnit.AreSame Nothing, Trunk.Parent, GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


