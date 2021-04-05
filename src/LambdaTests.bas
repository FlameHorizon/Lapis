Attribute VB_Name = "LambdaTests"
'@Folder("Tests")
Option Explicit

Private Const ModuleName As String = "LambdaTests"

Public Sub Start()

    ArithmeticOperationsTest
    LogicalOperationsTest
    ArgumentsTest
    PropertyAccessTest
    EvaluateMethodsAccessTest
    InlineIfTest
    PureFunctionsTest
    MultilineUsingTest
    MultilineWithVariablesTest
    SinglelineWithVariablesTest
    FunctionDefinitionRecursionTest
    FunctionDefinitionWithCallToDifferentFunctionTest
    FunctionDefinitionWithLocalVarsTest
    FunctionDefinitionWithNestedFunctionsTest
    BindingTest
    CreationOfNewBindingDoesNotEraseOldBindingTest
    BindMultipleArgumentsSimultaneously
    GlobalVariablesTest
    DictionaryDeclarationTest
    LateBoundGlobalBindTest
    
End Sub


Private Sub ArithmeticOperationsTest()

    On Error GoTo ErrHandler
    Const MethodName = "ArithmeticOperationsTest"

    ' Assert
    ExUnit.AreEqual 1, Lambda.Create("(3*(2+5)+5*8/2^(2+1))/26").Run(), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Function GetSig(ByVal MethodName As String) As String
    GetSig = ModuleName & "." & MethodName
End Function


Private Sub LogicalOperationsTest()

    On Error GoTo ErrHandler
    Const MethodName = "LogicalOperationsTest"

    ' Assert
    ExUnit.IsTrue Lambda.Create("5<3 or 5>3").Run(), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub ArgumentsTest()

    On Error GoTo ErrHandler
    Const MethodName = "ArgumentsTest"

    ' Assert
    ExUnit.AreEqual 14, Lambda.Create("$1 + $2").Run(5, 9), GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub PropertyAccessTest()

    On Error GoTo ErrHandler
    Const MethodName = "PropertyAccessTest"

    ' Assert
    ExUnit.AreEqual "Lapis", Lambda.Create("$1.VBProject.Name").Run(ThisWorkbook), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub EvaluateMethodsAccessTest()

    On Error GoTo ErrHandler
    Const MethodName = "EvaluateMethodsAccessTest"
    
    ' Act
    Lambda.Create("$1#activate").Run ThisWorkbook

    ' Assert
    ExUnit.AreEqual "Lapis.xlam", ActiveWorkbook.Name, GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub InlineIfTest()

    On Error GoTo ErrHandler
    Const MethodName = "InlineIfTest"

    ' Arrange
    Dim Expr As Variant
    
    ' Act
    Set Expr = Lambda.Create("if $1 then 0 else if $2 then 1 else 1 + 1")
    
    ' Assert
    ExUnit.AreEqual 0, Expr.Run(True, True), GetSig(MethodName)
    ExUnit.AreEqual 1, Expr.Run(False, True), GetSig(MethodName)
    ExUnit.AreEqual 2, Expr.Run(False, False), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub PureFunctionsTest()

    On Error GoTo ErrHandler
    Const MethodName = "PureFunctionsTest"
    
    ' Assert
    ExUnit.AreEqual "ORANGES8", Lambda.Create("uCase(trim(""          oranges        "")) & len(""potatoes"")").Run(), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub MultilineUsingTest()

    On Error GoTo ErrHandler
    Const MethodName = "MultilineUsingTest"
    
    ' Assert
    ExUnit.AreEqual 10, Lambda.Create("2+2: 5*2").Run(), GetSig(MethodName)
    
    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub MultilineWithVariablesTest()

    On Error GoTo ErrHandler
    Const MethodName = "MultilineWithVariablesTest"

    ' Arrange
    Dim Expr As ICallable
    
    ' Act
    Set Expr = Lambda.CreateMultiline(Array( _
         "test = 2", _
         "if $1 then", _
         "   smth = test + 2", _
         "   test = smth * 2", _
         "else", _
         "   test = test + 4", _
         "end", _
         "test" _
    ))

    ' Assert
   
    ExUnit.AreEqual 8, Expr.Run(True), GetSig(MethodName)
    ExUnit.AreEqual 6, Expr.Run(False), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub SinglelineWithVariablesTest()

    On Error GoTo ErrHandler
    Const MethodName = "SinglelineWithVariablesTest"

    ' Arrange
    Dim Expr As ICallable
    
    ' Act
    Set Expr = Lambda.Create("test = 2: if $1 then smth = test + 2: test = smth * 2 else test = test + 4 end: test ")
    
    ' Assert
    ExUnit.AreEqual 8, Expr.Run(True), GetSig(MethodName)
    ExUnit.AreEqual 6, Expr.Run(False), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub FunctionDefinitionRecursionTest()

    On Error GoTo ErrHandler
    Const MethodName = "FunctionDefinitionRecursionTest"

    ' Arrange
    Dim Expr As ICallable
    
    ' Act
    Set Expr = Lambda.CreateMultiline(Array( _
         "fun fib(v)", _
         "  if v<=1 then", _
         "    v", _
         "  else ", _
         "    fib(v-2) + fib(v-1)", _
         "  end", _
         "end", _
         "fib($1)" _
    ))
    
    ' Assert
    ExUnit.AreEqual 6765, Expr.Run(20), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub FunctionDefinitionWithCallToDifferentFunctionTest()

    On Error GoTo ErrHandler
    Const MethodName = "FunctionDefinitionWithCallToDifferentFunctionTest"

    ' Arrange
    Dim Expr As ICallable
    
    ' Act
    Set Expr = Lambda.CreateMultiline(Array( _
         "fun mul3(v) v * 3 end", _
         "fun mul3Add1(v) mul3(v) + 2 end", _
         "mul3Add1(2) + mul3Add1(2)"))

    ' Assert
    ExUnit.AreEqual 16, Expr.Run, GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub FunctionDefinitionWithLocalVarsTest()

    On Error GoTo ErrHandler
    Const MethodName = "FunctionDefinitionWithLocalVarsTest"

    ' Arrange
    Dim Expr As ICallable
    
    ' Act
    Set Expr = Lambda.CreateMultiline(Array( _
         "someVar = 12", _
         "fun localVars(v)", _
         "  smth = 3", _
         "  if v < 2 then ", _
         "    smth = smth + 2", _
         "  end ", _
         "  smth", _
         "end", _
         "someVar + localVars(1)"))

    ' Assert
    ExUnit.AreEqual 17, Expr.Run, GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub FunctionDefinitionWithNestedFunctionsTest()

    On Error GoTo ErrHandler
    Const MethodName = "FunctionDefinitionWithNestedFunctionsTest"

    ' Arrange
    Dim Expr As ICallable
    
    ' Act
    Set Expr = Lambda.CreateMultiline(Array( _
         "fun somth()", _
         "  fun nested()", _
         "    2", _
         "  end", _
         "  nested() + nested()", _
         "end", _
         "somth()"))

    ' Assert
    ExUnit.AreEqual 4, Expr.Run, GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub BindingTest()

    On Error GoTo ErrHandler
    Const MethodName = "BindingTest"
    
    ' Arrange
    Dim Expr As ICallable
    
    ' Act
    Set Expr = Lambda.Create("Array($1,$2,$3)")

    ' Assert
    ExUnit.AreEqual "1|2|3", VBA.Join(Expr.Bind(1).Run(2, 3), "|"), GetSig(MethodName)
    ExUnit.AreEqual "1|2|3", VBA.Join(Expr.Bind(1, 2).Run(3), "|"), GetSig(MethodName)
    ExUnit.AreEqual "1|2|3", VBA.Join(Expr.Bind(1, 2, 3).Run(), "|"), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub CreationOfNewBindingDoesNotEraseOldBindingTest()

    On Error GoTo ErrHandler
    Const MethodName = "CreationOfNewBindingDoesNotEraseOldBindingTest"

    ' Assert
    ExUnit.AreEqual "1|2|3", VBA.Join(Lambda.Create("Array($1,$2,$3)").Run(1, 2, 3), "|"), GetSig(MethodName)
    ExUnit.AreEqual "1|2|3", VBA.Join(Lambda.Create("Array($1,$2,$3)").Bind(1).Run(2, 3), "|"), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub BindMultipleArgumentsSimultaneously()

    On Error GoTo ErrHandler
    Const MethodName = "BindMultipleArgumentsSimultaneously"
    
    ' Arrange
    Dim Expr As ICallable
    
    ' Act
    Set Expr = Lambda.Create("Array($1,$2,$3)").Bind(1, 2, "hello")
    
    ' Assert
    ExUnit.AreEqual "1|2|hello", VBA.Join(Expr.Run, "|"), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub GlobalVariablesTest()

    On Error GoTo ErrHandler
    Const MethodName = "GlobalVariablesTest"
    
    ' Arrange
    Dim Expr As ICallable
    
    ' Act
    Set Expr = Lambda.Create("hello + 2").BindGlobal("hello", 1)
    
    ' Assert
    ExUnit.AreEqual 3, Expr.Run, GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub DictionaryDeclarationTest()

    On Error GoTo ErrHandler
    Const MethodName = "DictionaryDeclarationTest"

    ' Arrange
    Dim Dict As New Scripting.Dictionary
    Dim Expr As ICallable
    
    ' Act
    Set Expr = Lambda.Create("$1.TEST")
    Dict.Add "TEST", True

    ' Assert
    ExUnit.IsTrue Expr.Run(Dict), GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub


Private Sub LateBoundGlobalBindTest()

    On Error GoTo ErrHandler
    Const MethodName = "LateBoundGlobalBindTest"

    ' Arrange
    Dim Expr As ICallable
    Dim Success As Boolean
    Dim NotSuccess As Boolean
    
    ' Act
    Set Expr = Lambda.Create("hello")
    Expr.SendMessage "bindGlobal", Success, Array("hello", True)
    Expr.SendMessage vbNullString, NotSuccess, Null

    ' Assert
    ExUnit.IsTrue Success, GetSig(MethodName)
    ExUnit.IsTrue Expr.Run, GetSig(MethodName)
    ExUnit.IsFalse NotSuccess, GetSig(MethodName)

    Exit Sub
ErrHandler:
    ExUnit.TestFailRunTime GetSig(MethodName)

End Sub
