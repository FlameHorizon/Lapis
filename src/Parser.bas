Attribute VB_Name = "Parser"
'@Folder("LambdaExpr")
Option Explicit

Public StackSize As Long
Public ScopeCount As Long
Public TokenIndex As Long
Public Tokens() As Token
Public OperationIndex As Long
Public Operations() As Operation
Public FuncScope As Long

Private pScopes() As Variant
Private pScopesArgCount() As Variant

Public Type Operation
    Type As EvalOperationType
    SubType As ExpressionType
    Value As Variant
End Type


'Evaluation operation types
Public Enum EvalOperationType
    oPush = 1
    oPop = 2
    oMerge = 3
    oAccess = 4
    oSet = 5
    oArithmetic = 6
    oLogic = 7
    oFunc = 8
    oComparison = 9
    oMisc = 10
    oJump = 11
    oReturn = 12
    oObject = 13
End Enum


Public Enum ExpressionType
    'Arithmetic
    oAdd = 1
    oSub = 2
    oMul = 3
    oDiv = 4
    oPow = 5
    oNeg = 6
    oMod = 7
    'Logic
    oAnd = 8
    oOr = 9
    oNot = 10
    oXor = 11
    'comparison
    oEql = 12
    oNeq = 13
    oLt = 14
    oLte = 15
    oGt = 16
    oGte = 17
    oIs = 18
    'misc operators
    oCat = 19
    oLike = 20
    'misc
    ifTrue = 21
    ifFalse = 22
    withValue = 23
    argument = 24
    'object
    oPropGet = 25
    oPropLet = 26
    oPropSet = 27
    oMethodCall = 28
    oEquality = 29    'Yet to be implemented
    oIsOperator = 30  'Yet to be implemented
    oEnum = 31        'Yet to be implemented
End Enum


' NOTE: This method gets called recursively.
Public Sub ParseBlock(ParamArray EndToken() As Variant)
    
    AddScope
    Dim Size As Long: Size = StackSize + 1
    
    ' Consume multiple lines
    Dim bLoop As Boolean: bLoop = True
    Do
        While OptConsume("colon"): Wend
        ParseStatement
        
        Dim i As Long
        For i = LBound(EndToken) To UBound(EndToken)
            If Peek(EndToken(i)) Then
                bLoop = False
            End If
        Next
    Loop While bLoop
    
    ' Get rid of all extra expression results and declarations
    While StackSize > Size
        AddOperation oMerge, , , -1
    Wend
    ScopeCount = ScopeCount - 1
    
End Sub


Private Sub AddScope()

    ScopeCount = ScopeCount + 1
    Dim Scope As Long: Scope = ScopeCount
    ReDim Preserve pScopes(1 To Scope)
    ReDim Preserve pScopesArgCount(1 To Scope)
    Set pScopes(Scope) = CreateObject("Scripting.Dictionary")
    Set pScopesArgCount(Scope) = CreateObject("Scripting.Dictionary")
    
End Sub


Private Sub ParseStatement()

    If Peek("var") And Peek("equal", 2) Then
        ParseAssignment
    ElseIf Peek("fun") Then
        ParseFunctionDeclaration
    Else
        ParseExpression
    End If
    
End Sub


' Combines peek and consume, consuming a token only if matched, without throwing an error if not
' @param {string} token The token that is expected
' @returns {vbNullString|string} Whether the expected token was found
Private Function OptConsume(ByVal TokenType As String) As Boolean

    Dim Matched As Boolean: Matched = Peek(TokenType)
    If Matched Then
        Consume TokenType
    End If
    OptConsume = Matched
    
End Function


'Checks whether the token at TokenIndex is of the given type
'@param {string} token  The token that is expected
'@param {long} offset   The number of tokens to look into the future, defaults to 1
'@returns {boolean} Whether the expected token was found
Private Function Peek(ByVal TokenType As String, Optional ByVal Offset As Long = 1) As Boolean

    If TokenIndex = 0 Then TokenIndex = 1
    If TokenIndex + Offset - 1 <= UBound(Tokens) Then
        Peek = Tokens(TokenIndex + Offset - 1).Type.Name = TokenType
    Else
        Peek = False
    End If
    
End Function


'Adds an operation to the instance operations list
'@param {IType} kType         The main type of the operation
'@param {ISubType} subType    The sub type of the operation
'@param {Variant} value       The value associated with the operation
'@param {Integer} stackDelta  The effect this has on the stack size (increasing or decreasing it)
'@returns {Integer} The index of the created operation
Private Function AddOperation(ByVal kType As EvalOperationType, _
                              Optional ByVal SubType As ExpressionType, _
                              Optional ByVal Value As Variant, _
                              Optional ByVal StackDelta As Long) As Long
    
    If OperationIndex = 0 Then
        ReDim Preserve Operations(0 To 1)
    Else
        Dim Size As Long: Size = UBound(Operations)
        If OperationIndex > Size Then
            ReDim Preserve Operations(0 To Size * 2)
        End If
    End If
    
    With Operations(OperationIndex)
        .Type = kType
        .SubType = SubType
        System.CopyVariant .Value, Value
    End With
    AddOperation = OperationIndex
    StackSize = StackSize + StackDelta
    
    OperationIndex = OperationIndex + 1
    
End Function


Private Sub ParseAssignment()

    Dim VarName As String: VarName = Consume("var")
    Consume "equal"
    ParseExpression
    Dim Offset As Long: Offset = FindVariable(VarName)
    If Offset >= 0 Then
        ' If the variable already existed, move the data to that pos on the stack
        AddOperation oSet, , Offset, -1
        AddOperation oAccess, , Offset, 1 ' To keep a return value
    Else
        ' If the variable didn't exist yet, treat this stack pos as its source
        pScopes(ScopeCount).Add VarName, StackSize
    End If
    
End Sub


Private Sub ParseFunctionDeclaration()
    
    ' Create a dedicated scope for this funcion
    AddScope
    Dim PrevFuncScope As Long: PrevFuncScope = FuncScope
    FuncScope = ScopeCount
    
    ' Add operation to skip this code in normal operation flow
    Dim SkipToIndex As Long: SkipToIndex = AddOperation(oJump)
    
    ' Obtain the signature
    Consume "fun"
    Dim FuncName As String: FuncName = Consume("var")
    Consume "lBracket"
    Dim ArgCount As Long
    While Not Peek("rBracket")
        If ArgCount > 0 Then
            Consume "comma"
        End If
        
        ParseParameterDeclaration
        ArgCount = ArgCount + 1
    Wend
    Consume "rBracket"
    
    ' Register the function
    pScopes(ScopeCount - 1).Add FuncName, OperationIndex
    pScopesArgCount(ScopeCount - 1).Add FuncName, ArgCount
    
    ' Obtain the body
    ParseBlock "end"
    Consume "end"
    While ArgCount > 0
        AddOperation oMerge, , , -1
        ArgCount = ArgCount - 1
    Wend
    AddOperation oReturn, withValue, , -1
    Operations(SkipToIndex).Value = OperationIndex
    
    ' Reset the scope
    ScopeCount = ScopeCount - 1
    FuncScope = PrevFuncScope
    
End Sub


Private Sub ParseExpression()
    ParseLogicPriority1
End Sub


Private Sub ParseLogicPriority1() 'xor

    ParseLogicPriority2
    Dim bLoop As Boolean: bLoop = True
    Do
        If OptConsume("xor") Then
            ParseLogicPriority2
            AddOperation oLogic, oXor, , -1
        Else
            bLoop = False
        End If
    Loop While bLoop
    
End Sub


Private Sub ParseLogicPriority2() 'or

    ParseLogicPriority3
    Dim bLoop As Boolean: bLoop = True
    Do
        If OptConsume("or") Then
            ParseLogicPriority3
            AddOperation oLogic, oOr, , -1
        Else
            bLoop = False
        End If
    Loop While bLoop
    
End Sub


Private Sub ParseLogicPriority3() 'and

    ParseLogicPriority4
    Dim bLoop As Boolean: bLoop = True
    Do
        If OptConsume("and") Then
            ParseLogicPriority4
            AddOperation oLogic, oAnd, , -1
        Else
            bLoop = False
        End If
    Loop While bLoop
    
End Sub


Private Sub ParseLogicPriority4() 'not

    Dim invert As Variant: invert = vbNull
    While OptConsume("not")
        If invert = vbNull Then invert = False
        invert = Not invert
    Wend
    
    ParseComparisonPriority1
    
    If invert <> vbNull Then
        AddOperation oLogic, oNot
        If invert = False Then
            AddOperation oLogic, oNot
        End If
    End If
    
End Sub


Private Sub ParseComparisonPriority1() '=, <>, <, <=, >, >=, is, Like

    ParseArithmeticPriority1
    Dim bLoop As Boolean: bLoop = True
    Do
        If OptConsume("lessThan") Then
            ParseArithmeticPriority1
            AddOperation oComparison, oLt, , -1
        ElseIf OptConsume("lessThanEqual") Then
            ParseArithmeticPriority1
            AddOperation oComparison, oLte, , -1
        ElseIf OptConsume("greaterThan") Then
            ParseArithmeticPriority1
            AddOperation oComparison, oGt, , -1
        ElseIf OptConsume("greaterThanEqual") Then
            ParseArithmeticPriority1
            AddOperation oComparison, oGte, , -1
        ElseIf OptConsume("equal") Then
            ParseArithmeticPriority1
            AddOperation oComparison, oEql, , -1
        ElseIf OptConsume("notEqual") Then
            ParseArithmeticPriority1
            AddOperation oComparison, oNeq, , -1
        ElseIf OptConsume("is") Then
            ParseArithmeticPriority1
            AddOperation oComparison, oIs, , -1
        ElseIf OptConsume("like") Then
            ParseArithmeticPriority1
            AddOperation oComparison, oLike, , -1
        Else
            bLoop = False
        End If
    Loop While bLoop
    
End Sub


Private Sub ParseArithmeticPriority1() '&

    ParseArithmeticPriority2
    Dim bLoop As Boolean: bLoop = True
    Do
        If OptConsume("concatenate") Then
            ParseArithmeticPriority2
            AddOperation oMisc, oCat, , -1
        Else
            bLoop = False
        End If
    Loop While bLoop
    
End Sub


Private Sub ParseArithmeticPriority2() '+, -

    ParseArithmeticPriority3
    Dim bLoop As Boolean: bLoop = True
    Do
        If OptConsume("add") Then
            ParseArithmeticPriority3
            AddOperation oArithmetic, oAdd, , -1
        ElseIf OptConsume("subtract") Then
            ParseArithmeticPriority3
            AddOperation oArithmetic, oSub, , -1
        Else
            bLoop = False
        End If
    Loop While bLoop
    
End Sub


Private Sub ParseArithmeticPriority3() 'mod

    ParseArithmeticPriority4
    Dim bLoop As Boolean: bLoop = True
    Do
        If OptConsume("mod") Then
            ParseArithmeticPriority4
            AddOperation oArithmetic, oMod, , -1
        Else
            bLoop = False
        End If
    Loop While bLoop
    
End Sub


Private Sub ParseArithmeticPriority4() '*, /

    ParseArithmeticPriority5
    Dim bLoop As Boolean: bLoop = True
    Do
        If OptConsume("multiply") Then
            ParseArithmeticPriority4
            AddOperation oArithmetic, oMul, , -1
        ElseIf OptConsume("divide") Then
            ParseArithmeticPriority4
            AddOperation oArithmetic, oDiv, , -1
        Else
            bLoop = False
        End If
    Loop While bLoop
    
End Sub


Private Sub ParseArithmeticPriority5() '+, - (unary)

    If OptConsume("subtract") Then
        ParseArithmeticPriority5 'recurse
        AddOperation oArithmetic, oNeg
    ElseIf OptConsume("add") Then
        ParseArithmeticPriority5 'recurse
    Else
        ParseArithmeticPriority6
    End If
    
End Sub


Private Sub ParseArithmeticPriority6() '^
    
    Dim bLoop As Boolean: bLoop = True
    ParseFlowPriority1
    Do
        If OptConsume("power") Then
            ParseArithmeticPriority6andahalf '- and + are still identity operators
            AddOperation oArithmetic, oPow, , -1
        Else
            bLoop = False
        End If
    Loop While bLoop
    
End Sub


Private Sub ParseArithmeticPriority6andahalf() '+, - (unary)

    If OptConsume("subtract") Then
        ParseArithmeticPriority6andahalf 'recurse
        AddOperation oArithmetic, oNeg
    ElseIf OptConsume("add") Then
        ParseArithmeticPriority6andahalf 'recurse
    Else
        ParseFlowPriority1
    End If
    
End Sub


Private Sub ParseFlowPriority1() 'if then else

    If OptConsume("if") Then
        ParseExpression
        Dim skipThenJumpIndex As Long: skipThenJumpIndex = AddOperation(oJump, ifFalse, , -1)
        
        Dim Size As Long: Size = StackSize
        Consume "then"
        ParseBlock "else", "end"
        Dim SkipElseJumpIndex As Long: SkipElseJumpIndex = AddOperation(oJump)
        Operations(skipThenJumpIndex).Value = OperationIndex
        StackSize = Size
        
        If OptConsume("end") Then
            AddOperation oPush, , 0, 1 'Expressions should always return a value
            Operations(SkipElseJumpIndex).Value = OperationIndex
        Else
            Consume "else"
            ParseBlock "eof", "rBracket", "end"
            Operations(SkipElseJumpIndex).Value = OperationIndex
        
            OptConsume "end"
        End If
    Else
        ParseValuePriority1
    End If
    
End Sub


Private Sub ParseValuePriority1() 'numbers, $vars, strings, booleans, (expressions)

    If Peek("literalNumber") Then
        AddOperation oPush, , CDbl(Consume("literalNumber")), 1
    ElseIf Peek("arg") Then
        AddOperation oAccess, argument, Consume("arg"), 1
        ParseManyAccessors
    ElseIf Peek("literalString") Then
        ParseString
    ElseIf Peek("literalBoolean") Then
        AddOperation oPush, , Consume("literalBoolean") = "true", 1
    ElseIf Peek("var") Then
        If Not ParseScopeAccess Then
            ParseFunction
        End If
        ParseManyAccessors
    Else
        Consume "lBracket"
        ParseExpression
        Consume "rBracket"
        ParseManyAccessors
    End If
    
End Sub


' Consumes a token
' @param {string} token The token type name to consume
' @throws If the expected token wasn't found
' @returns {string} The value of the token
Private Function Consume(ByVal TknName As String) As String

    Dim FirstToken As Token
    FirstToken = ShiftTokens()
    If FirstToken.Type.Name <> TknName Then
        Lapis.Errors.OnInvalidOperation vbNullString, _
                                        "Unexpected token, found: " & FirstToken.Type.Name _
                                        & " but expected: " & TknName
    Else
        Consume = FirstToken.Value
    End If
    
End Function


Private Function FindVariable(ByRef VarName As String) As Long

    Dim Scope As Long: Scope = ScopeCount
    FindVariable = -1
    While Scope > 0
        If pScopes(Scope).Exists(VarName) Then
            If Scope < FuncScope Then
                Lapis.Errors.OnInvalidOperation "VarName", _
                                                "Can't access """ & VarName _
                                                & """, functions can unfortunately not access data outside their block"
            ElseIf pScopesArgCount(Scope).Exists(VarName) Then
                Lapis.Errors.OnInvalidOperation "VarName", _
                                                "Expected a variable, but found a function for name " & VarName
            Else
                FindVariable = StackSize - pScopes(Scope).Item(VarName)
                Scope = 0
            End If
        End If
        Scope = Scope - 1
    Wend
    
End Function


Private Sub ParseParameterDeclaration()

    Dim VarName As String: VarName = Consume("var")
    Dim Offset As Long: Offset = FindVariable(VarName)
    If Offset >= 0 Then
        Lapis.Errors.OnInvalidOperation vbNullString, "You can't declare multiple parameters with the same name"
    Else
        ' Reserve a spot for this parameter, it will be pushed by the caller
        StackSize = StackSize + 1
        pScopes(ScopeCount).Add VarName, StackSize
    End If
    
End Sub


Private Sub ParseManyAccessors()

    Dim bLoop As Boolean: bLoop = True
    Do
        bLoop = False
        If ParseOptObjectProperty() Then bLoop = True
        If ParseOptObjectMethod() Then bLoop = True
    Loop While bLoop
    
End Sub


Private Sub ParseString()

    Dim Res As String: Res = Consume("literalString")
    Res = Mid$(Res, 2, Len(Res) - 2)
    Res = VBA.Replace(Res, """""", """")
    AddOperation oPush, , Res, 1
    
End Sub


Private Function ParseScopeAccess() As Boolean

    If Peek("lBracket", 2) Then
        ParseScopeAccess = ParseFunctionAccess()
    Else
        ParseScopeAccess = ParseVariableAccess()
    End If
    
End Function


Private Sub ParseFunction()

    AddOperation oPush, , Consume("var"), 1
    Dim Size As Long: Size = StackSize
    ParseOptParameters
    AddOperation oFunc
    StackSize = Size
    
End Sub


'Shifts the Tokens array (uses an index)
'@returns {token} The token at the tokenIndex
Private Function ShiftTokens() As Token

    If TokenIndex = 0 Then
        TokenIndex = 1
    End If
    
    'Get next token
    ShiftTokens = Tokens(TokenIndex)
    
    'Increment token index
    TokenIndex = TokenIndex + 1
    
End Function


Private Function ParseOptObjectProperty() As Boolean

    ParseOptObjectProperty = False
    If OptConsume("propertyAccess") Then
        Dim Size As Long: Size = StackSize
        AddOperation oPush, , Consume("var"), 1
        ParseOptParameters
        AddOperation oObject, oPropGet
        StackSize = Size
        ParseOptObjectProperty = True
    End If
    
End Function


Private Function ParseOptObjectMethod() As Boolean

    ParseOptObjectMethod = False
    If OptConsume("methodAccess") Then
        Dim Size As Long: Size = StackSize
        AddOperation oPush, , Consume("var"), 1
        ParseOptParameters
        AddOperation oObject, oMethodCall
        StackSize = Size
        ParseOptObjectMethod = True
    End If
    
End Function


Private Function ParseFunctionAccess() As Boolean

    ParseFunctionAccess = False
    Dim FuncName As String: FuncName = Consume("var")
    Dim ArgCount As Long
    Dim FuncPos As Long: FuncPos = FindFunction(FuncName, ArgCount)
    If FuncPos <> -1 Then
        ParseFunctionAccess = True
        Dim ReturnPosIndex As Long: ReturnPosIndex = AddOperation(oPush, , , 1)
        
        ' Consume the arguments
        Consume ("lBracket")
        Dim iArgCount As Long
        While Not Peek("rBracket")
            If iArgCount > 0 Then
                Consume "comma"
            End If
            ParseExpression
            iArgCount = iArgCount + 1
        Wend
        Consume "rBracket"
        If iArgCount <> ArgCount Then
            Lapis.Errors.OnInvalidOperation vbNullString, _
                                            ArgCount & " arguments should have been provided to " & FuncName _
                                            & " but only " & iArgCount & " were received"
        End If
        
        ' Add call and return data
        AddOperation oJump, , FuncPos, -iArgCount 'only -argCount since pushing Result and popping return pos cancel out
        Operations(ReturnPosIndex).Value = OperationIndex
    Else
        TokenIndex = TokenIndex - 1 ' Revert token consumption
    End If
    
End Function


Private Function ParseVariableAccess() As Boolean

    ParseVariableAccess = False
    Dim VarName As String: VarName = Consume("var")
    Dim Offset As Long: Offset = FindVariable(VarName)
    If Offset >= 0 Then
        ParseVariableAccess = True
        AddOperation oAccess, , 1 + Offset, 1
    Else
        TokenIndex = TokenIndex - 1 ' Revert token consumption
    End If
    
End Function


Private Function ParseOptParameters() As Boolean

    ParseOptParameters = False
    If OptConsume("lBracket") Then
        Dim iArgCount As Long
        While Not Peek("rBracket")
            If iArgCount > 0 Then
                Consume "comma"
            End If
            ParseExpression
            iArgCount = iArgCount + 1
        Wend
        Consume "rBracket"
        If iArgCount > 0 Then
            AddOperation oPush, , iArgCount, 1
        End If
        ParseOptParameters = True
    End If
    
End Function


Private Function FindFunction(ByRef VarName As String, Optional ByRef ArgCount As Long) As Long

    Dim Scope As Long: Scope = ScopeCount
    FindFunction = -1
    While Scope > 0
        If pScopes(Scope).Exists(VarName) Then
            If Not pScopesArgCount(Scope).Exists(VarName) Then
                Lapis.Errors.OnInvalidOperation "VarName", _
                                                "Expected a function, but found a variable for name " & VarName
            Else
                FindFunction = pScopes(Scope).Item(VarName)
                ArgCount = pScopesArgCount(Scope).Item(VarName)
                Scope = 0
            End If
        End If
        Scope = Scope - 1
    Wend
    
End Function


