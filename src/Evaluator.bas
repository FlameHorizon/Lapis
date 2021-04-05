Attribute VB_Name = "Evaluator"
'@Folder("LambdaExpr")
Option Explicit

Private Const ModuleName As String = "Evaluator"

'Direct call convention of VBA.CallByName
#If Not Mac Then
    #If VBA7 Then
        'VBE7 is interchangable with msvbvm60.dll    however VBE7.dll appears to always be present where as msvbvm60 is only occasionally present.
        Private Declare PtrSafe Function rtcCallByName Lib "VBE7.dll" (ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Variant
    #Else
        Private Declare Function rtcCallByName Lib "msvbvm60" (ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Variant
    #End If
#End If

Public UsePerformanceCache As Boolean
Public PerformanceCache As Object
Public FunctExt As Object 'Dictionary<string => stdCallback>

Private Const pMinStackSize As Long = 30 'note that the stack size may become smaller than this

'Evaluates the given list of operations
'@param {Operation()} operations  The operations to evaluate
'@returns {Variant} The result of the operations
Public Function Evaluate(ByRef Ops() As Operation, ByVal vLastArgs As Variant) As Variant

    Const MethodName = "Evaluate"

    Dim Stack() As Variant
    ReDim Stack(0 To 5)
    Dim StackPtr As Long: StackPtr = 0
    
    Dim Op As Operation
    Dim v1 As Variant
    Dim v2 As Variant
    Dim v3 As Variant
    Dim OpIndex As Long: OpIndex = 0
    Dim OpCount As Long: OpCount = UBound(Ops)
    
    'If result is in performance cache then return it immediately
    If UsePerformanceCache Then
        Dim PerformanceCacheID As String: PerformanceCacheID = GetPerformanceCacheID(vLastArgs)
        If PerformanceCache.Exists(PerformanceCacheID) Then
            System.CopyVariant Evaluate, PerformanceCache(PerformanceCacheID)
            Exit Function
        End If
    End If

    'Evaluate operations to identify result
    While OpIndex <= OpCount
        Op = Ops(OpIndex)
        OpIndex = OpIndex + 1
        Select Case Op.Type
            Case EvalOperationType.oPush
                PushV Stack, StackPtr, Op.Value
                'Arithmetic
            Case EvalOperationType.oArithmetic
                v2 = PopV(Stack, StackPtr)
                Select Case Op.SubType
                    Case ExpressionType.oAdd
                        v1 = PopV(Stack, StackPtr)
                        v3 = v1 + v2
                    Case ExpressionType.oSub
                        v1 = PopV(Stack, StackPtr)
                        v3 = v1 - v2
                    Case ExpressionType.oMul
                        v1 = PopV(Stack, StackPtr)
                        v3 = v1 * v2
                    Case ExpressionType.oDiv
                        v1 = PopV(Stack, StackPtr)
                        v3 = v1 / v2
                    Case ExpressionType.oPow
                        v1 = PopV(Stack, StackPtr)
                        v3 = v1 ^ v2
                    Case ExpressionType.oMod
                        v1 = PopV(Stack, StackPtr)
                        v3 = v1 Mod v2
                    Case ExpressionType.oNeg
                        v3 = -v2
                    Case Else
                        v3 = Empty
                End Select
                PushV Stack, StackPtr, v3
                'Comparison
            Case EvalOperationType.oComparison
                v2 = PopV(Stack, StackPtr)
                v1 = PopV(Stack, StackPtr)
                Select Case Op.SubType
                    Case ExpressionType.oEql
                        v3 = v1 = v2
                    Case ExpressionType.oNeq
                        v3 = v1 <> v2
                    Case ExpressionType.oGt
                        v3 = v1 > v2
                    Case ExpressionType.oGte
                        v3 = v1 >= v2
                    Case ExpressionType.oLt
                        v3 = v1 < v2
                    Case ExpressionType.oLte
                        v3 = v1 <= v2
                    Case Else
                        v3 = Empty
                End Select
                PushV Stack, StackPtr, v3
                'Logic
            Case EvalOperationType.oLogic
                v2 = PopV(Stack, StackPtr)
                Select Case Op.SubType
                    Case ExpressionType.oAnd
                        v1 = PopV(Stack, StackPtr)
                        v3 = v1 And v2
                    Case ExpressionType.oOr
                        v1 = PopV(Stack, StackPtr)
                        v3 = v1 Or v2
                    Case ExpressionType.oNot
                        v3 = Not v2
                    Case ExpressionType.oXor
                        v1 = PopV(Stack, StackPtr)
                        v3 = v1 Xor v2
                    Case Else
                        v3 = Empty
                End Select
                PushV Stack, StackPtr, v3
                'Object
            Case EvalOperationType.oObject
                ObjectCaller Stack, StackPtr, Op
                'Func
            Case EvalOperationType.oFunc
                Dim Args() As Variant
                Args = GetArgs(Stack, StackPtr)
                v1 = PopV(Stack, StackPtr)
                PushV Stack, StackPtr, EvaluateFunc(v1, Args)
                'Misc
            Case EvalOperationType.oMisc
                v2 = PopV(Stack, StackPtr)
                v1 = PopV(Stack, StackPtr)
                Select Case Op.SubType
                    Case ExpressionType.oCat
                        v3 = v1 & v2
                    Case ExpressionType.oLike
                        v3 = v1 Like v2
                    Case Else
                        v3 = Empty
                End Select
                PushV Stack, StackPtr, v3
                'Variable
            Case EvalOperationType.oAccess
                Select Case Op.SubType
                    Case ExpressionType.argument
                        Dim ArgIndex As Long: ArgIndex = val(Mid$(Op.Value, 2)) + LBound(vLastArgs) - 1
                        If ArgIndex <= UBound(vLastArgs) Then
                            PushV Stack, StackPtr, vLastArgs(ArgIndex)
                        Else
                            Lapis.Errors.OnArgumentError "ArgIndex", "Argument not supplied to Lambda. " & ModuleName & "." & MethodName
                        End If
                    Case Else
                        PushV Stack, StackPtr, Stack(StackPtr - Op.Value)
                End Select
            Case EvalOperationType.oSet
                v1 = PopV(Stack, StackPtr)
                Stack(StackPtr - Op.Value) = v1
                'Flow
            Case EvalOperationType.oJump
                Select Case Op.SubType
                    Case ExpressionType.ifTrue
                        v1 = PopV(Stack, StackPtr)
                        If v1 Then
                            OpIndex = Op.Value
                        End If
                    Case ExpressionType.ifFalse
                        v1 = PopV(Stack, StackPtr)
                        If Not v1 Then
                            OpIndex = Op.Value
                        End If
                    Case Else
                        OpIndex = Op.Value
                End Select
            Case EvalOperationType.oReturn
                Select Case Op.SubType
                    Case ExpressionType.withValue
                        v1 = PopV(Stack, StackPtr)
                        OpIndex = Stack(StackPtr - 1)
                        Stack(StackPtr - 1) = v1
                    Case Else
                        OpIndex = PopV(Stack, StackPtr)
                End Select
                'Data
            Case EvalOperationType.oMerge
                System.CopyVariant v1, PopV(Stack, StackPtr)
                System.CopyVariant Stack(StackPtr - 1), v1
            Case EvalOperationType.oPop
                PopV Stack, StackPtr
            Case Else
                OpIndex = 10 ^ 6 'TODO: replace by infinity or something
        End Select
    Wend

    'Add result to performance cache
    If UsePerformanceCache Then
        If IsObject(Stack(0)) Then
            Set PerformanceCache(PerformanceCacheID) = Stack(0)
        Else
            PerformanceCache(PerformanceCacheID) = Stack(0)
        End If
    End If

    System.CopyVariant Evaluate, Stack(0)
    
End Function



'Serializes the argument array passed to a string.
'@param {ByRef Variant()} Arguments to serialize
'@returns {String} Serialized representation of the arguments.
'@remark Objects cannot be split into their components and thus are cached as a conglomerate of type and pointer (e.g. Dictionary<12341234123>).
'@TODO: Potentially use [StgSerializePropVariant ](https://docs.microsoft.com/en-us/windows/win32/api/propvarutil/nf-propvarutil-stgserializepropvariant) as this'd be more optimal
'@example
'  Debug.Print getPerformanceCacheID(Array())=""
'  Debug.Print getPerformanceCacheID(Array(Array(1, 2, Null), "yop", Empty, "", Nothing, New Collection, DateSerial(2020, 1, 1), False, True)) = "Array[1;2;null;];""yop"";empty;"""";Nothing;Collection<1720260481920>;01/01/2020;False;True;"
'returns
'  True
'  True
Private Function GetPerformanceCacheID(ByRef Arguments As Variant) As String

    Dim Length As Long: Length = UBound(Arguments) - LBound(Arguments) + 1
    If Length > 0 Then
        Dim Serialized As String: Serialized = vbNullString
        Dim i As Long
        For i = LBound(Arguments) To UBound(Arguments)
            Select Case VarType(Arguments(i))
                Case vbBoolean, vbByte, vbInteger, vbLong, vbLongLong, vbCurrency, vbDate, vbDecimal, vbDouble, vbSingle
                    Serialized = Serialized & Arguments(i) & ";"
                Case vbString
                    Serialized = Serialized & """" & Arguments(i) & """;"
                Case vbObject, vbDataObject
                    If Arguments(i) Is Nothing Then
                        Serialized = Serialized & "Nothing;"
                    Else
                        Serialized = Serialized & TypeName(Arguments(i)) & "<" & ObjPtr(Arguments(i)) & ">;"
                    End If
                Case vbEmpty
                    Serialized = Serialized & "empty;"
                Case vbNull
                    Serialized = Serialized & "null;"
                Case vbError
                    Serialized = Serialized & "error;"
                Case Else
                    If CBool(VarType(Arguments(i)) And vbArray) Then
                        Serialized = Serialized & "Array[" & GetPerformanceCacheID(Arguments(i)) & "];"
                    Else
                        Serialized = Serialized & "Unknown;"
                    End If
            End Select
        Next
    End If
    GetPerformanceCacheID = Serialized
    
End Function


Private Sub PushV(ByRef Stack() As Variant, ByRef Index As Long, ByVal Item As Variant)

    Dim Size As Long: Size = UBound(Stack)
    If Index > Size Then
        ReDim Preserve Stack(0 To Size * 2)
    End If
    If IsObject(Item) Then
        Set Stack(Index) = Item
    Else
        Stack(Index) = Item
    End If
    Index = Index + 1
    
End Sub


Private Function PopV(ByRef Stack() As Variant, ByRef Index As Variant) As Variant

    Dim Size As Long: Size = UBound(Stack)
    If Index < Size / 3 And Index > pMinStackSize Then
        ReDim Preserve Stack(0 To CLng(Size / 2))
    End If
    
    Index = Index - 1
    If IsObject(Stack(Index)) Then
        Set PopV = Stack(Index)
    Else
        PopV = Stack(Index)
    End If
    
    #If devMode Then
        Stack(Index) = Empty
    #End If
    
End Function


'Calls an object method/setter/getter/letter
'@param {ByRef Variant()} stack     The stack to get the data from and add the result to
'@param {ByRef Long} stackPtr       The pointer that indicates the position of the top of the stack
'@param {ByRef Operation} op        The operation to execute
Private Sub ObjectCaller(ByRef Stack() As Variant, ByRef StackPtr As Long, ByRef Op As Operation)

    'Get the name and arguments
    Dim Args() As Variant: Args = GetArgs(Stack, StackPtr)
    Dim FuncName As Variant: FuncName = PopV(Stack, StackPtr)
    
    'Get caller type
    Dim CallerType As VbCallType
    Select Case Op.SubType
        Case ExpressionType.oPropGet:     CallerType = VbGet
        Case ExpressionType.oMethodCall:  CallerType = VbMethod
        Case ExpressionType.oPropLet:     CallerType = VbLet
        Case ExpressionType.oPropSet:     CallerType = VbSet
    End Select
                
    'Call rtcCallByName
    Dim Obj As Object
    Set Obj = PopV(Stack, StackPtr)
    PushV Stack, StackPtr, StdCallByName(Obj, FuncName, CallerType, Args)
    
End Sub


'Retrieves the arguments from the stack
'@param {ByRef Variant()} stack     The stack to get the data from and add the result to
'@param {ByRef Long} stackPtr       The pointer that indicates the position of the top of the stack
'@returns {Variant()} The args list
Private Function GetArgs(ByRef Stack() As Variant, ByRef StackPtr As Long) As Variant

    Dim ArgCount As Variant: ArgCount = Stack(StackPtr - 1)
    Dim Args() As Variant
    If VarType(ArgCount) = vbString Then
        'If no argument count is specified, there are no arguments
        ArgCount = 0
        Args = Array()
    Else
        'If an argument count is provided, extract all arguments into an array
        PopV Stack, StackPtr
        ReDim Args(1 To ArgCount)
        
        Dim i As Long
        'Arguments are held on the stack in order, which means that we need to fill the array in reverse order.
        For i = ArgCount To 1 Step -1
            System.CopyVariant Args(i), PopV(Stack, StackPtr)
        Next
    End If
    
    GetArgs = Args
    
End Function


'Evaluates the built in standard functions
'@param {String} FuncName The name of the function to invoke
'@param {Variant} args() The arguments
'@returns The result
Private Function EvaluateFunc(ByVal FuncName As String, ByVal Args As Variant) As Variant

    Const MethodName = "EvaluateFunc"

    Dim ArgStart As Long: ArgStart = LBound(Args)
    If TypeName(FunctExt) = "Dictionary" Then
        If FunctExt.Exists(FuncName) Then
            Dim InjectedVar As Variant
            System.CopyVariant InjectedVar, FunctExt(FuncName)
            If TypeOf InjectedVar Is ICallable Then
                System.CopyVariant EvaluateFunc, FunctExt(FuncName).RunEx(Args)
            Else
                System.CopyVariant EvaluateFunc, InjectedVar
            End If
            Exit Function
        End If
    End If
    
    Select Case LCase$(FuncName)
        Case "eval": EvaluateFunc = Lambda.Create(Args(ArgStart)).Run()

            'Useful OOP constants
        Case "thisworkbook": Set EvaluateFunc = ThisWorkbook
        Case "application":  Set EvaluateFunc = Application

            'MATH:
            '-----
        Case "abs": EvaluateFunc = VBA.Math.Abs(Args(ArgStart))
        Case "int": EvaluateFunc = VBA.Int(Args(ArgStart))
        Case "fix": EvaluateFunc = VBA.Fix(Args(ArgStart))
        Case "exp": EvaluateFunc = VBA.Math.Exp(Args(ArgStart))
        Case "log": EvaluateFunc = VBA.Math.Log(Args(ArgStart))
        Case "sqr": EvaluateFunc = VBA.Math.Sqr(Args(ArgStart))
        Case "sgn": EvaluateFunc = VBA.Math.Sgn(Args(ArgStart))
        Case "rnd": EvaluateFunc = VBA.Math.Rnd(Args(ArgStart))

            'Trigonometry
        Case "cos":  EvaluateFunc = VBA.Math.Cos(Args(ArgStart))
        Case "sin":  EvaluateFunc = VBA.Math.Sin(Args(ArgStart))
        Case "tan":  EvaluateFunc = VBA.Math.Tan(Args(ArgStart))
        Case "atn":  EvaluateFunc = VBA.Math.Atn(Args(ArgStart))
        Case "asin": EvaluateFunc = VBA.Math.Atn(Args(ArgStart) / VBA.Math.Sqr(-1 * Args(ArgStart) * Args(ArgStart) + 1))
        Case "acos": EvaluateFunc = VBA.Math.Atn(-1 * Args(ArgStart) / VBA.Math.Sqr(-1 * Args(ArgStart) * Args(ArgStart) + 1)) + 2 * Atn(1)

            'VBA Constants:
        Case "vbcrlf":          EvaluateFunc = vbCrLf
        Case "vbcr":            EvaluateFunc = vbCr
        Case "vblf":            EvaluateFunc = vbLf
        Case "vbnewline":       EvaluateFunc = vbNewLine
        Case "vbnullchar":      EvaluateFunc = vbNullChar
        Case "vbnullstring":    EvaluateFunc = vbNullString
        Case "vbobjecterror":   EvaluateFunc = vbObjectError
        Case "vbtab":           EvaluateFunc = vbTab
        Case "vbback":          EvaluateFunc = vbBack
        Case "vbformfeed":      EvaluateFunc = vbFormFeed
        Case "vbverticaltab":   EvaluateFunc = vbVerticalTab


            'VBA Structure
        Case "array": EvaluateFunc = Args
            'TODO: Case "callbyname": evaluateFunc = CallByName(args(ArgStart))
        Case "createobject"
            Select Case UBound(Args)
                Case ArgStart
                    Set EvaluateFunc = CreateObject(Args(ArgStart))
                Case ArgStart + 1
                    Set EvaluateFunc = CreateObject(Args(ArgStart), Args(ArgStart + 1))
            End Select
        Case "getobject"
            Select Case UBound(Args)
                Case ArgStart
                    Set EvaluateFunc = GetObject(Args(ArgStart))
                Case ArgStart + 1
                    Set EvaluateFunc = GetObject(Args(ArgStart), Args(ArgStart + 1))
            End Select
        Case "iff"
            If CBool(Args(ArgStart)) Then
                EvaluateFunc = Args(ArgStart + 1)
            Else
                EvaluateFunc = Args(ArgStart + 2)
            End If
        Case "typename"
            EvaluateFunc = TypeName(Args(ArgStart))
            
            'VBA Casting
        Case "cbool": EvaluateFunc = VBA.Conversion.CBool(Args(ArgStart))
        Case "cbyte": EvaluateFunc = VBA.Conversion.CByte(Args(ArgStart))
        Case "ccur":  EvaluateFunc = VBA.Conversion.CCur(Args(ArgStart))
        Case "cdate": EvaluateFunc = VBA.Conversion.CDate(Args(ArgStart))
        Case "csng":  EvaluateFunc = VBA.Conversion.CSng(Args(ArgStart))
        Case "cdbl":  EvaluateFunc = VBA.Conversion.CDbl(Args(ArgStart))
        Case "cint":  EvaluateFunc = VBA.Conversion.CInt(Args(ArgStart))
        Case "clng":  EvaluateFunc = VBA.Conversion.CLng(Args(ArgStart))
        Case "cstr":  EvaluateFunc = VBA.Conversion.CStr(Args(ArgStart))
        Case "cvar":  EvaluateFunc = VBA.Conversion.CVar(Args(ArgStart))
        Case "cverr": EvaluateFunc = VBA.Conversion.CVErr(Args(ArgStart))
        
            'Conversion
        Case "asc":     EvaluateFunc = VBA.Asc(Args(ArgStart))
        Case "chr":     EvaluateFunc = VBA.Chr$(Args(ArgStart))
        
        Case "format"
            Select Case UBound(Args)
                Case ArgStart
                    EvaluateFunc = Format$(Args(ArgStart))
                Case ArgStart + 1
                    EvaluateFunc = Format$(Args(ArgStart), Args(ArgStart + 1))
                Case ArgStart + 2
                    EvaluateFunc = Format$(Args(ArgStart), Args(ArgStart + 1), Args(ArgStart + 2))
                Case ArgStart + 3
                    EvaluateFunc = Format$(Args(ArgStart), Args(ArgStart + 1), Args(ArgStart + 2), Args(ArgStart + 3))
            End Select
        Case "hex":     EvaluateFunc = VBA.Conversion.Hex$(Args(ArgStart))
        Case "oct":     EvaluateFunc = VBA.Conversion.Oct$(Args(ArgStart))
        Case "str":     EvaluateFunc = VBA.Conversion.Str$(Args(ArgStart))
        Case "val":     EvaluateFunc = VBA.Conversion.val(Args(ArgStart))
        
            'String functions
        Case "trim":  EvaluateFunc = VBA.Trim$(Args(ArgStart))
        Case "lcase": EvaluateFunc = VBA.LCase$(Args(ArgStart))
        Case "ucase": EvaluateFunc = VBA.UCase$(Args(ArgStart))
        Case "right": EvaluateFunc = VBA.right(Args(ArgStart), Args(ArgStart + 1))
        Case "left":  EvaluateFunc = VBA.Left$(Args(ArgStart), Args(ArgStart + 1))
        Case "len":   EvaluateFunc = VBA.Len(Args(ArgStart))

        Case "mid"
            Select Case UBound(Args)
                Case ArgStart + 1
                    EvaluateFunc = VBA.Mid$(Args(ArgStart), Args(ArgStart + 1))
                Case ArgStart + 2
                    EvaluateFunc = VBA.Mid$(Args(ArgStart), Args(ArgStart + 1), Args(ArgStart + 2))
            End Select
            'Misc
        Case "now": EvaluateFunc = VBA.Now()
        Case "switch"
            'TODO: Switch caching and use of dictionary would be good here
            Dim i As Long
            For i = ArgStart + 1 To UBound(Args) Step 2
                If i + 1 > UBound(Args) Then
                    System.CopyVariant EvaluateFunc, Args(i)
                    Exit For
                Else
                    If IsObject(Args(ArgStart)) And IsObject(Args(i)) Then
                        If Args(ArgStart) Is Args(i) Then
                            Set EvaluateFunc = Args(i + 1)
                            Exit For
                        End If
                    ElseIf (Not IsObject(Args(ArgStart))) And (Not IsObject(Args(i))) Then
                        If Args(ArgStart) = Args(i) Then
                            EvaluateFunc = Args(i + 1)
                            Exit For
                        End If
                    End If
                End If
            Next
        Case "any"
            EvaluateFunc = False
            'Detect if comparee is an object or a value
            If IsObject(Args(ArgStart)) Then
                For i = ArgStart + 1 To UBound(Args)
                    If IsObject(Args(i)) Then
                        If Args(ArgStart) Is Args(i) Then
                            EvaluateFunc = True
                            Exit For
                        End If
                    End If
                Next
            Else
                For i = ArgStart + 1 To UBound(Args)
                    If Not IsObject(Args(i)) Then
                        If Args(ArgStart) = Args(i) Then
                            EvaluateFunc = True
                            Exit For
                        End If
                    End If
                Next
            End If
        Case Else
            Lapis.Errors.OnInvalidOperation vbNullString, "No such function: " & FuncName & ". " _
                                                          & ModuleName & "." & MethodName
    End Select
    
End Function



'Calls an object method/setter/getter/letter. Treats dictionary properties as direct object properties, I.E. `A.B` ==> `A.item("B")`
'@param {ByRef Object} - The object to call
'@param {ByVal String} - The method name to call
'@param {ByVal VbCallType} - The property/method call type
'@param {ByVal Variant()}  - An array of arguments. This function supports up to 30 arguments, akin to Application.Run
'@returns Variant - The return value of the called function
Public Function StdCallByName(ByRef Obj As Object, _
                              ByVal FuncName As String, _
                              ByVal CallerType As VbCallType, _
                              ByRef Args() As Variant) As Variant
                              
    'If Dictionary and
    If TypeName(Obj) = "Dictionary" Then
        Select Case FuncName
            Case "add", "exists", "items", "keys", "remove", "removeall", "comparemode", "count", "item", "key"
                'These methods already exist on dictionary, do not override
            Case Else
                'Call DictionaryInstance.Item(funcName) only if funcName exists on the item
                If Obj.Exists(FuncName) Then
                    'TODO: Make this work for callerType.VbLet
                    System.CopyVariant StdCallByName, Obj.Item(FuncName)
                    Exit Function
                End If
        End Select
    End If

    'Call CallByName from DLL or
    #If Mac Then
        Call System.CopyVariant(StdCallByName, macCallByName(Obj, FuncName, CallerType, Args))
    #Else
        System.CopyVariant StdCallByName, rtcCallByName(Obj, StrPtr(FuncName), CallerType, Args, &H409)
    #End If
    
End Function


