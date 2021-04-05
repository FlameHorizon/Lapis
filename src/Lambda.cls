VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("LambdaExpr")
Option Explicit

'For testing purposes
'#Const Mac = True

Implements ICallable

#Const devMode = True 'Used for enabling some debugging feature

Public Enum LambdaType
    StandardLambda = 1
    BoundLambda = 2
End Enum

Private pFuncExt As Object 'Dictionary<string => stdCallback>
Private pOperations() As Operation
Private pIsBound As Boolean
Private pBound As Lambda
Private pVarBound As Variant
Private pUsePerformanceCache As Boolean
Private pPerformanceCache As Object


Public Function Create(ByVal Equation As String, _
                       Optional ByVal UsePerformanceCache As Boolean = False, _
                       Optional ByVal SandboxExtras As Boolean = False) As Lambda
    
    'Cache Lambda created
    Dim pCache As Object
    If pCache Is Nothing Then
        Set pCache = CreateObject("Scripting.Dictionary")
    End If
    
    Dim ID As String: ID = UsePerformanceCache & "-" & SandboxExtras & ")" & Equation
    If Not pCache.Exists(ID) Then
        Set pCache(ID) = New Lambda
        pCache(ID).Init LambdaType.StandardLambda, Equation, UsePerformanceCache, SandboxExtras
    End If
    
    'Return cached lambda
    Set Create = pCache(ID)
    
End Function


Public Function CreateMultiline(ByRef Equation As Variant, _
                                Optional ByVal UsePerformanceCache As Boolean = False, _
                                Optional ByVal SandboxExtras As Boolean = False) As Lambda
    Set CreateMultiline = Create(Join(Equation, " "), UsePerformanceCache, SandboxExtras)
End Function


Public Function BindEx(ByVal Params As Variant) As Lambda

    Set BindEx = New Lambda
    Dim Callable As ICallable: Set Callable = Me
    BindEx.Init LambdaType.BoundLambda, Callable, Params
    
End Function


'Bind a global variable to the function
'@param {String} - New global name
'@param {Variant}- Data to store in global variable
'@returns {stdLambda} The lambda existing lambda
Public Function BindGlobal(ByVal GlobalName As String, ByVal Variable As Variant) As Lambda

    Set BindGlobal = Me
    If pIsBound Then
        pBound.BindGlobal GlobalName, Variable
    Else
        If IsObject(Variable) Then
            Set pFuncExt(GlobalName) = Variable
        Else
            pFuncExt(GlobalName) = Variable
        End If
    End If
    
End Function


Public Sub Init(ByVal Kind As LambdaType, ParamArray Params() As Variant)

    Select Case Kind
        Case LambdaType.StandardLambda
            Dim Equation As String: Equation = Params(0)
            pUsePerformanceCache = Params(1)
            Dim SandboxExtras As Boolean: SandboxExtras = Params(2)
            
            'Performance cache
            If pUsePerformanceCache Then
                Set pPerformanceCache = CreateObject("Scripting.Dictionary")
            End If

            'Function extensions
            If SandboxExtras Or pFuncExt Is Nothing Then
                Set pFuncExt = CreateObject("Scripting.Dictionary")
            End If

            pIsBound = False
            Dim pTokens() As Token
            pTokens = Tokenizer.Tokenise(Equation)
            
            ' Initial state of parser.
            Parser.TokenIndex = 0
            Parser.OperationIndex = 0
            Parser.StackSize = 1
            Parser.ScopeCount = 0
            Parser.FuncScope = 0
            Parser.Tokens = pTokens
            
            Parser.ParseBlock "eof"
            pOperations = Parser.Operations

            'Resizes the pOperations list so there are no more empty pOperations
            ReDim Preserve pOperations(0 To Parser.OperationIndex)

        Case LambdaType.BoundLambda
            pIsBound = True
            Set pBound = Params(0)
            pVarBound = Params(1)
            
            'Function extensions

            If SandboxExtras Or pFuncExt Is Nothing Then
                Set pFuncExt = CreateObject("Scripting.Dictionary")
            End If
        Case Else
            'error?
    End Select
    
End Sub


Private Function ICallable_Run(ParamArray Params() As Variant) As Variant

    If Not pIsBound Then
        'Execute top-down parser
        System.CopyVariant ICallable_Run, Evaluate(pOperations, Params)
    Else
        System.CopyVariant ICallable_Run, pBound.RunEx(ConcatArrays(pVarBound, Params))
    End If
    
End Function


Private Function Evaluate(ByRef Ops() As Operation, ByVal vLastArgs As Variant) As Variant

    Evaluator.UsePerformanceCache = pUsePerformanceCache
    Set Evaluator.PerformanceCache = pPerformanceCache
    Set Evaluator.FunctExt = pFuncExt
    Evaluate = Evaluator.Evaluate(Ops, vLastArgs)

End Function


Private Function ICallable_RunEx(ByVal Params As Variant) As Variant

    If Not IsArray(Params) Then
        Lapis.Errors.OnArgumentError "Params", "Params to be supplied as array of arguments"
    End If
    
    If Not pIsBound Then
        'Execute top-down parser
        System.CopyVariant ICallable_RunEx, Evaluate(pOperations, Params)
    Else
        System.CopyVariant ICallable_RunEx, pBound.RunEx(ConcatArrays(pVarBound, Params))
    End If
    
End Function


Public Function Run(ParamArray Params() As Variant) As Variant

    If Not pIsBound Then
        'Execute top-down parser
        System.CopyVariant Run, Evaluate(pOperations, Params)
    Else
        System.CopyVariant Run, pBound.RunEx(ConcatArrays(pVarBound, Params))
    End If
    
End Function


Public Function RunEx(ByVal Params As Variant) As Variant

    If Not pIsBound Then
        If Not IsArray(Params) Then
            Lapis.Errors.OnArgumentError "Params", "Params to be supplied as array of arguments"
        End If
        
        'Execute top-down parser
        System.CopyVariant RunEx, Evaluate(pOperations, Params)
    Else
        System.CopyVariant RunEx, pBound.RunEx(ConcatArrays(pVarBound, Params))
    End If
    
End Function


'Bind a parameter to the function
Private Function ICallable_Bind(ParamArray Params() As Variant) As ICallable
    Set ICallable_Bind = BindEx(Params)
End Function


Public Function Bind(ParamArray Params() As Variant) As Lambda
    Set Bind = BindEx(Params)
End Function


'Low-dependency function calling
'@protected
'@param {ByVal String}       - Message to send
'@param {ByRef Boolean}      - Success of message. If message wasn't processed return false.
'@param {Paramarray Variant} - Parameters to pass along with message
'@returns {Variant} - Anything returned by the function
Private Function ICallable_SendMessage(ByVal Message As String, _
                                          ByRef Success As Boolean, _
                                          ByVal Params As Variant) As Variant

    Select Case Message
        Case "obj"
            Set ICallable_SendMessage = Me
            Success = True
        Case "className"
            ICallable_SendMessage = "stdLambda"
            Success = True
        Case "bindGlobal"
            'Bind global based whether this is a bound lambda or not
            BindGlobal Params(0), Params(1)
            Success = True
        Case Else
            Success = False
    End Select
    
End Function


'Used by Bind() for binding arguments ontop of BoundArgs and binding bound args to passed arguments
'@param {Variant()} The 1st array which will
'@param {Variant()} The 2nd array which will be concatenated after the 1st
Private Function ConcatArrays(ByVal Arr1 As Variant, ByVal Arr2 As Variant) As Variant

    Dim ub1 As Long: ub1 = UBound(Arr1)
    Dim lb1 As Long: lb1 = LBound(Arr1)
    Dim ub2 As Long: ub2 = UBound(Arr2)
    Dim lb2 As Long: lb2 = LBound(Arr2)
    Dim iub As Long: iub = ub1 + ub2 - lb2 + 1
    
    If iub > -1 Then
        Dim v() As Variant
        ReDim v(lb1 To iub)
        
        Dim i As Long
        For i = LBound(v) To UBound(v)
            If i <= ub1 Then
                System.CopyVariant v(i), Arr1(i)
            Else
                System.CopyVariant v(i), Arr2(i - ub1 - 1 + lb2)
            End If
        Next
        ConcatArrays = v
    Else
        ConcatArrays = Array()
    End If
    
End Function


'Reimplementation of rtcCallByName() but for Mac OS
'@param {ByRef Object} - The object to call
'@param {ByVal String} - The method name to call
'@param {ByVal VbCallType} - The property/method call type
'@param {ByVal Variant()}  - An array of arguments. This function supports up to 30 arguments, akin to Application.Run
'@returns Variant - The return value of the called function
'@Ignore ProcedureNotUsed
Private Function macCallByName(ByVal Obj As Object, _
                               ByVal FuncName As String, _
                               ByVal CallerType As VbCallType, _
                               ByVal Args As Variant) As Variant

    'Get currentLength
    Dim CurrentLength As Long: CurrentLength = UBound(Args) - LBound(Args) + 1
    Dim i As Long: i = LBound(Args)
    
    'Cant use same trick as in stdCallback, as it seems CallByName doesn't support the Missing value... So have to do it this way...
    'Will go up to 30 as per Application.Run() Also seems that you can't pass args array directly to CallByName() because it causes an Overflow error,
    'instead we need to convert the args to vars first... Yes this doesn't look at all pretty, but at least it's compartmentalised to the end of the code...
    '@Ignore MultipleDeclarations, VariableTypeNotDeclared
    Dim a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26, a27, a28, a29
    If CurrentLength - 1 >= 0 Then System.CopyVariant a0, Args(i + 0) Else GoTo macJmpCall
    If CurrentLength - 1 >= 1 Then System.CopyVariant a1, Args(i + 1) Else GoTo macJmpCall
    If CurrentLength - 1 >= 2 Then System.CopyVariant a2, Args(i + 2) Else GoTo macJmpCall
    If CurrentLength - 1 >= 3 Then System.CopyVariant a3, Args(i + 3) Else GoTo macJmpCall
    If CurrentLength - 1 >= 4 Then System.CopyVariant a4, Args(i + 4) Else GoTo macJmpCall
    If CurrentLength - 1 >= 5 Then System.CopyVariant a5, Args(i + 5) Else GoTo macJmpCall
    If CurrentLength - 1 >= 6 Then System.CopyVariant a6, Args(i + 6) Else GoTo macJmpCall
    If CurrentLength - 1 >= 7 Then System.CopyVariant a7, Args(i + 7) Else GoTo macJmpCall
    If CurrentLength - 1 >= 8 Then System.CopyVariant a8, Args(i + 8) Else GoTo macJmpCall
    If CurrentLength - 1 >= 9 Then System.CopyVariant a9, Args(i + 9) Else GoTo macJmpCall
    If CurrentLength - 1 >= 10 Then System.CopyVariant a10, Args(i + 10) Else GoTo macJmpCall
    If CurrentLength - 1 >= 11 Then System.CopyVariant a11, Args(i + 11) Else GoTo macJmpCall
    If CurrentLength - 1 >= 12 Then System.CopyVariant a12, Args(i + 12) Else GoTo macJmpCall
    If CurrentLength - 1 >= 13 Then System.CopyVariant a13, Args(i + 13) Else GoTo macJmpCall
    If CurrentLength - 1 >= 14 Then System.CopyVariant a14, Args(i + 14) Else GoTo macJmpCall
    If CurrentLength - 1 >= 15 Then System.CopyVariant a15, Args(i + 15) Else GoTo macJmpCall
    If CurrentLength - 1 >= 16 Then System.CopyVariant a16, Args(i + 16) Else GoTo macJmpCall
    If CurrentLength - 1 >= 17 Then System.CopyVariant a17, Args(i + 17) Else GoTo macJmpCall
    If CurrentLength - 1 >= 18 Then System.CopyVariant a18, Args(i + 18) Else GoTo macJmpCall
    If CurrentLength - 1 >= 19 Then System.CopyVariant a19, Args(i + 19) Else GoTo macJmpCall
    If CurrentLength - 1 >= 20 Then System.CopyVariant a20, Args(i + 20) Else GoTo macJmpCall
    If CurrentLength - 1 >= 21 Then System.CopyVariant a21, Args(i + 21) Else GoTo macJmpCall
    If CurrentLength - 1 >= 22 Then System.CopyVariant a22, Args(i + 22) Else GoTo macJmpCall
    If CurrentLength - 1 >= 23 Then System.CopyVariant a23, Args(i + 23) Else GoTo macJmpCall
    If CurrentLength - 1 >= 24 Then System.CopyVariant a24, Args(i + 24) Else GoTo macJmpCall
    If CurrentLength - 1 >= 25 Then System.CopyVariant a25, Args(i + 25) Else GoTo macJmpCall
    If CurrentLength - 1 >= 26 Then System.CopyVariant a26, Args(i + 26) Else GoTo macJmpCall
    If CurrentLength - 1 >= 27 Then System.CopyVariant a27, Args(i + 27) Else GoTo macJmpCall
    If CurrentLength - 1 >= 28 Then System.CopyVariant a28, Args(i + 28) Else GoTo macJmpCall
    If CurrentLength - 1 >= 29 Then System.CopyVariant a29, Args(i + 29) Else GoTo macJmpCall

macJmpCall:
    Select Case CurrentLength
        Case 0:  System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType)
        Case 1:  System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0)
        Case 2:  System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1)
        Case 3:  System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2)
        Case 4:  System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3)
        Case 5:  System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4)
        Case 6:  System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5)
        Case 7:  System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6)
        Case 8:  System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7)
        Case 9:  System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8)
        Case 10: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9)
        Case 11: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10)
        Case 12: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11)
        Case 13: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12)
        Case 14: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13)
        Case 15: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14)
        Case 16: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15)
        Case 17: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16)
        Case 18: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17)
        Case 19: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18)
        Case 20: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19)
        Case 21: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20)
        Case 22: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21)
        Case 23: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22)
        Case 24: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23)
        Case 25: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24)
        Case 26: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25)
        Case 27: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26)
        Case 28: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26, a27)
        Case 29: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26, a27, a28)
        Case 30: System.CopyVariant macCallByName, CallByName(Obj, FuncName, CallerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26, a27, a28, a29)
    End Select
    
End Function


