Attribute VB_Name = "LambdaTests"
'@Folder("Tests")
Option Explicit

Public Sub Start()

    On Error Resume Next
    Debug.Print "Arithmetic operations", Lambda.Create("(3*(2+5)+5*8/2^(2+1))/26").Run() = 1
    Debug.Print "Logical operations", Lambda.Create("5<3 or 5>3").Run() = True
    Debug.Print "Arguments", Lambda.Create("$1 + $2").Run(5, 9) = 14
    Debug.Print "Property access", Lambda.Create("$1.Range(""A1"")").Run(Sheets(1)).Address(True, True, xlA1, True) Is Sheets(1).Range("A1").Address(True, True, xlA1, True)
    
    Lambda.Create("$1#select").Run (ActiveSheet.Range("A1"))
    Debug.Print "Evaluate methods access", Selection.Address(True, True, xlA1, True) = ActiveSheet.Range("A1").Address(True, True, xlA1, True)
    
    'inline if
    Dim Lam As Variant
    Set Lam = Lambda.Create("if $1 then 0 else if $2 then 1 else 1 + 1")
    Debug.Print "Inline if 1", Lambda.Run(True, True) = 0
    Debug.Print "Inline if 2", Lambda.Run(False, True) = 1
    Debug.Print "Inline if 3", Lambda.Run(False, False) = 2
    
    Debug.Print "Pure functions", Lambda.Create("uCase(trim(""          oranges        "")) & len(""potatoes"")").Run() = "ORANGES8"
    Debug.Print "Multiline using :", Lambda.Create("2+2: 5*2").Run() = 10 'not really a test for whether the 1st line executed

    
    'variables
    With Lambda.CreateMultiline(Array( _
         "test = 2", _
         "if $1 then", _
         "   smth = test + 2", _
         "   test = smth * 2", _
         "else", _
         "   test = test + 4", _
         "end", _
         "test" _
    ))
        Debug.Print "Variables 1", .Run(True) = 8
        Debug.Print "Variables 2", .Run(False) = 6
    End With
    With Lambda.Create("test = 2: if $1 then smth = test + 2: test = smth * 2 else test = test + 4 end: test ")
        Debug.Print "Variables 3", .Run(True) = 8
        Debug.Print "Variables 4", .Run(False) = 6
    End With
    
    'function definition
    Debug.Print "Function 1 fibonacci recursion", Lambda.CreateMultiline(Array( _
         "fun fib(v)", _
         "  if v<=1 then", _
         "    v", _
         "  else ", _
         "    fib(v-2) + fib(v-1)", _
         "  end", _
         "end", _
         "fib($1)" _
    )).Run(20) = 6765

    Debug.Print "Function 2 functions calling functions", Lambda.CreateMultiline(Array( _
         "fun mul3(v) v * 3 end", _
         "fun mul3Add1(v) mul3(v) + 2 end", _
         "mul3Add1(2) + mul3Add1(2)" _
    )).Run() = 16
    
    Debug.Print "Function 3 local vars", Lambda.CreateMultiline(Array( _
         "someVar = 12", _
         "fun localVars(v)", _
         "  smth = 3", _
         "  if v < 2 then ", _
         "    smth = smth + 2", _
         "  end ", _
         "  smth", _
         "end", _
         "someVar + localVars(1)" _
    )).Run() = 17
    
    Debug.Print "Function 4 nested functions", Lambda.CreateMultiline(Array( _
         "fun somth()", _
         "  fun nested()", _
         "    2", _
         "  end", _
         "  nested() + nested()", _
         "end", _
         "somth()" _
    )).Run() = 4
    
    'not allowed
    'Debug.Print "", Lambda.CreateMultiline(Array( _
    '     "fun somth()", _
    '     "  fun nested()", _
    '     "    2", _
    '     "  end", _
    '     "  nested() + nested()", _
    '     "end", _
    '     "nested()" _
    ')).Run()
    
    'Debug.Print "", Lambda.CreateMultiline(Array( _
    '     "someVar = 12", _
    '     "fun globalVars(v)", _
    '     "  smth = 3", _
    '     "  if v < 2 then ", _
    '     "    smth = smth + someVar", _
    '     "  end ", _
    '     "  smth", _
    '     "end", _
    '     "someVar + globalVars(1)" _
    ')).Run()

    'Test Lambda::bind()
    With Lambda.Create("Array($1,$2,$3)").Bind(1)
        Debug.Print "Lambda::Bind() 1 Example", Join(.Run(2, 3), "|") = "1|2|3"
        With .Bind(2)
            Debug.Print "Lambda::Bind() 2 Example", Join(.Run(3), "|") = "1|2|3"
            With .Bind(3)
                Debug.Print "Lambda::Bind() 3 Example", Join(.Run(), "|") = "1|2|3"
            End With
        End With
        
        'In a historical version of Lambda these would fail:
        Debug.Print "Lambda::Bind() 4 Ensure creation of new bindings doesn't erase old bindings", Join(Lambda.Create("Array($1,$2,$3)").Run(1, 2, 3), "|") = "1|2|3"
        Debug.Print "Lambda::Bind() 5 Ensure creation of new bindings doesn't erase old bindings", Join(.Run(2, 3), "|") = "1|2|3"
        
        'Can also bind multiple arguments simultaneously
        With .Bind(2, "hello")
            Debug.Print "Lambda::Bind() 6 multiple arg binding", Join(.Run(), "|") = "1|2|hello"
        End With
    End With

    'Testing global variables
    With Lambda.Create("hello + 2")
        .BindGlobal "hello", 1
        Debug.Print "Lambda::BindGlobal() 1 Can bind global variables", .Run() = 3
    End With

    Dim oDict As Object: Set oDict = CreateObject("Scripting.Dictionary")

    'Testing dictionary declaration
    With Lambda.Create("$1.TEST")
        oDict("TEST") = True
        Debug.Print "Dictionary.Method syntax", .Run(oDict)
    End With

    'Ensure late-bound bindGlobal() works
    Dim ICallable As ICallable, bSuccess As Boolean
    Set ICallable = Lambda.Create("hello")
    ICallable.SendMessage "bindGlobal", bSuccess, Array("hello", True)
    Debug.Print "SendMessage Success Parameter set", bSuccess
    Debug.Print "SendMessage Successful binding", ICallable.Run()
    ICallable.SendMessage "", bSuccess, Null
    Debug.Print "SendMessage Fail Parameter set", Not bSuccess

    'Ensure performance chaching is faster when params are the same:
    oDict("number") = 1
    
    'Time without performance cache
    Dim iStart As Double
    iStart = Timer
    
    Dim i As Long
    With Lambda.Create("$1.number+1")
      For i = 1 To 10 ^ 4
          .Run (oDict)
      Next
    End With
    Dim timeA As Double: timeA = (Timer - iStart)
    
    'Time with performance cache
    iStart = Timer
    With Lambda.Create("$1.number+1", True)
      For i = 1 To 10 ^ 4
          .Run (oDict)
      Next
    End With
    
End Sub
