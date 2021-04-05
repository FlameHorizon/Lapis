Attribute VB_Name = "Tokenizer"
'@Folder("LambdaExpr")
Option Explicit

Private Const ModuleName As String = "Tokenizer"

Public Type TokenDefinition
    Name As String
    Regex As String
    RegexObj As Object
End Type


Public Type Token
    Type As TokenDefinition
    Value As Variant
    BracketDepth As Long
End Type


'Tokenise the input string
'@param {string} Text String to tokenise
'@return {token[]} A list of Token structs
Public Function Tokenise(ByVal Text As String) As Token()
    
    Const MethodName = "Tokenise"
    
    Dim Defs() As TokenDefinition
    Defs = GetTokenDefinitions()
    
    Dim Tokens() As Token
    Dim TokenDef As Long
    ReDim Tokens(1 To 1)
    
    Dim TextInput As String
    TextInput = Text
    
    Dim NumTokens As Long
    NumTokens = 0
    While Len(TextInput) > 0
        Dim Matched As Boolean
        Matched = False
        
        For TokenDef = 1 To UBound(Defs)
            'Test match, if matched then add token
            If Defs(TokenDef).RegexObj.Test(TextInput) Then
                'Get match details
                Dim Match As Object: Set Match = Defs(TokenDef).RegexObj.Execute(TextInput)
                
                'Create new token
                NumTokens = NumTokens + 1
                ReDim Preserve Tokens(1 To NumTokens)
                
                'Tokenise
                Tokens(NumTokens).Type = Defs(TokenDef)
                Tokens(NumTokens).Value = Match(0)
                
                'Trim string to unmatched range
                TextInput = Mid$(TextInput, Len(Match(0)) + 1)
                
                'Flag that a match was made
                Matched = True
                Exit For
            End If
        Next
        
        'If no match made then syntax error
        If Not Matched Then
            Lapis.Errors.OnArgumentError "TextInput", _
                                         "Syntax Error unexpected character """ & Mid$(TextInput, 1, 1) & """" & ". " _
                                         & ModuleName & "." & MethodName
        End If
    Wend
    
    'Add eof token
    ReDim Preserve Tokens(1 To NumTokens + 1)
    Tokens(NumTokens + 1).Type.Name = "eof"
    
    Tokenise = RemoveTokens(Tokens, "space")
    
End Function


Private Function GetTokenDefinitions() As TokenDefinition()
    
    Dim Arr() As TokenDefinition
    ReDim Arr(1 To 99)
    
    Dim i As Long: i = 0
    'Whitespace
    i = i + 1: Arr(i) = GetTokenDefinition("space", "\s+") 'String

    'Literal
    i = i + 1: Arr(i) = GetTokenDefinition("literalString", """(?:""""|[^""])*""") 'String
    i = i + 1: Arr(i) = GetTokenDefinition("literalNumber", "\d+(?:\.\d+)?") 'Number
    i = i + 1: Arr(i) = GetTokenDefinition("literalBoolean", "True|False", IsKeyword:=True)
    
    'Named operators
    i = i + 1: Arr(i) = GetTokenDefinition("is", "is", IsKeyword:=True)
    i = i + 1: Arr(i) = GetTokenDefinition("mod", "mod", IsKeyword:=True)
    i = i + 1: Arr(i) = GetTokenDefinition("and", "and", IsKeyword:=True)
    i = i + 1: Arr(i) = GetTokenDefinition("or", "or", IsKeyword:=True)
    i = i + 1: Arr(i) = GetTokenDefinition("xor", "xor", IsKeyword:=True)
    i = i + 1: Arr(i) = GetTokenDefinition("not", "not", IsKeyword:=True)

    'Structural
    ' Inline if
    i = i + 1: Arr(i) = GetTokenDefinition("if", "if", IsKeyword:=True)
    i = i + 1: Arr(i) = GetTokenDefinition("then", "then", IsKeyword:=True)
    i = i + 1: Arr(i) = GetTokenDefinition("else", "else", IsKeyword:=True)
    i = i + 1: Arr(i) = GetTokenDefinition("end", "end", IsKeyword:=True)
    ' Brackets
    i = i + 1: Arr(i) = GetTokenDefinition("lBracket", "\(")
    i = i + 1: Arr(i) = GetTokenDefinition("rBracket", "\)")
    ' Functions
    i = i + 1: Arr(i) = GetTokenDefinition("fun", "fun", IsKeyword:=True)
    i = i + 1: Arr(i) = GetTokenDefinition("comma", ",") 'params
    ' Lines
    i = i + 1: Arr(i) = GetTokenDefinition("colon", ":")
    
    'VarName
    i = i + 1: Arr(i) = GetTokenDefinition("arg", "\$\d+")
    i = i + 1: Arr(i) = GetTokenDefinition("var", "[a-zA-Z][a-zA-Z0-9_]*")
    
    'Operators
    i = i + 1: Arr(i) = GetTokenDefinition("propertyAccess", "\.")
    i = i + 1: Arr(i) = GetTokenDefinition("methodAccess", "\#")
    i = i + 1: Arr(i) = GetTokenDefinition("multiply", "\*")
    i = i + 1: Arr(i) = GetTokenDefinition("divide", "\/")
    i = i + 1: Arr(i) = GetTokenDefinition("power", "\^")
    i = i + 1: Arr(i) = GetTokenDefinition("add", "\+")
    i = i + 1: Arr(i) = GetTokenDefinition("subtract", "\-")
    i = i + 1: Arr(i) = GetTokenDefinition("equal", "\=")
    i = i + 1: Arr(i) = GetTokenDefinition("notEqual", "\<\>")
    i = i + 1: Arr(i) = GetTokenDefinition("greaterThanEqual", "\>\=")
    i = i + 1: Arr(i) = GetTokenDefinition("greaterThan", "\>")
    i = i + 1: Arr(i) = GetTokenDefinition("lessThanEqual", "\<\=")
    i = i + 1: Arr(i) = GetTokenDefinition("lessThan", "\<")
    i = i + 1: Arr(i) = GetTokenDefinition("concatenate", "\&")
    
    ReDim Preserve Arr(1 To i)
    GetTokenDefinitions = Arr
    
End Function


'Copies one variant to a destination
'@param {ByRef Token()} tokens Tokens to remove the specified type from
'@param {string} RemoveType   Token type to remove.
'@returns {Token()} The modified token array.
Private Function RemoveTokens(ByRef Tokens() As Token, ByVal RemoveType As String) As Token()
    
    Dim CountRemoved As Long: CountRemoved = 0
    Dim Token As Long
    For Token = LBound(Tokens) To UBound(Tokens)
        If Tokens(Token).Type.Name <> RemoveType Then
            Tokens(Token - CountRemoved) = Tokens(Token)
        Else
            CountRemoved = CountRemoved + 1
        End If
    Next Token
    
    ReDim Preserve Tokens(LBound(Tokens) To (UBound(Tokens) - CountRemoved))
    RemoveTokens = Tokens
    
End Function


Private Function GetTokenDefinition(ByVal Name As String, _
                                    ByVal Regex As String, _
                                    Optional ByVal IgnoreCase As Boolean = True, _
                                    Optional ByVal IsKeyword As Boolean = False) As TokenDefinition

    GetTokenDefinition.Name = Name
    GetTokenDefinition.Regex = Regex & IIf(IsKeyword, "\b", vbNullString)
    Set GetTokenDefinition.RegexObj = CreateObject("VBScript.Regexp")
    GetTokenDefinition.RegexObj.Pattern = "^(?:" & Regex & IIf(IsKeyword, "\b", vbNullString) & ")"
    GetTokenDefinition.RegexObj.IgnoreCase = IgnoreCase
    
End Function


