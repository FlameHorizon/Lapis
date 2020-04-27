Attribute VB_Name = "Comparers"
Option Explicit

Public Function Default(ByVal Value As Variant) As IComparer
    
    Dim VarT As VbVarType
    VarT = VarType(Value)
    
    Select Case VarT
        Case VbVarType.vbString
            Set Default = ComparerStatic.Default
        
        Case VbVarType.vbLong Or VbVarType.vbInteger Or VbVarType.vbByte Or VbVarType.vbBoolean
            Set Default = New LongComparer
        
        Case VbVarType.vbSingle Or VbVarType.vbDouble Or VbVarType.vbCurrency Or VbVarType.vbDecimal Or VbVarType.vbDate
            Set Default = New DecimalComparer
            
        Case Else
            Errors.OnArgumentOutOfRange "VarType", "Comparers.Default"
    
    End Select
    
End Function

