Attribute VB_Name = "Comparers"
Option Explicit

Public Function Default(ByVal Value As Variant) As IComparer
    
    Dim VarT As VbVarType
    VarT = VarType(Value)
    
    Select Case VarT
        Case VbVarType.vbString
            Set Default = StringComparer.OrdinalIgnoreCase
        
        Case VbVarType.vbLong, VbVarType.vbInteger, VbVarType.vbByte, VbVarType.vbBoolean
            Set Default = New LongComparer
        
        Case VbVarType.vbSingle, VbVarType.vbDouble, VbVarType.vbCurrency, VbVarType.vbDecimal, VbVarType.vbDate
            Set Default = New DecimalComparer
            
        Case Else
            Errors.OnArgumentOutOfRange "VarType", "Comparers.Default"
    
    End Select
    
End Function

