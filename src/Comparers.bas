Attribute VB_Name = "Comparers"
'@Folder("Collection")
Option Explicit


Public Function Default(ByVal Value As Variant) As IComparer
    
    Dim VarT As VbVarType
    VarT = VarType(Value)
    
    Select Case VarT
        Case VbVarType.vbString
            Set Default = StringComparers.OrdinalIgnoreCase
        
        Case VbVarType.vbLong, VbVarType.vbInteger, VbVarType.vbByte, VbVarType.vbBoolean, VbVarType.vbLongLong
            Set Default = New LongLongComparer
        
        Case VbVarType.vbSingle, VbVarType.vbDouble, VbVarType.vbCurrency, VbVarType.vbDecimal, VbVarType.vbDate
            Set Default = New DecimalComparer
            
        Case Else
            Set Default = Nothing
    
    End Select
    
End Function


