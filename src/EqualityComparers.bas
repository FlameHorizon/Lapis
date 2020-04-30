Attribute VB_Name = "EqualityComparers"
'@Folder("Collection")
Option Explicit

Public Function Default(ByVal Value As Variant) As IEqualityComparer
    
    Dim VarT As VbVarType
    VarT = VarType(Value)
    
    Select Case VarT
        Case VbVarType.vbString
            Set Default = StringEqualityComparers.OrdinalIgnoreCase
        
        Case VbVarType.vbLong, VbVarType.vbInteger, VbVarType.vbByte, VbVarType.vbBoolean, VbVarType.vbLongLong
            Set Default = New LongLongEqualityComparer
        
        Case VbVarType.vbSingle, VbVarType.vbDouble, VbVarType.vbCurrency, VbVarType.vbDecimal, VbVarType.vbDate
            Set Default = New DecimalEqualityComparer
            
        Case Else
            Set Default = Nothing
    
    End Select
    
End Function
