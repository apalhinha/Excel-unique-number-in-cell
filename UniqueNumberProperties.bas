Attribute VB_Name = "UniqueNumberProperties"
Public Function GetUniqueNumber() As Long
Dim res As Long

On Error GoTo e
    'Selection.Formula = 123
    res = ActiveWorkbook.CustomDocumentProperties("Unique Number")
    res = res + 1
    ActiveWorkbook.CustomDocumentProperties("Unique Number") = res
    GetUniqueNumber = res
    Exit Function
    
e:
    On Error Resume Next
    ActiveWorkbook.CustomDocumentProperties.Add Name:="Unique Number", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:=1
    GetUniqueNumber = 1
    Exit Function
    
End Function

Public Sub InsertUniqueNumber()
    Selection.Formula = GetUniqueNumber()
End Sub
