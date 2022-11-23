# Excel-Bold-Formula
Excel Formula for Bold Text

![Bold Formula](/repository/asset/imgbold.png?raw=true "Employee Data title")

```vb
Function BOLD(txt As String) As String
Dim X() As Byte
Dim Y() As Long
Dim Temp As String
Dim WF As WorksheetFunction

Set WF = WorksheetFunction
X = StrConv(txt, vbFromUnicode)
ReDim Y(UBound(X))

For i = 0 To UBound(X)
    Select Case X(i)
    Case 48 To 57  'Number 0-9
        Y(i) = X(i) + 120734
        Temp = Temp & WF.Unichar(Y(i))
    Case 65 To 90   'A-Z
        Y(i) = X(i) + 120211
        Temp = Temp & WF.Unichar(Y(i))
    Case 97 To 122   'a-z
        Y(i) = X(i) + 120205
        Temp = Temp & WF.Unichar(Y(i))
    Case Else
        Temp = Temp & Chr(X(i))
    End Select
Next
BOLD = Temp
End Function
```
