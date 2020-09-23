<div align="center">

## Easy Tokenizer


</div>

### Description

Break up variables in a string, whatever separator is used.
 
### More Info
 
String to be Tokenized, Separator between variables

List of individual variables


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Paul Crowdy](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-crowdy.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/paul-crowdy-easy-tokenizer__1-26012/archive/master.zip)

### API Declarations

```
Public Type TokenList
  Tokens() As String
  TokenCount As Integer
End Type
Public Function Tokenize(strString As String, strSeparator As String) As TokenList
  Dim iCount As Integer
  Dim iStart As Integer
  Dim iTokens As Integer
  ReDim Tokenize.Tokens(0)
  iTokens = 0
  iCount = 1
  iStart = 1
  Do Until iCount = Len(strString)
    If Mid$(strString, iCount, Len(strSeparator)) = strSeparator Then
      ReDim Preserve Tokenize.Tokens(iTokens + 1)
      Tokenize.Tokens(iTokens) = Mid$(strString, iStart, iCount - iStart)
      iStart = iCount + Len(strSeparator)
      iTokens = iTokens + 1
    End If
    iCount = iCount + 1
  Loop
  Tokenize.Tokens(iTokens) = Mid$(strString, iStart)
  Tokenize.TokenCount = iTokens + 1
End Function
```


### Source Code

```
Private Sub Form_Load()
  Dim tk As TokenList
  Dim strTest As String
  Dim strSeparator As String
  strTest = "String, Tokenization, By, Paul, Crowdy, www.kmcpartnership.co.uk"
  strSeparator = ", "
  tk = Tokenize(strTest, strSeparator)
  For i = 0 To tk.TokenCount - 1
    MsgBox "Token " & i + 1 & " = " & tk.Tokens(i), vbInformation, strTest
  Next i
  End
End Sub
```

