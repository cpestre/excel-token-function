Public Function TOKEN(ParamArray var() As Variant) As String

' Excel add-in function; takes as input:
' - a string to tokenize,
' - an integer token position (optional),
' - a string of delimiters (optional),
' and returns:
' - a string (token).
' The function displays its call syntax when invoked without argument.
' The function treats any sequence of characters from the string of delimiters as a single delimiter.
' Variables:
' - var() may hold 0-3 parameters; if none, the function displays its call syntax.
' - var(0) = my_str = string to tokenize; entered as a string or a cell.
' - var(1) = my_pos = position of token (default = first = 1); entered as an int or a cell.
' - var(2) = my_del = string of delimiters (default = space = " "); entered as a string or a cell.
' - TOKEN = the output token.
' VBScript.RegExp documentation at https://www.regular-expressions.info/vbscript.html
' - my_regex = object of TypeName IRegExp2.
' - matches = object of TypeName IMatchCollection2 = a collection.
' - matches(0) = matches.Item(0) = object of TypeName IMatch2.
' - matches(0).submatches = matches.Item(0).submatches = object of TypeName ISubMatches = a collection.
' - matches(0).submatches(0) = matches.Item(0).submatches.Item(0) = a string.

    Dim my_str As String
    Dim my_pos As Integer
    Dim my_del As String
    Dim my_pattern As String
    Dim my_regex As Object
    Dim matches As Variant
    
    Select Case (UBound(var) - LBound(var) + 1)
        Case 0
            TOKEN = "Syntax: TOKEN(<string to tokenize> [, <position of token(def=1)> [, <string of delimiters(def="" "")])"
            Exit Function
        Case 1
            my_str = CStr(var(0)): my_pos = 1: my_del = " "
        Case 2
            my_str = CStr(var(0)): my_pos = CInt(var(1)): my_del = " "
        Case 3
            my_str = CStr(var(0)): my_pos = CInt(var(1)): my_del = CStr(var(2))
        Case Else
            TOKEN = "Too many arguments."
            Exit Function
    End Select

    ' The regular expression (with one capturing group) is [my_del]*([^my_del]+).
    my_pattern = "[" & my_del & "]*" & "([^" & my_del & "]+)"

    Set my_regex = CreateObject("VBScript.RegExp")
    
    With my_regex
        .pattern = my_pattern
        .Global = True
    End With
    
    Set matches = my_regex.Execute(my_str)
    
    If my_pos > matches.Count Then
        TOKEN = "Not enough tokens."
        Exit Function
    End If
    
    ' The only capturing group, SubMatches(0), is the TOKEN.
    TOKEN = matches.Item(my_pos - 1).submatches.Item(0)

End Function