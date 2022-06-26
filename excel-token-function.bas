Public Function TOKEN(ParamArray var() As Variant) As String

' Excel add-in function that returns a single substring from a string
' expression, given custom delimiters and the substring position.

' The function takes as input:
' - A string to tokenize, as required 1st parameter,
' - A string of delimiters, as optional 2nd or 3rd parameter,
' - An integer token position (counted from 1), as optional 2nd or 3rd
'   parameter,

' The function returns:
' - A substring (token) of the input string.

' Function's logic:
' - The function identifies within the input string the longest sequences of
'   characters from delimiters and uses these sequences to identify substrings.

' Input modes:
' - Parameters may be entered by value (e.g., TOKEN("abc def", " ", 2)) or by
'   reference (e.g., TOKEN(A1, A2, A3)), or by a mix of the two.
' - If the string of delimiters is omitted, the default is a string made of the
'   space character (" ").
' - If the string of delimiters is passed by value (e.g., TOKEN(A1, " ,;/", 2),
'   the surrounding quotation marks are ignored; if it is passed by reference
'   (e.g., TOKEN(A1, A2, 2), the surrounding quotation marks are included in
'   the set of delimiter characters.
' - If the position is omitted, the default is 1.

' Exception handling:
' - If the function is invoked without parameters, it returns a description of
'   its call syntax.
' - If the input string is an empty string ("") or an empty cell, the function
'   returns an empty string.
' - If the string of delimiters is an empty string (""), the function returns
'   the entire input string (if position=1, otherwise "Not enough tokens.".
' - If the 2nd parameter is numeric while the 3rd parameter is omitted or empty,
'   or if the 3rd parameter is numeric while the 2nd parameter is empty, then
'   that numeric parameter is interpreted as token position.

' VBScript.RegExp documentation and types of RegExp objects used by the
'   function:
' - https://www.regular-expressions.info/vbscript.html
' - my_regex = object of TypeName IRegExp2.
' - matches = object of TypeName IMatchCollection2 = a collection.
' - matches(0) = matches.Item(0) = object of TypeName IMatch2.
' - matches(0).submatches = matches.Item(0).submatches = object of TypeName
'   ISubMatches = a collection.
' - matches(0).submatches(0) = matches.Item(0).submatches.Item(0) = a string.

    Dim my_str As String
    Dim my_pos As Integer
    Dim my_del As String
    Dim my_pattern As String
    Dim my_regex As Object
    Dim matches As Variant
    
    Select Case (UBound(var) - LBound(var) + 1)
        Case 0
            TOKEN = "Syntax: TOKEN(<string to tokenize> [, <string of " + _
            "delimiters(def="" "")>] [, <position of token(def=1)>])"
            Exit Function
        Case 1
            If IsEmpty(var(0)) Or var(0) = "" Then
                TOKEN = ""
                Exit Function
            Else
                my_str = CStr(var(0)): my_pos = 1: my_del = " "
            End If
        Case 2
            If IsEmpty(var(0)) Or var(0) = "" Then
                TOKEN = ""
                Exit Function
            ElseIf IsEmpty(var(1)) Or var(1) = "" Then
                TOKEN = var(0)
                Exit Function
            ' If var(1) is a number, var(1)=position.
            ElseIf WorksheetFunction.IsNumber(var(1)) Then
                my_str = CStr(var(0))
                my_pos = CInt(var(1))
                my_del = " "
            Else
                my_str = CStr(var(0))
                my_pos = 1
                my_del = CStr(var(1))
            End If
        Case 3
            If IsEmpty(var(0)) Or var(0) = "" Then
                TOKEN = ""
                Exit Function
            ElseIf IsEmpty(var(1)) Or var(1) = "" Then
                If IsEmpty(var(2)) Or var(2) = "" Then
                    ' Assume delimiters="" and position=1.
                    TOKEN = var(0)
                    Exit Function
                ElseIf WorksheetFunction.IsNumber(var(2)) Then
                    ' var(2)=position, assume delimiters="".
                    If var(2) = 1 Then
                        TOKEN = var(0)
                        Exit Function
                    Else
                        TOKEN = "Not enough tokens."
                        Exit Function
                    End If
                Else
                    ' var(2)=delimiters, assume position=1.
                    my_str = CStr(var(0))
                    my_pos = 1
                    my_del = CStr(var(2))
                End If
            ElseIf IsEmpty(var(2)) Or var(2) = "" Then
                If IsEmpty(var(1)) Or var(1) = "" Then
                    ' Assume delimiters="" and position=1.
                    TOKEN = var(0)
                    Exit Function
                ElseIf WorksheetFunction.IsNumber(var(1)) Then
                    ' var(1)=position, assume delimiters="".
                    If var(1) = 1 Then
                        TOKEN = var(0)
                        Exit Function
                    Else
                        TOKEN = "Not enough tokens."
                        Exit Function
                    End If
                Else
                    ' var(1)=delimiters, assume position=1.
                    my_str = CStr(var(0))
                    my_pos = 1
                    my_del = CStr(var(1))
                End If
            ' If var(1) and var(2) are both numbers,
            ' assume var(1)=delimiters, var(2)=position.
            ElseIf WorksheetFunction.IsNumber(var(2)) Then
                my_str = CStr(var(0))
                my_pos = CInt(var(2))
                my_del = CStr(var(1))
            ElseIf WorksheetFunction.IsNumber(var(1)) Then
                my_str = CStr(var(0))
                my_pos = CInt(var(1))
                my_del = CStr(var(2))
            Else
                TOKEN = "One of the two optional parameters should be a number."
                Exit Function
            End If
        Case Else
            TOKEN = "Too many arguments."
            Exit Function
    End Select

    ' The regular expression (with one capturing group) is:
    ' [my_del]*([^my_del]+)
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
