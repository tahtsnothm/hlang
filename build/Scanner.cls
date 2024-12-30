VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Scanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_source As String
Private m_tokens As Variant
Private m_start As Long
Private m_current As Long
Private m_line As Long
Private m_h As H
Private m_keywords As New Scripting.Dictionary
Private Sub Class_Initialize()
    m_source = ""
    m_tokens = Array()
    m_start = 0
    m_current = 0
    m_line = 1
    m_keywords("and") = AND_
    m_keywords("class") = CLASS_
m_keywords("else") = ELSE_
m_keywords("false") = FALSE_
m_keywords("for") = FOR_
m_keywords("func") = FUNC_
m_keywords("if") = IF_
m_keywords("null") = NULL_
m_keywords("or") = OR_
m_keywords("print") = PRINT_
m_keywords("return") = RETURN_
m_keywords("super") = SUPER_
m_keywords("this") = THIS_
m_keywords("true") = TRUE_
m_keywords("var") = VAR_
m_keywords("while") = WHILE_

    
End Sub
Public Function init(source As String, h_instance As H)
    m_source = preprocess(source)
    Set m_h = h_instance
End Function
Private Function preprocess(source As String) As String
    preprocess = Replace(source, "[", "(")
    preprocess = Replace(preprocess, "]", ")")
End Function

Public Function scanTokens() As Variant ' lsittokens
While (Not isAtEnd())
    m_start = m_current
    scanToken
Wend
    Dim newtoken As New Token
    newtoken.init EOF_, "", Null, m_line
    
    arad m_tokens, newtoken
    scanTokens = m_tokens



End Function


Private Function isAtEnd() As Boolean
    isAtEnd = m_current >= Len(m_source)
End Function

Private Function scanToken()
    Dim c As String 'char data type
    c = advance()
    Select Case c
        Case "(": addToken LEFT_PAREN
        Case ")": addToken RIGHT_PAREN
        Case "{": addToken LEFT_BRACE
        Case "}": addToken RIGHT_BRACE
        Case ",": addToken COMMA
        Case ".": addToken DOT
        Case ":": addToken COLON
        Case "-":
            If match("=") Then
                addToken MINUS_EQUAL
            Else
                addToken MINUS
            End If
        Case "+":
            If match("=") Then
                addToken PLUS_EQUAL
            Else
                addToken PLUS
            End If
        Case ";": addToken SEMICOLON
        Case "*":
            If match("=") Then
                addToken STAR_EQUAL
            Else
                addToken STAR
            End If
        Case "&":
            If match("&") Then
                addToken BRANCHING_AND_
            Else
                addToken AMPERSAND
            End If
        Case "#": raw_code_f
        Case "!":
            If match("=") Then
                addToken BANG_EQUAL
            Else
                addToken BANG
            End If
        Case "=":
            If match("=") Then
                addToken EQUAL_EQUAL
            Else
                addToken EQUAL
            End If
        Case ">":
            If match("=") Then
                addToken GREATER_EQUAL
            Else
                addToken GREATER
            End If
        Case "<":
            If match("=") Then
                addToken LESS_EQUAL
            Else
                addToken LESS
            End If
        Case "/":
            If (match("/")) Then
                While (peek() <> vbLf And Not isAtEnd()) ' NEWLINE CHARACTER POTENTIAL SOURCE OF ISSUES
                    advance
                Wend
            ElseIf match("=") Then
                addToken SLASH_EQUAL
            ElseIf (match("*")) Then
                blockcomment_f
            Else
                addToken SLASH
            End If
        Case " ":
        Case Chr(9): 'tab
        Case vbLf:
            m_line = m_line + 1
        Case """":
            string_f
        
        Case Else
            If (IsNumeric(c)) Then
                number_f
            ElseIf (isAlpha(c)) Then
                identifier_F
            Else
                m_h.error line, "Unexpected character."
            End If
            
    End Select
    

End Function
Private Function addToken(type_ As tokentype)
    addTokenInternal type_, Null
End Function
Private Function addTokenInternal(type_ As tokentype, literal As Variant)
    Dim text As String
    text = substring(m_source, m_start, m_current)
    Dim newtoken As New Token
    newtoken.init type_, text, literal, m_line
    m_tokens = aradv(m_tokens, newtoken)
End Function
Private Function advance()
m_current = m_current + 1
advance = charAt(m_source, m_current - 1)

End Function
Private Function match(expected As String) As Boolean
    If (isAtEnd()) Then
        match = False
        Exit Function
    End If
    If (Mid(m_source, m_current + 1, 1) <> expected) Then
        match = False
        Exit Function
    End If
    
    m_current = m_current + 1
    match = True
End Function
Private Function peek() As String
    If (isAtEnd()) Then
        peek = ""
        Exit Function
    End If
    peek = substring(m_source, m_current, m_current + 1)
End Function
Private Function raw_code_f()

While (peek() <> "#" And Not isAtEnd())
    If peek() = vbLf Then
        m_line = m_line + 1
    End If
    advance
Wend
If (isAtEnd()) Then
    m_h.error m_line, "Unterminated Raw Code Block"
    Exit Function
End If

advance ' closing #
Dim value As String
value = substring(m_source, m_start + 1, m_current - 1)
'value = Mid(m_source, m_start + 2, m_current - 1 - (m_start + 1) + 1) ' convoluted
addTokenInternal RAW_, value

End Function

Private Function string_f()
While (peek() <> """" And Not isAtEnd())
    If peek() = vbLf Then
        m_line = m_line + 1
    End If
    advance
Wend
If (isAtEnd()) Then
    m_h.error m_line, "Unterminated String"
    Exit Function
End If

advance ' closing "
Dim value As String
value = substring(m_source, m_start + 1, m_current - 1)
'value = Mid(m_source, m_start + 2, m_current - 1 - (m_start + 1) + 1) ' convoluted
addTokenInternal STRING_, value
End Function

Private Function blockcomment_f()
    Dim endloop As Boolean: endloop = False
    While (Not endloop)
        If peek() = vbLf Then
            m_line = m_line + 1
        End If
        If (peek() = "*") Then
            If (peeknext() = "/") Then
                endloop = True
                GoTo hehehaha
            End If
        End If
        advance
        If (isAtEnd()) Then endloop = True
hehehaha:
    Wend
    
    advance
    advance 'for the skibidi
    Dim value As String
        value = substring(m_source, m_start + 2, m_current - 2)
    If (isAtEnd()) Then
        'm_h.error m_line, "Unterminated block comment."
    Exit Function
End If
End Function

Private Function number_f()
While IsNumeric(peek())
    advance
Wend

If (peek() = "." And IsNumeric(peeknext())) Then
    advance
    
    While (IsNumeric(peek()))
        advance
    Wend
End If
addTokenInternal NUMBER_, CLng(substring(m_source, m_start, m_current))
End Function

Private Function identifier_F()
While (isAlphaNumeric(peek()) Or peek() = ".")
advance
Wend
Dim s As String

s = substring(m_source, m_start, m_current)
Dim type_ As tokentype

If (Not m_keywords.Exists(s)) Then
    type_ = IDENTIFIER
Else
    type_ = m_keywords(s)
End If



addToken type_

End Function


Private Function peeknext()
If (m_current + 1 >= Len(m_source)) Then
    peeknext = ""
    Exit Function
End If
peeknext = substring(m_source, m_current + 1, m_current + 2)

End Function

Private Function isAlpha(c As String)
    If ((Asc(c) >= Asc("a") And Asc(c) <= Asc("z")) Or (Asc(c) >= Asc("A") And Asc(c) <= Asc("Z")) Or c = "_") Then
        isAlpha = True
    Else
        isAlpha = False
    End If
End Function
Private Function isAlphaNumeric(c As String)
    isAlphaNumeric = isAlpha(c) Or IsNumeric(c)
End Function
