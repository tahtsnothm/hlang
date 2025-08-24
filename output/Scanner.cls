VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Scanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
private m_source as string
private m_tokens() as variant
private m_start as long
private m_current as long
private m_line as long
private m_h as H
private m_keywords as Dictionary
public function init(byval source as string, byref h_instance as H) 
    nt st (m_source, preprocess(source))
    nt st (m_tokens, array())
    nt st (m_start, 0)
    nt st (m_current, 0)
    nt st (m_line, 1)
    nt st (m_h, h_instance)
    nt st (m_keywords, new Dictionary)
    nt st (m_keywords("and"), AND_)
    nt st (m_keywords("class"), CLASS_)
    nt st (m_keywords("else"), ELSE_)
    nt st (m_keywords("false"), FALSE_)
    nt st (m_keywords("for"), FOR_)
    nt st (m_keywords("func"), FUNC_)
    nt st (m_keywords("if"), IF_)
    nt st (m_keywords("null"), NULL_)
    nt st (m_keywords("or"), OR_)
    nt st (m_keywords("print"), PRINT_)
    nt st (m_keywords("return"), RETURN_)
    nt st (m_keywords("super"), SUPER_)
    nt st (m_keywords("this"), THIS_)
    nt st (m_keywords("true"), TRUE_)
    nt st (m_keywords("var"), VAR_)
    nt st (m_keywords("while"), WHILE_)
    nt st (m_keywords("new"), NEW_)
    nt st (m_keywords("public"), PUBLIC_)
    nt st (m_keywords("private"), PRIVATE_)
    nt st (m_keywords("static"), STATIC_)
    nt st (m_keywords("break"), BREAK_)
    nt st (m_keywords("continue"), CONTINUE_)
    nt st (m_keywords("enum"), ENUM_)
end function


private function preprocess(byval source as string) as string
    st preprocess, replace(source, "_", "_")
    exit function
end function


public function scanTokens() as variant
    while (not isAtEnd())
        nt st (m_start, m_current)
        call scanToken()
        looplabel0bef:
        nt 1
    wend
    looplabel0aft:
    dim  newtoken as Token: st newtoken, new Token
    call newtoken.init(EOF_, "", "null", m_line)
    call arad(m_tokens, newtoken)
    st scanTokens, m_tokens
    exit function
end function


private function isAtEnd() as boolean
    st isAtEnd, m_current >= len(m_source)
    exit function
end function


private function scanToken() 
     dim rawquote as string : rawquote = """" 
    dim  c as string: st c, advance()
    if (c = "(") then 
        call addToken(LEFT_PAREN)
    else
        if (c = ")") then 
            call addToken(RIGHT_PAREN)
        else
            if (c = "{") then 
                call addToken(LEFT_BRACE)
            else
                if (c = "}") then 
                    call addToken(RIGHT_BRACE)
                else
                    if (c = "[") then 
                        call addToken(LEFT_SQUARE)
                    else
                        if (c = "]") then 
                            call addToken(RIGHT_SQUARE)
                        else
                            if (c = ",") then 
                                call addToken(COMMA)
                            else
                                if (c = ".") then 
                                    call addToken(DOT)
                                else
                                    if (c = ":") then 
                                        call addToken(COLON)
                                    else
                                        if (c = "%") then 
                                            call addToken(MOD_)
                                        else
                                            if (c = "-") then 
                                                if (match("=")) then 
                                                    call addToken(MINUS_EQUAL)
                                                else
                                                    if (match("-")) then 
                                                        call addToken(MINUS_MINUS)
                                                    else
                                                        call addToken(MINUS)
                                                    endif
                                                endif
                                            else
                                                if (c = "+") then 
                                                    if (match("=")) then 
                                                        call addToken(PLUS_EQUAL)
                                                    else
                                                        if (match("+")) then 
                                                            call addToken(PLUS_PLUS)
                                                        else
                                                            call addToken(PLUS)
                                                        endif
                                                    endif
                                                else
                                                    if (c = ";") then 
                                                        call addToken(SEMICOLON)
                                                    else
                                                        if (c = "*") then 
                                                            if (match("=")) then 
                                                                call addToken(STAR_EQUAL)
                                                            else
                                                                call addToken(STAR)
                                                            endif
                                                        else
                                                            if (c = "&") then 
                                                                if (match("&")) then 
                                                                    call addToken(BRANCHING_AND_)
                                                                else
                                                                    call addToken(AMPERSAND)
                                                                endif
                                                            else
                                                                if (c = "#") then 
                                                                    call raw_code_f()
                                                                else
                                                                    if (c = "!") then 
                                                                        if (match("=")) then 
                                                                            call addToken(BANG_EQUAL)
                                                                        else
                                                                            call addToken(BANG)
                                                                        endif
                                                                    else
                                                                        if (c = "=") then 
                                                                            if (match("=")) then 
                                                                                call addToken(EQUAL_EQUAL)
                                                                            else
                                                                                if (match(">")) then 
                                                                                    call addToken(ARROW_)
                                                                                else
                                                                                    call addToken(EQUAL)
                                                                                endif
                                                                            endif
                                                                        else
                                                                            if (c = ">") then 
                                                                                if (match("=")) then 
                                                                                    call addToken(GREATER_EQUAL)
                                                                                else
                                                                                    call addToken(GREATER)
                                                                                endif
                                                                            else
                                                                                if (c = "<") then 
                                                                                    if (match("=")) then 
                                                                                        call addToken(LESS_EQUAL)
                                                                                    else
                                                                                        call addToken(LESS)
                                                                                    endif
                                                                                else
                                                                                    if (c = "/") then 
                                                                                        if (match("/")) then 
                                                                                            while (peek() <> vblf and not isAtEnd())
                                                                                                call advance()
                                                                                                looplabel1bef:
                                                                                                nt 1
                                                                                            wend
                                                                                            looplabel1aft:
                                                                                        else
                                                                                            if (match("=")) then 
                                                                                                call addToken(SLASH_EQUAL)
                                                                                            else
                                                                                                if (match("*")) then 
                                                                                                    call blockcomment_f()
                                                                                                else
                                                                                                    call addToken(SLASH)
                                                                                                endif
                                                                                            endif
                                                                                        endif
                                                                                    else
                                                                                        if (c = " " or c = chr(9)) then 
                                                                                        else
                                                                                            if (c = vblf) then 
                                                                                                nt st (m_line, m_line + 1)
                                                                                            else
                                                                                                if (c = rawquote) then 
                                                                                                    call string_f()
                                                                                                else
                                                                                                    if (isDigit(c)) then 
                                                                                                        call number_f()
                                                                                                    else
                                                                                                        if (isAlpha(c)) then 
                                                                                                            call identifier_f()
                                                                                                        else
                                                                                                            call m_h.error(m_line, "Unexpected character.")
                                                                                                        endif
                                                                                                    endif
                                                                                                endif
                                                                                            endif
                                                                                        endif
                                                                                    endif
                                                                                endif
                                                                            endif
                                                                        endif
                                                                    endif
                                                                endif
                                                            endif
                                                        endif
                                                    endif
                                                endif
                                            endif
                                        endif
                                    endif
                                endif
                            endif
                        endif
                    endif
                endif
            endif
        endif
    endif
end function


private function addToken(byval type_ as tokentype) 
    call addTokenInternal(type_, "null")
end function


private function addTokenInternal(byval type_ as tokentype, byval literal as variant) 
    dim  text as string: st text, substring(m_source, m_start, m_current)
    dim  newtoken as Token: st newtoken, new Token
    call newtoken.init(type_, text, literal, m_line)
    call arad(m_tokens, newtoken)
end function


private function advance() as string
    nt st (m_current, m_current + 1)
    st advance, charAt(m_source, m_current - 1)
    exit function
end function


private function match(byval expected as string) as boolean
    if (isAtEnd()) then 
        st match, False
        exit function
    endif
    if (charAt(m_source, m_current) <> expected) then 
        st match, False
        exit function
    endif
    nt st (m_current, m_current + 1)
    st match, True
    exit function
end function


private function peek() as string
    if (isAtEnd()) then 
        st peek, ""
        exit function
    endif
    st peek, substring(m_source, m_current, m_current + 1)
    exit function
end function


private function peeknext() as string
    if (m_current + 1 >= len(m_source)) then 
        st peeknext, ""
        exit function
    endif
    st peeknext, substring(m_source, m_current + 1, m_current + 2)
    exit function
end function


private function raw_code_f() 
    while (peek() <> "#" and not isAtEnd())
        if (peek() = vblf) then 
            nt st (m_line, m_line + 1)
        endif
        call advance()
        looplabel2bef:
        nt 1
    wend
    looplabel2aft:
    if (isAtEnd()) then 
        call m_h.error(m_line, "Unterminated Raw Code Block")
        exit function
    endif
    call advance()
    dim  value as string: st value, substring(m_source, m_start + 1, m_current - 1)
    call addTokenInternal(RAW_, value)
end function


private function string_f() 
     dim rawquote as string : rawquote = """" 
    while (peek() <> rawquote and not isAtEnd())
        if (peek() = vblf) then 
            nt st (m_line, m_line + 1)
        endif
        call advance()
        looplabel3bef:
        nt 1
    wend
    looplabel3aft:
    if (isAtEnd()) then 
        call m_h.error(m_line, "Unterminated string")
        exit function
    endif
    call advance()
    dim  value as string: st value, substring(m_source, m_start + 1, m_current - 1)
    call addTokenInternal(STRING_, value)
end function


private function blockcomment_f() 
    while (True)
        if (peek() = vblf) then 
            nt st (m_line, m_line + 1)
        endif
        if (peek() = "*" and peeknext() = "/") then 
            call advance()
            call advance()
            goto looplabel4aft:
        endif
        if (isAtEnd()) then 
            exit function
        endif
        call advance()
        looplabel4bef:
        nt 1
    wend
    looplabel4aft:
end function


private function number_f() 
    while (isDigit(peek()))
        call advance()
        looplabel5bef:
        nt 1
    wend
    looplabel5aft:
    if (peek() = "." and isDigit(peeknext())) then 
        call advance()
        while (isDigit(peek()))
            call advance()
            looplabel6bef:
            nt 1
        wend
        looplabel6aft:
    endif
    dim  num as string: st num, substring(m_source, m_start, m_current)
    call addTokenInternal(NUMBER_, clng(num))
end function


private function identifier_f() 
    while (isAlphaNumeric(peek()) or peek() = ".")
        call advance()
        looplabel7bef:
        nt 1
    wend
    looplabel7aft:
    dim  s as string: st s, substring(m_source, m_start, m_current)
    dim  type_ as tokentype
    if (m_keywords.exists(s)) then 
        nt st (type_, m_keywords(s))
    else
        nt st (type_, IDENTIFIER)
    endif
    call addToken(type_)
end function


private function isAlpha(byval c as string) as boolean
    st isAlpha, (asc(c) >= asc("a") and asc(c) <= asc("z")) or (asc(c) >= asc("A") and asc(c) <= asc("Z")) or c = "_"
    exit function
end function


private function isDigit(byval c as string) as boolean
    st isDigit, isNumeric(c)
    exit function
end function


private function isAlphaNumeric(byval c as string) as boolean
    st isAlphaNumeric, isAlpha(c) or isDigit(c)
    exit function
end function



