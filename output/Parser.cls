Attribute VB_Name = "Parser"
private tokens as variant
private current as long
private m_h as H
private statements as variant
private functcounter as long
public function init(byval tokens_ as variant, byval hh as H) 
    nt st (tokens, tokens_)
    nt st (current, 0)
    nt st (statements, array())
    nt st (functcounter, 0)
    nt st (m_h, hh)
end function


public function parse() 
    while (not isAtEnd())
        call arad(statements, declaration_f())
        looplabel0bef:
        nt 1
    wend
    looplabel0aft:
    st parse, statements
    exit function
end function


private function declaration_f() as stmt
    dim  vis as string: st vis, ""
    if (match(PUBLIC_)) then 
        nt st (vis, "public")
    else
        if (match(PRIVATE_)) then 
            nt st (vis, "private")
        endif
    endif
    if ((peek().m_type_ = IDENTIFIER) and (peeknext().m_type_ = IDENTIFIER)) then 
        st declaration_f, varDeclaration_f(vis)
        exit function
    else
        if (match(FUNC_)) then 
            st declaration_f, function_("function", vis)
            exit function
        else
            if (match(CLASS_)) then 
                st declaration_f, classDeclaration_f()
                exit function
            else
                if (match(ENUM_)) then 
                    st declaration_f, enumDeclaration_f()
                    exit function
                endif
            endif
        endif
    endif
    st declaration_f, statement_f()
    exit function
end function


private function enumDeclaration_f() as EnumStmt
    dim  name as Token: st name, consume(IDENTIFIER, "expected enum name")
    dim  types_() as variant
    call consume(LEFT_BRACE, "Expected '{' after enum")
    call arad(types_, consume(IDENTIFIER, "Enum needs at least one name"))
    while (match(COMMA))
        call arad(types_, consume(IDENTIFIER, "Expected identifier"))
        looplabel1bef:
        nt 1
    wend
    looplabel1aft:
    call consume(RIGHT_BRACE, "Expected '}' after enum declaration")
    dim  temp as EnumStmt: st temp, new EnumStmt
    call temp.init(name, types_)
    st enumDeclaration_f, temp
    exit function
end function


private function classDeclaration_f() as ClassStmt
    dim  name as Token: st name, consume(IDENTIFIER, "Expected class name")
    dim  supercs() as variant
    dim  superc as token
    dim  hassuperclass as boolean: st hassuperclass, False
    if (match(LESS)) then 
        nt st (superc, consume(IDENTIFIER, "Expected parent class name"))
        call arad(supercs, superc)
        while (match(COMMA))
            nt st (superc, consume(IDENTIFIER, "Expected parent class name"))
            call arad(supercs, superc)
            looplabel2bef:
            nt 1
        wend
        looplabel2aft:
    endif
    call consume(LEFT_BRACE, "Expect '{' before class body")
    dim  methods() as variant
    dim  properties() as variant
    while (not check(RIGHT_BRACE) and not isAtEnd())
        looplabel3bef:
        nt 1
    wend
    looplabel3aft:
end function



