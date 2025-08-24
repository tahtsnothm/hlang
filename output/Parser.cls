VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
        dim  vis as string: st vis, ""
        if (match(PUBLIC_)) then 
            nt st (vis, "public")
        else
            if (match(PRIVATE_)) then 
                nt st (vis, "private")
            else
                if (match(STATIC_)) then 
                    nt st (vis, "static")
                endif
            endif
        endif
        if ((peek().m_type_ = IDENTIFIER) and (peeknext().m_type_ = IDENTIFIER)) then 
            call arad(methods, function_("method", vis))
        else
            call arad(properties, varDeclaration_f(vis))
        endif
        looplabel3bef:
        nt 1
    wend
    looplabel3aft:
    call consume(RIGHT_BRACE, "Expected '{' after class body")
    dim  temp as ClassStmt: st temp, new ClassStmt
    call temp.init(name, methods, supercs, properties)
    st classDeclaration_f, temp
    exit function
end function


private function function_(byval kind as string, byval vis as string) as FunctionStmt
    dim  returntype as Token: st returntype, consume(IDENTIFIER, "Expected return type before function name")
    dim  name as Token: st name, consume(IDENTIFIER, "Expected " & kind & " name")
    call consume(LEFT_PAREN, "Expected '(' after " & kind & " name")
    dim  parameters() as variant
    dim  parammethods() as variant
    dim  parametertypes() as variant
    if (not check(RIGHT_PAREN)) then 
        if (al(parameters) > 254) then 
            call err(peek(), "cant have more than 255 params")
        endif
        call arad(parametertypes, consume(IDENTIFIER, "expected param type"))
        call arad(parammethods, match(AMPERSAND))
        call arad(parameters, consume(IDENTIFIER, "Expected parameter name"))
        while (match(COMMA))
            if (al(parameters) > 254) then 
                call err(peek(), "cant have more than 255 params")
            endif
            call arad(parametertypes, consume(IDENTIFIER, "expected param type"))
            call arad(parammethods, match(AMPERSAND))
            call arad(parameters, consume(IDENTIFIER, "Expected parameter name"))
            looplabel4bef:
            nt 1
        wend
        looplabel4aft:
    endif
    call consume(RIGHT_PAREN, "Expected ')' after parameters")
    call consume(LEFT_BRACE, "Expect '{' before " & kind & " body")
    dim  body as BlockStmt: st body, block_f()
    dim  temp as FunctionStmt: st temp, new FunctionStmt
    call temp.init(name, returntype, parameters, parametertypes, parammethods, body, vis)
    st function_, temp
    exit function
end function


private function isTypeName(byval tk as Token) as boolean
    if (tk.m_type = IDENTIFIER) then 
        if (tk.m_lexeme = "int") then 
            st isTypeName, True
            exit function
        else
            if (tk.m_lexeme = "string") then 
                st isTypeName, True
                exit function
            else
                if (tk.m_lexeme = "bool") then 
                    st isTypeName, True
                    exit function
                else
                    if (tk.m_lexeme = "long") then 
                        st isTypeName, True
                        exit function
                    else
                        st isTypeName, False
                        exit function
                    endif
                endif
            endif
        endif
    endif
end function


private function varDeclaration_f(byval vis as string) as VarStmt
    dim  vartype as Token: st vartype, consume(IDENTIFIER, "Expected type before variable name")
    dim  varName as Token: st varName, consume(IDENTIFIER, "Expected variable name after type")
    dim  arrcount as Expr: st arrcount, new Expr
    dim  hasarrcount as boolean: st hasarrcount, False
    if (match(LEFT_SQUARE)) then 
        nt st (arrcount, expression_())
        nt st (hasarrcount, True)
        call consume(RIGHT_SQUARE, "Expected closing ']' for array declaration")
    endif
    dim  initializer as expr: st initializer, new expr
    dim  addnew as boolean: st addnew, False
    dim  temp as VarStmt: st temp, new VarStmt
    if (match(EQUAL)) then 
        if (match(NEW_)) then 
            nt st (addnew, True)
        endif
        call temp.init(vis, vartype, varname, addnew, True, initializer, arrcount, hasarrcount)
    else
        call temp.init(vis, vartype, varname, addnew, False, initializer, arrcount, hasarrcount)
    endif
    call consume(SEMICOLON, "Expect ';' after variable declaration")
    st varDeclaration_f, temp
    exit function
end function


private function statement_f() as Stmt
    if (match(IF_)) then 
        st statement_f, ifStatement_f()
        exit function
    else
        if (match(PRINT_)) then 
            st statement_f, printStatement_f()
            exit function
        else
            if (match(LEFT_BRACE)) then 
                dim  block as BlockStmt: st block, new BlockStmt
                call block.init(block_f())
                st statement_f, block
                exit function
            else
                if (match(WHILE_)) then 
                    st statement_f, whileStatement_f()
                    exit function
                else
                    if (match(FOR_)) then 
                        st statement_f, forStatement_f()
                        exit function
                    else
                        if (match(RETURN_)) then 
                            st statement_f, returnStatement_f()
                            exit function
                        else
                            if (match(RAW_)) then 
                                st statement_f, codeStatement_f()
                                exit function
                            else
                                if (match(CONTINUE_, BREAK_)) then 
                                    st statement_f, controlStatement_f()
                                    exit function
                                endif
                            endif
                        endif
                    endif
                endif
            endif
        endif
    endif
    st statement_f, expressionStatement_f()
    exit function
end function


private function returnStatement_f() as ReturnStmt
    dim  keyword as Token: st keyword, previous()
    dim  value as Expr
    dim  temp as ReturnStmt: st temp, new ReturnStmt
    if (not check(SEMICOLON)) then 
        nt st (value, expression_())
        call temp.init(keyword, value, True)
    else
        dim  empty as LiteralExpr: st empty, new LiteralExpr
        call empty.init("null")
        call temp.init(keyword, empty, False)
    endif
    call consume(SEMICOLON, "Expected ';' after return")
    st returnStatement_f, temp
    exit function
end function


private function codeStatement_f() as CodeStmt
    dim  codeblock as Token: st codeblock, previous()
    dim  temp as CodeStmt: st temp, new CodeStmt
    call temp.init(codebloc.m_literal)
    st codeStatement_f, temp
    exit function
end function


private function forStatement_f() as Stmt
    call consume(LEFT_PAREN, "Expected '(' after 'for'")
    dim  isForEach as boolean: st isForEach, False
    dim  iteratingobj as Token
    dim  vardecin as VarStmt
    dim  exdecin as ExprStmt
    dim  varexprin as VariableExpr
    dim  hasDeclarator as boolean: st hasDeclarator, False
    dim  initializer as Stmt
    dim  iteratingobjtype as Token
    dim  iterable as Expr
    if (check(IDENTIFIER) and (peeknext().m_type_ = COLON or peeknextnext().m_type_ = COLON)) then 
        nt st (isForEach, True)
        if (peeknext.m_type_ = IDENTIFIER) then 
            nt st (hasDeclarator, True)
            nt st (iteratingobjtype, consume(IDENTIFIER, "Expected type name"))
        endif
        nt st (iteratingobj, consume(IDENTIFIER, "Expected Variable Name"))
        call consume(COLON, "Expected ':' in for clause")
        nt st (iterable, expression_())
        call consume(RIGHT_PAREN, "Expected ')' after 'for' clause")
    else
        dim  hasinitializer as boolean: st hasinitializer, True
        if (match(SEMICOLON)) then 
            nt st (hasinitializer, False)
        else
            if (isTypeName(peek())) then 
                nt st (initializer, varDeclaration_f())
            else
                nt st (initializer, expressionStatement_f())
            endif
        endif
        dim  condition as Expr
        dim  hascondition as boolean: st hascondition, False
        dim  increment as Expr
        dim  hasincrement as boolean: st hasincrement, False
        if (not check(SEMICOLON)) then 
            nt st (condition, expression_())
            nt st (hascondition, True)
        endif
        call consume(SEMICOLON, "Expect ';' after loop condition")
        if (not check(RIGHT_PAREN)) then 
            nt st (increment, expression_())
            nt st (hasincrement, True)
        endif
        call consume(RIGHT_PAREN, "Expected ')' after for clause")
    endif
    dim  iterator as Token
    dim  iteratortype as Token
    dim  hasIteratorDeclarator as boolean: st hasIteratorDeclarator, False
    if (isForEach) then 
        call consume(LEFT_PAREN, "Expected '(' for 'foreach' loop")
        if (peeknext().m_type_ = IDENTIFIER) then 
            nt st (hasIteratorDeclarator, True)
            nt st (iteratortype, consume(IDENTIFIER, "Expected iterator type"))
        endif
        nt st (iterator, consume(IDENTIFIER, "Expected variable name"))
        call consume(RIGHT_PAREN, "Expected ')' after for clause")
    endif
    dim  body as Stmt: st body, statement_f()
    dim  temp as BlockStmt: st temp, new BlockStmt
    dim  tempinner as ExprStmt: st tempinner, new ExprStmt
    dim  itx as Stmt
    dim  one as LiteralExpr: st one, new LiteralExpr
    dim  incr as ExprStmt: st incr, new ExprStmt
    call one.init(1)
    call incr.init(one)
    nt st (itx, incr)
    if (not isForEach) then 
        if (hasincrement) then 
            call incr.init(increment)
            nt st (itx, incr)
        endif
        if (not hascondition) then 
            dim  tmp as LiteralExpr: st tmp, new LiteralExpr
            call tmp.init(True)
            nt st (condition, tmp)
        endif
        dim  tempwhile as WhileStmt: st tempwhile, new WhileStmt
        call tempwhile.init(condition, body, itx)
        nt st (body, tempwhile)
        if (hasinitializer) then 
            dim  tempblock as BlockStmt: st tempblock, new BlockStmt
            call tempblock.init(array(initializer, body))
            st forStatement_f, tempblock
            exit function
        else
            st forStatement_f, body
            exit function
        endif
    else
        dim  tempforeach as ForEachStmt: st tempforeach, new ForEachStmt
        dim  iteratingobjdeclarator as VarStmt: st iteratingobjdeclarator, new VarStmt
        dim  iteratordeclarator as VarStmt: st iteratordeclarator, new VarStmt
        dim  e as Expr: st e, new Expr
        call iteratingobjdeclarator.init("", iteratingobjtype, iteratingobj, False, False, e, e, False)
        call iteratordeclarator.init("", iteratortype, iterator, False, False, e, e, False)
        call tempforeach.init(iteratingobjdeclarator, iteratingobj, iterable, iteratordeclarator, iterator, body, hasDeclarator, hasIteratorDeclarator)
        st forStatement_f, tempforeach
        exit function
    endif
end function


private function whileStatement_f() as WhileStmt
    call consume(LEFT_PAREN, "Expected '(' after 'while'")
    dim  condition as Expr: st condition, expression_()
    call consume(RIGHT_PAREN, "Expected ')' after while condition")
    dim  body as Stmt: st body, statement_f()
    dim  one as LiteralExpr: st one, new LiteralExpr
    dim  incr as ExprStmt: st incr, new ExprStmt
    call one.init(1)
    call incr.init(one)
    dim  tempwhile as WhileStmt: st tempwhile, new WhileStmt
    call tempwhile.init(condition, body, incr)
    st whileStatement_f, tempwhile
    exit function
end function


private function ifStatement_f() as IfStmt
    call consume(LEFT_PAREN, "Expected '(' after 'if'.")
    dim  condition as Expr: st condition, expression_()
    call consume(RIGHT_PAREN, "Expected ')' after condition")
    dim  thenBranch as Stmt: st thenBranch, statement_f()
    dim  elseBranch as Stmt
    dim  temp as IfStmt: st temp, new IfStmt
    if (match(ELSE_)) then 
        nt st (elseBranch, statement_f())
    endif
    call temp.init(condition, thenBranch, elseBranch)
    st ifStatement_f, temp
    exit function
end function


private function printStatement_f() as PrintStmt
    dim  val as Expr: st val, expression_()
    call consume(SEMICOLON, "Expect ';' after value")
    dim  temp as PrintStmt: st temp, new PrintStmt
    call temp.init(val)
    st printStatement_f, temp
    exit function
end function


private function expressionStatement_f() as ExprStmt
    dim  temp as ExprStmt: st temp, new ExprStmt
    call temp.init(expression_())
    call consume(SEMICOLON, "Expect ';' after value")
    st expressionStatement_f, temp
    exit function
end function


private function block_f() as variant
    dim  statements() as variant
private function funcl380c0() as Boolean
    funcl380c0 = false
    if (not check(RIGHT_BRACE)) then
        if (not isAtEnd()) then
            funcl380c0 = true
        endif
    endif
end function
    while (funcl380c0())
        dim  temp as Stmt: st temp, declaration_f()
        call arad(statements, temp)
        looplabel5bef:
        nt 1
    wend
    looplabel5aft:
    call consume(RIGHT_BRACE, "Expect '}' after block")
    st block_f, statements
    exit function
end function


private function expression_() as Expr
    st expression_, assignment_()
    exit function
end function


private function assignment_() as Expr
    dim  expr_ as Expr: st expr_, br_and_f()
    dim  eqs as Token
    dim  value as Expr
    dim  temp as AssignExpr: st temp, new AssignExpr
    dim  tempbinary as BinaryExpr: st tempbinary, new BinaryExpr
    dim  ops as Token: st ops, new Token
    dim  one as LiteralExpr: st one, new LiteralExpr
    dim  tempunary as UnaryExpr: st tempunary, new UnaryExpr
    dim  addnew as boolean: st addnew, False
    if (match(EQUAL)) then 
        if (match(NEW_)) then 
            nt st (addnew, True)
        endif
        nt st (eqs, previous())
        nt st (value, assignment_())
        if (TypeName(expr_) = "VariableExpr") then 
            call temp.init(expr_, value, addnew)
            nt st (expr_, temp)
        else
            if (TypeName(expr_) = "CallExpr") then 
                call temp.init(expr_, value, addnew)
                nt st (expr_, temp)
            else
                if (TypeName(expr_) = "AccessorExpr") then 
                    call temp.init(expr_, value, addnew)
                    nt st (expr_, temp)
                endif
            endif
        endif
    else
        if (match(PLUS_EQUAL)) then 
            call ops.init(PLUS, "+", "null", previous().m_line)
            nt st (value, assignment_())
            call tempbinary.init(expr_.copy(), ops, value)
            call temp.init(expr_, tempbinary, addnew)
            nt st (expr_, temp)
        else
            if (match(MINUS_EQUAL)) then 
                call ops.init(MINUS, "-", "null", previous().m_line)
                nt st (value, assignment_())
                call tempbinary.init(expr_.copy(), ops, value)
                call temp.init(expr_, tempbinary, addnew)
                nt st (expr_, temp)
            else
                if (match(STAR_EQUAL)) then 
                    call ops.init(STAR, "*", "null", previous().m_line)
                    nt st (value, assignment_())
                    call tempbinary.init(expr_.copy(), ops, value)
                    call temp.init(expr_, tempbinary, addnew)
                    nt st (expr_, temp)
                else
                    if (match(SLASH_EQUAL)) then 
                        call ops.init(SLASH, "/", "null", previous().m_line)
                        nt st (value, assignment_())
                        call tempbinary.init(expr_.copy(), ops, value)
                        call temp.init(expr_, tempbinary, addnew)
                        nt st (expr_, temp)
                    else
                        if (match(PLUS_PLUS)) then 
                            call ops.init(PLUS, "+", "null", previous().m_line)
                            call one.init(1)
                            call tempbinary.init(expr_.copy(), ops, one)
                            call temp.init(expr_, tempbinary, addnew)
                            nt st (expr_, temp)
                        else
                            if (match(MINUS_MINUS)) then 
                                call ops.init(MINUS, "-", "null", previous().m_line)
                                call tempunary.init(ops, expr_.copy())
                                nt st (expr_, tempunary)
                            endif
                        endif
                    endif
                endif
            endif
        endif
    endif
    st assignment_, expr_
    exit function
end function


private function br_and_f() as Expr
    dim  expr_ as Expr: st expr_, or_f()
    while (match(BRANCHING_AND_))
        dim  operator as Token: st operator, previous()
        dim  right as Expr: st right, or_f()
        dim  temp as LogicalExpr: st temp, new LogicalExpr
        call temp.init(expr_.copy(), operator, right)
        nt st (expr_, temp)
        looplabel6bef:
        nt 1
    wend
    looplabel6aft:
    st br_and_f, expr_
    exit function
end function


private function or_f() as Expr
    dim  expr_ as Expr: st expr_, and_f()
    while (match(OR_))
        dim  operator as Token: st operator, previous()
        dim  right as Expr: st right, and_f()
        dim  temp as BinaryExpr: st temp, new BinaryExpr
        dim  prevExpr as Expr: st prevExpr, expr_.copy()
        call temp.init(prevExpr, operator, right)
        nt st (expr_, temp)
        looplabel7bef:
        nt 1
    wend
    looplabel7aft:
    st or_f, expr_
    exit function
end function


private function and_f() as Expr
    dim  expr_ as Expr: st expr_, equality_()
    while (match(AND_))
        dim  operator as Token: st operator, previous()
        dim  right as Expr: st right, equality_()
        dim  temp as BinaryExpr: st temp, new BinaryExpr
        dim  prevExpr as Expr: st prevExpr, expr_.copy()
        call temp.init(prevExpr, operator, right)
        nt st (expr_, temp)
        looplabel8bef:
        nt 1
    wend
    looplabel8aft:
    st and_f, expr_
    exit function
end function


private function equality_() as Expr
    dim  expr_ as Expr: st expr_, comparison_()
    while (match(BANG_EQUAL, EQUAL_EQUAL))
        dim  operator as Token: st operator, previous()
        dim  right as Expr: st right, comparison_()
        dim  temp as BinaryExpr: st temp, new BinaryExpr
        call temp.init(expr_.copy(), operator, right)
        nt st (expr_, temp)
        looplabel9bef:
        nt 1
    wend
    looplabel9aft:
    st equality_, expr_
    exit function
end function


private function comparison_() as Expr
    dim  expr_ as Expr: st expr_, term_()
    while (match(GREATER, GREATER_EQUAL, LESS, LESS_EQUAL))
        dim  operator as Token: st operator, previous()
        dim  right as Expr: st right, term_()
        dim  temp as BinaryExpr: st temp, new BinaryExpr
        call temp.init(expr_.copy(), operator, right)
        nt st (expr_, temp)
        looplabel10bef:
        nt 1
    wend
    looplabel10aft:
    st comparison_, expr_
    exit function
end function


private function term_() as Expr
    dim  expr_ as Expr: st expr_, factor_()
    while (match(MINUS, PLUS, AMPERSAND, MOD_))
        dim  operator as Token: st operator, previous()
        dim  left as Expr: st left, expr_.copy()
        dim  right as Expr: st right, factor_()
        dim  temp as BinaryExpr: st temp, new BinaryExpr
        call temp.init(left, operator, right)
        nt st (expr_, temp)
        looplabel11bef:
        nt 1
    wend
    looplabel11aft:
    st term_, expr_
    exit function
end function


private function factor_() as Expr
    dim  expr_ as Expr: st expr_, unary_()
    while (match(SLASH, STAR))
        dim  operator as Token: st operator, previous()
        dim  right as Expr: st right, unary_()
        dim  left as Expr: st left, expr_.copy()
        dim  temp as BinaryExpr: st temp, new BinaryExpr
        call temp.init(left, operator, right)
        nt st (expr_, temp)
        looplabel12bef:
        nt 1
    wend
    looplabel12aft:
    st factor_, expr_
    exit function
end function


private function unary_() as Expr
    if (match(BANG, MINUS)) then 
        dim  operator as Token: st operator, previous()
        dim  right as Expr: st right, unary_()
        dim  temp as UnaryExpr: st temp, new UnaryExpr
        call temp.init(operator, right)
        st unary_, temp
        exit function
    else
        st unary_, call_()
        exit function
    endif
end function


private function call_() as Expr
    dim  expr_ as Expr: st expr_, primary_()
    while (True)
        if (match(LEFT_PAREN)) then 
            nt st (expr_, finishCall(expr_.copy()))
        else
            if (match(DOT)) then 
                nt st (expr_, finishDot(expr_.copy()))
            else
                if (match(LEFT_SQUARE)) then 
                    nt st (expr_, finishAccessor(expr_.copy()))
                else
                    goto looplabel13aft:
                endif
            endif
        endif
        looplabel13bef:
        nt 1
    wend
    looplabel13aft:
    st call_, expr_
    exit function
end function


private function finishAccessor(byval accessee as Expr) as AccessorExpr
    dim  arguments() as variant
    dim  colonCounter as long: st colonCounter, 0
    dim  commaPresent as boolean: st commaPresent, False
    if (not check(RIGHT_SQUARE)) then 
        call arad(arguments, expression_())
        while (not check(RIGHT_SQUARE) and (check(COLON) or check(COMMA)) and (not commaPresent and colonCounter < 2))
            if (check(COLON)) then 
                call consume(COLON, "")
                nt st (colonCounter, colonCounter + 1)
            else
                if (check(COMMA)) then 
                    call consume(COMMA, "")
                    nt st (commaPresent, True)
                endif
            endif
            call arad(arguments, expression_())
            looplabel14bef:
            nt 1
        wend
        looplabel14aft:
    endif
    dim  paren as Token: st paren, consume(RIGHT_SQUARE, "Expect ']' after indices")
    dim  valMode as string
    if (commaPresent) then 
        nt st (valMode, "val")
    else
        if (colonCounter = 0) then 
            nt st (valMode, "access")
        else
            if (colonCounter = 1) then 
                nt st (valMode, "slice")
                dim  ee as LiteralExpr: st ee, new LiteralExpr
                call ee.init(1)
                call arad(arguments, ee)
            else
                if (colonCounter = 2) then 
                    nt st (valMode, "slice")
                endif
            endif
        endif
    endif
    dim  temp as AccessorExpr: st temp, new AccessorExpr
    call temp.init(accessee, paren, arguments, valMode)
    st finishAccessor, temp
    exit function
end function


private function finishCall(byval callee as Expr) as CallExpr
    dim  arguments() as variant
    if (not check(RIGHT_PAREN)) then 
        if (al(arguments) > 254) then 
            call err(peek(), "Can't have more than 255 arguments")
        endif
        call arad(arguments, callArgument_f())
        while (match(COMMA))
            if (al(arguments) > 254) then 
                call err(peek(), "Can't have more than 255 arguments")
            endif
            call arad(arguments, callArgument_f())
            looplabel15bef:
            nt 1
        wend
        looplabel15aft:
    endif
    dim  paren as Token: st paren, consume(RIGHT_PAREN, "Expect ')' after arguments")
    dim  temp as CallExpr: st temp, new CallExpr
    call temp.init(callee, paren, arguments)
    st finishCall, temp
    exit function
end function


private function callArgument_f() as Expr
    if (match(LEFT_SQUARE)) then 
        dim  myfunctExpr as FunctExpr: st myfunctExpr, new FunctExpr
        dim  inputs() as variant
        call arad(inputs, consume(IDENTIFIER, "Expected identifier"))
        while (match(COMMA))
            call arad(inputs, consume(IDENTIFIER, "Expected identifier"))
            looplabel16bef:
            nt 1
        wend
        looplabel16aft:
        call consume(RIGHT_SQUARE, "Expected ']'")
        dim  arrowkey as Token: st arrowkey, consume(ARROW_, "Expected arrow for arrow sign")
        dim  functname as string: st functname, "functl" & previous().m_line & "c" & functcounter
        nt st (functcounter, functcounter + 1)
        call myfunctExpr.init(functname)
        dim  halfbody as Stmt
        if (match(LEFT_BRACE)) then 
            nt st (halfbody, block_f())
        else
            dim  tempexpr as Expr: st tempexpr, expression_()
            dim  tempbodyreturn as ReturnStmt: st tempbodyreturn, new ReturnStmt
            call tempbodyreturn.init(arrowkey, tempexpr.copy(), True)
            nt st (halfbody, tempbodyreturn)
        endif
        dim  body() as variant
        dim  param as Token
        dim  i as long
        for i = 0 to al(inputs) - 1
            st param, inputs(i)
            dim  tassignexpr as AssignExpr: st tassignexpr, new AssignExpr
            dim  inputname as VariableExpr: st inputname, new VariableExpr
            dim  arraccessor as AccessorExpr: st arraccessor, new AccessorExpr
            dim  arrname as VariableExpr: st arrname, new VariableExpr
            dim  arrnametoken as Token: st arrnametoken, new Token
            dim  fakeBracket as Token: st fakeBracket, new Token
            dim  numberr as LiteralExpr: st numberr, new LiteralExpr
            dim  wraps as ExprStmt: st wraps, new ExprStmt
            call arrnametoken.init(IDENTIFIER, "arr", "arr", arrowkey.m_line)
            call fakeBracket.init(LEFT_SQUARE, "[fake", "[fake", arrowkey.m_line)
            call numberr.init(i)
            call arrname.init(arrnametoken)
            call inputname.init(param)
            call arraccessor.init(arrname, fakeBracket, array(numberr), "access", False)
            call tassignexpr.init(inputname, arraccessor, False)
            call wraps.init(tassignexpr)
            call arad(body, wraps)
            looplabel17bef:
        next i
        looplabel17aft:
        call arad(body, halfbody)
        dim  functClass as ClassStmt: st functClass, new ClassStmt
        dim  fstmt as FunctionStmt: st fstmt, new FunctionStmt
        call fstmt.init(mkident("funct_run", arrowkey.m_line), mkident("variant", arrowkey.m_line))
        call functClass.init(mkident(functname, arrowkey.m_line), array(fstmt), array(mkident("funct", arrowkey.m_line)), array())
        call arad(statements, functClass)
        st callArgument_f, myfunctExpr
        exit function
    else
        st callArgument_f, expression_()
        exit function
    endif
end function


private function mkident(byval name as string, byval linenum as long) as Token
    dim  t as Token: st t, new Token
    call t.init(IDENTIFIER, name, name, linenum)
    st mkident, t
    exit function
end function


private function finishDot(byval accessee as Expr) as DotExpr
    dim  rhs as Expr: st rhs, call_()
    dim  temp as DotExpr: st temp, new DotExpr
    call temp.init(accessee, rhs)
    st finishDot, temp
    exit function
end function


private function primary_() as Expr
    dim  temp as LiteralExpr: st temp, new LiteralExpr
    if (match(FALSE_)) then 
        call temp.init(False)
        st primary_, temp
        exit function
    else
        if (match(TRUE_)) then 
            call temp.init(True)
            st primary_, temp
            exit function
        else
            if (match(NULL_)) then 
                call temp.init("null")
                st primary_, temp
                exit function
            else
                if (match(NUMBER_, STRING_)) then 
                    call temp.init(previous().m_literal)
                    st primary_, temp
                    exit function
                else
                    if (match(LEFT_PAREN)) then 
                        dim  expr_ as Expr: st expr_, expression_()
                        call consume(RIGHT_PAREN, "Expect ')' after expression.")
                        dim  tempp as GroupingExpr: st tempp, new GroupingExpr
                        call tempp.init(expr_)
                        st primary_, tempp
                        exit function
                    else
                        if (match(IDENTIFIER)) then 
                            dim  temppp as VariableExpr: st temppp, new VariableExpr
                            call temppp.init(previous())
                            st primary_, temppp
                            exit function
                        endif
                    endif
                endif
            endif
        endif
    endif
    call err(peek(), "Unexpected Token")
end function


private function match( paramarray arr() as variant) as boolean
    dim  typ as tokentype
    dim  i as long
    for i = 0 to al(arr) - 1
        st typ, arr(i)
        if (check(ttyp)) then 
            call advance()
            st match, True
            exit function
        endif
        looplabel18bef:
    next i
    looplabel18aft:
    st match, False
    exit function
end function


private function check(byval typ as tokentype) as boolean
    if (isAtEnd()) then 
        st check, False
        exit function
    endif
    st check, peek().m_type_ = typ
    exit function
end function


private function advance() as Token
    if (not isAtEnd()) then 
        nt st (current, current + 1)
    endif
    st advance, previous()
    exit function
end function


private function isAtEnd() as boolean
    st isAtEnd, peek().m_type_ = EOF_
    exit function
end function


private function peek() as Token
    st peek, tokens(current)
    exit function
end function


private function peeknext() as Token
    if (current + 1 >= al(tokens)) then 
        st peeknext, peek()
        exit function
    endif
    st peeknext, tokens(current + 1)
    exit function
end function


private function peeknextnext() as Token
    if (current + 2 >= al(tokens)) then 
        dim  a as Token: st a, new Token
        call a.init(EOF_, "", "", 999)
        st peeknextnext, peek()
        exit function
    endif
    st peeknextnext, tokens(current + 2)
    exit function
end function


private function previous() as Token
    st previous, tokens(current - 1)
    exit function
end function


private function consume(byval typ as tokentype, byval message as string) as Token
    if (check(typ)) then 
        dim  t as Token: st t, peek()
        call advance()
        st consume, t
        exit function
    endif
    call err(peek(), message)
end function


private function err(byval tk as Token, byval message as string) 
    call m_h.err(tk, message)
end function



