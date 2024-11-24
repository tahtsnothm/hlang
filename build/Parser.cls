VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private tokens As Variant
Private current As Long
Private m_h As H

Public Function init(tokens_ As Variant, hh As H)
tokens = tokens_
current = 0
Set m_h = hh
End Function

Public Function parse()
    Dim statements() As Variant
    While (Not isAtEnd())
        arad statements, declaration_f()
    Wend
    st parse, statements
End Function
Private Function declaration_f()
    If ((peek().m_type_ = IDENTIFIER) And (peeknext().m_type_ = IDENTIFIER)) Then
        st declaration_f, varDeclaration_f()
        Exit Function
    ElseIf (match(FUNC_)) Then
            st declaration_f, function_("function")
            Exit Function
     ElseIf (match(CLASS_)) Then
        st declaration_f, classDeclaration_f()
        Exit Function
    End If
    
    st declaration_f, statement_f()
    
End Function
Private Function classDeclaration_f()
 Dim name As Token
st name, consume(IDENTIFIER, "Expect class name.")
consume LEFT_BRACE, "Expect '{' before class body"
Dim methods() As Variant
While (Not check(RIGHT_BRACE) And Not isAtEnd())
    arad methods, function_("method")
Wend

consume RIGHT_BRACE, "Expect '}' after class body."

Dim temp As New ClassStmt
temp.Stmt_init name, methods
        
End Function

Private Function function_(kind As String)

    Dim returntype As Token
    st returntype, consume(IDENTIFIER, "Expected type before function name")
    Dim name As Token
    st name, consume(IDENTIFIER, "Expected " & kind & " name.")
    
    consume LEFT_PAREN, "Expected '(' after " & kind & " name."
    Dim parameters() As Variant
    Dim parammethods() As Variant
    Dim parametertypes() As Variant
    If (Not check(RIGHT_PAREN)) Then
        Do
            If (al(parameters) > 254) Then
                err peek(), "Can't have more than 255 parameters"
            End If
            arad parametertypes, consume(IDENTIFIER, "Expected parameter type.")
            If (match(AMPERSAND)) Then
                arad parammethods, True
            Else
                arad parammethods, False
            End If
            arad parameters, consume(IDENTIFIER, "Expected parameter name.")
        Loop While (match(COMMA))
        
        
    End If
    consume RIGHT_PAREN, "Expected ')' after parameters"
    
    consume LEFT_BRACE, "Expect '{' before " & kind & " body"
    Dim body As Variant
    st body, block_f()
    Dim temp As New FunctionStmt
    temp.Stmt_init name, returntype, parameters, parametertypes, parammethods, body
    st function_, temp
    
    
End Function
Private Function isTypeName(tk As Token)

    If tk.m_type_ = IDENTIFIER Then
        Select Case tk.m_lexeme
        Case "int":
            isTypeName = True
        Case "string":
            isTypeName = True
        Case "bool":
            isTypeName = True
        Case "long":
            isTypeName = True
        
        End Select
        
    Else
        isTypeName = False
    End If
    
End Function
Private Function varDeclaration_f()
    Dim varType As Token
    st varType, consume(IDENTIFIER, "Expected type before variable name")
    Dim varName As Token

    st varName, consume(IDENTIFIER, "Expected variable name")
    Dim initializer As Expr
    
    
    Dim temp As New VarStmt
    If (match(EQUAL)) Then
        st initializer, expression_()
        temp.Stmt_init varType, varName, initializer
    Else
        temp.Stmt_init varType, varName 'do we need a null node hmm (no thats in the codeGen, for the init function to handle the boolean)
    End If
    consume SEMICOLON, "Expect ';' after variable declaration"
    st varDeclaration_f, temp
End Function
Private Function statement_f()
    If (match(IF_)) Then
        st statement_f, ifStatement_f()
        Exit Function
    End If
    If (match(PRINT_)) Then
        st statement_f, printStatement_f()
        Exit Function
    End If
    If (match(LEFT_BRACE)) Then
    
        Dim blocknode As New BlockStmt
        blocknode.Stmt_init block_f()
        st statement_f, blocknode
        Exit Function
    End If

    If (match(WHILE_)) Then
        st statement_f, whileStatement_f()
        Exit Function
    End If
    
    If (match(FOR_)) Then
        st statement_f, forStatement_f()
        Exit Function
    End If
    If (match(RETURN_)) Then
        st statement_f, returnStatement_f()
        Exit Function
    End If
    If (match(RAW_)) Then
        st statement_f, codeStatement_f()
        Exit Function
    End If
    st statement_f, expressionStatement_f()
End Function

Private Function returnStatement_f()
Dim keyword As Token: st keyword, previous()
Dim value As Expr
Dim temp As New ReturnStmt
If (Not check(SEMICOLON)) Then
    st value, expression_()
    temp.Stmt_init keyword, value, True
Else
    Dim faketemp As New LiteralExpr
    faketemp.Expr_init "null"
    temp.Stmt_init keyword, faketemp, False
End If
consume SEMICOLON, "Expected ';' after return"
st returnStatement_f, temp
End Function
Private Function codeStatement_f()
Dim codeblock As Token
st codeblock, previous()

Dim temp As New CodeStmt
temp.Stmt_init codeblock.m_literal
st codeStatement_f, temp
End Function
Private Function forStatement_f()
    consume LEFT_PAREN, "Expected '(' after 'for'"
    
    Dim initializer As Stmt
    Dim hasinitializer As Boolean: hasinitializer = True
    If (match(SEMICOLON)) Then
        hasinitializer = False
    ElseIf isTypeName(peek()) Then
        st initializer, varDeclaration_f()
    Else
        st initializer, expressionStatement_f()
    End If
    
    Dim condition As Expr
    Dim hascondition As Boolean: hascondition = False
    If (Not check(SEMICOLON)) Then
        st condition, expression_()
        hascondition = True
    End If
    consume SEMICOLON, "Expect ';' after loop condition"
    
    Dim increment As Expr
    Dim hasincrement As Boolean: hasincrement = False
    If (Not check(RIGHT_PAREN)) Then
        st increment, expression_()
        hasincrement = True
    End If
    consume RIGHT_PAREN, "Expected ')' after 'for' clauses"
    
    
    Dim body As Stmt
    st body, statement_f()
    
    If (hasincrement) Then
        Dim temp As New BlockStmt
        Dim tempinner As New ExprStmt
        tempinner.Stmt_init increment
        temp.Stmt_init Array(body, tempinner)
        st body, temp
    End If
    
    If (Not hascondition) Then
        Dim tmp As New LiteralExpr
        tmp.Expr_init True
        st condition, tmp
        
    End If
    Dim tempwhile As New WhileStmt
    tempwhile.Stmt_init condition, body
    st body, tempwhile
    
    If (hasinitializer) Then
        Dim temptemp As New BlockStmt
        temptemp.Stmt_init Array(initializer, body)
        st forStatement_f, temptemp
    Else
        st forStatement_f, body
    End If
    
End Function

Private Function whileStatement_f()
    consume LEFT_PAREN, "Expected '(' after 'while'"
    Dim condition As New Expr
    st condition, expression_()
    
    consume RIGHT_PAREN, "Expected ')' after while condition"
    Dim body As New Expr
    st body, expression_()
    Dim temp As New WhileStmt
    temp.Stmt_init condition, body
    
    st whileStatement_f, temp
End Function

Private Function ifStatement_f()
    consume LEFT_PAREN, "Expected '(' after 'if'."
    Dim condition As New Expr
    st condition, expression_f()
    consume RIGHT_PAREN, "Expected ')' after condition"
    
    Dim thenBranch As New Stmt
    Dim elseBranch As Stmt
    st thenBranch, statement_f()

    Dim temp As New IfStmt
    If (match(ELSE_)) Then
        st elseBranch, statement_f()
        temp.Stmt_init condition, thenBranch, elseBranch
    Else
        temp.Stmt_init condition, thenBranch
    End If
    
    st ifStatement_f, temp

End Function


Private Function printStatement_f()
    Dim val As Expr
    st val, expression_()
    consume SEMICOLON, "Expect ';' after value"
    Dim temp As New PrintStmt
    temp.Stmt_init val
    st printStatement_f, temp
End Function

Private Function expressionStatement_f()
    Dim expr_ As New Expr
    st expr_, expression_()
    consume SEMICOLON, "Expect ';' after value"
    Dim temp As New ExprStmt
    temp.Stmt_init expr_
    st expressionStatement_f, temp
End Function

Private Function block_f()
    Dim statements() As Variant
    While (Not check(RIGHT_BRACE) And Not isAtEnd())
        Dim temp As New Stmt
        st temp, declaration_f()
        arad statements, temp
    Wend
    consume RIGHT_BRACE, "Expect '}' after block"
    st block_f, statements
    
End Function

Private Function expression_() As Expr
    Set expression_ = assignment_()
End Function
Private Function expression_f() As Expr
    Set expression_f = assignment_()
End Function

Private Function assignment_()
    Dim expr_ As New Expr
    st expr_, br_and_f()
    Dim eqs As New Token
    Dim value As Expr
    Dim temp As New AssignExpr
    Dim tempbinary As New BinaryExpr
    Dim ops As New Token
    If match(EQUAL) Then
        
        st eqs, previous()
        st value, assignment_()
        If (TypeName(expr_) = "VariableExpr") Then
            temp.Expr_init expr_, value
            st expr_, temp
        ElseIf (TypeName(expr_) = "CallExpr") Then
            temp.Expr_init expr_, value
            st expr_, temp
        End If
        'raise error here
    ElseIf match(PLUS_EQUAL) Then
        ops.init PLUS, "+", Null, previous.m_line
        st value, assignment_()
        tempbinary.Expr_init expr_.copy(), ops, value 'minor bug if expr_ (LHS of assignment) is not pure
        temp.Expr_init expr_, tempbinary
        st expr_, temp
    ElseIf match(MINUS_EQUAL) Then
        ops.init MINUS, "-", Null, previous.m_line
        st value, assignment_()
        tempbinary.Expr_init expr_.copy(), ops, value 'minor bug if expr_ (LHS of assignment) is not pure
        temp.Expr_init expr_, tempbinary
        st expr_, temp
    ElseIf match(STAR_EQUAL) Then
        ops.init STAR, "*", Null, previous.m_line
        st value, assignment_()
        tempbinary.Expr_init expr_.copy(), ops, value 'minor bug if expr_ (LHS of assignment) is not pure
        temp.Expr_init expr_, tempbinary
        st expr_, temp
    ElseIf match(SLASH_EQUAL) Then
        ops.init SLASH, "/", Null, previous.m_line
        st value, assignment_()
        tempbinary.Expr_init expr_.copy(), ops, value 'minor bug if expr_ (LHS of assignment) is not pure
        temp.Expr_init expr_, tempbinary
        st expr_, temp
    End If
    
    st assignment_, expr_
End Function

Private Function br_and_f()
'branching and
    Dim expr_ As New Expr
    st expr_, or_f()
    While (match(BRANCHING_AND_))
        Dim operator As New Token
        st operator, previous()
        
        Dim right As New Expr
        st right, or_f()
        
        Dim temp As New LogicalExpr
        temp.Expr_init expr_.copy(), operator, right
        st expr_, temp
    Wend
    st br_and_f, expr_
End Function

Private Function or_f()
    Dim expr_ As New Expr
    st expr_, and_f()
    
    While (match(OR_))
        Dim operator As New Token
        st operator, previous()
        
        Dim right As New Expr
        st right, and_f()
        
        Dim temp As New BinaryExpr
        Dim prevExpr As New Expr
        st prevExpr, expr_
        temp.Expr_init prevExpr.copy(), operator, right
        st expr_, temp
    Wend
    st or_f, expr_
End Function

Private Function and_f()
     Dim expr_ As New Expr
    st expr_, equality_()
    
    While (match(AND_))
        Dim operator As New Token
        st operator, previous()
        
        Dim right As New Expr
        st right, equality_()
        
        Dim temp As New BinaryExpr
        Dim prevExpr As New Expr
        st prevExpr, expr_
        temp.Expr_init prevExpr.copy(), operator, right
        st expr_, temp
    Wend
    
    st and_f, expr_
End Function

Private Function equality_()
    Dim expr_ As Expr
    Set expr_ = comparison_()
    
    While (match(BANG_EQUAL, EQUAL_EQUAL))
        Dim operator As Token
        Set operator = previous()
        Dim right As Expr
        Set right = comparison_()
        Dim temp As New BinaryExpr
        temp.Expr_init expr_.copy(), operator, right
        Set expr_ = temp
    Wend
    
    Set equality_ = expr_
End Function

Private Function comparison_()
    Dim expr_ As Expr
    Set expr_ = term_()
    
    While (match(GREATER, GREATER_EQUAL, LESS, LESS_EQUAL))
        Dim operator As Token
        Set operator = previous()
        Dim right As Expr
        Set right = term_()
        Dim temp As New BinaryExpr
        temp.Expr_init expr_.copy(), operator, right
        Set expr_ = temp
    Wend
    
    Set comparison_ = expr_
End Function

Private Function term_()
    Dim expr_ As Expr
    Set expr_ = factor_()
    While (match(MINUS, PLUS, AMPERSAND))
        Dim operator As Token
        Set operator = previous()
        Dim left As Expr
        st left, expr_.copy()
        
        Dim right As Expr
        st right, factor_()
        'pt expr_.accept(New AstPrinter), right.accept(New AstPrinter)
        
        
        Dim temp As New BinaryExpr
        'pt TypeName(left), TypeName(operator), TypeName(right)
        temp.Expr_init left, operator, right
        
        Dim tempcast As Expr
        st tempcast, temp
        'pt tempcast.accept(New AstPrinter)
        st expr_, temp
    Wend
    
    Set term_ = expr_
End Function

Private Function factor_()
    Dim expr_ As Expr
    Set expr_ = unary_()
    While (match(SLASH, STAR))
        Dim operator As Token
        Set operator = previous()
        Dim right As New Expr
        Set right = unary_()
        Dim temp As New BinaryExpr
        Dim left As New Expr
        st left, expr_.copy()
        temp.Expr_init left, operator, right
        Set expr_ = temp
    Wend
    
    Set factor_ = expr_
End Function
Private Function unary_()
    Dim expr_ As Expr
    If (match(BANG, MINUS)) Then
        Dim operator As Token
        Set operator = previous()
        Dim right As Expr
        Set right = unary_()
        Dim temp As New UnaryExpr
        temp.Expr_init operator, right
        Set expr_ = temp
    Else
        Set expr_ = call_()
    End If
    
    Set unary_ = expr_
End Function

Private Function call_()
Dim expr_ As Expr
st expr_, primary_()

Dim exitwhile As Boolean: exitwhile = False
While Not exitwhile
    If (match(LEFT_PAREN)) Then
        Dim temp As Expr
        st temp, finishCall(expr_.copy())
        st expr_, temp
    Else
        exitwhile = True
    End If
Wend
st call_, expr_
End Function
Private Function finishCall(callee As Expr)
    Dim arguments() As Variant
    If (Not check(RIGHT_PAREN)) Then
        Do
            If (al(arguments) > 254) Then
                err peek(), "Can't have more than 255 arguments"
            End If
            arad arguments, expression_()
        Loop While (match(COMMA))
    End If
    
    Dim paren As Token
    st paren, consume(RIGHT_PAREN, "Expect ')' after arguments")
    
    
    Dim temp As New CallExpr
    temp.Expr_init callee, paren, arguments
    
    st finishCall, temp
    
End Function
Private Function primary_()
Dim temp As New LiteralExpr
If (match(FALSE_)) Then
    
    temp.Expr_init False
    Set primary_ = temp
    Exit Function
ElseIf (match(TRUE_)) Then

    temp.Expr_init True
    Set primary_ = temp
    Exit Function
ElseIf (match(NULL_)) Then
    temp.Expr_init "null"
    Set primary_ = temp
    Exit Function
ElseIf (match(NUMBER_, STRING_)) Then

    temp.Expr_init previous().m_literal
    Set primary_ = temp
    Exit Function
ElseIf match(LEFT_PAREN) Then
    Dim expr_ As Expr
    Set expr_ = expression_()
    consume RIGHT_PAREN, "Expect ')' after expression."
    Dim tempp As New GroupingExpr
    tempp.Expr_init expr_
    Set primary_ = tempp
    Exit Function
ElseIf match(IDENTIFIER) Then
    Dim temppp As New VariableExpr
    temppp.Expr_init previous()
    st primary_, temppp
    Exit Function
End If

    err peek(), "Unexpected Token"
End Function

Private Function match(ParamArray arr() As Variant) As Boolean
    Dim typ As Variant
    For Each typ In arr
        Dim ttyp As tokentype
        ttyp = typ
        If (check(ttyp)) Then
            advance
            match = True
            Exit Function
        End If
    Next typ
    match = False
End Function
Private Function check(typ As tokentype) As Boolean
If (isAtEnd()) Then
    check = False
    Exit Function
End If
check = peek().m_type_ = typ
End Function

Private Function advance() As Token
    If (Not isAtEnd()) Then
     current = current + 1
    End If
    
    Set advance = previous()
End Function

Private Function isAtEnd() As Boolean
isAtEnd = peek().m_type_ = EOF_
End Function

Private Function peek() As Token
    Set peek = tokens(current)
End Function

Private Function peeknext() As Token
If (current + 1 >= al(tokens)) Then
    Set peeknext = peek() 'guaranteed eof
    Exit Function
End If
st peeknext, tokens(current + 1)
End Function

Private Function previous() As Token
    Set previous = tokens(current - 1)
End Function

Private Function consume(typ As tokentype, message As String)
    If (check(typ)) Then
        st consume, peek()
        advance
        Exit Function
    End If
    
    err peek(), message
End Function
Private Function err(tk As Token, message As String)
    m_h.err tk, message
    
End Function
