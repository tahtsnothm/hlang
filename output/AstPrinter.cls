VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AstPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
implements ExprVisitor
implements StmtVisitor
public function prt(byval statements as variant) 
    dim  statement as Stmt
    dim  i as long
    for i = 0 to al(statements) - 1
        st statement, statements(i)
        dim  s as string: st s, cstr(statement.accept(me))
        call pt(s)
        looplabel0bef:
    next i
    looplabel0aft:
end function


private function parenthesize( paramarray arr() as variant) as string
    dim  s as string: st s, "(" & arr(0)
    dim  expression as Expr
    dim  i as long: st i, 1
    while (i < al(arr))
        nt st (expression, arr(i))
        nt st (s, s + " ")
        nt st (s, s + cstr(expression.accept(me)))
        looplabel1bef:
        nt st (i, i + 1)
    wend
    looplabel1aft:
    nt st (s, s + ")")
    st parenthesize, s
    exit function
end function


private function parenthesizestatements( paramarray arr() as variant) as string
    dim  s as string: st s, "(" + arr(0)
    dim  statement as Stmt
    dim  i as long: st i, 1
    while (i < al(arr))
        nt st (statement, arr(i))
        nt st (s, s + " ")
        nt st (s, s + cstr(statement.accept(me)))
        looplabel2bef:
        nt st (i, i + 1)
    wend
    looplabel2aft:
    nt st (s, s + ")")
    st parenthesizestatements, s
    exit function
end function


private function parenthesizelots( paramarray arr() as variant) as string
    dim  s as string: st s, "(" + arr(0)
    dim  statement as Stmt
    dim  i as long: st i, 0
    while (i < al(arr))
        nt st (statement, arr(1)(i))
        nt st (s, s + cstr(statement.accept(me)))
        nt st (s, s + " ")
        looplabel3bef:
        nt st (i, i + 1)
    wend
    looplabel3aft:
    if (al(arr(1)) > 0) then 
        nt st (s, left(s, len(s) - 1))
    endif
    nt st (s, s + ")")
    st parenthesizelots, s
    exit function
end function


private function ExprVisitor_visitAssignExpr(byval expression as AssignExpr) as variant
    st ExprVisitor_visitAssignExpr, "(assign " & expression.assignee.accept(me) & " " & expression.value.accept(me) & ")"
    exit function
end function


private function ExprVisitor_visitBinaryExpr(byval expression as BinaryExpr) as variant
    st ExprVisitor_visitBinaryExpr, parenthesize(expression.operator.m_lexeme, expression.left, expression.right)
    exit function
end function


private function ExprVisitor_visitCallExpr(byval expression as CallExpr) as variant
    dim  args as string: st args, ""
    dim  innerexpr as Expr
    dim  i as long
    for i = 0 to al(expression.arguments) - 1
        st innerexpr, expression.arguments(i)
        nt st (args, args & innerexpr.accept(me) & ",")
        looplabel4bef:
    next i
    looplabel4aft:
    if (al(expression.arguments) > 0) then 
        nt st (args, left(args, len(args) - 1))
    endif
    st ExprVisitor_visitCallExpr, "(call " & expression.callee.accept(me) & "(" & args & "))"
    exit function
end function


private function ExprVisitor_visitDotExpr(byval expression as DotExpr) as variant
    st ExprVisitor_visitDotExpr, expression.left.accept(me) & "." & expression.right.accept(me)
    exit function
end function


private function ExprVisitor_visitGroupingExpr(byval expression as GroupingExpr) as variant
    st ExprVisitor_visitGroupingExpr, parenthesize("group", expression.expression)
    exit function
end function


private function ExprVisitor_visitLiteralExpr(byval expression as LiteralExpr) as variant
    if (expression.value = "null") then 
        st ExprVisitor_visitLiteralExpr, "null"
        exit function
    else
         dim rawquote as string : rawquote = """" 
        if (typename(expression.value) = "String") then 
            st ExprVisitor_visitLiteralExpr, rawquote & expression.value & rawquote
            exit function
        else
            st ExprVisitor_visitLiteralExpr, expression.value
            exit function
        endif
    endif
end function


private function ExprVisitor_visitLogicalExpr(byval expression as LogicalExpr) as variant
    st ExprVisitor_visitLogicalExpr, parenthesize(expression.operator.m_lexeme, expression.left, expression.right)
    exit function
end function


private function ExprVisitor_visitUnaryExpr(byval expression as UnaryExpr) as variant
    st ExprVisitor_visitUnaryExpr, parenthesize(expression.operator.m_lexeme, expression.right)
    exit function
end function


private function ExprVisitor_visitVariableExpr(byval expression as VariableExpr) as variant
    st ExprVisitor_visitVariableExpr, expression.name.m_lexeme
    exit function
end function


private function ExprVisitor_visitAccessorExpr(byval expression as AccessorExpr) as variant
    dim  args as string: st args, ""
    dim  innerexpr as Expr
    dim  i as long
    for i = 0 to al(expression.arguments) - 1
        st innerexpr, expression.arguments(i)
        nt st (args, args & innerexpr.accept(me) & ",")
        looplabel5bef:
    next i
    looplabel5aft:
    if (al(expression.arguments) > 0) then 
        nt st (args, left(args, len(args) - 1))
    endif
    st ExprVisitor_visitAccessorExpr, "(" & expression.accessee.accept(me) & "[" & args & "])"
    exit function
end function


private function ExprVisitor_visitFunctExpr(byval expression as FunctExpr) as variant
    st ExprVisitor_visitFunctExpr, "lambda"
    exit function
end function


private function StmtVisitor_visitVarStmt(byval statement as VarStmt) as variant
    dim  varex as VariableExpr: st varex, new VariableExpr
    call varex.init(statement.name)
    if (statement.hasinitializer) then 
        st StmtVisitor_visitVarStmt, parenthesize("set", varex, statement.initializer)
        exit function
    else
        st StmtVisitor_visitVarStmt, parenthesize("set", varex)
        exit function
    endif
end function


private function StmtVisitor_visitExprStmt(byval statement as ExprStmt) as variant
    st StmtVisitor_visitExprStmt, statement.expression.accept(me)
    exit function
end function


private function StmtVisitor_visitPrintStmt(byval statement as PrintStmt) as variant
    st StmtVisitor_visitPrintStmt, parenthesize("print", statement.expression)
    exit function
end function


private function StmtVisitor_visitBlockStmt(byval statement as BlockStmt) as variant
    st StmtVisitor_visitBlockStmt, parenthesizelots("block", statement.statements)
    exit function
end function


private function StmtVisitor_visitIfStmt(byval statement as IfStmt) as variant
    if (statement.hasElse) then 
        st StmtVisitor_visitIfStmt, "(if " & cstr(statement.condition.accept(me)) & " " & cstr(statement.thenBranch.accept(me)) & Cstr(statement.elseBranch.accept(me)) & ")"
        exit function
    else
        dim  cs as string: st cs, cstr(statement.condition.accept(me))
        dim  bs as string: st bs, cstr(statement.thenBranch.accept(me))
        st StmtVisitor_visitIfStmt, "(if " & cs & " " & bs & ")"
        exit function
    endif
end function


private function StmtVisitor_visitWhileStmt(byval statement as WhileStmt) as variant
    st StmtVisitor_visitWhileStmt, "(while " & statement.condition.accept(me) & statement.body.accept(me) & ")"
    exit function
end function


private function StmtVisitor_visitFunctionStmt(byval statement as FunctionStmt) as variant
    dim  args as string: st args, ""
    dim  i as long: st i, 0
    while (i < al(statement.params))
        dim  innerexpr as Token: st innerexpr, statement.params(i)
        dim  innertoken as Token: st innertoken, statement.paramtypes(i)
        nt st (args, args & innertoken.m_lexeme & " " & innerexpr.m_lexeme & ",")
        looplabel6bef:
        nt st (i, i + 1)
    wend
    looplabel6aft:
    if (al(statement.params) > 0) then 
        nt st (args, left(args, len(args) - 1))
    endif
    st StmtVisitor_visitFunctionStmt, "(fun " & statement.name.m_lexeme & "(" & args & ")" & parenthesizelots("body", statement.body) & ")"
    exit function
end function


private function StmtVisitor_visitClassStmt(byval statement as ClassStmt) as variant
    st StmtVisitor_visitClassStmt, "class" & statement.name.m_lexeme
    exit function
end function


private function StmtVisitor_visitCodeStmt(byval statement as CodeStmt) as variant
    st StmtVisitor_visitCodeStmt, statement.code
    exit function
end function


private function StmtVisitor_visitReturnStmt(byval statement as ReturnStmt) as variant
    st StmtVisitor_visitReturnStmt, "return"
    exit function
end function


private function StmtVisitor_visitForEachStmt(byval statement as ForEachStmt) as variant
    st StmtVisitor_visitForEachStmt, "(foreach" & statement.iterable.accept(me) & statement.body.accept(me) & ")"
    exit function
end function


private function StmtVisitor_visitControlStmt(byval statement as ControlStmt) as variant
    st StmtVisitor_visitControlStmt, statement.keyword.m_lexeme
    exit function
end function


private function StmtVisitor_visitEnumStmt(byval statement as EnumStmt) as variant
    st StmtVisitor_visitEnumStmt, "enum " & statement.name.m_lexeme
    exit function
end function



