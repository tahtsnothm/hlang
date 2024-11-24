VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AstPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements ExprVisitor
Implements StmtVisitor

Public Function prt(statements As Variant)
    Dim statement As Stmt
    Dim i As Long
    For i = 0 To al(statements) - 1
        st statement, statements(i)
        Dim s As String:
        s = CStr(statement.accept(Me))
        pt s
    Next i
    'st prt, expression.accept(Me)
End Function


Private Function parenthesize(ParamArray arr() As Variant) As String
    Dim s As String
    Dim i As Long
    s = ""
    s = s + "("
    s = s + arr(0)
    Dim expression As Expr
    For i = 1 To al(arr) - 1
        st expression, arr(i)
        s = s + " "
        s = s + CStr(expression.accept(Me))

    Next i
    s = s + ")"
    parenthesize = s
End Function

Private Function parenthesizestatements(ParamArray arr() As Variant) As String
    Dim s As String
    Dim i As Long
    s = ""
    s = s + "("
    s = s + arr(0)
    Dim expression As Stmt
    For i = 1 To al(arr) - 1
        st expression, arr(i)
        s = s + " "
        s = s + CStr(expression.accept(Me))

    Next i
    s = s + ")"
    parenthesize = s
End Function
Private Function parenthesizelots(ParamArray arr() As Variant) As String
    Dim s As String
    Dim i As Long
    s = ""
    s = s + "("
    s = s + arr(0)
    
    Dim expression As Stmt
    For i = 0 To al(arr(1)) - 1
        st expression, arr(1)(i)
        s = s + CStr(expression.accept(Me))
        s = s + " "
    Next i
    s = left(s, Len(s) - 1)
    s = s + ")"
    parenthesizelots = s
End Function



Private Function ExprVisitor_visitAssignExpr(expression As AssignExpr) As Variant
    st ExprVisitor_visitAssignExpr, parenthesize("assign", expression.name, expression.value)
End Function

Private Function ExprVisitor_visitBinaryExpr(expression As BinaryExpr) As Variant
    ExprVisitor_visitBinaryExpr = parenthesize(expression.operator.m_lexeme, expression.left, expression.right)
End Function

Private Function ExprVisitor_visitGroupingExpr(expression As GroupingExpr) As Variant
    ExprVisitor_visitGroupingExpr = parenthesize("group", expression.expression)
End Function

Private Function ExprVisitor_visitLiteralExpr(expression As LiteralExpr) As Variant
If (expression.value = Null) Then
        ExprVisitor_visitLiteralExpr = "null"
    Else
        If (TypeName(expression.value) = "String") Then
            ExprVisitor_visitLiteralExpr = """" & expression.value & """"
        Else
            ExprVisitor_visitLiteralExpr = expression.value
        End If
        
        
    End If
    
End Function

Private Function ExprVisitor_visitLogicalExpr(expression As LogicalExpr) As Variant
    st ExprVisitor_visitLogicalExpr, parenthesize(expression.operator.m_lexeme, expression.left, expression.right)
End Function

Private Function ExprVisitor_visitUnaryExpr(expression As UnaryExpr) As Variant
ExprVisitor_visitUnaryExpr = parenthesize(expression.operator.m_lexeme, expression.right)
End Function

Private Function ExprVisitor_visitVariableExpr(expression As VariableExpr) As Variant
    st ExprVisitor_visitVariableExpr, expression.name.m_lexeme
End Function


Private Function StmtVisitor_visitBlockStmt(statement As BlockStmt) As Variant
    st StmtVisitor_visitBlockStmt, parenthesizelots("block", statement.statements)
End Function

Private Function StmtVisitor_visitExprStmt(statement As ExprStmt) As Variant
    st StmtVisitor_visitExprStmt, statement.expression.accept(Me)
End Function

Private Function StmtVisitor_visitIfStmt(statement As IfStmt) As Variant
    If (statement.hasElse) Then
        st StmtVisitor_visitIfStmt, "(if " & CStr(statement.condition.accept(Me)) & " " & CStr(statement.thenBranch.accept(Me)) & " " & CStr(statement.elseBranch.accept(Me)) & ")"
    Else
        Dim cs As String
        Dim bs As String
        cs = CStr(statement.condition.accept(Me))
        bs = CStr(statement.thenBranch.accept(Me))
        st StmtVisitor_visitIfStmt, "(if " & CStr(statement.condition.accept(Me)) & " " & CStr(statement.thenBranch.accept(Me)) & ")"
    End If

End Function

Private Function StmtVisitor_visitPrintStmt(statement As PrintStmt) As Variant
    st StmtVisitor_visitPrintStmt, parenthesize("print", statement.expression)
End Function

Private Function StmtVisitor_visitVarStmt(statement As VarStmt) As Variant
    Dim varex As New VariableExpr
    varex.Expr_init statement.name
    
    If statement.hasinitializer Then
        st StmtVisitor_visitVarStmt, parenthesize("set", varex, statement.initializer)
    Else
        st StmtVisitor_visitVarStmt, parenthesize("set", varex)
    End If
    
    
End Function

Private Function StmtVisitor_visitWhileStmt(statement As WhileStmt) As Variant
    st StmtVisitor_visitWhileStmt, "(while " & statement.condition.accept(Me) & statement.body.accept(Me) & ")"
End Function
