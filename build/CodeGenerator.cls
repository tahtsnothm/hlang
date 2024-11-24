VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements ExprVisitor
Implements StmtVisitor

Public output_file_path As String
Public functions As String
Public inMain As Boolean
Public indentLevel As Long
Public indentType As String
Private f As Object


Private Sub Class_Initialize()
    ' default function will be in main
    ' theres a main thread for stuff
    ' and another one for functions
    output_file_path = ThisWorkbook.path & "\out.bas"
    inMain = True
    indentLevel = 0
    indentType = "  "
    ' to change when classes are supported
End Sub


Public Function generate(statements As Variant)
    Dim statement As Stmt
    Dim i As Long
    
    Dim fso As New FileSystemObject
    st f, fso.CreateTextFile(output_file_path)
    
    write_ "public sub main ()"
    indentLevel = 1
    For i = 0 To al(statements) - 1
        st statement, statements(i)
        statement.accept Me
    Next i
    indentLevel = 0
    write_ "end sub"
    f.Close
End Function
Private Function indent()
    indentLevel = indentLevel + 1
End Function
Private Function unindent()
    indentLevel = indentLevel - 1
    If indentLevel < 0 Then indentLevel = 0
End Function

Private Function write_(s As String)
    Dim i As Long
    For i = 0 To indentLevel - 1
        s = indentType & s
    Next i
    If inMain Then
        f.WriteLine s
    Else
        functions = functions & vbLf & s
    End If
    
End Function

Private Function ExprVisitor_visitAssignExpr(expression As AssignExpr) As Variant
    st ExprVisitor_visitAssignExpr, CStr(expression.assignee.accept(Me)) & " = " & expression.value.accept(Me)
End Function

Private Function ExprVisitor_visitBinaryExpr(expression As BinaryExpr) As Variant
    
    st ExprVisitor_visitBinaryExpr, expression.left.accept(Me) & " " & expression.operator.m_lexeme & " " & expression.right.accept(Me)
End Function

Private Function ExprVisitor_visitCallExpr(expression As CallExpr) As Variant
    Dim args As String: args = ""
    Dim ee As Expr
    Dim i As Long:
    For i = 0 To al(expression.arguments) - 1
        st ee, expression.arguments(i)
        args = args & ee.accept(Me) & ", "
    Next i
    If al(expression.arguments) > 0 Then args = left(args, Len(args) - 2)
    st ExprVisitor_visitCallExpr, expression.callee.accept(Me) & "(" & args & ")"
End Function

Private Function ExprVisitor_visitGroupingExpr(expression As GroupingExpr) As Variant
    'is this really needed half the time but sure
    st ExprVisitor_visitGroupingExpr, "(" & expression.expression.copy().Expr_accept(Me) & ")"
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
    
End Function

Private Function ExprVisitor_visitUnaryExpr(expression As UnaryExpr) As Variant
    If (expression.operator.m_lexeme = "!") Then
        st ExprVisitor_visitUnaryExpr, "not " & expression.right.accept(Me)
    ElseIf (expression.operator.m_lexeme = "-") Then
        st ExprVisitor_visitUnaryExpr, "-" & expression.right.accept(Me)
    End If
End Function

Private Function ExprVisitor_visitVariableExpr(expression As VariableExpr) As Variant
    st ExprVisitor_visitVariableExpr, expression.name.m_lexeme
End Function

Private Function StmtVisitor_visitBlockStmt(statement As BlockStmt) As Variant
    Dim statemt As Stmt
    Dim i As Long
    For i = 0 To al(statement.statements) - 1
        st statemt, statement.statements(i)
        statemt.accept Me
    Next i
End Function

Private Function StmtVisitor_visitClassStmt(statement As ClassStmt) As Variant

End Function

Private Function StmtVisitor_visitCodeStmt(statement As CodeStmt) As Variant
    write_ statement.code
End Function

Private Function StmtVisitor_visitExprStmt(statement As ExprStmt) As Variant
    write_ statement.expression.accept(Me)
End Function

Private Function StmtVisitor_visitFunctionStmt(statement As FunctionStmt) As Variant

End Function

Private Function StmtVisitor_visitIfStmt(statement As IfStmt) As Variant
    write_ "if (" & statement.condition.accept(Me) & ") then "
    indent
    statement.thenBranch.accept Me
    unindent
    If (statement.hasElse) Then
        write_ "else"
        indent
        statement.elseBranch.accept Me
        unindent
    End If
    write_ "endif"
End Function

Private Function StmtVisitor_visitPrintStmt(statement As PrintStmt) As Variant
    write_ "pt " & statement.expression.accept(Me)
End Function

Private Function StmtVisitor_visitReturnStmt(statement As ReturnStmt) As Variant

End Function

Private Function StmtVisitor_visitVarStmt(statement As VarStmt) As Variant
    Dim s As String: s = "dim " & statement.name.m_lexeme & " as " & statement.type_.m_lexeme
    If statement.hasinitializer Then
        s = s & ": " & statement.name.m_lexeme & " = " & statement.initializer.accept(Me)
    End If
    write_ s
End Function

Private Function StmtVisitor_visitWhileStmt(statement As WhileStmt) As Variant
    write_ "while (" & statement.condition.accept(Me) & ")"
    indent
    statement.body.accept Me
    unindent
    write_ "wend"
End Function
