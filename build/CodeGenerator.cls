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
Public indentLevel As Long
Public indentType As String
Public ilevels As Variant
Public scopes As Variant

Public fcounter As Long

Public current_function_name As String
Private f As Object


Private Sub Class_Initialize()
    ' default function will be in main
    ' theres a main thread for stuff
    ' and another one for functions
    output_file_path = ThisWorkbook.path & "\out.bas"
    indentLevel = 0
    current_function_name = "main"
    
    ilevels = Array(0) '
    scopes = Array("main") ' just as a buffer for pop,
    indentType = "    "
    fcounter = 0
    
    ' to change when classes are supported
End Sub


Public Function generate(statements As Variant)
    Dim statement As Stmt
    Dim i As Long
    
    Dim fso As New FileSystemObject
    st f, fso.CreateTextFile(output_file_path)
    
    write_ "public sub main ()"
    indent
    For i = 0 To al(statements) - 1
        st statement, statements(i)
        statement.accept Me
    Next i
    unindent
    write_ "end sub"
    
    write_ functions
    f.Close
End Function
Private Function indent()
    indentLevel = indentLevel + 1
End Function
Private Function unindent()
    indentLevel = indentLevel - 1
    If indentLevel < 0 Then indentLevel = 0
End Function
Private Function pushscope(Optional fname As String = "")

arad ilevels, indentLevel
arad scopes, current_function_name

current_function_name = fname
indentLevel = 0

End Function
Private Function popscope()
indentLevel = ilevels(al(ilevels) - 1)
pop ilevels, al(ilevels) - 1
current_function_name = scopes(al(scopes) - 1)
pop scopes, al(scopes) - 1

End Function

Private Function write_(s As String)
    Dim i As Long
    For i = 0 To indentLevel - 1
        s = indentType & s
    Next i
    If current_function_name = "main" Then
        f.WriteLine s
    Else
        functions = functions & vbLf & s
    End If
    
End Function

Private Function ctyp(s As String)
'cleans types from int/ string/ array/ to the proper types
Select Case s
    Case "int":
        ctyp = "integer"
    Case "bool":
        ctyp = "boolean"
    Case "array":
        ctyp = "variant"
    Case "void":
        ctyp = "variant"
    Case Else:
        ctyp = s
End Select
End Function

Private Function ExprVisitor_visitAssignExpr(expression As AssignExpr) As Variant
    st ExprVisitor_visitAssignExpr, "st (" & CStr(expression.assignee.accept(Me)) & ", " & expression.value.accept(Me) & ")"
End Function

Private Function ExprVisitor_visitBinaryExpr(expression As BinaryExpr) As Variant
    If (expression.operator.m_lexeme = "==") Then expression.operator.m_lexeme = "="
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
    Dim left As String: st left, expression.left.accept(Me)
    Dim right As String: st right, expression.right.accept(Me)
    Dim fname As String: fname = "funcl" & expression.operator.m_line & "c" & fcounter
    fcounter = fcounter + 1
    pushscope fname
    'resolver please
    write_ "private function " & fname & "()" & " as Boolean"
    indent
        write_ fname & " = false"
        write_ "if (" & left & ") then"
        indent
            write_ "if (" & right & ") then"
            indent
                write_ fname & " = true"
            unindent
            write_ "endif"
        unindent
        write_ "endif"
    unindent
    write_ "end function"
    popscope
    st ExprVisitor_visitLogicalExpr, fname & "()"
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
    If (TypeName(statement.expression) = "CallExpr") Then
        write_ "call " & statement.expression.accept(Me)
    Else
        write_ "nt " & statement.expression.accept(Me)
    End If
    
    'write_ "nt " & statement.expression.accept(Me)
End Function

Private Function StmtVisitor_visitForEachStmt(statement As ForEachStmt) As Variant


If (statement.hasDeclarator) Then
    statement.objdeclarator.Stmt_accept Me
End If

If (statement.hasIteratorDeclarator) Then
    statement.iteratordeclarator.Stmt_accept Me
End If


 write_ "for " & statement.iterator.m_lexeme & " = 0 to al(" & statement.iterable.accept(Me) & ") - 1"
    indent
    write_ "st " & statement.obj.m_lexeme & ", " & statement.iterable.accept(Me) & "(" & statement.iterator.m_lexeme & ")"
    statement.body.accept Me
    unindent
    write_ "next " & statement.iterator.m_lexeme
End Function

Private Function StmtVisitor_visitFunctionStmt(statement As FunctionStmt) As Variant

pushscope statement.name.m_lexeme

Dim args As String: args = ""
Dim i As Long
For i = 0 To al(statement.params) - 1

    Dim bywhat As String
    If statement.parammethod(i) Then
        bywhat = "byref"
    Else
        bywhat = "byval"
    End If
    Dim pname As Token: st pname, statement.params(i)
    Dim ptype As Token: st ptype, statement.paramtypes(i)
    Dim paramname As String: paramname = statement.params(i).m_lexeme
    Dim paramtype As String: paramtype = statement.paramtypes(i).m_lexeme
    args = args + bywhat + " " + paramname + " as " + ctyp(paramtype) + ", "

Next i
If al(statement.params) > 0 Then args = left(args, Len(args) - 2)
If statement.returntype.m_lexeme = "void" Then
    write_ "function " & statement.name.m_lexeme & "(" & args & ") "
Else
    write_ "function " & statement.name.m_lexeme & "(" & args & ") " & "as " & ctyp(statement.returntype.m_lexeme)
End If
indent

Dim bodyLine As Stmt
st i, 0
While (i < al(statement.body))
    st bodyLine, statement.body(i)
    bodyLine.accept Me
    i = i + 1
Wend


unindent
write_ "end function"
write_ vbLf
popscope
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
    If (current_function_name = "") Then
        pt "Error: line", statement.keyword.m_line, " Cannot return"
    Else
        If (statement.hasValue) Then
        write_ "st " & current_function_name & ", " & statement.value.accept(Me)
        End If
        write_ "exit function"
    End If
    
End Function

Private Function StmtVisitor_visitVarStmt(statement As VarStmt) As Variant
    Dim s As String:
    If statement.type_.m_lexeme = "array" Then
        s = "dim " & statement.name.m_lexeme & "() as " & ctyp(statement.type_.m_lexeme)
    Else
        s = "dim " & statement.name.m_lexeme & " as " & ctyp(statement.type_.m_lexeme)
    End If
   
    If statement.hasinitializer Then
        s = s & ": st " & statement.name.m_lexeme & ", " & statement.initializer.accept(Me)
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
