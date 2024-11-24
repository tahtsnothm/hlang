VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VarStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Stmt
Public type_ As Token
Public name As Token
Public initializer As Expr
Public hasinitializer As Boolean

Public Function Stmt_init(ParamArray arr() As Variant)
    st type_, arr(0)
    st name, arr(1)
    If (al(arr) = 2) Then
        hasinitializer = False
    Else
        st initializer, arr(2)
        hasinitializer = True
    End If
    
End Function

Public Function Stmt_accept(visitor As StmtVisitor) As Variant
    st Stmt_accept, visitor.visitVarStmt(Me)
End Function


