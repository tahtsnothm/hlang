VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IfStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Stmt
Public condition As Expr
Public thenBranch As Stmt
Public elseBranch As Stmt
Public hasElse As Boolean

Public Function Stmt_init(ParamArray arr() As Variant)
    st condition, arr(0)
    st thenBranch, arr(1)
    If (al(arr) = 2) Then
        hasElse = False
    Else
        st elseBranch, arr(2)
        hasElse = True
    End If
    
End Function

Public Function Stmt_accept(visitor As StmtVisitor) As Variant
    st Stmt_accept, visitor.visitIfStmt(Me)
End Function



