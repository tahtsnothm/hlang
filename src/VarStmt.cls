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
Public vis As String
Public type_ As Token
Public name As Token
Public initializer As Expr
Public addnew As Boolean
Public hasinitializer As Boolean
Public arrcount As Expr
Public hasarrcount As Boolean
Public Function Stmt_init(ParamArray arr() As Variant)
    st vis, arr(0)
    st type_, arr(1)
    st name, arr(2)
    st addnew, arr(3)
    st hasinitializer, arr(4)
    st initializer, arr(5)
    st arrcount, arr(6)
    st hasarrcount, arr(7)
End Function

Public Function Stmt_accept(visitor As StmtVisitor) As Variant
    st Stmt_accept, visitor.visitVarStmt(Me)
End Function


