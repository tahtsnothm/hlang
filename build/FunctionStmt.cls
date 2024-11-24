VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FunctionStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Stmt
Public name As Token
Public params As Variant ' of tokens
Public paramtypes As Variant ' of tokens
Public body As Variant 'of Stmts
Public returntype As Token


Public Function Stmt_init(ParamArray arr() As Variant)
    st name, arr(0)
    st returntype, arr(1)
    st params, arr(2)
    st paramtypes, arr(3)
    st body, arr(4)
End Function

Public Function Stmt_accept(visitor As StmtVisitor) As Variant
    st Stmt_accept, visitor.visitFunctionStmt(Me)
End Function



