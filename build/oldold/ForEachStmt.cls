VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ForEachStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Stmt
Public objdeclarator As VarStmt
Public obj As Token

Public iterable As Expr
Public iteratordeclarator As VarStmt
Public iterator As Token
Public body As Stmt
Public hasDeclarator As Boolean
Public hasIteratorDeclarator As Boolean



Public Function Stmt_init(ParamArray arr() As Variant)
    st objdeclarator, arr(0)
    st obj, arr(1)
    st iterable, arr(2)
    st iteratordeclarator, arr(3)
    st iterator, arr(4)
    st body, arr(5)
    st hasDeclarator, arr(6)
    st hasIteratorDeclarator, arr(7)
    
    
    ' iterable , iteratingobj, iterator, iteratingobjtype, iteratorType, hasDeclarator, hasIteratorDeclarator
End Function

Public Function Stmt_accept(visitor As StmtVisitor) As Variant
    st Stmt_accept, visitor.visitForEachStmt(Me)
End Function

