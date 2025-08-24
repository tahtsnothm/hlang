VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WhileStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
implements Stmt
public condition as expr
public body as Stmt
public itx as Stmt
public function init(byval a as expr, byval b as stmt, byval c as stmt) 
    nt st (condition, a)
    nt st (body, b)
    nt st (itx, c)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitWhileStmt(Me)
    exit function
end function



