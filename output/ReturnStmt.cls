VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReturnStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
implements Stmt
public keyword as Token
public hasValue as boolean
public value as Expr
public function init(byval a as Token, byval b as Expr, byval c as boolean) 
    nt st (keyword, a)
    nt st (value, b)
    nt st (hasValue, c)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitReturnStmt(Me)
    exit function
end function



