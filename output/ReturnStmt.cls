Attribute VB_Name = "ReturnStmt"
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



