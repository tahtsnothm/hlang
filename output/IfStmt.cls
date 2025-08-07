Attribute VB_Name = "IfStmt"
implements Stmt
public condition as Expr
public thenBranch as Stmt
public elseBranch as Stmt
public hasElse as boolean
public function init(byval a as Expr, byval b as Stmt, byval c as Stmt, byval d as boolean) 
    nt st (condition, a)
    nt st (thenBranch, b)
    nt st (elseBranch, c)
    nt st (hasElse, d)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitIfStmt(Me)
    exit function
end function



