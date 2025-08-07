Attribute VB_Name = "PrintStmt"
implements Stmt
public expression as Expr
public function init(byval a as Expr) 
    nt st (expression, a)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitPrintStmt(Me)
    exit function
end function



