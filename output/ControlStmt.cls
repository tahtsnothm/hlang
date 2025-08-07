Attribute VB_Name = "ControlStmt"
implements Stmt
public keyword as Token
public function init(byval a as Token) 
    nt st (keyword, a)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitControlStmt(Me)
    exit function
end function



