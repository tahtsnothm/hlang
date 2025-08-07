Attribute VB_Name = "CodeStmt"
implements Stmt
public code as string
public function init(byval a as string) 
    nt st (code, a)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitCodeStmt(Me)
    exit function
end function



