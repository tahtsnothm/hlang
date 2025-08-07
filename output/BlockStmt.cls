Attribute VB_Name = "BlockStmt"
implements Stmt
public statements as variant
public function init(byval a as variant) 
    nt st (statements, a)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitBlockStmt(Me)
    exit function
end function



