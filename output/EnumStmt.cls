Attribute VB_Name = "EnumStmt"
implements Stmt
public name as Token
public types_ as variant
public function init(byval a as Token, byval b as variant) 
    nt st (name, a)
    nt st (types_, b)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitEnumStmt(Me)
    exit function
end function



