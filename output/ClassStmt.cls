Attribute VB_Name = "ClassStmt"
implements Stmt
public name as Token
public methods as variant
public superclasses as variant
public properties as variant
public function init(byval a as Token, byval b as variant, byval c as variant, byval d as variant) 
    nt st (name, a)
    nt st (methods, b)
    nt st (superclasses, c)
    nt st (properties, d)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitClassStmt(Me)
    exit function
end function



