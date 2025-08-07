Attribute VB_Name = "VarStmt"
implements Stmt
public vis as string
public type_ as token
public name as token
public initializer as Expr
public addnew as boolean
public hasinitializer as boolean
public arrcount as Expr
public hasarrcount as boolean
public function init(byval a as string, byval b as token, byval c as token, byval d as expr, byval e as boolean, byval f as boolean, byval g as Expr, byval h as boolean) 
    nt st (vis, a)
    nt st (type_, b)
    nt st (name, c)
    nt st (initializer, d)
    nt st (addnew, e)
    nt st (arrcount, f)
    nt st (hasarrcount, g)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitVarStmt(Me)
    exit function
end function



