Attribute VB_Name = "ForEachStmt"
implements Stmt
public objdeclarator as VarStmt
public obj as Token
public iterable as Expr
public iteratordeclarator as VarStmt
public iterator as Token
public body as Stmt
public hasDeclarator as boolean
public hasIteratorDeclarator as boolean
public function init(byval a as VarStmt, byval b as Token, byval c as Expr, byval d as VarStmt, byval e as Token, byval f as Stmt, byval g as boolean, byval h as boolean) 
    nt st (objdeclarator, a)
    nt st (obj, b)
    nt st (iterable, c)
    nt st (iteratordeclarator, d)
    nt st (iterator, e)
    nt st (body, f)
    nt st (hasDeclarator, g)
    nt st (hasIteratorDeclarator, h)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitForEachStmt(Me)
    exit function
end function



