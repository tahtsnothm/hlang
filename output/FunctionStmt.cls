VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FunctionStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
implements Stmt
public name as Token
public params as variant
public paramtypes as variant
public parammethod as variant
public body as variant
public returntype as token
public visibility as string
public function init(byval a as Token, byval b as variant, byval c as variant, byval d as variant, byval e as variant, byval f as token, byval g as string) 
    nt st (name, a)
    nt st (params, b)
    nt st (paramtypes, c)
    nt st (parammethod, d)
    nt st (body, e)
    nt st (visibility, f)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitFunctionStmt(Me)
    exit function
end function



