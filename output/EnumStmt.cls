VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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



