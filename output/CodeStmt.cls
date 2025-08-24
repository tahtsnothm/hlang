VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
implements Stmt
public code as string
public function init(byval a as string) 
    nt st (code, a)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitCodeStmt(Me)
    exit function
end function



