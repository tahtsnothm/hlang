VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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



