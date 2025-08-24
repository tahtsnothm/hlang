VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExprVisitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
public function visitUnaryExpr(byval expression as UnaryExpr) as variant
end function


public function visitBinaryExpr(byval expression as BinaryExpr) as variant
end function


public function visitLiteralExpr(byval expression as LiteralExpr) as variant
end function


public function visitGroupingExpr(byval expression as GroupingExpr) as variant
end function


public function visitAssignExpr(byval expression as AssignExpr) as variant
end function


public function visitVariableExpr(byval expression as VariableExpr) as variant
end function


public function visitLogicalExpr(byval expression as LogicalExpr) as variant
end function


public function visitCallExpr(byval expression as CallExpr) as variant
end function


public function visitDotExpr(byval expression as DotExpr) as variant
end function


public function visitAccessorExpr(byval expression as AccessorExpr) as variant
end function


public function visitFunctExpr(byval expression as FunctExpr) as variant
end function



