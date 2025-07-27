VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExprVisitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function visitUnaryExpr(expression As UnaryExpr)
End Function

Public Function visitBinaryExpr(expression As BinaryExpr)
End Function
Public Function visitLiteralExpr(expression As LiteralExpr)
End Function
Public Function visitGroupingExpr(expression As GroupingExpr)
End Function
Public Function visitAssignExpr(expression As AssignExpr)
End Function
Public Function visitVariableExpr(expression As VariableExpr)
End Function
Public Function visitLogicalExpr(expression As LogicalExpr)
End Function
Public Function visitCallExpr(expression As CallExpr)
End Function
Public Function visitDotExpr(expression As DotExpr)
End Function
