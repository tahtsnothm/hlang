VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StmtVisitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function visitVarStmt(statement As VarStmt)
End Function
Public Function visitExprStmt(statement As ExprStmt)
End Function
Public Function visitPrintStmt(statement As PrintStmt)
End Function

Public Function visitBlockStmt(statement As BlockStmt)
End Function

Public Function visitIfStmt(statement As IfStmt)
End Function
Public Function visitWhileStmt(statement As WhileStmt)
End Function

Public Function visitFunctionStmt(statement As FunctionStmt)
End Function
