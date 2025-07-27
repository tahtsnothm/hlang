VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Token"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_type_ As tokentype
Public m_lexeme As String
Public m_literal As Variant
Public m_line As Long

Public Function init(type_ As tokentype, lexeme As String, literal As Variant, line As Long)
m_type_ = type_
m_lexeme = lexeme
m_literal = literal
m_line = line
End Function

Public Function toString() As String
    toString = m_type_ & " " & m_lexeme & " " & literal
End Function

