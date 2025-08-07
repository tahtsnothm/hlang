Attribute VB_Name = "Token"
public m_type_ as tokentype
public m_lexeme as string
public m_literal as variant
public m_line as long
public function init(byval a as tokentype, byval b as string, byval c as variant, byval d as long) 
    nt st (m_type_, a)
    nt st (m_lexeme, b)
    nt st (m_literal, c)
    nt st (m_line, d)
end function


public function toString() as string
    st toString, m_type_ & " " & m_lexeme & " " & literal
    exit function
end function



