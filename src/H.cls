VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "H"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private hadError As Boolean
Private Sub Class_Initialize()
    hadError = False
    
End Sub

Public Function runFile(filename As String)
Dim fso As New FileSystemObject
Dim ts As TextStream
Set ts = fso.OpenTextFile(filename)

Dim wodecode As String
wodecode = ts.ReadAll
wodecode = Replace(wodecode, vbCrLf, vbLf)
runCode wodecode
ts.Close

End Function

Public Function runPrompt()
' lazy to implement
End Function

Public Function writeToFile(filename As String, text As String)
Dim fso As New FileSystemObject
Dim ts As TextStream
If fso.FileExists(filename) Then
    ts = fso.OpenTextFile(filename)
Else
    ts = fso.CreateTextFile(filename)
End If

ts.Write text
ts.Close
End Function
Static Function error(ByVal line As Long, message As String)
    report line, "", message
End Function
Private Function report(line As Long, where As String, message As String)
    pt "[line", line, "] Error", where, ": ", message
    hadError = True
End Function

Static Function err(tk As Token, message As String)
    If (tk.m_type_ = EOF_) Then
        report tk.m_line, " at end.", message
    Else
        report tk.m_line, " at '" + tk.m_lexeme + "'", message
        
    End If
        
    
End Function

Private Function runCode(source As String)
Dim Scanner As New Scanner
Dim tk As Variant
Dim tokens() As Variant ' list<Tokens>
Scanner.init source, Me
tokens = Scanner.scanTokens()
Dim Parser As New Parser
Parser.init tokens, Me
Dim statements As Variant
st statements, Parser.parse()

If hadError Then
    pt "ERROROROROR"
    Exit Function
End If
Dim astp As New AstPrinter
'debugger
astp.prt statements

Dim gen As New CodeGenerator

gen.generate statements
'For Each tk In tokens
 '   pt tk.toString()
'Next tk
End Function


