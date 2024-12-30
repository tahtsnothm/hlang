Attribute VB_Name = "Module2"
Function test()
    Dim a() As Variant: st a, Array("one", "two", "three", "four", "five")
    Dim s As String
    Dim i As Integer
    For i = 0 To al(a) - 1
        st s, a(i)
        pt (s)
    Next i
End Function
Sub t()
test

End Sub
