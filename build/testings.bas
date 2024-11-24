Attribute VB_Name = "testings"
Public Sub main()
    pt fibonacci(10)
End Sub

Function fibonacci(ByVal n As Integer) As Integer
    If (n <= 1) Then
        st fibonacci, 1
        Exit Function
    End If
    st fibonacci, fibonacci(n - 1) + fibonacci(n - 2)
    Exit Function
End Function



