Attribute VB_Name = "sheetutils"
Public outfold As String
Public templates As String
Const cvsource As String = "D:\"




Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Public Sub updatePaths()
sheetutils.templates = "D:\usersp\" & Environ("UserName") & "\Desktop\Templates\"
sheetutils.outfold = "D:\usersp\" & Environ("UserName") & "\Desktop\"
End Sub

Public Function rtd(row As Range, rowheaders As Variant) As Object
Dim i As Long
Dim res As Object
Set res = CreateObject("Scripting.Dictionary")
For i = 1 To al(rowheaders)
    res(CStr(rowheaders(i - 1))) = row(i)
    

Next i

Set rtd = res
End Function

Public Function dta(ByVal row As Object, rowheaders As Variant, Optional startIndex As Long = 1) As Variant

Dim i As Long

Dim res() As Variant
ReDim res(startIndex To al(rowheaders) - 1 + startIndex)
For i = startIndex To al(rowheaders) - 1 + startIndex

    If row.Exists(CStr(rowheaders(i))) Then
        res(i) = row(CStr(rowheaders(i)))
    Else
        res(i) = ""
    End If
    
    
    
Next i
 dta = res
End Function

Public Function rta(ByVal row As Range, Optional startIndex As Integer = 1) As Variant
Dim i As Long
Dim res() As Variant
ReDim res(startIndex To row.Count + startIndex - 1)
For i = 1 To row.Count
    res(startIndex + i - 1) = row(i)
Next i
 rta = res
End Function

Public Function al(ByVal arr As Variant, Optional dimension As Long = 1) As Long
Dim ub As Long
ub = -1
On Error Resume Next
ub = UBound(arr, dimension)
On Error GoTo hh
hh:
Dim lb As Long
lb = 0
On Error Resume Next
lb = LBound(arr, dimension)
On Error GoTo 0

al = ub - lb + 1
End Function


Static Function puts(ByVal i As Long, ByVal j As Long, ByVal valuee As Variant, Optional ByVal ss As String)
Dim deff As String
If ss <> "" Then
deff = ss
End If

Dim sheet As Worksheet
Set sheet = Worksheets(deff)
Worksheets(deff).Cells(i, j).value = valuee
puts = valuee

End Function

Public Function GetFileDialog() As String
    Dim filedialog As Object
    Set filedialog = Application.filedialog(msoFileDialogFilePicker)
    filedialog.Show
    
    GetFileDialog = filedialog.SelectedItems(1)
End Function
Public Function arrcontains(arr As Variant, item As Variant) As Boolean
Dim i As Long
arrcontains = False
If al(arr) = 0 Then
    Exit Function
End If
For i = LBound(arr) To UBound(arr)
    If (arr(i) = item) Then
        arrcontains = True
        Exit Function
    End If
Next i
End Function

Public Function arad(ByRef arr As Variant, ByRef item As Variant)
    ReDim Preserve arr(0 To al(arr))
    If IsObject(item) Then
        Set arr(al(arr) - 1) = item
    Else
        arr(al(arr) - 1) = item
    End If
End Function

Public Function aradv(ByVal arr As Variant, ByVal item As Variant) As Variant
    ReDim Preserve arr(0 To al(arr))
    If TypeName(item) = "Dictionary" Then
        Set arr(al(arr) - 1) = item
    ElseIf IsObject(item) Then
        Set arr(al(arr) - 1) = item
    Else
        arr(al(arr) - 1) = item
    End If
    aradv = arr
End Function


Static Function pop(ByRef arr As Variant, ByVal index As Long)
Dim arrlength As Long
arrlength = al(arr)

If arrlength = 0 Then
    Exit Function
End If

If arrlength = 1 Then
    arr = Array()
    Exit Function
End If

If arrlength - 1 < index Then
    Exit Function
End If


Dim i As Long
If arrlength - 1 <> index Then
    For i = index To arrlength - 2
        arr(i) = arr(i + 1)
    Next i
End If
ReDim Preserve arr(0 To arrlength - 2)
End Function
Static Function insert(ByRef arr As Variant, ByVal item As Variant, ByVal index As Long) As Variant
Dim arrlength As Long
arrlength = al(arr)

If arrlength = 0 Or index = al(arr) Then
    arad arr, item
    insert = arr
    Exit Function
End If

If index < 0 Or index > al(arr) Then
    Exit Function
End If
    

Dim i As Long
arad arr, "temp"
For i = arrlength - 1 To index Step -1
    arr(i + 1) = arr(i)
Next i
arr(index) = item
insert = arr

End Function


'part 2
Function printFile(file As String) As Long
printFile = ShellExecute(0, "print", file, vbNullChar, vbNullChar, 0)
End Function

Public Function createFilename(instancename As String, templatename As String)
updatePaths
Dim template As String
template = templates + templatename + ".docx"

Dim outputfilename As String
outputfilename = outfold + instancename + " " + templatename + ".docx"
FileCopy template, outputfilename
createFilename = outputfilename

End Function
Public Function textDate(ByVal numberdate As Date) As String
textDate = Application.WorksheetFunction.text(numberdate, "D MMM YYYY")
End Function


Public Function strreplace(ByVal text As String, ByVal replacements As Object)
Dim key As Variant

For Each key In replacements
    text = Replace(text, CStr(key), replacements(key))
Next key

strreplace = (text)

End Function


Function transpose(a As Variant)
'm by n to n by m
Dim m As Long, n As Long
Dim b() As Variant
Dim brow() As Variant
Dim i As Long
m = al(a)
n = al(a(0))
If m = 0 Then
    transpose = b
    Exit Function
End If
If n = 0 Then
    ReDim Preserve b(0 To m - 1)
    For i = 0 To m - 1
        b(i) = Array(a(i))
    Next i
    transpose = b
    Exit Function
End If

ReDim Preserve b(0 To n - 1)
ReDim brow(0 To m - 1)

Dim j As Long
For j = 0 To n - 1
    For i = 0 To m - 1
        brow(i) = a(i)(j)
    Next i
    b(j) = brow
Next j
transpose = b
End Function

Function unwrap(ByVal arr As Variant, ByVal values As Variant) As Variant
Dim value As Variant
Dim result() As Variant
For Each value In values
    arad result, arr(value)
Next value
unwrap = result
End Function


Public Function decomp(ByRef arr As Variant, Optional ByRef rows As Long, Optional ByRef cols As Long) As Variant 'changes a 2d array (i, j) to (i)(j)
Dim rc As Long
Dim cc As Long
rc = al(arr, 1)
cc = al(arr, 2)

If (rc = 0 And cc = 0) Then
    'single value
    decomp = Array(Array(arr))
    If (Not IsMissing(rows)) Then
    rows = 1
End If
If (Not IsMissing(cols)) Then
    cols = 1
End If
    Exit Function
End If


Dim i As Long
Dim j As Long
Dim temparray() As Variant
Dim res() As Variant
For i = 1 To rc
    temparray = Array()
    For j = 1 To cc
        arad temparray, arr(i, j)

    Next j
    arad res, temparray
Next i
decomp = res
If (Not IsMissing(rows)) Then
    rows = rc
End If
If (Not IsMissing(cols)) Then
    cols = cc
End If
End Function

Public Function comp(ByRef arr As Variant) As Variant ' changes a (i)(j) to (i, j) 1 indexed
Dim rc As Long
Dim cc As Long
rc = al(arr)
cc = al(arr(0))

Dim res(1 To rc, 1 To cc) As Variant
For i = 1 To rc
    For j = 1 To cc
        res(i, j) = arr(i - 1)(j - 1)
    Next j
Next i

comp = res
End Function
Public Function rv(range_name As String, Optional sheet As Worksheet) As Variant
If IsMissing(sheet) Or IsEmpty(sheet) Or sheet Is Nothing Then
    rv = Range(range_name).Cells.value
Else
    rv = sheet.Range(range_name).Cells.value
End If

rv = decomp(rv)

End Function

Public Function pt(ParamArray outputs() As Variant)
Dim item As Variant
Dim s As String: s = ""
For Each item In outputs
    s = s & item & " "
Next item
s = left(s, Len(s) - 1)
Debug.Print (s)
End Function




Function substring(s As String, start As Long, finish As Long) '[start, end)
substring = Mid(s, start + 1, finish - start)
End Function
Function charAt(s As String, position As Long)
charAt = substring(s, position, position + 1)
End Function

Function st(ByRef obj As Variant, ByVal item As Variant)
If IsObject(item) Then
        Set obj = item
    Else
        obj = item
    End If
End Function

Function nt(ByVal n As Variant)
' does nothing
End Function

Public Function assert()

End Function


