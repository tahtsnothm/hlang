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

Static Function val(ByVal i As Long, ByVal j As Long, Optional ByVal s As String) As Variant
Dim def As String
If s <> "" Then
def = s
End If
val = Worksheets(def).Cells(i, j).value
'x then y

End Function
Static Function valb(ByVal i As Long, ByVal j As Long, wb As String, Optional ByVal s As String) As Variant
Dim def As String
If s <> "" Then
def = s
End If
valb = Workbooks(wb).Worksheets(def).Cells(i, j).value
'x then y

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

Public Sub arad(ByRef arr As Variant, ByRef item As Variant)
    ReDim Preserve arr(0 To al(arr))
    If IsObject(item) Then
        Set arr(al(arr) - 1) = item
    Else
        arr(al(arr) - 1) = item
    End If
End Sub

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

Public Function tdrtad(rng As Range, Optional ByRef headerarray As Variant) As Variant
'2d range to array of dictionaries
If rng.rows.Count < 2 Then
    Exit Function
End If
Dim i As Long
i = 2
'second row onwards
Dim headers As Variant
headers = rta(rng.rows(1).Cells, 0) ' is a row a range

Dim res() As Variant
Dim currow As Dictionary
For i = 2 To rng.rows.Count
    arad res, rtd(rng.rows(i).Cells, headers)
Next i

If Not IsMissing(headerarray) Then
    headerarray = rta(rng.rows(1).Cells, 0)
End If
tdrtad = res
End Function

Public Function tdrtadt(rng As Range, Optional ByRef headerarray As Variant) As Variant
'TRANSPOSED : left column as key
'2d range to array of dictionaries
If rng.columns.Count < 2 Then
    Exit Function
End If
Dim i As Long
i = 2
'second row onwards
Dim headers As Variant
headers = rta(rng.columns(1).Cells, 0) ' is a row a range

Dim res() As Variant
Dim currow As Dictionary
For i = 2 To rng.columns.Count
    arad res, rtd(rng.columns(i).Cells, headers)
Next i

If Not IsMissing(headerarray) Then
    headerarray = rta(rng.columns(1).Cells, 0)
End If
tdrtadt = res
End Function

Public Function adttda(ad As Variant, Optional ByRef headerarray As Variant)
    Dim i As Long
    Dim res() As Variant
    For i = LBound(ad) To UBound(ad)
        arad res, dta(ad(i))
    Next i
    If Not IsMissing(headerarray) Then
        headerarray = ad(0).Keys
    End If
    
    adttda = res
End Function
Public Function adttdah(ad As Variant) ' headers
    Dim i As Long
    Dim res() As Variant
    i = 0
    res = ad(0).Keys
    adttdah = res
End Function

Public Sub atr(arr As Variant, cell As String, sheet As Variant, Optional down As Boolean = True) ' hori/verti
Dim r As Integer, c As Integer, i As Integer
With Range(cell)
r = .row
c = .column

End With
Dim she As Worksheet
If TypeName(sheet) = "Worksheet" Then
    Set she = sheet
Else
    Set she = Worksheets(sheet)
End If


Dim item As Variant
i = 0
For i = 0 To al(arr) - 1
item = arr(i)

If down Then
    she.Cells(r + i, c).value = item
Else
    she.Cells(r, c + i).value = item
End If

Next i

End Sub

Public Sub tdattdr(tda As Variant, cell As String, sheet As Variant)
Dim row As Variant
Dim r As Integer, c As Integer, i As Integer, j As Integer
With Range(cell)
r = .row
c = .column
End With

Dim she As Worksheet
If TypeName(sheet) = "Worksheet" Then
    Set she = sheet
Else
    Set she = Worksheets(sheet)
End If

i = 0

For i = 0 To al(tda) - 1
    For j = 0 To al(tda(i)) - 1
        she.Cells(r + i, c + j).value = tda(i)(j)
        
    Next j
Next i
End Sub

Public Sub sheetSplit(column As String, tdr As Range)
Dim pk, ad, ar, headers As Variant
ad = tdrtad(tdr, headers)

Dim resDict As New Scripting.Dictionary
Dim row As Scripting.Dictionary
Dim i As Long
For i = 0 To al(ad) - 1
    pk = ad(i)(column)
    If Not resDict.Exists(pk) Then
        Dim temparray() As Variant
        resDict.Add pk, aradv(temparray, dta(ad(i), headers, 0))
    Else
        resDict(pk) = aradv(resDict(pk), dta(ad(i), headers, 0))
    End If
Next i

Dim sheet As Worksheet
For Each pk In resDict.Keys
    Set sheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    sheet.name = CStr(pk)
    atr headers, "A1", sheet, False
    
    tdattdr resDict(pk), "A2", sheet
Next pk

End Sub
Static Function putst(id As Variant, Target As Variant, value As Variant, Optional ByVal sheet As String, Optional idcol As Variant)
'puts t (Sheet)
' will scan through the first col to find id and place value in the target column

Dim deff As String
If sheet <> "" Then
deff = sheet
End If


Dim tabb As Worksheet

Set tabb = Worksheets(deff)
'primary key a1
Dim idcolnum As Long
idcolnum = 1
If Not IsMissing(idcol) Then
    idcolnum = Application.match(idcol, tabb.rows(1).Cells, 0)
End If

Dim persrow As Long
persrow = (Application.match(id, tabb.columns(idcolnum).Cells, 0))
Dim perscol As Long
perscol = CLng(Application.match(CStr(Target), tabb.rows(1).Cells, 0)) 'might cause issues

puts persrow, perscol, value, deff
putst = value
End Function

Static Function valt(id As Variant, Target As Variant, Optional ByVal sheet As String, Optional idcol As Variant) As Variant
Dim tabb As Worksheet

Dim deff As String
If sheet <> "" Then
deff = sheet
End If
Dim idcolnum As Long
idcolnum = 1
Set tabb = Worksheets(deff)
If Not IsMissing(idcol) Then
    idcolnum = Application.match(CStr(idcol), tabb.rows(1).Cells, 0)
End If
id = (id)

'primary key a1
On Error GoTo notfound:
Dim persrow As Long
persrow = (Application.match((id), tabb.columns(idcolnum).Cells, 0))
Dim perscol As Long
perscol = (Application.match(CStr(Target), tabb.rows(1).Cells, 0)) 'might cause issues


GoTo found

notfound:
valt = ""
Exit Function

found:
valt = val(persrow, perscol, deff)

End Function

Static Function valtb(id As Variant, Target As Variant, wb As String, Optional ByVal sheet As String, Optional idcol As Variant) As Variant
Dim tabb As Worksheet

Dim deff As String
If sheet <> "" Then
deff = sheet
End If
Dim idcolnum As Long
idcolnum = 1
Set tabb = Workbooks(wb).Worksheets(deff)
If Not IsMissing(idcol) Then
    idcolnum = Application.match(CStr(idcol), tabb.rows(1).Cells, 0)
End If
id = (id)

'primary key a1
On Error GoTo notfound:
Dim persrow As Long
persrow = (Application.match((id), tabb.columns(idcolnum).Cells, 0))
Dim perscol As Long
perscol = (Application.match(CStr(Target), tabb.rows(1).Cells, 0)) 'might cause issues


GoTo found

notfound:
valtb = ""
Exit Function

found:
valtb = valb(persrow, perscol, wb, deff)

End Function

Static Function pop(ByRef arr As Variant, ByVal index As Long)
Dim arrlength As Long
arrlength = al(arr)

If arrlength = 0 Then
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

Public Function wordReplace(targetfile As String, ByVal replacements As Object, Optional ByRef format As func)
Dim rdict As New Scripting.Dictionary
updatePaths
Set rdict = replacements

Dim book As Word.Application
Dim sheet As Word.Document
Set book = CreateObject("word.application")
book.Visible = True
Set sheet = book.Documents.Open(targetfile)

Dim i As Long

Dim key As Variant
For Each key In rdict.Keys()
    If key = "" Or rdict(key) = "" Then
        GoTo continue
    End If
    
    
    If Len(rdict(key)) > 250 Then
    Dim orng As Word.Range
    Dim sel As Word.Selection
    Set orng = sheet.Content
    With orng.Find
        .text = key
        While .Execute
            orng.Select
            book.Selection.TypeText rdict(key)
            orng.Collapse wdCollapseEnd
        Wend
    End With
    
    Else
    
        With sheet.Content.Find
        .text = key
        Dim value As Variant
        value = rdict(key)
        
        If IsMissing(format) Or format Is Nothing Then
            If IsDate(value) And key <> "{ldsfull}" Then
            .Replacement.text = textDate(CDate(value))
            Else
            .Replacement.text = value
            End If
        Else
            .Replacement.text = format.run(key, value)
        End If
         .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        
        End With
    End If
    
    sheet.Content.Find.Execute Replace:=wdReplaceAll
    
continue:

Next key
sheet.Close savechanges:=True
book.Quit

End Function
Public Function wordReplaceT(targetfile As String, ByVal replacements As Object, Optional ByRef format As func)
Dim rdict As Scripting.Dictionary

Set rdict = replacements

Dim book As Word.Application
Dim sheet As Word.Document
Set book = CreateObject("word.application")
book.Visible = True
Set sheet = book.Documents.Open(targetfile)

Dim i As Long

Dim key As Variant
For Each key In rdict.Keys()
    If key = "" Or rdict(key) = "" Then
        GoTo continue
    End If
    With sheet.Content.Find
    .text = key
    Dim value As Variant
    value = rdict(key)
    
    If IsMissing(format) Or format Is Nothing Then
        If IsDate(value) Then
        .Replacement.text = textDate(value)
        Else
        .Replacement.text = value
        End If
    Else
        .Replacement.text = format.run(key, value)
    End If
     .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
    
    
    End With
    sheet.Content.Find.Execute Replace:=wdReplaceAll
    
    
    Dim shp As Object
    For Each shp In sheet.Shapes
        With shp
        If shp.Type = msoTextBox Then
        If .TextFrame.HasText Then
        With shp.TextFrame.TextRange.Find
        .text = key
    If IsMissing(func) Then
        If IsDate(value) Then
        .Replacement.text = textDate(value)
        Else
        .Replacement.text = value
        End If
    Else
        .Replacement.text = format.run(key, value)
    End If
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    
            End With
            shp.TextFrame.TextRange.Find.Execute Replace:=wdReplaceAll
        End If
        End If
        
        End With
    Next
    
continue:

Next key
sheet.Close savechanges:=True
book.Quit

End Function


Public Function strreplace(ByVal text As String, ByVal replacements As Object)
Dim key As Variant

For Each key In replacements
    text = Replace(text, CStr(key), replacements(key))
Next key

strreplace = (text)

End Function

Public Function draft(ByVal person As String, ByVal cc As String, ByVal subject As String, ByVal body As String, Optional folder As String = "OUT", Optional ByVal replacements As Object)
Dim objoutlook As Outlook.Application
Dim objmail As Outlook.MailItem
Set objoutlook = Outlook.Application
Set objmail = objoutlook.CreateItem(olMailItem)
'
objmail.BodyFormat = olFormatPlain
'change format for html
objmail.to = person
objmail.cc = cc

If IsMissing(replacements) Then
objmail.subject = subject
objmail.HTMLBody = body
Else
objmail.subject = strreplace(subject, replacements)
objmail.HTMLBody = strreplace(body, replacements)
End If


'objmail.body = body


' ... & HTMLBody for sig
objmail.Move (objoutlook.GetNamespace("MAPI").GetDefaultFolder(6).folders(folder))
objmail.Close (olSave)
Set objmail = Nothing
End Function

Public Function drafth(ByVal person As String, ByVal cc As String, ByVal subject As String, ByVal body As String, Optional folder As String = "OUT", Optional ByVal replacements As Object)
Dim objoutlook As Outlook.Application
Dim objmail As Outlook.MailItem
Set objoutlook = Outlook.Application
Set objmail = objoutlook.CreateItem(olMailItem)
'
objmail.BodyFormat = olFormatHTML
'change format for html
objmail.to = person
objmail.cc = cc

If IsMissing(replacements) Then
objmail.subject = subject
objmail.HTMLBody = body
Else
objmail.subject = strreplace(subject, replacements)
objmail.HTMLBody = strreplace(body, replacements)
End If


'objmail.body = body


' ... & HTMLBody for sig
objmail.Move (objoutlook.GetNamespace("MAPI").GetDefaultFolder(6).folders(folder))
objmail.Close (olSave)
Set objmail = Nothing
End Function

Function draftatch(ByVal person As String, ByVal cc As String, ByVal subject As String, ByVal body As String, ByVal attc As Variant, Optional folder As String = "OUT", Optional replacements As Object, Optional embeds As Variant)
Dim objoutlook As Outlook.Application
Dim objmail As Outlook.MailItem
Set objoutlook = Outlook.Application
Set objmail = objoutlook.CreateItem(olMailItem)
Dim olkPA As Outlook.PropertyAccessor
Const ATTACH_STUFF = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
Dim oatch As Outlook.attachment

objmail.BodyFormat = olFormatHTML
objmail.to = person
objmail.cc = cc

Dim attcarr() As Variant
Dim embedarr() As Variant

Dim itt As Variant

If IsMissing(embeds) Or IsEmpty(embeds) Then
Else
    If IsArray(embeds) Then
        embedarr = embeds
    Else
        arad embedarr, embeds
    End If
    
    For Each itt In embedarr
    
        Set oatch = objmail.Attachments.Add(itt, 1, 1)
        Set olkPA = oatch.PropertyAccessor
        olkPA.SetProperty ATTACH_STUFF, itt

    Next itt
End If

If IsMissing(attc) Or IsEmpty(attc) Then
Else
    If IsArray(attc) Then
        attcarr = attc
    Else
        arad attcarr, attc
    End If
        For Each itt In attcarr
    
        Set oatch = objmail.Attachments.Add(itt)

    Next itt
End If

If IsMissing(replacements) Then
objmail.subject = subject
objmail.HTMLBody = body
Else
objmail.subject = strreplace(subject, replacements)
objmail.HTMLBody = strreplace(body, replacements)
End If
objmail.Close olSave



objmail.Close (olSave)
objmail.Move (objoutlook.GetNamespace("MAPI").GetDefaultFolder(6).folders(folder))
Set objmail = Nothing
End Function

Function draftatc(ByVal person As String, ByVal cc As String, ByVal subject As String, ByVal body As String, ByVal attc As Variant, Optional folder As String = "OUT", Optional replacements As Object)
Dim objoutlook As Outlook.Application
Dim objmail As Outlook.MailItem
Set objoutlook = Outlook.Application
Set objmail = objoutlook.CreateItem(olMailItem)
Dim olkPA As Outlook.PropertyAccessor
Const ATTACH_STUFF = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
Dim oatch As Outlook.attachment

objmail.BodyFormat = olFormatPlain
objmail.to = person
objmail.cc = cc

Dim attcarr() As Variant
Dim itt As Variant

If IsMissing(attc) Or IsEmpty(attc) Then
Else
    If IsArray(attc) Then
        attcarr = attc
    Else
        arad attcarr, attc
    End If
        For Each itt In attcarr
    
        Set oatch = objmail.Attachments.Add(itt)

    Next itt
End If

If IsMissing(replacements) Then
objmail.subject = subject
objmail.body = body
Else
objmail.subject = strreplace(subject, replacements)
objmail.body = strreplace(body, replacements)
End If
objmail.Close olSave

objmail.Move (objoutlook.GetNamespace("MAPI").GetDefaultFolder(6).folders(folder))
Set objmail = Nothing
End Function

Function sret(persno As Variant, rescrit As Variant) As Variant
'range
Dim ressers() As String
Dim rnge As ListObject
Set rnge = Worksheets("res").ListObjects("res")

Dim resarray() As Variant
Dim m, n As Integer
Dim mm, nn As Integer
m = persno.rows.Count
n = persno.columns.Count

If TypeName(rescrit) = "String" Then
mm = 1
nn = 1
Else

mm = rescrit.rows.Count
nn = rescrit.columns.Count
End If

Dim qm, qn As Long
If m > n Then
'hori
qm = m
qn = Application.WorksheetFunction.Max(mm, nn)
Else
qn = n
qm = Application.WorksheetFunction.Max(mm, nn)
End If



ReDim Preserve resarray(0 To qm - 1, 0 To qn - 1)

Dim lll As Long
lll = 1
'Dim resbig() As String
Dim ee As Long
'ee = UBound(persno) - LBound(persno) + 1
Dim matchres As Long
Dim inside As Variant




Dim r, c As Long
Dim rr, cc As Long
Dim i As Long
Dim j As Long

For i = 0 To qm - 1
    For j = 0 To qn - 1
    
        If TypeName(rescrit) = "String" Then
            matchres = WorksheetFunction.match(persno.rows(i + 1).columns(1).value, rnge.ListColumns("PERNR").DataBodyRange, 0)
            resarray(i, j) = WorksheetFunction.index(rnge.ListColumns(rescrit).DataBodyRange, matchres)
        Else
        If (m > n) Then
            'if horizontal persno
            matchres = WorksheetFunction.match(persno.rows(i + 1).columns(1).value, rnge.ListColumns("PERNR").DataBodyRange, 0)
            resarray(i, j) = WorksheetFunction.index(rnge.ListColumns(rescrit.rows(1).columns(j + 1).value).DataBodyRange, matchres)
        Else
            matchres = WorksheetFunction.match(persno.rows(1).columns(j + 1).value, rnge.ListColumns("PERNR").DataBodyRange, 0)
            resarray(i, j) = WorksheetFunction.index(rnge.ListColumns(rescrit.rows(i + 1).columns(1).value).DataBodyRange, matchres)
        End If
        End If
    Next j
Next i



'For r = 0 To m - 1
    'For c = 0 To n - 1
           'matchres = WorksheetFunction.Match(persno.Rows(r + 1).Columns(c + 1).value, rnge.ListColumns("PERNR").DataBodyRange, 0)
            'resarray(r, c) = WorksheetFunction.index(rnge.ListColumns(rescrit).DataBodyRange, matchres)
    'Next c
'Next r

sret = resarray()



End Function

Function sres(persno As Variant, searchcrit As String, rescrit As Variant) As Variant
'range
Dim ressers() As String
Dim rnge As ListObject
Set rnge = Worksheets("res").ListObjects("res")

Dim resarray() As Variant
Dim m, n As Integer
Dim mm, nn As Integer
m = persno.rows.Count
n = persno.columns.Count

If TypeName(rescrit) = "String" Then
mm = 1
nn = 1
Else

mm = rescrit.rows.Count
nn = rescrit.columns.Count
End If

Dim qm, qn As Long
If m > n Then
'hori
qm = m
qn = Application.WorksheetFunction.Max(mm, nn)
Else
qn = n
qm = Application.WorksheetFunction.Max(mm, nn)
End If



ReDim Preserve resarray(0 To qm - 1, 0 To qn - 1)

Dim lll As Long
lll = 1
'Dim resbig() As String
Dim ee As Long
'ee = UBound(persno) - LBound(persno) + 1
Dim matchres As Long
Dim inside As Variant




Dim r, c As Long
Dim rr, cc As Long
Dim i As Long
Dim j As Long

For i = 0 To qm - 1
    For j = 0 To qn - 1
        If (m > n) Then
            'if horizontal persno
            matchres = WorksheetFunction.match(persno.rows(i + 1).columns(1).value, rnge.ListColumns(searchcrit).DataBodyRange, 0)
            resarray(i, j) = WorksheetFunction.index(rnge.ListColumns(rescrit.rows(1).columns(j + 1).value).DataBodyRange, matchres)
        Else
            matchres = WorksheetFunction.match(persno.rows(1).columns(j + 1).value, rnge.ListColumns(searchcrit).DataBodyRange, 0)
            resarray(i, j) = WorksheetFunction.index(rnge.ListColumns(rescrit.rows(i + 1).columns(1).value).DataBodyRange, matchres)
        End If
    Next j
Next i



'For r = 0 To m - 1
    'For c = 0 To n - 1
           'matchres = WorksheetFunction.Match(persno.Rows(r + 1).Columns(c + 1).value, rnge.ListColumns("PERNR").DataBodyRange, 0)
            'resarray(r, c) = WorksheetFunction.index(rnge.ListColumns(rescrit).DataBodyRange, matchres)
    'Next c
'Next r

sres = resarray()



End Function

Public Function tdatt(ByVal arr As Variant, ByVal table As ListObject)
'preserves header?
Dim rc As Long
Dim cc As Long
rc = al(arr)
cc = al(arr(0))
'to table
Dim tda() As Variant
ReDim tda(0 To rc - 1, 0 To cc - 1)


Dim i As Long
Dim j As Long
For i = 0 To rc - 1
    For j = 0 To cc - 1
        tda(i, j) = arr(i)(j)
    Next j
Next i

table.Resize table.Range.Resize(rc + 1, cc)
table.DataBodyRange.ClearContents ' moved this down
If rc = 1 Then
Dim oda() As Variant
For j = 0 To cc - 1
    arad oda, tda(0, j)
Next j
Dim newrow As ListRow

Set newrow = table.ListRows.Add
newrow.Range.value = oda

Else
table.DataBodyRange.value = tda
End If


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

Function gettbl(ws As String, tblname As String) As ListObject

Set gettbl = Worksheets(ws).ListObjects(tblname)
End Function

Function globalgettbl(ByVal tblname As String) As ListObject
Dim wb As Workbook
Dim ws As Worksheet
Dim tbl As ListObject
For Each wb In Workbooks
    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.name = tblname Then
                Set globalgettbl = tbl
            End If
        Next tbl
    Next ws
Next wb
End Function
Function unwrap(ByVal arr As Variant, ByVal values As Variant) As Variant
Dim value As Variant
Dim result() As Variant
For Each value In values
    arad result, arr(value)
Next value
unwrap = result
End Function


Function artt(ByRef tbl As ListObject, ByVal row As Variant)
'addrowtotable
    Dim headers As Variant

    headers = rta(tbl.HeaderRowRange.Cells)
    Dim newrowrange() As Variant
    Dim i As Long
    Dim ii As Integer
    Dim sheet As ListObject
    Set sheet = tbl
    If TypeName(row) = "Dictionary" Then
         newrowrange = dta(row, headers)
    
    ElseIf TypeName(row) = "Range" Then
        ReDim newrowrange(1 To row.Cells.Count)
        For i = 1 To row.Cells.Count
            newrowrange(i) = row
        Next i
         newrowrange = dta(row, headers)
    Else
        If (al(headers) <> UBound(row) - LBound(row) + 1) Then
            'error
            Exit Function
        Else
            ReDim newrowrange(1 To UBound(row) - LBound(row) + 1)
            ii = 1
            For i = LBound(row) To UBound(row)
                newrowrange(i + 1) = row(i)
            Next i
        End If
    End If
    Dim newrow As ListRow
        Set newrow = sheet.ListRows.Add()
        Dim rheight As Long: rheight = newrow.Range.RowHeight
        newrow.Range.Formula = newrowrange
        newrow.Range.RowHeight = rheight
        
        
        'With newrow
                'For ii = 1 To al(headers)
                '''.Range(ii) = rawrow.Cells(ii)
                '.Range(ii) = newrowrange(ii)
            'Next ii
        'End With
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

Public Function hstack(ParamArray columns() As Variant) As Variant 'returns a rows x sum(cols) array ( rows must be the same)
'TODO: make it work for arbitrarily thick matrices
Dim res() As Variant
Dim i As Long
Dim j As Long
Dim k As Long
Dim rc As Long
Dim cc As Long ' cc is the number of arguments
'row count and column count

If (al(columns()) = 0) Then
    Exit Function ' at least one column
End If



rc = columns(0).Cells.rows.Count ' row count
cc = al(columns())
Dim temparray() As Variant
For i = 0 To rc - 1
    temparray = Array()
    For j = 0 To cc - 1
        For k = 1 To columns(j).columns.Count
            arad temparray, columns(j).rows(i + 1).Cells(k).value
        Next k
    Next j
    res = aradv(res, temparray)
Next i

hstack = res

End Function

Public Function vstack(ParamArray rows() As Variant) As Variant 'returns a rows x sum(cols) array ( rows must be the same)
'TODO: make it work for arbitrarily thick matrices
Dim res() As Variant
Dim i As Long
Dim j As Long
Dim k As Long
Dim rc As Long
Dim cc As Long ' cc is the number of arguments
'row count and column count

If (al(rows()) = 0) Then
    Exit Function ' at least one column
End If



cc = rows(0).Cells.columns.Count ' row count
rc = al(rows())
Dim temparray() As Variant


    
    For i = 0 To rc - 1
        For k = 1 To rows(j).rows.Count
            temparray = Array()
            For j = 0 To cc - 1
            arad temparray, rows(i).rows(j + 1).Cells(k).value
        Next j
            res = aradv(res, temparray)
    Next k
    
    Next i

vstack = res

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

Sub tt()
Dim a(1 To 2) As Variant
pt al(a), al(a, 1), al(a, 2), "flkdfhjskl"

End Sub

Function addOne(rng As Variant)
Dim i, j As Long
Dim rc As Long, cc As Long
decomp rng, rc, cc
For i = 0 To rc
    For j = 0 To cc
        rng(i)(j) = m_addOne(rng(i)(j))
    Next j
Next i

addOne = comp(rng)

End Function
Function m_addOne(ByVal number As Long)
m_addOne = number + 1
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
