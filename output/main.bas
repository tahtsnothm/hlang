
Attribute VB_Name = "main"
public sub main ()
end sub

 function hcc(byval filename as string) 
    dim  hh as H: st hh, new H
    call hh.init()
    call hh.compile(ThisWorkbook.Path + "\" + filename)
end function


