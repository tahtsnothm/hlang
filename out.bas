public sub main ()
    if (funcl5c2()) then 
        pt ("eligible")
    endif
end sub

private function funcl5c0() as Boolean
    funcl5c0 = false
    if (a.rank == "COL") then
        if (a.cep >= "LTC") then
            funcl5c0 = true
        endif
    endif
end function
private function funcl5c1() as Boolean
    funcl5c1 = false
    if (person != Nothing) then
        if ((funcl5c0())) then
            funcl5c1 = true
        endif
    endif
end function
private function funcl5c2() as Boolean
    funcl5c2 = false
    if (funcl5c1()) then
        if (not a.retiring) then
            funcl5c2 = true
        endif
    endif
end function
