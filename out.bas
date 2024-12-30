public sub main ()
    pt fibonacci(10)
end sub

function fibonacci(byval n as integer) as integer
    if (n <= 1) then 
        st fibonacci, 1
        exit function
    endif
    st fibonacci, fibonacci(n - 1) + fibonacci(n - 2)
    exit function
end function


