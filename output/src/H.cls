Attribute VB_Name = "H"
private hadError as boolean
public function init() 
    nt st (hadError, False)
end function


 function removeExt(byval filename as string) as string
    dim  dotPos as long: st dotPos, instrrev(filename, ".")
    if (dotPos > 0) then 
        st removeExt, left(filename, dotPos - 1)
        exit function
    endif
    st removeExt, filename
    exit function
end function


public function compile(byval filepath as string) 
    dim  fso as FileSystemObject: st fso, new FileSystemObject
    if (fso.FileExists(filepath)) then 
        call runFile(filepath)
    else
        if (fso.FolderExists(filepath)) then 
            dim  file as variant: st file, dir(filepath + "\*.*")
            dim  nextfile as variant
            while (file <> "")
                nt st (nextfile, dir)
                call runFile(filepath + "\" + file)
                nt st (file, nextfile)
                looplabel0bef:
                nt 1
            wend
            looplabel0aft:
        else
        endif
    endif
end function


public function runFile(byval filename as string) 
    dim  fso as FileSystemObject: st fso, new FileSystemObject
    dim  ts as TextStream: st ts, fso.OpenTextFile(filename)
    dim  rawcode as string: st rawcode, ts.ReadAll
    nt st (rawcode, replace(rawcode, vbcrlf, vblf))
    call runCode(rawcode, removeExt(dir(filename)))
     ts.close 
end function


static function error(byval line as long, byval message as string) 
    call report(line, "", message)
end function


private function report(byval line as long, byval where as string, byval msg as string) 
    call pt("[line", line, "] Error", where, ": ", msg)
    nt st (hadError, True)
end function


static function err(byval tk as Token, byval message as string) 
    if (tk.m_type_ = EOF_) then 
        call report(tk.m_line, " at end.", message)
    else
        call report(tk.m_line, " at '" + tk.m_lexeme + "'", message)
    endif
end function


private function runCode(byval source as string, byval filename as string) 
    dim  s as Scanner: st s, new Scanner
    dim  tk as variant
    dim  tokens() as variant
    call s.init(source, me)
    nt st (tokens, s.scanTokens())
    dim  parser as Parser: st parser, new parser
    call parser.init(tokens, me)
    dim  statements as variant: st statements, parser.parse()
    if (hadError) then 
        call pt("ERROR")
        exit function
    endif
    dim  astp as AstPrinter: st astp, new AstPrinter
    call astp.prt(statements)
    dim  gen as CodeGenerator: st gen, new CodeGenerator
    call gen.generate(statements, filename)
    exit function
end function



