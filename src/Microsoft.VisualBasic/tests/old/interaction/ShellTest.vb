'ShellTest.vb

imports Microsoft.VisualBasic

Module Test

sub main()
    FileOpen(1, "vbcrun-ShellTest.out", OpenMode.Output )
    Print(1, "Shell test passed" & Chr(13) & Chr(10))
    FileClose(1)
end sub

end Module