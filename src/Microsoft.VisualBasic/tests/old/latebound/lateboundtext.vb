option compare text
option strict off

Imports Microsoft.VisualBasic
Imports System
Imports System.Console

Public Module LateboundCompareTextTest

    Sub TestLateBoundCompareText()

        Dim o1 as Object
        Dim o2 as Object 
     
        Console.WriteLine("*** TestLateboundCompareText ***")

        Console.WriteLine("Nothing")

        o1 = Nothing
        o2 = Nothing

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, True)
        TestScenario(o1 <> o2, False)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, False)

        Console.WriteLine("integers")

        o1 = CInt(1)
        o2 = CInt(2)

        TestScenario(o1 < o2, true)
        TestScenario(o1 <= o2, true)
        TestScenario(o1 = o2, false)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, false)
        TestScenario(o1 > o2, false)

        o1 = CInt(2)

        TestScenario(o1 < o2, false)
        TestScenario(o1 <= o2, true)
        TestScenario(o1 = o2, true)
        TestScenario(o1 <> o2, False)
        TestScenario(o1 >= o2, true)
        TestScenario(o1 > o2, false)

        o1 = CInt(3)

        TestScenario(o1 < o2, false)
        TestScenario(o1 <= o2, false)
        TestScenario(o1 = o2, false)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, true)
        TestScenario(o1 > o2, true)

        o1 = Nothing

        TestScenario(o1 < o2, True)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, False)
        TestScenario(o1 > o2, False)

        o1 = CInt(3)
        o2 = Nothing

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, False)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, True)

        Console.WriteLine("Strings")

        o1 = "aaaa"
        o2 = "bbbb"

        TestScenario(o1 < o2, True)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, False)
        TestScenario(o1 > o2, False)

        o1 = "bbbb"

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, True)
        TestScenario(o1 <> o2, False)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, False)

        o1 = "cccc"

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, False)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, True)

        o1 = Nothing

        TestScenario(o1 < o2, True)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, False)
        TestScenario(o1 > o2, False)

        o2 = ""

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, True)
        TestScenario(o1 <> o2, False)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, False)

        o1 = "cccc"
        o2 = Nothing

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, False)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, True)

        o1 = ""

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, True)
        TestScenario(o1 <> o2, False)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, False)

        Console.WriteLine("reference cases")

        On Error Resume Next
    
        Dim b as boolean

        o1 = new cls1
        o2 = Nothing

        b = o1 <  o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 <= o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1  = o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 <> o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 >= o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 >  o2 : Console.WriteLine(Err.Description) : Err.Clear

        o2 = new cls1

        b = o1 <  o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 <= o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1  = o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 <> o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 >= o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 >  o2 : Console.WriteLine(Err.Description) : Err.Clear

        o2 = new cls2

        b = o1 <  o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 <= o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1  = o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 <> o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 >= o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 >  o2 : Console.WriteLine(Err.Description) : Err.Clear

        o1 = Nothing

        b = o1 <  o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 <= o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1  = o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 <> o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 >= o2 : Console.WriteLine(Err.Description) : Err.Clear
        b = o1 >  o2 : Console.WriteLine(Err.Description) : Err.Clear

        On Error Goto 0

    End Sub
        
End Module
