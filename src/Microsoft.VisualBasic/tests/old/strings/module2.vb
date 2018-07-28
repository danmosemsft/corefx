Option Compare Text
Option Strict On

Imports System
Imports Microsoft.VisualBasic
Imports System.Globalization

Module StringsSuite2

    Dim b As Boolean
    Dim i As Integer
    Dim s As String
    Dim TestNumber As Integer
    Dim TestCounter As Integer
    Dim ch As Char
    Dim m_lBugID As Long



    Sub InstrTextTests()
        NewApiTest("InStr tests (Option Compare Text)")

        Try
            Console.Write("1) ")
            PassFail(InStr(1, "test", "TEST", CompareMethod.Text) = 1)

            Console.Write("2) ") 
            PassFail(InStr(1, "test", "TEST", CompareMethod.Binary) = 0)

            Console.Write("3) ") 
            PassFail(InStr("test", "test") = 1)

            Console.Write("4) ")
            PassFail(InStr(1, "test", "test") = 1)

            Console.Write("5) ")
            PassFail(InStr(1, "test", "test", CompareMethod.Binary) = 1)

            Console.Write("6) ")
            PassFail(InStr(1, "test", "TEST") = 1)

            Console.Write("7) ")
            PassFail(InStr(1, "test", "TEST", CompareMethod.Binary) = 0)

            Console.Write("8) ")
            PassFail(InStr(3, "test", "") = 3)

            'UNDONE: fix ligature handling
            '            s = ChrW(&H9C)  'oe Ligature
            '            Console.Write("9) ")
            '            PassFail(InStr(1, "abcoed", s, CompareMethod.Text) = 4)

            Console.Write("10) ")
            PassFail(InStr("A", "") = 1 AndAlso InStr("A", Nothing) = 1)


            'BUGBUG 304152 - Regressed in NDP 3305
            'Console.Write("11) ")
            'PassFail(InStr("ABCD", ChrW(0)) = 0)

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub


    Sub TestLikeText()
        Dim i As Integer

    	NewApiTest("Like tests (Option Compare Text)")
	Test("testtest", "Testtest", True)
        Test("abc4", "a[a-c]?#", True)
        Test("a", "[b]", False)
        Test( "a", "[][a]", True)
        Test( "abc", "*", True)
        Test( "abc", "a*", True)
        Test( "abc", "abc*", True)
        Test( "abc", "abcc", False)
        Test( "a", "[!b]", True)
        Test( "a", "!b", False)
        Test( "!", "!", True)
        Test( "!", "[!]", True)
        Test( "!", "[!!]", False)
        Test( "-", "-", True)
        Test( "abd", "!bbc", False)
        Test( "b", "[a-c]", True)
        Test( "b", "[b-b]", True)
        Test( "b", "[!a-c]", False)
        Test( "d", "[!a-c]", True)
        Test( "d", "![a-c]", False)
        Test( "0", "[abc0-9]", True)
        Test( "b", "![a-c]", False)
        Test( "t", "[abcd-gxyz]", False)
        TestNonPrintable( ChrW(90), "[" + ChrW(0) + "-" + ChrW(255) + "]", True, "90 [0-255]")
        TestNonPrintable( ChrW(122), "[" + ChrW(0) + "-" + ChrW(255) + "]", True, "122 [0-255]")
        Test("-[?*#!at", "[-][[][?][*][#][!][a][a-z]", True)
        Test("[?*#!at", "[[][?][*][#][!][b][a-z]", False)
        Test( "aabc", "*c", True)
        Test( "abcdef-", "*[a-c][!abc]?f[-]", True)
        Test( "abc", "****", True)
        Test( "aaa.bbb.ccc", "*.ccc", True)
        Test( "a.c", "*.*", True)
        Test( "aaaaaaabcdddddeeeabcdffee", "*[a-z]bcd*ee", True)
        Test( "aaaaaabbccccdaaaaaaaaabbcdd", "*bb*d", True)
        Test( "abbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb", "*a*", True)
        Test( "abbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb", "*c*", False)
        Test( "", "", True)
        Test( "", "*", True)
        Test( "", "[ab", False)
        Test( "[", "[!-]", True)
        Test( "--", "-?", True)
        Test( "---", "--?", True)
        Test( "----", "---?", True)
        TestException("c", "[c-a]")
        TestException("c", "[a--c]")
        TestException("c", "[abc")

        'TRUN generates an access violation if I try Call Test ... here.
        Console.WriteLine("Source=50000 x's & a, Pattern=x*, Actual=" & ((StrDup(50000,"x") & "a") Like "x*") )

        'Can't pass Nulls to Test
        'Console.WriteLine("DBNull Pattern returns: " & ("bob" Like VbVarType.DBNull) )
        'Console.WriteLine("DBNull Source returns: " & (VbVarType.DBNull Like "bob") )
        'Console.WriteLine("DBNull Source with illegal pattern returns: " & (VbVarType.DBNull Like "[ab") )

        'RAID 58311
        Test( "?", "[!?]", False)
        Test( "#", "[!#]", False)
        Test( "*", "[!*]", False)
        Test( "-", "[!-]", False)
        Test( "[", "[![]", False)

        Test( "?", "[!#]", True)
        Test( "!", "[!?]", True)
        Test( "#", "[!*]", True)
        Test( "*", "[!#]", True)
        Test( "-", "[![]", True)
        Test( "[", "[!-]", True)

        'RAID 68130
        Test( "--", "[--]", False)
        Test( "---", "[---]", False)

        'RAID 230305
        Test( "fu", "*????", False)
        Test( "foo.txt", "*[;-?[* ]*", False)
        Test( "foo.txt", "*]*", False)

        'RAID 230255
        For i = 0 to 127
            TestNonPrintable( Chr(i), "[!" + Chr(0) + "-" + Chr(255) + "]", False, i & " [!0-255]")
        Next i   

        TestNonPrintable( Chr(128), "[!" + Chr(0) + "-" + Chr(255) + "]", True, i & " [!0-255]")
        TestNonPrintable( Chr(129), "[!" + Chr(0) + "-" + Chr(255) + "]", False, i & " [!0-255]")

        For i = 130 to 140
            TestNonPrintable( Chr(i), "[!" + Chr(0) + "-" + Chr(255) + "]", True, i & " [!0-255]")
        Next i   

        TestNonPrintable( Chr(141), "[!" + Chr(0) + "-" + Chr(255) + "]", False, i & " [!0-255]")
        TestNonPrintable( Chr(142), "[!" + Chr(0) + "-" + Chr(255) + "]", True, i & " [!0-255]")
        TestNonPrintable( Chr(143), "[!" + Chr(0) + "-" + Chr(255) + "]", False, i & " [!0-255]")
        TestNonPrintable( Chr(144), "[!" + Chr(0) + "-" + Chr(255) + "]", False, i & " [!0-255]")

        For i = 145 to 156
            TestNonPrintable( Chr(i), "[!" + Chr(0) + "-" + Chr(255) + "]", True, i & " [!0-255]")
        Next i  
         
        TestNonPrintable( Chr(157), "[!" + Chr(0) + "-" + Chr(255) + "]", False, i & " [!0-255]")
        TestNonPrintable( Chr(158), "[!" + Chr(0) + "-" + Chr(255) + "]", True, i & " [!0-255]")
        TestNonPrintable( Chr(159), "[!" + Chr(0) + "-" + Chr(255) + "]", True, i & " [!0-255]")

        For i = 160 to 255
            TestNonPrintable( Chr(i), "[!" + Chr(0) + "-" + Chr(255) + "]", False, i & " [!0-255]")
        Next i   
    End Sub


    Sub Test( ByVal Source As String, ByVal Pattern As String, ByVal Expected As Boolean)
        Dim Actual As Boolean
        Actual = Source Like Pattern
        If  Actual = Expected Then
            Console.WriteLine("Source=" & Source & " Pattern=" & Pattern &" " & Actual &" passed" )
        Else
            Console.WriteLine("Source=" & Source & " Pattern=" & Pattern & " " & Actual &" FAILED" )
        End If
    End Sub

End Module
