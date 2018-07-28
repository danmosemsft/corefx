Option Compare Binary
Option Strict On

Imports System
Imports Microsoft.VisualBasic
Imports System.Globalization

Module StringsSuite

    Dim b As Boolean
    Dim i As Integer
    Dim s As String
    Dim TestNumber As Integer
    Dim TestCounter As Integer
    Dim ch As Char
    Dim m_lBugID As Long



    Sub Main()
        MyCulture.Init(&H409)

        Console.WriteLine("Begin Tests")
	    TestPublicFuncs
        TestLikeBinary
        TestLikeText
        TestDates
        TestErrors
        NewApiTest("Regression Tests")
        Bug304082.Test
        Bug316642.Test
        Console.WriteLine("End Tests")
    End Sub



    Sub TestPublicFuncs()
        AscTests
        ChrTests
        GetCharTests
        FormatTests
        FormatDateTimeTests
        FormatCurrencyTests
        FormatPercentTests
        FormatNumberTests
        InstrBinaryTests
        InStrTextTests
        InstrRevTests
        JoinTests
        SplitTests
        ReplaceTests
        SpaceTests
        LCaseTests
        UCaseTests
        LeftTests
        RightTests
        MidTests
        MidStmtTests
        TrimTests
        StrCompTests
        StrConvTests
        ValTests
        FilterTest
        StrReverseTests
    End Sub


    
    Sub Failed(ByVal ex As Exception)
        If ex Is Nothing Then
            Console.WriteLine("NULL System.Exception")
        Else
            Console.WriteLine(ex.GetType().FullName & ": " & ex.Message)
        End If
        Console.WriteLine("FAILED !!!")
    End Sub



    Sub Failed()
        Console.WriteLine("FAILED !!!")
    End Sub



    Sub Passed()
        Console.WriteLine("passed")
    End Sub



    Sub PrintResult(ByVal bPass As Boolean)
        If bPass Then
            Console.WriteLine("passed")
        Else
            Console.WriteLine("FAILED !!!")
        End If
        m_lBugID = 0
    End Sub



    Sub SetBugID(ByVal lID As Long, Optional ByVal lExceptedErr As Long = 0)
        If m_lBugID <> 0 Then
            Console.WriteLine("fail (unexpected exception)")
        End If
        m_lBugID = lID
        Console.Write("Bug" & CStr(m_lBugID) & ": ")
    End Sub



    Sub PassFail(ByVal bPassed As Boolean)
        If bPassed Then
            Console.WriteLine("passed")
        Else
            Console.WriteLine("FAILED !!!")
        End If
    End Sub


    Sub TestCase(ByVal TestPrefix As String, ByVal Result As String, ByVal Expected As String)
        Console.Write(TestPrefix)
        If Result = Expected Then
            Console.WriteLine("passed")
        Else
            Console.Write("failed : result = '")
            Console.WriteLine(CStr(Result) & "'")
        End If
    End Sub



    Sub NewApiTest(ByVal TestName As String)
        TestCounter = 0
        Console.WriteLine("----------------------------------")
        Console.WriteLine(TestName)
    End Sub



    Sub NewSubTest(ByVal SubTestName As String)
        Console.WriteLine(SubTestName)
        TestCounter += 1
    End Sub



    Sub NamedParamTest()
#if 0
        ' NOTE:  ***** DO NOT MODIFY THIS FUNCTION *****
        '       (unless you really know what you're doing)
        '
        ' This function is intended to just to ensure compilation
        ' If someone changes a name in the runtime, this should break
        ' we will rely on the test team to make sure named params
        ' return results as expected
        '
        Dim v As Object
        Dim av(1) As Object
        Dim o As Object
        Dim s As String

        v = VBA.Asc(String := "string")
        v = VBA.AscW(String := "string")
        v = VBA.ChrW(CharCode := 1)
        v = VBA.ChrW(CharCode := 1)
        v = VBA.Filter(SourceArray := v, Match := "match", Include := True, Compare := CompareMethod.Text)
        'UNDONE: the Format parameter conflicts with the Format function name
        ' when this is fixed, we should enable this test
        'v = VBA.Format(Expression := v, Style := v, FirstDayOfWeek := vbSunday, FirstWeekOfYear := vbFirstJan1)
        'Use the next line for now
        v = VBA.Format(Expression := v, Style := v, FirstDayOfWeek := vbSunday, FirstWeekOfYear := vbFirstJan1)
        v = VBA.Format$(Expression := v, Style := v, FirstDayOfWeek := vbSunday, FirstWeekOfYear := vbFirstJan1)
        v = VBA.FormatCurrency(Expression := v, NumDigitsAfterDecimal := 1, IncludeLeadingDigit := vbUseDefault, _
            UseParensForNegativeNumbers := vbUseDefault, GroupDigits := vbUseDefault)
        v = VBA.FormatDateTime(Expression := v, NamedFormat := vbGeneralDate)
        v = VBA.FormatNumber(Expression := v, NumDigitsAfterDecimal := 1, IncludeLeadingDigit := vbUseDefault, _
            UseParensForNegativeNumbers := vbUseDefault, GroupDigits := vbUseDefault)
        v = VBA.FormatPercent(Expression := v, NumDigitsAfterDecimal := 1, IncludeLeadingDigit := vbUseDefault, _
            UseParensForNegativeNumbers := vbUseDefault, GroupDigits := vbUseDefault)
        v = VBA.InStr(Start := v, String1 := v, String2 := v, Compare := CompareMethod.Text)
        v = VBA.InStrRev(StringCheck := "check", StringMatch := "match", Start := 1, Compare := CompareMethod.Text)
        v = VBA.Join(SourceArray := v, Delimiter := v)
        v = VBA.LCase(String := v)
        v = VBA.LCase$(String := v)
        v = VBA.Left(String := s, Length := 1)
        v = VBA.Left$(String := s, Length := 1)
        v = VBA.LeftB$(String := s, Length := 1)
        v = VBA.Len(Expression := v)
        v = VBA.LTrim(String := s)
        v = VBA.LTrim$(String := s)
        v = VBA.Mid(String := s, Start := 1, Length := 1)
        v = VBA.Mid$(String := s, Start := 1, Length := 1)
        v = VBA.MonthName(Month := 1, Abbreviate := True)
        'UNDONE: Replace parameter name conflicts with function name
        ' when this is fixed, we should enable this test
        'v = VBA.Replace(Expression := v, Find := s, Replace := s, Start := 1, Count := 1, Compare := CompareMethod.Text)
        v = VBA.Right(String := s, Length := 1)
        v = VBA.Right$(String := s, Length := 1)
        v = VBA.RTrim(String := s)
        v = VBA.RTrim$(String := s)
        v = VBA.Space(Number := 1)
        v = VBA.Space$(Number := 1)
        v = VBA.Split(Expression := v, Delimiter := v, Limit := 1, Compare := CompareMethod.Text)
        v = VBA.StrComp(String1 := s, String2 := s, Compare := CompareMethod.Text)
        v = VBA.StrConv(String := s, Conversion := 0, LocaleID := 1)
        v = VBA.String(Number := 1, Character := 1)
        v = VBA.String$(Number := 1, Character := 1)
        v = VBA.StrReverse(Expression := s)
        v = VBA.Trim(String := s)
        v = VBA.Trim$(String := s)
        v = VBA.UCase(String := s)
        v = VBA.UCase$(String := s)
        v = VBA.WeekdayName(Weekday := 1, Abbreviate := True, FirstDayOfWeek := vbSunday)
#end if
    End Sub



    Sub ValTests
        NewApiTest("Val tests")
	    Dim i as Char
	    i = Chr(49)
	    Console.writeline(i & " "& val(i))	'RAID 124981
    End Sub



    Sub AscTests()
        NewApiTest("Asc tests")

        NewSubTest("String is Nothing")
        Try
            Dim sEmpty As String
            i = Asc(sEmpty)
            Failed
        Catch ae As ArgumentException
            Passed
        Catch e As Exception
            Failed(e)
        End Try

        NewSubTest("String is EQ Length Zero")
        Try
            Dim sZero as String = ""
            i = Asc(sZero)
            Failed
        Catch ae As ArgumentException
            Passed
        Catch e As Exception
            Failed(e)
        End Try

        NewSubTest("String is GT Length Zero")
        Try
            PassFail(Asc("ABCD") = 65)
        Catch e As Exception
            Failed(e)
        End Try
    End Sub



    Sub ChrTests()
        NewApiTest("Chr tests")

        NewSubTest("CharCode LT -32768")
        Try
            i = -32769
            s = ChrW(i)
            Failed
        Catch ae As ArgumentException
            Passed
        Catch e As Exception
            Failed(e)
        End Try

        NewSubTest("CharCode GE -32768 And LE 65535")
        Try
            i = 65
            PassFail(ChrW(i) = "A"c)
        Catch e As Exception
            Failed(e)
        End Try
    End Sub



    Sub GetCharTests()
        NewApiTest("GetChar tests")

        NewSubTest("Index < 0")
        Try
            ch = GetChar("ABC", -1)
            Failed
        Catch ae As ArgumentException
            Passed
        Catch e As Exception
            Failed(e)
        End Try

        NewSubTest("Index > String.Length")
        Try
            ch = GetChar("ABC", 4)
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try

        NewSubTest("Valid Index: 2 tests")
        Try
            PassFail(GetChar("ABC", 1) = "A"c)
            PassFail(GetChar("ABC", 3) = "C"c)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub FormatCurrencyTests
        NewApiTest("FormatCurrencyTests tests")
        Try
            Dim dec As Decimal
            Dim ExpectedResult As String

            Console.WriteLine("Test with different combinations with valid expressions")

            'FormatCurrency(    ByVal Expression As Object, 
            '                   Optional ByVal NumDigitsAfterDecimal As Integer = -1, 
            '                   Optional ByVal IncludeLeadingDigit As TriState = TriState.UseDefault, 
            '                   Optional ByVal UseParensForNegativeNumbers As TriState = TriState.UseDefault, 
            '                   Optional ByVal GroupDigits As TriState = TriState.UseDefault) As String

            TestCase("1a) ", FormatCurrency(1234.56@), "$1,234.56")
            TestCase("1b) ", FormatCurrency(-1234.56@), "($1,234.56)")
            TestCase("1c) ", FormatCurrency(1234.56@, 4), "$1,234.5600")
            TestCase("1d) ", FormatCurrency(-1234.56@, 4), "($1,234.5600)")
            TestCase("1e) ", FormatCurrency(1234.56@, 3, TriState.True, TriState.False), "$1,234.560")
            TestCase("1f) ", FormatCurrency(-1234.56@, 3, TriState.True, TriState.False), "-$1,234.560")
            TestCase("1g) ", FormatCurrency(1234.56@, 3, TriState.True, TriState.True, TriState.False), "$1234.560")
            TestCase("1h) ", FormatCurrency(-1234.56@, 3, TriState.True, TriState.True, TriState.False), "($1234.560)")
            TestCase("1i) ", FormatCurrency(.56@, -1, TriState.False, TriState.False), "$.56")
            TestCase("1j) ", FormatCurrency(-.56@, -1, TriState.False, TriState.False), "-$.56")
            TestCase("1k) ", FormatCurrency(.56@, -1, TriState.False, TriState.True), "$.56")
            TestCase("1l) ", FormatCurrency(-.56@, -1, TriState.False, TriState.True), "($.56)")

        Catch ex As Exception
            Console.WriteLine
            Failed(ex)
        End Try
    End Sub



    Sub FormatPercentTests
        NewApiTest("FormatPercentTests tests")
        Dim dec As Decimal
        Dim sng As Single
        Dim dbl As Double
        Dim i As Integer
        Dim l As Long
        Dim sh As Short
        Dim ch As Char
        Dim s As String
        Dim byt As Byte
        Dim dt As Date

        dec = 12.3456@
        sng = 12.3456
        dbl = 12.3456
        i = 123
        l = 123
        sh = 123
        ch = ChrW(12)
        byt = 123
        dt = #1/1/2001#
        s = "12.3456"

        Try
            TestCase("1a) ", FormatPercent(sh), "12,300.00%")
            TestCase("1b) ", FormatPercent(i), "12,300.00%")
            TestCase("1c) ", FormatPercent(l), "12,300.00%")
            TestCase("1d) ", FormatPercent(dec), "1,234.56%")
            TestCase("1e) ", FormatPercent(sng), "1,234.56%")
            TestCase("1f) ", FormatPercent(dbl), "1,234.56%")
            TestCase("1g) ", FormatPercent(s), "1,234.56%")
            TestCase("1h) ", FormatPercent(byt), "12,300.00%")
            'TestCase("1i) ", FormatPercent(ch), "1,234.56%")
        Catch ex As Exception
            Console.WriteLine
            Failed(ex)
        End Try
    End Sub



    Sub FormatNumberTests
        NewApiTest("FormatNumberTests tests")

        'Test of the following API
        '        FormatNumber(ByVal Expression As Object, _
        '                    Optional ByVal NumDigitsAfterDecimal As Integer = -1, 
        '                    Optional ByVal IncludeLeadingDigit As TriState = TriState.UseDefault, 
        '                    Optional ByVal UseParensForNegativeNumbers As TriState = TriState.UseDefault, 
        '                    Optional ByVal GroupDigits As TriState = TriState.UseDefault) As String
        Dim dec As Decimal
        Dim sng As Single
        Dim dbl As Double
        Dim i As Integer
        Dim l As Long
        Dim sh As Short
        Dim ch As Char
        Dim s As String
        Dim byt As Byte
        Dim dt As Date

        dec = 12345.67@
        sng = 12345.67
        dbl = 12345.67
        i = 12345
        l = 12345
        sh = 12345
        ch = ChrW(12345)
        byt = 123
        dt = #1/1/2001#
        s = "12345.67"

        Try
            TestCase("1a) ", FormatNumber(sh), "12,345.00")
            TestCase("1b) ", FormatNumber(i), "12,345.00")
            TestCase("1c) ", FormatNumber(l), "12,345.00")
            TestCase("1d) ", FormatNumber(dec), "12,345.67")
            TestCase("1e) ", FormatNumber(sng), "12,345.67")
            TestCase("1f) ", FormatNumber(dbl), "12,345.67")
            TestCase("1g) ", FormatNumber(s), "12,345.67")
            TestCase("1h) ", FormatNumber(byt), "123.00")
            TestCase("1i) ", FormatNumber(True), "-1.00")
            TestCase("1j) ", FormatNumber(False), "0.00")
            TestCase("1k) ", FormatNumber(45, 5), "45.00000")
            'TestCase("1m) ", FormatNumber(ch), "")
        Catch ex As Exception
            Console.WriteLine
            Failed(ex)
        End Try

        Try
            TestCase("1l) ", FormatNumber(45, -234), "45")
            Failed
        Catch ex As ArgumentException
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub FormatDateTimeTests()
        NewApiTest("FormatDateTime tests")

        Try
            Dim dt As DateTime
            Dim DateOnly As Date = Today
            Dim TimeOnly As Date = New DateTime(1,1,1, 12, 15, 39)

            Dim ExpectedResult As String

            dt = New DateTime(2000, 12, 25, 12,15,39, 0)

            Console.WriteLine("Test with date and time")
            TestCase("1a) ", FormatDateTime(dt, DateFormat.GeneralDate), "12/25/2000 12:15:39 PM")
            TestCase("1b) ", FormatDateTime(dt, DateFormat.ShortDate), "12/25/2000")
            TestCase("1c) ", FormatDateTime(dt, DateFormat.LongDate), "Monday, December 25, 2000")
            TestCase("1d) ", FormatDateTime(dt, DateFormat.ShortTime), "12:15")
            TestCase("1e) ", FormatDateTime(dt, DateFormat.LongTime), "12:15:39 PM")
            TestCase("1f) ", CStr(dt), "12/25/2000 12:15:39 PM")

            Console.WriteLine("Test with date only")
            dt = New DateTime(2000,12,25, 0, 0, 0, 0)
            TestCase("2a) ", FormatDateTime(dt, DateFormat.GeneralDate), "12/25/2000")
            TestCase("2b) ", FormatDateTime(dt, DateFormat.ShortDate), "12/25/2000")
            TestCase("2c) ", FormatDateTime(dt, DateFormat.LongDate), "Monday, December 25, 2000")
            TestCase("2d) ", FormatDateTime(dt, DateFormat.ShortTime), "00:00")
            TestCase("2e) ", FormatDateTime(dt, DateFormat.LongTime), "12:00:00 AM")
            TestCase("2f) ", CStr(dt), "12/25/2000")

            Console.WriteLine("Test with time only")
            dt = New DateTime(1,1,1, 12, 15, 39, 0)
            TestCase("3a) ", FormatDateTime(dt, DateFormat.GeneralDate), "12:15:39 PM")
            TestCase("3b) ", FormatDateTime(dt, DateFormat.ShortDate), "1/1/0001") 'DIFFERENT from VB6
            TestCase("3c) ", FormatDateTime(dt, DateFormat.LongDate), "Monday, January 01, 0001") 'DIFFERENT from VB6
            TestCase("3d) ", FormatDateTime(dt, DateFormat.ShortTime), "12:15")
            TestCase("3e) ", FormatDateTime(dt, DateFormat.LongTime), "12:15:39 PM")
            TestCase("3f) ", CStr(dt), "12:15:39 PM")

        Catch ex As Exception
            Console.WriteLine
            Failed(ex)
        End Try
    End Sub



    Sub FormatTests()
        NewApiTest("Format tests")

        Try
            Console.WriteLine("Testing named formats")

            Console.WriteLine("Test with lower case")
            Console.Write("1) ")  
            PassFail("4.6" = Format(1.2 + 3.4, "general number"))

            Console.Write("2) ")  
            PassFail("12:34:00 PM" = Format(#12:34#, "long time"))

            Console.Write("3) ")  
            PassFail("12:34:00 PM" = Format(#12:34#, "medium time"))

            Console.Write("4) ")  
            PassFail("12:34 PM" = Format(#12:34#, "short time"))

            Console.Write("5) ")  
            PassFail("4/24/2001 12:00:00 AM" = Format(#4/24/2001#, "general date"))
   
            Console.Write("6) ")  
            PassFail("Tuesday, April 24, 2001" = Format(#4/24/2001#, "long date"))
   
            Console.Write("7) ")  
            PassFail("Tuesday, April 24, 2001" = Format(#4/24/2001#, "medium date"))
   
            Console.Write("8) ")  
            PassFail("4/24/2001" = Format(#4/24/2001#, "short date"))

            Console.Write("9) ")  
            PassFail("123.46" = Format(123.456, "fixed"))

            Console.Write("10) ")  
            PassFail("123.46" = Format(123.456, "standard"))

            Console.Write("11) ")  
            PassFail("45.60%" = Format(0.456, "percent"))

            Console.Write("12) ")  
            PassFail("1.23E+02" = Format(123.456, "scientific"))

            Console.Write("13) ")  
            PassFail("$123.46" = Format(123.456, "currency"))

            Console.Write("14) ")  
            PassFail("True" = Format(1, "true/false"))

            Console.Write("15) ")  
            PassFail("False" = Format(0, "true/false"))

            Console.Write("16) ")  
            PassFail("Yes" = Format(1, "yes/no"))

            Console.Write("17) ")  
            PassFail("No" = Format(0, "yes/no"))

            Console.Write("18) ")  
            PassFail("On" = Format(1, "on/off"))

            Console.Write("19) ")  
            PassFail("Off" = Format(0, "on/off"))


            Console.WriteLine("Test with upper case")
            Console.Write("1a) ")  
            PassFail("4.6" = Format(1.2 + 3.4, "GENERAL NUMBER"))

            Console.Write("2a) ")  
            PassFail("12:34:00 PM" = Format(#12:34#, "LONG TIME"))

            Console.Write("3a) ")  
            PassFail("12:34:00 PM" = Format(#12:34#, "MEDIUM TIME"))

            Console.Write("4a) ")  
            PassFail("12:34 PM" = Format(#12:34#, "SHORT TIME"))

            Console.Write("5a) ")  
            PassFail("4/24/2001 12:00:00 AM" = Format(#4/24/2001#, "GENERAL DATE"))
   
            Console.Write("6a) ")  
            PassFail("Tuesday, April 24, 2001" = Format(#4/24/2001#, "LONG DATE"))
   
            Console.Write("7a) ")  
            PassFail("Tuesday, April 24, 2001" = Format(#4/24/2001#, "MEDIUM DATE"))
   
            Console.Write("8a) ")  
            PassFail("4/24/2001" = Format(#4/24/2001#, "SHORT DATE"))

            Console.Write("9a) ")  
            PassFail("123.46" = Format(123.456, "FIXED"))

            Console.Write("10a) ")  
            PassFail("123.46" = Format(123.456, "STANDARD"))

            Console.Write("11a) ")  
            PassFail("45.60%" = Format(0.456, "PERCENT"))

            Console.Write("12a) ")  
            PassFail("1.23E+02" = Format(123.456, "SCIENTIFIC"))

            Console.Write("13a) ")  
            PassFail("$123.46" = Format(123.456, "CURRENCY"))

            Console.Write("14a) ")  
            PassFail("True" = Format(1, "TRUE/FALSE"))

            Console.Write("15a) ")  
            PassFail("False" = Format(0, "TRUE/FALSE"))

            Console.Write("16a) ")  
            PassFail("Yes" = Format(1, "YES/NO"))

            Console.Write("17a) ")  
            PassFail("No" = Format(0, "YES/NO"))

            Console.Write("18a) ")  
            PassFail("On" = Format(1, "ON/OFF"))

            Console.Write("19a) ")  
            PassFail("Off" = Format(0, "ON/OFF"))


            Console.WriteLine("Test with mixed case")
            Console.Write("1b) ")  
            PassFail("4.6" = Format(1.2 + 3.4, "General numbeR"))

            Console.Write("2b) ")  
            PassFail("12:34:00 PM" = Format(#12:34#, "Long timE"))

            Console.Write("3b) ")  
            PassFail("12:34:00 PM" = Format(#12:34#, "Medium timE"))

            Console.Write("4b) ")  
            PassFail("12:34 PM" = Format(#12:34#, "Short timE"))

            Console.Write("5b) ")  
            PassFail("4/24/2001 12:00:00 AM" = Format(#4/24/2001#, "General datE"))
   
            Console.Write("6b) ")  
            PassFail("Tuesday, April 24, 2001" = Format(#4/24/2001#, "Long datE"))
   
            Console.Write("7b) ")  
            PassFail("Tuesday, April 24, 2001" = Format(#4/24/2001#, "Medium datE"))
   
            Console.Write("8b) ")  
            PassFail("4/24/2001" = Format(#4/24/2001#, "Short datE"))

            Console.Write("9b) ")  
            PassFail("123.46" = Format(123.456, "FixeD"))

            Console.Write("10b) ")  
            PassFail("123.46" = Format(123.456, "StandarD"))

            Console.Write("11b) ")  
            PassFail("45.60%" = Format(0.456, "PercenT"))

            Console.Write("12b) ")  
            PassFail("1.23E+02" = Format(123.456, "ScientifiC"))

            Console.Write("13b) ")  
            PassFail("$123.46" = Format(123.456, "CurrencY"))

            Console.Write("14b) ")  
            PassFail("True" = Format(1, "True/False"))

            Console.Write("15b) ")  
            PassFail("False" = Format(0, "True/False"))

            Console.Write("16b) ")  
            PassFail("Yes" = Format(1, "Yes/No"))

            Console.Write("17b) ")  
            PassFail("No" = Format(0, "Yes/No"))

            Console.Write("18b) ")  
            PassFail("On" = Format(1, "On/Off"))

            Console.Write("19b) ")  
            PassFail("Off" = Format(0, "On/Off"))

#if 0
            Console.Write("2) ")  
            PassFail("xx" = Format(CStr("xx")))

            Console.Write("3) ")  
            PassFail("xx" = Format(CStr("xx"), "@@"))

            Console.Write("4) ")  
            PassFail("5" = Format(CInt(5), "#"))

            Console.Write("5) ")  
            PassFail("123,456.789" = Format(CDbl(123456.789), "###,###.####"))

            Console.Write("6) ")  
            PassFail("96.00%" = Format(CDbl(0.96), "Percent"))

            Console.Write("7) ")  
            PassFail("XX" = Format(CStr("xx"), ">@@"))

            Console.Write("8) ")  
            PassFail("(5)" = Format(CInt(-5), "##&(##)"))

            Console.Write("9) ")  
            PassFail(Format(2147483647) = "2147483647")

            Console.Write("9) ")  
            PassFail(Format(1.234@, "000.000") = "001.234")
#end if
            'Bug295297
            Console.Write("   Bug295297: ")  
	        PassFail(Format(True, "The answer is {0}") = "The answer is True")

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub InstrBinaryTests()
        NewApiTest("InStr tests (Option Compare Binary)")

        Try
            Console.Write("1) ")
            PassFail(InStr(1, "test", "TEST", CompareMethod.Text) = 1)

            Console.Write("2) ") 
            PassFail(InStr(1, "test", "TEST", CompareMethod.Binary) = 0)

            Console.Write("3) ") 
            PassFail(InStr("test", "test") = 1)

            Console.Write("4) ")
            PassFail(InStr(1, "test", "test", CompareMethod.Text) = 1)

            Console.Write("5) ")
            PassFail(InStr(1, "test", "test", CompareMethod.Binary) = 1)

            Console.Write("6) ")
            PassFail(InStr(1, "test", "TEST", CompareMethod.Text) = 1)

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

            Console.Write("11) ")
            PassFail(InStr("ABCD", ChrW(0)) = 0)

        Catch ex As Exception
            Failed(ex)
        End Try

        InStrBug31477

#if 0
        Dim x7var(2) As Byte
        Dim x7var2(2) as Byte
        x7var(0) = 88
        x7var(1) = 0
        x7var2(0) = 120
        x7var2(1) = 0
        'Console.WriteLine("9) ")  PassFail(VBA.InStr(1, x7var, x7var2, CompareMethod.Text) = 1)
#end if
    End Sub



    Sub InStrRevTests()
        NewApiTest("InStrRev tests")
        Try

            Console.Write("1) ") 
            PassFail(InStrRev("test", "test") = 1)

            Console.Write("2) ") 
            PassFail(InStrRev("test", "test", 1) = 0)

            Console.Write("3) ") 
            PassFail(InStrRev("test", "TEST", 1) = 0)

            s = ChrW(&H9C)  'oe Ligature
            Console.Write("4) ") 
            PassFail(InStrRev("Rex " & s & "dipus", "oe", , CompareMethod.Binary) = 0)

'UNDONE: fix ligature handling
'            Console.Write("5) ") 
'            PassFail(InStrRev("Rex " & s & "dipus", "oe", , CompareMethod.Text) = 5)

            Console.Write("6) ") 
            PassFail(InStrRev("", "wow") = 0)

            Console.Write("7) ") 
            PassFail(InStrRev("wow", "") = 3)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub JoinTests()
        NewApiTest("Join tests")

        Try
            Dim Ary() as String = { "Fred", "Flintstone" }

            Console.Write("1) ")
            PassFail(Join(Ary, "%") = "Fred%Flintstone")

            Console.Write("2) ")
            PassFail(Join(Ary, "<>") = "Fred<>Flintstone")

            Console.Write("3) ")
            PassFail(Join(Ary, "") = "FredFlintstone")

            Console.Write("4) ")
            PassFail(Join(Ary) = "Fred Flintstone")

            Dim EmptyArray as String()
            Console.Write( "5) ")
            PassFail(Join(EmptyArray, "") = "") 'Bug 34985 + 68126

            Dim str1(1) As String
            str1(0) = "Hello"
            str1(1) = "World"
            Console.Write( "6) ")
            Console.WriteLine( Join(str1, "$$") )

            Dim ObAry() as Object = { "Fred", "Flintstone" }
            Console.Write("7) ")
            PassFail(Join(ObAry, "%") = "Fred%Flintstone")
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub SplitTests()
        NewApiTest("Split tests")

        Try
            Dim vntAry As String()

            vntAry = Split("aBcbBdb", "b", , CompareMethod.Text)

            Console.Write("1) ")
            PassFail(vntAry.GetUpperBound(0) = 4)

            Console.Write("2) ")
            PassFail(vntAry(0) = "a")

            Console.Write("3) ")
            PassFail(vntAry(1) = "c")

            Console.Write("4) ")
            PassFail(vntAry(2) = "")

            Console.Write("5) ")
            PassFail(vntAry(3) = "d")

            Console.Write("6) ")
            PassFail(vntAry(4) = "")

            Console.Write("7) ")
            vntAry = Split("", "")
#if 1
'Workaround for codegen assert
            Dim b As Boolean
            PassFail(vntAry(0) = "")
            PassFail(vntAry.GetUpperBound(0) = 0)
#else
            PassFail(vntAry(0) = "" And vntAry.GetUpperBound(0) = 0)
#end if
            Console.Write("9) ")
            vntAry = Split("Bug 34077 test")
#if 1
'Workaround for codegen assert
            PassFail(vntAry.GetUpperBound(0) = 2) 
            PassFail(vntAry(0) = "Bug")
            PassFail(vntAry(1) = "34077") 
            PassFail(vntAry(2) = "test")
#else
            PassFail((vntAry.GetUpperBound(0) = 2) And (vntAry(0) = "Bug") And (vntAry(1) = "34077") And (vntAry(2) = "test"))
#end if
#if 0
            Try
                Dim vEmpty As Object
                vntAry = Split(vEmpty, " &")
                Console.Write("10) "
                Console.WriteLine(vntAry(0))
                Console.Write(PassFail(Err.Number = 9))
            Catch ex As Exception
                Failed(ex)
            End Try
#end if

            vntAry = Split("aaa", "a")
            Console.Write("11) ")
            PassFail(vntAry.Length = 4 And vntAry(0) = "" And vntAry(1) = "" And vntAry(2) = "" And vntAry(3) = "")

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub ReplaceTests()
        NewApiTest("Replace tests")

        Try
            Console.Write("1) ") 
            PassFail(Replace("a.b.c.d", ".", "-", 1, , CompareMethod.Text) = "a-b-c-d")

            Console.Write("2) ") 
            PassFail(Replace("a.b.c.d", ".", "-", 1, , CompareMethod.Binary) = "a-b-c-d")

            Console.Write("3) ") 
            PassFail(Replace("a.b.c.d", ".", "-", 1, 2, CompareMethod.Text) = "a-b-c.d")

            Console.Write("4) ") 
            PassFail(Replace("a.b.c.d", ".", "-", 1, 2, CompareMethod.Binary) = "a-b-c.d")

            Console.Write("5) ") 
            PassFail(Replace("aaa", "a", "b", 3, , CompareMethod.Binary) = "b")

            Console.Write("6) ") 
            PassFail(Replace("aaa", "a", "b", 1, 2, CompareMethod.Binary) = "bba")

            Console.Write("7) ") 
            Dim Res() As String
            Res = Split("adbDc", "d", 2, CompareMethod.Text) 
#if 1
'Workaround for codegen assert
            PassFail(Res(0) = "a")
            PassFail(Res(1) = "bDc")
#else
            PassFail((Res(0) = "a") And (Res(1) = "bDc"))
#end if
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub SpaceTests()
        NewApiTest("Space tests")

        Try
            Console.Write("1) ")
            PassFail(Space(0) = "")

            Console.Write("2) ")
            PassFail(Space(10) = "          ")

        Catch ex As Exception
            Failed(ex)
        End Try

        'Error cases
        Try
            Console.Write("3) ")
            Console.WriteLine(Space(-1))
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub StrReverseTests()


        NewApiTest("StrReverse tests")
        Try
            Dim s, sReverse As String
            Dim StringWithSurrogateEmbedded As String = New String(New Char() {ChrW(65), ChrW(55299), ChrW(57088), ChrW(66)})   ' "A??B"
            Dim StringWithSurrogateAtEnd As String = New String(New Char() {ChrW(65), ChrW(66), ChrW(55299), ChrW(57088)})   ' "AB??"
            Dim StringWithSurrogateAtStart As String = New String(New Char() {ChrW(55299), ChrW(57088), ChrW(65), ChrW(66)})   ' "??AB"
            Dim bSuccess As Boolean

            s = "ABCDEFGHIJKLMNOP"
            Console.Write("1) ")
            PassFail(StrReverse(s) = "PONMLKJIHGFEDCBA")

            Console.Write("2) ")
            PassFail(StrReverse("") = "")

            s = Nothing
            Console.Write("3) ")
            PassFail(StrReverse(s) = "")

            Console.Write("4) ")
            sReverse = StrReverse(StringWithSurrogateEmbedded)
            bSuccess = (sReverse.Chars(1) = StringWithSurrogateEmbedded.Chars(1)) AndAlso _
                        (sReverse.Chars(2) = StringWithSurrogateEmbedded.Chars(2)) AndAlso _
                        (sReverse.Chars(0) = StringWithSurrogateEmbedded.Chars(3)) AndAlso _
                        (sReverse.Chars(3) = StringWithSurrogateEmbedded.Chars(0))
            PassFail(bSuccess)

            Console.Write("5) ")
            sReverse = StrReverse(StringWithSurrogateAtEnd)
            bSuccess = (sReverse.Chars(0) = StringWithSurrogateAtEnd.Chars(2)) AndAlso _
                        (sReverse.Chars(1) = StringWithSurrogateAtEnd.Chars(3)) AndAlso _
                        (sReverse.Chars(2) = StringWithSurrogateAtEnd.Chars(1)) AndAlso _
                        (sReverse.Chars(3) = StringWithSurrogateAtEnd.Chars(0))
            PassFail(bSuccess)

            Console.Write("6) ")
            sReverse = StrReverse(StringWithSurrogateAtStart)
            bSuccess = (sReverse.Chars(0) = StringWithSurrogateAtStart.Chars(3)) AndAlso _
                        (sReverse.Chars(1) = StringWithSurrogateAtStart.Chars(2)) AndAlso _
                        (sReverse.Chars(2) = StringWithSurrogateAtStart.Chars(0)) AndAlso _
                        (sReverse.Chars(3) = StringWithSurrogateAtStart.Chars(1))
            PassFail(bSuccess)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub LCaseTests()
        NewApiTest("LCase tests")

        Try
            s = "ABCDE"
            Console.Write("1) ")
            PassFail(LCase(s) = "abcde")

            Console.Write("2) ")
            PassFail(LCase("") = "")

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub UCaseTests()
        NewApiTest("UCase tests")

        Try

            Console.Write("1) ")
            PassFail(UCase("abcdefg") = "ABCDEFG")

            Console.Write("2) ")
            PassFail(UCase("") = "")

            'Console.Write("3) ")  PassFail(IsDBNull(UCase(DBNull))))

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub LeftTests()
        NewApiTest("Left tests")

        Try
            s = "ABCDE"

            Console.Write("1) ")
            PassFail(Left(s, 0) = "")

            Console.Write("2) ")
            PassFail(Left(s, 1) = "A")

            Console.Write("3) ")
            PassFail(Left(s, 3) = "ABC")

            Console.Write("4) ")
            PassFail(Left(s, 5) = "ABCDE")

            Console.Write("5) ")
            PassFail(Left(s, 6) = "ABCDE")

        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("6) ")
            Console.WriteLine(Left(s, -1))
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub RightTests()
        NewApiTest("Right tests")

        Try
            s = "ABCDE"

            Console.Write("1) ")
            PassFail(Right(s, 0) = "")

            Console.Write("2) ")
            PassFail(Right(s, 1) = "E")

            Console.Write("3) ")
            PassFail(Right(s, 3) = "CDE")

            Console.Write("4) ")
            PassFail(Right(s, 5) = "ABCDE")

            Console.Write("5) ")
            PassFail(Right(s, 6) = "ABCDE")

        Catch ex As Exception
            Failed(ex)
        End Try

        'Error cases
        Try

            Console.Write("6) ")
            Console.WriteLine(Right(s, -1))
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub MidTests()
        NewApiTest("Mid tests")

        Try
            s = "ABCDE"

            Console.Write("1) ")
            PassFail(Mid(s, 1) = "ABCDE")

            Console.Write("2) ")
            PassFail(Mid(s, 3) = "CDE")

            Console.Write("3) ")
            PassFail(Mid(s, 5) = "E")

            Console.Write("4) ")
            PassFail(Mid(s, 6) = "")

            Console.Write("5) ")
            PassFail(Mid(s, 1, 0) = "")

            Console.Write("6) ")
            PassFail(Mid(s, 1, 5) = "ABCDE")

            Console.Write("7) ")
            PassFail(Mid(s, 1, 6) = "ABCDE")

            Console.Write("8) ")
            PassFail(Mid(s, 5, 3) = "E")

            Console.Write("9) ")
            PassFail(Mid(s, 6, 1) = "")

        Catch ex As Exception
            Failed(ex)
        End Try

        'The following tests broken into separate subs to work around assert in metaemit
        MidSubTest1
        MidSubTest2
    End Sub



    Sub MidSubTest1()
        'Enable when metaemit assert fixed

        Try
            s = "ABC"
            Console.Write("10) ")
            Console.WriteLine(Mid(s, -1))
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub MidSubTest2()
        Try
            s = "ABC"
            Console.Write("11) ")
            Console.WriteLine(Mid(s, 0))
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub MidStmtTests()
        NewApiTest("Mid (String) tests")

        Try
            s = "ABCDEFG"
            Console.Write("1) " )
            Mid(s, 4, 1) = "-"
            PassFail(s = "ABC-EFG")

            s = "ABCDEFG"
            Console.Write("2) " )
            Mid(s, 4, 3) = "123456789"
            PassFail(s = "ABC123G")

            s = "ABCDEFG"
            Console.Write("3) ")
            Mid(s, 4, 4) = "123456789"
            PassFail(s = "ABC1234")

            s = "ABCDEFG"
            Console.Write("4) " )
            Mid(s, 4, 40) = "123456789"
            PassFail(s = "ABC1234")

            s = "ABCDEFG"
            Console.Write("5) " )
            Mid(s, 1, 40) = "123456789"
            PassFail(s = "1234567")

            s = "ABCDEFG"
            Console.Write("6) " )
            Mid(s, 7, 1) = "123456789"
            PassFail(s = "ABCDEF1")

        Catch ex As Exception
            Failed(ex)
        End Try

        MidStmtSubTest1
    End Sub



    Sub MidStmtSubTest1()
        'Broken out into separate sub because of metaemit Assert
        Try
            s = "ABCDEFG"
            Console.Write("7) ")
            Mid(s, 8, 1) = "123456789"
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub TrimTests()
        NewApiTest("LTrim/Trim/RTrim tests")

        Try
            s = "    ABCD EFG    "
            Console.Write("1) ")
            PassFail(Trim(s) = "ABCD EFG")

            Console.Write("2) ")
            PassFail(LTrim(s) = "ABCD EFG    ")

            Console.Write("3) ")
            PassFail(RTrim(s) = "    ABCD EFG")

            s = "    ABCD EFG"
            Console.Write("4) ")
            PassFail(Trim(s) = "ABCD EFG")

            Console.Write("5) ")
            PassFail(LTrim(s) = "ABCD EFG")

            Console.Write("6) ")
            PassFail(RTrim(s) = "    ABCD EFG")

            s = "ABCD EFG    "
            Console.Write("7) ")
            PassFail(Trim(s) = "ABCD EFG")

            Console.Write("8) ")
            PassFail(LTrim(s) = "ABCD EFG    ")

            Console.Write("9) ")
            PassFail(RTrim(s) = "ABCD EFG")

            s = "ABCD EFG"
            Console.Write("10) ")
            PassFail(Trim(s) = "ABCD EFG")

            Console.Write("11) ")
            PassFail(LTrim(s) = "ABCD EFG")

            Console.Write("12) ")
            PassFail(RTrim(s) = "ABCD EFG")

            s = ""
            Console.Write("13) ")
            PassFail(Trim(s) = "")

            Console.Write("14) ")
            PassFail(LTrim(s) = "")

            Console.Write("15) ")
            PassFail(RTrim(s) = "")

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub StrCompTests()
        Dim Result As Integer
        NewApiTest("StrComp tests")

        Try
            Console.Write("1) ")  
            PassFail(StrComp("abcd", "abcd", CompareMethod.Text) = 0)

            Console.Write("2) ")  
            PassFail(StrComp("abcd", "abcd", CompareMethod.Binary) = 0)

            Console.Write("3) ")  
            PassFail(StrComp("ABCD", "abcd", CompareMethod.Text) = 0)

            Console.Write("4) ")  
            PassFail(StrComp("ABCD", "abcd", CompareMethod.Binary) = -1)

            Console.Write("5) ")  
            PassFail(StrComp("BCDE", "ABCD", CompareMethod.Text) = 1)

            Console.Write("6) ")  
            PassFail(StrComp("BCDE", "ABCD", CompareMethod.Binary) = 1)

            Console.Write("7) ")  
            PassFail(StrComp("ABCD", "BCDE", CompareMethod.Text) = -1)

            Console.Write("8) ")  
            PassFail(StrComp("ABCD", "BCDE", CompareMethod.Binary) = -1)

        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("9) ")  
            Result = StrComp("ABCD", "BCDE", CType(2, CompareMethod)) 'Test Illegal CompareMethod RAID 128422
            Failed()
        Catch
            Passed()
        End Try
    End Sub



    Sub StrConvTests()
        NewApiTest("StrConv Tests")

        Try
            Console.Write("1) ")
            PassFail(StrConv("AbCd", VbStrConv.UpperCase) = "ABCD")

            Console.Write("2) ")
            PassFail(StrConv("AbCd", VbStrConv.LowerCase) = "abcd")

            Console.Write("3) ") 
            PassFail(StrConv("AbCd", VbStrConv.ProperCase) = "Abcd")

            'UNDONE: Enable when this is working
            Console.Write("4) ")
            PassFail(StrConv(ChrW(0), VbStrConv.ProperCase) = ChrW(0)) 'Bug 31582

            Console.Write("5) ")
            PassFail(StrConv("a simple" & ChrW(0) & "little test", VbStrConv.ProperCase) = "A Simple" & ChrW(0) & "Little Test") 'Bug 31582

            Console.Write("6) ")
            PassFail(StrConv("this is A sentence.  and this Is too.", VbStrConv.ProperCase) = "This Is A Sentence.  And This Is Too.")

            'Certain Japanese string conversions will INCREASE the length
            Console.Write("7) ")
#If 0 Then
'UNDONE: Need to make this work on Win98 or on machines that don't have Japan locale support
            Dim fJapanLocaleSupported As Boolean
            Try
                Dim loc As CultureInfo = New CultureInfo(1041)
                fJapanLocaleSupported = True
            Catch 
                'Eat the exception
            End Try

            If fJapanLocaleSupported Then
                Dim s1, s2 As String
                s1 = ChrW(&H30AC)
                PassFail(Len(s1) = 1 And Len(StrConv(s1, VbStrConv.Narrow, 1041)) = 2)
            Else
                Passed()
                Passed()
            End If
#Else
            Passed()
#End If

        Catch ex As Exception
            Failed(ex)
        End Try

        StrConvSubTest1
        StrConvSubTest2
        StrConvSubTest3
        StrConvSubTest4
    End Sub



    Sub StrConvSubTest1()
        Try
            Console.Write("8) ")
            b = (StrConv("abcd", VbStrConv.Wide) = "abcd")
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub StrConvSubTest2()
        Try
            Console.Write("9) ")
            b = (StrConv("abcd", VbStrConv.Narrow) = "abcd")
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub StrConvSubTest3()
        Try
            Console.Write("10) ")
            b = (StrConv("abcd", VbStrConv.Katakana) = "abcd")
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub StrConvSubTest4()
        Try
            Console.Write("11) ")
            b = (StrConv("abcd", VbStrConv.Hiragana) = "abcd")
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub ErrorCheck(ByVal iErrExpected As Long)
        If iErrExpected = Err.Number Then
            Console.WriteLine("passed")
        Else
            Console.WriteLine("FAILED !!! Error: " & CStr(Err.Number) & "  Expected " & CStr(iErrExpected))
        End If
    End Sub



#if 0
    Sub TestStringConcat()
	    Dim v1 As Object, v2 As Object, v3 As Object

	    v1 = DBNull
	    v2 = DBNull
	    v3 = DBNull

	    NewApiTest("StrCatVar tests")

	    Console.Write("1) ")  
        PassFail(VarType(v1 & v2) = vbNull)

	    v1 = "ABC"
	    Console.Write("2) ")  
        PassFail((v1 & v2) = "ABC")

	    v2 = "ABC"
	    Console.Write("3) ")  PassFail((v3 & v2) = "ABC")

	    Console.Write("4) ")  PassFail((v1 & v2) = "ABCABC")

    End Sub
#end if



    Sub TestLikeBinary()
        Dim i As Integer

	    NewApiTest("Like tests (Option Compare Binary)")
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



    'Same as Test, but used when non-printable chars exist in Source or Pattern, prints Title instead.
    Sub TestNonPrintable( ByVal Source As String, ByVal Pattern As String, ByVal Expected As Boolean, ByVal Title As String)
        Dim Actual As Boolean
        Actual = Source Like Pattern
        If  Actual = Expected Then
            Console.WriteLine( Title & " " & Actual & " passed" )
        Else
            Console.WriteLine( Title & " " & Actual & " FAILED" )
        End If
    End Sub



    Sub TestException( ByVal Source As String, ByVal Pattern As String)
        Dim Actual As Boolean
        Try
            Actual = Source Like Pattern
            Console.WriteLine("Source=" & Source & " Pattern=" & Pattern & " FAILED to throw expected exception" )
        Catch e As System.Exception
            Console.WriteLine("Source=" & Source & " Pattern=" & Pattern & " Threw expected exception: passed" )
        End Try
    End Sub

   

    Sub TestDates()
        NewApiTest( "Date tests" )

        Console.WriteLine( WeekDayName(1, , FirstDayOfWeek.System) ) 'Sunday
        Console.WriteLine( WeekDayName(1, , FirstDayOfWeek.Sunday) ) 'Sunday
        Console.WriteLine( WeekDayName(2, True, FirstDayOfWeek.Sunday) ) 'Mon
        Console.WriteLine( WeekDayName(2, False, FirstDayOfWeek.Monday) ) 'Tuesday
        Console.WriteLine( WeekDayName(7, True, FirstDayOfWeek.Monday) ) 'Sun
        Console.WriteLine( WeekDayName(7, False, FirstDayOfWeek.Saturday) ) 'Friday

        Try
            Console.WriteLine( WeekDayName(0, , FirstDayOfWeek.Sunday) )   
        Catch
            Console.WriteLine( "WeekDayName(0,,1) caught expected exception" )
        End Try  
        
        Try
            Console.WriteLine( WeekDayName(1, , CType(-1, FirstDayOfWeek)) )   
        Catch
            Console.WriteLine( "WeekDayName(1,,-1) caught expected exception" )
        End Try  
        
        Try
            Console.WriteLine( WeekDayName(8,,FirstDayOfWeek.Saturday) )   
        Catch
            Console.WriteLine( "WeekDayName(8,,7) caught expected exception" )
        End Try      
           
        Try
            Console.WriteLine( WeekDayName(7,,CType(8, FirstDayOfWeek)) )   
        Catch
            Console.WriteLine( "WeekDayName(7,,8) caught expected exception" )
        End Try 
        
        Console.WriteLine( "WeekOfYear for 1/1/2000: " & DatePart(DateInterval.WeekOfYear, #1/1/2000# ) ) '1

        'Waiting for bugfix 58803
        'Console.WriteLine( "WeekOfYear for 1/2/2000: " & DatePart(DateInterval.WeekOfYear, #1/2/2000# ) ) '2
        'Console.WriteLine( "WeekOfYear for 1/9/2000: " & DatePart(DateInterval.WeekOfYear, #1/9/2000# ) ) '3  

        Dim i As integer
        Dim fr As New System.Globalization.CultureInfo("fr-fr") 'French
        System.Threading.Thread.CurrentThread.CurrentCulture = fr
        If System.Threading.Thread.CurrentThread.CurrentCulture Is fr Then
            Console.WriteLine( "Current culture set to french" )
        Else
            Console.WriteLine( "Error setting current culture to french" )
        End If
        For i = 1 to 7
            Console.WriteLine( WeekDayName(i) )
        Next i

        Dim de As New System.Globalization.CultureInfo("de-de") 'German
        System.Threading.Thread.CurrentThread.CurrentCulture = de
        If System.Threading.Thread.CurrentThread.CurrentCulture Is de Then
            Console.WriteLine( "Current culture set to german" )
        Else
            Console.WriteLine( "Error setting current culture to german" )
        End If
        For i = 1 to 7
            Console.WriteLine( WeekDayName(i) )
        Next i
    End Sub


    
    Sub TestErrors
        NewApiTest("TestErrors" )

        dim i as short = 32767
        dim f As Single = 0.0
        Dim a(9) As Integer
        on error resume next

        a(20) = 0
        Console.writeline("Index Out of Range: " & err.number & " " & err.description)

        i = CShort(i + 1)
        Console.writeline("Overflow " & err.number & " " & err.description)
    End Sub



    Sub FilterTest()
	    NewApiTest("Filter tests")

        Dim Source(4) As String
        Dim Match As String
        Dim MultiDim(1,1) As String
        Dim NonArray As Object
        Dim NonString(1) As Integer
        Dim IgnoreCase(5) As String
    
        Source(0) = "abcd"
        Source(1) = "bcde"
        Source(2) = "cdef"
        Source(3) = "defg"
        Source(4) = "efgh"

        MultiDim(0,0) = "abcd"
        MultiDim(0,1) = "bcde"
        MultiDim(1,0) = "cdef"
        MultiDim(1,1) = "defg"

        NonArray = "defg"

        NonString(0) = 0
        NonString(1) = 1

        IgnoreCase(0) = "abcd"
        IgnoreCase(1) = "BCDE"
        IgnoreCase(2) = "cdef"
        IgnoreCase(3) = "DEFG"
        IgnoreCase(4) = "efgh"
        IgnoreCase(5) = "FGHI"

        Match = "a"
        TestF( "Filter1", Source, Match )

        Match = "b"
        TestF( "Filter2", Source, Match )

        Match = "c"
        TestF( "Filter3", Source, Match )
    
        Match = "d"
        TestF( "Filter4", Source, Match )

        Match = "e"
        TestF( "Filter5", Source, Match )

        Match = "f"
        TestF( "Filter6", Source, Match )

        Match = "g"
        TestF( "Filter7", Source, Match )

        Match = "h"
        TestF( "Filter8", Source, Match )

        Match = "i"
        TestF( "Filter9", Source, Match )

        Match = "a"
        TestF( "Filter10", Source, Match, False )

        Match = "b"
        TestF( "Filter11", Source, Match, False )

        Match = "c"
        TestF( "Filter12", Source, Match, False )

        Match = "d"
        TestF( "Filter13", Source, Match, False )

        Match = "e"
        TestF( "Filter14", Source, Match, False )

        Match = "f"
        TestF( "Filter15", Source, Match, False )

        Match = "g"
        TestF( "Filter16", Source, Match, False )

        Match = "h"
        TestF( "Filter17", Source, Match, False )

        Match = "i"
        TestF( "Filter18", Source, Match, False )

        Match = "b"
        TestF( "Filter50", IgnoreCase, Match, True, CompareMethod.Text )

        Match = "B"
        TestF( "Filter51", IgnoreCase, Match, True, CompareMethod.Text )

        Match = "c"
        TestF( "Filter52", IgnoreCase, Match, False, CompareMethod.Text )

        Match = "C"
        TestF( "Filter53", IgnoreCase, Match, False, CompareMethod.Text )

        FilterBug41895
        FilterBug41897
        FilterBug41900
    End Sub



    Sub TestF( ByVal Test As String, ByVal Source() As String, ByVal Match As String, Optional ByVal Include As Boolean = True, <Microsoft.VisualBasic.CompilerServices.OptionCompareAttribute> Optional ByVal [Compare] As CompareMethod = CompareMethod.Binary)
        Dim Result As String()
        Dim s As String

        Console.Write(Test & "  ")
        Result = Filter(Source, Match, Include, [Compare])

        If Not Result Is Nothing
            For Each s In Result
                Console.Write(s & " " )
            Next s
        End If

        Console.WriteLine()
    End Sub



#if 0
    Sub Bug9747()
        SetBugID( 9747 )

        Dim b As Boolean
        Dim v As Object

        b =  ("" <> String$(257, ChrW(0)))
        v = ("" <> String$(257, ChrW(0)))
        b = b And CBool(v)

        PrintResult(b)
    End Sub



    Sub Bug31584()
        SetBugID( 31584 )
        PassFail(ChrW(153) = String$(1, 153))
    End Sub
#end if



    Sub InStrBug31477()
        Console.Write( "Bug31477: " )

        Dim Temp As Integer
        Dim NullString As String = ""

        NullString = ""

        Temp = InStr(1, "ABCDEFG-ABCDEFG", NullString)
        PassFail(Temp = 1)
    End Sub



    Sub FilterBug41895()
        'Test empty Source Array   
        Console.Write( "FilterBug41895: " )
        Dim a(-1) As String
        Dim x As Object
        x = Filter(a, "bob", True, CompareMethod.Binary)
        Console.WriteLine(TypeName(x))
    End Sub


      
    Sub FilterBug41897()
        Console.Write( "FilterBug41897: " )
        'Test Object Source Array   
        Dim x As String()
        Dim b(0) As Object

        b(0) = "bob"
        x = Filter(b, "bob", True, CompareMethod.Binary)
        Console.WriteLine(x(0))
    End Sub



    Sub FilterBug41900()
        Console.Write( "FilterBug41900: " )
        Dim c(2) As String
        Dim x As String()

        c(0) = "aaaaa1"
        c(1) = "bbbbb1aa"
        c(2) = "aaabbb1"

        x = Filter(c, "aaa", True, CompareMethod.Binary)
        Console.WriteLine(x(0) & " " & x(1) & " " & UBound(x))
        x = Filter(c, "aaa", False, CompareMethod.Binary)
        Console.WriteLine(x(0) & " " & UBound(x))
    End Sub


End Module



Class MyCulture
    Inherits CultureInfo

    Private m_nfi As NumberFormatInfo
    Private m_dtfi As DateTimeFormatInfo
    Shared m_culture As MyCulture



    Shared Sub Init(ByVal lcid As Integer)
        'Set culture to Enlish to avoid locale diffs
        Dim ci As CultureInfo = New CultureInfo(&H409)
        Dim nfi As NumberFormatInfo = CType(ci.GetFormat(GetType(NumberFormatInfo)), NumberFormatInfo)
        Dim dtfi As DateTimeFormatInfo = CType(ci.GetFormat(GetType(DateTimeFormatInfo)), DateTimeFormatInfo)

        System.Threading.Thread.CurrentThread.CurrentCulture = New MyCulture(&H409, dtfi, nfi)
    End Sub



    Private Sub New(lcid As Integer, dtfi As DateTimeFormatInfo, nfi As NumberFormatInfo)
        MyBase.New(lcid)

        nfi = CType(nfi.Clone(), NumberFormatInfo)
        dtfi = CType(dtfi.Clone(), DateTimeFormatInfo)

        dtfi.DateSeparator = "/"
        dtfi.ShortDatePattern = "M/d/yyyy"

        m_nfi = NumberFormatInfo.ReadOnly(nfi)
        m_dtfi = DateTimeFormatInfo.ReadOnly(dtfi)
    End Sub



    Public Overrides Function GetFormat(ByVal service As Type) As Object
        If service Is GetType(NumberFormatInfo) Then
            Return m_nfi
        ElseIf service Is GetType(DateTimeFormatInfo) Then
            Return m_dtfi
        End If

        Throw New ArgumentException() 'UNDONE
    End Function



End Class



'CONSIDER: adding this to a Japanese only checkin suite
Module Bug239321
    'NOTE: REQUIRES SETTING JAPANESE LOCALE
    Sub Test()
        Console.Write("*** Bug 239321 : ")
        Dim  tst As Integer
        Dim s As String

        tst = &H3042 '  UNI Hex value of  First  Japanese Hiragana character. ANSI Hex Value : 82AA 
        'This tests an invalid ANSI character getting passed
        'when a unicode or other value is passed, only the hibyte is used
        PassFail(AscW(Chr(tst)) = &H30)
    End Sub
End Module


Module Bug304082

    Structure s1
        Dim i As Integer
    End Structure

    Sub Test()
        Console.Write("*** Bug 304082: ")
        Try
            Dim i As Integer
            Dim j As IntPtr     'same exception for uintptr
            i = Len(j)
            Console.WriteLine(i)
        Catch ex As ArgumentException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug316642

    Structure MyStruct
        Public Shared x As MyStruct
        Public y As Integer
    End Structure

    Sub Test()
        Console.Write("*** Bug 316642: ")
        Try
            Dim i As Integer
            Dim g As MyStruct

            i = Len(g)
            PassFail(i = 4)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module
