Imports System
Imports Microsoft.VisualBasic
Imports System.Globalization

Module Test

    Dim m_lBugID As Long

    Sub PrintResult(ByVal bPass As Boolean)
        If bPass Then
            Console.WriteLine( "passed" )
        Else
            Console.WriteLine( "FAILED !!!" )
        End If
        m_lBugID = 0
    End Sub

    Sub SetBugID(ByVal lID As Long, Optional ByVal lExceptedErr As Long = 0)
        If m_lBugID <> 0 Then
            Console.WriteLine( "FAILED !!! (unexpected exception)" )
        End If
        m_lBugID = lID
        Console.Write( "Bug" & CStr(m_lBugID) & ": ")
    End Sub

    Sub PassFail(ByVal bPassed As Boolean)
        If bPassed Then
            Console.WriteLine("passed")
        Else
            Console.WriteLine("FAILED !!!")
        End If
    End Sub

    Sub PassFailErrNumber(ByVal iErr As Integer)
        If Err.Number <> iErr Then
            Console.WriteLine("FAILED: incorrect error " & Err.Number)
        Else
            Console.WriteLine("passed: correct error thrown")
        End If

    End Sub

    Sub Main()
        MyCulture.Init(&H409)

        'Don't run NamedParamTest, just ensure compile
        'NamedParamTest
        TestDateAdd
        TestDateDiff
        TestIntrinsicFunctions
        TestDatePart
        TestPartFunctions
	    TestDateValue
	    TestTimeValue
        TestTimeSerial
        TestDateSerial
        TestWeekday

        RegressionTests
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

        v = VBA.DateAdd(Interval := "m", Number := 234, Date := v)
        v = VBA.DateDiff(Interval := "m", Date1 := v, Date2 := v, FirstDayOfWeek := vbSunday, FirstWeekOfYear := vbFirstJan1)
        v = VBA.DatePart(Interval := "m", Date := v, FirstDayOfWeek := vbSunday, FirstWeekOfYear := vbFirstJan1)
        v = VBA.DateSerial(Year := 1999, Month := 1, Day := 1)
        v = VBA.DateValue(Date := "1/1/99")
        v = VBA.Day(Date := v)
        v = VBA.Month(Date := v)
        v = VBA.Year(Date := v)
        v = VBA.Hour(Time := v)
        v = VBA.Minute(Time := v)
        v = VBA.Second(Time := v)
        v = VBA.TimeSerial(Hour := 1, Minute := 1, Second := 1)
        v = VBA.TimeValue(Time := "1:01:01 AM")
        v = VBA.Weekday(Date := v, FirstDayOfWeek := vbSunday)
        v = VBA.Year(Date := v)
#end if
    End Sub


    Public Sub TestDateAdd()
        Console.WriteLine( "Begin Test DateAdd" )
        'Test each of the interval types
        '================================
        Console.Write( "1) ") 
        PassFail(DateAdd("m", 1, #12/31/1999#) = #1/31/2000#)

        Console.Write( "2) ")
        PassFail(DateAdd("d", 1, #12/31/1999#) = #1/1/2000#)

        Console.Write( "3) ")
        PassFail(DateAdd("y", 1, #12/31/1999#) = #1/1/2000#)

        Console.Write( "4) " )
        PassFail(DateAdd("w", 1, #12/31/1999#) = #1/1/2000#)

        Console.Write( "5) " )
        PassFail(DateAdd("ww", 1, #12/31/1999#) = #1/7/2000#)

        Console.Write( "6) " )
        PassFail(DateAdd("h", 1, #12/31/1999#) = #12/31/1999 1:00:00 AM#)

        Console.Write( "7) " )
        PassFail(DateAdd("n", 1, #12/31/1999#) = #12/31/1999 12:01:00 AM#)

        Console.Write( "8) " )
        PassFail(DateAdd("s", 1, #12/31/1999#) =  #12/31/1999 12:00:01 AM#)

        Console.Write( "9) " )
        PassFail(DateAdd("yyyy", 1, #12/31/1999#) = #12/31/2000#)

        Console.Write( "10) " )
        PassFail(DateAdd("q", 1, #12/31/1999#) = #3/31/2000#)

        Try
            Console.Write( "11) " )
            DateAdd("x", 1, #12/31/1999#)
            Console.WriteLine( "FAILED! error not thrown" )
        Catch
            PassFailErrNumber(5)
        End Try

        'Test several numbers
        '=====================
        Console.Write( "12) " )
        PassFail(DateAdd("m", 0, #12/31/1999#) = #12/31/1999#)

        Console.Write( "13) " )
        PassFail(DateAdd("m", -1, #12/31/1999#) = #11/30/1999#)

        Console.Write( "14) " )
        PassFail(DateAdd("m", 1000, #12/31/1999#) = #4/30/2083#)

        Console.Write( "15) " )
        PassFail(DateAdd("m", -1000, #12/31/1999#) = #8/31/1916#)

        Console.Write( "16) " )
        PassFail(DateAdd("m", 1.34, #12/31/1999#) = #1/31/2000#)

        Console.Write( "17) " )
        PassFail(DateAdd("m", -0, #12/31/1999#) = #12/31/1999#)

        Console.Write( "18) " )
        PassFail(DateAdd("m", 0.345, #12/31/1999#) = #12/31/1999#)

        Try
            Console.Write( "19) " )
            DateAdd("m", 1.23444321444141E+15, #12/31/1999#)
            Console.WriteLine( "FAILED! error not thrown" )
        Catch
            PassFailErrNumber(6)
        End Try

        'Test various types of Object
        '==============================
        Dim var As Object
        var = #12/31/1999#
        Console.Write( "20) " )
        PassFail(DateAdd("m", 0, var) = #12/31/1999#)

        var = "12/31/99"
        Console.Write( "21) " )
        PassFail(DateAdd("m", 0, var) = #12/31/1999#)

        var = "1 Jan, 1999"
        Console.Write( "22) " )
        PassFail(DateAdd("m", 0, var) = #1/1/1999#)

        Console.WriteLine( "End Test DateAdd" )
        Console.WriteLine()
        Exit Sub
    
    End Sub

    Public Sub TestDateDiff()
        Console.WriteLine( "Begin Test DateDiff")
        'Test each of the interval types
        '================================
        Console.WriteLine("DateDiff Tests")

        Console.Write( "1) " )
        PassFail(DateDiff("m", #12/31/1999#, #12/31/2000 1:30:00 PM#) = 12)

        Console.Write( "2) " )
        PassFail(DateDiff("d", #12/31/1999#, #12/31/2000 1:30:00 PM#) = 366)

        Console.Write( "3) " )
        PassFail(DateDiff("y", #12/31/1999#, #12/31/2000 1:30:00 PM#) = 366)

        Console.Write( "4) " )
        PassFail(DateDiff("w", #12/31/1999#, #12/31/2000 1:30:00 PM#) = 52)

        Console.Write( "5) " )
        PassFail(DateDiff("ww", #12/31/1999#, #12/31/2000 1:30:00 PM#) = 53)

        Console.Write( "6) " )
        PassFail(DateDiff("h", #12/31/1999#, #12/31/2000 1:30:00 PM#) = 8797)

        Console.Write( "7) " )
        PassFail(DateDiff("n", #12/31/1999#, #12/31/2000 1:30:00 PM#) = 527850)

        Console.Write( "8) " )
        PassFail(DateDiff("yyyy", #12/31/1999#, #12/31/2000 1:30:00 PM#) = 1)

        Console.Write( "9) " )
        PassFail(DateDiff("q", #12/31/1999#, #12/31/2000 1:30:00 PM#) = 4)

        Try
            Console.Write( "10) " )
            DateDiff("x", #12/31/1999#, #12/31/2000 1:30:00 PM#)
        Catch
            PassFailErrNumber(5)
        End Try

        'Test First day of week, first day of year
        '==========================================
        Console.Write( "11) " )
        PassFail(DateDiff("ww", #12/31/1999#, #12/31/2000 1:30:00 PM#, vbSunday) = 53)

        Console.Write( "12) " )
        PassFail(DateDiff("ww", #12/31/1999#, #12/31/2000 1:30:00 PM#, vbThursday, vbFirstFourDays) = 52)

        Console.Write( "13) " )
        PassFail(DateDiff("ww", #12/31/1999#, #12/31/2000 1:30:00 PM#, , vbFirstFourDays) = 53)

        Try
            Console.Write( "14) " )
            DateDiff("ww", "12/32/1999", #12/31/2000 1:30:00 PM#, , vbFirstFourDays)
            Console.WriteLine( "FAILED! error not thrown" )
        Catch
            PassFailErrNumber(13)
        End Try


        Try
            Console.Write( "15) " )
            DateDiff("ww", #12/31/1999#, "13/31/2000 1:30:00 PM", , vbFirstFourDays)
        Catch
            PassFailErrNumber(13)
        End Try

        Dim dt1, dt2 As Date
        Dim v As Long
    
        dt1 = #1/1/2000#
        dt2 = #1/1/2001#

        Console.WriteLine("Test using WeekDay 'w' ")

        Console.Write("1) ")
        PassFail(DateDiff("w", dt1, dt2, vbUseSystemDayOfWeek) = 52)

        Console.Write("1) ")
        PassFail(DateDiff("w", dt1, dt2, vbMonday) = 52)

        Console.Write("1) ")
        PassFail(DateDiff("w", dt1, dt2, vbTuesday) = 52)

        Console.Write("1) ")
        PassFail(DateDiff("w", dt1, dt2, vbWednesday) = 52)

        Console.Write("1) ")
        PassFail(DateDiff("w", dt1, dt2, vbThursday) = 52)

        Console.Write("1) ")
        PassFail(DateDiff("w", dt1, dt2, vbFriday) = 52)

        Console.Write("1) ")
        PassFail(DateDiff("w", dt1, dt2, vbSaturday) = 52)

        Console.Write("1) ")
        PassFail(DateDiff("w", dt1, dt2, vbSunday) = 52)

        Console.WriteLine("Test using WeekOfYear 'ww'")

        Console.Write("1) ")
        PassFail(DateDiff("ww", dt1, dt2, vbUseSystemDayOfWeek) = 53)

        Console.Write("1) ")
        PassFail(DateDiff("ww", dt1, dt2, vbMonday) = 53)

        Console.Write("1) ")
        PassFail(DateDiff("ww", dt1, dt2, vbTuesday) = 52)

        Console.Write("1) ")
        PassFail(DateDiff("ww", dt1, dt2, vbWednesday) = 52)

        Console.Write("1) ")
        PassFail(DateDiff("ww", dt1, dt2, vbThursday) = 52)

        Console.Write("1) ")
        PassFail(DateDiff("ww", dt1, dt2, vbFriday) = 52)

        Console.Write("1) ")
        PassFail(DateDiff("ww", dt1, dt2, vbSaturday) = 52)

        Console.Write("1) ")
        PassFail(DateDiff("ww", dt1, dt2, vbSunday) = 53)

        Console.WriteLine( "End Test DateDiff")
        Console.WriteLine()
        Exit Sub
    
    End Sub

    Public Sub TestDatePart()
        Console.WriteLine( "Begin Test DatePart")
        'Test each of the interval types
        '================================
        Console.Write( "1) ") : PassFail(DatePart("yyyy", #12/31/1999#) = 1999)
        Console.Write( "2) ") : PassFail(DatePart("m", #12/31/1999#) = 12)
        Console.Write( "3) ") : PassFail(DatePart("d", #12/31/1999#) = 31)
        Console.Write( "4) ") : PassFail(DatePart("h", #12/31/1999#) = 0)
        Console.Write( "5) ") : PassFail(DatePart("n", #12/31/1999#) = 0)
        Console.Write( "6) ") : PassFail(DatePart("s", #12/31/1999#) = 0)
        'UNDONE: BUGBUG Console.Write( "7) ") : PassFail(DatePart("ww", #4/4/1999#) = 15)
        Console.Write( "8) ") : PassFail(DatePart("w", #4/4/1999#) = 1)
        Console.Write( "9) ") : PassFail(DatePart("q", #12/31/1999#) = 4)
        Console.Write( "10) ") : PassFail(DatePart("y", #12/31/1999#) = 365)

    
        'Test for error conditions
        '==========================
        Dim var As Object

        Try
            Console.Write( "11) " )
            var = DatePart("x", #12/31/1999#)
            Console.WriteLine( "FAILED! error not thrown" )
        Catch
            PassFailErrNumber(5)
        End Try

        Try
            Console.Write( "12) " )
            var = DatePart("m", "13/32/1999")
        Catch
            PassFailErrNumber(13)
        End Try

        'DatePart testing firstdayofweek
        Dim dt1 As Date
        Dim v As Long
    
        dt1 = #1/1/2000#

        Console.WriteLine("DatePart using 'w' (FirstDayOfWeek)")

        Console.Write("1) " )
        PassFail(DatePart("w", dt1, vbUseSystemDayOfWeek) = 7)

        Console.Write("2) " )
        PassFail(DatePart("w", dt1, vbMonday) = 6)

        Console.Write("3) " )
        PassFail(DatePart("w", dt1, vbTuesday) = 5)

        Console.Write("4) " )
        PassFail(DatePart("w", dt1, vbWednesday) = 4)

        Console.Write("5) " )
        PassFail(DatePart("w", dt1, vbThursday) = 3)

        Console.Write("6) " )
        PassFail(DatePart("w", dt1, vbFriday) = 2)

        Console.Write("7) " )
        PassFail(DatePart("w", dt1, vbSaturday) = 1)

        Console.Write("8) " )
        PassFail(DatePart("w", dt1, vbSunday) = 7)

        Console.WriteLine( "End Test DatePart")
        Console.WriteLine()
        Exit Sub
    
    End Sub

    Public Sub TestIntrinsicFunctions()
        Console.WriteLine( "Start Test Intrinsic Functions")

        Const FiveMinuteTicks As Int64 = Timespan.TicksPerMinute * 5
        Dim ID As String
        Dim var As Object
        Dim s As String
        Dim dt As Date

        'These tests are fairly limited since they would
        ' otherwise have to change the system date/time
        Try
            Console.Write("Today Get : ")
            dt = Today
            If dt.TimeOfDay.Ticks <> 0 Then
                Failed()
            Else
                Passed()

                Console.Write("Today Set : ")

                'If we are less than 5 minutes to midnight, then skip skip this test
                If (TimeSpan.TicksPerDay - System.DateTime.Now.TimeOfDay.Ticks) < FiveMinuteTicks Then
                    'Skip this test to avoid possible messing up system date
                    Passed()

                Else
                    'Only test set if Get passed
                    Try
                        Today = dt
                        If Today = dt Then
                            Passed()
                        Else
                            Failed()
                        End If
                    Catch Ex as Security.SecurityException
                        'Allow this to pass if insufficient security
                        Passed()
                    Catch ex As Exception
                        Failed(ex)
                    End Try

                End If

            End If

            Try
                Console.Write("*** Bug 267119: ")
                Today = #01/01/0101#
                Failed()
            Catch ex As ArgumentException
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("DateString Get : ")
            s = DateString
            'DateString should always return yyyy-MM-dd or mm-dd-yyyy
            If (Not ((s.Chars(4) = "-"c OrElse s.Chars(7) = "-"c) OrElse _
                     (s.Chars(2) = "-"c OrElse s.Chars(5) = "-"c)) OrElse s.Length <> 10) Then
Console.WriteLine(s)
                Failed()
            Else
                Passed()

                Console.Write("DateString Set : ")

                'If we are less than 5 minutes to midnight, then skip skip this test
                If (TimeSpan.TicksPerDay - System.DateTime.Now.TimeOfDay.Ticks) < FiveMinuteTicks Then
                    'Skip this test to avoid possible messing up system date
                    Passed()

                Else
                    'Only test valid set if Get passed
                    Try
                        DateString = s
                        If Today = dt Then
                            Passed()
                        Else
                            Failed()
                        End If
                    Catch Ex as Security.SecurityException
                        'Allow this to pass if insufficient security
                        Passed()
                    Catch ex As Exception
                        Failed(ex)
                    End Try

                End If

            End If

            Console.Write("DateString Set (invalid arg): ")
            Try
                DateString = "XYZ"
                Failed()
            Catch
                ErrorCheck(5)
            End Try

            Console.Write("Now : ")
            dt = Now
            If dt.Year < 2000 OrElse dt.TimeOfDay.Ticks = 0 Then
                'This could fail if executed exactly at midnight, 
                ' but that's not likely
                Failed()
            Else
                Passed()
            End If

            Console.Write("TimeOfDay : ")
            dt = TimeOfDay()
            'UNDONE: any checks we can really do here?
            If dt.Year <> 1 OrElse _ 
                dt.Month <> 1 OrElse _
                dt.Day <> 1 OrElse _
                dt.Ticks = 0 Then
                Failed()
            Else
                Passed()

                'Only test set if Get passed
                Console.Write("TimeOfDay Set : ")
                If (TimeSpan.TicksPerDay - System.DateTime.Now.TimeOfDay.Ticks) < FiveMinuteTicks Then
                    'Skip this test to avoid possible messing up system date
                    Passed()

                Else
                    Try
                        TimeOfDay = dt
                        Passed()
                    Catch Ex as Security.SecurityException
                        'Allow this to pass if insufficient security
                        Passed()
                    Catch ex As Exception
                        Failed(ex)
                    End Try

                End If

            End If

            Console.Write("TimeString Get : ")
            s = TimeString
            'TimeString should always return HH:mm
            If s.Chars(2) <> ":"c OrElse _
               s.Chars(5) <> ":"c OrElse _
                s.Length <> 8 Then
                Failed()
            Else
                Passed()

                'Only test valid set if Get passed
                '  just trying to avoid someone's machine
                '  getting the system time wacked
                Console.Write("TimeOfDay Set : ")

                If (TimeSpan.TicksPerDay - System.DateTime.Now.TimeOfDay.Ticks) < FiveMinuteTicks Then
                    'Skip this test to avoid possible messing up system date
                    Passed()

                Else
                    Try
                        TimeString = s
                        Passed()
                    Catch Ex as Security.SecurityException
                        'Allow this to pass if insufficient security
                        Passed()
                    Catch ex As Exception
                        Failed(ex)
                    End Try

                End If

            End If

            Console.Write("TimeString Set (invalid arg): ")
            Try
                TimeString = "XYZ"
                Failed()
            Catch
                ErrorCheck(5)
            End Try

        Catch ex As Exception
            Failed(ex)
        End Try


        Console.WriteLine( "End Test Intrinsic Functions")
        Console.WriteLine()

    End Sub

    Public Sub TestPartFunctions()
        Console.WriteLine( "Start Test Part Functions")
        'Test for each of the parts
        '===========================
        Console.Write( "1) ") : PassFail(Day(#1/1/1900#) = 1)
        Console.Write( "2) ") : PassFail(Hour(#1/1/2000 1:20:00 PM#) = 13)
        Console.Write( "3) ") : PassFail(Minute(#1/1/1999 1:20:00 PM#) = 20)
        Console.Write( "4) ") : PassFail(Second(#1:20:34 PM#) = 34)
        Console.Write( "5) ") : PassFail(Month(#1/1/2000#) = 1)
        Console.Write( "6) ") : PassFail(Weekday(#1/1/2000#) = 7)
        Console.Write( "7) ") : PassFail(Year(#1/1/2000#) = 2000)

        Console.WriteLine( "End Test Part Functions")
        Console.WriteLine()
    End Sub


    Sub TestDateValue()
	    Console.WriteLine( "*** DateValue tests")

        On Error Resume Next
	    Console.Write( "1) ")
        PassFail(DateValue("01/01/1999") = #01/01/1999#)

	    Console.Write( "2) ")
        PassFail(DateValue("12/31/1999") = #12/31/1999#)

	    Console.Write( "3) ")
        PassFail(DateValue("01/01/2000") = #01/01/2000#)

	    Console.Write( "4) ")
        PassFail(DateValue("01/01/1999 12:00:00 AM") = #01/01/1999#)

	    Console.Write( "5) ")
        'UNDONE: FIX BUG PassFail(DateValue("12/31/1999 23:59:59 PM") = #12/31/1999#)
        Console.WriteLine()

	    Console.Write( "6) ")
        PassFail(DateValue("01/01/2000 06:00:00 AM") = #01/01/2000#)

	    Console.Write( "7) ")
        Err.Number = 0 : DateValue("1997") : ErrorCheck( 13 )
    End Sub

    Sub TestTimeValue()

        On Error Resume Next

	    Console.WriteLine( "*** TimeValue tests" )
	    Console.Write( "1) ")
        PassFail(TimeValue("1:15:30") = New DateTime(1,1,1,1,15,30,0))

	    Console.Write( "2) ")
        PassFail(TimeValue("1:00:59") = New DateTime(1,1,1,1,0,59,0))

	    Console.Write( "3) ")
        PassFail(TimeValue("12:59:59") = New DateTime(1,1,1,12,59,59,0))

	    Console.Write( "4) ")
        PassFail(TimeValue("12:00:00 AM") = New DateTime(1,1,1,0,0,0,0))

	    Console.Write( "5) ")
        'UNDONE: BUGBUG PassFail(TimeValue("12:00:00 PM") = New DateTime(1,1,1,12,0,0,0))
        Console.WriteLine()

	    Console.Write( "6) ")
        'UNDONE: BUGBUG PassFail(TimeValue("23:59:59 PM") = New DateTime(1,1,1,23,59,59,0))
        Console.WriteLine()

	    Console.Write( "7) ")
        PassFail(TimeValue("06:00:00 AM") = New DateTime(1,1,1,6,0,0,0))

	    Console.Write( "8) ")
        Err.Number = 0 : Call TimeValue("1997") : ErrorCheck(13)

    End Sub

    Sub ErrorCheck(ByVal iErrExpected As Long)
        If iErrExpected = Err.Number Then
            Console.WriteLine( "passed" )
        Else
            Console.WriteLine( "FAILED Error: " & CStr(Err.Number) & "  Expected " & CStr(iErrExpected))
        End If
    End Sub


    Sub Passed()
        Console.WriteLine("passed")
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


    Sub TestTimeSerial()
        ON Error Resume Next
        Console.WriteLine( "*** TimeSerial tests")
	    Console.Write( "1) ")
        PassFail(TimeSerial(12, 0, 1) = New DateTime(1,1,1,12,0,1,0))
    End Sub

    Sub TestDateSerial()
    End Sub


    Sub RegressionTests()
        Bug31387
    End Sub

    Sub Bug31387()
        Dim s As String

        SetBugID(31387)

        On Error Resume Next
        s = Nothing
        Call DateDiff(s,Now,Now)
        PrintResult (Err.Number = 5)
    End Sub


    Sub TestWeekday()

        Console.WriteLine( "*** Weekday Tests")

        SetBugID(31381)
	    PrintResult (vbTuesday = 3)

        'SetBugID(31381)
	    'PrintResult (weekday(1, vbTuesday) = 6)

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

