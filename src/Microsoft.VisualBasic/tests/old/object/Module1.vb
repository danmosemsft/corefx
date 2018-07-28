Option Strict Off

Imports System
Imports Microsoft.VisualBasic
Imports System.Globalization

Namespace Conversion


    Module Module1

        Dim m_lBugID As Long

        Dim sh As Short
        Dim i As Integer
        Dim l As Long
        Dim sng As Single
        Dim dbl As Double
        Dim s As String
        Dim obj As Object
        Dim b As Boolean
        Dim byt As Byte
        Dim dt As Date
        Dim vMissing As Object
        Dim dec As Decimal
        Dim c As Char
        Dim TestCounter As Integer

        Sub Main()
            console.writeline(systemtypename("boolean"))
            ' Force the culture and date formats to be the same to avoid locale diffs
            MyCulture.Init(&H409)

            Console.WriteLine("Begin Tests")

            Module2.MathTests

            TestPublicFuncs()
            Bug231113.Test
            Bug235693.Test
            Bug238590.Test
            Bug239831.Test
            Bug239390.Test
            Bug238890.Test
            Bug239162.Test
            Bug239972.Test
            Bug240143.Test
            Bug240141.Test
            Bug244220.Test
            Bug244917.Test
            Bug244958.Test
            Bug246245.Test
            Bug254927.Test
            Bug287339.Test
            Bug288405.Test

            Console.WriteLine("End Tests")

        End Sub

        Sub TestPublicFuncs()

            CompareTests
            MathTests

        End Sub

        Sub CompareTests
            On Error Resume Next
            StringCompareTests
            DateCompareTests
            IntegerCompareTests
            LongCompareTests
            ShortCompareTests
            SingleCompareTests
            DoubleCompareTests
            DecimalCompareTests
        End Sub

        Sub MathTests
            On Error Resume Next
            StringMathTests
            DateMathTests
            IntegerMathTests
            ByteMathTests
            BooleanMathTests
            LongMathTests
            ShortMathTests
            SingleMathTests
            DoubleMathTests
            DecimalMathTests
        End Sub


        Sub StringCompareTests
            Console.WriteLine()
            Console.WriteLine("*** StringCompareTests")

            Dim s1, s2, s3, s4 As String
            Dim o1, o2, o3, o4 As Object
            Dim b As Boolean

            s1 = "ABC"
            s2 = "abc"
            s3 = "tRuE"
            s4 = "FaLsE"

            o1 = s1
            o2 = s2
            o3 = s3
            o4 = s4

            Console.Write("1) ")
            Try
                b = CBool(o1 <> o2)
                PassFail(b)
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("2) ")
            Try
                b = CBool(o1 = o2)
                PassFail(Not b)
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("3) ")
            Try
                o3 = "tRuE"
                PassFail( o3 = True )
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("4) ")
            Try
                o3 = "tRuE"
                PassFail( True = o3 )
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("5) ")
            Try
                o4 = "FaLsE"
                PassFail( o4 = False )
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("6) ")
            Try
                o4 = "FaLsE"
                PassFail( False = o4 )
            Catch ex As Exception
                Failed(ex)
            End Try

            'These should throw exceptions
            Console.Write("7) ")
            Try
                o2 = "abc"
                b = (o2 = False)
                PassFail(False)
            Catch ex As InvalidCastException
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("8) ")
            Try
                o2 = "abc"
                b = (False = o2)
                PassFail(False)
            Catch ex As InvalidCastException
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try
    End Sub

        Sub DateCompareTests
            Console.WriteLine()
            Console.WriteLine("*** DateCompareTests")

            Dim dt1, dt2 As Date
            Dim o1, o2 As Object

            dt1 = New Date(1,1,1,1,1,1,1)
            dt2 = New Date(1,1,1,1,1,1,1)

            o1 = dt1
            o2 = dt2

            Console.WriteLine("a) comparing equal dates")
            Console.Write("1a) ")
            b = CBool(o1 = o2)
            PassFail(b)

            Console.Write("2a) ")
            b = CBool(o1 <> o2)
            PassFail(Not b)

            Console.Write("3a) ")
            b = CBool(o1 < o2)
            PassFail(Not b)

            Console.Write("4a) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5a) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6a) ")
            b = CBool(o1 >= o2)
            PassFail(b)



            dt1 = New Date(1999,1,1,1,1,1,1)
            dt2 = New Date(2000,1,1,1,1,1,1)

            o1 = dt1
            o2 = dt2

            Console.WriteLine("b) comparing different values")
            Console.Write("1b) ")
            b = CBool(o1 = o2)
            PassFail(Not b)

            Console.Write("2b) ")
            b = CBool(o1 <> o2)
            PassFail(b)

            Console.Write("3b) ")
            b = CBool(o1 < o2)
            PassFail(b)

            Console.Write("4b) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5b) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6b) ")
            b = CBool(o1 >= o2)
            PassFail(Not b)

        End Sub

        Sub IntegerCompareTests
            Console.WriteLine()
            Console.WriteLine("*** IntegerCompareTests")

            Dim o1, o2 As Object

            o1 = Int32.MaxValue
            o2 = Int32.MaxValue

            Console.WriteLine("a) comparing equal dates")
            Console.Write("1a) ")
            b = CBool(o1 = o2)
            PassFail(b)

            Console.Write("2a) ")
            b = CBool(o1 <> o2)
            PassFail(Not b)

            Console.Write("3a) ")
            b = CBool(o1 < o2)
            PassFail(Not b)

            Console.Write("4a) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5a) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6a) ")
            b = CBool(o1 >= o2)
            PassFail(b)



            o1 = Int32.MinValue
            o2 = Int32.MaxValue

            Console.WriteLine("b) comparing different values")
            Console.Write("1b) ")
            b = CBool(o1 = o2)
            PassFail(Not b)

            Console.Write("2b) ")
            b = CBool(o1 <> o2)
            PassFail(b)

            Console.Write("3b) ")
            b = CBool(o1 < o2)
            PassFail(b)

            Console.Write("4b) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5b) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6b) ")
            b = CBool(o1 >= o2)
            PassFail(Not b)
        End Sub

        Sub LongCompareTests
            Console.WriteLine()
            Console.WriteLine("*** LongCompareTests")

            Dim o1, o2 As Object

            o1 = Int64.MaxValue
            o2 = Int64.MaxValue

            Console.WriteLine("a) comparing equal values")
            Console.Write("1a) ")
            b = CBool(o1 = o2)
            PassFail(b)

            Console.Write("2a) ")
            b = CBool(o1 <> o2)
            PassFail(Not b)

            Console.Write("3a) ")
            b = CBool(o1 < o2)
            PassFail(Not b)

            Console.Write("4a) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5a) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6a) ")
            b = CBool(o1 >= o2)
            PassFail(b)



            o1 = Int64.MinValue
            o2 = Int64.MaxValue

            Console.WriteLine("b) comparing different values")
            Console.Write("1b) ")
            b = CBool(o1 = o2)
            PassFail(Not b)

            Console.Write("2b) ")
            b = CBool(o1 <> o2)
            PassFail(b)

            Console.Write("3b) ")
            b = CBool(o1 < o2)
            PassFail(b)

            Console.Write("4b) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5b) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6b) ")
            b = CBool(o1 >= o2)
            PassFail(Not b)
        End Sub

        Sub ShortCompareTests
            Console.WriteLine()
            Console.WriteLine("*** ShortCompareTests")

            Dim o1, o2 As Object

            o1 = Int16.MaxValue
            o2 = Int16.MaxValue

            Console.WriteLine("a) comparing equal values")
            Console.Write("1a) ")
            b = CBool(o1 = o2)
            PassFail(b)

            Console.Write("2a) ")
            b = CBool(o1 <> o2)
            PassFail(Not b)

            Console.Write("3a) ")
            b = CBool(o1 < o2)
            PassFail(Not b)

            Console.Write("4a) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5a) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6a) ")
            b = CBool(o1 >= o2)
            PassFail(b)



            o1 = Int16.MinValue
            o2 = Int16.MaxValue

            Console.WriteLine("b) comparing different values")
            Console.Write("1b) ")
            b = CBool(o1 = o2)
            PassFail(Not b)

            Console.Write("2b) ")
            b = CBool(o1 <> o2)
            PassFail(b)

            Console.Write("3b) ")
            b = CBool(o1 < o2)
            PassFail(b)

            Console.Write("4b) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5b) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6b) ")
            b = CBool(o1 >= o2)
            PassFail(Not b)
        End Sub

        Sub SingleCompareTests
            Console.WriteLine()
            Console.WriteLine("*** SingleCompareTests")

            Dim o1, o2 As Object

            o1 = Single.MaxValue
            o2 = Single.MaxValue

            Console.WriteLine("a) comparing equal values")
            Console.Write("1a) ")
            b = CBool(o1 = o2)
            PassFail(b)

            Console.Write("2a) ")
            b = CBool(o1 <> o2)
            PassFail(Not b)

            Console.Write("3a) ")
            b = CBool(o1 < o2)
            PassFail(Not b)

            Console.Write("4a) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5a) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6a) ")
            b = CBool(o1 >= o2)
            PassFail(b)



            o1 = Single.MinValue
            o2 = Single.MaxValue

            Console.WriteLine("b) comparing different values")
            Console.Write("1b) ")
            b = CBool(o1 = o2)
            PassFail(Not b)

            Console.Write("2b) ")
            b = CBool(o1 <> o2)
            PassFail(b)

            Console.Write("3b) ")
            b = CBool(o1 < o2)
            PassFail(b)

            Console.Write("4b) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5b) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6b) ")
            b = CBool(o1 >= o2)
            PassFail(Not b)
        End Sub

        Sub DoubleCompareTests
            Console.WriteLine()
            Console.WriteLine("*** DoubleCompareTests")

            Dim o1, o2 As Object

            o1 = Double.MaxValue
            o2 = Double.MaxValue

            Console.WriteLine("a) comparing equal values")
            Console.Write("1a) ")
            b = CBool(o1 = o2)
            PassFail(b)

            Console.Write("2a) ")
            b = CBool(o1 <> o2)
            PassFail(Not b)

            Console.Write("3a) ")
            b = CBool(o1 < o2)
            PassFail(Not b)

            Console.Write("4a) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5a) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6a) ")
            b = CBool(o1 >= o2)
            PassFail(b)



            o1 = Double.MinValue
            o2 = Double.MaxValue

            Console.WriteLine("b) comparing different values")
            Console.Write("1b) ")
            b = CBool(o1 = o2)
            PassFail(Not b)

            Console.Write("2b) ")
            b = CBool(o1 <> o2)
            PassFail(b)

            Console.Write("3b) ")
            b = CBool(o1 < o2)
            PassFail(b)

            Console.Write("4b) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5b) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6b) ")
            b = CBool(o1 >= o2)
            PassFail(Not b)
        End Sub

        Sub DecimalCompareTests
            Console.WriteLine()
            Console.WriteLine("*** DecimalCompareTests")

            Dim o1, o2 As Object

            o1 = Decimal.MaxValue
            o2 = Decimal.MaxValue

            Console.WriteLine("a) comparing equal values")
            Console.Write("1a) ")
            b = CBool(o1 = o2)
            PassFail(b)

            Console.Write("2a) ")
            b = CBool(o1 <> o2)
            PassFail(Not b)

            Console.Write("3a) ")
            b = CBool(o1 < o2)
            PassFail(Not b)

            Console.Write("4a) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5a) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6a) ")
            b = CBool(o1 >= o2)
            PassFail(b)



            o1 = Decimal.MinValue
            o2 = Decimal.MaxValue

            Console.WriteLine("b) comparing different values")
            Console.Write("1b) ")
            b = CBool(o1 = o2)
            PassFail(Not b)

            Console.Write("2b) ")
            b = CBool(o1 <> o2)
            PassFail(b)

            Console.Write("3b) ")
            b = CBool(o1 < o2)
            PassFail(b)

            Console.Write("4b) ")
            b = CBool(o1 > o2)
            PassFail(Not b)

            Console.Write("5b) ")
            b = CBool(o1 <= o2)
            PassFail(b)

            Console.Write("6b) ")
            b = CBool(o1 >= o2)
            PassFail(Not b)
        End Sub


        Sub StringMathTests()
            Console.WriteLine()
            Console.WriteLine("*** StringMathTests")
        End Sub

        Sub DateMathTests()
            Console.WriteLine()
            Console.WriteLine("*** DateMathTests")
        End Sub

        Sub IntegerMathTests()
            Console.WriteLine()
            Console.WriteLine("*** IntegerMathTests")

            Dim o1, o2, oResult, oExpected As Object
            Dim int1, int2 As Integer

            Try
                int1 = 0
                o1 = int1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ int1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- int1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                int1 = &H0808
                o1 = int1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ int1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- int1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                int1 = &H8080
                o1 = int1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ int1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- int1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 0, o2 = 0")
                int1 = 0
                int2 = 0
                o1 = int1
                o2 = int2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (int1 + int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (int1 - int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (int1 * int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (int1 \ int2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (int1 / int2)
                PassFail(Double.IsNaN(oResult) AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                Try
                    oResult = (o1 Mod o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (int1 Mod int2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator ^   ")
                PassFail((o1 ^ o2) = (int1 ^ int2))
            Catch ex As Exception
                Failed(ex)
            End Try
        

            Try
                Console.WriteLine("   o1 = 0, o2 = Int32.MaxValue")
                int1 = 0
                int2 = Int32.MaxValue
                o1 = int1
                o2 = int2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (int1 + int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (int1 - int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (int1 * int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (int1 / int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (int1 \ int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (int1 Mod int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (int1 ^ int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 0, o2 = Int32.MinValue")
                int1 = 0
                int2 = Int32.MinValue
                o1 = int1
                o2 = int2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (int1 + int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (CLng(int1) - CLng(int2))
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (int1 * int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (int1 / int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (int1 \ int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (int1 Mod int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (int1 ^ int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Int32.MaxValue, o2 = 1")
                int1 = Int32.MaxValue
                int2 = 1
                o1 = int1
                o2 = int2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = CLng(int1) + CLng(int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (int1 - int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = o1 * o2 : oExpected = (int1 * int2)
                PassFail((o1 * o2) = (int1 * int2) AndAlso TypeName(oResult) = TypeName(int1 * int2))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (int1 / int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (int1 \ int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (int1 Mod int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (int1 ^ int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Int32.MinValue, o2 = 1")
                int1 = Int32.MinValue
                int2 = 1
                o1 = int1
                o2 = int2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (int1 + int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (CLng(int1) - CLng(int2))
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (int1 * int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (int1 / int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (int1 \ int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (int1 Mod int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (int1 ^ int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Nothing, o2 = 1")
                int1 = CInt(Nothing)
                int2 = 1
                o1 = Nothing
                o2 = int2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (int1 + int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (int1 - int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (int1 * int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (int1 / int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (int1 \ int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (int1 Mod int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (int1 ^ int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 1, o2 = Nothing")
                int1 = 1
                int2 = CInt(Nothing)
                o1 = int1
                o2 = Nothing

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (int1 + int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (int1 - int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (int1 * int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (int1 / int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (int1 \ int2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator Mod ")
                Try
                    oResult = (o1 Mod o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (int1 Mod int2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (int1 ^ int2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

        End Sub

        Sub LongMathTests()
            Console.WriteLine()
            Console.WriteLine("*** LongMathTests")

            Dim o1, o2, oResult, oExpected As Object
            Dim lng1, lng2 As Long

            Try
                lng1 = 0
                o1 = lng1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ lng1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- lng1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                lng1 = &H0808
                o1 = lng1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ lng1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- lng1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                lng1 = &H8080
                o1 = lng1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ lng1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- lng1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 0, o2 = 0")
                lng1 = 0
                lng2 = 0
                o1 = lng1
                o2 = lng2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (lng1 + lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (lng1 - lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (lng1 * lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (lng1 \ lng2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (lng1 / lng2)
                PassFail(Double.IsNaN(oResult) AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                Try
                    oResult = (o1 Mod o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (lng1 Mod lng2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator ^   ")
                PassFail((o1 ^ o2) = (lng1 ^ lng2))
            Catch ex As Exception
                Failed(ex)
            End Try
        

            Try
                Console.WriteLine("   o1 = 0, o2 = Int64.MaxValue")
                lng1 = 0
                lng2 = Int64.MaxValue
                o1 = lng1
                o2 = lng2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (lng1 + lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (lng1 - lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (lng1 * lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (lng1 / lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (lng1 \ lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (lng1 Mod lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (lng1 ^ lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Nothing, o2 = Int64.MaxValue")
                lng1 = CLng(Nothing)
                lng2 = Int64.MaxValue
                o1 = Nothing
                o2 = lng2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (lng1 + lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (lng1 - lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (lng1 * lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (lng1 / lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (lng1 \ lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (lng1 Mod lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (lng1 ^ lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 0, o2 = Int64.MinValue")
                lng1 = 0
                lng2 = Int64.MinValue
                o1 = lng1
                o2 = lng2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (lng1 + lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (CDec(lng1) - CDec(lng2))
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (lng1 * lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (lng1 / lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (lng1 \ lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (lng1 Mod lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (lng1 ^ lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Int64.MaxValue, o2 = 1")
                lng1 = Int64.MaxValue
                lng2 = 1
                o1 = lng1
                o2 = lng2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = CDec(lng1) + CDec(lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (lng1 - lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = o1 * o2 : oExpected = (lng1 * lng2)
                PassFail((o1 * o2) = (lng1 * lng2) AndAlso TypeName(oResult) = TypeName(lng1 * lng2))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (lng1 / lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (lng1 \ lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (lng1 Mod lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (lng1 ^ lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Int64.MinValue, o2 = 1")
                lng1 = Int64.MinValue
                lng2 = 1
                o1 = lng1
                o2 = lng2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (lng1 + lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (CDec(lng1) - CDec(lng2))
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (lng1 * lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (lng1 / lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (lng1 \ lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (lng1 Mod lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (lng1 ^ lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Nothing, o2 = 1")
                lng1 = CLng(Nothing)
                lng2 = 1
                o1 = Nothing
                o2 = lng2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (lng1 + lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (lng1 - lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (lng1 * lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (lng1 / lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (lng1 \ lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (lng1 Mod lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (lng1 ^ lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 1, o2 = Nothing")
                lng1 = 1
                lng2 = CLng(Nothing)
                o1 = lng1
                o2 = Nothing

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (lng1 + lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (lng1 - lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (lng1 * lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (lng1 / lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (lng1 \ lng2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator Mod ")
                Try
                    oResult = (o1 Mod o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (lng1 Mod lng2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (lng1 ^ lng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub

        Sub ShortMathTests()
            Console.WriteLine()
            Console.WriteLine("*** ShortMathTests")

            Dim o1, o2, oResult, oExpected As Object
            Dim sh1, sh2 As Short

            Try
                sh1 = 0
                o1 = sh1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ sh1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- sh1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                sh1 = &H0808S
                o1 = sh1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ sh1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- sh1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                sh1 = &H8080S
                o1 = sh1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ sh1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- sh1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 0, o2 = 0")
                sh1 = 0S
                sh2 = 0S
                o1 = sh1
                o2 = sh2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (sh1 + sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (sh1 - sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (sh1 * sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (sh1 \ sh2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (sh1 / sh2)
                PassFail(Double.IsNaN(oResult) AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                Try
                    oResult = (o1 Mod o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (sh1 Mod sh2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator ^   ")
                PassFail((o1 ^ o2) = (sh1 ^ sh2))
            Catch ex As Exception
                Failed(ex)
            End Try
        

            Try
                Console.WriteLine("   o1 = 0, o2 = Int16.MaxValue")
                sh1 = 0S
                sh2 = Int16.MaxValue
                o1 = sh1
                o2 = sh2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (sh1 + sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (sh1 - sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (sh1 * sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (sh1 / sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (sh1 \ sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (sh1 Mod sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (sh1 ^ sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Nothing, o2 = Int16.MaxValue")
                sh1 = CShort(Nothing)
                sh2 = Int16.MaxValue
                o1 = Nothing
                o2 = sh2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (sh1 + sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (sh1 - sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (sh1 * sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (sh1 / sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (sh1 \ sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (sh1 Mod sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (sh1 ^ sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 0, o2 = Short.MinValue")
                sh1 = 0S
                sh2 = Short.MinValue
                o1 = sh1
                o2 = sh2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (sh1 + sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (CInt(sh1) - CInt(sh2))
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (sh1 * sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (sh1 / sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (sh1 \ sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (sh1 Mod sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (sh1 ^ sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Int16.MaxValue, o2 = 1")
                sh1 = Int16.MaxValue
                sh2 = 1S
                o1 = sh1
                o2 = sh2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = CInt(sh1) + CInt(sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (sh1 - sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = o1 * o2 : oExpected = (sh1 * sh2)
                PassFail((o1 * o2) = (sh1 * sh2) AndAlso TypeName(oResult) = TypeName(sh1 * sh2))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (sh1 / sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (sh1 \ sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (sh1 Mod sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (sh1 ^ sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Short.MinValue, o2 = 1")
                sh1 = Short.MinValue
                sh2 = 1S
                o1 = sh1
                o2 = sh2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (sh1 + sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (CInt(sh1) - CInt(sh2))
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (sh1 * sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (sh1 / sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (sh1 \ sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (sh1 Mod sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (sh1 ^ sh2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

        End Sub


        Sub ByteMathTests()
            Console.WriteLine()
            Console.WriteLine("*** ByteMathTests")

            Dim o1, o2, oResult, oExpected As Object
            Dim by1, by2 As Byte

            Try
                Console.WriteLine("   o1 = CByte(0)")
                by1 = 0
                o1 = by1
                    Console.Write("      unary +   ")
                    oResult = (+ o1) : oExpected = (+ by1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("      unary -   ")
                    oResult = (- o1) : oExpected = (- by1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.WriteLine("   o1 = CByte(&H08)")
                by1 = &H08
                o1 = by1
                    Console.Write("      unary +   ")
                    oResult = (+ o1) : oExpected = (+ by1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("      unary -   ")
                    oResult = (- o1) : oExpected = (- by1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.WriteLine("   o1 = CByte(&H80)")
                by1 = &H80
                o1 = by1
                    Console.Write("      unary +   ")
                    oResult = (+ o1) : oExpected = (+ by1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("      unary -   ")
                    oResult = (- o1) : oExpected = (- by1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 0, o2 = 0")
                by1 = 0
                by2 = 0
                o1 = by1
                o2 = by2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (by1 + by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (by1 - by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (by1 * by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (by1 \ by2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (by1 / by2)
                PassFail(Double.IsNaN(oResult) AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                Try
                    oResult = (o1 Mod o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (by1 Mod by2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator ^   ")
                PassFail((o1 ^ o2) = (by1 ^ by2))
            Catch ex As Exception
                Failed(ex)
            End Try
        

            Try
                Console.WriteLine("   o1 = 0, o2 = Byte.MaxValue")
                by1 = 0
                by2 = Byte.MaxValue
                o1 = by1
                o2 = by2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (by1 + by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (CShort(by1) - CShort(by2))
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (by1 * by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (by1 / by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (by1 \ by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (by1 Mod by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (by1 ^ by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Nothing, o2 = Byte.MaxValue")
                by1 = CByte(Nothing)
                by2 = Byte.MaxValue
                o1 = Nothing
                o2 = by2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (by1 + by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (CShort(by1) - CShort(by2))
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (by1 * by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (by1 / by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (by1 \ by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (by1 Mod by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (by1 ^ by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try


            Try
                Console.WriteLine("   o1 = Byte.MaxValue, o2 = Nothing")
                by1 = Byte.MaxValue
                by2 = CByte(Nothing)
                o1 = by1
                o2 = Nothing

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (by1 + by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (by1 - CByte(o2))
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (by1 * by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (by1 / by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2) 
                    Console.WriteLine("Exception should have been thrown")
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (by1 \ by2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator Mod ")
                Try
                    oResult = (o1 Mod o2) 
                    Console.WriteLine("Exception should have been thrown")
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (by1 Mod by2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (by1 ^ by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try


            '  ********* Byte.MinValue is zero, so no need to test 0 and Byte.MinValue
            '            since it was done above

            Try
                Console.WriteLine("   o1 = Byte.MaxValue, o2 = 1")
                by1 = Byte.MaxValue
                by2 = 1S
                o1 = by1
                o2 = by2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = CShort(by1) + CShort(by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (by1 - by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = o1 * o2 : oExpected = (by1 * by2)
                PassFail((o1 * o2) = (by1 * by2) AndAlso TypeName(oResult) = TypeName(by1 * by2))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (by1 / by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (by1 \ by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (by1 Mod by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (by1 ^ by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Byte.MinValue, o2 = 1")
                by1 = Byte.MinValue
                by2 = 1
                o1 = by1
                o2 = by2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (by1 + by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (CShort(by1) - CShort(by2))
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (by1 * by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (by1 / by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (by1 \ by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (by1 Mod by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (by1 ^ by2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

        End Sub


        Sub BooleanMathTests()
            Console.WriteLine()
            Console.WriteLine("*** BooleanMathTests")
            Dim o1, o2, oResult, oExpected As Object
            Dim b1, b2 As Boolean

            Try
                b1 = False
                o1 = b1
                    Console.Write("   unary (+ False) ")
                    oResult = (+ o1) : oExpected = (+ b1)
                    PassFail(TypeOf oResult Is Short AndAlso TypeOf oExpected Is Short AndAlso oResult = oExpected)

                    Console.Write("   unary (- False) ")
                    oResult = (- o1) : oExpected = (- b1)
                    PassFail(TypeOf oResult Is Short AndAlso TypeOf oExpected Is Short AndAlso oResult = oExpected)

                b1 = True
                o1 = b1
                    Console.Write("   unary (+ True)  ")
                    oResult = (+ o1) : oExpected = (+ b1)
                    PassFail(TypeOf oResult Is Short AndAlso TypeOf oExpected Is Short AndAlso oResult = oExpected)

                    Console.Write("   unary (- True)  ")
                    oResult = (- o1) : oExpected = (- b1)
                    PassFail(TypeOf oResult Is Short AndAlso TypeOf oExpected Is Short AndAlso oResult = oExpected)

            Catch ex As Exception
                Failed(ex)
            End Try
                        
            Try
                Console.WriteLine("   b1 = False, b2 = False")

                b1 = False
                b2 = False
                o1 = b1
                o2 = b2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (b1 + b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (b1 - b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (b1 * b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (b1 / b2)
                PassFail(Double.IsNaN(oResult) AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2) 
                    Console.WriteLine("Exception should have been thrown")
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (b1 \ b2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator Mod ")
                Try
                    oResult = (o1 Mod o2) 
                    Console.WriteLine("Exception should have been thrown")
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (b1 Mod b2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (b1 ^ b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   b1 = False, b2 = True")
                b1 = False
                b2 = True
                o1 = b1
                o2 = b2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (b1 + b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (CShort(b1) - CShort(b2))
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (b1 * b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (b1 / b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (b1 \ b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (b1 Mod b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (b1 ^ b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   b1 = True, b2 = False")
                b1 = True
                b2 = False
                o1 = b1
                o2 = b2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (b1 + b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (b1 - b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (b1 * b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (b1 / b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2) 
                    Console.WriteLine("Exception should have been thrown")
                    PassFail(False)
                Catch ex As Exception

                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (b1 \ b2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try

                End Try

                Console.Write("      operator Mod ")
                Try
                    oResult = (o1 Mod o2) 
                    Console.WriteLine("Exception should have been thrown")
                    PassFail(False)
                Catch ex As Exception

                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (b1 Mod b2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try

                End Try

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (b1 ^ b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   b1 = True, b2 = True")
                b1 = True
                b2 = True
                o1 = b1
                o2 = b2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (b1 + b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (b1 - b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (b1 * b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (b1 / b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                oResult = (o1 \ o2) : oExpected = (b1 \ b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (b1 Mod b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (b1 ^ b2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

        End Sub

        Sub SingleMathTests()
            Console.WriteLine()
            Console.WriteLine("*** SingleMathTests")

            Dim o1, o2, oResult, oExpected As Object
            Dim sng1, sng2 As Single

            Try
                sng1 = 0
                o1 = sng1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ sng1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- sng1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                sng1 = Single.MaxValue
                o1 = sng1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ sng1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- sng1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                sng1 = Single.MinValue
                o1 = sng1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ sng1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- sng1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 0, o2 = 0")
                sng1 = 0
                sng2 = 0
                o1 = sng1
                o2 = sng2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (sng1 + sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (sng1 - sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (sng1 * sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (sng1 \ sng2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (sng1 / sng2)
                PassFail(Double.IsNaN(oResult) AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (sng1 Mod sng2)
                PassFail(Single.IsNaN(oResult) AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (sng1 ^ sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try
        

            Try
                Console.WriteLine("   o1 = 0, o2 = Single.MaxValue")
                sng1 = 0
                sng2 = Single.MaxValue
                o1 = sng1
                o2 = sng2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (sng1 + sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (sng1 - sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (sng1 * sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (sng1 / sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (sng1 \ sng2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (sng1 Mod sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (sng1 ^ sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 0, o2 = Single.MinValue")
                sng1 = 0
                sng2 = Single.MinValue
                o1 = sng1
                o2 = sng2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (sng1 + sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (sng1 - sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (sng1 * sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (sng1 / sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (sng1 \ sng2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (sng1 Mod sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (sng1 ^ sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Single.MaxValue, o2 = 1")
                sng1 = Single.MaxValue
                sng2 = 1
                o1 = sng1
                o2 = sng2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (sng1 + sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (sng1 - sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = o1 * o2 : oExpected = (sng1 * sng2)
                PassFail((o1 * o2) = (sng1 * sng2) AndAlso TypeName(oResult) = TypeName(sng1 * sng2))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (sng1 / sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (sng1 \ sng2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (sng1 Mod sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (sng1 ^ sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Single.MinValue, o2 = 1")
                sng1 = Single.MinValue
                sng2 = 1
                o1 = sng1
                o2 = sng2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (sng1 + sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (sng1 - sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = CSng(CDbl(sng1) * CDbl(sng2))
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (sng1 / sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (sng1 \ sng2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (sng1 Mod sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (sng1 ^ sng2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub


        Sub DoubleMathTests()
            Console.WriteLine()
            Console.WriteLine("*** DoubleMathTests")

            Dim o1, o2, oResult, oExpected As Object
            Dim dbl1, dbl2 As Double

            Try
                dbl1 = 0
                o1 = dbl1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ dbl1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- dbl1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                dbl1 = Double.MaxValue
                o1 = dbl1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ dbl1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- dbl1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                dbl1 = Double.MinValue
                o1 = dbl1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ dbl1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- dbl1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 0, o2 = 0")
                dbl1 = 0
                dbl2 = 0
                o1 = dbl1
                o2 = dbl2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (dbl1 + dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (dbl1 - dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (dbl1 * dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (dbl1 \ dbl2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (dbl1 / dbl2)
                PassFail(Double.IsNaN(oResult) AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (dbl1 Mod dbl2)
                PassFail(Double.IsNaN(oResult) AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (dbl1 ^ dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try
        

            Try
                Console.WriteLine("   o1 = 0, o2 = Double.MaxValue")
                dbl1 = 0
                dbl2 = Double.MaxValue
                o1 = dbl1
                o2 = dbl2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (dbl1 + dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (dbl1 - dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (dbl1 * dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (dbl1 / dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (dbl1 \ dbl2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (dbl1 Mod dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (dbl1 ^ dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 0, o2 = Double.MinValue")
                dbl1 = 0
                dbl2 = Double.MinValue
                o1 = dbl1
                o2 = dbl2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (dbl1 + dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (dbl1 - dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (dbl1 * dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (dbl1 / dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (dbl1 \ dbl2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (dbl1 Mod dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (dbl1 ^ dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Double.MaxValue, o2 = 1")
                dbl1 = Double.MaxValue
                dbl2 = 1
                o1 = dbl1
                o2 = dbl2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (dbl1 + dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (dbl1 - dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = o1 * o2 : oExpected = (dbl1 * dbl2)
                PassFail((o1 * o2) = (dbl1 * dbl2) AndAlso TypeName(oResult) = TypeName(dbl1 * dbl2))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (dbl1 / dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (dbl1 \ dbl2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (dbl1 Mod dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (dbl1 ^ dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Double.MinValue, o2 = 1")
                dbl1 = Double.MinValue
                dbl2 = 1
                o1 = dbl1
                o2 = dbl2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (dbl1 + dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (dbl1 - dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (dbl1 * dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (dbl1 / dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (dbl1 \ dbl2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (dbl1 Mod dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (dbl1 ^ dbl2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub

        Sub DecimalMathTests()
            Console.WriteLine()
            Console.WriteLine("*** DecimalMathTests")

            Dim o1, o2, oResult, oExpected As Object
            Dim dec1, dec2 As Decimal

            Try
                dec1 = 0
                o1 = dec1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ dec1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- dec1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                dec1 = Decimal.MaxValue
                o1 = dec1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ dec1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- dec1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                dec1 = Decimal.MinValue
                o1 = dec1
                    Console.Write("         unary +   ")
                    oResult = (+ o1) : oExpected = (+ dec1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                    Console.Write("         unary -   ")
                    oResult = (- o1) : oExpected = (- dec1)
                    PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 0, o2 = 0")
                dec1 = 0@
                dec2 = 0@
                o1 = dec1
                o2 = dec2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (dec1 + dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (dec1 - dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (dec1 * dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                Try
                    oResult = (o1 / o2) 
                    Console.WriteLine("Exception should have been thrown")
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (dec1 / dec2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try

                End Try

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2) 
                    Console.WriteLine("Exception should have been thrown")
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (dec1 \ dec2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try

                End Try

                Console.Write("      operator Mod ")
                Try
                    oResult = (o1 Mod o2)
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (dec1 Mod dec2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try
                End Try

                Console.Write("      operator ^   ")
                PassFail((o1 ^ o2) = (dec1 ^ dec2))
            Catch ex As Exception
                Failed(ex)
            End Try
        

            Try
                Console.WriteLine("   o1 = 0, o2 = Decimal.MaxValue")
                dec1 = 0@
                dec2 = Decimal.MaxValue
                o1 = dec1
                o2 = dec2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (dec1 + dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (dec1 - dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (dec1 * dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (dec1 / dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2) 
                    Console.WriteLine("Exception should have been thrown")
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (dec1 \ dec2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try

                End Try

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (dec1 Mod dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (dec1 ^ dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = 0, o2 = Decimal.MinValue")
                dec1 = 0@
                dec2 = Decimal.MinValue
                o1 = dec1
                o2 = dec2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (dec1 + dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (dec1 - dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (dec1 * dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (dec1 / dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2) 
                    Console.WriteLine("Exception should have been thrown")
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (dec1 \ dec2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try

                End Try

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (dec1 Mod dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (dec1 ^ dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Decimal.MaxValue, o2 = 1")
                dec1 = Decimal.MaxValue
                dec2 = 1@
                o1 = dec1
                o2 = dec2
#if 0
                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = CSng(dec1) + CSng(dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))
#Else
                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = CDbl(dec1) + CDbl(dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))
#End If
                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (dec1 - dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = o1 * o2 : oExpected = (dec1 * dec2)
                PassFail((o1 * o2) = (dec1 * dec2) AndAlso TypeName(oResult) = TypeName(dec1 * dec2))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (dec1 / dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2) 
                    Console.WriteLine("Exception should have been thrown")
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (dec1 \ dec2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try

                End Try

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (dec1 Mod dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (dec1 ^ dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.WriteLine("   o1 = Decimal.MinValue, o2 = 1")
                dec1 = Decimal.MinValue
                dec2 = 1@
                o1 = dec1
                o2 = dec2

                Console.Write("      operator +   ")
                oResult = (o1 + o2) : oExpected = (dec1 + dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator -   ")
                oResult = (o1 - o2) : oExpected = (CDbl(dec1) + CDbl(dec2))
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator *   ")
                oResult = (o1 * o2) : oExpected = (dec1 * dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator /   ")
                oResult = (o1 / o2) : oExpected = (dec1 / dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator \   ")
                Try
                    oResult = (o1 \ o2) 
                    Console.WriteLine("Exception should have been thrown")
                    PassFail(False)
                Catch ex As Exception
                    Dim ExceptionType As System.Type
                    ExceptionType = ex.GetType()
                    Try
                        oResult = (dec1 \ dec2)
                        Console.WriteLine("Exception should have been thrown")
                        PassFail(False)
                    Catch ex2 As Exception
                        If ExceptionType Is ex2.GetType() Then
                            PassFail(True)
                        Else
                            PassFail(False)
                        End If
                    End Try

                End Try

                Console.Write("      operator Mod ")
                oResult = (o1 Mod o2) : oExpected = (dec1 Mod dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

                Console.Write("      operator ^   ")
                oResult = (o1 ^ o2) : oExpected = (dec1 ^ dec2)
                PassFail(oResult = oExpected AndAlso TypeName(oResult) = TypeName(oExpected))

            Catch ex As Exception
                Failed(ex)
            End Try

        End Sub

        Sub CObjTest()
            Console.WriteLine()
            Console.WriteLine("*** CObj tests")

            obj = sh    : Console.Write("1) ") : PassFail(TypeOf obj Is Short)
            obj = i     : Console.Write("2) ") : PassFail(TypeOf obj Is Integer)
            obj = l     : Console.Write("3) ") : PassFail(TypeOf obj Is Long)
            obj = sng   : Console.Write("4) ") : PassFail(TypeOf obj Is Single)
            obj = dbl   : Console.Write("5) ") : PassFail(TypeOf obj Is Double)
            obj = dec   : Console.Write("6) ") : PassFail(TypeOf obj Is Decimal)
            obj = s     : Console.Write("7) ") : PassFail(TypeOf obj Is String)
            obj = b     : Console.Write("8) ") : PassFail(TypeOf obj Is Boolean)
            obj = dt    : Console.Write("9) ") : PassFail(TypeOf obj Is Date)
            obj = c     : Console.Write("10) ") : PassFail(TypeOf obj Is Char)
        End Sub

        Sub ErrorCheck(ByVal iErrExpected As Long)
            If iErrExpected = Err.Number Then
                Console.WriteLine("PASSED")
            Else
                Console.WriteLine("FAILED Error: " & CStr(Err.Number) & "  Expected " & CStr(iErrExpected))
            End If
            m_lBugID = 0
        End Sub


        Sub SetBugID(ByVal lID As Long, Optional ByVal lExceptedErr As Long = 0)
            If m_lBugID <> 0 Then
                Console.WriteLine("fail (unexpected exception): " & lID)
            End If
            m_lBugID = lID
            Console.Write("Bug" &  CStr(m_lBugID) & ": " )
        End Sub

        Sub PassFail(ByVal bPassed As Boolean)
            If bPassed Then
                Console.WriteLine("passed")
            Else
                Console.WriteLine("FAILED !!!")
            End If
        End Sub

        Sub Passed()
            Console.WriteLine("passed")
        End Sub
        Sub Failed()
            Console.WriteLine("FAILED !!!")
        End Sub

        Sub Failed(ByVal ex As Exception)
            If ex Is Nothing Then
                Console.WriteLine("NULL System.Exception")
            Else
                Console.WriteLine(ex.GetType().FullName & ": " & ex.Message)
            End If
            Console.WriteLine("FAILED !!!")
        End Sub
    End Module

    Module Bug231113

        Sub Test()

            Console.Write("*** Bug 231113 : ")
            Try
                Dim y As Object, z As VariantType
                y = VariantType.Integer
                z = VariantType.Integer
                Dim x As Boolean = y OrElse z

                PassFail(x = True)

            Catch ex As Exception
                Failed(ex)
            End Try

        End Sub

    End Module

    Module Bug235693

        Sub Test()

            Console.Write("*** Bug 235693 : ")
            Try
                Dim o1, o2 As Object

                o1 = CShort(-1)
                o2 = True
                If o1 <> o2 Then
                    Throw New ArgumentException("NOT EQUAL")
                End If

                o1 = CInt(-1)
                o2 = True
                If o1 <> o2 Then
                    Throw New ArgumentException("NOT EQUAL")
                End If

                o1 = CLng(-1)
                o2 = True
                If o1 <> o2 Then
                    Throw New ArgumentException("NOT EQUAL")
                End If

                o1 = CDec(-1)
                o2 = True
                If o1 <> o2 Then
                    Throw New ArgumentException("NOT EQUAL")
                End If

                o1 = CSng(-1)
                o2 = True
                If o1 <> o2 Then
                    Throw New ArgumentException("NOT EQUAL")
                End If

                o1 = CDbl(-1)
                o2 = True
                If o1 <> o2 Then
                    Throw New ArgumentException("NOT EQUAL")
                End If

                Console.WriteLine("passed")
            Catch ex As Exception
                Failed(ex)
            End Try

        End Sub

    End Module



    Module Bug238590
    
        Sub Test()

            Console.WriteLine("*** Bug 238590 : ")
            Try
                SubTest()
            Catch ex As Exception
                Failed(ex)
            End Try

        End Sub

        Sub SubTest()
            Dim x As Object = "1224"
            Dim y As Integer = 123
            Dim z As Boolean

            Console.Write("  1) ")
            PassFail(x > y)
            z = x > y

            Dim s As String
            Dim dt As Date

            dt = #12/31/2001#
            s = dt.ToString()
            Console.Write("  2) ")
            PassFail(s = dt)

            Dim o1, o2 As Object

            o1 = s
            o2 = dt
            Console.Write("  3) ")
            PassFail(o1 = o2)

            o1 = dt
            o2 = s
            Console.Write("  4) ")
            PassFail(o1 = o2)

            o1 = "A"c
            o2 = "A"
            Console.Write("  5) ")
            PassFail(o1 = o2)

            o1 = "A"
            o2 = "A"c
            Console.Write("  6) ")
            PassFail(o1 = o2)
        End Sub

    End Module

    Module Bug238833

        Sub Test()
            Console.Write("*** Bug 238833 : ")
            Dim o1, o2 As Object

            o1 = True
            o2 = False

            Dim o As Object
            o = o1 And o2

            PassFail(TypeOf o Is Boolean)

        End Sub

    End Module

    Module Bug238890
    
        Sub Test()
            Console.Write("*** Bug 239390 : ")

            Dim y1 = -2147483647#
            Dim VARandorshort4
            Dim VARandorshort5
            Dim VARandorshort6
            Dim VARandorshort7
            VARandorshort5 = CDec(-y1 - 1)
            VARandorshort4 = CDec(-y1)
            VARandorshort7 = CDec(y1)
            VARandorshort6 = CDec(y1 + 1)

            Dim obj As Object
            obj = (VARandorshort4 < VARandorshort5) Or (VARandorshort6 = VARandorshort7)

            PassFail((TypeOf obj Is Boolean) AndAlso (obj = False))

        End Sub

    End Module


    Module Bug239390
   
        Sub Test()
            Console.Write("*** Bug 239390 : ")

            Dim a As Object
            Dim b As Object
            a = CDec("10000000000.1")
            b = CLng("10000000000")

            Dim obj1, obj2 As Object

            obj1 = a < b
            obj2 = b < a
            PassFail(obj1 = False AndAlso obj2 = True)

        End Sub


    End Module

    Module Bug239831
   
        Sub Test()
            Console.Write("*** Bug 239831 : ")

            Dim a
            Dim b
            Dim c
            Dim sng1, sng2 As Single

            sng1 = CSng(1.2341234E+31)
            sng2 = CSng(1.2341234E+34)
            a = sng1
            b = sng2
            c = a * b

            'Verify promotion to double
            PassFail((TypeOf c Is Double) AndAlso (c = (CDbl(sng1) * CDbl(sng2))))
        End Sub


    End Module

    Module Bug239162
	    Enum enumb As Byte
		    e0
		    e1
		    e2
		    e253 = 253
		    e254 = 254
		    e255 = 255
	    End Enum

	    Sub Test()

            Console.Write("*** Bug 239162 : ")

		    Dim bn10_1 As Object 
            Dim result As Object

		    bn10_1 = enumb.e1
            result = Not bn10_1
		    PassFail((TypeName(result) = "enumb") AndAlso result = 254)
	    End Sub
    End Module

    Module Bug239972

	    Sub Test()

            Console.Write("*** Bug 239972 : ")

            Dim bn9_2 As Object = CShort(22)
            Dim result As Object

            result = Not bn9_2
		    PassFail((TypeOf result Is Int16) AndAlso result = -23)
	    End Sub
    End Module

    Module Bug240141
    
        Enum Color
            Red = 1
            Green = 2
        End Enum

        Sub Test()

            Console.Write("*** Bug 240141 : ")

            'Main bound case
            Dim X1 As Color
            Dim Y1 As Integer
            X1 = Color.Green
            Y1 = 10
            Y1 ^= X1   ' No error Error

            'Late bound case
            Dim X2, Y2 As Object
            X2 = Color.Green
            Y2 = 10
            Y2 ^= X2   'Error
            PassFail((TypeName(Y2) = TypeName(10I ^ X1)) _
                AndAlso (Y2 = Y1))
        End Sub


    End Module

    Module Bug240143
    
        Enum Color
            Red = 1
            Green = 2
        End Enum

        Sub Test()
            Console.Write("*** Bug 240143 : ")

            'Early bound case
            Dim X1 As Color
            Dim Y1 As Integer
            X1 = Color.Green
            Y1 = 10

            Try
                Y1 /= X1   ' No error Error

                'Late bound case
                Dim X2, Y2 As Object
                X2 = Color.Green
                Y2 = 10

                Y2 /= X2   'Error
                PassFail((TypeName(Y2) = TypeName(10I / X1)) AndAlso (Y2 = Y1))
            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub


    End Module

    Module Bug244220
        Sub Test()
            Console.Write("*** Bug 244220 : ")
            Try
                Dim o2 As Object = Nothing
                Dim o3 As Object = 1 / o2
                Dim o4 As Object

                o4 = 1 * o2

                PassFail(o3.GetType() Is GetType(Double) AndAlso _
                    o3.ToString() = Globalization.NumberFormatInfo.CurrentInfo.PositiveInfinitySymbol AndAlso _
                    o4.GetType() Is GetType(Integer) AndAlso o4 = 0)
            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub
    End Module

    Module Bug244917

        Sub Test()
            Console.WriteLine("*** Bug 244917 ")
            Try

                Const a As Boolean = False
                Const b As Boolean = True

                Dim result As Object
                Dim result2 As Object
                Dim obj1 As Object
                Dim obj2 As Object

                obj1 = a
                obj2 = b

                result = a Or b
                result2 = obj1 Or obj2

                Console.Write("    ") : PassFail(TypeOf obj1 Is Boolean)
                Console.Write("    ") : PassFail(TypeOf obj2 Is Boolean)
                Console.Write("    ") : PassFail(TypeOf result Is Boolean)
                Console.Write("    ") : PassFail(TypeOf result2 Is Boolean)
            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub

    End Module

    Module Bug244958
        Sub Test
            Console.Write("*** Bug 244958: ")

            Dim fixarg(10) As Char
            DIm o1, o2 As Object
            Dim expected

            expected = "cartridges"
            fixarg = expected

            Try
                PassFail((fixarg = expected) AndAlso (expected = fixarg))
            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub
    End Module

    Module Bug246245

        Sub test1()
            Dim o1 As Byte = 0
            Dim o2 As Boolean = True
            Dim o3 As Object = o1
            Dim o4 As Object = o2
            Dim A, B As Object

            Console.Write("    ")

            A = o1 \ o2
            B = o3 \ o4
            PassFail(A = B AndAlso A.GetType() Is B.GetType())
        End Sub

        Sub test2()
            Dim o1 As Byte = 0
            Dim o2 As Boolean = True
            Dim o3 As Object = o1
            Dim o4 As Object = o2
            Dim A, B As Object

            Console.Write("    ")

            A = o1 Mod o2
            B = o3 Mod o4
            PassFail(A = B AndAlso A.GetType() Is B.GetType())
        End Sub

        Sub Test()
            Console.WriteLine("*** Bug246245")
            test1()
            test2()
        End Sub

    End Module

    Class Bug254927
   
        Shared Sub Test()
            Console.WriteLine("*** Bug 254927")

            Dim result As Object
            Dim xint321 As Int32

            Try
                Console.Write("    1) ")
                result = Nothing
                xint321 = 536870912
                result = xint321 + xint321
                result = xint321 + result
                result = xint321 + result
                result = xint321 + result   ' -------------------> No exception is thrown here
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.Write("    2) ")
                result = Nothing
                xint321 = -536870912
                result = xint321 - (-xint321)
                result = xint321 - (-result)
                result = xint321 - (-result)
                result = xint321 - (-result)   ' ----------------------> This throws an overflow exception
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub

    End Class


    Class Bug287339
        Shared Sub Test()
            Dim o1 As Object = 54
            Dim o2 As Object = "hello"

            Console.Write("*** Bug 287339: ")
            Try
                Console.WriteLine(o1 <> o2)		'the error message given by this illegal comparison is wrong
            Catch e As System.Exception
                Console.WriteLine(e.Message)
            End Try
        End Sub
    End Class


    Class Bug288405
        Shared Sub Test()
            Console.WriteLine("*** Bug 288405")

            Dim Ary() As Char = "10"
            Dim s As String = "Hello"
            Dim o1 As Object = 10
            Dim o2 As Object = Ary

            Console.Write("    1) ")
            Try
                Console.WriteLine(o1 = o2)
                Failed()
            Catch ex As InvalidCastException
                Passed()
                Console.WriteLine(ex.Message)
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("    2) ")
            Try
                o1 = Ary
                o1 += 10  'No exception given here,
                Failed()
            Catch ex As InvalidCastException
                Passed()
                Console.WriteLine(ex.Message)
            Catch ex As Exception
                Failed(ex)
            End Try

        End Sub

    End Class


    Class MyCulture
        Inherits CultureInfo

        Private m_nfi As NumberFormatInfo
        Private m_dtfi As DateTimeFormatInfo
        Shared m_culture As MyCulture

        Shared Sub Init(ByVal lcid As Integer)
            'Set culture to Enlish to avoid locale diffs
            Dim ci As CultureInfo = New CultureInfo(lcid)
            Dim nfi As NumberFormatInfo = CType(ci.GetFormat(GetType(NumberFormatInfo)), NumberFormatInfo)
            Dim dtfi As DateTimeFormatInfo = CType(ci.GetFormat(GetType(DateTimeFormatInfo)), DateTimeFormatInfo)

            System.Threading.Thread.CurrentThread.CurrentCulture = New MyCulture(lcid, dtfi, nfi)
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

            Return Nothing
        End Function

    End Class


End Namespace



