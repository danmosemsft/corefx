Imports System
Imports Microsoft.VisualBasic

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
            Console.WriteLine("Begin Tests")
            TestPublicFuncs()
            RegressionTests()
            Console.WriteLine("End Tests")
        End Sub

        Sub RegressionTests()
            Console.WriteLine("Regression tests")
            Bug229397.Test
            Bug249566.Test
        End Sub

        Sub TestPublicFuncs()
            CBoolTests()
            CByteTests()
            CDecTests()
            CDblTests()
            CDateTest()
            CIntTest
            CLngTest
            CSngTest()
            CStrTest()
            CObjTest
            ErrorTest
            ErrorStrTest
            FixTest
            IntTest
            HexTest
            OctTest
            StrTest
            ValTest
            ChrTest
            ChrWTest
            AscTest
            AscWTest
            AscAndChrConstTest
        End Sub



        Sub CObjTest()
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



        Sub CBoolTests()
            Console.WriteLine("*** CBool tests")
            Try

                i = 1 : Console.Write("1a) ") : PassFail(CBool(i) = True)
                i = -1 : Console.Write("1b) ") : PassFail(CBool(i) = True)
                i = 0 : Console.Write("1c) ") : PassFail(CBool(i) = False)
                i = &H8000I : Console.Write("1d) ") : PassFail(CBool(i) = True)

                l = 1 : Console.Write("2a) ") : PassFail(CBool(l) = True)
                l = -1 : Console.Write("2b) ") : PassFail(CBool(l) = True)
                l = 0 : Console.Write("2c) ") : PassFail(CBool(l) = False)
                l = &H80000000L
                 : Console.Write("2d) ") : PassFail(CBool(l) = True)
                l = 1 : Console.Write("2e) ") : PassFail(CBool(l) = True)

                sng = 1 : Console.Write("3a) ") : PassFail(CBool(sng) = True)
                sng = -1 : Console.Write("3b) ") : PassFail(CBool(sng) = True)
                sng = 0 : Console.Write("3c) ") : PassFail(CBool(sng) = False)
                sng = CSng(1E+16) : Console.Write("3d) ") : PassFail(CBool(sng) = True)

                dbl = 1 : Console.Write("4a) ") : PassFail(CBool(dbl) = True)
                dbl = -1 : Console.Write("4b) ") : PassFail(CBool(dbl) = True)
                dbl = 0 : Console.Write("4c) ") : PassFail(CBool(dbl) = False)
                dbl = 1E+32 : Console.Write("4d) ") : PassFail(CBool(dbl) = True)

                obj = 1 : Console.Write("5) ") : PassFail(CBool(obj) = True)

                b = True : Console.Write("6a) ") : PassFail(CBool(b) = True)
                b = False : Console.Write("6b) ") : PassFail(CBool(b) = False)

                byt = 1 : Console.Write("7a) ") : PassFail(CBool(byt) = True)
                byt = 0 : Console.Write("7b) ") : PassFail(CBool(byt) = False)
                byt = 255 : Console.Write("7c) ") : PassFail(CBool(byt) = True)
                byt = 128 : Console.Write("7d) ") : PassFail(CBool(byt) = True)
                byt = 127 : Console.Write("7e) ") : PassFail(CBool(byt) = True)

                dec = 1 : Console.Write("8a) ") : PassFail(CBool(dec) = True)
                dec = -1 : Console.Write("8b) ") : PassFail(CBool(dec) = True)
                dec = 0 : Console.Write("8c) ") : PassFail(CBool(dec) = False)
                dec = 1234 : Console.Write("8d) ") : PassFail(CBool(dec) = True)

                l = 1 : Console.Write("9a) ") : PassFail(CBool(l) = True)
                l = -1 : Console.Write("9b) ") : PassFail(CBool(l) = True)
                l = 0 : Console.Write("9c) ") : PassFail(CBool(l) = False)
                l = &H7FFFFFFFFFFFFFFFL
                 : Console.Write("9d) ") : PassFail(CBool(l) = True)
                l = 1 : Console.Write("9e) ") : PassFail(CBool(l) = True)

                s = "True" : Console.Write("10a) ") : PassFail(CBool(s) = True)
                s = "False" : Console.Write("10b) ") : PassFail(CBool(s) = False)
                s = "&H123" : Console.Write("10c) ") : PassFail(CBool(s) = True)
                s = " &H123" : Console.Write("10d) ") : PassFail(CBool(s) = True)
                s = "&O123" : Console.Write("10e) ") : PassFail(CBool(s) = True)
                s = " &O123" : Console.Write("10f) ") : PassFail(CBool(s) = True)

            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub



        Sub CByteTests()
            Console.WriteLine("*** CByte tests")

            Try
                sh = 1 : Console.Write("1a) ") : PassFail(CByte(sh) = 1)
                sh = 0 : Console.Write("1b) ") : PassFail(CByte(sh) = 0)

                i = &HA1 : Console.Write("2a) ") : PassFail(CByte(i) = &HA1)
                i = 0 : Console.Write("2b) ") : PassFail(CByte(i) = 0)

                l = &HAB : Console.Write("3a) ") : PassFail(CByte(l) = &HAB)
                l = &H0 : Console.Write("3b) ") : PassFail(CByte(l) = 0)

                sng = 1 : Console.Write("4a) ") : PassFail(CByte(sng) = 1)
                sng = 0 : Console.Write("4b) ") : PassFail(CByte(sng) = 0)

                dbl = 1 : Console.Write("5a) ") : PassFail(CByte(dbl) = 1)
                dbl = 0 : Console.Write("5b) ") : PassFail(CByte(dbl) = 0)

                'Date to Byte not supported

                obj = 0 : Console.Write("6a) ") : PassFail(CByte(obj) = 0)
                obj = 127 : Console.Write("6b) ") : PassFail(CByte(obj) = 127)
                obj = 128 : Console.Write("6c) ") : PassFail(CByte(obj) = 128)
                obj = 255 : Console.Write("6d) ") : PassFail(CByte(obj) = 255)

                b = True : Console.Write("7a) ") : PassFail(CByte(b) = 255)
                b = False : Console.Write("7b) ") : PassFail(CByte(b) = 0)

                byt = 1 : Console.Write("8a) ") : PassFail(CByte(byt) = 1)
                byt = 0 : Console.Write("8b) ") : PassFail(CByte(byt) = 0)
                byt = 255 : Console.Write("8c) ") : PassFail(CByte(byt) = 255)
                byt = 128 : Console.Write("8d) ") : PassFail(CByte(byt) = 128)
                byt = 127 : Console.Write("8e) ") : PassFail(CByte(byt) = 127)

                dec = &H12 : Console.Write("9a) ") : PassFail(CByte(dec) = &H12)
                dec = 0 : Console.Write("9b) ") : PassFail(CByte(dec) = 0)

                s = "1" : Console.Write("10a) ") : PassFail(CByte(s) = 1)
                s = "0" : Console.Write("10b) ") : PassFail(CByte(s) = 0)

            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub

    

        Sub CDblTests()
            Console.WriteLine("*** CDbl tests")

            Try
                sh = 1 : Console.Write("1a) ") : PassFail(CDbl(sh) = 1)
                sh = -1 : Console.Write("1b) ") : PassFail(CDbl(sh) = -1)
                sh = 0 : Console.Write("1c) ") : PassFail(CDbl(sh) = 0)
                sh = &H8000S : Console.Write("1d) ") : PassFail(CDbl(sh) = &H8000S) : sh = 0

                i = 1 : Console.Write("2a) ") : PassFail(CDbl(i) = 1)
                i = -1 : Console.Write("2b) ") : PassFail(CDbl(i) = -1)
                i = 0 : Console.Write("2c) ") : PassFail(CDbl(i) = 0)
                i = &H80000000I
                 : Console.Write("2d) ") : PassFail(CDbl(i) = &H80000000I)

                l = 1 : Console.Write("3a) ") : PassFail(CDbl(l) = 1)
                l = -1 : Console.Write("3b) ") : PassFail(CDbl(l) = -1)
                l = 0 : Console.Write("3c) ") : PassFail(CDbl(l) = 0)
                'UNDONE
                'l = &H7FFFFFFFFFFFFFFFL
                '            : Console.Write( "3d) " ) : PassFail( CDbl( l ) = &H7FFFFFFFFFFFFFFFL )
                'l = &H8000000000000000L
                '            : Console.Write( "3d) " ) : PassFail( CDbl( l ) = &H8000000000000000L )

                sng = 1 : Console.Write("4a) ") : PassFail(CDbl(sng) = 1)
                sng = -1 : Console.Write("4b) ") : PassFail(CDbl(sng) = -1)
                sng = 0 : Console.Write("4c) ") : PassFail(CDbl(sng) = 0)
                'UNDONE     sng = CSng(1E+16) : Console.Write( "4d) " ) : PassFail( CDbl( sng ) = 1E+16 )

                dbl = 1 : Console.Write("5a) ") : PassFail(CDbl(dbl) = 1)
                dbl = -1 : Console.Write("5b) ") : PassFail(CDbl(dbl) = -1)
                dbl = 0 : Console.Write("5c) ") : PassFail(CDbl(dbl) = 0)
                dbl = 1E+32 : Console.Write("5d) ") : PassFail(CDbl(dbl) = 1E+32)

                obj = 1 : Console.Write("6) ") : PassFail(CDbl(obj) = 1)

                b = True : Console.Write("7a) ") : PassFail(CDbl(b) = -1)
                b = False : Console.Write("7b) ") : PassFail(CDbl(b) = 0)

                byt = 1 : Console.Write("8a) ") : PassFail(CDbl(byt) = 1)
                byt = 0 : Console.Write("8b) ") : PassFail(CDbl(byt) = 0)
                byt = 255 : Console.Write("8c) ") : PassFail(CDbl(byt) = 255)
                byt = 128 : Console.Write("8d) ") : PassFail(CDbl(byt) = 128)
                byt = 127 : Console.Write("8e) ") : PassFail(CDbl(byt) = 127)

                dec = 1 : Console.Write("9a) ") : PassFail(CDbl(dec) = 1)
                dec = -1 : Console.Write("9b) ") : PassFail(CDbl(dec) = -1)
                dec = 0 : Console.Write("9c) ") : PassFail(CDbl(dec) = 0)
                dec = 1234 : Console.Write("9d) ") : PassFail(CDbl(dec) = 1234)

                s = "1" : Console.Write("10a) ") : PassFail(CDbl(s) = 1)
                s = "-1" : Console.Write("10b) ") : PassFail(CDbl(s) = -1)
                s = "0" : Console.Write("10c) ") : PassFail(CDbl(s) = 0)
                s = "1234" : Console.Write("10d) ") : PassFail(CDbl(s) = 1234)
                s = "&H123" : Console.Write("10e) ") : PassFail(CDbl(s) = &H123)
                s = " &H123" : Console.Write("10f) ") : PassFail(CDbl(s) = &H123)
                s = "&O123" : Console.Write("10g) ") : PassFail(CDbl(s) = &O123)
                s = " &O123" : Console.Write("10h) ") : PassFail(CDbl(s) = &O123)
                s = "1E10" : Console.Write("10i) ") : PassFail(CDbl(s) = 10000000000.0)
                s = "1E200" : Console.Write("10j) ") : PassFail(CDbl(s) = 1.0E+200)
                Try
                    s = "1E20000" : Console.Write("10k) ")
                    PassFail(CDbl(s) = 42)
                Catch ex As OverflowException
                    PassFail(True)
                End Try
                s = "Infinity" : Console.Write("10l) ") : PassFail(CDbl(s) = System.Double.PositiveInfinity)
                s = "NaN" : Console.Write("10m) ") : PassFail(System.Double.IsNaN(CDbl(s)))

            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub



        Sub CDecTests()
            Console.WriteLine("*** CDec tests")

            Try

                i = 1 : Console.Write("1a) ") : PassFail(CDec(i) = 1)
                i = -1 : Console.Write("1b) ") : PassFail(CDec(i) = -1)
                i = 0 : Console.Write("1c) ") : PassFail(CDec(i) = 0)
                i = &H8000 : Console.Write("1d) ") : PassFail(CDec(i) = &H8000)

                l = 1 : Console.Write("2a) ") : PassFail(CDec(l) = 1)
                l = -1 : Console.Write("2b) ") : PassFail(CDec(l) = -1)
                l = 0 : Console.Write("2c) ") : PassFail(CDec(l) = 0)
                l = &H80000000L : Console.Write("2d) ") : PassFail(CDec(l) = &H80000000L)
                l = &H7FFFFFFFFFFFFFFFL : Console.Write("2e) ") : PassFail(CDec(l) = &H7FFFFFFFFFFFFFFFL)

                sng = 1 : Console.Write("3a) ") : PassFail(CDec(sng) = 1)
                sng = -1 : Console.Write("3b) ") : PassFail(CDec(sng) = -1)
                sng = 0 : Console.Write("3c) ") : PassFail(CDec(sng) = 0)
                sng = CSng(1E+16) : Console.Write("3d) ") : PassFail(CDec(sng) = 1E+16)

                dbl = 1 : Console.Write("4a) ") : PassFail(CDec(dbl) = 1)
                dbl = -1 : Console.Write("4b) ") : PassFail(CDec(dbl) = -1)
                dbl = 0 : Console.Write("4c) ") : PassFail(CDec(dbl) = 0)
                'UNDONE:    dbl = 1E+32 : Console.Write( "4d) " ) : PassFail( CDec( dbl ) = 1E+32 )

                obj = 1 : Console.Write("5) ") : PassFail(CDec(obj) = 1)

                b = True : Console.Write("6a) ") : PassFail(CDec(b) = -1)
                b = False : Console.Write("6b) ") : PassFail(CDec(b) = 0)

                byt = 1 : Console.Write("7a) ") : PassFail(CDec(byt) = 1)
                byt = 0 : Console.Write("7b) ") : PassFail(CDec(byt) = 0)
                byt = 255 : Console.Write("7c) ") : PassFail(CDec(byt) = 255)
                byt = 128 : Console.Write("7d) ") : PassFail(CDec(byt) = 128)
                byt = 127 : Console.Write("7e) ") : PassFail(CDec(byt) = 127)

                dec = 1 : Console.Write("8a) ") : PassFail(CDec(dec) = 1)
                dec = -1 : Console.Write("8b) ") : PassFail(CDec(dec) = -1)
                dec = 0 : Console.Write("8c) ") : PassFail(CDec(dec) = 0)
                dec = 1234 : Console.Write("8d) ") : PassFail(CDec(dec) = 1234)

                s = "1" : Console.Write("9a) ") : PassFail(CDec(s) = 1)
                s = "-1" : Console.Write("9b) ") : PassFail(CDec(s) = -1)
                s = "0" : Console.Write("9c) ") : PassFail(CDec(s) = 0)
                s = "1234" : Console.Write("9d) ") : PassFail(CDec(s) = 1234)
                s = "&H123" : Console.Write("9e) ") : PassFail(CDec(s) = &H123)
                s = " &H123" : Console.Write("9f) ") : PassFail(CDec(s) = &H123)
                s = "&O123" : Console.Write("9g) ") : PassFail(CDec(s) = &O123)
                s = " &O123" : Console.Write("9h) ") : PassFail(CDec(s) = &O123)

            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub



        Sub Failed(ByVal ex As Exception)
            If ex Is Nothing Then
                Console.WriteLine("NULL System.Exception")
            Else
                Console.WriteLine(ex.GetType().FullName)
            End If
            Console.WriteLine("FAILED !!!")
        End Sub



        Sub Failed()
            Console.WriteLine("FAILED !!!")
        End Sub



        Sub Passed()
            Console.WriteLine("passed")
        End Sub



        Sub PassFail(ByVal bPassed As Boolean)
            If bPassed Then
                Console.WriteLine("passed")
            Else
                Console.WriteLine("FAILED !!!")
            End If
            m_lBugId = 0
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
            ' NOTE:  ***** DO NOT MODIFY THIS FUNCTION *****
            '       (unless you really know what you're doing)
            '
            ' This function is intended to just to ensure compilation
            ' If someone changes a name in the runtime, this should break
            ' we will rely on the test team to make sure named params
            ' return results as expected
            ' 
            'obj = VBA.CBool(Expression := obj)
            'obj = VBA.CByte(Expression := obj)
            'obj = VBA.CCur(Expression := obj)
            'obj = VBA.CDate(Expression := obj)
            'obj = VBA.CDbl(Expression := obj)
            'obj = VBA.CDec(Expression := obj)
            'obj = VBA.CInt(Expression := obj)
            'obj = VBA.CLng(Expression := obj)
            'obj = VBA.CSng(Expression := obj)
            'obj = VBA.CStr(Expression := obj)
            'obj = VBA.CVar(Expression := obj)
            'obj = VBA.CVDate(Expression := obj)
            'obj = VBA.CVErr(Expression := obj)
#if 0
'UNDONE: BUG BUG BUG
            b = ErrorToString(ErrorNumber := i)
            b = Fix(Number := sh)
            b = Fix(Number := i)
            b = Fix(Number := l)
            b = Fix(Number := b)
            b = Int(Number := sh)
            b = Int(Number := i)
            b = Int(Number := l)
            b = Int(Number := b)
            b = Hex(Number := obj)
            b = Oct(Number := obj)
            b = Str(Number := obj)
            b = Val(String := "1234")
#end if
        End Sub



        Sub CDateTest()
            Console.WriteLine("*** CDate tests")

            Dim s As String

            Try

                s = "1/31/99"
                Console.Write("1a) ")
                PassFail(CDate(s) = #1/31/1999#)

                s = "12/31/2099"
                Console.Write("1b) ")
                PassFail(CDate(s) = #12/31/2099#)

                s = "PM 10:10:10"
                Console.Write("1c) ")
	            PassFail(CDate(s) = #22:10:10#)
#if 0
'UNDONE: Enable after 244341 fixed
                s = "AM 10:10:10"
                Console.Write("1d) ")
	            PassFail(CDate(s) = #10:10:10#)
#End If
            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub



        Sub CIntTest()
            On Error Resume Next
            Console.WriteLine("*** CInt tests")
            i = 1 : Console.WriteLine("1a) " & CInt(i))
            i = -1 : Console.WriteLine("1b) " & CInt(i))
            i = 0 : Console.WriteLine("1c) " & CInt(i))
            i = &H8000 : Console.WriteLine("1d) " & CInt(i))

            l = 1 : Console.WriteLine("2a) " & CInt(l))
            l = -1 : Console.WriteLine("2b) " & CInt(l))
            l = 0 : Console.WriteLine("2c) " & CInt(l))
            l = &H80000000&
             : Console.WriteLine("2d) " & CInt(l))
            l = 1 : Console.WriteLine("2e) " & CInt(l))

            sng = 1 : Console.WriteLine("3a) " & CInt(sng))
            sng = -1 : Console.WriteLine("3b) " & CInt(sng))
            sng = 0 : Console.WriteLine("3c) " & CInt(sng))
            Err.Number = 0
            sng = 1E+16 : Console.Write("3d) ") : i = CInt(sng)
            ErrorCheck(6)

            dbl = 1 : Console.WriteLine("4a) " & CInt(dbl))
            dbl = -1 : Console.WriteLine("4b) " & CInt(dbl))
            dbl = 0 : Console.WriteLine("4c) " & CInt(dbl))
            Err.Number = 0
            dbl = 1E+32 : Console.Write("4d) ") : i= CInt(dbl)
            ErrorCheck(6)

            s = "1" : Console.WriteLine("7a) " & CInt(s))
            s = "-1" : Console.WriteLine("7b) " & CInt(s))
            s = "0" : Console.WriteLine("7c) " & CInt(s))
            s = "1234" : Console.WriteLine("7d) " & CInt(s))
            s = "&H123" : Console.Write("7e) ") : PassFail(CInt(s) = &H123)
            s = " &H123" : Console.Write("7f) ") : PassFail(CInt(s) = &H123)
            s = "&O123" : Console.Write("7g) ") : PassFail(CInt(s) = &O123)
            s = " &O123" : Console.Write("7h) ") : PassFail(CInt(s) = &O123)

            b = True : Console.WriteLine("9a) " & CInt(b))
            b = False : Console.WriteLine("9b) " & CInt(b))

            obj = 1 : Console.WriteLine("10) " & CInt(obj))

            byt = 1 : Console.WriteLine("11a) " & CInt(byt))
            byt = 0 : Console.WriteLine("11b) " & CInt(byt))
            byt = 255 : Console.WriteLine("11c) " & CInt(byt))
            byt = 128 : Console.WriteLine("11d) " & CInt(byt))
            byt = 127 : Console.WriteLine("11e) " & CInt(byt))

            l = 1 : dec = l : Console.WriteLine("13a) " & CInt(dec))
            l = -1 : dec = l : Console.WriteLine("13b) " & CInt(dec))
            l = 0 : dec = l : Console.WriteLine("13c) " & CInt(dec))
            l = 1234 : dec = l : Console.WriteLine("13d) " & CInt(dec))

            c = "a"c : Console.WriteLine("14a) " & Asc(c))
            c = "b"c : Console.WriteLine("14b) " & Asc(c))
            c = Chr(0) : Console.WriteLine("14c) " & Asc(c))
            c = Chr(&HFFFF) : Console.WriteLine("14d) " & Asc(c))
            c = Chr(&H7FFF) : Console.WriteLine("14e) " & Asc(c))
        End Sub



        Sub CLngTest()
            On Error Resume Next

            Console.WriteLine("*** CLng tests")
            i = 1 : Console.WriteLine("1a) " & CLng(i))
            i = -1 : Console.WriteLine("1b) " & CLng(i))
            i = 0 : Console.WriteLine("1c) " & CLng(i))
            i = &H8000 : Console.WriteLine("1d) " & CLng(i))

            l = 1 : Console.WriteLine("2a) " & CLng(l))
            l = -1 : Console.WriteLine("2b) " & CLng(l))
            l = 0 : Console.WriteLine("2c) " & CLng(l))
            l = &H80000000&
             : Console.WriteLine("2d) " & CLng(l))
            l = 1 : Console.WriteLine("2e) " & CLng(l))

            sng = 1 : Console.WriteLine("3a) " & CLng(sng))
            sng = -1 : Console.WriteLine("3b) " & CLng(sng))
            sng = 0 : Console.WriteLine("3c) " & CLng(sng))
            sng = 2147483583 : Console.WriteLine("3d) " & CLng(sng))
            Err.Number = 0
            sng = 1E+38 : Console.Write("3e) ") : Console.WriteLine(CLng(sng))
            ErrorCheck(6)

            dbl = 1 : Console.WriteLine("4a) " & CLng(dbl))
            dbl = -1 : Console.WriteLine("4b) " & CLng(dbl))
            dbl = 0 : Console.WriteLine("4c) " & CLng(dbl))
            Err.Number = 0
            dbl = 1E+32 : Console.Write("4d) ") : l = CLng(dbl)
            ErrorCheck(6)

            s = "1" : Console.WriteLine("7a) " & CLng(s))
            s = "-1" : Console.WriteLine("7b) " & CLng(s))
            s = "0" : Console.WriteLine("7c) " & CLng(s))
            s = "1234" : Console.WriteLine("7d) " & CLng(s))
            s = "&H123" : Console.Write("7e) ") : PassFail(CLng(s) = &H123)
            s = " &H123" : Console.Write("7f) ") : PassFail(CLng(s) = &H123)
            s = "&O123" : Console.Write("7g) ") : PassFail(CLng(s) = &O123)
            s = " &O123" : Console.Write("7h) ") : PassFail(CLng(s) = &O123)

            b = True : Console.WriteLine("9a) " & CLng(b))
            b = False : Console.WriteLine("9b) " & CLng(b))

            obj = 1 : Console.WriteLine("10) " & CLng(obj))

            byt = 1 : Console.WriteLine("11a) " & CLng(byt))
            byt = 0 : Console.WriteLine("11b) " & CLng(byt))
            byt = 255 : Console.WriteLine("11c) " & CLng(byt))
            byt = 128 : Console.WriteLine("11d) " & CLng(byt))
            byt = 127 : Console.WriteLine("11e) " & CLng(byt))

            'udt = 1    : Console.WriteLine( "12) " & CLng(udt))

            l = 1 : dec = l : Console.WriteLine("13a) " & CLng(dec))
            l = -1 : dec = l : Console.WriteLine("13b) " & CLng(dec))
            l = 0 : dec = l : Console.WriteLine("13c) " & CLng(dec))
            l = 1234 : dec = l : Console.WriteLine("13d) " & CLng(dec))
        End Sub



        Sub CSngTest()
            Console.WriteLine("*** CSng tests")

            Try
                sh = 1 : Console.Write("1a) ") : PassFail(CSng(sh) = 1)
                sh = -1 : Console.Write("1b) ") : PassFail(CSng(sh) = -1)
                sh = 0 : Console.Write("1c) ") : PassFail(CSng(sh) = 0)
                sh = &H8000S : Console.Write("1d) ") : PassFail(CSng(sh) = &H8000S)

                i = 1 : Console.Write("2a) ") : PassFail(CSng(i) = 1)
                i = -1 : Console.Write("2b) ") : PassFail(CSng(i) = -1)
                i = 0 : Console.Write("2c) ") : PassFail(CSng(i) = 0)
                i = &H80000000I : Console.Write("2d) ") : PassFail(CSng(i) = &H80000000I)

                l = 1 : Console.Write("3a) ") : PassFail(CSng(l) = 1)
                l = -1 : Console.Write("3b) ") : PassFail(CSng(l) = -1)
                l = 0 : Console.Write("3c) ") : PassFail(CSng(l) = 0)
                l = &H8000000000000000L
                 : Console.Write("3d) ") : PassFail(CSng(l) = &H8000000000000000L)

                sng = 1 : Console.Write("4a) ") : PassFail(CSng(sng) = 1)
                sng = -1 : Console.Write("4b) ") : PassFail(CSng(sng) = -1)
                sng = 0 : Console.Write("4c) ") : PassFail(CSng(sng) = 0)
                sng = 1E+16 : Console.Write("4d) ") : PassFail(CSng(sng) = CSng(1E+16))

                dbl = 1 : Console.Write("5a) ") : PassFail(CSng(dbl) = 1)
                dbl = -1 : Console.Write("5b) ") : PassFail(CSng(dbl) = -1)
                dbl = 0 : Console.Write("5c) ") : PassFail(CSng(dbl) = 0)
                dbl = 1E+32 : Console.Write("5d) ") : PassFail(CSng(dbl) = CSng(1E+32))

                s = "1" : Console.Write("6a) ") : PassFail(CSng(s) = 1)
                s = "-1" : Console.Write("6b) ") : PassFail(CSng(s) = -1)
                s = "0" : Console.Write("6c) ") : PassFail(CSng(s) = 0)
                s = "12.34" : Console.Write("6d) ") : PassFail(CSng(s) = CSng(12.34))
                s = "&H123" : Console.Write("7e) ") : PassFail(CSng(s) = &H123)
                s = " &H123" : Console.Write("7f) ") : PassFail(CSng(s) = &H123)
                s = "&O123" : Console.Write("7g) ") : PassFail(CSng(s) = &O123)
                s = " &O123" : Console.Write("7h) ") : PassFail(CSng(s) = &O123)
                s = "1E10" : Console.Write("7i) ") : PassFail(CSng(s) = 10000000000.0)
                Try
                    s = "1E200" : Console.Write("7j) ")
                    PassFail(CSng(s) = 42)
                Catch ex As OverflowException
                    PassFail(True)
                End Try
                Try
                    s = "1E20000" : Console.Write("7k) ")
                    PassFail(CSng(s) = 42)
                Catch ex As OverflowException
                    PassFail(True)
                End Try
                s = "Infinity" : Console.Write("7l) ") : PassFail(CSng(s) = System.Single.PositiveInfinity)
                s = "NaN" : Console.Write("7m) ") : PassFail(System.Single.IsNaN(CSng(s)))

                obj = 1.234 : Console.Write("7) ") : PassFail(CSng(obj) = CSng(1.234))

                b = True : Console.Write("8a) ") : PassFail(CSng(b) = -1)
                b = False : Console.Write("8b) ") : PassFail(CSng(b) = 0)

                obj = 1 : Console.Write("9) ") : PassFail(CSng(obj) = 1)

                byt = 1 : Console.Write("10a) ") : PassFail(CSng(byt) = 1)
                byt = 0 : Console.Write("10b) ") : PassFail(CSng(byt) = 0)
                byt = 255 : Console.Write("10c) ") : PassFail(CSng(byt) = 255)
                byt = 128 : Console.Write("10d) ") : PassFail(CSng(byt) = 128)
                byt = 127 : Console.Write("10e) ") : PassFail(CSng(byt) = 127)

                dec = 1 : Console.Write("11a) ") : PassFail(CSng(dec) = 1)
                dec = -1 : Console.Write("11b) ") : PassFail(CSng(dec) = -1)
                dec = 0 : Console.Write("11c) ") : PassFail(CSng(dec) = 0)
                dec = 1234 : Console.Write("11d) ") : PassFail(CSng(dec) = 1234)

            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub


        Sub CStrTest()
            Console.WriteLine("*** CStr tests")
            i = 1 : Console.WriteLine("1a) " & CStr(i))
            i = -1 : Console.WriteLine("1b) " & CStr(i))
            i = 0 : Console.WriteLine("1c) " & CStr(i))
            i = &H8000 : Console.WriteLine("1d) " & CStr(i))

            l = 1 : Console.WriteLine("2a) " & CStr(l))
            l = -1 : Console.WriteLine("2b) " & CStr(l))
            l = 0 : Console.WriteLine("2c) " & CStr(l))
            l = &H80000000L : Console.WriteLine("2d) " & CStr(l))
            l = 1 : Console.WriteLine("2e) " & CStr(l))

            sng = 1 : Console.WriteLine("3a) " & CStr(sng))
            sng = -1 : Console.WriteLine("3b) " & CStr(sng))
            sng = 0 : Console.WriteLine("3c) " & CStr(sng))
            sng = 1.0E+16 : Console.WriteLine("3d) " & CStr(sng))

            dbl = 1 : Console.WriteLine("4a) " & CStr(dbl))
            dbl = -1 : Console.WriteLine("4b) " & CStr(dbl))
            dbl = 0 : Console.WriteLine("4c) " & CStr(dbl))
            dbl = 1.0E+32 : Console.WriteLine("4d) " & CStr(dbl))

            Dim dt As Date
            Dim s As String

            Console.WriteLine("6a) " & CStr(dt))
            dt = #1/1/1999# : s = CStr(dt)
            Console.Write("6b) ") : PassFail(CDate(s) = #1/1/1999#)

            s = "1" : Console.WriteLine("7a) " & CStr(s))
            s = "-1" : Console.WriteLine("7b) " & CStr(s))
            s = "0" : Console.WriteLine("7c) " & CStr(s))
            s = "1234" : Console.WriteLine("7d) " & CStr(s))

            b = True : Console.WriteLine("9a) " & CStr(b))
            b = False : Console.WriteLine("9b) " & CStr(b))

            obj = 1 : Console.WriteLine("10) " & CStr(obj))

            byt = 1 : Console.WriteLine("11a) " & CStr(byt))
            byt = 0 : Console.WriteLine("11b) " & CStr(byt))
            byt = 255 : Console.WriteLine("11c) " & CStr(byt))
            byt = 128 : Console.WriteLine("11d) " & CStr(byt))
            byt = 127 : Console.WriteLine("11e) " & CStr(byt))

            'udt = 1    : Console.WriteLine "12) " &  CStr(udt)

            l = 1 : dec = l : Console.WriteLine("13a) " & CStr(dec))
            l = -1 : dec = l : Console.WriteLine("13b) " & CStr(dec))
            l = 0 : dec = l : Console.WriteLine("13c) " & CStr(dec))
            l = 1234 : dec = l : Console.WriteLine("13d) " & CStr(dec))
        End Sub



        Sub CStrDblTest(ByVal Expected As String, ByVal dbl As Double)
            Dim s As String = CStr(dbl)
            If s <> Expected Then
                Console.WriteLine("Failed: Expected '" & Expected & "' returned '" & s & "'")
            Else
                Console.WriteLine("Passed")
            End If
        End Sub



        Sub CStrDoubleTest()
            Console.WriteLine("* Test of CStr(Double)")
            CStrDblTest("0.1", 0.1)
            CStrDblTest("0.01", 0.01)
            CStrDblTest("0.001", 0.001)
            CStrDblTest("0.0001", 0.0001)
            CStrDblTest("0.00001", 0.00001)
            CStrDblTest("0.000001", 0.000001)
            CStrDblTest("0.0000001", 0.0000001)
            CStrDblTest("0.00000001", 0.00000001)
            CStrDblTest("0.000000001", 0.000000001)
            CStrDblTest("0.0000000001", 0.0000000001)
            CStrDblTest("0.00000000001", 0.00000000001)
            CStrDblTest("0.000000000001", 0.000000000001)
            CStrDblTest("0.0000000000001", 0.0000000000001)
            CStrDblTest("0.00000000000001", 0.00000000000001)
            CStrDblTest("0.000000000000001", 0.000000000000001)
            CStrDblTest("1E-16", 0.0000000000000001) '15

            CStrDblTest("999999999999999", 999999999999999)
            CStrDblTest("999999999999999", 999999999999999.0# + 0.4#)
            CStrDblTest("1E+15", 999999999999999.0# + 0.5#)
            CStrDblTest("0.00000000000001", 0.00000000000001)
            CStrDblTest("0.000000000000001", 0.000000000000001)
            CStrDblTest("1E-16", 0.0000000000000001)
            CStrDblTest("100000000000000", 100000000000000.0)
            CStrDblTest("1E+15", 1.0E+15)
            CStrDblTest("1E+16", 1.0E+16)
        End Sub



        Sub CStrSngTest(ByVal Expected As String, ByVal Sng As Single)
            Dim s As String = CStr(Sng)
            If s <> Expected Then
                Console.WriteLine("Failed: Expected '" & Expected & "' returned '" & s & "'")
            Else
                Console.WriteLine("Passed")
            End If
        End Sub



        Sub CStrSingleTest()
            Console.WriteLine("* Test of CStr(Single)")
            CStrSngTest("0.1", 0.1)
            CStrSngTest("0.01", 0.01)
            CStrSngTest("0.001", 0.001)
            CStrSngTest("0.0001", 0.0001)
            CStrSngTest("0.00001", 0.00001)
            CStrSngTest("0.000001", 0.000001)
            CStrSngTest("0.0000001", 0.0000001)
            CStrSngTest("0.00000001", 0.00000001)
            CStrSngTest("0.000000001", 0.000000001)
            CStrSngTest("0.0000000001", 0.0000000001)
            CStrSngTest("0.00000000001", 0.00000000001)
            CStrSngTest("0.000000000001", 0.000000000001)
            CStrSngTest("0.0000000000001", 0.0000000000001)
            CStrSngTest("0.00000000000001", 0.00000000000001)
            CStrSngTest("0.000000000000001", 0.000000000000001)
            CStrSngTest("1E-16", 0.0000000000000001) '15

            CStrSngTest("999999999999999", 999999999999999)
            CStrSngTest("999999999999999", 999999999999999.0# + 0.4#)
            CStrSngTest("1E+15", 999999999999999.0# + 0.5#)
            CStrSngTest("0.00000000000001", 0.00000000000001)
            CStrSngTest("0.000000000000001", 0.000000000000001)
            CStrSngTest("1E-16", 0.0000000000000001)
            CStrSngTest("100000000000000", 100000000000000.0)
            CStrSngTest("1E+15", 1.0E+15)
            CStrSngTest("1E+16", 1.0E+16)
        End Sub



        Sub IntTest()
            On Error Resume Next
            Console.WriteLine("Int tests")
            Console.WriteLine("Testing helper IntR4")

            Dim sng As Single
            sng = 5.5 : Console.Write("1) ") : PassFail(Int(sng) = 5)
            sng = 6.5 : Console.Write("2) ") : PassFail(Int(sng) = 6)
            sng = 0 : Console.Write("3) ") : PassFail(Int(sng) = 0)
            sng = -0.4 : Console.Write("4) ") : PassFail(Int(sng) = -1)
            sng = -0.5 : Console.Write("5) ") : PassFail(Int(sng) = -1)
            sng = -1 : Console.Write("6) ") : PassFail(Int(sng) = -1)
            sng = -1.1 : Console.Write("7) ") : PassFail(Int(sng) = -2)
            sng = -1.5 : Console.Write("8) ") : PassFail(Int(sng) = -2)
            sng = -2.5 : Console.Write("9) ") : PassFail(Int(sng) = -3)

            Console.WriteLine("Testing helper IntR8")
            Dim dbl As Double
            dbl = 5.5 : Console.Write("1) ") : PassFail(Int(dbl) = 5.0#)
            dbl = 6.5 : Console.Write("2) ") : PassFail(Int(dbl) = 6.0#)
            dbl = 0 : Console.Write("3) ") : PassFail(Int(dbl) = 0.0#)
            dbl = -0.4 : Console.Write("4) ") : PassFail(Int(dbl) = -1.0#)
            dbl = -0.5 : Console.Write("5) ") : PassFail(Int(dbl) = -1.0#)
            dbl = -1 : Console.Write("6) ") : PassFail(Int(dbl) = -1.0#)
            dbl = -1.1 : Console.Write("7) ") : PassFail(Int(dbl) = -2.0#)
            dbl = -1.5 : Console.Write("8) ") : PassFail(Int(dbl) = -2.0#)
            dbl = -2.5 : Console.Write("9) ") : PassFail(Int(dbl) = -3.0#)
            dbl = -12345678.12 : Console.Write("10) ") : PassFail(Int(dbl) = -12345679)
            dbl = 12345678.12 : Console.Write("11) ") : PassFail(Int(dbl) = 12345678)
        End Sub



        Sub HexTest()
            On Error Resume Next

            Console.WriteLine("Hex tests")
            l = -1 : Console.Write("1a) ") : PassFail(Hex(l) = "FFFFFFFFFFFFFFFF")
            l = 0 : Console.Write("2a) ") : PassFail(Hex(l) = "0")
            l = 1 : Console.Write("3a) ") : PassFail(Hex(l) = "1")
            l = &H1A2B3C4D1A2B3C4DL
            : Console.Write("4a) ") : PassFail(Hex(l) = "1A2B3C4D1A2B3C4D")

            i = -1 : Console.Write("1b) ") : PassFail(Hex(i) = "FFFFFFFF")
            i = 0 : Console.Write("2b) ") : PassFail(Hex(i) = "0")
            i = 1 : Console.Write("3b) ") : PassFail(Hex(i) = "1")
            i = &H1A2B3C4DL
            : Console.Write("4b) ") : PassFail(Hex(i) = "1A2B3C4D")

            sh = -1 : Console.Write("1c) ") : PassFail(Hex(sh) = "FFFF")
            sh = 0 : Console.Write("2c) ") : PassFail(Hex(sh) = "0")
            sh = 1 : Console.Write("3c) ") : PassFail(Hex(sh) = "1")
            sh = &H1A2B : Console.Write("4c) ") : PassFail(Hex(sh) = "1A2B")

            dbl = -1 : Console.Write("1d) ") : PassFail(Hex(dbl) = "FFFFFFFF")
            dbl = -2 : Console.Write("2d) ") : PassFail(Hex(dbl) = "FFFFFFFE")
            dbl = 0 : Console.Write("3d) ") : PassFail(Hex(dbl) = "0")
            dbl = -1L + System.Int32.MinValue : Console.Write("4d) ") : PassFail(Hex(dbl) = "FFFFFFFF7FFFFFFF")
            dbl = +1L + System.Int32.MaxValue : Console.Write("5d) ") : PassFail(Hex(dbl) = "80000000")
            dbl = &H1A2B : Console.Write("6d) ") : PassFail(Hex(dbl) = "1A2B")

            dec = -1 : Console.Write("1e) ") : PassFail(Hex(dec) = "FFFFFFFF")
            dec = -2 : Console.Write("2e) ") : PassFail(Hex(dec) = "FFFFFFFE")
            dec = 0 : Console.Write("3e) ") : PassFail(Hex(dec) = "0")
            dec = -1L + System.Int32.MinValue : Console.Write("4e) ") : PassFail(Hex(dec) = "FFFFFFFF7FFFFFFF")
            dec = +1L + System.Int32.MaxValue : Console.Write("5e) ") : PassFail(Hex(dec) = "80000000")
            dec = &H1A2B : Console.Write("6e) ") : PassFail(Hex(dec) = "1A2B")

            s = "-1" : Console.Write("1f) ") : PassFail(Hex(s) = "FFFFFFFF")
            s = "-2" : Console.Write("2f) ") : PassFail(Hex(s) = "FFFFFFFE")
            s = "0" : Console.Write("3f) ") : PassFail(Hex(s) = "0")
            s = "32768" : Console.Write("4f) ") : PassFail(Hex(s) = "8000")
            s = "2147483648" : Console.Write("5f) ") : PassFail(Hex(s) = "80000000")
            s = "&H1A2B" : Console.Write("6f) ") : PassFail(Hex(s) = "1A2B")

            sng = -1 : Console.Write("1g) ") : PassFail(Hex(sng) = "FFFFFFFF")
            sng = -2 : Console.Write("2g) ") : PassFail(Hex(sng) = "FFFFFFFE")
            sng = 0 : Console.Write("3g) ") : PassFail(Hex(sng) = "0")
            'Cannot test with -1L + System.Int32.MinValue since it loses precision and doesn't print Int64 value
            sng = -2147490000.0 : Console.Write("4g) ") : PassFail(Hex(sng) = "FFFFFFFF7FFFE700")
            sng = +1L + System.Int32.MaxValue : Console.Write("5g) ") : PassFail(Hex(sng) = "80000000")
            sng = &H1A2B : Console.Write("6g) ") : PassFail(Hex(sng) = "1A2B")
        End Sub



        Sub OctTest()
            On Error Resume Next

            'UNDONE: COM+ bug in retail crashes using Oct(), so changed to Oct()
            ' raid against COM+ and then return to using Oct when fixed
            Console.WriteLine("Oct tests")
            l = -1 : Console.Write("1a) ") : PassFail(Oct(l) = "1777777777777777777777")
            l = 0 : Console.Write("2a) ") : PassFail(Oct(l) = "0")
            l = 1 : Console.Write("3a) ") : PassFail(Oct(l) = "1")
            l = &O1234567012345670123456L
            : Console.Write("4a) ") : PassFail(Oct(l) = "1234567012345670123456")

            i = -1 : Console.Write("1b) ") : PassFail(Oct(i) = "37777777777")
            i = 0 : Console.Write("2b) ") : PassFail(Oct(i) = "0")
            i = 1 : Console.Write("3b) ") : PassFail(Oct(i) = "1")
            i = &O12345670123I
            : Console.Write("4b) ") : PassFail(Oct(i) = "12345670123")

            sh = -1 : Console.Write("1c) ") : PassFail(Oct(sh) = "177777")
            sh = 0 : Console.Write("2c) ") : PassFail(Oct(sh) = "0")
            sh = 1 : Console.Write("3c) ") : PassFail(Oct(sh) = "1")
            sh = &O123456S : Console.Write("4c) ") : PassFail(Oct(sh) = "123456")

            dbl = -1 : Console.Write("1d) ") : PassFail(Oct(dbl) = "37777777777")
            dbl = -2 : Console.Write("2d) ") : PassFail(Oct(dbl) = "37777777776")
            dbl = 0 : Console.Write("3d) ") : PassFail(Oct(dbl) = "0")
            dbl = -1L + System.Int32.MinValue : Console.Write("4d) ") : PassFail(Oct(dbl) = "1777777777757777777777")
            dbl = +1L + System.Int32.MaxValue : Console.Write("5d) ") : PassFail(Oct(dbl) = "20000000000")
            dbl = &H1A2B : Console.Write("6d) ") : PassFail(Oct(dbl) = "15053")

            dec = -1 : Console.Write("1e) ") : PassFail(Oct(dec) = "37777777777")
            dec = -2 : Console.Write("2e) ") : PassFail(Oct(dec) = "37777777776")
            dec = 0 : Console.Write("3e) ") : PassFail(Oct(dec) = "0")
            dec = -1L + System.Int32.MinValue : Console.Write("4e) ") : PassFail(Oct(dec) = "1777777777757777777777")
            dec = +1L + System.Int32.MaxValue : Console.Write("5e) ") : PassFail(Oct(dec) = "20000000000")
            dec = &H1A2B : Console.Write("6e) ") : PassFail(Oct(dec) = "15053")

            s = "-1" : Console.Write("1f) ") : PassFail(Oct(s) = "37777777777")
            s = "-2" : Console.Write("2f) ") : PassFail(Oct(s) = "37777777776")
            s = "0" : Console.Write("3f) ") : PassFail(Oct(s) = "0")
            s = "32768" : Console.Write("4f) ") : PassFail(Oct(s) = "100000")
            s = "2147483648" : Console.Write("5f) ") : PassFail(Oct(s) = "20000000000")
            s = "&H1A2B" : Console.Write("6f) ") : PassFail(Oct(s) = "15053")

            sng = -1 : Console.Write("1g) ") : PassFail(Oct(sng) = "37777777777")
            sng = -2 : Console.Write("2g) ") : PassFail(Oct(sng) = "37777777776")
            sng = 0 : Console.Write("3g) ") : PassFail(Oct(sng) = "0")
            'Cannot test with -1L + System.Int32.MinValue since it loses precision and doesn't print Int64 value
            sng = -2147490000.0 : Console.Write("4g) ") : PassFail(Oct(sng) = "1777777777757777763400")
            sng = +1L + System.Int32.MaxValue : Console.Write("5g) ") : PassFail(Oct(sng) = "20000000000")
            sng = &H1A2B : Console.Write("6g) ") : PassFail(Oct(sng) = "15053")
        End Sub



        Sub StrTest()
            On Error Resume Next

            Console.WriteLine("Str tests")
            Console.Write("1) ") : PassFail(Str(0) = " 0")
            Console.Write("2) ") : PassFail(Str(123) = " 123")
            Console.Write("3) ") : PassFail(Str(-123) = "-123")
            Console.Write("4) ") : PassFail(Str(123456) = " 123456")
            Console.Write("5) ") : PassFail(Str(-123456) = "-123456")
            Console.Write("6) ") : PassFail(Str(0.0) = " 0")
            Console.Write("7) ") : PassFail(Str(123.456) = " 123.456")
            Console.Write("8) ") : PassFail(Str(-123.456) = "-123.456")
            Console.Write("9) ") : PassFail(Str(12345.6789) = " 12345.6789")
            Console.Write("10) ") : PassFail(Str(-12345.6789) = "-12345.6789")
            Console.Write("11) ") : PassFail(Str(1234560000.0) = " 1234560000")
            Console.Write("12) ") : PassFail(Str(-1234560000.0) = "-1234560000")
            Console.Write("13) ") : PassFail(Str(1.23456E+99) = " 1.23456E+99")
            Console.Write("14) ") : PassFail(Str(-1.23456E+99) = "-1.23456E+99")
            Console.Write("15) ") : PassFail(Str(0.12345) = " .12345")
            Console.Write("16) ") : PassFail(Str(-0.12345) = "-.12345")
            obj = 12345 : Console.Write("17) ") : PassFail(TypeOf Str(obj) Is String)
            'obj = Null : Console.Write( "1) " ) : PassFail(IsDBNull(Str(obj)))
            'obj = Empty : Console.Write( "2) " ) : PassFail(Str(obj) = " 0")
#If False Then 'UNDONE,stephwe,9/12/02: VSWhidbey:28645
            Err.Number = 0 : Console.Write("18) ") : Console.WriteLine(Str("ABC")) : ErrorCheck(13)
#End If

            Console.Write("19) ") : PassFail(Str(True) = "True")
            Console.Write("20) ") : PassFail(Str(False) = "False")
        End Sub



        Sub ValTest()

            Try
                Console.WriteLine("Val tests")
                Console.Write("1) ") : PassFail(Val("0") = 0)
                Console.Write("2) ") : PassFail(Val("123") = 123)
                Console.Write("3) ") : PassFail(Val("-123") = -123)
                Console.Write("4) ") : PassFail(Val("123456") = 123456)
                Console.Write("5) ") : PassFail(Val("-123456") = -123456)
                Console.Write("6) ") : PassFail(Val("0.0") = 0)
                Console.Write("7) ") : PassFail(Val("123.456") = 123.456)
                Console.Write("8) ") : PassFail(Val("-123.456") = -123.456)
                Console.Write("9) ") : PassFail(Val("12345.6789") = 12345.6789)
                Console.Write("10) ") : PassFail(Val("-12345.6789") = -12345.6789)
                Console.Write("11) ") : PassFail(Val("1.23456e+9") = 1234560000)
                Console.Write("12) ") : PassFail(Val("-1.23456e+9") = -1234560000)
                Console.Write("13) ") : PassFail(Val("1.23456e+99") = 1.23456E+99)
                Console.Write("14) ") : PassFail(Val("-1.23456e+99") = -1.23456E+99)
                Console.Write("15) ") : PassFail(Val("&HABCD") = &HABCDS)
                Console.Write("16) ") : PassFail(Val("&H1234ABCD&") = &H1234ABCD)
                Console.Write("17) ") : PassFail(Val(" &O1234") = &O1234)
                Console.Write("18) ") : PassFail(Val(" &O76543210&") = &O76543210)
                Console.Write("19) ") : PassFail(Val("&H7FFFFFFFFFFFFFG") = 36028797018963967)
                Console.Write("20) ") : PassFail(Val("&O777777777777778") = 4398046511103)
                Console.Write("21) ") : PassFail(Val("") = 0)
                Console.Write("22) ") : PassFail(Val(" ") = 0)
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.Write("Bug8636: ")
                s = "1" & ChrW(11) & "2"
                PassFail(Val(s) = 1)
            Catch ex As Exception
                Failed(ex)
            End Try


            Try
                Console.Write("Bug31431: ")
                dbl = Val("1.797693134862315D400")
                Failed()
            Catch ex As Exception When Err.Number = 6
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

            Err.Number = 0
            Try
                Console.Write("Bug31435: ")
                Call Val("123.456%")
                Failed()
            Catch ex As Exception When Err.Number = 13
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

            Dim strvar As String
            Console.WriteLine("Bug293428: ")
            Try
                Console.Write("   1) ")
                strvar = "1.0E-14"
                PassFail(CStr(Val(strvar)) = "1E-14")
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.Write("   2) ")
                strvar = "1E-14"
                PassFail(CStr(Val(strvar)) = "1E-14")
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.Write("   3) ")
                strvar = "1.1E-14"
                PassFail(CStr(Val(strvar)) = "1.1E-14")
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.Write("   4) ")
                strvar = "1.0E+14"
                PassFail(CDbl(strvar) = Val(strvar))
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.Write("   5) ")
                strvar = "1E+14"
                PassFail(CDbl(strvar) = Val(strvar))
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.Write("   6) ")
                strvar = "1.1E+14"
                PassFail(CDbl(strvar) = Val(strvar))
            Catch ex As Exception
                Failed(ex)
            End Try

            Try
                Console.Write("Bug270951: ")
                Dim o As Object
                o = 1.0E-308
                PassFail(CDbl(o) = Val(o))
            Catch ex As Exception
                Failed(ex)
            End Try

        End Sub



        Sub FixTest()
            On Error Resume Next
            Console.WriteLine("Fix tests")

            Dim sng As Single
            sng = 5.5 : Console.Write("1) ") : PassFail(Fix(sng) = 5)
            sng = 6.5 : Console.Write("2) ") : PassFail(Fix(sng) = 6)
            sng = 0 : Console.Write("3) ") : PassFail(Fix(sng) = 0)
            sng = -0.4 : Console.Write("4) ") : PassFail(Fix(sng) = 0)
            sng = -0.5 : Console.Write("5) ") : PassFail(Fix(sng) = 0)
            sng = -1 : Console.Write("6) ") : PassFail(Fix(sng) = -1)
            sng = -1.1 : Console.Write("7) ") : PassFail(Fix(sng) = -1)
            sng = -1.5 : Console.Write("8) ") : PassFail(Fix(sng) = -1)
            sng = -2.5 : Console.Write("9) ") : PassFail(Fix(sng) = -2)
        End Sub



        Sub ErrorTest()
            On Error Resume Next
            Console.WriteLine("Error tests")
            Resume Next                         'Test Resume without error
            Console.WriteLine(Err.Number & " " & Err.Description)
        End Sub



        Sub ErrorStrTest()
            Dim i As Integer

            On Error Resume Next
            Console.WriteLine("ErrorToString tests")
            Console.WriteLine(">" & ErrorToString() & "<")

            For i = 0 To 99
                Console.WriteLine(ErrorToString(i))
            Next i
        End Sub



        Sub ErrorCheck(ByVal iErrExpected As Long)
            If iErrExpected = Err.Number Then
                Console.WriteLine("passed")
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
            Console.Write("Bug" & CStr(m_lBugID) & ": ")
        End Sub



        Sub ChrTest()
            Console.WriteLine("** ChrTest **")
            Dim c As Char
            Dim i As Integer = 0

            'c = Chr(-1)          : PassFail(Asc(c) = -1)
            'c = Chr(-40000)
            c = Chr(1) : PassFail(Asc(c) = 1)
            c = Chr(i) : PassFail(Asc(c) = i)
            c = Chr(128) : PassFail(Asc(c) = 128)
            c = Chr(Nothing) : PassFail(Asc(c) = Nothing)
        End Sub



        Sub ChrWTest()
            Console.WriteLine("** ChrWTest **")
            Dim c As Char
            Dim i As Integer = 0

            'c = ChrW(-1)
            'c = ChrW(-40000)
            c = ChrW(1) : PassFail(AscW(c) = 1)
            c = ChrW(i) : PassFail(AscW(c) = i)
            c = ChrW(128) : PassFail(AscW(c) = 128)
            c = ChrW(Nothing) : PassFail(AscW(c) = Nothing)
        End Sub



        Sub AscTest()
            Console.WriteLine("** AscTest **")
            Dim s As String
            Dim c As Char
            Dim i As Integer

            i = Asc(c) : PassFail(Chr(i) = c)
            i = Asc("Ç"c) : PassFail(Chr(i) = "Ç"c)
            i = Asc("H"c) : PassFail(Chr(i) = "H"c)
            i = Asc("H") : PassFail(Chr(i) = "H"c)
            'i = Asc("")
            i = Asc(Nothing) : PassFail(Chr(i) = Nothing)
            i = Asc("Hello") : PassFail(Chr(i) = "H"c)
            i = Asc("Çsdf") : PassFail(Chr(i) = "Ç"c)
            'i = Asc(3)

            Try
                i = Asc(s)
            Catch
                Console.WriteLine("passed: exception caught")
            End Try
        End Sub



        Sub AscWTest()
            Console.WriteLine("** AscWTest **")
            Dim s As String
            Dim c As Char
            Dim i As Integer

            i = AscW(c) : PassFail(ChrW(i) = c)
            i = AscW("Ç"c) : PassFail(ChrW(i) = "Ç"c)
            i = AscW("H"c) : PassFail(ChrW(i) = "H"c)
            i = AscW("H") : PassFail(ChrW(i) = "H"c)
            'i = AscW("")
            i = AscW(Nothing) : PassFail(ChrW(i) = Nothing)
            i = AscW("Hello") : PassFail(ChrW(i) = "H"c)
            i = AscW("Çsdf") : PassFail(ChrW(i) = "Ç"c)
            'i = AscW(3)

            Try
                i = Asc(s)
            Catch
                Console.WriteLine("passed: exception caught")
            End Try
        End Sub



        Public Const cx As Char = ChrW(1)
        Public Const cy As Char = ChrW(-1)
        Public Const cj As Char = ChrW(-32768)
        Public Const ck As Char = ChrW(CShort(-32768))
        Public Const cl As Char = ChrW(&H8000S)
        Public Const cm As Char = ChrW(65535)

        Public Const ca As Char = Chr(0)
        Public Const cc As Char = Chr(127)

        Public Const cd As Integer = AscW("ab")
        Public Const ce As Integer = AscW("a"c)

        Public Const cg As Integer = Asc("ab")
        Public Const ch As Integer = Asc("a"c)



        Sub AscAndChrConstTest()
            Console.WriteLine("** AscAndChrConstTest **")

            Console.WriteLine(cx)
            Console.WriteLine(cy)
            Console.WriteLine(cj)
            Console.WriteLine(ck)
            Console.WriteLine(cl)
            Console.WriteLine(cm)
            Console.WriteLine(ca)
            Console.WriteLine(cc)
            Console.WriteLine(cd)
            Console.WriteLine(ce)
            Console.WriteLine(cg)
            Console.WriteLine(ch)
        End Sub



    End Module


    Module Bug229397

        Sub Test()
            Console.WriteLine("*** Bug229397: ")
            Try
            
                Const CS1 = 2.372E+13!
                Dim sng As Single = CS1

                Console.Write("    ")
                PassFail(Format(sng) = "2.372E+13")

                Console.Write("    ")
                PassFail(CStr(sng) = "2.372E+13")

                Console.Write("    ")
                PassFail(Sng.ToString() = "2.372E+13")

            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub

    End Module

    Module Bug249566

        Sub test()
            Console.WriteLine("Bug 249566")
            Try
                Console.Write("   ") : PassFail(Val("&HF") = 15)
                Console.Write("   ") : PassFail(Val("&HFF") = 255)
                Console.Write("   ") : PassFail(Val("&HFFFF") = -1)
                Console.Write("   ") : PassFail(Val("&HFFFFFFFF") = -1)
                Console.Write("   ") : PassFail(Val("&HFFFFFFFFFFFFFFFF") = -1)
                Console.Write("   ") : PassFail(Val("&HF%") = 15)
                Console.Write("   ") : PassFail(Val("&HFF%") = 255)
                Console.Write("   ") : PassFail(Val("&HFFFF&") = 65535)
                Console.Write("   ") : PassFail(Val("&HFFFFFFFFL") = -1)
            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub

    End Module

End Namespace



