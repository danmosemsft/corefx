Option Compare Binary

Imports System
Imports Microsoft.VisualBasic
Imports System.Globalization

Class MyCulture
    Inherits CultureInfo

    Private m_nfi As NumberFormatInfo
    Private m_dtfi As DateTimeFormatInfo



    Shared Sub Init(ByVal lcid As Integer)
        'Set culture to Enlish to avoid locale diffs
        Dim ci As CultureInfo = New CultureInfo(&H409)
        Dim nfi As NumberFormatInfo = CType(ci.GetFormat(GetType(NumberFormatInfo)), NumberFormatInfo)
        Dim dtfi As DateTimeFormatInfo = CType(ci.GetFormat(GetType(DateTimeFormatInfo)), DateTimeFormatInfo)

        System.Threading.Thread.CurrentThread.CurrentCulture = New MyCulture(&H409, dtfi, nfi)
    End Sub



    Sub New(lcid As Integer, dtfi As DateTimeFormatInfo, nfi As NumberFormatInfo)
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



Module Test
    Dim b As Boolean
    Dim i As Integer
    Dim s As String
    Dim TestNumber As Integer
    Dim TestCounter As Integer
    Dim ch As Char
    Dim m_sPrintInputFileName As String
    Dim m_sFileName As String
    Dim m_sFileNameDynamic As String
    Dim m_sFileNameLenCheck As String
    Public m_sTestRootPath As String
    Public m_sTestTempPath As String



    Sub Main()
        MyCulture.Init(&H409)
        Console.WriteLine("Begin Tests")
        m_sTestRootPath = Environ("TESTDIR")	

        if m_sTestRootPath.Length <> 0 Then
            m_sTestRootPath = m_sTestRootPath & "\"
        End If
        m_sTestTempPath = m_sTestRootPath & "temp\"

	if m_sTestTempPath = "" OrElse not System.IO.Directory.Exists(m_sTestRootPath) then
		failed()
		exit sub
	end if

	'Delete the temp directory and its subdirectories, if they exist, then
	'  create the temp directory anew
	Try
		System.IO.Directory.Delete(m_sTestTempPath, True)
	Catch ex As System.IO.DirectoryNotFoundException
		'Ignore if doesn't exist
	End Try
	System.IO.Directory.CreateDirectory(m_sTestTempPath)

        Try
            TestOutput
            WriteInputTests
            GetPutTests
            LockTests
            DirTests
            Bug170856.Main
            EOFTests
            LOCTests
            Bug258106
            Bug280613.Test
            Bug289099.Test
            Bug298737.Test
            Bug300062.Test
            Bug300161.Test
            Bug340556.Test
            Bug550463.Test
        Catch ex As Exception
            Console.WriteLine("Unhandled exception in Sub Main")
            Failed(ex)
        End Try

        Console.WriteLine("End Tests")
    End Sub



    Sub  Bug258106()
        Dim FileNumber As Integer = FreeFile

        m_sFileName = m_sTestTempPath & "opentest.out"
        Console.Write( "Bug258106: " )

        Try
            FileOpen( FreeFile, m_sFileName, OpenMode.Binary, OpenAccess.Write)
            Passed
            FileClose( FreeFile )
        Catch ex As Exception
            Failed
        End Try
    End Sub



    Sub EOFTests()
        Dim strTmp As String
        Dim FileNumber As Integer = FreeFile
        Dim i As Integer

        m_sFileName = m_sTestTempPath & "eoftest.out"
        
        Console.WriteLine( "*** EOF test" )

        Try
            Kill(m_sFileName)
        Catch ex As Exception
        End Try

        FileOpen(FileNumber, m_sFileName, OpenMode.Output)
        PrintLine(FileNumber, Strdup(1000, "X"))
        Console.Write("1): ")
        PassFail( EOF(FileNumber) = True)

        FileClose(FileNumber)

        FileOpen(FileNumber, m_sFileName, OpenMode.Input)
        Console.Write("2): ")
        PassFail( EOF(FileNumber) = False)

        input(FileNumber, strtmp)
        Console.Write("3): ")
        PassFail( EOF(FileNumber) = True)
                
        Seek(FileNumber, 500)         
        Console.Write("4): ")
        PassFail( EOF(FileNumber) = False)
        FileClose(FileNumber)

        FileOpen(FileNumber, m_sFileName, OpenMode.Output)

        For i = 1 To 10
            Write(FileNumber, i)
        Next I

        FileClose(FileNumber)

        FileOpen(FileNumber, m_sFileName, OpenMode.Input)
        Input(FileNumber, i)
        Console.Write("5): ")
        PassFail( EOF(FileNumber) = False)
        FileClose(FileNumber)
    End Sub



    Sub LOCTests()
        Dim strTmp As String
        Dim FileNumber As Integer = FreeFile()
        Dim i As Integer

        m_sFileName = m_sTestTempPath & "loctest.out"
        Console.WriteLine( "*** LOC test" )

        Try
            Kill(m_sFileName)
        Catch ex As Exception
        End Try

        FileOpen(FileNumber, m_sFileName, OpenMode.Output)

        For i = 1 To 9
            WriteLine(FileNumber, i)
        Next i

        FileClose(FileNumber)


        FileOpen(FileNumber, m_sFileName, OpenMode.Binary)

        For i = 1 To 9
            Console.Write(i & "): ")
            Input(FileNumber, i)
            PassFail(Loc(FileNumber) = i*3)
        Next i

        FileClose(FileNumber)
    End Sub



    Sub DirTests()
        Dim TestDir As String = "TestDir"
        Dim TestFile1 As String = "TestFile1.dat"
        Dim TestFile2 As String = "TestFile2.dat"
        Dim DirPath As String = m_sTestTempPath & TestDir
        Dim TestPath1 As String = DirPath & "\" & TestFile1
        Dim TestPath2 As String = DirPath & "\" & TestFile2
        Dim s As String
        Dim FileNumber As Integer = FreeFile
        Dim f As FileAttribute

        Console.WriteLine( "*** Dir test" )

        'Ensure everything is cleaned from prior runs
        Try
            Kill( TestPath1 )
        Catch
        End Try

        Try
            SetAttr( TestPath2, FileAttribute.Normal  )
            Kill( TestPath2 )
        Catch
        End Try

        Try
            RmDir( DirPath )
        Catch ex As Exception
            'Console.WriteLine( "Cannot clean out the old TestDir! " & ex.Message )
            'Exit Sub
        End Try

        Try
            Console.Write( "MkDir of null path: " )
            MkDir( "" )
            Failed
        Catch e1 As ArgumentException
            Passed
        Catch e2 As Exception
            Failed( e2 )
        End Try

        Try
            'Create a test directory
            Console.Write( "MkDir: " )
            MkDir( DirPath )
            Passed

            'Execute dir on empty directory
            Console.Write( "Dir of empty directory: " )
            s = Dir( DirPath )
            PassFail( s, "" )

            'Execute dir on empty directory
            Console.Write( "Dir of empty directory: " )
            s = Dir( DirPath & "\" )
            PassFail( s, "" )

            'Add 2 files 
            Console.Write( "Creating two files: ")
            FileOpen( FileNumber, TestPath1, OpenMode.Output ) 
            FileClose( FileNumber )

            FileOpen( FileNumber, TestPath2, OpenMode.Output ) 
            FileClose( FileNumber )

            Passed

            Console.Write( "Dir of first file: " )
            s = Dir( DirPath & "\*.*" )
            PassFail( s, TestFile1 )

            Console.Write( "Dir of second file: " )
            s = Dir()
            PassFail( s, TestFile2 )

            Console.Write( "Dir beyond end of filelist: " )
            s = Dir()
            PassFail( s, "" )

            Console.Write( "Setting hidden attribute: " )
            SetAttr( TestPath1, FileAttribute.Hidden  )
            Passed

            Console.Write( "Dir of hidden file: " )
            s = Dir( DirPath & "\*.dat", FileAttribute.Hidden )
            PassFail( s, TestFile1 )

            Console.Write( "GetAttr of hidden file: " )
            f = GetAttr(TestPath1)
            PassFail( f, vbHidden)

            Console.Write( "Dir of normal file: " )
            s = Dir()
            PassFail( s, TestFile2 )

            Console.Write( "GetAttr of normal file: " )
            f = GetAttr(TestPath2)
            PassFail( f, vbArchive )

            Console.Write( "Removing hidden attribute: " )
            SetAttr( TestPath1, FileAttribute.Normal  )
            Passed

            Console.Write( "Dir of '.' first file: " )
            s = Dir( DirPath & "\." )
            PassFail( s, TestFile1 )

            Console.Write( "Dir of '.' second file: " )
            s = Dir()
            PassFail( s, TestFile2 )

            Console.WriteLine( "Dir of .. with Directory attribute: " )
            s = Dir( DirPath & "\..", FileAttribute.Directory )

            While s <> ""
                If (GetAttr( DirPath & "\..\" & s ) And vbDirectory) = vbDirectory
                    Console.WriteLine( s )
                End If

                s = Dir()
            End While

            ChDir(DirPath)
            Console.Write( "Dir of first file using """": " )
            s = Dir( "" )
            PassFail( s, TestFile1 )

        Catch ex As Exception
            Failed(ex)
            Console.WriteLine( ex.Message )
        End Try

        Try
            Console.Write( "Dir of nonexistant path: " )
            s = Dir( DirPath & "\NonExistant" )
            PassFail( s, "" )
        Catch e1 As ArgumentException
            Passed
        Catch e2 As Exception
            Failed( e2 )
        End Try

        Console.WriteLine("*** End Dir test" )
    End Sub



    Sub LockTests()
        Dim s As String

        Console.WriteLine( "*** Lock test" )

        Try
            Dim FileNumber As Integer = FreeFile
            FileOpen( FileNumber, m_sTestTempPath & "foo1.txt", OpenMode.Output, OpenAccess.Default, OpenShare.Default, 1)
            Dim FileNumber2 As Integer = FreeFile

            PrintLine( FileNumber, "ABCDEFGHIJKLMNOPQRSTUVWXYZ" )
            PrintLine( FileNumber, "ABCDEFGHIJKLMNOPQRSTUVWXYZ" )
            PrintLine( FileNumber, "ABCDEFGHIJKLMNOPQRSTUVWXYZ" )
            PrintLine( FileNumber, "ABCDEFGHIJKLMNOPQRSTUVWXYZ" )

            Lock( FileNumber )
            Console.WriteLine( "Locked entire file" )
            Unlock( FileNumber )
            Console.WriteLine( "Unlocked entire file" )

            Lock( FileNumber, 1 )
            Console.WriteLine( "Locked 1 record" )
            Unlock( FileNumber, 1 )
            Console.WriteLine( "Unlocked 1 record" )

            Lock( FileNumber, 2, 2 )
            Console.WriteLine( "Locked 2 records" )
            Unlock( FileNumber, 2, 2 )
            Console.WriteLine( "Unlocked 2 records" )
            FileClose( FileNumber )

            Console.WriteLine( "Trying to input when file is locked" )
            FileOpen(FileNumber, m_sTestTempPath & "foo1.txt", OpenMode.Input, , OpenShare.Shared)
            FileOpen(FileNumber2, m_sTestTempPath & "foo1.txt", OpenMode.Input, , OpenShare.Shared)
            Lock(FileNumber)

            Try
                Input(FileNumber2, s) 'This should fail with permission denied
                Failed()
            Catch ex As System.IO.IOException
                Console.WriteLine( "Permission denied as expected" )  
            Catch ex As Exception
                Failed(ex)
            End Try

            FileClose()
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine( "*** End Lock test" )
    End Sub



    Sub TestOutput()
        Dim shrt As Short
        Dim int As Integer
        Dim lng As Long
        Dim sng As Single
        Dim dbl As Double
        Dim dt As Date
        Dim dec As Decimal
        Dim bool As Boolean
        Dim FileName As String
        Dim FileNumber As Integer

        Console.WriteLine( "*** Output test" )

        Try
            FileNumber = FreeFile
            shrt = 1
            int = 2
            lng = 3
            sng = CSng(4.5)
            dbl = 6.7
            dt = New Date(2000, 4, 27)
            dec = New Decimal(CDbl(1234.5678))
            bool = True
            FileName = Environ("TESTDIR")

            If FileName.Length <> 0 Then
                FileName = FileName & "\"
            End If

            FileName = m_sTestTempPath & "OutputTest.tmp"
            FileOpen( FileNumber, FileName, OpenMode.Output, OpenAccess.Write, OpenShare.LockReadWrite )
            PrintLine( FileNumber, "ABC", bool, shrt, int, lng, sng, dbl, dt, dec )
            PrintLine( FileNumber, "ABC", TAB, bool, TAB, shrt, TAB, int, TAB, lng, TAB, sng, TAB, dbl, TAB, dt, TAB, dec )
            Print( FileNumber, "ABC" )
            Print( FileNumber, bool )
            Print( FileNumber, shrt )
            Print( FileNumber, int )
            Print( FileNumber, lng )
            Print( FileNumber, sng )
            Print( FileNumber, dbl )
            Print( FileNumber, dt )
            PrintLine( FileNumber, dec )

            PrintLine( FileNumber, "Last", Tab(40), "First", "MI", SPC(20), "Date Of Birth" )
            PrintLine( FileNumber, "Customer", TAB(40), "John", "Q", SPC(20), #4/27/2000# )
            PrintLine( FileNumber, int )
            PrintLine( FileNumber )

	        PrintLine( FileNumber, "a", Tab(6), "b") 'Eqv to VB6 'Print #1, "a"; Tab(6); "b"
	        PrintLine( FileNumber, "a", Tab, "b")	
	        PrintLine( FileNumber, "a", Spc(10), "b")

	        PrintLine( FileNumber, "a", TAB, Tab(6), TAB, "b") 'Eqv to VB6 'Print #1, "a"; Tab(6); "b"
	        PrintLine( FileNumber, "a", TAB, Tab, TAB, "b")	
	        PrintLine( FileNumber, "a", TAB, Spc(10), TAB, "b")
        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
            Console.WriteLine(ex.StackTrace)
        Finally
            FileClose( FileNumber )
        End Try

        Console.WriteLine( "*** End Output test" )
    End Sub



    Private Sub TestInputFunction()
        Dim FileNumber As Integer
        Dim caught As Boolean

        Console.WriteLine( "*** InputString function test" )
        FileNumber = FreeFile

        Try
            Dim s As String

            FileNumber = FreeFile

            FileOpen( FileNumber, m_sTestTempPath & "foo.txt", OpenMode.Output, OpenAccess.Write, OpenShare.LockReadWrite )
            PrintLine( FileNumber, "ABCDEFGHIJKLMNOPQRSTUVWXYZ" )
            FileClose( FileNumber )

            s = "1234567890"
            FileOpen( FileNumber, m_sTestTempPath & "foo.txt", OpenMode.Input, OpenAccess.Read, OpenShare.LockReadWrite  )
            s = InputString(FileNumber, 5)
            FileClose( FileNumber)
            Console.WriteLine( "OpenMode.Input: >" & s & "< length = " & CStr(s.Length) )

            s = "1234567890"
            FileOpen( FileNumber, m_sTestTempPath & "foo.txt", OpenMode.Binary, OpenAccess.Read, OpenShare.LockReadWrite  )
            s = InputString(FileNumber, 9)
            FileClose( FileNumber)
            Console.WriteLine( "OpenMode.Binary: >" & s & "< length = " & CStr(s.Length) )
        
            'Test embedded CTRL-Zs RAID 229130
            FileOpen( FileNumber, m_sTestTempPath & "foo.txt", OpenMode.Output, OpenAccess.Write  )
            PrintLine( FileNumber, "ABC" & Chr(26) & "DEF" )
            FileClose( FileNumber )

            FileOpen( FileNumber, m_sTestTempPath & "foo.txt", OpenMode.Input, OpenAccess.Read  )
            s = ""

            Try
                s = InputString(FileNumber, LOF(FileNumber))
            Catch ex As IO.EndOfStreamException    
                Console.WriteLine( "Reading past embedded CTRL-Zs: passed" )
                caught = True
            End Try

            If Not caught Then
                Console.WriteLine( "Reading past embedded CTRL-Zs: FAILED" )
            End If

            FileClose( FileNumber)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub



    Private Sub WriteInputTests()
        TestWrite( False )
        TestInput( False )
        TestInputFunction
        'TestWrite True
        'TestInput True
    End Sub



    Sub TestWrite(ByVal bBinary As Boolean)
        Dim FileNumber As Integer

        Dim i As Integer
        Dim byt As Byte
        Dim l As Long
        Dim dt As Date
        Dim s As String
        Dim b As Boolean
        Dim dbl As Double
        Dim sng As Single
        Dim dec As Decimal
        Dim cc(4) As Char
        Dim v1, v2, v3, v4, v5, v6, v7, v8, v9, v10 As Object

        Try
            i = 1234
            l = 56789
            dt = #6/28/1999#
            byt = 129
            s = "A 16 char string"
            dbl = 1.234765E+102
            b = True
            sng = 4.5
            dec = 12345.67@
            cc = "asdf"

            m_sPrintInputFileName = m_sTestTempPath & "PrintInputTest-vb7.out"
            Console.WriteLine( "*** Begin Write Test" )
            FileNumber = FreeFile

            If bBinary Then
                FileOpen( FileNumber, m_sPrintInputFileName, OpenMode.Binary, OpenAccess.Write, OpenShare.LockReadWrite)
            Else
                FileOpen( FileNumber, m_sPrintInputFileName, OpenMode.Output )
            End If

            WriteLine( FileNumber, i )
            WriteLine( FileNumber, l )
            WriteLine( FileNumber, dt )
            WriteLine( FileNumber, byt )
            WriteLine( FileNumber, dbl )
            WriteLine( FileNumber, sng )
            WriteLine( FileNumber, b )
            WriteLine( FileNumber, s )
            WriteLine( FileNumber, dec )
            WriteLine( FileNumber, cc )
            WriteLine( FileNumber, i, l, dt, byt, dbl, sng, b, s, dec, cc )

            v1 = i: WriteLine( FileNumber, v1 )
            v2 = l: WriteLine( FileNumber, v2 )
            v3 = dt: WriteLine( FileNumber, v3 )
            v4 = byt: WriteLine( FileNumber, v4 )
            v5 = dbl: WriteLine( FileNumber, v5 )
            v6 = sng: WriteLine( FileNumber, v6 )
            v7 = b: WriteLine( FileNumber, v7 )
            v8 = s: WriteLine( FileNumber, v8 )
            v9 = dec : WriteLine( FileNumber,  v9 )
            v10 = cc : WriteLine( FileNumber,  v10 )

            WriteLine( FileNumber, v1, v2, v3, v4, v5, v6, v7, v8, v9, v10 )
            FileClose( FileNumber )

        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine( "*** End Write Test" )
    End Sub



    Sub TestInput(ByVal bBinary As Boolean)
        Dim FileNumber As Integer
        Dim i As Integer
        Dim byt As Byte
        Dim l As Long
        Dim dt As Date
        Dim s As String
        Dim b As Boolean
        Dim dbl As Double
        Dim sng As Single
        Dim dec As Decimal
        Dim cc(4) As Char
        Dim v As Object

        Console.WriteLine( "*** Begin Input Test" )

        Try
            FileNumber = FreeFile

            If bBinary Then
                'FileOpen( FileNumber, m_sPrintInputFileName, OpenMode.Binary)
            Else
                FileOpen(FileNumber, m_sPrintInputFileName, OpenMode.Input)
            End If

            Console.WriteLine( "*** Input with one data type" )

            Console.Write( "Integer: " )
            Input( FileNumber, i )
            PassFail(i = 1234)

            Console.Write( "Long: " ) 
            Input( FileNumber, l )
            PassFail(l = 56789)

            Console.Write( "Date: " ) 
            Input( FileNumber, dt )
            PassFail(dt = #6/28/1999#)

            Console.Write( "Byte: " ) 
            Input( FileNumber, byt )
            PassFail(byt = 129)

            Console.Write( "Double: " ) 
            Input( FileNumber, dbl )
            PassFail(dbl = 1.234765E+102)

            Console.Write( "Single: " ) 
            Input( FileNumber, sng )
            PassFail(sng = CSng(4.5))

            Console.Write( "Boolean: " ) 
            Input( FileNumber, b )
            PassFail(b = True)

            Console.Write( "String: " ) 
            Input( FileNumber, s )
            PassFail(s = "A 16 char string")

            Console.Write( "Decimal: " ) 
            Input( FileNumber,  dec )
            PassFail(CBool(dec = 12345.67@))

            Console.Write( "Char array: " ) 
            Input( FileNumber,  cc )
            PassFail(cc = "asdf")


            Console.WriteLine( "*** Input an Object containing different data types" )

            'v must have the expected vartype
            v = i
            Console.Write( "Object(Integer): " )
            Input( FileNumber, v )
            PassFail(CInt(v) = 1234)

            v = l
            Console.Write( "Object(Long): " ) 
            Input( FileNumber, v )
            PassFail(CLng(v) = 56789)

            v = dt 
            Console.Write( "Object(Date): " ) 
            Input( FileNumber, v )
            PassFail(CDate(v) = #6/28/1999#)

            v = byt
            Console.Write( "Object(Byte): " )
            Input( FileNumber, v )
            PassFail(CByte(v) = 129)

            v = dbl
            Console.Write( "Object(Double): " )
            Input( FileNumber, v )
            PassFail(CDbl(v) = 1.234765E+102)

            v = sng
            Console.Write( "Object(Single): " )
            Input( FileNumber, v )
            PassFail(CSng(v) = 4.5)

            v = True 
            Console.Write( "Object(Boolean): " )
            Input( FileNumber, v )  
            PassFail(CBool(v) = True)

            v = "ABC"
            Console.Write( "Object(String): " )
            Input( FileNumber, v )
            PassFail(CStr(v) = "A 16 char string" )

            v = dec
            Console.Write( "Object(Decimal): " ) 
            Input( FileNumber,  v )
            PassFail(CDec(v) = 12345.67)

            v = cc
            Console.Write( "Object(Char array): " ) 
            Input( FileNumber,  v )
            PassFail(CStr(v) = "asdf")

#if false
            Dim v1 As Object
            Dim v2 As Object
            Dim v3 As Object
            Dim v4 As Object
            Dim v5 As Object
            Dim v6 As Object
            Dim v7 As Object
            Dim v8 As Object
            Dim v9 As Object
            Dim v10 As Object
            Dim v11 As Object

            'ByRef ParamArrays not supported in VB7
            v1 = i : v2 = l : v3 = dt : v5 = byt : v6 = dbl : v7 = sng : v8 = b : v9 = "ABC" : v10 = "ABC"

            Console.WriteLine( "*** Input multiple variants containing different data types" )
            Input( FileNumber, v1, v2, v3, v4, v5, v6, v7, v8, v9, v10 )

            Console.Write( "Object(Integer): " ) 
            PassFail(VarType(v1) = VariantType.Integer)
            PassFail(CInt(v1) = 1234)

            Console.Write( "Object(Long): " ) 
            PassFail(VarType(v2) = VariantType.Long)
            PassFail(CLng(v2) = 56789)

            Console.Write( "Object(Date): " ) 
            PassFail(VarType(v3) = VariantType.Date)
            PassFail(CDate(v3) = #6/28/1999#)

            'YES, Object(Byte) should return an Integer type
            Console.Write( "Object(Byte): " ) 
            PassFail(VarType(v5) = VariantType.Integer)
            PassFail(CByte(v5) = 129)

            Console.Write( "Object(Double): " )
            PassFail(VarType(v6) = VariantType.Double)
            PassFail(CDbl(v6) = 1.234765E+102)

            'YES, Object(Single) should return an Double type
            Console.Write( "Object(Single): " ) 
            PassFail(VarType(v7) = VariantType.Double)
            PassFail(CSng(v7) = 4.5)

            Console.Write( "Object(Boolean): " )
            PassFail(VarType(v8) = VariantType.Boolean)
            PassFail(CBool(v8) = True)

            Console.Write( "Object(String): ") 
            PassFail(VarType(v9) = VariantType.String)
            PassFail(CStr(v9) = "A 16 char string" )
#end if
        Catch ex As Exception
            Failed(ex)
        Finally
            FileClose(FreeFile)
        End Try
                    
        Console.WriteLine( "*** End Input Test" ) 
    End Sub



    Sub GetPutTests()
        On Error Resume Next

        VB6CompatGetTest
        VB6CompatPutTest

        VB7OnlyTests
    End Sub



    Sub VB7OnlyTests()
        Console.WriteLine( "*** Put/Get Tests" )
        TestPut
        TestGet
        TestPutGetDynamicFixed
        TestVBFixedArrayAttribute
        TestDynamicArrayPut
        Console.WriteLine( "*** End Put/Get Tests" )
    End Sub



    Structure s1
        Dim i1 As Integer
        Dim a1() As Long
    End Structure

    Structure s2
        Dim i2 As Integer
        <VBFixedArray(3)> Dim a2() As Long
        Dim j As s1
    End Structure

    Structure s3
        Dim i2 As Integer
        <VBFixedArray(5)> Dim a2() As Long
        Dim j As s1
    End Structure

    Structure s4
        Dim i2 As Integer
        <VBFixedArray(1)> Dim a2() As Long
        Dim j As s1
    End Structure


    Sub TestVBFixedArrayAttribute()
        Dim s As s2
        Dim t As s2
        Dim u As s3
        Dim v As s3
        Dim w As s4
        Dim x As s4
        Dim fn As Integer = FreeFile()

        Try
            Console.WriteLine("*** VBFixedArrayAttribute test")
            m_sFileName = m_sTestTempPath & "fixedarray-vb7.out"

            Try
                Kill(m_sFileName)
            Catch
            End Try

            Console.WriteLine("Testing nested structure with matching VBFixedArray dimensions")
            FileOpen(fn, m_sFileName, OpenMode.Random)

            s.i2 = 1
            ReDim s.a2(3)
            s.a2(0) = 2
            s.a2(1) = 3
            s.a2(2) = 4
            s.a2(3) = 5
            s.j.i1 = 6
            ReDim s.j.a1(2)
            s.j.a1(0) = 7
            s.j.a1(1) = 8
            s.j.a1(2) = 9
            FilePut(fn, s)

            ReDim t.a2(3)
            t.a2(0) = 0
            ReDim t.j.a1(2)
            t.j.a1(0) = 0
            FileGet(fn, t, 1)
            PassFail(t.i2 = s.i2 AndAlso t.a2(0) = s.a2(0) AndAlso t.a2(1) = s.a2(1) AndAlso t.a2(2) = s.a2(2) AndAlso t.a2(3) = s.a2(3))
            PassFail(t.j.i1 = s.j.i1 AndAlso t.j.a1(0) = s.j.a1(0) AndAlso t.j.a1(1) = s.j.a1(1) AndAlso t.j.a1(2) = s.j.a1(2))
            FileClose(fn)
            Kill(m_sFileName)

            Console.WriteLine("Testing nested structure with VBFixedArray dimension > array dimension")
            FileOpen(fn, m_sFileName, OpenMode.Random)
            u.i2 = 1
            ReDim u.a2(3)
            u.a2(0) = 2
            u.a2(1) = 3
            u.a2(2) = 4
            u.a2(3) = 5
            u.j.i1 = 6
            ReDim u.j.a1(2)
            u.j.a1(0) = 7
            u.j.a1(1) = 8
            u.j.a1(2) = 9
            FilePut(fn, u)

            ReDim v.a2(3)
            v.a2(0) = 0
            ReDim v.j.a1(2)
            v.j.a1(0) = 0

            FileGet(fn, v, 1)
            PassFail(v.i2 = u.i2 AndAlso v.a2(0) = u.a2(0) AndAlso v.a2(1) = u.a2(1) AndAlso v.a2(2) = u.a2(2) AndAlso v.a2(3) = u.a2(3))
            PassFail(v.j.i1 = u.j.i1 AndAlso v.j.a1(0) = u.j.a1(0) AndAlso v.j.a1(1) = u.j.a1(1) AndAlso v.j.a1(2) = u.j.a1(2))
            PassFail(v.a2(4) = 0 AndAlso v.a2(5) = 0)
            FileClose(fn)
            Kill(m_sFileName)

            Console.WriteLine("Testing nested structure with VBFixedArray dimension < array dimension")
            FileOpen(fn, m_sFileName, OpenMode.Random)
            w.i2 = 1
            ReDim w.a2(3)
            w.a2(0) = 2
            w.a2(1) = 3
            w.a2(2) = 4
            w.a2(3) = 5
            w.j.i1 = 6
            ReDim w.j.a1(2)
            w.j.a1(0) = 7
            w.j.a1(1) = 8
            w.j.a1(2) = 9

            Try
                FilePut(fn, w)
                Failed()
            Catch Ex As ArgumentException
                Passed()
            End Try

        Catch Ex As Exception
            Failed(Ex)
        End Try

        Console.WriteLine("*** End VBFixedArrayAttribute test")
    End Sub


    Structure DynamicArrayStruct
        Public x As Integer()
    End Structure

    Sub TestDynamicArrayPut()
        Dim a, b As DynamicArrayStruct
        Dim fn As Integer = FreeFile()

        Try
            Console.WriteLine("*** Put of dynamic array test")
            m_sFileName = m_sTestTempPath & "dynamicarray-vb7.out"

            Try
                Kill(m_sFileName)
            Catch
            End Try

            Console.Write("   1) ")
            Try
                FileOpen(fn, m_sFileName, OpenMode.Random)
                'Should write out 0 dim array
                FilePut(fn, a)
                Passed()
            Catch ex As Exception
                Failed(ex)
            Finally
                FileClose(fn)
            End Try


            a.x = New Integer() { 1, 2, 3, 4, 5 }

            Console.Write("   2) ")
            Try
                FileOpen(fn, m_sFileName, OpenMode.Random)
                FilePut(fn, a)
                Passed()
            Catch ex As Exception
                Failed(ex)
            Finally
                FileClose(fn)
            End Try

            Console.Write("   3) ")
            Try
                FileOpen(fn, m_sFileName, OpenMode.Random)
                FileGet(fn, b)
                PassFail(b.x.Length = 5 AndAlso b.x(0) = 1 AndAlso b.x(4) = 5)
            Catch ex As Exception
                Failed(ex)
            Finally
                FileClose(fn)
            End Try

        Catch Ex As Exception
            Failed(Ex)
        End Try

        Console.WriteLine("*** End Put of dynamic array test")
    End Sub


    Sub VB6CompatPutTest()
        'Data types to test
        Dim sh As Short
        Dim i As Integer
        Dim byt As Byte
        Dim l As Long
        Dim dt As Date
        Dim s As String
        Dim b As Boolean
        Dim dbl As Double
        Dim sng As Single
        Dim v As Object
        Dim dec As Decimal
        Dim FileNumber As Integer = FreeFile

        sh = 1234
        i = 56789
        dt = #6/28/1999#
        byt = 129
        s = "A 16 char string"
        dbl = 1234567.7654321
        b = True
        sng = 1234.567
        dec = 12345.67@

        Console.WriteLine("*** VB6 Put compatibility test")
        m_sFileName = m_sTestTempPath & "getputtest-vb7.out"

        Try
            Kill( m_sFileName)
        Catch
        End Try

        FileOpen(FileNumber, m_sFileName, OpenMode.Random)

        FilePut(FileNumber, sh)
        FilePut(FileNumber, i)
        FilePut(FileNumber, dt)
        FilePut(FileNumber, byt)
        FilePut(FileNumber, dbl)
        FilePut(FileNumber, sng)
        FilePut(FileNumber, b)
        FilePut(FileNumber, s)
        'VB6 Decimal cannot be written out except in variant

        v = sh : FilePutObject(FileNumber, v)
        v = i : FilePutObject(FileNumber, v)
        v = dt : FilePutObject(FileNumber, v)
        v = byt : FilePutObject(FileNumber, v)
        v = dbl : FilePutObject(FileNumber, v)
        v = sng : FilePutObject(FileNumber, v)
        v = b : FilePutObject(FileNumber, v)
        v = s : FilePutObject(FileNumber, v)
        v = dec : FilePutObject(FileNumber, v)
        dec = -12345.67@
        v = dec : FilePutObject(FileNumber, v)
        v = DBNull.Value : FilePutObject(FileNumber, v)

        FileClose(FileNumber)

        Console.WriteLine("NOTE: file compare results listed further down in the results file" )
        Console.WriteLine("*** End VB6 Put compatibility test")
    End Sub



    Sub VB6CompatGetTest()
        'Data types to test
        Dim sh As Short
        Dim i As Integer
        Dim byt As Byte
        Dim l As Long
        Dim dt As Date
        Dim s As String
        Dim b As Boolean
        Dim dbl As Double
        Dim sng As Single
        Dim v As Object
        Dim dec As Decimal
        Dim FileNumber As Integer = FreeFile

        Console.WriteLine("*** VB6 Get compatibility test")

        Try
            FileOpen(FileNumber, m_sTestRootPath & "GetPutTest-vb6.out", OpenMode.Random, OpenAccess.Read)

            Console.Write("Short: ")
            FileGet(FileNumber, sh)
            PassFail(sh = 1234)

            Console.Write("Integer: ")
            FileGet(FileNumber, i)
            PassFail(i = 56789)

            Console.Write("Date: ")
            FileGet(FileNumber, dt)
            PassFail(dt = #6/28/1999#)

            Console.Write("Byte: ")
            FileGet(FileNumber, byt)
            PassFail(byt = 129)

            Console.Write("Double: ")
            FileGet(FileNumber, dbl)
            PassFail(dbl = 1234567.7654321)

            Console.Write("Single: ")
            FileGet(FileNumber, sng)
            PassFail(sng = CSng(1234.567))

            Console.Write("Boolean: ")
            FileGet(FileNumber, b)
            PassFail(b = True)

            Console.Write("String: ")
            FileGet(FileNumber, s)
            PassFail(s = "A 16 char string")

            'VB6 cannot read/write a decimal except inside a variant

            Console.Write("VARIANT(Short): ")
            FileGetObject(FileNumber, v)
            PassFail(VarType(v) = VariantType.Short AndAlso CInt(v) = 1234)

            Console.Write("VARIANT(Integer): ")
            FileGetObject(FileNumber, v)
            PassFail(VarType(v) = VariantType.Integer AndAlso CLng(v) = 56789)

            Console.Write("VARIANT(Date): ")
            FileGetObject(FileNumber, v)
            PassFail(VarType(v) = VariantType.Date AndAlso CDate(v) = #6/28/1999#)

            Console.Write("VARIANT(Byte): ")
            FileGetObject(FileNumber, v)
            PassFail(VarType(v) = VariantType.Byte AndAlso CByte(v) = 129)

            Console.Write("VARIANT(Double): ")
            FileGetObject(FileNumber, v)
            PassFail(VarType(v) = VariantType.Double AndAlso CDbl(v) = 1234567.7654321)

            Console.Write("VARIANT(Single): ")
            FileGetObject(FileNumber, v)
            PassFail(VarType(v) = VariantType.Single AndAlso CSng(v) = CSng(1234.567))

            Console.Write("VARIANT(Boolean): ")
            FileGetObject(FileNumber, v)
            PassFail(VarType(v) = VariantType.Boolean AndAlso CBool(v) = True)

            Console.Write("VARIANT(String): ")
            FileGetObject(FileNumber, v)
            PassFail(VarType(v) = VariantType.String AndAlso CStr(v) = "A 16 char string")

            Console.Write("VARIANT(Decimal) positive: ")
            FileGetObject(FileNumber, v)
            PassFail(VarType(v) = VariantType.Decimal AndAlso CDec(v) = 12345.67@)

            Console.Write("VARIANT(Decimal) negative: ")
            FileGetObject(FileNumber, v)
            PassFail(VarType(v) = VariantType.Decimal AndAlso CDec(v) = -12345.67@)

            Console.Write("VARIANT(DBNull): ")
            FileGetObject(FileNumber, v)
            PassFail(VarType(v) = VariantType.Null)
       
            FileClose(FileNumber)
        Catch ex As Exception
            Console.WriteLine()
            Console.WriteLine(ex.GetType().Name & ": " & ex.Message)
            Failed()
        End Try

        Console.WriteLine("*** End VB6 Get compatibility test")
    End Sub



    Private Structure foo
        Public c() As Char
        Public objDummy As Object
    End Structure



    Private Structure Person
        Dim ID As Short
        Dim MonthlySalary As Double
        Dim LastReviewDate As Date
        <VBFixedString(10)> Dim FirstName As String
        <VBFixedString(10)> Dim LastName As String
        <VBFixedString(10)> Dim Title As String
        <VBFixedString(10)> Dim ReviewComments As String
    End Structure



    Sub TestPutGetDynamicFixed()
        'Test Put and Get with all combinations of:
        '   dynamic and fixed arrays, 
        '   variable and fixed length strings
        '   random and binary open mode
        Dim s(3) As String
        Dim t(3) As String
        Dim FileNumber As Integer

        s(0) = "aaaaaaaaaaa"
        s(1) = "bbbbbbbbbbb"
        s(2) = "ccccccccccc"
        s(3) = "ddddddddddd"
        
        Try
            Console.WriteLine("*** VB7 Begin Dynamic and Fixed Test")
            m_sFileName = m_sTestTempPath & "dynfixtest.out"
            FileNumber = FreeFile

            Try
                Kill(m_sFileName)
            Catch ex As Exception
            End Try

            Console.WriteLine("Random mode: ")

            Console.WriteLine("Put of fixed array, fixed strings: ")
            FileOpen(FileNumber, m_sFileName, OpenMode.Random)
            FilePut(FileNumber, s, , False, True)
            PassFail(Lof(FileNumber) = 11*4)
            t(0) = "12345678901" 
            FileGet(FileNumber, t, 1, False, True)
            PassFail(t(0) = s(0) And t(1) = s(1) And t(2) = s(2) And t(3) = s(3))
            FileClose(FileNumber)
            Kill(m_sFileName)

            Console.WriteLine("Put of fixed array, variable strings: ")
            FileOpen(FileNumber, m_sFileName, OpenMode.Random)
            FilePut(FileNumber, s, , False, False)
            PassFail(Lof(FileNumber) = (11+2)*4)
            t(0) = "12345678901" 
            FileGet(FileNumber, t, 1, False, False)
            PassFail(t(0) = s(0) And t(1) = s(1) And t(2) = s(2) And t(3) = s(3))
            FileClose(FileNumber)
            Kill(m_sFileName)

            Console.WriteLine("Put of dynamic array, fixed strings: ")
            FileOpen(FileNumber, m_sFileName, OpenMode.Random)
            FilePut(FileNumber, s, , True, True)
            PassFail(Lof(FileNumber) = 10+(11*4))
            t(0) = "12345678901" 
            FileGet(FileNumber, t, 1, True, True)
            PassFail(t(0) = s(0) And t(1) = s(1) And t(2) = s(2) And t(3) = s(3))
            FileClose(FileNumber)
            Kill(m_sFileName)

            Console.WriteLine("Put of dynamic array, variable strings: ")
            FileOpen(FileNumber, m_sFileName, OpenMode.Random)
            FilePut(FileNumber, s, , True, False)
            PassFail(Lof(FileNumber) = 10+((11+2)*4))
            t(0) = "12345678901" 
            FileGet(FileNumber, t, 1, True, False)
            PassFail(t(0) = s(0) And t(1) = s(1) And t(2) = s(2) And t(3) = s(3))
            FileClose(FileNumber)
            Kill(m_sFileName)

            Console.WriteLine("Binary mode: ")

            Console.WriteLine("Put of fixed array, fixed strings: ")
            FileOpen(FileNumber, m_sFileName, OpenMode.Binary)
            FilePut(FileNumber, s, , False, True)
            PassFail(Lof(FileNumber) = 11*4)
            t(0) = "12345678901" 
            FileGet(FileNumber, t, 1, False, True)
            PassFail(t(0) = s(0) And t(1) = s(1) And t(2) = s(2) And t(3) = s(3))
            FileClose(FileNumber)
            Kill(m_sFileName)

            Console.WriteLine("Put of fixed array, variable strings: ")
            FileOpen(FileNumber, m_sFileName, OpenMode.Binary)
            FilePut(FileNumber, s, , False, False)
            PassFail(Lof(FileNumber) = (11+2)*4)
            t(0) = "12345678901" 
            FileGet(FileNumber, t, 1, False, False)
            PassFail(t(0) = s(0) And t(1) = s(1) And t(2) = s(2) And t(3) = s(3))
            FileClose(FileNumber)
            Kill(m_sFileName)

            Console.WriteLine("Put of dynamic array, fixed strings: ")
            FileOpen(FileNumber, m_sFileName, OpenMode.Binary)
            FilePut(FileNumber, s, , True, True)
            PassFail(Lof(FileNumber) = 10+(11*4))
            t(0) = "12345678901" 
            FileGet(FileNumber, t, 1, True, True)
            PassFail(t(0) = s(0) And t(1) = s(1) And t(2) = s(2) And t(3) = s(3))
            FileClose(FileNumber)
            Kill(m_sFileName)

            Console.WriteLine("Put of dynamic array, variable strings: ")
            FileOpen(FileNumber, m_sFileName, OpenMode.Binary)
            FilePut(FileNumber, s, , True, False)
            PassFail(Lof(FileNumber) = 10+((11+2)*4))
            t(0) = "12345678901" 
            FileGet(FileNumber, t, 1, True, False)
            PassFail(t(0) = s(0) And t(1) = s(1) And t(2) = s(2) And t(3) = s(3))
            FileClose(FileNumber)
            Kill(m_sFileName)
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine("*** VB7 End Dynamic and Fixed Test")
    End Sub



    Sub TestPut()
        'Data types to test
        Dim i As Integer
        Dim byt As Byte
        Dim l As Long
        Dim dt As Date
        Dim s, s2 As String
        Dim b As Boolean
        Dim dbl As Double
        Dim sng As Single
        Dim v As Object
        Dim dec, negdec As Decimal
        Dim FileNumber As Integer
        Dim a(2) As object
        Dim f As New foo()
        Dim cc(2) As Char
        Dim dd(2) As Decimal
        Dim g(1) As Person
        Dim rec As Integer
        Dim bb(109) As Byte
        Dim j As Integer
        Const reclen As Integer = 128 ' default size in VB6

        i = 1234
        l = 56789
        dt = #6/28/1999#
        byt = 129
        s = "A 16 char string"
        s2 = "ABCDEF"
        dbl = 1234567.7654321
        b = True
        sng = 1234.567
        dec = 12345.67@
        negdec = -12345.67@
        a(0) = CShort(1)
        a(1) = CShort(2)
        a(2) = CShort(3)
        f.c = CType(StrDup(2, "q"), Char())
        f.objDummy = "World"
        cc(0) = "a"c
        cc(1) = "b"c
        cc(2) = "c"c
        dd(0) = 1D
        dd(1) = 2D
        dd(2) = 3D

        g(0).ID = 5
        g(0).MonthlySalary = 55.55
        g(0).LastReviewDate = #4/1/2001#
        g(0).FirstName = "Todd"
        g(0).LastName = "Apley"
        g(0).Title = "nice guy"
        g(0).ReviewComments = "Wow"

        For j = 0 To 109
            bb(j) = j
        Next

        Console.WriteLine("*** VB7 Begin Test Put")

        Try
            m_sFileName = m_sTestTempPath & "getputtest.out"
            m_sFileNameDynamic = m_sTestTempPath & "getputtest2.out"

            Try
                Kill(m_sFileName)
            Catch ex As Exception
            End Try

            FileNumber = FreeFile

            Console.WriteLine("Open: ")
            FileOpen(FileNumber, m_sFileName, OpenMode.Random)
            rec = 1
            PassFail(0 = Lof(FileNumber))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)
            'FAILED PassFail(False = Eof(FileNumber))

            Console.WriteLine("Put Integer: ")
            FilePut(FileNumber, i)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)
            PassFail(Len(i) = Lof(FileNumber))
            'FAILED PassFail(False = Eof(FileNumber))

            Console.WriteLine("Put Long: ")
            FilePut(FileNumber, l)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)
            PassFail((reclen + 8) = Lof(FileNumber))
            'FAILED PassFail(False = Eof(FileNumber))

            Console.WriteLine("Put Date: ")
            FilePut(FileNumber, dt)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Byte: ")
            FilePut(FileNumber, byt)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Double: ")
            FilePut(FileNumber, dbl)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Single: ")
            FilePut(FileNumber, sng)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Boolean: ")
            FilePut(FileNumber, b)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put String: ")
            FilePut(FileNumber, s)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Decimal (positive): ")
            FilePut(FileNumber, dec)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Decimal (negative): ")
            FilePut(FileNumber, negdec)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Structure of char array and object: ")
            FilePut(FileNumber, f)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put array of structures with fixed length strings: ")
            FilePut(FileNumber, g)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            'Console.WriteLine("Put char array: ")  'Waiting on RAID 235736, postponed
            'FilePut(FileNumber, cc)
            'rec += 1
            'PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            'PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put decimal array: ")  
            FilePut(FileNumber, dd)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Object(Integer): ")
            v = i : FilePutObject(FileNumber, v)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Object(Long): ")
            v = l : FilePutObject(FileNumber, v)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Object(Date): ")
            v = dt : FilePutObject(FileNumber, v)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Object(Byte): ")
            v = byt : FilePutObject(FileNumber, v)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Object(Double): ")
            v = dbl : FilePutObject(FileNumber, v)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Object(Single): ")
            v = sng : FilePutObject(FileNumber, v)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Object(Byte): ")
            v = b : FilePutObject(FileNumber, v)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Object(String): ")
            v = s : FilePutObject(FileNumber, v)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Object(String2): ")
            v = s2 : FilePutObject(FileNumber, v)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Object(Decimal) (positive): ")
            v = dec : FilePutObject(FileNumber, v)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Object(Decimal) (negative): ")
            v = negdec : FilePutObject(FileNumber, v)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Object(DBNull): ")
            v = DBNull.Value : FilePutObject(FileNumber, v)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            Console.WriteLine("Put Object(Object Array): ")
            v = a : FilePutObject(FileNumber, v)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)
            'FAILED PassFail((((rec - 2) * reclen) + 8) = Lof(FileNumber))
            'FAILED PassFail(False = Eof(FileNumber))

            FileClose(FileNumber)

            Try
                Kill(m_sFileNameDynamic)
            Catch ex As Exception
            End Try

            FileNumber = FreeFile

            Console.WriteLine("Open Record size 120: ")
            FileOpen(FileNumber, m_sFileNameDynamic, OpenMode.Random, , , 120)
            rec = 1

            Console.WriteLine("Put dynamic byte array: ")  
            FilePut(FileNumber, bb, , True)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(rec = RecordFromPosition(Lof(FileNumber), reclen) + 1)

            FileClose(FileNumber)
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine("*** VB7 End Test Put")
    End Sub



    Sub TestGet()
        'Data types to test
        Dim i As Integer
        Dim s As String
        Dim byt As Byte
        Dim l As Long
        Dim dt As Date
        Dim b As Boolean
        Dim dbl As Double
        Dim sng As Single
        Dim v As Object
        Dim dec, negdec As Decimal
        Dim FileNumber As Integer
        Dim a(3) as object
        Dim f As New foo()
        Dim cc(2) As Char
        Dim dd(2) As Decimal
        Dim g(1) As Person
        Dim bb(109) As Byte 
        Dim uninitialized() As Byte
        Dim initialized() As Byte = {}
        Dim rec As Integer
        Dim recEof As Integer
        Const reclen As Integer = 128 'Default record length

        f.c = CType(StrDup(2, "a"), Char()) 'Must be initialized to avoid Bad record length errors

        Console.WriteLine("*** VB7 Begin Test Get")

        Try
            Console.WriteLine("Open: ")
            FileNumber = FreeFile
            FileOpen(FileNumber, m_sFileName, OpenMode.Random)
            rec = 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(False = Eof(FileNumber))

            Console.WriteLine("Integer: ")
            FileGet(FileNumber, i)
            rec += 1
            PassFail(i = 1234)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("Long: ")
            FileGet(FileNumber, l)
            rec += 1
            PassFail(l = 56789)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            
            Console.WriteLine("Date: ")
            FileGet(FileNumber, dt)
            rec += 1
            PassFail(dt = #6/28/1999#)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("Byte: ")
            FileGet(FileNumber, byt)
            rec += 1
            PassFail(byt = 129)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("Double: ")
            FileGet(FileNumber, dbl)
            rec += 1
            PassFail(dbl = 1234567.7654321)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("Single: ")
            FileGet(FileNumber, sng)
            rec += 1
            PassFail(sng = CSng(1234.567))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("Boolean: ")
            FileGet(FileNumber, b)
            rec += 1
            PassFail(b = True)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("String: ")
            FileGet(FileNumber, s)
            rec += 1
            PassFail(s = "A 16 char string")
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("Decimal (positive): ")
            FileGet(FileNumber, dec)
            rec += 1
            PassFail(dec = 12345.67@)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("Decimal (negative): ")
            FileGet(FileNumber, dec)
            rec += 1
            PassFail(dec = -12345.67@)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("Structure of char array and object: ")
            FileGet(FileNumber, f)
            rec += 1
            PassFail(f.c(0) = "q"c And f.c(1) = "q"c And f.objDummy = "World")
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("Array of structures with fixed length strings: ")
            FileGet(FileNumber, g)
            rec += 1
            PassFail(g(0).ID = 5 And g(0).MonthlySalary = 55.55 And g(0).LastReviewDate = #4/1/2001# And _
                g(0).FirstName = "Todd      " And g(0).LastName = "Apley     " And g(0).Title = "nice guy  " And _
                g(0).ReviewComments = "Wow       ")
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            'Console.WriteLine("Char array: ")   'Waiting on RAID 235736, postponed
            'FileGet(FileNumber, cc)
            'rec += 1
            'PassFail(cc(0) = "a"c And cc(1) = "b"c And cc(2) = "c"c)
            'PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("Decimal array: ")   
            FileGet(FileNumber, dd)
            rec += 1
            PassFail(dd(0) = 1D And dd(1) = 2D And dd(2) = 3D)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("VARIANT(Integer): ")
            FileGetObject(FileNumber, v)
            rec += 1
            PassFail((VarType(v) = VariantType.Integer) AndAlso (CInt(v) = 1234))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("VARIANT(Long): ")
            FileGetObject(FileNumber, v)
            rec += 1
            PassFail((VarType(v) = VariantType.Long) AndAlso (CLng(v) = 56789))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("VARIANT(Date): ")
            FileGetObject(FileNumber, v)
            rec += 1
            PassFail((VarType(v) = VariantType.Date) AndAlso (CDate(v) = #6/28/1999#))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("VARIANT(Byte): ")
            FileGetObject(FileNumber, v)
            rec += 1
            PassFail((VarType(v) = VariantType.Byte) AndAlso (CByte(v) = 129))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("VARIANT(Double): ")
            FileGetObject(FileNumber, v)
            rec += 1
            PassFail((VarType(v) = VariantType.Double) AndAlso (CDbl(v) = 1234567.7654321))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("VARIANT(Single): ")
            FileGetObject(FileNumber, v)
            rec += 1
            PassFail((VarType(v) = VariantType.Single) AndAlso (CSng(v) = CSng(1234.567)))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("VARIANT(Boolean): ")
            FileGetObject(FileNumber, v)
            rec += 1
            PassFail((VarType(v) = VariantType.Boolean) AndAlso (CBool(v) = True))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("VARIANT(String) into object: ")
            FileGetObject(FileNumber, v)
            rec += 1
            PassFail((VarType(v) = VariantType.String) AndAlso (CStr(v) = "A 16 char string"))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("VARIANT(String) into uninitialized string: ") 'RAID 233851
            v = ""
            FileGetObject(FileNumber, v)
            rec += 1
            PassFail(VarType(v) = VariantType.String AndAlso CStr(v) = "ABCDEF")
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("VARIANT(Decimal) positive: ")
            FileGetObject(FileNumber, v)
            rec += 1
            PassFail((VarType(v) = VariantType.Decimal) AndAlso (CDec(v) = 12345.67@))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("VARIANT(Decimal) negative: ")
            FileGetObject(FileNumber, v)
            rec += 1
            PassFail((VarType(v) = VariantType.Decimal) AndAlso (CDec(v) = -12345.67@))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(False = Eof(FileNumber))

            Console.WriteLine("VARIANT(DBNull): ")
            FileGetObject(FileNumber, v)
            rec += 1
            PassFail((VarType(v) = VariantType.Null))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(False = Eof(FileNumber))

            Console.WriteLine("VARIANT(Object Array): ")
            FileGetObject(FileNumber, v)
            rec += 1
            PassFail((VarType(v) = (VariantType.Array Or VariantType.Object)) AndAlso _
                (CShort(v(0)) = 1) AndAlso (CShort(v(1)) = 2) AndAlso (CShort(v(2)) = 3))
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))

            Console.WriteLine("EOF: ")
            recEof = rec
            PassFail(True = Eof(FileNumber))
            PassFail(recEof = RecordFromPosition(Lof(FileNumber), reclen) + 1)	

            Console.WriteLine("Seek 5: ")
            rec = 5
            Seek(FileNumber, rec)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(False = Eof(FileNumber))	

            Console.WriteLine("Seek 1: ")
            rec = 1
            Seek(FileNumber, rec)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(False = Eof(FileNumber))	

            Console.WriteLine("Seek 7: ")
            rec = 7
            Seek(FileNumber, rec)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(False = Eof(FileNumber))	

            Console.WriteLine("Seek(Last Record): ")
            rec = recEof - 1
            Seek(FileNumber, rec)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(False = Eof(FileNumber))	

            Console.WriteLine("Seek(Past Last Record): ")
            rec = recEof
            Seek(FileNumber, rec)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            'FAILED PassFail(False = Eof(FileNumber))	

            Console.WriteLine("Seek(Further Past Last Record): ")
            rec += 10
            Seek(FileNumber, rec)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(recEof = RecordFromPosition(Lof(FileNumber), reclen) + 1)
            'FAILED PassFail(False = Eof(FileNumber))	

            Console.WriteLine("Get(Past Last Record): ")
            FileGet(FileNumber, s)
            rec += 1
            'FAILED PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail(recEof = RecordFromPosition(Lof(FileNumber), reclen) + 1)
            PassFail(True = Eof(FileNumber))	

            FileClose(FileNumber)

            Console.WriteLine("Test Dynamic arrays: ")
            FileNumber = FreeFile

            Console.WriteLine("Open Record size 120: ")
            FileOpen(FileNumber, m_sFileNameDynamic, OpenMode.Random, , , 120)
            rec = 1

            Console.WriteLine("Get dynamic byte array: ")  
            FileGet(FileNumber, bb, , True)
            rec += 1
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail( bb(0) = 0 AndAlso bb(109) = 109)

            Console.WriteLine("Get initialized dynamic byte array: ")  
            FileGet(FileNumber, initialized, 1, True)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            PassFail( initialized(0) = 0 AndAlso initialized(109) = 109)

            Try
                Console.WriteLine("Get uninitialized dynamic byte array: ")  
                FileGet(FileNumber, uninitialized, 1, True)
                Failed
            Catch ex As ArgumentException
                Passed
            Catch ex As Exception
                Failed(ex)
            End Try

            PassFail( uninitialized Is Nothing)

            Console.WriteLine("Get dynamic byte array past EOF: ")  
            FileGet(FileNumber, bb, , True)
            PassFail(((rec - 1) = Loc(FileNumber)) AndAlso (rec = Seek(FileNumber)))
            FileClose(FileNumber)

            m_sFileNameLenCheck = m_sTestTempPath & "getputtest3.out"

            Try
                Kill(m_sFileNameLenCheck)
            Catch ex As Exception
            End Try

            FileNumber = FreeFile

            Console.WriteLine("Open for Length Check: ")
            FileOpen(FileNumber, m_sFileNameLenCheck, OpenMode.Output)
            Console.WriteLine("Print string with no descriptor: ")  
            Print(FileNumber, StrDup(60, "X"))

            Try
                FileGet(FileNumber, s, 1)
            Catch ex As IO.IOException
                Passed
            Catch ex As Exception
                Failed(ex)
            End Try

            PassFail( s.Length = 0)

            FileClose(FileNumber)
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine("*** VB7 End Test Get")
    End Sub
    


    Function RecordFromPosition(ByVal Position As Integer, ByVal RecordLength As Integer) As Integer
        RecordFromPosition = (Position + RecordLength - 1) \ RecordLength
    End Function



    Sub Failed(ByVal ex As Exception)
        If ex Is Nothing Then
            Console.WriteLine("NULL System.Exception")
        Else
            Console.WriteLine(ex.GetType().FullName & " " & ex.Message)
        End If
        Console.WriteLine("FAILED !!!")
    End Sub



    Sub Failed()
        Console.WriteLine("FAILED !!!")
    End Sub



    Sub Passed()
        Console.WriteLine("passed")
    End Sub



    Sub PassFail(ByVal Actual As String, ByVal Expected As String)
        If Actual = Expected Then
            Console.WriteLine("passed")
        Else
            Console.WriteLine("FAILED: expected >" & Expected &"<, but got >" & Actual & "<" )
        End If
    End Sub



    Sub PassFail(ByVal bPassed As Boolean)
        If bPassed Then
            Console.WriteLine("passed")
        Else
            Console.WriteLine("FAILED")
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



End Module


Module Bug170856

    Function Function1(ByRef arg)
        Function1 = Arg
    End Function



    Sub Main()
        Console.Write("*** Bug 170856: ")

        Try
            Dim Stringvar1 As Object = "XXXXX"
            Dim RandomFile As String = Test.m_sTestTempPath & "Bug170856.DAT"
            Dim DummyArg As Integer = FreeFile

            Try
                Kill(RandomFile)
            Catch
            End Try

            FileOpen(DummyArg, RandomFile, OpenMode.Random)
            FilePutObject(Function1(DummyArg), Stringvar1) 
            FileClose(DummyArg)

            Stringvar1 = "ABCDEFG"

            FileOpen(DummyArg, RandomFile, OpenMode.Random)
            FileGetObject(Function1(DummyArg), Stringvar1) 
            PassFail(StringVar1 = "XXXXX")
            FileClose(DummyArg)
        Catch ex as Exception
            Failed(ex)
        End Try

    End Sub

End Module


Module Bug280613

    Sub Test()

        Console.Write("*** Bug 280613: ")
        Dim I As object
        Dim FileNumber As Integer = FreeFile
        Try
            FileOpen(FileNumber, "TESTFILE.txt", OpenMode.Output)
            PrintLine(FileNumber, 14&)
            FileClose(FileNumber)

            FileOpen(FileNumber, "TESTFILE.txt", OpenMode.Input)
            Input(FileNumber, I)
            FileClose(FileNumber)
            PassFail(VarType(I) = VariantType.Short)
        Catch ex as Exception
            Failed(ex)
        End Try

    End Sub


End Module

Module Bug289099

    Sub Test()

        Console.Write("*** Bug 289099: ")
        Dim I As object
        Dim FileNumber As Integer = FreeFile
        Dim strvar As String
        Dim fn As String = "289099.dat"

        Try
            FileOpen(FileNumber, fn, OpenMode.Binary, OpenAccess.ReadWrite)
            strvar = "Check it out"
            FilePut(FileNumber, strvar)
            FileClose(FileNumber)

            strvar = New String(" "c, 12)
            FileOpen(FileNumber, fn, OpenMode.Binary, OpenAccess.Read)
            FileGet(FileNumber, strvar)
            FileClose(FileNumber)
            PassFail(strvar = "Check it out")

        Catch ex as Exception
            Failed(ex)
        End Try

    End Sub


End Module


Module Bug298737

    Private Structure FixedArray1
        <VBFixedArray(1, 2)> Dim m_dwReserved1(,) As Short
    End Structure

    Private Structure FixedArray2
        <VBFixedArray(1, 2), VBFixedString(5)> Dim m_dwReserved2(,) As String
    End Structure

    Private Structure FixedArray3
        <VBFixedArray(1), VBFixedString(5)> Dim m_dwReserved1() As String
        <VBFixedArray(1, 2), VBFixedString(5)> Dim m_dwReserved2(,) As String
    End Structure

    Sub Test()
        Dim f1 As FixedArray1
        Dim f2 As FixedArray2
        Dim f3 As FixedArray3

        Console.WriteLine("*** Bug 298737: ")
        Try
            Console.Write("   1) ")
            PassFail(Len(f1) = 12)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   2) ")
            PassFail(Len(f2) = 30)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   3) ")
            PassFail(Len(f3) = 40)
        Catch ex As Exception
            Failed(ex)
        End Try

        Dim FileName As String = m_sTestRootPath & "bug298737.dat"
        Dim FileNumber As Integer = FreeFile

        Try
            Console.Write("   4) ")

            Try
                Kill(FileName & "1")
            Catch ex As Exception
            End Try

            FileOpen(FileNumber, FileName & "1", OpenMode.Binary, OpenAccess.ReadWrite)
            FilePut(FileNumber, f1)
            PassFail(LOF(FileNumber) = 12)
            FileClose(FileNumber)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   5) ")
            f2.m_dwReserved2 = New String(1, 2) {}
            f2.m_dwReserved2(0, 0) = "a"
            f2.m_dwReserved2(1, 0) = "b"
            f2.m_dwReserved2(0, 1) = "c"
            f2.m_dwReserved2(1, 1) = "d"
            f2.m_dwReserved2(0, 2) = "e"
            f2.m_dwReserved2(1, 2) = "f"

            Try
                Kill(FileName & "2")
            Catch ex As Exception
            End Try

            FileOpen(FileNumber, FileName & "2", OpenMode.Binary, OpenAccess.ReadWrite)
            FilePut(FileNumber, f2)
            PassFail(LOF(FileNumber) = 30)
            FileClose(FileNumber)

        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   6) ")
            f3.m_dwReserved1 = New String(1) {"abc", "def"}

            Try
                Kill(FileName & "3")
            Catch ex As Exception
            End Try

            FileOpen(FileNumber, FileName & "3", OpenMode.Binary, OpenAccess.ReadWrite)
            FilePut(FileNumber, f3)
            PassFail(LOF(FileNumber) = 40)
            FileClose(FileNumber)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module


'Test bad array dimensions
Module Bug300062
    Private Structure struct1
        Public arr1(,,) As Integer
        <VBFixedString(4)> Public arr2 As String
    End Structure

    Private Structure struct2
        <VBFixedArray(1,2)> Public arr1(,) As Integer
    End Structure

    Sub Test()
        Console.WriteLine("*** Bug300062")

        Dim structVar1 As struct1
        Dim structVar2 as struct2
        Dim i, j, k, FileNumber As Integer
        Dim FileName As String

        ReDim structVar1.arr1(6, 3, 4)
        structVar1.arr2 = "jhagjangkjngakd"

        For i = 0 To structVar1.arr1.GetUpperBound(0)
            For j = 0 To structVar1.arr1.GetUpperBound(1)
                For k = 0 To structVar1.arr1.GetUpperBound(2)
                    structVar1.arr1(i, j, k) = 9
                Next
            Next
        Next


        FileNumber = FreeFile
        FileName = m_sTestTempPath & "bug300062a.dat"

        Try
            Console.Write("   1) ")
            FileOpen(FileNumber, FileName, OpenMode.Binary, OpenAccess.Write)
            FilePut(FileNumber, structVar1)
        Catch ex As ArgumentException
            Passed()
            Console.WriteLine("      " & ex.Message)
        Catch ex As Exception
            Failed(ex)
        Finally
            FileClose(FileNumber)
        End Try


        ReDim structVar2.arr1(2,3)

        Dim counter As Integer = 0
        For j = 0 To structVar2.arr1.GetUpperBound(1)
            For i = 0 To structVar2.arr1.GetUpperBound(0)
                structVar2.arr1(i, j) = counter
                counter += 1
            Next
        Next

        Try
            Console.Write("   2) ")
            FileName = m_sTestTempPath & "bug300062b.dat"
            FileOpen(FileNumber, FileName, OpenMode.Binary, OpenAccess.Write)
            FilePut(FileNumber, structVar2)
        Catch ex As ArgumentException
            Passed()
            Console.WriteLine("      " & ex.Message)
        Catch ex As Exception
            Failed(ex)
        Finally
            FileClose(FileNumber)
        End Try

    End Sub
End Module


Module Bug300161
    Private Structure struct1
        <VBFixedArray(4)> Public arr1() As Integer
    End Structure

    Sub Test()
        Console.WriteLine("*** Bug300161")

        Dim structVar As struct1
        Dim result As struct1
        Dim i  As Integer
        Dim FileNumber As Integer
        Dim FileName As String

        ReDim structVar.arr1(2)

        For i = 0 To 2
           structVar.arr1(i) = 9
        Next

        FileName = m_sTestTempPath & "bug300161.dat"
        FileNumber = FreeFile

        Try
            Kill(FileName)
        Catch
        End Try

        Try
            Console.Write("   1) ")
            FileOpen(FileNumber, FileName, OpenMode.Binary, OpenAccess.Write)
            FilePut(FileNumber, structVar)
            PassFail(Lof(FileNumber) = 20)
        Catch ex As Exception
            Failed(ex)
        Finally
            FileClose(FileNumber)
        End Try

        Try
            Console.Write("   2) ")
            FileOpen(FileNumber, FileName, OpenMode.Binary, OpenAccess.Read)
            FileGet(FileNumber, result)
            PassFail(Len(result) = 20)
        Catch ex As Exception
            Failed(ex)
        Finally
            FileClose(FileNumber)
        End Try

    End Sub
End Module

Module Bug340556
    Sub Test()
        Console.WriteLine("*** Bug340556")

        Dim FileName As String = m_sTestRootPath & "bug340556.dat"
        Dim FileNumber As Integer = 2
        FileOpen(FileNumber, FileName, OpenMode.Input)

        Console.WriteLine(LineInput(FileNumber))
        Console.WriteLine(LineInput(FileNumber))

        Seek(FileNumber, 1)
        Console.WriteLine(LineInput(FileNumber))
        FileClose(FileNumber)
    End Sub
End Module

Module Bug550463
    Sub Test
        Console.WriteLine("*** Bug550463")

        Dim FileName As String = m_sTestRootPath & "bug550463.dat"
        Dim FileNumber As Integer = FreeFile()
        FileOpen(FileNumber, FileName, OpenMode.Output)

        Dim outputstring As String = "Double quotes are ""trouble"" aren't they?"
        Write(FileNumber, outputstring)

        FileClose(FileNumber)
        FileOpen(FileNumber, FileName, OpenMode.Input)

        Dim inputstring As String
        Input(FileNumber, inputstring)

        Console.Write("1): ") : Console.WriteLine(inputstring)

        FileClose(FileNumber)
    End Sub
End Module

