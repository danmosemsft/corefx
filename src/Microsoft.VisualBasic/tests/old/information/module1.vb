Option Compare Binary

Imports System.Globalization
Imports System
Imports System.Text
Imports System.Threading
Imports Microsoft.VisualBasic

Module Test

    Structure MyType
        Public s As String
    End Structure

    Enum MyColors
        Red
        Blue
        Green
    End Enum

    Dim shrt As Short
    Dim i As Integer
    Dim l As Long
    Dim sng As Single
    Dim dbl As Double
    Dim dec As Decimal
    Dim dt As Date
    Dim dtNull As Date
    Dim s As String
    Dim obj As Object
    Dim sb As StringBuilder
    Dim b As Boolean
    Dim byt As Byte
    Dim udt As MyType
    Dim ch As Char
    Dim ai() As Integer
    Dim m_lBugID As Long
    Dim objDBNull As Object



    Sub Main()
        'Don't run NamedParamTest, just ensure compile
        'NamedParamTest

        Console.WriteLine("Begin Tests")

        InitTest
        ErrObjectTest
        IsArrayTest
        IsDateTest
        IsErrorTest
        IsNumericTest
        QBColorTest
        RGBTest
        TypeNameTest
        VarTypeTest

        Bug247905.Test

        Console.WriteLine("End Tests")
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
            Console.WriteLine( "fail (unexpected exception)" )
        End If
        m_lBugID = lID
        Console.WriteLine( "Bug" & CStr(m_lBugID) & ": " )
    End Sub



    Function PassFail(ByVal bPassed As Boolean) As String
        If bPassed Then
            Console.WriteLine( "passed" )
        Else
            Console.WriteLine( "FAILED !!!" )
        End If
    End Function



    Sub NamedParamTest()
#if false
        ' NOTE:  ***** DO NOT MODIFY THIS FUNCTION *****
        '       (unless you really know what you're doing)
        '
        ' This function is intended to just to ensure compilation
        ' If someone changes a name in the runtime, this should break
        ' we will rely on the test team to make sure named params
        ' return results as expected
        '
        Dim v As Variant
        Dim eo As IErrObject

        v = VBA.Err()   'No params
        eo.Raise Number := v, Source := v, Description := v, HelpFile := v, HelpContext := v
        v = VBA.IMEStatus() 'No params
        v = VBA.IsArray(VarName := v)
        v = VBA.IsDate(Expression := v)
        v = VBA.IsEmpty(Expression := v)
        v = VBA.IsError(Expression := v)
        v = VBA.IsMissing(ArgName := v)
        v = VBA.IsNull(Expression := v)
        v = VBA.IsNumeric(Expression := v)
        v = VBA.IsObject(Expression := v)
        v = VBA.QBColor(Color := 1)
        v = VBA.RGB(Red := 1, Green := 1, Blue := 1)
        v = VBA.TypeName(VarName := v)
        v = VBA.VarType(VarName := v)
#end if
    End Sub



    Sub InitTest()
        obj = Nothing
        sb = New StringBuilder
        dt = #1/1/1999#
        dtNull = #12:00:00 AM#
        s = "A string"
    End Sub



    Sub ErrObjectTest()
         1:  Dim a As Double
         2:  Dim b, c As Integer

         8:  On Error Resume Next
         9:  Console.WriteLine( "*** Err object test" )
        10:  FileOpen(1, "nonexistantfile", OpenMode.Input )
        11:  Console.WriteLine( Err.Number )
        12:  Console.WriteLine( Err.Erl )
        '13:  Console.WriteLine( Err.Description )

        19:  Err.Clear
        20:  Err.Raise( 5 )
        21:  Console.WriteLine( Err.Number )
        22:  Console.WriteLine( Err.Erl )
        23:  Console.WriteLine(Err.Description)
        24:  Console.WriteLine()
    End Sub



    Sub IsErrorTest()
        Dim e As New System.Exception

        Console.WriteLine( "*** IsError test" )
        Console.WriteLine( "1: " & IsError("") )
        Console.WriteLine( "2: " & IsError( e ) )
        Console.WriteLine()
    End Sub



    Sub IsArrayTest()
        Dim s(0) As String
        Dim i(0) As Integer
        Dim o(0) As Object

        Console.WriteLine( "*** IsArray test" )

        Try
            Console.Write("Nothing: ")
            PassFail(IsArray(Nothing) = False)

            Console.Write("Object(): ")
            PassFail(IsArray(o) = True)

            Console.Write("Integer(): ")
            PassFail(IsArray(i) = True)

            Console.Write("String(): ")
            PassFail(IsArray(s) = True)

        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine()
    End Sub



    Sub IsDateTest()
        Dim dt As Date
        Dim o(0) As Object
        Dim s As String

        Console.WriteLine("*** IsDate test")

        Try
            Console.Write("Nothing: ")
            PassFail(IsDate(Nothing) = False)

            Console.Write("Object(): ")
            PassFail(IsDate(o) = False)

            Console.Write("String(): ")
            PassFail(IsDate(s) = False)

            Console.WriteLine("Misc strings:")

            Console.Write("1: ") : PassFail(IsDate("1/1/111") = True)
            Console.Write("2: ") : PassFail(IsDate("1/1/1111") = True)
            Console.Write("3: ") : PassFail(IsDate("1/1/9999") = True)
            Console.Write("4: ") : PassFail(IsDate("1/1/10000") = False)
            Console.Write("5: ") : PassFail(IsDate("2/29/00") = True)
            Console.Write("6: ") : PassFail(IsDate("2/29/01") = False)
            Console.Write("7: ") : PassFail(IsDate(" " & vbTab & "        2/29/00           ") = True)
            Console.Write("8: ") : PassFail(IsDate("2/29/01a") = False)
            Console.Write("9: ") : PassFail(IsDate("a2/29/01") = False)
            Console.Write("10: ") : PassFail(IsDate("2-29.00") = True)  ' VSWhidbey 427066: This is a valid format. Log bug against BCL if this fails again.
            Dim OldCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
            Thread.CurrentThread.CurrentCulture = New CultureInfo("de-de")
            Console.Write("11: ") : PassFail(IsDate("2/29/00") = False)
            Console.Write("12: ") : PassFail(IsDate("29/2/00") = True)
            Thread.CurrentThread.CurrentCulture = OldCulture

        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine()
    End Sub

    Sub IsNumericTest()
        Console.WriteLine( "*** IsNumeric test" )

        Try
            Console.Write("'1234': ")
            PassFail(IsNumeric("1234") = True)

            Console.Write("Short: ")
            PassFail(IsNumeric(shrt) = True)

            Console.Write( "Integer: " ) 
            PassFail(IsNumeric(i) = True)

            Console.Write( "Long: " )
            PassFail(IsNumeric(l) = True)

            Console.Write( "Decimal: " ) 
            PassFail(IsNumeric(dec) = True)

            Console.Write( "Single: " ) 
            PassFail(IsNumeric(sng) = True)

            Console.Write( "Double: " ) 
            PassFail(IsNumeric(dbl) = True)

            Console.Write( "Byte: " ) 
            PassFail(IsNumeric(byt) = True)

            Console.Write( "UDT: " ) 
            PassFail(IsNumeric(udt) = False)

#If False Then             'UNDONE: VSWhidbey:28645
            Console.Write( "Char: " ) 
            PassFail(IsNumeric(ch) = False)

            Console.Write( "AB: " ) 
            PassFail(IsNumeric("AB") = False)
#End If

            Console.Write( "&HAB: " ) 
            PassFail(IsNumeric("&HAB") = True)

            Console.Write( "Date: " ) 
            PassFail(IsNumeric(dt) = False)

            Console.Write( "Object(Nothing): " ) 
            PassFail(IsNumeric(obj) = False)

            Console.Write( "Object(StringBuilder): " ) 
            PassFail(IsNumeric(sb) = False)

            Console.Write( "Boolean: " ) 
            PassFail(IsNumeric(False) = True)

            Console.Write( "Integer(): " ) 
            PassFail(IsNumeric(ai) = False)

        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine()
    End Sub



    Sub QBColorTest()
        Console.WriteLine( "*** QBColor test" )

        Try
            Console.Write( "0) " ) 
            PassFail(QBColor(0) = &H000000&)  '   0 - black

            Console.Write( "1) " ) 
            PassFail(QBColor(1) = &H800000&)  '   1 - blue

            Console.Write( "2) " ) 
            PassFail(QBColor(2) = &H008000&)  '   2 - green

            Console.Write( "3) " ) 
            PassFail(QBColor(3) = &H808000&)  '   3 - cyan

            Console.Write( "4) " ) 
            PassFail(QBColor(4) = &H000080&)  '   4 - red

            Console.Write( "5) " ) 
            PassFail(QBColor(5) = &H800080&)  '   5 - magenta

            Console.Write( "6) " ) 
            PassFail(QBColor(6) = &H008080&)  '   6 - yellow

            Console.Write( "7) " ) 
            PassFail(QBColor(7) = &Hc0c0c0&)  '   7 - white

            Console.Write( "8) " ) 
            PassFail(QBColor(8) = &H808080&)  '   8 - gray

            Console.Write( "9) " ) 
            PassFail(QBColor(9) = &Hff0000&)  '   9 - light blue

            Console.Write( "10) " ) 
            PassFail(QBColor(10) = &H00ff00&) '  10 - light green

            Console.Write( "11) " ) 
            PassFail(QBColor(11) = &Hffff00&) '  11 - light cyan

            Console.Write( "12) " ) 
            PassFail(QBColor(12) = &H0000ff&) '  12 - light red

            Console.Write( "13) " ) 
            PassFail(QBColor(13) = &Hff00ff&) '  13 - light magenta

            Console.Write( "14) " ) 
            PassFail(QBColor(14) = &H00ffff&) '  14 - light yellow

            Console.Write( "15) " ) 
            PassFail(QBColor(15) = &Hffffff&) '  15 - bright white
    
            Console.WriteLine()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub RGBTest()
        On Error Resume Next
        Console.WriteLine( "*** RGB test" )

        Err.Number = 0
        Console.WriteLine( RGB(-1, 0, 0) )
        ErrorCheck( 5 )
        Err.Number = 0 
        Console.WriteLine( RGB(0, -1, 0)  )                
        ErrorCheck( 5 )
        Err.Number = 0 
        Console.WriteLine( RGB(0, 0, -1) )                  
        ErrorCheck( 5 )
        Console.WriteLine( RGB(0, 0, 0) = 0 )
        Console.WriteLine( RGB(&HFF, &HFF, &HFF) = &H00FFFFFF& )
        Console.WriteLine( RGB(&H1, &H2, &H3) = &H30201& )
        Console.WriteLine()
    End Sub



    Sub TypeNameTest()
        Console.WriteLine( "*** TypeName test" )

        Try
            Console.WriteLine( "* TypeName non-array tests *" )

            Console.Write( "1) " ) 
            PassFail(TypeName(i) = "Integer")

            Console.Write( "2) " ) 
            PassFail(TypeName(l) = "Long")

            Console.Write( "3) " ) 
            PassFail(TypeName(sng) = "Single")

            Console.Write( "4) " ) 
            PassFail(TypeName(dbl) = "Double")

            dt = #1/1/1999#
            Console.Write( "5) " ) 
            PassFail(TypeName(dt) = "Date")

            Console.Write( "6) " ) 
            PassFail(TypeName(s) = "String")

            Console.Write( "7) " ) 
            PassFail(TypeName(obj) = "Nothing")

            Console.Write( "8) " ) 
            PassFail(TypeName(sb) = "StringBuilder")

            Console.Write( "9) " ) 
            PassFail(TypeName(b) = "Boolean")

            Console.Write( "10) " ) 
            PassFail(TypeName(dec) = "Decimal")

            Console.Write( "11) " ) 
            PassFail(TypeName(byt) = "Byte")

            Console.Write( "12) " ) 
            PassFail(TypeName(ch) = "Char")

            Console.Write( "13) " ) 
            PassFail(TypeName(shrt) = "Short")

            Dim ai(0) As Integer
            Dim al(0) As Long
            Dim asng(0) As Single
            Dim adbl(0) As Double
            Dim adt(0) As Date
            Dim adtNull(0) As Date
            Dim astr(0) As String
            Dim aobj(0) As System.Object
            Dim ab(0) As Boolean
            Dim adec(0) As Decimal
            Dim abyt(0) As Byte
            Dim audt(0) As MyType
            Dim ac(0) As Char
            Dim ashrt(0) As Short
            Dim asb(0) As StringBuilder

            Console.WriteLine( "* TypeName array tests *" )
            Console.Write( "1) " ) 
            PassFail(TypeName(ai) = "Integer()")

            Console.Write( "2) " ) 
            PassFail(TypeName(al) = "Long()")

            Console.Write( "3) " ) 
            PassFail(TypeName(asng) = "Single()")

            Console.Write( "4) " ) 
            PassFail(TypeName(adbl) = "Double()")

            Console.Write( "5) " ) 
            PassFail(TypeName(adt) = "Date()")

            Console.Write( "6) " ) 
            PassFail(TypeName(astr) = "String()")

            Console.Write( "7) " ) 
            PassFail(TypeName(aobj) = "Object()")

            Console.Write( "8) " ) 
            PassFail(TypeName(asb) = "StringBuilder()")

            Console.Write( "9) " ) 
            PassFail(TypeName(ab) = "Boolean()")

            Console.Write( "10) " ) 
            PassFail(TypeName(adec) = "Decimal()")

            Console.Write( "11) " ) 
            PassFail(TypeName(abyt) = "Byte()")

            Console.Write( "12) " ) 
            PassFail(TypeName(audt) = "MyType()")

            Console.Write( "13) " ) 
            PassFail(TypeName(ac) = "Char()")

            Console.Write( "14) " ) 
            PassFail(TypeName(ashrt) = "Short()")

            Dim clr As MyColors
            Console.Write( "15) " ) 
            PassFail(TypeName(clr) = "MyColors")

        Catch ex As Exception
            Failed(ex)
        End Try
        
        Console.WriteLine()
    End Sub



    Sub VarTypeTest()
        Dim ai(0) As Integer
        Dim al(0) As Long
        Dim asng(0) As Single
        Dim adbl(0) As Double
        Dim adt(0) As Date
        Dim adtNull(0) As Date
        Dim astr(0) As String
        Dim aobj(0) As System.Object
        Dim ab(0) As Boolean
        Dim adec(0) As Decimal
        Dim abyt(0) As Byte
        Dim audt(0) As MyType
        Dim ach(0) As Char
        Dim asb(0) As StringBuilder
        Dim ashrt(0) As Short

        Console.WriteLine( "*** VarType test" )

        Try
            Console.WriteLine( " * Vartype non-array tests *" )

            Console.Write( "1) " ) 
            PassFail(VarType(i) = VariantType.Integer)

            Console.Write( "2) " ) 
            PassFail(VarType(l) = VariantType.Long)

            Console.Write( "3) " ) 
            PassFail(VarType(sng) = VariantType.Single)

            Console.Write( "4) " ) 
            PassFail(VarType(dbl) = VariantType.Double)

            Console.Write( "5) " ) 
            PassFail(VarType(dt) = VariantType.Date)

            Console.Write( "6) " ) 
            PassFail(VarType(s) = VariantType.String)

            Console.Write( "7) " ) 
            PassFail(VarType(obj) = VariantType.Object)

            Console.Write( "8) " ) 
            PassFail(VarType(b) = VariantType.Boolean)

            Console.Write( "9) " ) 
            PassFail(VarType(dec) = VariantType.Decimal)

            Console.Write( "10) " ) 
            PassFail(VarType(byt) = VariantType.Byte)

            Console.Write( "11) " ) 
            PassFail(VarType(ch) = VariantType.Char)

            Console.Write( "12) " ) 
            PassFail(VarType(shrt) = VariantType.Short)

            Console.Write( "13) " ) 
            PassFail(VarType(ai) = (VariantType.Array Or VariantType.Integer))

            Console.Write( "14) " ) 
            PassFail(VarType(aobj) = (VariantType.Array Or VariantType.Object))

            Console.Write( "15) " ) 
            PassFail(VarType(ab) = (VariantType.Array Or VariantType.Boolean))

            Console.WriteLine( " * VarType array tests *" )
            Console.Write( "1) " ) 
            PassFail(VarType(ai) = VariantType.Integer Or VariantType.Array)

            Console.Write( "2) " ) 
            PassFail(VarType(al) = VariantType.Long Or VariantType.Array)

            Console.Write( "3) " ) 
            PassFail(VarType(asng) = VariantType.Single Or VariantType.Array)

            Console.Write( "4) " ) 
            PassFail(VarType(adbl) = VariantType.Double Or VariantType.Array)

            Console.Write( "5) " ) 
            PassFail(VarType(adt) = VariantType.Date Or VariantType.Array)

            Console.Write( "6) " ) 
            PassFail(VarType(astr) = VariantType.String Or VariantType.Array)

            Console.Write( "7) " ) 
            PassFail(VarType(aobj) = VariantType.Object Or VariantType.Array)

            Console.Write( "8) " ) 
            PassFail(VarType(ab) = VariantType.Boolean Or VariantType.Array)

            Console.Write( "9) " ) 
            PassFail(VarType(adec) = VariantType.Decimal Or VariantType.Array)

            Console.Write( "10) " ) 
            PassFail(VarType(abyt) = VariantType.Byte Or VariantType.Array)

            Console.Write( "11) " ) 
            PassFail(VarType(audt) = VariantType.UserDefinedType Or VariantType.Array)

            Console.Write( "12) " ) 
            PassFail(VarType(ach) = VariantType.Char Or VariantType.Array)

            Console.Write( "13) " ) 
            PassFail(VarType(ashrt) = VariantType.Short Or VariantType.Array)

        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine()
    End Sub



    Sub ErrorCheck(ByVal iErrExpected As Long)
        If iErrExpected = Err.Number Then
            Console.WriteLine( "PASSED" )
        Else
            Console.WriteLine( "FAILED Error: " & CStr(Err.Number) & "  Expected " & CStr(iErrExpected) )
        End If
    End Sub



End Module


Module Bug247905
    Sub Test()
        Console.Write("*** Bug 247905: ")
        Try
            Dim s As String
            PassFail(IsNumeric("1") AndAlso IsNumeric(Chr(49)))
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module

