Imports System
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Conversion
Imports Microsoft.VisualBasic.Information
Imports Microsoft.VisualBasic.Financial

Module PPmt_Test

    Sub Main()
        Console.WriteLine( "----PPmt----" )
        FIN_PPmt0001
    End Sub


    Sub FIN_PPmt0001()

        Dim dRetVal As Double

        Try

        '   Testing for the correct value for first set of input
            Console.WriteLine( "CorVal - 1" )
            dRetVal = PPmt(0.008, 4, 12, 3000, 0, 0)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for second set of input
            Console.WriteLine( "CorVal - 2" )
            dRetVal = PPmt(0.012, 15, 79, 2387, 200, 1)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for third set of input
            Console.WriteLine( "CorVal -3" )
            dRetVal = PPmt(0.0096, 54, 123, 4760, 0, 0)
            Console.WriteLine( dRetVal )

        '   Testing for default values being taken correctly for optional argument
            Console.WriteLine( "DefVal" )
            dRetVal = PPmt(0.008, 4, 12, 3000)
            Console.WriteLine( dRetVal )

            dRetVal = PPmt(0.008, 4, 12, 3000, 0)
            Console.WriteLine( dRetVal )

        '   Test for incorrect/inconsistent/special cases
            Console.WriteLine( "Spl" )
            dRetVal = PPmt(-0.008, 4, 12, 3000, 0, 0)      ' rate < 0
            Console.WriteLine( dRetVal )

        '   dRetVal = PPmt(0.008, -4, 12, 3000, 0, 0)      ' per  < 0
        '   Debug.Print Error(Err.Number)
        '	Err.Clear

        '   dRetVal = PPmt(0.008, 4, -12, 3000, 0, 0)      ' nper < 0
        '   Debug.Print Error(Err.Number)
        '	Err.Clear

            dRetVal = PPmt(0.008, 4, 12, 3000, 0, 7)       ' Type <> 0 and Type <> 1
            Console.WriteLine( dRetVal )

        '   dRetVal = PPmt(0, 4, 12, 3000, 0, 0)           ' rate = 0
        '   Debug.Print Error(Err.Number)
        '	Err.Clear

        '   dRetVal = PPmt(0.008, 12, 4, 3000, 0, 0)       ' per > nper
        '   Debug.Print Error(Err.Number)
        '	Err.Clear

        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
            Console.WriteLine("FAILED !!!")
        End Try

    End Sub

End Module


