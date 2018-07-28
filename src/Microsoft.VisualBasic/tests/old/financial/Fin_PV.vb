Imports System
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Conversion
Imports Microsoft.VisualBasic.Information
Imports Microsoft.VisualBasic.Financial

Module PV_Test

    Sub Main()
        Console.WriteLine( "----PV----" )
        FIN_PV0001
    End Sub


    Sub FIN_PV0001()

        Dim dRetVal As Double

        Try
        '   Testing for the correct value for first set of input
            Console.WriteLine( "CorVal - 1" )
            dRetVal = PV(0.008, 31, 2000.0, 0, 0)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for second set of input
            Console.WriteLine( "CorVal - 2" )
            dRetVal = PV(0.012, 15,  780.0, 2000.0, 1)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for third set of input
            Console.WriteLine( "CorVal -3" )
            dRetVal = PV(0.0096, 54, 123.0, 4760.0, 0)
            Console.WriteLine( dRetVal )

        '   Testing for default values being taken correctly for optional argument
            Console.WriteLine( "DefVal" )
            dRetVal = PV(0.008, 4, 12, 3000)
            Console.WriteLine( dRetVal )
	        dRetVal = PV(0.008, 4, 12, 3000, 0)
	        Console.WriteLine( dRetVal )

        '	Test for incorrect/inconsistent/special cases
            Console.WriteLine( "Spl" )
            dRetVal = PV(-0.008,  31, 2000.0, 0, 0)   ' rate < 0
	        Console.WriteLine( dRetVal )

            dRetVal = PV( 0,      31, 2000.0, 0, 0)   ' rate = 0
            Console.WriteLine( dRetVal )

            dRetVal = PV( 0.008, -31, 2000.0, 0, 0)   ' nper < 0
            Console.WriteLine( dRetVal )

            Console.WriteLine( "OverFlow" )
            dRetVal = PV(1E25, 12, 1797, 0, 1)
            Console.WriteLine( dRetVal )

        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
            Console.WriteLine("FAILED !!!")
        End Try

    End Sub

End Module


