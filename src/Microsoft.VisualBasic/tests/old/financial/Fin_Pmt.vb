Imports System
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Conversion
Imports Microsoft.VisualBasic.Information
Imports Microsoft.VisualBasic.Financial

Module Pmt_Test

    Sub Main()
        Console.WriteLine( "----Pmt----" )
        FIN_Pmt0001
    End Sub


    Sub FIN_Pmt0001()

        Dim dRetVal As Double

	    Try

        '   Testing for the correct value for first set of input
            Console.WriteLine( "CorVal - 1" )
            dRetVal = Pmt(0.007, 25, -3000, 0, 0)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for second set of input
            Console.WriteLine( "CorVal - 2" )
            dRetVal = Pmt(0.019, 70, 80000, 20000, 1)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for third set of input
            Console.WriteLine( "CorVal -3" )
            dRetVal = Pmt(0.0012, 5, 500, 0, 0)
            Console.WriteLine( dRetVal )

        '   Testing for default values being taken correctly for optional argument
            Console.WriteLine( "DefVal" )
            dRetVal = Pmt(0.007, 25, -3000)
            Console.WriteLine( dRetVal )
	        dRetVal = Pmt(0.007, 25, -3000, 0)
	        Console.WriteLine( dRetVal )

        '	Test for incorrect/inconsistent/special cases
            Console.WriteLine( "Spl" )
            dRetVal = Pmt(-0.007,  25, -3000, 0, 0)   ' rate < 0
	        Console.WriteLine( dRetVal )

            dRetVal = Pmt( 0.007, -25,  3000, 0, 0)   ' nper < 0
            Console.WriteLine( dRetVal )

	        dRetVal = Pmt( 0.007,  25,  3000, 0, 7)   ' Type <> 0 and Type <> 1
            Console.WriteLine( dRetVal )

            dRetVal = Pmt( 0    ,  25,  3000, 0, 0)   ' rate = 0
            Console.WriteLine( dRetVal )

        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
            Console.WriteLine("FAILED !!!")
        End Try

    End Sub

End Module


