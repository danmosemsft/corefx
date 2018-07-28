Imports System
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Conversion
Imports Microsoft.VisualBasic.Information
Imports Microsoft.VisualBasic.Financial

Module SLN_Test

    Sub Main()
        Console.WriteLine( "----SLN----" )
        FIN_SLN0001
    End Sub


    Sub FIN_SLN0001()

        Dim dRetVal As Double

	    Try

        '    Testing for the correct value for first set of input
            Console.WriteLine( "CorVal - 1" )
            dRetVal = SLN(5000, 1000, 20)
            Console.WriteLine( dRetVal )

        '    Testing for the correct value for second set of input
            Console.WriteLine( "CorVal - 2" )
            dRetVal = SLN(54870, 21008, 7)
            Console.WriteLine( dRetVal )

        '    Testing for the correct value for third set of input
            Console.WriteLine( "CorVal -3" )
            dRetVal = SLN(2, 1.1, 12)
            Console.WriteLine( dRetVal )

        '    Test for incorrect/inconsistent/special cases
            Console.WriteLine( "Spl" )
            dRetVal = SLN(1000, 5000, 20)      ' salvage > cost
            Console.WriteLine( dRetVal )

            dRetVal = SLN(5000, 1000, -20)     ' life < 0
            Console.WriteLine( dRetVal )

        '   dRetVal = SLN(5000, 1000, 0)       ' life = 0
        '   Console.WriteLine Error(Err.Number)
        '   Err.Clear

            dRetVal = SLN(5000, 0, 12)         ' salvage = 0
            Console.WriteLine( dRetVal )

            dRetVal = SLN(-5000, -1000, -20)   ' All parameter -ve
            Console.WriteLine( dRetVal )

        '   dRetVal = SLN(0, 0, 0)             ' All parameter 0
        '   Console.WriteLine Error(Err.Number)
        '   Err.Clear

        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
            Console.WriteLine("FAILED !!!")
        End Try

    End Sub

End Module


