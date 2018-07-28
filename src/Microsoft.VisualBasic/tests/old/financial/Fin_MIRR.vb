Imports System
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Conversion
Imports Microsoft.VisualBasic.Information
Imports Microsoft.VisualBasic.Financial

Module MIRR_Test

    Sub Main()
        Console.WriteLine( "----MIRR----" )
        FIN_MIRR0001
        TestErrors
    End Sub


    Sub FIN_MIRR0001()

        Dim dRetVal As Double
	    Dim Values(4) as Double
	    
        Try

        '   Testing for the correct value for first set of input
            Console.WriteLine( "CorVal - 1" )
	        Values(0) = -70000
	        Values(1) = 22000 : Values(2) = 25000
	        Values(3) = 28000 : Values(4) = 31000
            dRetVal = MIRR(Values, 0.1, 0.12)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for second set of input
            Console.WriteLine( "CorVal - 2" )
	        Values(0) = -10000
	        Values(1) = 6000 : Values(2) = -2000
	        Values(3) = 7000 : Values(4) =  1000
            dRetVal = MIRR(Values, 0.13, 0.18)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for third set of input
            Console.WriteLine( "CorVal -3" )
	        Values(0) = -30000
	        Values(1) = -10000 : Values(2) = 25000
	        Values(3) =  12000 : Values(4) = 15000
            dRetVal = MIRR(Values, 0.2, 0.23)
            Console.WriteLine( dRetVal )

        '	Test for incorrect/inconsistent/special cases
	        Values(0) = -70000
	        Values(1) = -22000 : Values(2) = -25000
	        Values(3) = -28000 : Values(4) = -31000
            dRetVal = MIRR(Values, 0.1, 0.12)         ' All negative cash flows
            Console.WriteLine( dRetVal )

        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
            Console.WriteLine("FAILED !!!")
        End Try
    End Sub



    Sub TestErrors
        'Try/Catch and On Error must be in separate routines
        Dim Result As Double
        Dim Values(4) As Double

		On Error Resume Next

	    Values(0) = 70000
	    Values(1) = 22000 : Values(2) = 25000
	    Values(3) = 28000 : Values(4) = 31000
        Result = MIRR(Values, 0.1, 0.12)         ' 
        Console.WriteLine( "All positive cash flows: Error =" & Err.Number )
    End Sub

End Module

