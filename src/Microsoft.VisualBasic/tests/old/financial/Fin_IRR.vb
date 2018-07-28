Imports System
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Conversion
Imports Microsoft.VisualBasic.Information
Imports Microsoft.VisualBasic.Financial

Module IRR_Test



    Sub Main()
        Console.WriteLine( "----IRR----" )
        FIN_IRR0001
        TestErrors
    End Sub



    Sub FIN_IRR0001()
        Dim dRetVal As Double
	    Dim Values(4) as Double

        Try
            'Testing for the correct value for first set of input
            Console.WriteLine( "CorVal - 1" )
	        Values(0) = -70000
	        Values(1) = 22000 : Values(2) = 25000
	        Values(3) = 28000 : Values(4) = 31000
            dRetVal = IRR(Values)
            Console.WriteLine( dRetVal )

            'Testing for the correct value for second set of input
            Console.WriteLine( "CorVal - 2" )
	        Values(0) = -10000
	        Values(1) = 6000 : Values(2) = -2000
	        Values(3) = 7000 : Values(4) =  1000
            dRetVal = IRR(Values)
            Console.WriteLine( dRetVal )

            'Testing for the correct value for third set of input
            Console.WriteLine( "CorVal -3" )
	        Values(0) = -30000
	        Values(1) = -10000 : Values(2) = 25000
	        Values(3) =  12000 : Values(4) = 15000
            dRetVal = IRR(Values)
            Console.WriteLine( dRetVal )

        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
            Console.WriteLine("FAILED !!!")
        End Try
    End Sub



    Sub TestErrors
        'Try/Catch and On Error must be in separate routines
        Dim Result As Double
        Dim Empty() as Double
	    Dim Values(4) as Double

		On Error Resume Next

        Result = IRR( Empty )
        Console.WriteLine( "IRR with empty array: Error =" & Err.Number )

	    Values(0) = 70000
	    Values(1) = 22000
        Values(2) = 25000
	    Values(3) = 28000
        Values(4) = 31000
        Result = IRR(Values)        
        Console.WriteLine( "IRR with all positive cash flows: Error =" & Err.Number )

	    Values(0) = -70000
	    Values(1) = -22000
        Values(2) = -25000
	    Values(3) = -28000
        Values(4) = -31000
        Result = IRR(Values)         
        Console.WriteLine( "IRR with all negative cash flows: Error =" & Err.Number )
    End Sub


End Module

