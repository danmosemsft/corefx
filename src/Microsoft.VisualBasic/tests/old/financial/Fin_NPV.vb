Imports System
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Conversion
Imports Microsoft.VisualBasic.Information
Imports Microsoft.VisualBasic.Financial

Module NPV_Test

    Sub Main()
        Console.WriteLine( "----NPV----" )
        FIN_NPV0001     
        TestErrors   
    End Sub


    Sub FIN_NPV0001()
        Dim dRetVal As Double
        Dim Values(4) as Double

        Try

            Console.WriteLine( "First set of input" )
	        Values(0) = -70000
	        Values(1) = 22000 : Values(2) = 25000
	        Values(3) = 28000 : Values(4) = 31000
            dRetVal = NPV(0.0625, Values)
            Console.WriteLine( dRetVal )

            Console.WriteLine( "Second set of input" )
	        Values(0) = -10000
	        Values(1) = 6000 : Values(2) = -2000
	        Values(3) = 7000 : Values(4) =  1000
            dRetVal = NPV(0.089, Values)
            Console.WriteLine( dRetVal )

            Console.WriteLine( "Third set of input" )
	        Values(0) = -30000
	        Values(1) = -10000 : Values(2) = 25000
	        Values(3) =  12000 : Values(4) = 15000
            dRetVal = NPV(0.011, Values)
            Console.WriteLine( dRetVal )

            Console.WriteLine( "rate < 0" )
            dRetVal = NPV(-0.011, Values)         ' 
            Console.WriteLine( dRetVal )

            Console.WriteLine( "All positive cash flows" )
	        Values(0) = 70000
	        Values(1) = 22000 : Values(2) = 25000
	        Values(3) = 28000 : Values(4) = 31000
            dRetVal = NPV(0.0625, Values)         ' 
	        Console.WriteLine( dRetVal )

            Console.WriteLine( "All negative cash flows" )
	        Values(0) = -70000
	        Values(1) = -22000 : Values(2) = -25000
	        Values(3) = -28000 : Values(4) = -31000
            dRetVal = NPV(0.0625, Values)         ' 
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

		On Error Resume Next

        Result = NPV(-1, Empty)
        Console.WriteLine( "NPV with Rate = -1: Error =" & Err.Number )

        Result = NPV(.12, Empty)
        Console.WriteLine( "NPV with empty array: Error =" & Err.Number )
    End Sub

End Module
