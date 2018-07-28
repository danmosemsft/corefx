Imports System
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Conversion
Imports Microsoft.VisualBasic.Information
Imports Microsoft.VisualBasic.Financial

Module Rate_Test

    Sub Main()
        Console.WriteLine( "----Rate----" )
        FIN_Rate0001
        TestErrors
    End Sub


    Sub FIN_Rate0001()
        Dim dRetVal As Double

	    Try

        '   Testing for the correct value for first set of input
            Console.WriteLine( "CorVal - 1" )
            dRetVal = Rate(12, -263.0, 3000, 0, 0)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for second set of input
            Console.WriteLine( "CorVal - 2" )
            dRetVal = Rate(48, -570, 24270.0, 0, 1)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for third set of input
            Console.WriteLine( "CorVal -3" )
            dRetVal = Rate(96, -1000.0, 56818, 0, 0)
            Console.WriteLine( dRetVal )

        '   Testing for default values being taken correctly for optional argument
            Console.WriteLine( "Default Value 1" )
            dRetVal = Rate(12, -263.0, 3000)
            Console.WriteLine( dRetVal )

            Console.WriteLine( "Default Value 2" )
	        dRetVal = Rate(12, -263.0, 3000, 0)
	        Console.WriteLine( dRetVal )

            Console.WriteLine( " pmt > pv" )
            dRetVal = Rate( 12, -3000.0,  300, 0, 0)   
            Console.WriteLine( dRetVal )

        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
            Console.WriteLine("FAILED !!!")
        End Try
    End Sub



    Sub TestErrors
        'Try/Catch and On Error must be in separate routines
        Dim Result As Double

		On Error Resume Next

        Result = Rate(-12,  -263.0, 3000, 0, 0)   ' 
        Console.WriteLine( "nper < 0: Error =" & Err.Number )
    End Sub


End Module
