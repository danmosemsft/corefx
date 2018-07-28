Imports System
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Conversion
Imports Microsoft.VisualBasic.Information
Imports Microsoft.VisualBasic.Financial

Module NPER_Test

    Sub Main()
        Console.WriteLine( "----NPer----" )
        FIN_NPer0001
    End Sub


    Sub FIN_NPer0001()
        Dim dRetVal As Double

        Try
        '   Testing for the correct value for first set of input
            Console.WriteLine( "CorVal - 1" )
            dRetVal = NPer(0.0072, -350.0, 7000.0, 0, 0)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for second set of input
            Console.WriteLine( "CorVal - 2" )
            dRetVal = NPer(0.018, -982.0, 33000.0, 2387, 1)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for third set of input
            Console.WriteLine( "CorVal -3" )
            dRetVal = NPer(0.0096, 1500.0, -70000.0, 10000, 0)
            Console.WriteLine( dRetVal )

        '   Testing for default values being taken correctly for optional argument
            Console.WriteLine( "Default Values 1" )
            dRetVal = NPer(0.0072, -350.0, 7000.0)
            Console.WriteLine( dRetVal )

            Console.WriteLine( "Default Values 2" )
	        dRetVal = NPer(0.0072, -350.0, 7000.0, 0)
	        Console.WriteLine( dRetVal )

        '	Test for incorrect/inconsistent/special cases
            Console.WriteLine( "rate < 0" )
            dRetVal = NPer(-0.0072,  -350.0, 7000.0, 0, 0)            
            Console.WriteLine( dRetVal )
    
            Console.WriteLine( "rate = 0" )
            dRetVal = NPer( 0,       -350.0, 7000.0, 0, 0)  
            'UNDONE:  dretval should be 20
            Console.WriteLine( dRetVal )
    
            Console.WriteLine( "Type <> 0 and Type <> 1" )
            dRetVal = NPer( 0.0072,  -350.0, 7000.0, 0, 7)    
            Console.WriteLine( dRetVal )
    
            Console.WriteLine( "pmt > pv" )
            dRetVal = NPer( 0.0072, -9000.0,  200.0, 0, 0)  ' dretval should be 0.022303910926714
            Console.WriteLine( dRetVal )

        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
            Console.WriteLine("FAILED !!!")
        End Try

    End Sub

End Module
