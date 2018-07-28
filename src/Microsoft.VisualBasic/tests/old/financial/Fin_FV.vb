Imports System
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Conversion
Imports Microsoft.VisualBasic.Information
Imports Microsoft.VisualBasic.Financial

Module FV_Test

    Sub Main()
        Console.WriteLine( "----FV----" )
        FIN_FV0001
    End Sub


    Sub FIN_FV0001()
        Dim dRetVal As Double

        Try

        '   Testing for the correct value for first set of input
            Console.WriteLine( "CorVal - 1" )
            dRetVal = FV(0.0083, 15, 263#, 0, DueDate.EndOfPeriod)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for second set of input
            Console.WriteLine( "CorVal - 2" )
            dRetVal = FV(0.013, 90, 81#, 5000#, DueDate.BegOfPeriod)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for third set of input
            Console.WriteLine( "CorVal - 3" )
            dRetVal = FV(0.01, 37, 100#, 0, DueDate.BegOfPeriod)
            Console.WriteLine( dRetVal )

            Console.WriteLine( "Default Values 1" )
            dRetVal = FV(0.0083, 15, 263#, DueDate.EndOfPeriod)
            Console.WriteLine( dRetVal )

            Console.WriteLine( "Default Values 2" )
            dRetVal = FV(0.0083, 15, 263#)
            Console.WriteLine( dRetVal )

        '   Test for incorrect/inconsistent/special input values
            Console.WriteLine( "rate < 0" )
            dRetVal = FV(-0.0083, 15, 263.0, 0, DueDate.EndOfPeriod)       
            Console.WriteLine( dRetVal )
   
            Console.WriteLine( "rate = 0" )
            dRetVal = FV(0, 15, 263, 0, DueDate.EndOfPeriod)                
            Console.WriteLine( dRetVal )

            Console.WriteLine( "type <> 0 and type <> 1" )
            dRetVal = FV(0.0083, 15, 263.0, 0, 8)     
            Console.WriteLine( dRetVal )

            Console.WriteLine( "OverFlow" )
            dRetVal = FV(1E+25, 12, 1797, 0, 1)
            Console.WriteLine( dRetVal )

        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
            Console.WriteLine("FAILED !!!")
        End Try

    End Sub


End Module

