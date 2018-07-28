Imports System
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Conversion
Imports Microsoft.VisualBasic.Information
Imports Microsoft.VisualBasic.Financial

Module SYD_Test

    Sub Main()
        Console.WriteLine( "----SYD----" )
        FIN_SYD0001
        TestErrors
    End Sub


    Sub FIN_SYD0001()
        Dim dRetVal As Double

        Try
            'Testing for the correct value for first set of input
            Console.WriteLine( "CorVal - 1" )
            dRetVal = SYD(4322#, 1009#, 73, 23)
            Console.WriteLine( dRetVal )

            'Testing for the correct value for second set of input
            Console.WriteLine( "CorVal - 2" )
            dRetVal = SYD(78000#, 21008, 8, 2)
            Console.WriteLine( dRetVal )

            'Testing for the correct value for third set of input
            Console.WriteLine( "CorVal -3" )
            dRetVal = SYD(23#, 7#, 21, 9)
            Console.WriteLine( dRetVal )

            Console.WriteLine( "Overflow test" )
            dRetVal = SYD(9.9999999999999E+305, 0, 100, 10)
            Console.WriteLine( dRetVal )

            'Test for incorrect/inconsistent/special cases
            Console.WriteLine( "Salvage > cost" )
            dRetVal = SYD(1009#, 4322#, 73, 23)         
            Console.WriteLine( dRetVal )


        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
            Console.WriteLine("FAILED !!!")
        End Try       
    End Sub



    Sub TestErrors()
        Dim dRetVal As Double

        On Error Resume Next

        dRetVal = SYD(4322#, 1009#, 23, 73)          
        Console.WriteLine( "period > life: Error =" & Err.Number)

        dRetVal = SYD(4322#, 1009#, 0, 23)           
        Console.WriteLine( "life = 0: Error =" & Err.Number)

        dRetVal = SYD(4322#, 1009#, 73, 0)          
        Console.WriteLine( "period = 0: Error =" & Err.Number)

        dRetVal = SYD(-4322#, -1009#, -73, -23)     
        Console.WriteLine( "All parameters negative: Error =" & Err.Number)

        dRetVal = SYD(0#, 0#, 0, 0)                 
        Console.WriteLine( "All parameters zero: Error =" & Err.Number)
    End Sub
End Module


