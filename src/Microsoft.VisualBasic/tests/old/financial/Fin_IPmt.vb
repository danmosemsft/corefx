Imports System
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Conversion
Imports Microsoft.VisualBasic.Information
Imports Microsoft.VisualBasic.Financial

Module IPMT_Test

    Sub Main()
        Console.WriteLine(   "----IPmt----" )
        FIN_IPmt0001
        TestErrors
    End Sub


    Sub FIN_IPmt0001()
        Dim dRetVal As Double

        Try
        '   Testing for the correct value for first set of input
            Console.WriteLine( "CorVal - 1" )
            dRetVal = IPmt(0.008, 4, 12, 3000, 0, 0)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for second set of input
            Console.WriteLine( "CorVal - 2" )
            dRetVal = IPmt(0.012, 15, 79, 2387, 200, 1)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for third set of input
            Console.WriteLine( "CorVal - 3" )
            dRetVal = IPmt(0.0096, 54, 123, 4760, 0, 0)
            Console.WriteLine( dRetVal )

        '   Testing for default values being taken correctly for optional argument
            Console.WriteLine( "DefVal" )
            dRetVal = IPmt(0.008, 4, 12, 3000)

            Console.WriteLine( dRetVal )
            dRetVal = IPmt(0.008, 4, 12, 3000, 0)
            Console.WriteLine( dRetVal )

        '   Test for incorrect/inconsistent/special cases
            Console.WriteLine( "Spl" )
            dRetVal = IPmt(-0.008, 4, 12, 3000, 0, 0)      ' rate < 0
            Console.WriteLine( dRetVal )

            dRetVal = IPmt(0.008, 4, 12, 3000, 0, 7)       ' Type <> 0 and Type <> 1
            Console.WriteLine( dRetVal )

            dRetVal = IPmt(0, 4, 12, 3000, 0, 0)           ' rate = 0
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

        Result = IPmt(0.008, -4, 12, 3000, 0, 0)      
        Console.WriteLine( "per  < 0: Error =" & Err.Number )

        Result = IPmt(0.008, 4, -12, 3000, 0, 0)       
        Console.WriteLine( "nper < 0: Error =" & Err.Number )

        Result = IPmt(0.008, 12, 4, 3000, 0, 0)       ' per > nper
        Console.WriteLine( "per > nper: Error =" & Err.Number )
    End Sub

End Module

