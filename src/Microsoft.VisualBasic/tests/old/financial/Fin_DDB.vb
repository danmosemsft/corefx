Imports System
Imports System.Console
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Conversion
Imports Microsoft.VisualBasic.Information
Imports Microsoft.VisualBasic.Financial

Module DDB_Test

    Sub Main()
        Console.WriteLine( "----DDB---" )
        FIN_DDB0001
        TestErrors
    End Sub


    Sub FIN_DDB0001()

        Dim dRetVal As Double

        Try

        '   Testing for the correct value for first set of input
            Console.WriteLine( "CorVal - 1" )
            dRetVal = DDB(10000#, 4350#, 84#, 35#, 2#)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for second set of input
            Console.WriteLine( "CorVal - 2" )
            dRetVal = DDB(15006#, 6350#, 81#, 23#, 1.5)
            Console.WriteLine( dRetVal )

        '   Testing for the correct value for third set of input
            Console.WriteLine( "CorVal - 3" )
            dRetVal = DDB(11606#, 6350#, 74#, 17#, 2.1)
            Console.WriteLine( dRetVal )

        '   Testing for default values being taken correctly for optional argument
            Console.WriteLine( "DefVal" )
            dRetVal = DDB(10000#, 4350#, 84#, 35#)
            Console.WriteLine( dRetVal )

        '   Test for incosistent input values (cost < salvage)
            Console.WriteLine( "InconVal" )
            dRetVal = DDB(11606#, 16245#, 71#, 17#, 2#)
            Console.WriteLine( dRetVal )

        '   Test for special case  (cost = salvage)
            Console.WriteLine( "SplCase" )
            dRetVal = DDB(10100#, 10100#, 70#, 20#, 2#)
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

        Result = DDB(-10000#, 5211#, -81#, 35#, 2#)
        Console.WriteLine( "Negative input values: Error =" & Err.Number )
    End Sub


End Module
