Option Strict Off

Imports System
Imports Microsoft.VisualBasic
Imports TestHarness

Public Enum MyEnum As Byte
    a = 1
    b = 2
    c = 3
End Enum

Public Class OverloadsBaseClass0

End Class

Public Class OverloadsBaseClass
    Inherits OverloadsBaseClass0

End Class

Public Class OverloadsClass1
    Inherits OverloadsBaseClass

    Public Overloads Function Foo() As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As Byte) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As Char) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As Short) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As Integer) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As Long) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As Single) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As Double) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As Decimal) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As String) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As MyEnum) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As Object) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As OverloadsBaseClass0) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As OverloadsBaseClass) As Integer
        DumpCaller()
    End Function

End Class

Public Class OverloadsClass2

'    Public Overloads Function Foo(ByVal x As Byte) As Integer
'        DumpCaller()
'    End Function

    Public Overloads Function Foo(ByVal x As Char) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As Short) As Integer
        DumpCaller()
    End Function

    Public Overloads Function Foo(ByVal x As Double) As Integer
        DumpCaller()
    End Function

    'Public Overloads Function Foo(ByVal x As Decimal) As Integer
    '    DumpCaller()
    'End Function

End Class

Module OverloadsTests

    Private Int32Value As Integer
    Private Int16Value As Short
    Private ByteValue As Byte
    Private CharValue As Char
    Private Int64Value As Long
    Private SingleValue As Single
    Private DoubleValue As Double
    Private DecimalValue As Decimal
    Private MyEnumValue As MyEnum
    Private StringValue As String = "ABC"

    Sub LateboundOverloadResolutionTests()
        Try
            TestGroup1()
            TestGroup2()
            RegressionTests()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

    Sub TestGroup1()
        Dim o, o1 As Object
        Dim c As OverloadsClass1 = New OverloadsClass1()

        Console.WriteLine()
        Console.WriteLine("Overloads testing: Group 1")
        Try
            o1 = c

            'Below are matching early and latebound calls

            o = c.Foo()
            o = o1.Foo()

            o = c.Foo(ByteValue)
            o = o1.Foo(ByteValue)

            o = c.Foo(CharValue)
            o = o1.Foo(CharValue)

            o = c.Foo(Int16Value)
            o = o1.Foo(Int16Value)

            o = c.Foo(Int32Value)
            o = o1.Foo(Int32Value)

            o = c.Foo(Int64Value)
            o = o1.Foo(Int64Value)

            o = c.Foo(SingleValue)
            o = o1.Foo(SingleValue)

            o = c.Foo(DoubleValue)
            o = o1.Foo(DoubleValue)

            o = c.Foo(DecimalValue)
            o = o1.Foo(DecimalValue)

            o = c.Foo(StringValue)
            o = o1.Foo(StringValue)

            o = c.Foo(MyEnumValue)
            o = o1.Foo(MyEnumValue)

            '        o = c.Foo(New TimeSpan())
            '        o = o1.Foo(New TimeSpan())

            o = c.Foo(c)
            o = o1.Foo(c)

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

    Sub TestGroup2()
        Dim o, o1 As Object
        Dim c As OverloadsClass2 = New OverloadsClass2()
        Dim ByteValue As Byte
        Dim SingleValue As Single

        Console.WriteLine()
        Console.WriteLine("Overloads testing: Group 2")

        Try

            o1 = c

            'Below are matching early and latebound calls

            o = c.Foo(ByteValue)
            o = o1.Foo(ByteValue)

            o = c.Foo(SingleValue)
            o = o1.Foo(SingleValue)

            'o = c.Foo(MyEnumValue)
            'o = o1.Foo(MyEnumValue)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

    Sub RegressionTests()
        Console.WriteLine()
        Console.WriteLine("Regression tests")
        Bug218010
    End Sub

    Sub Bug218010()
        Try
            Dim X0 As Integer = -32767
            Dim X1 As Object = -32767

            Console.Write("Bug218010: ")

            If Math.Abs(X0) = Math.Abs(X1) AndAlso TypeName(Math.Abs(X0)) = Typename(Math.Abs(X1)) Then
                Passed()
            Else
                Failed()
            End If
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

    '***
    '***  IMPORTANT NOTE!!! 
    '***
    '***  /debug+ /optimize- must be used to compile this module
    '***  for this StackFrame code to work
    Public Sub DumpCaller()
        Dim sf As System.Diagnostics.StackFrame = New System.Diagnostics.StackFrame(1)
        Console.WriteLine(sf.GetMethod().ToString())
    End Sub

End Module

