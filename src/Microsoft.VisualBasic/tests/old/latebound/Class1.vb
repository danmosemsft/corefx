Imports System
Imports Microsoft.VisualBasic

Namespace ShadowsClassLibrary

    Public Class c1
        Public Property p(ByVal i As Char) As Integer
            Get
                System.Console.WriteLine("   c1.p get")
            End Get
            Set(ByVal Value As Integer)
                System.Console.WriteLine("   c1.p set")
            End Set
        End Property

        Public Sub foo()
            Console.WriteLine("   c1.foo(): FAILED")
        End Sub

        Public Sub bar(ByVal o As Object)
            Console.WriteLine("   c1.bar(Object)")
        End Sub

        Public Overridable Sub test(ByVal i As Int16)
            Console.WriteLine("   c1.test(Int16): FAILED")
        End Sub

        Public Overridable Sub test2(ByVal i As Int16)
            Console.WriteLine("   c1.test2(Int16): FAILED")
        End Sub

    End Class

    Public Class c2
        Inherits c1
        Public Shadows Property p(ByVal i As Integer) As Integer
            Get
                System.Console.WriteLine("   c2.p get")
            End Get
            Set(ByVal Value As Integer)
                System.Console.WriteLine("   c2.p set")
            End Set
        End Property

        Public Shadows Sub foo()
            Console.WriteLine("   c2.foo()")
        End Sub

        Public Overloads Sub bar(ByVal i As Int16)
            Console.WriteLine("   c2.bar(Int16)")
        End Sub

        Public Shadows Sub test(ByVal i As Int32)
            Console.WriteLine("   c2.test(Int32)")
        End Sub

        Public Overloads Overrides Sub test2(ByVal i As Int16)
            Console.WriteLine("   c2.test2(Int16)")
        End Sub
    End Class

End namespace
