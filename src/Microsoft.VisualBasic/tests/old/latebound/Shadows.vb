Imports System
Imports Microsoft.VisualBasic
Imports System.ComponentModel 'For DesignerSerializationVisibilityAttribute Also requires a command line reference: /r:c:\vs\vsbuilt\debug\bin\i386\complus\system.dll
Imports TestHarness

Namespace ShadowsTests

    Interface I1
        Function Foo() As Integer
    End Interface

    Interface I2
        Inherits I1
        Shadows Function Foo() As Double
    End Interface

    Class c1
        Implements i2

        Function foo1() As Integer Implements i1.foo
            console.writeline("   foo() as integer")
        End Function

        Function foo2() As Double Implements i2.foo
            console.writeline("   foo() as double")
        End Function
    End Class

    Interface BaseInterface
        Sub func1()
        Event bob()
        Event AnEvent(ByVal x As Integer)
    End Interface

    Interface IntermediateInterface
        Inherits BaseInterface
        Shadows Sub func1()
        Shadows Sub bob(ByVal x As Integer) 'ok since shadows event
    End Interface

    Interface DerivedInterface
        Inherits IntermediateInterface
        Shadows Event AnEvent()
    End Interface

    Class BaseClass

        Public variable As Integer

        Event MyEvent()

        Structure AStruct
            Dim x As Integer
        End Structure

        Enum AnEnum
            a
        End Enum

        Class InnerClass
        End Class

        Public Sub foo()
            console.writeline("   In Base Sub Foo()")
        End Sub

        Public Sub FuncVarCollision()
        End Sub

        Public Sub OverloadedSub()
        End Sub

        Public Sub OverloadedSub2()
        End Sub

        Default ReadOnly Property DefaultProperty(ByVal x As Integer) As Integer
            Get
                Return 50
            End Get
        End Property

    End Class

    Class IntermediateClass
        Inherits BaseClass
        Default Overridable Shadows Property DefaultProperty(ByVal x As Integer) As Integer
            Set(ByVal Value As Integer)
            End Set
            Get
                Return 100
            End Get
        End Property
    End Class

    Class DerivedClass
        Inherits IntermediateClass

        Shadows MyEvent As Integer     'should shadow all the synthetic code, as well
        Shadows variable As Integer
        Shadows FuncVarCollision As Integer     'ok - shadows sub Function() in base

        Shadows Structure AStruct
            Dim x As Integer
        End Structure

        Shadows Enum AnEnum
            a
        End Enum

        Shadows Class InnerClass
        End Class

        Shadows Sub foo()
            variable = 42
            FuncVarCollision = 84
            console.writeline("   Variable: " & variable)
            console.writeline("   FuncVarCollision: " & FuncVarCollision)
            console.writeline("   In Shadows Sub Foo()")
        End Sub

        Public Shadows Sub OverloadedSub(ByVal x As Integer) 'ok - Doesn't require Overloads ok because we shadowed
        End Sub

        Overloads Sub OverloadedSub2(ByVal x As Integer) 'shadowing by name and signature 'overloads'
        End Sub

        Default Overrides Property DefaultProperty(ByVal x As Integer) As Integer
            Set(ByVal Value As Integer)
            End Set
            Get
                Return 150
            End Get
        End Property

    End Class

    ' --- Test shadowing WithEvent vars ( shadows the synthetic artifacts as well )
    Class Source1
        Event E1(ByVal x As String)
        Sub go()
            RaiseEvent e1("source1.go")
        End Sub
    End Class

    Class Source2
        Event E1(ByVal x As Integer)
        Sub go()
            RaiseEvent e1(42)
        End Sub
    End Class

    Class Base
        Public WithEvents x As source1
        Sub foo(ByVal x As String) Handles x.e1
            console.writeline("   Base: foo(" & x & ") handles x.e1")
        End Sub
        Sub go()
            x = New source1()
            x.go()
        End Sub
    End Class

    Class Derived
        Inherits Base

        Public Shadows WithEvents x As source2
        Shadows Sub foo(ByVal x As Integer) Handles x.e1
            console.writeline("   derived: foo(" & x & ") handles x.e1")
        End Sub
        Shadows Sub go()
            x = New source2()
            x.go()
        End Sub
    End Class
    ' --- End Test shadowing WithEvent vars ( shadows the synthetic artifcats as well )

    ' --- Test Shadowing Events On SubObjects ---
    Class EventSource3
        Event MyEvent()
        Sub test()
            RaiseEvent MyEvent()
        End Sub
    End Class

    Class OuterClass2
        Private SubObject As New EventSource3()
        <DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Content)> _
                 Public ReadOnly Property SomeProperty() As EventSource3
            Get
                SomeProperty = SubObject
            End Get
        End Property

        Sub Test()
            SubObject.Test()
        End Sub
    End Class

    Class DerivedOuterClass
        Inherits OuterClass2

        Private Shadows SubObject As New EventSource3()

        <DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Content)> _
        Public Shadows ReadOnly Property SomeProperty() As EventSource3
            Get
                SomeProperty = SubObject
                console.writeline("   In Shadowing EventSource Property")
            End Get
        End Property

        Shadows Sub Test()
            SubObject.Test()
        End Sub

    End Class

    Class Sink
        Dim WithEvents x As DerivedOuterClass

        Sub foo() Handles x.SomeProperty.MyEvent
            console.writeline("   It Worked!")
        End Sub

        Sub test()
            x.Test()
        End Sub

        Sub New()
            x = New DerivedOuterClass()
        End Sub
    End Class

    ' --- End Test Shadowing Events On SubObjects ---

    Module Module1
        Sub Test()
            Test1
            Test2
            Bug256610.Test
            Bug254931.Test
            Bug207083.Test
            Bug301640.Test
            Bug301641.Test
        End Sub

        Sub Test1()
            Dim base As New BaseClass()
            Dim derived As New DerivedClass()
            Dim o As Object

            Console.WriteLine()
            Console.WriteLine("*** Shadows Tests (same project)")
            Console.WriteLine()
            Console.Writeline("*** Test 1a")
            Try
                base.foo()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 1b")
            Try
                o = base
                o.foo()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 2a")
            Try
                derived.foo()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 2b")

            Try
                o = derived
                o.foo()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 3a")
            Try
                console.Writeline("   " & base(1))
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 3b")
            Try
                o = base
                console.Writeline("   " & o(1))
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 4a")
            Try
                console.Writeline("   " & derived(1))
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 4b")
            Try
                o = derived
                console.Writeline("   " & o(1))
            Catch ex As Exception
                Failed(ex)
            End Try

            Dim x As i1
            Dim y As i2

            x = New c1()

#if 0
'This tests interface shadowing
' which is not supported for latebinding
            Console.Writeline("*** Test 5a")
            Try
                x.foo()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 5b")
            Try
                o = x
                o.foo()
            Catch ex As Exception
                Failed(ex)
            End Try
#End If

#if 0
'This tests interface shadowing
' which is not supported for latebinding
            Console.Writeline("*** Test 6a")
            Try
                y = New c1()
                y.foo()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 6b")
            Try
                o = y
                o.foo()
            Catch ex As Exception
                Failed(ex)
            End Try
#End If

            'Shadowed WithEvent vars on classes test
            Dim x1 As New base()
            Dim y1 As New derived()

            Console.Writeline("*** Test 7a")
            Try
                x1.go()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 7b")
            Try
                o = x1 
                o.go()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 8a")
            Try
                y1.go()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 8b")
            Try
                o = y1
                o.go()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 9a")
            'Shadowing events on sub objects
            Dim x3 As Sink
            Try
                x3 = New Sink()
                x3.Test()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Writeline("*** Test 9b")
            Try
                o = New Sink()
                o.Test
            Catch ex As Exception
                Failed(ex)
            End Try

        End Sub



        Sub Test2()
            Dim c As New ShadowsClassLibrary.c2()
            Dim i As Integer
            Dim dec As Decimal
            dec = 1@

            Console.WriteLine()
            Console.WriteLine("*** Shadows Tests (project import)")
            Console.WriteLine()
            Console.WriteLine("Earlybound")

            Try
                Console.WriteLine("*** Test 1")
                c.p(0) = 1
            Catch ex As Exception
                failed(ex)
            End Try

            Try
                Console.WriteLine("*** Test 2")
                'Should attempt to call c2.p(Integer) which should give an invalid cast compile time error
                'c.p("a"c) = 1
                Passed()
            Catch ex As Exception
                failed(ex)
            End Try

            Try
                Console.WriteLine("*** Test 3")
                c.foo()
            Catch ex As Exception
                failed(ex)
            End Try
            Try
                Console.WriteLine("*** Test 4")
                c.bar(1)
            Catch ex As Exception
                failed(ex)
            End Try
            Try
                Console.WriteLine("*** Test 5")
                c.test(1)
            Catch ex As Exception
                failed(ex)
            End Try
            Try
                Console.WriteLine("*** Test 6")
                c.test2(1)
            Catch ex As Exception
                failed(ex)
            End Try

            Console.WriteLine("")
            Console.WriteLine("LateBound")
            Dim o As Object
            o = c

            Try
                Console.WriteLine("*** Test 1")
                o.p(0) = 1
            Catch ex As Exception
                failed(ex)
            End Try

            Try
                Console.WriteLine("*** Test 2")
                'Should attempt to call c2.p(Integer) which should give an invalid cast
                o.p("a"c) = 1
                Failed()
            Catch ex As InvalidCastException
                Console.Write("   ")
                Passed()
            Catch ex As Exception
                failed(ex)
            End Try

            Try
                Console.WriteLine("*** Test 3")
                o.foo()
            Catch ex As Exception
                failed(ex)
            End Try

            Try
                Console.WriteLine("*** Test 4")
                o.bar(1)
            Catch ex As Exception
                failed(ex)
            End Try

            Try
                Console.WriteLine("*** Test 5")
                o.test(1)
            Catch ex As Exception
                failed(ex)
            End Try

            Try
                Console.WriteLine("*** Test 6")
                o.test2(1)
            Catch ex As Exception
                failed(ex)
            End Try

        End Sub

    End Module

    Module Bug207083
        Class c1
            Private m_iValue As Integer
            Property p(ByVal i As Char) As Integer
                Get
                    Return m_iValue
                End Get
                Set(ByVal Value As Integer)
                    m_iValue = Value
                End Set
            End Property
        End Class

        Class c2
            Inherits c1
            Private m_iValue As Integer
            Shadows Property p(ByVal i As Integer) As Integer
                Get
                    Return m_iValue
                End Get
                Set(ByVal Value As Integer)
                    m_iValue = Value
                End Set
            End Property
        End Class

        Sub Test()
            Console.WriteLine("Bug 207083")

            Dim c1 As C1
            Dim c2 as C2
            Dim i As Integer
            Dim o As Object

            c2 = New C2()
            c1 = c2

            c1.p("a") = 1
            c2.p(9) = 2

            o = c2
            Console.Write("   1) ")
            Try
                i = o.p("c"c)
                Failed()
                Console.WriteLine("o.p(" & ChrW(34) & "c" & ChrW(34) & "c) = i" & i)
            Catch ex As InvalidCastException
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("   2) ")
            Try
                o.p("c"c) = i
                Failed()
                Console.WriteLine("o.p(" & ChrW(34) & "c" & ChrW(34) & "c) = i")
            Catch ex As InvalidCastException
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

        End Sub

    End Module

    Module Bug256610
        Dim Result As Integer

        Class Cls1
            Sub foo(ByVal arg As Short)
                Result = arg
            End Sub
        End Class

        Class Cls2
            Inherits Cls1

            Shadows Sub foo(ByVal arg As Short)
                Result = arg * 2
            End Sub
        End Class

        Sub Test()

            Console.Write("*** Bug 256610: ")
            Dim c As New Cls2()
            Dim Obj As Object = c
            Try
                Dim EarlyResult As Integer

                c.foo(10)
                EarlyResult = Result

                obj.foo(10)
                PassFail(EarlyResult = Result)

            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub
    End Module

    Class Bug254931

        Shared Result As Integer

        Class c1
            Sub Foo(ByVal i As Short)
                Result = 1
            End Sub
        End Class

        Class c2
            Inherits c1

            Shadows Sub Foo(ByVal i As Short)
                Result = 2
            End Sub

        End Class

        Shared Sub Test()
            Console.Write("*** Bug 254931: ")
            Dim o As Object = New c2()

            Try
                o.foo(CShort(1))
                PassFail(Result = 2)
            Catch ex as Exception
                Failed(ex)
            End Try

            'Same classes, different argument
            Console.Write("*** Bug 254933: ")
            Try
                Result = 0
                o.foo(CByte(1))
                PassFail(Result = 2)
            Catch ex as Exception
                Failed(ex)
            End Try

        End Sub
    End Class

    Class Bug301640

        Private Class c1
            Sub foo()

            End Sub
        End Class

        Private Class c2
            Inherits c1
            Shadows Class foo

            End Class
        End Class

        Shared Sub Test()

            Console.Write("*** Bug 301640: ")
            Try
                Dim c As New c2()
                Dim o As Object

                o = c
                o.foo()

                Failed()
            Catch ex As ArgumentException
                Passed()
                Console.WriteLine("   " & ex.Message)
            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub

    End Class


    Class Bug301641

        Private Class c1
            Sub x()
                Console.WriteLine("shouldn't have ran")
            End Sub
        End Class

        Private Class c2
            Inherits c1
            Overridable Shadows Sub x(ByVal i As Integer)
            End Sub
        End Class

        Private Class c3
            Inherits c2
            Overrides Sub x(ByVal i As Integer)
            End Sub
        End Class

        Shared Sub Test()
            Dim obj As Object = New c3()

            Console.Write("*** Bug 301641: ")
            Try
                obj.x()
                Failed()
            Catch ex As MissingMemberException
                Passed()
                Console.WriteLine("   " & ex.Message)
            Catch ex As Exception
                Failed(ex)
            End Try
        End Sub

    End Class

End Namespace

