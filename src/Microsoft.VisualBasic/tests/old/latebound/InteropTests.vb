Imports System
Imports Microsoft.VisualBasic
Imports System.ComponentModel 'For DesignerSerializationVisibilityAttribute Also requires a command line reference: /r:c:\vs\vsbuilt\debug\bin\i386\complus\system.dll
Imports TestHarness

Namespace InteropTests

    Class Bug255122

        Class c2
            Inherits Project1.Class1
            Public Shadows foo As String
        End Class

        Shared Sub Test()
            Dim c As New c2()
            Dim cc As Project1.Class1 = c
            Dim o As Object = c

            c.foo = "abcd"
            cc.foo = "efgh"
            o.foo = "cd"

            Console.WriteLine(c.foo)
            Console.WriteLine(cc.foo)
        End Sub

    End Class

End Namespace

