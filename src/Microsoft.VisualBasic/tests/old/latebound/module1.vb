Option Strict Off

Imports System
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports TestHarness
Imports System.Reflection

Public Class Class1
    Private m_default() As String

    Private Shared PrivateSharedValue As Integer = 12
    Protected Shared ProtectedSharedValue As Integer = 23
    Public Shared PublicSharedValue As Integer = 12345

    Sub New()
        ObjectValue = New Integer() {1, 2, 3, 4, 5}
        m_default = New String() {"A", "B", "C", "D", "E"}
        Ary = New Integer() {6, 7, 8, 9, 10}
    End Sub

    Public Ary(4) As Integer
    Public ObjectValue As Object
    Public ShortValue As Short
    Public IntegerValue As Integer
    Public LongValue As Long
    Public SingleValue As Single
    Public DoubleValue As Double
    Public DateValue As Date
    Public DecimalValue As Decimal
    Public StringValue As String

    Private PrivateValue As String

    Protected Property ProtectedProp() As Integer
        Get
            Return 12345
        End Get
        Set(ByVal Value As Integer)
            If value <> 12345 Then
                Throw New ArgumentException("Argument was not correct")
            End If
        End Set
    End Property

    Default Public Property DefaultProp(ByVal Index As Integer) As String
        Get
            Return m_default(Index)
        End Get
        Set(ByVal Value As String)
            m_default(Index) = Value
        End Set
    End Property

    Friend Declare Ansi Function AnsiStrFunction Lib "DeclExtNightly001.DLL" Alias "StrFunction" (ByVal Arg As String, ByVal Arg1 As Integer, ByVal Arg2 As Integer, ByVal Arg3 As Boolean) As Boolean
    Public Declare Sub MsgBeep Lib "user32.DLL" Alias "MessageBeep" (Optional ByVal x As Integer = 0)


End Class

Class Class2
    Default Public Property DefaultProp(ByVal Arg1 As Integer) As Integer
        Get
            If arg1 <> 2 Then
                Throw New ArgumentException("DefaultProp argument were not correct")
            End If
            Return 12345
        End Get
        Set(ByVal Value As Integer)
            If arg1 <> 2 OrElse Value <> 12345 Then
                Throw New ArgumentException("DefaultProp argument were not correct")
            End If
        End Set
    End Property
End Class

Class Class3
    Default Public Property DefaultProp(ByVal Arg1 As String, Optional ByVal Arg2 As Integer = 10) As Integer 'the Optional parameter here contributes to the unexpected exception
        Get
            If arg1 <> "ABC" OrElse Arg2 <> 3 Then
                Throw New ArgumentException("DefaultProp arguments were not correct")
            End If
            Return 12345
        End Get
        Set(ByVal Value As Integer)
            If arg1 <> "ABC" OrElse Arg2 <> 3 OrElse Value <> 12345 Then
                Throw New ArgumentException("DefaultProp arguments were not correct")
            End If
        End Set
    End Property
End Class

Class Class4
    Default Public ReadOnly Property DefaultProp(ByVal Arg1 As Integer) As Integer
        Get
            If arg1 <> 4 Then
                Throw New ArgumentException("DefaultProp argument were not correct")
            End If
            Return 12345
        End Get
    End Property
End Class

Class Class5
    Default Public WriteOnly Property DefaultProp(ByVal Arg1 As Integer) As Integer
        Set(ByVal Value As Integer)
            If arg1 <> 5 Then
                Throw New ArgumentException("DefaultProp argument were not correct")
            End If
            If Value <> 12345 Then
                Throw New ArgumentException("Value assigned to default property did not match")
            End If
        End Set
    End Property
End Class

Class Class6

    Public Sub Foo1(ByRef x As Integer, ByVal y As Integer)
        x = y
    End Sub

    Public Sub Foo2(ByVal x As Integer, ByRef y As Integer)
        y = x
    End Sub

    Public Sub Foo3(ByRef x As Integer, ByRef y As Integer, ByVal z As Integer)
        x = z
        y = z
    End Sub

    Public Sub Foo4(ByVal x As Integer, ByRef y As Integer, ByRef z As Integer)
        y = x
        z = x
    End Sub

    Public Function Foo5(ByRef x As Integer, ByVal y As Integer) As Integer
        x = y
        Return y
    End Function

    Public Function Foo6(ByVal x As Integer, ByRef y As Integer) As Integer
        y = x
        Return x
    End Function

    Public Function Foo7(ByRef x As Integer, ByRef y As Integer, ByVal z As Integer) As Integer
        x = z
        y = z
        Return z
    End Function

    Public Function Foo8(ByVal x As Integer, ByRef y As Integer, ByRef z As Integer) As Integer
        y = x
        z = x
        Return x
    End Function

    Public Sub Foo9(ByVal s1 As String, ByRef s2 As String)
        s2 = s1
    End Sub

End Class

Module Test

    Sub Main()

        Try
            Console.WriteLine("Begin Test")
            LateBoundFieldGetTests()
            LateBoundFieldSetTests()

            LateBoundPrivateFieldGetTest()
            LateBoundPrivateFieldSetTest()

            LateBoundProtectedPropertyTest()

            SharedFieldTests()
            DefaultPropertyTests()

            Console.WriteLine("* Latebound Array indexer tests")

            LateBoundArrayIndexGet1()
            LateBoundArrayIndexGet2()
            LateBoundArrayIndexSet1()
            LateBoundArrayIndexSet2()

            Bug168135.Test()
            Bug166464.Test()
            Bug126622.Test()
            Bug132982.Main()
            Bug149726.Main()
            Bug153295.Main()
            Bug168229.Main()
            Bug231364.Main()
            Bug231397.Main()
            Bug233217.Main()
            Bug236760.Main()
            Bug240327.Test()
            Bug240263.Test()
            Bug239809.Test()
            Bug240872.Test()
            Bug224845.Test()
            Bug257652.Test()
            Bug227370.Test()
            Bug256623.Test()
            Bug257437.Test()
            Bug256635.Test()
            Bug257640.Test()
            Bug257644.Test()
            Bug258409.Test()
            Bug263543.Test()
            Bug264334.Test()
            Bug236108.Test()
            Bug259716.Test()
            Bug266851.Test()
            Bug266843.Test()
            Bug266851a.Test()
            Bug264079.Test()
            Bug271007.Test()
            Bug271373.Test()
            Bug271377.Test()
            Bug271685.Test()
            Bug277488.Test()
            Bug278308.Test()
            Bug280519.Test()
            Bug279014.Test()
            Bug283039.Test()
            Bug257649.Test()
            Bug275937.Test()
            Bug287347.Test()
            Bug287352.Test()
            Bug287357.Test()
            Bug287831.Test()
            Bug287837.Test()
            Bug272002.Test()
            Bug270472.Test()
            Bug284107.Test()
            Bug291120.Test()
            Bug291287.Test()
            Bug293134.Test()
            Bug293485.Test()
            Bug294938.Test()
            Bug294942.Test()
            Bug294987.Test()
            Bug294970.Test()
            Bug294956.Test()
            Bug294958.Test()
            Bug297690.Test()
            Bug297717.Test()
            Bug297731.Test()
            Bug296923.Test
            Bug297792.Test
            Bug297815.Test
            Bug297835.Test
            Bug297892.Test
            'BUGBUG: not fixed yet Bug298131.Test
            Bug298658.Test
            Bug302246.Test
            Bug302354.Test
            Bug302374.Test
            Bug302374b.Test
            Bug305634.Test
            Bug304212.Test
            Bug306698.Test
            Bug307060.Test
            Bug307060b.Test
            Bug306786.Test
            Bug308339.Test
            Bug308219.Test
            Bug308240.Test
            Bug308245.Test
            Bug308345.Test
            Bug309526.Test 
            Bug309529.Test
            Bug309444.Test
            Bug309445.Test
            Bug309554.Test
            Bug309613.Test
            Bug309766.Test
            'Bug309766b.Test
            Bug307268.Test
            Bug309820.Test 'Enable when fix complete
            Bug309822.Test
            Bug309845.Test
            Bug309881.Test
            Bug309882.Test
            Bug309883.Test
            Bug309884.Test
            Bug309894.Test
            Bug310044.Test
            Bug309805.Test
            Bug310146.Test
            Bug310322.Test
            Bug310339.Test
            Bug310340.Test
            Bug310341.Test
            Bug310333.Test
            Bug310784.Test
            Bug310786.Test
            Bug310851.Test
            Bug310901.Test
            Bug298655.Test
            Bug311040.Test
            Bug311133.Test
            Bug337350.Test
            'POSTPONED Bug311283.Test
            'POSTPONED Bug311285.Test
            'POSTPONED Bug311223.Test
            BugVSW32724.Test
            BugVSW32699.Test
            BugVSW32809.Test
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin
            'Bug      .Test 'placeholders so the line numbers on warnings don't change on every checkin

            DecimalDateOptionalDefaults.Test()

            LateBoundStringCompare()

            LateBoundByRefTests()

            LateBoundNamedParamTests()

            TestLateBoundCompare()

            LateboundOverloadResolutionTests()

            LateBoundStructureTests.Main()

            TestByRefLateBoundParam()

            ShadowsTests.Module1.Test()

            TestRValueBaseAndOptimisticSet()
            Console.WriteLine("End Test")
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

    Friend Class NamedParamTestClass

        Sub foo1(Optional ByVal opta As String = "def A" & ChrW(0), Optional ByVal optb As String = "def B" & ChrW(0), Optional ByVal optc As String = "def C" & ChrW(0))
            Console.WriteLine("  opta = " & opta)
            Console.WriteLine("  optb = " & optb)
            Console.WriteLine("  optc = " & optc)
        End Sub

        Sub foo2(ByVal x As Integer, Optional ByVal opta As String = "def A" & ChrW(0), Optional ByVal optb As String = "def B" & ChrW(0), Optional ByVal optc As String = "def C" & ChrW(0))
            Console.WriteLine("  x    = " & x)
            Console.WriteLine("  opta = " & opta)
            Console.WriteLine("  optb = " & optb)
            Console.WriteLine("  optc = " & optc)
        End Sub

    End Class

    Sub LateBoundNamedParamTests()
        Console.WriteLine("LateBound Named Param Tests")

        Try
            Dim a As String
            Dim B As String
            Dim C As String

            Dim o As Object
            o = New NamedParamTestClass()
            A = "A"
            B = "B"
            C = "C"

            Console.WriteLine("* Test of foo1 taking 3 optional String parameters")
            Console.WriteLine("A")
            o.foo1(OPTA:=A)

            Console.WriteLine("B")
            o.foo1(OPTB:=B)

            Console.WriteLine("C")
            o.foo1(OPTC:=C)

            Console.WriteLine("AB")
            o.foo1(OPTA:=A, optB:=B)

            Console.WriteLine("BA")
            o.foo1(OPTB:=B, optA:=A)

            Console.WriteLine("AC")
            o.foo1(OPTA:=A, optC:=C)
            Console.WriteLine("CA")
            o.foo1(OPTC:=C, optA:=A)

            Console.WriteLine("BC")
            o.foo1(OPTB:=B, optC:=C)
            Console.WriteLine("CB")
            o.foo1(OPTC:=C, optB:=B)

            'Three arguments
            Console.WriteLine("ABC")
            o.foo1(OPTA:=A, optB:=B, optc:=C)

            Console.WriteLine("ACB")
            o.foo1(OPTA:=A, optC:=C, optB:=B)

            Console.WriteLine("BAC")
            o.foo1(OPTB:=B, optA:=A, optC:=C)

            Console.WriteLine("BCA")
            o.foo1(OPTB:=B, optC:=C, optA:=A)

            Console.WriteLine("CAB")
            o.foo1(OPTC:=C, opta:=A, optB:=B)

            Console.WriteLine("CBA")
            o.foo1(OPTC:=C, optB:=B, optA:=A)

        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Dim a As String
            Dim B As String
            Dim C As String

            Dim o As Object
            o = New NamedParamTestClass()
            A = "A"
            B = "B"
            C = "C"

            Console.WriteLine("* Test of foo1 taking 1 fixed and 3 optional String parameters")
            Console.WriteLine("A")
            o.foo2(1, OPTA:=A)

            Console.WriteLine("B")
            o.foo2(2, OPTB:=B)

            Console.WriteLine("C")
            o.foo2(3, OPTC:=C)

            Console.WriteLine("AB")
            o.foo2(4, OPTA:=A, optB:=B)

            Console.WriteLine("BA")
            o.foo2(5, OPTB:=B, optA:=A)

            Console.WriteLine("AC")
            o.foo2(6, OPTA:=A, optC:=C)

            Console.WriteLine("CA")
            o.foo2(7, OPTC:=C, optA:=A)

            Console.WriteLine("BC")
            o.foo2(8, OPTB:=B, optC:=C)
            Console.WriteLine("CB")
            o.foo2(9, OPTC:=C, optB:=B)

            'Three arguments
            Console.WriteLine("ABC")
            o.foo2(10, OPTA:=A, optB:=B, optc:=C)

            Console.WriteLine("ACB")
            o.foo2(11, OPTA:=A, optC:=C, optB:=B)

            Console.WriteLine("BAC")
            o.foo2(12, OPTB:=B, optA:=A, optC:=C)

            Console.WriteLine("BCA")
            o.foo2(13, OPTB:=B, optC:=C, optA:=A)

            Console.WriteLine("CAB")
            o.foo2(14, OPTC:=C, opta:=A, optB:=B)

            Console.WriteLine("CBA")
            o.foo2(15, OPTC:=C, optB:=B, optA:=A)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

    Sub LateBoundByRefTests()
        Dim o As Object
        Dim w, x, y, z As Integer

        o = New Class6()

        Console.WriteLine("LateBound ByRef Tests")

        x = 1 : y = 2 : z = 3

        Console.Write("1) ")
        o.Foo1(x, y)
        PassFail(x = y)

        x = 1 : y = 2 : z = 3

        Console.Write("2) ")
        o.Foo2(x, y)
        PassFail(x = y)

        x = 1 : y = 2 : z = 3

        Console.Write("3) ")
        o.Foo3(x, y, z)
        PassFail(x = y AndAlso y = z)

        x = 1 : y = 2 : z = 3

        Console.Write("3a) ")
        o.Foo3((x), (y), (z))
        PassFail(x = 1 AndAlso y = 2 AndAlso z = 3)

        x = 1 : y = 2 : z = 3

        Console.Write("4) ")
        o.Foo4(x, y, z)
        PassFail(x = y AndAlso y = z)

        x = 1 : y = 2 : z = 3

        Console.Write("5) ")
        z = o.Foo5(x, y)
        PassFail(x = y AndAlso y = z)

        x = 1 : y = 2 : z = 3

        Console.Write("6) ")
        z = o.Foo6(x, y)
        PassFail(x = y AndAlso y = z)

        x = 1 : y = 2 : z = 3

        Console.Write("7) ")
        w = o.Foo7(x, y, z)
        PassFail(x = y AndAlso y = z AndAlso w = x)

        x = 1 : y = 2 : z = 3

        Console.Write("8) ")
        w = o.Foo8(x, y, z)
        PassFail(x = y AndAlso y = z AndAlso w = x)

        x = 1 : y = 2 : z = 3

        Console.Write("8a) ")
        w = o.Foo8((x), (y), (z))
        PassFail(x = 1 AndAlso y = 2 AndAlso z = 3)

        Console.Write("9a) ")
        Dim s1 As String = "goodbye"
        Dim s2 As String = "hello"
        w = o.Foo9(s1, s2)
        PassFail(s1 = s2)
    End Sub

    Sub SharedFieldTests()
        Dim Obj As Object
        Dim retVal As Object

        Obj = New Class1()
        Console.WriteLine("* LateBound Private Shared Field Get Test")
        Try
            retVal = Obj.PublicSharedValue
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine("* LateBound Private Shared Field Get Test")
        Try
            retVal = Obj.PrivateSharedValue
            Console.WriteLine("FAILED! expected exception not thrown")
            'UNDONE: Remove the extra catch when the urt changes have been finalized
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine("* LateBound Protected Shared Field Get Test")
        Try
            retVal = obj.ProtectedSharedValue
            Console.WriteLine("FAILED! expected exception not thrown")
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

    Sub DefaultPropertyTests()
        Dim Obj As Object
        Dim retVal As Object

        Obj = New Class2()
        Console.WriteLine("* LateBound Default Property Get Test")
        Try
            retVal = Obj(2)
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine("* LateBound Default Property Set Test")
        Try
            Obj(2) = 12345
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine("* LateBound ReadOnly Default Property Set Test")
        Try
            obj = New Class4()
            Obj(4) = 12345
            Console.WriteLine("FAILED! expected exception not thrown")
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine("* LateBound ReadOnly Default Property Get Test")
        Try
            obj = New Class4()
            retVal = Obj(4)
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine("* LateBound WriteOnly Default Property Get Test")
        Try
            obj = New Class5()
            retVal = Obj(5)
            Console.WriteLine("FAILED! expected exception not thrown")
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.WriteLine("* LateBound WriteOnly Default Property Set Test")
        Try
            obj = New Class5()
            Obj(5) = 12345
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub


    Sub LateBoundFieldGetTests()
        Dim obj As Object
        Dim o As Object

        Console.WriteLine("* LateBound Field Get Tests")
        Try
            obj = New Class1()
            o = obj.ObjectValue
            o = obj.ShortValue
            o = obj.IntegerValue
            o = obj.LongValue
            o = obj.SingleValue
            o = obj.DoubleValue
            o = obj.DateValue
            o = obj.DecimalValue
            o = obj.StringValue
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

    Sub LateBoundFieldSetTests()
        Dim obj As Object
        Dim o As Object

        Console.WriteLine("* LateBound Field Set Tests")
        Try
            obj = New Class1()
            obj.ObjectValue = obj
            obj.ObjectValue = Nothing
            obj.ShortValue = 1L
            obj.ShortValue = Nothing
            obj.IntegerValue = 1L
            obj.IntegerValue = Nothing
            obj.LongValue = 1S
            obj.LongValue = Nothing
            obj.SingleValue = 1
            obj.SingleValue = Nothing
            obj.DoubleValue = 1
            obj.DoubleValue = Nothing
            obj.DateValue = #1/1/2001#
            obj.DateValue = Nothing
            obj.DecimalValue = 1@
            obj.DecimalValue = Nothing
            obj.StringValue = "ABC"
            obj.StringValue = Nothing
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

    Sub LateBoundPrivateFieldGetTest()
        Dim obj As Object
        Dim o As Object

        Console.WriteLine("* LateBound Get of Private Field Test")
        Try
            obj = New Class1()
            o = obj.PrivateValue
            Failed()
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

    Sub LateBoundProtectedPropertyTest()
        Dim obj As Object
        Dim o As Object

        Console.WriteLine("* LateBound Get of Protected Property Test")
        Try
            obj = New Class1()
            o = obj.ProtectedProp
            Failed()
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

    Sub LateBoundPrivateFieldSetTest()
        Dim obj As Object
        Dim o As Object

        Console.WriteLine("* LateBound Set of Private Field Tests")
        Try
            obj = New Class1()
            obj.PrivateValue = Nothing
            Failed()
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

    Sub LateBoundArrayIndexGet1()

        Dim values() As String = New String() {"Microsoft", "Visual", "Basic", "Compiler"}
        Dim obj As Object
        Dim i, j As Integer

        Console.WriteLine("* Test single dimension get")
        Try

            obj = values

            For i = 0 To values.GetUpperBound(0)
                Console.WriteLine(CStr(obj(i)))
            Next i

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

    Sub LateBoundArrayIndexGet2()

        Dim values(,) As String = New String(,) {{"A", "B"}, {"C", "D"}, {"E", "F"}}
        Dim obj As Object
        Dim i, j As Integer

        Console.WriteLine("* Test multiple dimension get")

        Try
            obj = values

            For i = 0 To values.GetUpperBound(0)
                For j = 0 To values.GetUpperBound(1)
                    Console.Write(CStr(obj(i, j)))
                Next j
                Console.WriteLine()
            Next i

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

    Sub LateBoundArrayIndexSet1()

        Dim values() As String = New String() {"Microsoft", "Visual", "Basic", "Compiler"}
        Dim obj As Object
        Dim i, j As Integer
        Dim s As String

        Console.WriteLine("* Test single dimension set")

        Try

            obj = values

            For i = 0 To values.GetUpperBound(0)
                s = "Val" & CStr(i)
                obj(i) = s
                Console.WriteLine(CStr(obj(i)))
            Next i

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub


    Sub LateBoundArrayIndexSet2()

        Dim values(,) As String = New String(,) {{"A", "B"}, {"C", "D"}, {"E", "F"}}
        Dim obj As Object
        Dim i, j As Integer
        Dim s As String

        Console.WriteLine("* Test multiple dimension set")

        Try
            obj = values

            For i = 0 To values.GetUpperBound(0)
                For j = 0 To values.GetUpperBound(1)
                    s = CStr(i) & ", " & CStr(j) & " - "
                    obj(i, j) = s
                    Console.Write(CStr(obj(i, j)))
                Next j
                Console.WriteLine()
            Next i

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub


    Sub LateBoundStringCompare()
        Dim strVar1 As String
        Dim strVar2 As String
        Dim Obj

        strVar1 = "-1"
        strVar2 = "0"
        Obj = "0"

        Console.WriteLine("** LateBound String Compare")
        PassFail(strVar1 <= strVar2)
        PassFail(strVar1 <= Obj)
    End Sub



End Module

Module Bug132982

    Class class1
        Sub foo(ByVal x As class1)

        End Sub
    End Class

    Class class2

    End Class

    Sub Main()

        Dim o1 As class1
        Dim o2 As Object
        Dim c As Object

        Console.Write("Bug 132982: earlybound ")

        Try

            o1 = New class1()
            c = New class2()

            o1.foo(c)
            Failed()
        Catch ex As InvalidCastException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.Write("Bug 132982: latebound ")
        Try
            o2 = New class1()
            c = New class2()

            o2.foo(c)
            Failed()
        Catch ex As InvalidCastException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module

Module Bug149726

    Sub Main()
        Dim Obj2 As Object = New Class1()

        Console.Write("Bug 149726: ")
        Try
            Obj2.MsgBeep()
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module

Module Bug153295
    Sub Main()
        Dim Obj As Object

        Dim cls1 As New Class1()

        Console.Write("Bug 153295: ")
        Try
            cls1.StringValue = New Char() {"H", "E"}

            obj = New Class1()
            obj.StringValue = New Char() {"H", "E"} 'this is causing an unexpected exception
            PassFail(obj.StringValue = "HE")
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module

Module Bug168229

    Sub Main()
        Dim obj, result, expected As Object

        Console.Write("Bug 168229: ")
        Try
            obj = "defghijklm"
            result = UCase(obj)

            Expected = "DEFGHIJKLM"
            PassFail(Result = Expected)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Public Module LateboundCompareBinTest

    Sub TestLateBoundCompareBin()

        Dim o1 As Object
        Dim o2 As Object

        Console.WriteLine("*** TestLateboundCompareBin ***")

        Console.WriteLine("Nothing")

        o1 = Nothing
        o2 = Nothing

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, True)
        TestScenario(o1 <> o2, False)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, False)

        Console.WriteLine("integers")

        o1 = CInt(1)
        o2 = CInt(2)

        TestScenario(o1 < o2, True)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, False)
        TestScenario(o1 > o2, False)

        o1 = CInt(2)

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, True)
        TestScenario(o1 <> o2, False)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, False)

        o1 = CInt(3)

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, False)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, True)

        o1 = Nothing

        TestScenario(o1 < o2, True)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, False)
        TestScenario(o1 > o2, False)

        o1 = CInt(3)
        o2 = Nothing

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, False)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, True)

        Console.WriteLine("Strings")

        o1 = "aaaa"
        o2 = "bbbb"

        TestScenario(o1 < o2, True)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, False)
        TestScenario(o1 > o2, False)

        o1 = "bbbb"

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, True)
        TestScenario(o1 <> o2, False)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, False)

        o1 = "cccc"

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, False)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, True)

        o1 = Nothing

        TestScenario(o1 < o2, True)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, False)
        TestScenario(o1 > o2, False)

        o2 = ""

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, True)
        TestScenario(o1 <> o2, False)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, False)

        o1 = "cccc"
        o2 = Nothing

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, False)
        TestScenario(o1 = o2, False)
        TestScenario(o1 <> o2, True)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, True)

        o1 = ""

        TestScenario(o1 < o2, False)
        TestScenario(o1 <= o2, True)
        TestScenario(o1 = o2, True)
        TestScenario(o1 <> o2, False)
        TestScenario(o1 >= o2, True)
        TestScenario(o1 > o2, False)

        Console.WriteLine("reference cases")

        On Error Resume Next

        Dim b As Boolean

        o1 = New cls1()
        o2 = Nothing

        b = o1 < o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 <= o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 = o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 <> o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 >= o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 > o2 : Console.WriteLine(Err.Description) : Err.Clear()

        o2 = New cls1()

        b = o1 < o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 <= o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 = o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 <> o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 >= o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 > o2 : Console.WriteLine(Err.Description) : Err.Clear()

        o2 = New cls2()

        b = o1 < o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 <= o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 = o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 <> o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 >= o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 > o2 : Console.WriteLine(Err.Description) : Err.Clear()

        o1 = Nothing

        b = o1 < o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 <= o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 = o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 <> o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 >= o2 : Console.WriteLine(Err.Description) : Err.Clear()
        b = o1 > o2 : Console.WriteLine(Err.Description) : Err.Clear()

        On Error GoTo 0

    End Sub

End Module

Module LateBoundCompareTest

    Class cls1
    End Class

    Class cls2
    End Class

    Sub TestScenario(ByVal b As Boolean, ByVal expected As Boolean)
        If b = expected Then
            Console.WriteLine("pass")
        Else
            Console.WriteLine("fail **")
        End If
    End Sub

    Sub TestLateBoundCompare()
        TestLateBoundCompareBin()
        TestLateboundCompareText()
    End Sub

End Module


Module LateBoundStructureTests

    Public Structure foo
        Public CharArrayField() As Char
    End Structure

    Sub Main()
        Console.WriteLine("LateBound Structure tests")
        SetTest()
        GetTest()
    End Sub

    Sub SetTest()
        Dim f As Object = New foo()

        Try
            f.CharArrayField = "ABCDEFG"
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

    Sub GetTest()
        Dim f As foo
        Dim o As Object
        Dim s As String

        Try
            s = "HIJKLMN"
            f.CharArrayField = s
            o = f

            If o.CharArrayField = s Then
                Passed()
            Else
                Failed()
            End If
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug231364

    Delegate Sub foo(ByVal x As Short, ByRef y As Long)
    Sub foo1(ByVal x As Short, Optional ByRef y As Long = 0)
        y = 8 / 2
    End Sub

    Sub Main()
        Console.Write("Bug 231364: ")
        Try
            Dim var As Object

            var = New foo(AddressOf foo1)

            Dim var2 As Long
            var2 = 8
            var.Invoke(10, y:=var2)
            PassFail(var2 = 4)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug231397

    Delegate Sub foo(ByVal x As Short, ByRef y As Long)

    Sub foo1(ByVal x As Short, Optional ByRef y As Long = 0)
        y = 8 / 2
        x = 8 / 2
    End Sub

    Sub Main()
        Console.Write("Bug 231397: ")
        Try
            Dim var As Object
            var = New foo(AddressOf foo1)
            Dim x As Short, y As Long
            x = 16
            y = 8
            var.Invoke(y:=y, x:=x)
            PassFail(x = 16 AndAlso y = 4)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Class Bug233217

    Shared Sub Main()

        Console.WriteLine("Bug 233217")
        Try
            Dim i As Integer

            Dim c As Cls3 = New Cls3()
            i = c.Moo(123)
            Console.Write("   1) ")
            PassFail(i = 1)

            i = c.Moo("Hello")
            Console.Write("   2) ")
            PassFail(i = 2)

            Dim o As Object = New Cls3()
            i = o.Moo(123)
            Console.Write("   3) ")
            PassFail(i = 1)

            i = o.Moo("Hello")
            Console.Write("   4) ")
            PassFail(i = 2)

            '#If BUGBUG
            'Bug 257652
            i = 1
            i = o.Moo()
            Console.Write("   5) ")
            PassFail(i = 2)
            '#End If

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

    Class Cls3

        Overloads Function Moo(ByVal Arg As Integer) As Integer
            Return 1
        End Function

        Overloads Function Moo(ByVal ParamArray PAry() As String) As Integer
            Return 2
        End Function

    End Class

End Class

Module Bug236760

    Structure str1
        Dim x As Single
    End Structure

    Sub Main()
        Console.Write("Bug 236760: ")
        Try
            Dim o As Object = New str1()

            o.x = 4242424242
            PassFail(o.x = 4242424242)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module ByRefLateBoundParamTest

    Class Late1
        Public ReadOnly Property fred() As Integer
            Get
                Return 12
            End Get
        End Property
        Default Public ReadOnly Property wilma(ByVal x As Integer) As Integer
            Get
                Return 114
            End Get
        End Property

        Public Function Blah() As Integer
            Return 5
        End Function

        Public Overloads ReadOnly Property betty(ByVal i As Integer) As Integer
            Get
                Return 6
            End Get
        End Property

        Public Overloads Property betty(ByVal s As String) As Integer
            Get
                Return 7
            End Get
            Set(ByVal Value As Integer)
                Console.WriteLine("Late1::betty(String)")
            End Set
        End Property
    End Class

    Class Late2
        Private _fred As Integer = 100
        Private _wilma() As Integer = {1, 2, 3, 4, 5}
        Private _bettyString As Integer = 11
        Private _bettyInteger As Integer() = {6, 7, 8, 9, 10}

        Public Property fred() As Integer
            Get
                Return _fred
            End Get
            Set(ByVal Value As Integer)
                _fred = Value
            End Set
        End Property

        Default Public Property wilma(ByVal x As Integer) As Integer
            Get
                Return _wilma(x)
            End Get
            Set(ByVal Value As Integer)
                _wilma(x) = Value
            End Set
        End Property

        Public Overloads ReadOnly Property betty(ByVal i As Integer) As Integer
            Get
                Return _bettyInteger(i)
            End Get
        End Property

        Public Overloads Property betty(ByVal s As String) As Integer
            Get
                Return _bettyString
            End Get
            Set(ByVal Value As Integer)
                _bettyString = Value
            End Set
        End Property

        Public Blah As Integer
    End Class

    Sub barney(ByRef x As Integer)
        x = 115
    End Sub

    Sub TestByRefLateBoundParam()

        Console.WriteLine("*** TestByRefLateBoundParam ***")

        Dim obj As Object
        Dim i As Integer

        Try
            Console.Write("   ReadOnly Property: ")
            obj = New Late1()

            i = obj.fred
            barney(obj.fred)
            PassFail(i = obj.fred)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   ReadOnly Default Indexed Property: ")
            i = obj(0)
            barney(obj(0))
            PassFail(i = obj(0))
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   Function: ")
            i = obj.Blah
            barney(obj.Blah)
            PassFail(i = obj.blah)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   Overloaded ReadOnly Indexed Property: ")
            i = obj.betty(1)
            barney(obj.Betty(1))
            PassFail(i = obj.betty(1))
        Catch ex As Exception
            Failed(ex)
        End Try


        Try
            Console.Write("   Overloaded Property: ")
            obj = New Late2()

            barney(obj.fred)
            PassFail(obj.fred = 115)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   Default Indexed Property: ")
            barney(obj(0))
            PassFail(obj(0) = 115)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   Field: ")
            barney(obj.Blah)
            PassFail(obj.blah = 115)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   Overloaded ReadOnly Indexed Property: ")
            i = obj.Betty(1)
            barney(obj.Betty(1))
            PassFail(obj.betty(1) = i)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   Overloaded Indexed Property (2): ")
            Dim c As New Late2()
            Dim o As Object
            o = ""
            barney(c.Betty(o))
            PassFail(c.betty(o) = 115)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub


End Module


Module Bug239809

    Class Bug239809Cls1
    End Class

    Sub Test()
        Console.Write("Bug 239809: ")

        Dim Obj As Object
        Obj = New Bug239809Cls1()
        On Error GoTo errorhandler
        Obj.foo = 10 'this gives err.number = 5 instead of 438 as in VB6
        Console.WriteLine("FAILED: no exception thrown")
        Exit Sub

errorhandler:
        If Err.Number <> 438 Then
            Console.WriteLine("FAILED: Err.Number = " & Err.Number)
        Else
            Passed()
        End If
    End Sub

End Module


Module Bug240263

    Sub Test()
        Dim o As Object = New Bug240263class1()

        Console.Write("Bug 240263: ")
        Try
            o.scen9(xxx:=5)
            Failed()
        Catch ex As InvalidCastException
            Passed()
        End Try

    End Sub

    Class Bug240263class1
        Sub scen9(ByVal ParamArray xxx() As Object)
            Console.WriteLine("Bug240263class1.scen9: FAILED!")
        End Sub
    End Class

End Module


Module Bug240327

    Class Bug240327cls1
        Public Sub foo(ByVal x() As Integer)
            If x.Length <> 1 OrElse x(0) <> 123 Then
                Throw New ArgumentException("FAILED")
            End If
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 240327: ")
        Try
            Dim o As Object = New Bug240327cls1()
            Dim a() As Integer = {123}
            o.foo(a)
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug240872

    Sub Test()

        Console.Write("Bug 240872: ")
        Try
            Dim x As New Collection()
            x.Add("Data", "z"c)

            PassFail(x.Item("z"c) = "Data")

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module



Class Bug224845

    Class c1
        Private i As Integer
        Private Sub foo()

        End Sub
    End Class

    Shared Sub Test()
        Console.WriteLine("*** Bug 224845: ")
        Dim c As Object = New c1()
        Try
            c.i = 1
            Failed()
        Catch e As Exception
            Console.Write("   ")
            PassFail(e.GetType() Is GetType(MissingMemberException))
        End Try

        Try
            c.foo()
        Catch e As Exception
            Console.Write("   ")
            PassFail(e.GetType() Is GetType(MissingMemberException))
        End Try
    End Sub

End Class



Public Module RValueBaseAndOptimisticSetTest

    Class Barney1
        Private Rep As Integer

        Public Property Z() As Integer
            Get
                Return Rep
            End Get
            Set(ByVal Value As Integer)
                Rep = Value
            End Set
        End Property

        Default Public Property Q(ByVal idx As Integer) As Integer
            Get
                Return Rep
            End Get
            Set(ByVal value As Integer)
                Rep = value
            End Set
        End Property

        Public Sub New(ByVal x As Integer)
            Rep = x
        End Sub
    End Class

    Structure Barney2
        Private Rep As Integer

        Public Property Z() As Integer
            Get
                Return Rep
            End Get
            Set(ByVal Value As Integer)
                Console.WriteLine("I shouldn't be called!")
                Rep = Value
            End Set
        End Property

        Default Public Property Q(ByVal idx As Integer) As Integer
            Get
                Return Rep
            End Get
            Set(ByVal value As Integer)
                Console.WriteLine("I shouldn't be called!")
                Rep = value
            End Set
        End Property

        Public Sub New(ByVal x As Integer)
            Rep = x
        End Sub
    End Structure

    Class Barney3
        Private Rep As Integer

        Public ReadOnly Property Z() As Integer
            Get
                Return Rep
            End Get
        End Property

        Default Public ReadOnly Property Q(ByVal idx As Integer) As Integer
            Get
                Return Rep
            End Get
        End Property

        Public Sub New(ByVal x As Integer)
            Rep = x
        End Sub
    End Class

    Structure Barney4
        Private Rep As Integer

        Public ReadOnly Property Z() As Integer
            Get
                Return Rep
            End Get
        End Property

        Default Public ReadOnly Property Q(ByVal idx As Integer) As Integer
            Get
                Return Rep
            End Get
        End Property

        Public Sub New(ByVal x As Integer)
            Rep = x
        End Sub
    End Structure

    Class Fred1
        Private Rep As New Barney1(1001)

        Public Property y() As Barney1
            Get
                Return Rep
            End Get
            Set(ByVal Value As Barney1)
                Rep = Value
            End Set
        End Property

        Default Public Property P(ByVal idx As Integer) As Barney1
            Get
                Return Rep
            End Get
            Set(ByVal Value As Barney1)
                Rep = Value
            End Set
        End Property

    End Class

    Class Fred2
        Private Rep As New Barney2(2001)

        Public Property y() As Barney2
            Get
                Return Rep
            End Get
            Set(ByVal Value As Barney2)
                Rep = Value
            End Set
        End Property

        Default Public Property P(ByVal idx As Integer) As Barney2
            Get
                Return Rep
            End Get
            Set(ByVal Value As Barney2)
                Rep = Value
            End Set
        End Property

    End Class

    Class Fred3
        Private Rep As New Barney3(3001)

        Public Property y() As Barney3
            Get
                Return Rep
            End Get
            Set(ByVal Value As Barney3)
                Rep = Value
            End Set
        End Property

        Default Public Property P(ByVal idx As Integer) As Barney3
            Get
                Return Rep
            End Get
            Set(ByVal Value As Barney3)
                Rep = Value
            End Set
        End Property

    End Class

    Class Fred4
        Private Rep As New Barney4(4001)

        Public Property y() As Barney4
            Get
                Return Rep
            End Get
            Set(ByVal Value As Barney4)
                Rep = Value
            End Set
        End Property

        Default Public Property P(ByVal idx As Integer) As Barney4
            Get
                Return Rep
            End Get
            Set(ByVal Value As Barney4)
                Rep = Value
            End Set
        End Property

    End Class

    Function honk(ByVal x As Integer) As Integer
        Return x
    End Function

    Function bonk(ByRef x As Integer) As Integer
        bonk = x
        x += 1
    End Function

    Public Sub TestRValueBaseAndOptimisticSet()
        Console.WriteLine("*** TestRValueBaseAndOptimisticSet ***")

        Dim x As Object


        Console.WriteLine("ReadWrite Class")
        Try
            x = New Fred1()
            'pass it byref
            Console.WriteLine(bonk(x.y.z))
            Console.WriteLine(bonk(x.y.z))
            Console.WriteLine(bonk(x(10)(10)))
            Console.WriteLine(bonk(x(10)(10)))
        Catch ex As Exception
            Failed(ex)
        End Try

        'pass it byval
        Try
            Console.WriteLine(honk(x.y.z))
            Console.WriteLine(honk(x(10)(10)))
        Catch ex As Exception
        Catch ex As Exception
            Failed(ex)
        End Try

        'assign to it
        Try
            x.y.z = 42
            Console.WriteLine(x.y.z)
            x(10)(10) = 52
            Console.WriteLine(x(10)(10))
        Catch ex As Exception
            Failed(ex)
        End Try


        Console.WriteLine("ReadWrite Structure")
        x = New Fred2()
        Try
            'pass it byref
            Console.WriteLine(bonk(x.y.z))
            Console.WriteLine("Failure - an exception should have been thrown")
        Catch e As Exception
            Console.WriteLine("Success - " & e.Message)
        End Try

        Try
            'pass it byref in With
            With x.y
                Console.WriteLine(bonk(.z))
                Console.WriteLine("Failure - an exception should have been thrown")
            End With
        Catch e As Exception
            Console.WriteLine("Success - " & e.Message)
        End Try

        Try
            'pass it byref
            Console.WriteLine(bonk(x(10)(10)))
            Console.WriteLine("Failure - an exception should have been thrown")
        Catch e As Exception
            Console.WriteLine("Success - " & e.Message)
        End Try

        Try
            'pass it byval
            Console.WriteLine(honk(x.y.z))
            Console.WriteLine(honk(x(10)(10)))
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            'pass it byval in With
            With x.y
                Console.WriteLine(honk(.z))
            End With
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            'assign to it
            x.y.z = 42
            Console.WriteLine("Failure - an exception should have been thrown")
        Catch e As Exception
            Console.WriteLine("Success - " & e.Message)
        End Try

        Try
            'assign to it in With
            With x.y
                .z = 42
                Console.WriteLine("Failure - an exception should have been thrown")
            End With
        Catch e As Exception
            Console.WriteLine("Success - " & e.Message)
        End Try

        Try
            'assign to it
            x(10)(10) = 52
            Console.WriteLine("Failure - an exception should have been thrown")
        Catch e As Exception
            Console.WriteLine("Success - " & e.Message)
        End Try


        Console.WriteLine("ReadOnly Class")
        x = New Fred3()
        'pass it byref
        Console.WriteLine(bonk(x.y.z))
        Console.WriteLine(bonk(x.y.z))
        Console.WriteLine(bonk(x(10)(10)))
        Console.WriteLine(bonk(x(10)(10)))

        'pass it byval
        Console.WriteLine(honk(x.y.z))
        Console.WriteLine(honk(x(10)(10)))

        Try
            'assign to it
            x.y.z = 42
            Console.WriteLine("Failure - an exception should have been thrown")
        Catch e As Exception
            Console.WriteLine("Success - " & e.Message)
        End Try

        Try
            'assign to it
            x(10)(10) = 52
            Console.WriteLine("Failure - an exception should have been thrown")
        Catch e As Exception
            Console.WriteLine("Success - " & e.Message)
        End Try



        Console.WriteLine("ReadOnly Structure")
        Try
            x = New Fred4()
            'pass it byref
            Console.WriteLine(bonk(x.y.z))
            Console.WriteLine(bonk(x.y.z))
            Console.WriteLine(bonk(x(10)(10)))
            Console.WriteLine(bonk(x(10)(10)))
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            'pass it byval
            Console.WriteLine(honk(x.y.z))
            Console.WriteLine(honk(x(10)(10)))
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            'assign to it
            x.y.z = 42
            Console.WriteLine("Failure - an exception should have been thrown")
        Catch e As Exception
            Console.WriteLine("Success - " & e.Message)
        End Try

        Try
            'assign to it
            x(10)(10) = 52
            Console.WriteLine("Failure - an exception should have been thrown")
        Catch e As Exception
            Console.WriteLine("Success - " & e.Message)
        End Try


        Dim o As Object = New Struct1()
        o.i = 10
        With o
            .i = 20
            .z = 42
        End With
        PassFail(o.i = 20)
        PassFail(o.z = 42)

        Try
            With Foo()
                .i = 20
                Failed()
            End With
        Catch e As Exception
            Console.WriteLine("Success - " & e.Message)
        End Try

        Try
            With Foo()
                .z = 20
                Failed()
            End With
        Catch e As Exception
            Console.WriteLine("Success - " & e.Message)
        End Try

    End Sub

    Structure Struct1
        Public i As Integer
        Private _z As Integer
        Public Property Z() As Integer
            Get
                Return _z
            End Get
            Set(ByVal Value As Integer)
                _z = Value
            End Set
        End Property
    End Structure

    Function Foo() As Object
        Return New Struct1()
    End Function

End Module

Namespace ConversionConsumptionTest

    Public Module Driver

        Public Function DynamicCast(Of T)(ByVal o As Object) As T
            Return DirectCast(Conversions.ChangeType(o, GetType(T)), T)
        End Function

        Function AddOne(ByVal s As String) As String
            Return "1" & s
        End Function

        Sub AddTwo(ByRef s As String)
            s = "2" & s
        End Sub

        Sub TestConversionConsumption()
            Console.WriteLine("---- TestConversionConsumption ----")

            Dim x As Complex

            x = DynamicCast(Of Complex)(42)
            Console.WriteLine(DynamicCast(Of String)(x))

            x = DynamicCast(Of Complex)(42 ^ 2)
            Console.WriteLine(DynamicCast(Of String)(x))

            x = DynamicCast(Of Complex)("74088+74088i")
            Console.WriteLine(DynamicCast(Of String)(x))

            x = DynamicCast(Of Complex)(AddOne(DynamicCast(Of String)(x)))
            Console.WriteLine(DynamicCast(Of String)(x))

            Dim s As String = DynamicCast(Of String)(x)
            AddTwo(s)
            Console.WriteLine(DynamicCast(Of String)(DynamicCast(Of Complex)(s)))

            Scenario1.Go()
            Scenario2.Go()
            Scenario3.Go()
            Scenario4.Go()
            Scenario5.Go()
            Scenario6.Go()
            Try
                Scenario7.Go()
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try

            Try
                Scenario8.Go()
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try

            Try
                Scenario9.Go()
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try

            WrapperForBug.C1.Go()
        End Sub
    End Module

    Public Structure Complex

        Private real As Integer
        Private imaginary As Integer

        Public Sub New(ByVal r As Integer, ByVal i As Integer)
            Me.real = r
            Me.imaginary = i
        End Sub

        Shared Widening Operator CType(ByVal x As Complex) As String
            Console.WriteLine("Complex::CType(Complex) As String")
            Return CStr(x.real) & "+" & CStr(x.imaginary) & "i"
        End Operator

        Shared Widening Operator CType(ByVal x As Integer) As Complex
            Console.WriteLine("Complex::CType(Integer) As Complex")
            Return New Complex(x, x)
        End Operator

        Shared Widening Operator CType(ByVal x As Double) As Complex
            Console.WriteLine("Complex::CType(Double) As Complex")
            Return New Complex(CInt(x), CInt(x))
        End Operator

        Shared Widening Operator CType(ByVal x As String) As Complex
            Console.WriteLine("Complex::CType(String) As Complex")
            Dim s As String() = Split(x, "+")
            Dim real As Integer = CInt(s(0))
            Dim imaginary As Integer = 0
            If UBound(s) = 1 Then
                imaginary = CInt(Left(s(1), Len(s(1)) - 1))
            End If
            Return New Complex(real, imaginary)
        End Operator

    End Structure


    Namespace Scenario1
        Public Class SBase
            Shared Narrowing Operator CType(ByVal x As SBase) As TBase
                Console.WriteLine("#6 SBase::CType(SBase) As TBase")
            End Operator
            Shared Narrowing Operator CType(ByVal x As SBase) As T
                Console.WriteLine("#3 SBase::CType(SBase) As T")
            End Operator
            Shared Narrowing Operator CType(ByVal x As SBase) As TDerived
                Console.WriteLine("#4 SBase::CType(SBase) As TDerived")
            End Operator
        End Class

        Public Class S : Inherits SBase
            Overloads Shared Narrowing Operator CType(ByVal x As S) As TBase
                Console.WriteLine("#5 S::CType(S) As TBase")
            End Operator
            Overloads Shared Narrowing Operator CType(ByVal x As S) As T
                Console.WriteLine("#1 S::CType(S) As T")
            End Operator
            Overloads Shared Narrowing Operator CType(ByVal x As S) As TDerived
                Console.WriteLine("#2 S::CType(S) As TDerived")
            End Operator
        End Class

        Public Class SDerived : Inherits S
        End Class

        Public Class TBase
            Shared Narrowing Operator CType(ByVal x As SDerived) As TBase
                Console.WriteLine("#8 TBase::CType(SDerived) As TBase")
            End Operator
        End Class

        Public Class T : Inherits TBase
            Overloads Shared Narrowing Operator CType(ByVal x As SDerived) As T
                Console.WriteLine("#7 T::CType(SDerived) As T")
            End Operator
        End Class

        Public Class TDerived : Inherits T
        End Class

        Public Module Driver
            Sub Go()
                Console.Write("Scenario 1: ")
                Dim x As S = New S
                Dim y As T = DynamicCast(Of T)(x)
            End Sub
        End Module
    End Namespace


    Namespace Scenario2
        Public Class SBase
            Shared Narrowing Operator CType(ByVal x As SBase) As TBase
                Console.WriteLine("#6 SBase::CType(SBase) As TBase")
            End Operator
            'Shared Narrowing Operator CType(ByVal x As SBase) As T
            '    Console.WriteLine("#3 SBase::CType(SBase) As T")
            'End Operator
            Shared Narrowing Operator CType(ByVal x As SBase) As TDerived
                Console.WriteLine("#4 SBase::CType(SBase) As TDerived")
            End Operator
        End Class

        Public Class S : Inherits SBase
            Overloads Shared Narrowing Operator CType(ByVal x As S) As TBase
                Console.WriteLine("#5 S::CType(S) As TBase")
            End Operator
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As T
            '    Console.WriteLine("#1 S::CType(S) As T")
            'End Operator
            Overloads Shared Narrowing Operator CType(ByVal x As S) As TDerived
                Console.WriteLine("#2 S::CType(S) As TDerived")
            End Operator
        End Class

        Public Class SDerived : Inherits S
        End Class

        Public Class TBase
            Shared Narrowing Operator CType(ByVal x As SDerived) As TBase
                Console.WriteLine("#8 TBase::CType(SDerived) As TBase")
            End Operator
        End Class

        Public Class T : Inherits TBase
            'Overloads Shared Narrowing Operator CType(ByVal x As SDerived) As T
            '    Console.WriteLine("#7 T::CType(SDerived) As T")
            'End Operator
        End Class

        Public Class TDerived : Inherits T
        End Class

        Public Module Driver
            Sub Go()
                Console.Write("Scenario 2: ")
                Dim x As S = New S
                Dim y As T = DynamicCast(Of T)(x)
            End Sub
        End Module
    End Namespace


    Namespace Scenario3
        Public Class SBase
            Shared Narrowing Operator CType(ByVal x As SBase) As TBase
                Console.WriteLine("#6 SBase::CType(SBase) As TBase")
            End Operator
            Shared Narrowing Operator CType(ByVal x As SBase) As T
                Console.WriteLine("#3 SBase::CType(SBase) As T")
            End Operator
            Shared Narrowing Operator CType(ByVal x As SBase) As TDerived
                Console.WriteLine("#4 SBase::CType(SBase) As TDerived")
            End Operator
        End Class

        Public Class S : Inherits SBase
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As TBase
            '    Console.WriteLine("#5 S::CType(S) As TBase")
            'End Operator
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As T
            '    Console.WriteLine("#1 S::CType(S) As T")
            'End Operator
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As TDerived
            '    Console.WriteLine("#2 S::CType(S) As TDerived")
            'End Operator
        End Class

        Public Class SDerived : Inherits S
        End Class

        Public Class TBase
            Shared Narrowing Operator CType(ByVal x As SDerived) As TBase
                Console.WriteLine("#8 TBase::CType(SDerived) As TBase")
            End Operator
        End Class

        Public Class T : Inherits TBase
            Overloads Shared Narrowing Operator CType(ByVal x As SDerived) As T
                Console.WriteLine("#7 T::CType(SDerived) As T")
            End Operator
        End Class

        Public Class TDerived : Inherits T
        End Class

        Public Module Driver
            Sub Go()
                Console.Write("Scenario 3: ")
                Dim x As S = New S
                Dim y As T = DynamicCast(Of T)(x)
            End Sub
        End Module
    End Namespace


    Namespace Scenario4
        Public Class SBase
            Shared Narrowing Operator CType(ByVal x As SBase) As TBase
                Console.WriteLine("#6 SBase::CType(SBase) As TBase")
            End Operator
            'Shared Narrowing Operator CType(ByVal x As SBase) As T
            '    Console.WriteLine("#3 SBase::CType(SBase) As T")
            'End Operator
            Shared Narrowing Operator CType(ByVal x As SBase) As TDerived
                Console.WriteLine("#4 SBase::CType(SBase) As TDerived")
            End Operator
        End Class

        Public Class S : Inherits SBase
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As TBase
            '    Console.WriteLine("#5 S::CType(S) As TBase")
            'End Operator
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As T
            '    Console.WriteLine("#1 S::CType(S) As T")
            'End Operator
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As TDerived
            '    Console.WriteLine("#2 S::CType(S) As TDerived")
            'End Operator
        End Class

        Public Class SDerived : Inherits S
        End Class

        Public Class TBase
            Shared Narrowing Operator CType(ByVal x As SDerived) As TBase
                Console.WriteLine("#8 TBase::CType(SDerived) As TBase")
            End Operator
        End Class

        Public Class T : Inherits TBase
            'Overloads Shared Narrowing Operator CType(ByVal x As SDerived) As T
            '    Console.WriteLine("#7 T::CType(SDerived) As T")
            'End Operator
        End Class

        Public Class TDerived : Inherits T
        End Class

        Public Module Driver
            Sub Go()
                Console.Write("Scenario 4: ")
                Dim x As S = New S
                Dim y As T = DynamicCast(Of T)(x)
            End Sub
        End Module
    End Namespace

    Namespace Scenario5
        Public Class SBase
            Shared Narrowing Operator CType(ByVal x As SBase) As TBase
                Console.WriteLine("#6 SBase::CType(SBase) As TBase")
            End Operator
            'Shared Narrowing Operator CType(ByVal x As SBase) As T
            '    Console.WriteLine("#3 SBase::CType(SBase) As T")
            'End Operator
            'Shared Narrowing Operator CType(ByVal x As SBase) As TDerived
            '    Console.WriteLine("#4 SBase::CType(SBase) As TDerived")
            'End Operator
        End Class

        Public Class S : Inherits SBase
            Overloads Shared Narrowing Operator CType(ByVal x As S) As TBase
                Console.WriteLine("#5 S::CType(S) As TBase")
            End Operator
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As T
            '    Console.WriteLine("#1 S::CType(S) As T")
            'End Operator
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As TDerived
            '    Console.WriteLine("#2 S::CType(S) As TDerived")
            'End Operator
        End Class

        Public Class SDerived : Inherits S
        End Class

        Public Class TBase
            Shared Narrowing Operator CType(ByVal x As SDerived) As TBase
                Console.WriteLine("#8 TBase::CType(SDerived) As TBase")
            End Operator
        End Class

        Public Class T : Inherits TBase
            'Overloads Shared Narrowing Operator CType(ByVal x As SDerived) As T
            '    Console.WriteLine("#7 T::CType(SDerived) As T")
            'End Operator
        End Class

        Public Class TDerived : Inherits T
        End Class

        Public Module Driver
            Sub Go()
                Console.Write("Scenario 5: ")
                Dim x As S = New S
                Dim y As T = DynamicCast(Of T)(x)
            End Sub
        End Module
    End Namespace


    Namespace Scenario6
        Public Class SBase
            Shared Narrowing Operator CType(ByVal x As SBase) As TBase
                Console.WriteLine("#6 SBase::CType(SBase) As TBase")
            End Operator
            'Shared Narrowing Operator CType(ByVal x As SBase) As T
            '    Console.WriteLine("#3 SBase::CType(SBase) As T")
            'End Operator
            'Shared Narrowing Operator CType(ByVal x As SBase) As TDerived
            '    Console.WriteLine("#4 SBase::CType(SBase) As TDerived")
            'End Operator
        End Class

        Public Class S : Inherits SBase
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As TBase
            '    Console.WriteLine("#5 S::CType(S) As TBase")
            'End Operator
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As T
            '    Console.WriteLine("#1 S::CType(S) As T")
            'End Operator
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As TDerived
            '    Console.WriteLine("#2 S::CType(S) As TDerived")
            'End Operator
        End Class

        Public Class SDerived : Inherits S
        End Class

        Public Class TBase
            Shared Narrowing Operator CType(ByVal x As SDerived) As TBase
                Console.WriteLine("#8 TBase::CType(SDerived) As TBase")
            End Operator
        End Class

        Public Class T : Inherits TBase
            'Overloads Shared Narrowing Operator CType(ByVal x As SDerived) As T
            '    Console.WriteLine("#7 T::CType(SDerived) As T")
            'End Operator
        End Class

        Public Class TDerived : Inherits T
        End Class

        Public Module Driver
            Sub Go()
                Console.Write("Scenario 6: ")
                Dim x As S = New S
                Dim y As T = DynamicCast(Of T)(x)
            End Sub
        End Module
    End Namespace


    Namespace Scenario7
        Public Class SBase
            'Shared Narrowing Operator CType(ByVal x As SBase) As TBase
            '    Console.WriteLine("#6 SBase::CType(SBase) As TBase")
            'End Operator
            'Shared Narrowing Operator CType(ByVal x As SBase) As T
            '    Console.WriteLine("#3 SBase::CType(SBase) As T")
            'End Operator
            'Shared Narrowing Operator CType(ByVal x As SBase) As TDerived
            '    Console.WriteLine("#4 SBase::CType(SBase) As TDerived")
            'End Operator
        End Class

        Public Class S : Inherits SBase
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As TBase
            '    Console.WriteLine("#5 S::CType(S) As TBase")
            'End Operator
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As T
            '    Console.WriteLine("#1 S::CType(S) As T")
            'End Operator
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As TDerived
            '    Console.WriteLine("#2 S::CType(S) As TDerived")
            'End Operator
        End Class

        Public Class SDerived : Inherits S
        End Class

        Public Class TBase
            Shared Narrowing Operator CType(ByVal x As SDerived) As TBase
                Console.WriteLine("#8 TBase::CType(SDerived) As TBase")
            End Operator
        End Class

        Public Class T : Inherits TBase
            Overloads Shared Narrowing Operator CType(ByVal x As SDerived) As T
                Console.WriteLine("#7 T::CType(SDerived) As T")
            End Operator
        End Class

        Public Class TDerived : Inherits T
        End Class

        Public Module Driver
            Sub Go()
                Console.Write("Scenario 7: ")
                Dim x As S = New S
                Dim y As T = DynamicCast(Of T)(x)
            End Sub
        End Module
    End Namespace


    Namespace Scenario8
        Public Class SBase
            'Shared Narrowing Operator CType(ByVal x As SBase) As TBase
            '    Console.WriteLine("#6 SBase::CType(SBase) As TBase")
            'End Operator
            'Shared Narrowing Operator CType(ByVal x As SBase) As T
            '    Console.WriteLine("#3 SBase::CType(SBase) As T")
            'End Operator
            'Shared Narrowing Operator CType(ByVal x As SBase) As TDerived
            '    Console.WriteLine("#4 SBase::CType(SBase) As TDerived")
            'End Operator
        End Class

        Public Class S : Inherits SBase
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As TBase
            '    Console.WriteLine("#5 S::CType(S) As TBase")
            'End Operator
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As T
            '    Console.WriteLine("#1 S::CType(S) As T")
            'End Operator
            'Overloads Shared Narrowing Operator CType(ByVal x As S) As TDerived
            '    Console.WriteLine("#2 S::CType(S) As TDerived")
            'End Operator
        End Class

        Public Class SDerived : Inherits S
        End Class

        Public Class TBase
            Shared Narrowing Operator CType(ByVal x As SDerived) As TBase
                Console.WriteLine("#8 TBase::CType(SDerived) As TBase")
            End Operator
        End Class

        Public Class T : Inherits TBase
            'Overloads Shared Narrowing Operator CType(ByVal x As SDerived) As T
            '    Console.WriteLine("#7 T::CType(SDerived) As T")
            'End Operator
        End Class

        Public Class TDerived : Inherits T
        End Class

        Public Module Driver
            Sub Go()
                Console.Write("Scenario 8: ")
                Dim x As S = New S
                Dim y As T = DynamicCast(Of T)(x)
            End Sub
        End Module
    End Namespace

    Namespace Scenario9
        Public Class SBase
            Shared Widening Operator CType(ByVal x As SBase) As TBase
                Console.WriteLine("#6 SBase::CType(SBase) As TBase")
            End Operator
        End Class

        Public Class S : Inherits SBase
        End Class

        Public Class TBase
            Shared Widening Operator CType(ByVal x As SBase) As TBase
                Console.WriteLine("#6 TBase::CType(SBase) As TBase")
            End Operator
        End Class

        Public Class T : Inherits TBase
        End Class

        Public Module Driver
            Sub Go()
                Console.Write("Scenario 9: ")
                Dim x As S = New S
                Dim y As T = DynamicCast(Of T)(x)
            End Sub
        End Module
    End Namespace

    Namespace WrapperForBug

        Public Class C1

            Private Class NestedPrivClass
                Shared Operator +(ByVal y As NestedPrivClass) As Integer
                    Console.WriteLine("NestedPrivClass::+(NestedPrivClass) As Integer")
                End Operator
            End Class

            Shared Sub Go()
                NestedPrivClass.op_UnaryPlus(New NestedPrivClass)
            End Sub

        End Class

    End Namespace

End Namespace


Namespace TestHarness

    Module TestHarness
        Sub Failed(ByVal ex As Exception)
            If ex Is Nothing Then
                Console.WriteLine("NULL System.Exception")
            Else
                Console.WriteLine(ex.GetType().FullName)
            End If
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine("FAILED !!!")
        End Sub

        Sub Failed()
            Console.WriteLine("FAILED !!!")
        End Sub

        Sub Passed()
            Console.WriteLine("passed")
        End Sub

        Sub PassFail(ByVal bPassed As Boolean)
            If bPassed Then
                Console.WriteLine("passed")
            Else
                Console.WriteLine("FAILED !!!")
            End If
        End Sub
    End Module
End Namespace


Class Bug257652

    Shared m_x As Integer()

    Shared Sub Test()
        Dim o As Object = New Bug257652class1()

        Console.WriteLine("Bug 257652")
        Try
            Console.Write("    1) ")
            o.foo1()
            PassFail(m_x.Length = 0)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("    2) ")
            m_x = Nothing
            o.foo2("hello")
            PassFail(m_x.Length = 0)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

    Class Bug257652Class1
        Sub foo1(ByVal ParamArray x() As Integer)
            m_x = x
        End Sub

        Sub foo2(ByVal s As String, ByVal ParamArray x() As Integer)
            m_x = x
        End Sub

    End Class

End Class

Module Bug227370
    Class Bug227370Cls1
        ReadOnly Property Foo(ByVal Arg As Integer) As Integer
            Get

            End Get
        End Property

        Property Foo2(ByVal Arg As Integer, ByVal arg2 As Integer) As Integer
            Get

            End Get
            Set(ByVal Value As Integer)
            End Set
        End Property
    End Class

    Sub Test()
        Console.WriteLine("*** Bug227370")
        Dim Obj As Object = New Bug227370Cls1()

        Try
            Console.Write("    1) ")
            Obj.Foo(10) = 10
            Failed()
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("    2) ")
            Obj.Foo2(10) = 10
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Class Bug256623
    Class Cls1
        Sub foo(ByVal Arg As Short)
        End Sub

        Sub foo(ByVal Arg As String)
        End Sub
    End Class

    Class Cls2
        Sub foo(ByVal Arg As Short, ByVal Arg2 As String)
        End Sub

        Sub foo(ByVal Arg As String, ByVal Arg2 As Short)
        End Sub
    End Class

    Shared Sub Test()
        Dim Obj As Object
        Dim i As Integer
        Dim c2 As Cls2
        Dim c As New Cls1()
        Dim s As String

        Console.WriteLine("*** Bug 256623")
        Try
            c2 = New Cls2()
            c = New Cls1()
            obj = c
            i = 10
            Console.Write("   1) ")
            Obj.foo(i)
            PassFail(False)
        Catch ex As Exception
            PassFail(InStr(ex.Message, ") As ") = 0)
            Console.WriteLine("   " & ex.Message)
        End Try


        Try
            Console.Write("   2) ")

            c2 = New Cls2()
            obj = c2
            i = 10
            Obj.foo(i, s)
            PassFail(False)
        Catch ex As Exception
            PassFail(InStr(ex.Message, ") As ") = 0)
            Console.WriteLine("   " & ex.Message)
        End Try
    End Sub
End Class


Class Bug257437
    Shared Result As Integer

    Shared Sub foo(ByVal i As Integer, ByVal b As Byte)
        Result = 1
    End Sub

    Shared Sub foo(ByVal i As Integer, ByVal b As Int16)
        Result = 2
    End Sub

    Shared Sub foo(ByVal i As Integer, ByVal b As Int32)
        Result = 3
    End Sub

    Shared Sub foo(ByVal i As Integer, ByVal b As String, Optional ByVal x As Integer = 1)
        Result = 4
    End Sub

    Shared Sub Test()
        Console.WriteLine("*** Bug 257437")
        Try
            Dim fnum

            Console.Write("   1) ")
            foo(fnum, CByte(255))
            PassFail(Result = 1)

            Console.Write("   2) ")
            foo(fnum, -1S)
            PassFail(Result = 2)

            Console.Write("   3) ")
            foo(fnum, -1I)
            PassFail(Result = 3)

            Console.Write("   4) ")
            foo(fnum, "abc")
            PassFail(Result = 4)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Class


Class Bug256635
    Shared Result As Integer
    Class Cls1
        Sub foo(ByVal Arg As Short)
            Result = 1
        End Sub

        Sub foo(ByVal Arg As UInt16)
            Result = 2
        End Sub
    End Class

    Shared Sub Test()
        Dim Obj As Object = New Cls1()
        Console.Write("*** Bug VS#256635 VSW#236186: ")
        Try
            Obj.foo(10)
            Failed()
        Catch ex As AmbiguousMatchException
            Passed()
        End Try
    End Sub
End Class


Class Bug257640

    Shared Result As Integer

    Class cls1
        Sub foo(ByVal ParamArray x() As Integer)
            Result = 1
        End Sub
        Sub foo(ByVal x As Integer, Optional ByVal z As Integer = 2)
            Result = 2
        End Sub
    End Class

    Shared Sub Test()
        Dim c As New Cls1()
        Dim x As Object = c
        Dim EarlyResult As Integer

        Console.Write("*** Bug 257640: ")
        Try
            c.foo(1)
            EarlyResult = Result

            Result = 0
            x.foo(1)
            PassFail(Result = EarlyResult)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Class


Module Bug166464
    Class Bug166464Class

        Public ReadOnly Property foo() As Short
            Get
                Return 1
            End Get
        End Property

        Public ReadOnly Property foo2() As Short
            Get
                Return 2
            End Get
        End Property

        Public ReadOnly Property abc() As Short
            Get
                Return 3
            End Get
        End Property

    End Class


    Sub Test()

        Console.WriteLine("Regression test Bug166464")

        Dim x As Bug166464Class = New Bug166464Class()
        Dim o1, o2 As Object
        Dim i As Integer

        o1 = x
        Try
            o2 = Interaction.CallByName(o1, "", Microsoft.VisualBasic.CallType.Get)
            Console.WriteLine("o2 = " & CStr(o2))
            Console.WriteLine("Bug166464: FAILED")
        Catch ex As Exception
            Console.WriteLine("Bug166464: PASSED")
        End Try

    End Sub
End Module


Module Bug126622

    Sub Test()

        Console.WriteLine("*** Bug 126622")

        Dim coll As Object

        coll = New Collection()

        Try
            Console.Write("   1) ")
            coll.Add("somedata1", "somekey1")
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   2) ")
            coll.Add("somedata2", "somekey2")
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   3) ")
            coll.Add("somedata3", "somekey3", Nothing, Nothing)
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module

Module Bug168135
    'This tests setting/getting of array fields
    Sub Test()
        Dim bFailed As Boolean

        Console.WriteLine("Regression test Bug168135")
        Try
            Dim o As Object = New Class1()
            If o.Ary(0) <> 6 Then
                Console.WriteLine("Bug168135: FAILED step 1a")
                bFailed = True
            End If
            o.Ary(4) = 11  'this is causing an unexpected MissingMethodException
            If o.Ary(4) <> 11 Then
                Console.WriteLine("Bug168135: FAILED step 1b")
                bFailed = True
            End If
        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name & ": " & ex.Message)
            Console.WriteLine("Bug168135: FAILED step 1c")
            bFailed = True
        End Try

        Try
            Dim o As Object = New Class1()
            If o.ObjectValue(0) <> 1 Then
                Console.WriteLine("Bug168135: FAILED step 2a")
                bFailed = True
            End If
            o.ObjectValue(4) = 6  'this is causing an unexpected MissingMethodException
            If o.ObjectValue(4) <> 6 Then
                Console.WriteLine("Bug168135: FAILED step 2b")
                bFailed = True
            End If
        Catch ex As Exception
            Failed(ex)
            Console.WriteLine("Bug168135: FAILED step 2c")
            bFailed = True
        End Try

        Try
            Dim o As Object = New Class1()
            If o(0) <> "A" Then
                Console.WriteLine("Bug168135: FAILED step 3a")
                bFailed = True
            End If
            o(4) = "X"  'this is causing an unexpected MissingMethodException
            If o(4) <> "X" Then
                Console.WriteLine("Bug168135: FAILED step 3b")
                bFailed = True
            End If
        Catch ex As Exception
            Failed(ex)
            Console.WriteLine("Bug168135: FAILED step 3c")
            bFailed = True
        End Try

        If Not bFailed Then
            Console.WriteLine("Bug168135: PASSED")
        End If
    End Sub
End Module

Class Bug257644

    Class cls1
        Property foo(ByVal ParamArray x() As Integer)
            Get

            End Get
            Set(ByVal Value)

            End Set
        End Property
    End Class

    Shared Sub Test()

        Console.Write("*** Bug 257644: ")
        Try
            Dim x As Object = New cls1()
            x.foo(1) = 1
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Class


Class Bug258409

    Shared m_result As String
    Class cls1
        Overloads Property P1(ByVal x As Integer) As String
            Get

            End Get
            Set(ByVal Value As String)
                m_result = "String"
            End Set
        End Property
        Overloads Property P1(ByVal x As Short) As Short
            Get

            End Get
            Set(ByVal Value As Short)
                m_result = "Short"
            End Set
        End Property
    End Class

    Shared Sub Test()

        Console.WriteLine("*** Bug258409: ")
        Try
            Dim x As Object = New cls1()
            Dim a As Short = "1234"
            Dim b As String = "moo"

            Console.Write("    1) ")
            x.P1(1) = a
            PassFail(m_result = "String")

            Console.Write("    2) ")
            x.P1(1) = b
            PassFail(m_result = "String")
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Class

Class Bug263543
    Shared m_result As Double

    Class cls1
        Sub foo(ByVal x As String)
            m_result = x
        End Sub
    End Class

    Shared Sub Test()
        Console.Write("*** Bug 263543: ")

        Try
            Dim x As cls1 = New cls1()
            Dim y As Object = x
            y.foo(1234.1234)
            PassFail(m_result = 1234.1234)
        Catch Ex As Exception
            Failed(ex)
        End Try
    End Sub

End Class

Class Bug265345

    Class Class1
        Sub sub1(ByRef P1 As Integer, Optional ByVal P2 As Integer = 1)

        End Sub
    End Class

    Shared Sub Test()
        Console.Write("*** Bug 265345: ")
        Dim x As Object

        Try
            x = New Class1()
            x.sub1(p1:=1)
        Catch Ex As Exception
            Failed(ex)
        End Try

    End Sub
End Class

Class Bug264334

    Shared Sub Test()

        Console.Write("*** Bug 264334: ")
        Dim o As Object = 1

        Try
            Console.WriteLine(o.a)
            Failed()
        Catch ex As MissingMemberException
            Passed()
        Catch Ex As Exception
            Failed(ex)
        End Try
    End Sub

End Class

Class Bug236108
    Class c1
        Public m_value1 As Integer
        Default Property foo(ByVal i As Integer) As Integer
            Get
                Return m_value1
            End Get
            Set(ByVal Value As Integer)
                m_value1 = Value
            End Set
        End Property
    End Class

    Class c2
        Inherits c1
        Public m_value2 As Integer
        Shadows Property foo(ByVal i As Integer) As Integer
            Get
                Return m_value2
            End Get
            Set(ByVal Value As Integer)
                m_value2 = Value
            End Set
        End Property
    End Class

    Shared Sub Test()
        Dim c As New c2()
        Dim o As Object = c

        Console.WriteLine("*** Bug 236108")
        Try
            Console.Write("   1) ")
            c(1) = 123
            passfail(c.m_value1 = 123)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   2) ")
            passfail(c(1) = 123)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   3) ")
            o(1) = 456
            passfail(c.m_value1 = 456)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   4) ")
            passfail(o(1) = 456)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub


End Class


Class Bug259716
    Class c1
        Public Function foo() As Integer
        End Function
    End Class

    Class c2
        Inherits c1
        Public Shadows foo As Integer

    End Class

    Shared Sub Test()
        Console.Write("*** Bug 259716: ")

        Dim c As Object = New c2()
        Try
            c.foo()
            Failed()
        Catch ex As ArgumentException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Class



Class Bug266851

    Private Shared m_Value As Integer

    Overloads Shared Sub Foo(ByVal Arg As Byte)
        m_Value = 1
    End Sub

    Overloads Shared Sub Foo(ByVal Arg As Integer)
        m_Value = 2
    End Sub

    Overloads Shared Sub Foo(ByVal Arg As Long)
        m_Value = 3
    End Sub

    Overloads Shared Sub Foo(ByVal Arg As Single)
        m_Value = 4
    End Sub

    Shared Sub Test()

        Console.Write("*** Bug 266851: ")
        Dim Obj As Object
        Dim EarlyResult As Integer

        Try
            Foo(Nothing)
            EarlyResult = m_Value

            Obj = Nothing

            m_Value = 0
            Foo(Obj)
            PassFail(EarlyResult = m_Value)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Class


Class Bug266851a

    Private Shared m_Value As Integer

    Overloads Shared Sub Foo(ByVal Arg As Byte)
        m_Value = 1
    End Sub

    Overloads Shared Sub Foo(ByVal Arg As Integer)
        m_Value = 2
    End Sub

    Overloads Shared Sub Foo(ByVal Arg As Long)
        m_Value = 3
    End Sub

    Overloads Shared Sub Foo(ByVal Arg As String)
        m_Value = 4
    End Sub

    Shared Sub Test()

        Console.Write("*** Bug 266851a: ")
        Dim Obj As Object
        Dim EarlyResult As Integer

        Try
            'Early bound gives expected ambiguous match
            '            Foo(Nothing)

            Obj = Nothing

            Foo(Obj)
            Failed()
        Catch ex As Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Class


Class bug266843

    Overloads Shared Sub Foo(ByVal Arg As Long)
    End Sub

    Class Cls1
    End Class

    Shared Sub test()
        Dim Obj As Object = 10

        Console.WriteLine("*** Bug 266843")
        Console.Write("   1) INT: ")
        Try
            Foo(Obj)
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.Write("   2) Ref: ")
        Try
            Obj = New Cls1()
            Foo(Obj)
        Catch ex As InvalidCastException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.Write("   3) DBNULL: ")
        Try
            Obj = DBNull.Value  'System.Reflection.AmbiguousMatchException thrown here instead of System.ArgumentException
            Foo(Obj)
        Catch ex As InvalidCastException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
        'Foo(1D)
    End Sub

End Class



Module DecimalDateOptionalDefaults

    Private m_dtValue As Date
    Private m_decValue As Decimal

    Class Class1

        Sub FooDecimal(Optional ByVal x As Decimal = 123.45@)
            m_decValue = x
        End Sub

        Sub FooDate(Optional ByVal x As DateTime = #10/12/1998 10:57:00 AM#)
            m_dtValue = x
        End Sub
    End Class

    Sub Test()
        Dim c As New Class1()
        Dim o As Object
        Dim decValue As Decimal
        Dim dtValue As Date

        o = c

        Try
            Console.Write("DateTime Default: ")

            c.FooDate()
            dtValue = m_dtValue

            'Clear value
            m_dtValue = New DateTime()

            o.FooDate()
            PassFail(dtValue = m_dtValue)
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.Write("Decimal Default: ")
        Try
            c.FooDecimal()
            decValue = m_decValue

            'Clear value
            m_decValue = 0@
            o.FooDecimal()
            PassFail(decValue = m_decValue)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug264079

    Private m_arr As Integer()

    Private Class cls1
        Public Sub sub1(ByVal ParamArray arr() As Integer)
            m_arr = arr
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 264079: ")
        Try
            Dim o As Object = New cls1()
            o.sub1()
            PassFail(m_arr.Length = 0)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub


End Module


Module Bug271007
    Private m_arr As Integer()
    Private Class cls1
        Sub foo(ByVal ParamArray x() As Integer)
            m_arr = x
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 271007: ")

        Try
            Dim x As New cls1()
            Dim y As Object = New cls1()
            Dim arr1, arr2 As Integer()

            x.foo(1, 2)
            arr1 = m_arr
            y.foo(1, 2)
            arr2 = m_arr
            PassFail((arr1.Length = arr2.Length) AndAlso _
                     (arr1(0) = arr2(0)) AndAlso _
                     (arr1(1) = arr2(1)))
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module

Module Bug271373
    Private m_i As Integer

    Private Class cls2
        Property P1(ByVal x As Integer) As Integer
            Get
                m_i = 1
            End Get
            Set(ByVal Value As Integer)
                m_i = 2
            End Set
        End Property

        Property P1(ByVal ParamArray x() As Integer) As Integer
            Get
                m_i = 3
            End Get
            Set(ByVal Value As Integer)
                m_i = 4
            End Set
        End Property

        Property P2(ByVal a As cls2, ByVal x As Integer()) As Integer
            Get
                m_i = 1
            End Get
            Set(ByVal Value As Integer)
                m_i = 2
            End Set
        End Property

        Property P2(ByVal a As String, ByVal ParamArray x() As Integer) As Integer
            Get
                m_i = 3
            End Get

            Set(ByVal Value As Integer)
                m_i = 4
            End Set
        End Property

    End Class

    Sub Test()

        Console.Write("Bug 271373: ")

        Try
            Dim iEarly As Integer
            Dim c As New Cls2()
            Dim o As Object = c

            'Should now give compiler error 
            'm_i = -1
            'c.P1(Nothing) = 1
            'iEarly = m_i

            m_i = -1
            o.P1(Nothing) = 1
            Failed()

        Catch ex As Reflection.AmbiguousMatchException
            Passed()

        Catch ex As Exception
            Failed(ex)
        End Try


        Console.Write("Bug 271373b: ")

        Try
            Dim iEarly As Integer
            Dim c As New cls2()
            Dim o As Object = c

            'Should now give compiler error 
            'm_i = -1
            'c.P2(Nothing, Nothing) = 1
            'iEarly = m_i

            m_i = -1
            o.P2(Nothing, Nothing) = 1
            Failed()

        Catch ex As Reflection.AmbiguousMatchException
            Passed()

        Catch ex As Exception
            Failed(ex)
        End Try


        Console.Write("Bug 271373c: ")

        Try
            Dim iEarly As Integer
            Dim c As New cls2()
            Dim o As Object = c

            'Should now give compiler error 
            'm_i = -1
            'c.P2(Nothing, New Integer() {1, 2}) = 1
            'iEarly = m_i

            m_i = -1
            o.P2(Nothing, New Integer() {1, 2}) = 1
            Failed()

        Catch ex As Reflection.AmbiguousMatchException
            Passed()

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module



Module Bug271377

    Private m_x As Integer()
    Private m_Value As Integer

    Private Class cls2
        Property P1(ByVal ParamArray x() As Integer) As Integer
            Get
                m_x = x
            End Get

            Set(ByVal Value As Integer)
                m_x = x
                m_Value = Value
            End Set
        End Property
    End Class

    Sub Test()

        Console.Write("Bug 271377: ")

        Try
            Dim iEarly As Integer
            Dim arrEarly As Integer()

            Dim o As Object = New cls2()
            o.P1() = 1
            PassFail(m_x.Length = 0 AndAlso m_Value = 1)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module


Module Bug271685

    Private m_arr As Object()

    Sub Test()
        Dim iRetCode As Int32  'only 10 means pass
        Dim obj1 As Object

        iRetCode = 11
        Console.Write("Bug 271685: ")

        Try
            obj1 = New C1()

            iRetCode = obj1.foo(1, 2, "SDF")

            PassFail(iRetCode = 10 AndAlso m_arr.Length = 3 AndAlso m_arr(0) = 1 AndAlso m_arr(1) = 2 AndAlso m_arr(2) = "SDF")

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub


    '-------------  class  ------------------------
    Private Class C1

        Public Function foo(ByVal ParamArray parmArr()) As Integer

            m_arr = parmArr
            foo = 10

        End Function

    End Class


End Module



Module Bug277488
    Private m_x As Object

    Private Class cls1
        Function foo(ByVal ParamArray x() As Short) As Integer
            m_x = x
        End Function
        Function foo(ByVal ParamArray x() As Integer) As Short
            m_x = x
        End Function
    End Class


    Sub Test()
        Console.Write("Bug 277488: ")
        Try
            Dim o As Object = New cls1()

            o.foo(1.1)
            Failed()
        Catch ex As System.Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug278308
    Private m_i As Integer

    Private Class cls1
        Public Sub foo(ByVal ParamArray x() As cls1)
            m_i = x.Length
        End Sub
    End Class

    Sub Test()

        Console.Write("Bug 278308: ")
        Try
            Dim x As Object = New cls1()

            x.foo(x, x)
            PassFail(m_i = 2)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module

Module Bug280519

    Sub Test()

        Console.Write("Bug 280519: ")
        Try
            Dim Obj As Object = Nothing
            Foo(Obj)  'this throws NullReferenceException instead of AmbiguousException
            Failed()
        Catch ex As Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

    Function Foo(ByVal Ary As Integer()(,,))
    End Function

    Function Foo(Optional ByVal Ary(,)(,,) As String = Nothing)
    End Function
End Module


Module Bug279014

    Private m_i As Integer

    Private Class c1
        Default Public Property foo(ByVal i As Integer) As Integer
            Get
                m_i = 1
            End Get
            Set(ByVal Value As Integer)
                m_i = 2
            End Set
        End Property
    End Class

    Private Class c2
        Inherits c1
        Default Public Overloads Property foo(ByVal i As Char) As Integer
            Get
                m_i = 3
            End Get
            Set(ByVal Value As Integer)
                m_i = 4
            End Set
        End Property
    End Class

    Sub Test()
        Console.Write("Bug 279014: ")

        Dim c As c2
        Dim d As Object
        Dim iEarly As Integer

        Try
            c = New C2()
            c(1) = 1
            iEarly = m_i

            d = c
            d(1) = 1
            PassFail(m_i = iEarly)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module



Module Bug283039
    Private m_i As Integer

    Private Class Cls1
        Public Sub Foo(ByVal Arg As Integer)
            m_i = 1
        End Sub

        Public Sub Foo(ByVal Arg As Char)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 283039: ")

        Try
            Dim Obj As Object = New Cls1()

            Obj.Foo(Nothing)
            Failed()

        Catch ex As Reflection.AmbiguousMatchException
            'verify that the message indicates a widening error
            PassFail(InStr(ex.Message, "narrow") = 0)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug257649
    Private m_i As Integer
    Private Class cls1
        Sub foo(ByVal ParamArray x() As Integer)
            m_i = 1
        End Sub

        Sub foo(ByVal x As Integer, ByVal ParamArray y() As Integer)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 257649: ")
        Try
            Dim x As Object = New Cls1()
            x.foo(1, 1, 1, 1)
            PassFail(m_i = 2)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug275937
    Private m_i As Integer
    Private m_foo1 As Integer()

    Private Class c1
        Public Property foo1(ByVal i As Integer) As Integer
            Get
                m_i = 1
            End Get
            Set(ByVal Value As Integer)
                m_i = 2
            End Set
        End Property
        Public Property foo2(ByVal i As Integer) As Integer
            Get
                m_i = 3
            End Get
            Set(ByVal Value As Integer)
                m_i = 4
            End Set
        End Property
    End Class

    Private Class c2
        Inherits c1
        Public Shadows foo1 As Integer()
        Public Shadows foo2 As Integer
        Sub New()
            foo1 = New Integer(5) {0, 1, 2, 3, 4, 5}
            m_foo1 = foo1
        End Sub
    End Class

    Sub Test()
        Dim i As Integer
        Dim c As Object
        'Dim c As c2

        m_i = -1
        c = New c2()

        Console.WriteLine("Bug 275937: ")
        Try
            Console.Write("   1) ")
            c.foo1(1) = 2
            PassFail(m_i = -1 AndAlso m_foo1(1) = 2)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write("   2) ")
            c.foo2(1) = 2
            Failed()
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module


Module Bug287347

    Private m_i As Integer

    Private Class CLs1
        Public Function Foo(ByVal Arg As String)
            m_i = 1
        End Function

        Public Function Foo(ByVal Arg() As Char)
            m_i = 2
        End Function
    End Class

    Sub Test()
        Dim Obj As CLs1
        Dim iEarly As Integer

        Console.Write("Bug 287347: ")
        Try
            Obj = New CLs1()
            Obj.Foo(54)
            iEarly = m_i
            m_i = -1
            Obj.Foo(CObj(54))
            PassFail(iEarly = m_i)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug287352
    Private m_x As Integer

    Private Class cls1
        Function foo(ByVal x() As Char) As Object
            m_x = 1
        End Function

        Function foo(ByVal x As Integer) As Object
            m_x = 2
        End Function
    End Class


    Sub Test()
        Console.Write("Bug 287352: ")
        Try
            Dim o As Object = New cls1()

            O.Foo(CObj("54"))
            Failed()
        Catch ex As System.Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug287357
    Private m_x As Integer

    Private Class cls1
        Function foo(ByVal x() As Char) As Object
            m_x = 1
        End Function

        Function foo(ByVal x As Char) As Object
            m_x = 2
        End Function
    End Class


    Sub Test()
        Console.Write("Bug 287357: ")
        Try
            Dim o As Object = New cls1()

            O.Foo(CObj("54"))
            Failed()
        Catch ex As System.Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug287831
    Private m_i As Integer

    Private Class cls1
        Sub foo(ByVal ParamArray x() As Integer)
            m_i = 1
        End Sub
        Sub foo(ByVal x As Integer, ByVal ParamArray y() As Integer)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 287831: ")
        Try
            Dim x As New cls1()
            Dim o As Object = x
            Dim iEarly As Integer

            x.foo(New Integer() {1, 1, 1})
            iEarly = m_i
            m_i = -1
            o.foo(New Integer() {1, 1, 1})
            PassFail(m_i = iEarly)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module

Module Bug287837
    Private m_i As Integer

    Private Class cls1
        Sub foo(ByVal ParamArray x() As Integer)
            m_i = 1
        End Sub
        Sub foo(ByVal x As Integer, ByVal ParamArray y() As Integer)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 287837: ")
        Try
            Dim x As cls1 = New cls1()
            Dim o As Object = x
            Dim iEarly As Integer

            x.foo(1, New Integer() {1, 1, 1})
            iEarly = m_i : m_i = -1
            o.foo(1, New Integer() {1, 1, 1})
            PassFail(m_i = iEarly)
            'Catch ex As System.Reflection.AmbiguousMatchException
            '    Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug272002

    Private m_i As Integer

    Private Class c1
        Overloads Sub foo(ByVal x As Short)
            m_i = 1
        End Sub
        Overloads Sub foo(Optional ByVal x As Integer = 10, Optional ByVal y As String = "")
            m_i = 2
        End Sub

    End Class

    Private Class c2
        Inherits c1
        Overloads Sub foo(ByVal x As Integer, ByVal y As String)
            m_i = 3
        End Sub
    End Class

    Sub Test()
        Dim c As C2
        Dim o As Object
        Dim s As Short = 10
        Dim iEarly As Integer

        c = New c2()
        o = c

        Console.WriteLine("Bug 272002:")

        Console.Write("   1) ")
        Try
            m_i = -1
            c.foo(10, "a")
            iEarly = m_i : m_i = -1
            o.foo(10, "a")
            PassFail(iEarly = m_i)
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.Write("   2) ")
        Try
            m_i = -1
            c.foo(s, "a")
            iEarly = m_i : m_i = -1
            o.foo(s, "a")
            PassFail(iEarly = m_i)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module

Module Bug289690
    Private m_i As Integer

    Private Class cls1
        Sub foo(ByVal ParamArray x() As Integer)
            m_i = 1
        End Sub
        Sub foo(ByVal x As Integer(), ByVal ParamArray y() As Integer)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 289690: ")
        Try
            Dim x As cls1 = New cls1()
            Dim o As Object = x

            o.foo(New Integer() {1, 1, 1})
            Failed()
        Catch ex As System.Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug270472

    Private m_i As Integer

    Private Class c1
        Sub foo(Optional ByVal x As Integer = 10)
            m_i = 1
        End Sub
    End Class

    Private Class c2
        Inherits c1
        Shadows Sub foo(ByVal x As Integer, Optional ByVal y As Integer = 5)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Dim c As C2
        Dim o As Object

        Console.Write("Bug 270472: ")
        c = New c2()
        o = c
        Try
            'c.foo() should throw exception
            o.foo()
            Failed()
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug284107
    Private m_i As Integer

    Private Class cls2
        Property P1(ByVal x As Integer, ByVal ParamArray y() As Integer) As Integer
            Get
            End Get
            Set(ByVal Value As Integer)
            End Set
        End Property
    End Class

    Sub Test()

        Dim c As New cls2()
        Dim o As Object = c
        Dim iEarly As Integer

        Console.Write("Bug 284107: ")
        Try
            c.P1(1, 1, 1) = 1
            iEarly = m_i

            o.P1(1, 1, 1) = 1
            PassFail(iEarly = m_i)

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module


Module Bug291120

    Private m_i As Integer

    Private Class c1
        Sub foo()
            m_i = 1
        End Sub
    End Class

    Private Class c2
        Inherits c1
        Shadows Property foo()
            Get
                m_i = 2
            End Get
            Set(ByVal Value)
                m_i = 3
            End Set
        End Property

    End Class

    Sub Test()

        Dim c As C2
        Dim o As Object

        c = New c2()

        Console.Write("Bug 291120: ")
        Try
            o = c
            o.foo()
            Failed()
        Catch ex As MissingMemberException
            'console.writeline(ex.message)
            'Passed()
            Failed()
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module


Module Bug291287
    Private Class Cls1
        Public ReadOnly i As Integer = 20
    End Class

    Sub Test()
        Dim c As New Cls1()
        Dim Obj As Object = c

        Console.Write("Bug 291287: ")
        Try
            Obj.i += 10 'this statement changes the value of the readonly field, but should not be doing so - BUG,
            Failed()
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module


Module Bug293485
    Private m_i As Integer

    Private Class cls1
        Overridable Sub foo(ByVal x() As Byte, ByVal ParamArray y() As Integer)
            m_i = 1
        End Sub
        Overridable Sub foo(ByVal x() As Integer, ByVal ParamArray y() As Integer)
            m_i = 2
        End Sub
    End Class

    Sub Test()

        Console.Write("Bug 293485: ")

        Dim x As cls1 = New cls1()
        Dim o As Object
        Try
            o = x
            o.foo(Nothing, New Integer() {1, 1, 1})
            Failed()
        Catch ex As Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module


Module Bug294938
    Private m_i As Integer

    Private Class cls1
        Sub foo1(ByVal ParamArray x As Integer())
            m_i = 1
        End Sub
        Sub foo1(ByVal x As Integer)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Dim c As New cls1()
        Dim o As Object
        o = c

        Console.Write("Bug 294938: ")
        Try
            o.foo1(True)
            PassFail(m_i = 2)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module

Module Bug294942
    Private m_i As Integer
    Private Class cls1
        Sub foo1(ByVal ParamArray x As Integer())
            m_i = 1
        End Sub
        Sub foo1(ByVal x As Integer)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Dim c As New cls1()
        Dim o As Object
        Dim iEarly As Integer

        Console.Write("Bug 294942: ")
        Try
            c.foo1(x:=1)
            iEarly = m_i : m_i = -1
            o = c
            o.foo1(x:=1)
            PassFail(m_i = iEarly)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug294987
    Private m_i As Integer
    Private Class cls1
        Overridable Sub foo(ByVal x As Integer, ByVal ParamArray y As Short())
            m_i = 1
        End Sub
        Overridable Sub foo(ByVal x As Integer, ByVal ParamArray q As Integer())
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Dim c As New cls1()
        Dim o As Object
        Dim iEarly As Integer

        o = c
        Console.Write("Bug 294987: ")
        Try
            c.foo(CByte(1), CByte(1))
            iEarly = m_i : m_i = -1
            o.foo(CByte(1), CByte(1))
            PassFail(m_i = iEarly)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module

Module Bug294970
    Private m_i As Integer

    Private Class cls1
        Sub foo(ByVal x As Integer, ByVal ParamArray y As Integer())
            m_i = 1
        End Sub
        Sub foo(ByVal x As Integer, ByVal z As Integer, ByVal y As Integer)
            m_i = 2
        End Sub
    End Class

    Sub test()
        Console.Write("Bug 294970: ")
        Dim c As New cls1()
        Dim o As Object
        Dim iEarly As Integer

        Try
            o = c
            c.foo(CByte(1), CByte(1), CByte(1))
            iEarly = m_i : m_i = -1
            o.foo(CByte(1), CByte(1), CByte(1))
            PassFail(iEarly = m_i)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module

Module Bug294956
    Private m_i As Integer

    Private Class cls1
        Sub foo1(ByVal x() As Short, ByVal ParamArray y As Integer())
            m_i = 1
        End Sub
        Sub foo1(ByVal x As Integer, ByVal ParamArray y As Integer())
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 294956: ")

        Dim x As cls1 = New cls1()
        Dim o As Object

        Try
            o = x
            o.foo1(Nothing)
            Failed()
        Catch ex As Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module

Module Bug294958
    Private m_i As Integer

    Private Class cls1
        Sub foo1(ByVal x() As Short, ByVal ParamArray y As Integer())
            m_i = 1
        End Sub
        Sub foo1(ByVal x As Integer, ByVal ParamArray y As Integer())
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 294958: ")

        Dim x As cls1 = New cls1()
        Dim o As Object

        Try
            o = x
            o.foo1(Nothing, Nothing)
            Failed()
        Catch ex As Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module

Module Bug293134

    Private Class Class1
        Sub Foo(ByVal x As Integer)
            PassFail(x = 42)
        End Sub
    End Class

    Private Class Class2
        Function Item(ByVal i As Integer) As Struct1
            Dim n As Struct1
            n.blah = i
            Return n
        End Function
    End Class

    Structure Struct1
        Public blah As Integer
    End Structure

    Sub Test()
        Console.Write("Bug 293134: ")

        Dim c, i, o As Object
        o = New Class1()
        c = New Class2()
        i = 42

        o.Foo(c.Item(i).blah)
    End Sub

End Module


Module Bug297690
    Private m_i As Integer

    Private Class cls1
        Public Property P1(ByVal ParamArray x As Integer()) As Integer
            Get

            End Get
            Set(ByVal Value As Integer)
                m_i = 1
            End Set
        End Property

        Public Property P1(ByVal y As Integer, ByVal ParamArray x As Integer()) As Integer
            Get

            End Get
            Set(ByVal Value As Integer)
                m_i = 2
            End Set
        End Property
    End Class

    Sub Test()
        Dim c As New cls1()
        Dim o As Object

        o = c
        Console.Write("Bug 297690: ")
        Try
            'early bound is ambiguous
            'c.p1(Nothing, Nothing) = 1
            o.p1(Nothing, Nothing) = 1
            PassFail(m_i = 2)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


#if True
Module Bug297690b
    Private m_i As Integer

    Private Class cls1
        Public Sub P1(ByVal ParamArray x As Integer())
            m_i = 1
        End Sub

        Public Sub P1(ByVal y As Integer, ByVal ParamArray x As Integer())
            m_i = 2
        End Sub

        Public Sub P2(ByVal ParamArray x As Integer())
            m_i = 3
        End Sub

        Public Sub P2(ByVal y As Object, ByVal x As Integer())
            m_i = 4
        End Sub
    End Class

    Sub Test()
        Dim c As New cls1()
        Dim o As Object

        o = c
        Console.Write("Bug 297690: ")
        Try
            o.p2(Nothing, Nothing)
            Failed()
        Catch ex As Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module
#End If

Module Bug297717
    Private m_i As Integer

    Private Class cls1
        Public Property P1(ByVal ParamArray x As Integer()) As Integer
            Get
                m_i = 1
            End Get
            Set(ByVal Value As Integer)
                m_i = 2
            End Set
        End Property

        Public Property P1(ByVal y As Short, ByVal ParamArray x As Integer()) As Integer
            Get
                m_i = 3
            End Get
            Set(ByVal Value As Integer)
                m_i = 4
            End Set
        End Property
    End Class

    Sub Test()
        Dim c As New cls1()
        Dim o As Object

        o = c
        Console.Write("Bug 297717: ")
        Try
            'early bound is ambiguous
            'c.p1(True)
            o.p1(True) = 1
            Failed()
        Catch ex As Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module



Module Bug297731
    Private m_i As Integer

    Private Class cls1
        Public Property P1(ByVal y As Integer, ByVal ParamArray x As Object()) As Integer
            Get
                m_i = 1
            End Get
            Set(ByVal Value As Integer)
                m_i = 2
            End Set
        End Property
    End Class

    Sub Test()
        Dim c As New cls1()
        Dim o As Object
        Dim iEarly As Integer

        o = c
        Console.Write("Bug 297731: ")
        Try
            'omitted argument cannot match a paramarray member
            c.p1(1, System.Reflection.Missing.Value) = True
            iEarly = m_i : m_i = -1
            o.P1(1, ) = True
            PassFail(iEarly = m_i)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module



Module Bug297792
    Private m_i As Integer

    Private Class cls1
        Public Property P1(Optional ByVal x As Integer = 1, Optional ByVal y As Integer = 2) As Integer
            Get
                m_i = 1
            End Get
            Set(ByVal Value As Integer)
                m_i = 2
            End Set
        End Property
    End Class

    Sub Test()
        Dim c As New cls1()
        Dim o As Object

        o = c
        Console.Write("Bug 297792: ")
        Try
            Dim iEarly As Integer
            c.p1(, 1) = True
            iEarly = m_i
            o.P1(, 1) = True
            PassFail(iEarly = m_i)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module



Module Bug297815
    Private m_i As Integer

    Private Class cls1
        Public Property P1() As Integer
            Get
                m_i = 1
            End Get
            Set(ByVal Value As Integer)
                m_i = 2
            End Set
        End Property
    End Class

    Sub Test()
        Dim c As New cls1()
        Dim o As Object

        o = c
        Console.Write("Bug 297815: ")
        Try
            Dim iEarly As Integer
            'This will cause a get of o.P1 and then do an indexed set using (y:=1), but no default will be found
            o.P1(y:=1) = True
            Failed()
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module



Module Bug297835
    Private m_i As Integer

    Private Class cls1
        Property P1(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As Integer
            Get
                m_i = 1
            End Get
            Set(ByVal Value As Integer)
                m_i = 2
            End Set
        End Property
    End Class

    Sub Test()
        Dim c As New cls1()
        Dim o As Object

        o = c
        Console.Write("Bug 297835: ")
        Try
            Dim iEarly As Integer
            'c.P1(x:=1, y:=1) = True
            'iEarly = m_i
            o.P1(x:=1, y:=1) = True
            Failed()
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module



Module Bug297892
    Private m_i As Integer

    Private Class cls1
        Public Property P1(ByVal ParamArray x As Integer()()) As Integer
            Get
                m_i = 1
            End Get
            Set(ByVal Value As Integer)
                m_i = 2
            End Set
        End Property
        Public Property P1(ByVal x As Integer(), ByVal ParamArray y As Integer()) As Integer
            Get
                m_i = 3
            End Get
            Set(ByVal Value As Integer)
                m_i = 4
            End Set
        End Property
    End Class

    Sub Test()
        Dim c As New cls1()
        Dim o As Object

        o = c
        Console.Write("Bug 297892: ")
        Try
            Dim iEarly As Integer
            c.P1(New Integer() {}, New Integer() {}, New Integer() {}) = True
            iEarly = m_i
            o.P1(New Integer() {}, New Integer() {}, New Integer() {}) = True
            PassFail(iEarly = m_i)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug298131
    Private m_i As Integer

    Private Class cls1
        Public ReadOnly Property P1()
            Get
                m_i = 1
            End Get
        End Property

        Public Property P1(ByVal ParamArray x As Integer()) As Integer
            Get
                m_i = 2
            End Get
            Set(ByVal Value As Integer)
                m_i = 3
            End Set
        End Property
    End Class

    Sub Test()
        Dim c As New cls1()
        Dim o As Object

        o = c
        Console.Write("Bug 298131: ")
        Try
            'early bound is ambiguous
            'c.p1() = 1
            o.p1() = 1
            Failed()
        Catch ex As MissingMemberException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug296923
    Private m_i As Integer

    Private Class cls1

        Sub foo(ByVal ParamArray x As Integer())
            m_i = 1
        End Sub

        Sub foo(ByVal x As Integer, ByVal y As Boolean)
            m_i = 2
        End Sub

    End Class

    Sub Test()
        Dim c As New cls1()
        Dim o As Object

        o = c
        Console.Write("Bug 296923: ")
        Try
            'early bound is ambiguous
            'c.p1() = 1
            o.foo(x:=1)
            Failed()
        Catch ex As InvalidCastException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug298658
    Private m_arg1 As Integer
    Private m_arg2 As Integer

    Private Class Cls1
        Public Sub Moo(ByVal Arg1 As Integer, ByVal Arg2 As Integer)
            m_arg1 = Arg1
            m_arg2 = Arg2
        End Sub
    End Class

    Sub Test()
        Dim c As Cls1
        Dim o As Object

        c = New Cls1()
        o = c

        Try
            Console.Write("Bug 298658: ")
            o.Moo(Arg2:=20, Arg1:=40)

            PassFail(m_arg2 = 20 AndAlso m_arg1 = 40)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module


Module Bug302246
    Private m_i As Integer
    Private m_ArgInteger As Integer
    Private m_Arg2 As Integer
    Private m_ArgString As Integer

    Public Class Class1
        Public Sub foo(ByVal Arg As Integer, ByVal Arg2 As Integer) ', Optional ByVal Arg2 As Long = 40)
            m_i = 1
            m_ArgInteger = Arg
            m_Arg2 = Arg2
        End Sub

        Public Sub Foo(ByVal Arg2 As Integer, ByVal Arg As String)
            m_i = 2
            m_ArgString = Arg
            m_Arg2 = Arg2
        End Sub
    End Class


    Sub Test()

        Console.Write("Bug 302246: ")

        Try
            Dim iEarly As Integer
            Dim c As New Class1()
            Dim o As Object = c

            m_i = -1
            c.foo(40, Arg:=50)
            iEarly = m_i

            m_i = -1
            o.foo(40, Arg:=50)  'this late bound case throws an unexpected exception - BUG
            PassFail(m_i = iEarly AndAlso m_ArgString = "50" AndAlso m_Arg2 = 40)

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module



Module Bug302354
    Private m_i() As Integer
    Private m_Count As Integer

    Public Class Class1
        Public Function Foo(ByVal Arg2 As UInt16, ByVal Arg1 As Integer)
            m_i(m_Count) = 1
            m_Count += 1
            'Console.WriteLine("in foo overload 1 : " & Arg2.ToString & " " & Arg1)
            Return Arg1
        End Function

        Public Function Foo(ByVAl Arg2 As Object, ByVal Arg1 As Integer)
            m_i(m_Count) = 2
            m_Count += 1
            'Console.WriteLine("in foo overload 2 : " & Arg2.ToString & " " & Arg1)
            Return Arg2
        End Function
    End Class

    Sub Test()
        Console.Write("Bug 302354: ")

        Try
            Dim c As Class1
            Dim o As Object
            Dim EarlyBoundOrder() As Integer

            c = New Class1
            o = c

            m_i = New Integer(2) { }

            c.foo(c.foo(New Object, 10), c.foo(New UInt16(), 42))
            EarlyBoundOrder = m_i

            m_Count = 0
            m_i = New Integer(2) { }
            o.foo(o.foo(New Object, 10), o.foo(New UInt16(), 42))

            PassFail(AreEqual(m_i, EarlyBoundOrder))

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

    Private Function AreEqual(a As Integer(), b As Integer()) As Boolean
        If a.Length <> b.Length THen
            Return False
        ENd If

        Dim i As Integer
        for i = 0 to a.GetUpperBound(0)
            If a(i) <> b(i) Then
                Return False
            End If
        Next i

        Return True
    End Function
End Module



Module Bug302374

    Private Class Class1
        Public x As Integer

        Public m_Arg1 As Integer
        Public m_Arg2 As Integer

        Public Function Foo(ByVal Arg1 As Integer, ByVal Arg2 As Integer)
            x += 1
            m_Arg1 = Arg1
            m_Arg2 = Arg2
            Return Arg1 + Arg2
        End Function
    End Class

    Sub Test()
        Console.Write("Bug 302374: ")
        Try
            Dim c As Class1
            Dim o As Object
            Dim iEarlyCount As Integer

            c = New Class1
            o = c

            c.foo(20, Arg2:=o.foo(o.foo(Arg1:=40, Arg2:=60), 200))
            iEarlyCount = c.x

            c.x = 0
            o.foo(20, Arg2:=o.foo(o.foo(Arg1:=40, Arg2:=60), 200))
            PassFail(iEarlyCount = c.x)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug302374b
    'Bug found while fixing 302374
    Public m_x, m_y, m_z As Integer

    Private Class class1

        'All arguments are byref
        Sub foo(Optional ByRef x As Integer = 1, Optional ByRef y As Integer = 2, Optional ByRef z As Integer = 3)
            m_x = x
            x = 1

            m_y = y
            y = 2

            m_z = z
            z = 3
        End Sub

        'Second argument is byref
        Sub foo2(Optional ByVal x As Integer = 1, Optional ByRef y As Integer = 2, Optional ByVal z As Integer = 3)
            m_x = x
            x = 1

            m_y = y
            y = 2

            m_z = z
            z = 3
        End Sub

    End Class

    Sub Test()
        
        Console.WriteLine("Bug 302374b: ")
        Try
            Dim c As class1
            Dim o As Object
            Dim v1, v2, v3 As Integer

            c = New class1()
            o = c

            v1 = 11
            v3 = 33
            Console.Write("   1): ")
            o.Foo(x:=v1, z:=v3)
            PassFail(v1 = 1 AndAlso v3 = 3 AndAlso m_x = 11 AndAlso m_z = 33)

            m_y = 0 : m_x = 0 : m_z = 0
            v2 = 22
            Console.Write("   2): ")
            o.Foo2(y:=v2)
            PassFail(v2 = 2 AndAlso m_y = 22)

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module


Module Bug305634
    Public m_i As Integer

    Private Class class1

        Sub foo(ByVal y As Integer, Optional ByVal x As Short = 2)
            m_i = 1
        End Sub

        Sub foo(ByVal ParamArray x() As Integer)
            m_i = 2
        End Sub

    End Class

    Sub Test()
        
        Console.Write("Bug 305634: ")
        Try
            Dim c As Class1 
            Dim o As Object
            Dim iEarly As Integer

            c = New class1()
            o = c

            c.foo(y:=1, x:=1)

            iEarly = m_i : m_i = -1

            o.foo(y:=1, x:=1)

            PassFail(m_i = iEarly)

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module


Module Bug304212

    Private Class Class1

        Public Function Func1(Optional ByVal Arg1 As Double = 200, Optional ByRef Arg2 As Single = -100) As Double
            Arg2 += 10
            Return Arg1 + Arg2 + 10
        End Function

        Public Function Func2(Optional ByRef Arg1 As Integer = -100) As Double
            Arg1 += 10
            Return Arg1
        End Function

    End Class

    Sub Test()
        Console.Write("Bug 304212: ")
        Try
            Dim c As Class1 = New Class1()
            Dim o As Object = c
            Dim i1, i2 As Integer

            i1 = o.Func1(500)
            i2 = o.Func1(500)
            PassFail(i1 = i2)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug306698
    Private m_i As Integer

    Private Class Cls2_2
        Sub foo(ByVal Arg2 As Double)
            m_i = 1
        End Sub
    End Class

    Private Class Cls2_1
        Inherits Cls2_2

        'NOTE: Intentially left off Overloads keyword here
        Sub foo(ByVal arg1 As Single)
            m_i = 2
        End Sub
    End Class

    Private Class Cls2
        Inherits Cls2_1

        'NOTE: Intentially left off Overloads keyword here
        Sub foo(ByVal arg2 As Integer)
            m_i = 3
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 306698: ")
        Try
            Dim d As Double
            Dim c As Cls2
            Dim o As Object
            Dim iEarly As Integer

            c = New Cls2()
            o = c

            c.foo(d)
            iEarly = m_i : m_i = -1

            o.foo(d)
            PassFail(iEarly = m_i)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug307060
    Private m_i As Integer

    Private Class base
        Public Overloads Sub foo(Optional ByVal x As String = "")
            m_i = 1
        End Sub
    End Class

    Private Class derived
        Inherits base
        Public Overloads Sub foo(ByVal ParamArray x() As String)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 306070: ")
        Try
            Dim c As Derived
            Dim o As Object
            Dim oString As Object
            Dim iFirst As Integer

            c = New Derived()
            oString = "Hello"

            c.foo()
            iFirst = m_i : m_i = -1

            c.foo(oString)

            PassFail(iFirst = m_i)

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module



Module Bug307060b
    Private m_i As Integer
    Const FooString As Integer = 1
    Const FooStringParamArray As Integer = 2

    Private Class base
        Public Overloads Sub foo(ByVal x As String)
            m_i = FooString
        End Sub
    End Class

    Private Class derived
        Inherits base
        Public Overloads Sub foo(ByVal ParamArray x() As String)
            m_i = FooStringParamArray
        End Sub
    End Class

    Private Class both
        Public Overloads Sub foo(ByVal x As String)
            m_i = FooString
        End Sub
        Public Overloads Sub foo(ByVal ParamArray x() As String)
            m_i = FooStringParamArray
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 306070b: ")
        Try
            Dim iNotInherited As Integer
            Dim c1 As New Both()
            Dim c2 As New Derived()
            Dim oString As Object = "Hello"

            c1.foo(oString)
            iNotInherited = m_i : m_i = -1

            c2.foo(oString)
            PassFail(m_i = iNotInherited)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug306786
    Private m_i As Integer

    Private Class cls1
        Overridable Sub foo(ByVal y As Long, ByVal z As Long, Optional ByVal x As Long() = Nothing)
            m_i = 1
        End Sub
        Overridable Sub foo(ByVal y As Long, ByVal ParamArray x() As Long)
            m_i = 2
        End Sub
    End Class

    Sub Test()

        Console.Write("Bug 306786: ")
        Try

            Dim c As cls1
            Dim o As Object
            
            c = New Cls1()
            o = c

            'c.foo(Nothing, Nothing, Nothing) 'earlybound
            o.foo(Nothing, Nothing, Nothing)
            Failed()

        Catch ex As Reflection.AmbiguousMatchException
            Passed()

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug308339
    Private m_i As Integer
    Private m_x As Object
    Private m_y As Integer()

    Private Class cls1
        Overridable Sub foo(ByVal x As Object)
            m_i = 1
            m_x = x
        End Sub
        Overridable Sub foo(ByVal ParamArray y() As Integer)
            m_i = 2
            m_y = y
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 308339: ")
        Try
            Dim c As cls1 = New cls1()
            Dim o As Object = c
            Dim iEarly As Integer
            Dim y As Integer()
            Dim x As Object
            Dim bPassed As Boolean

            c.foo(Nothing)
            iEarly = m_i : m_i = -1

            o.foo(Nothing)
            If iEarly = m_i Then
                If iEarly = 1 Then
                    bPassed = (m_x = x)
                Else
                    If m_y Is Nothing AndAlso y Is Nothing Then
                        bPassed = True
                    ElseIf m_y.Length = 1 AndAlso y.Length = 1 AndAlso m_y(0) = y(0) Then
                        bPassed = True
                    End If
                End If
            End If
            PassFail(bPassed)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module HarishK
    Private m_i As Integer

    Private Class cls2
        Function Foo() As Integer
            m_i = 1
        End Function

        Function foo(ByVal Arg() As Integer)
            m_i = 2
        End Function
    End Class

    Sub Test()
        
        Console.Write("Bug 308339: ")
        Try
            Dim c As cls2 = New cls2()
            Dim o As Object = c

            'c.Foo(20)
            o.foo(20)
        Catch ex As InvalidCastException
            Passed()

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug308219
    Private Class cls1
        Sub foo(ByVal x As Short)
        End Sub
        Sub foo(ByVal x As Long)
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 308219: ")
        Try
            Dim c As cls1 = New cls1()
            Dim o As Object = c

            o.foo(True)
            Failed()

        Catch ex As Reflection.AmbiguousMatchException
            Passed()

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug308240
    Private m_i As Integer

    Private Class cls1
        Sub foo(ByVal x As Object)
            m_i = 1
        End Sub
        Sub foo(ByVal ParamArray x() As Object)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 308240: ")
        Try
            Dim c As cls1 = New cls1()
            Dim o As Object = c
            Dim iEarly As Integer

            c.foo(1)
            iEarly = m_i : m_i = -1

            o.foo(1)
            PassFail(m_i = iEarly)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module

Module Bug308245
    Private m_i As Integer

    Private Class cls1
        Sub foo(ByVal x As Double)
            m_i = 1
        End Sub
        Sub foo(ByVal ParamArray x() As Object)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 308245: ")
        Try
            Dim c As cls1 = New cls1()
            Dim o As Object = c
            Dim iEarly As Integer

            c.foo(1)
            iEarly = m_i : m_i = -1

            o.foo(1)
            PassFail(m_i = iEarly)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug308345
    Private m_i As Integer

    Private Class Cls1
        Function Foo() As Integer
            m_i = 1
        End Function
    End Class

    Private Class cls2
        Inherits Cls1

        Overloads Function foo(ByVal ParamArray Arg() As Integer)
            m_i = 2
        End Function
    End Class


    Sub Test()
        Console.Write("Bug 308345: ")
        Try
            Dim c As cls2 = New cls2()
            Dim o As Object = c
            Dim iEarly As Integer

            c.foo()
            iEarly = m_i : m_i = -1

            o.foo()
            PassFail(m_i = iEarly)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug309526
    Private Class c4_1
        Public called = 0
        Public Sub x(ByVal i As Integer)
            called = 1
        End Sub
    End Class

    Private Class c4_2
        Inherits c4_1
        Public Shadows Sub x(ByVal i As Char)
            called = 3
        End Sub
        Public Shadows Sub x(ByVal i As c4_1)
            called = 4
        End Sub
    End Class

    'Test with implicit shadows
    Private Class c4_3
        Inherits c4_1
        Public Sub x(ByVal i As Char)
            called = 3
        End Sub
        Public Sub x(ByVal i As c4_1)
            called = 4
        End Sub
    End Class


    Sub Test()

        Console.Write("Bug 309526a: ")
        Try
            Dim c As c4_2
            Dim o As Object
            c = New c4_2()
            o = c

            'c.x(5) earlybound should give compiler error
            o.x(5)
            Failed()
       	Catch ex As AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.Write("Bug 309526b: ")
        Try
            Dim c As c4_3
            Dim o As Object
            c = New c4_3()
            o = c

            'c.x(5)  'earlybound should give compiler error
            o.x(5)
            Failed()
        Catch ex As AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug309529
    Private m_i As Integer

    Private Class cls1
        Sub foo(ByVal arg() As cls1)
            m_i = 1
        End Sub

        Sub foo(ByVal arg() As cls2)
            m_i = 2
        End Sub
    End Class

    Private Class cls2
        Inherits cls1
    End Class

    Private Class cls3
        Inherits cls2
    End Class

    Sub Test()

        Console.Write("Bug 309529: ")
        Try
            Dim c As New cls1()
            Dim o As Object = c
            Dim iEarly As Integer

            c.foo(New cls3() {})
            iEarly = m_i : m_i = -1

            o.foo(New cls3() {})  'unexpected exception here - BUG
            PassFail(iEarly = m_i)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug309444
    Private m_i As Integer

    Private Class cls1
        Sub foo(ByVal ParamArray arg() As Double)
            m_i = 1
        End Sub

        Sub foo(ByVal arg As Object)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 309444: ")
        Try
            Dim c As cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New cls1()
            o = c

            c.foo(20)
            iEarly = m_i : m_i = -1

            o.foo(20)
            PassFail(m_i = iEarly)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug309445

    Private Class cls1
        Sub foo(ByVal ParamArray arg() As Double)
            m_i = 1
        End Sub

        Sub foo(ByVal arg As Double)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 309445: ")
        Try
            Dim c As cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New cls1()
            o = c

            c.foo(20)
            iEarly = m_i : m_i = -1

            o.foo(20)
            PassFail(m_i = iEarly)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug309554
    Sub Test()
        Console.Write("Bug 309554: ")
        Try
            Dim c As cls1
            Dim o As Object

            c = New cls1()
            o = c

            'c.foo(New Cls2() { }) 'Ambiguous
            o.foo(New Cls2() { })
            Failed()

        Catch ex As Reflection.AmbiguousMatchException
            Passed()

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

    Private m_i As Integer

    Private Class cls1
        Sub foo(ByVal arg() As I1)
            m_i = 1
        End Sub

        Sub foo(ByVal ParamArray arg() As I2)
            m_i = 2
        End Sub
    End Class

    Private Class cls2
        Implements I2, I1
    End Class

    Private Interface I2
    End Interface

    Private Interface I1
    End Interface
End Module


Module Bug309613
    Private m_i As Integer

    Private Class cls1
        Sub foo(ByVal arg(,) As cls1)
            m_i = 1
        End Sub

        Sub foo(ByVal ParamArray arg() As cls1)
            m_i = 2
        End Sub
    End Class

    Private Class cls2
        Inherits cls1
    End Class

    Sub Main()

        Test()

    End Sub

    Sub Test()
        Console.Write("Bug 309613: ")
        Try
            Dim c As cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New cls1()
            o = c

            c.foo(New cls2() {})
            iEarly = m_i : m_i = -1

            o.foo(New cls2() {})
            PassFail(m_i = iEarly)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module BugUnknown
    Private m_i As Integer

    Private Class cls1
        Overloads Sub foo(ByVal arg() As cls1)
            m_i = 1
        End Sub
    End Class

    Private Class cls2
        Inherits cls1

        Overloads Sub foo(ByVal arg() As cls2)
            m_i = 2
        End Sub

    End Class

    Private Class cls3
        Inherits cls2
    End Class

    Sub Main()
        Dim c As New cls3()
        Dim o As Object = c

        c.foo(New cls3() {})
        o.foo(New cls3() {})  'unexpected exception here - BUG
    End Sub

End Module


Module Bug309766
    Private m_i As Integer

    Private Class cls1
        Sub foo(ByVal x As Object, Optional ByVal y As Double = 1)
            m_i = 1
        End Sub
        Sub foo(ByVal ParamArray y As Object())
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 309766: ")
        Try
            Dim c As cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New Cls1()
            o = c

            c.foo(Nothing, Nothing)
            iEarly = m_i : m_i = -1

            o.foo(Nothing, Nothing) '<--bug
            PassFail(iEarly = m_i)

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module


Module Bug309766b
    Private m_i As Integer

    Private Class cls1
        Sub foo(ByVal x As Object, Optional ByVal y As Object = Nothing)
            m_i = 1
        End Sub
        Sub foo(ByVal x As Object, ByVal ParamArray y As Object())
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 309766b: ")
        Try
            Dim c As cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New Cls1()
            o = c

            c.foo(Nothing, Nothing)
            iEarly = m_i : m_i = -1

            o.foo(Nothing, Nothing) '<--bug
            PassFail(iEarly = m_i)

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module


Module Bug307268
    Private m_i As Integer

    Private Class cls1
        Overridable Sub foo(ByVal y As Long, ByVal z As Long, Optional ByVal x As Long() = Nothing)
            m_i = 1
        End Sub
        Overridable Sub foo(ByVal y As Long, ByVal ParamArray x() As Long)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 307268: ")
        Try
            Dim c As cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New Cls1()
            o = c

            'c.foo(Nothing, Nothing) 'Ambiguous
            iEarly = m_i : m_i = -1

            o.foo(Nothing, Nothing)
            Failed()

        Catch ex As Reflection.AmbiguousMatchException
            Passed()

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module


Module Bug309820
    Private m_i As Integer

    Private Class cls1
        Sub foo(Optional ByVal y As Decimal = 1)
            m_i = 1
        End Sub
        Sub foo(ByVal ParamArray y As Decimal())
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 309820: ")
        Try
            Dim c As Cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New cls1
            o = c

            c.foo()
            iEarly = m_i : m_i = -1

            o.foo()
            PassFail(iEarly = m_i)

        Catch ex As Exception
            Failed(ex)
        End Try


    End Sub

End Module


Module Bug309822
    Private m_i As Integer

    Private Class cls1
        Overridable Sub foo(ByVal x As Object, Optional ByVal y As Object = 1)
            m_i = 1
        End Sub
        Overridable Sub foo(ByVal x As Object, ByVal ParamArray y As Object())
            m_i = 2
        End Sub
    End Class


    Sub test()
        Console.Write("Bug 309822: ")
        Try
            Dim c As cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New cls1()
            o = c

            c.foo(Nothing)
            iEarly = m_i : m_i = -1

            o.foo(Nothing)
            PassFail(iEarly = m_i)

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module


Module Bug309845
    Private m_i As Integer

    Private Class cls1
        Overridable Sub foo(ByVal x As Integer)
            m_i = 1
        End Sub
        Overridable Sub foo(ByVal x As Integer, ByVal ParamArray y As Object())
            m_i = 2
        End Sub
    End Class

    Sub test()
        Console.Write("Bug 309845: ")
        Try
            Dim c As cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New cls1()
            o = c

            c.foo(1)
            iEarly = m_i : m_i = -1

            o.foo(1)
            PassFail(iEarly = m_i)

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module


Module Bug309881
    Private m_i As Integer

    Private Class cls1
        Overridable Sub foo(ByVal ParamArray x() As Integer)
            m_i = 1
        End Sub
        Overridable Sub foo(ByVal x() As Integer, ByVal ParamArray y() As Integer)
            m_i = 2
        End Sub
    End Class

    Sub test()
        Console.Write("Bug 309881: ")
        Try
            Dim c As cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New cls1()
            o = c

            'c.foo(, )
            'iEarly = m_i : m_i = -1

            o.foo(, )
            Failed()
            
	Catch ex As AmbiguousMatchException
            Passed()

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module



Module Bug309882
    Private m_i As Integer

    Private Class cls2

        Property P1(ByVal x As Integer) As Integer
            Get
                m_i = 1
            End Get
            Set(ByVal Value As Integer)
                m_i = 2
            End Set
        End Property

        Property P2(ByVal ParamArray x() As Integer) As Integer
            Get
                m_i = 3
            End Get
            Set(ByVal Value As Integer)
                m_i = 4
            End Set
        End Property

    End Class

    Sub test()
        Console.WriteLine("Bug 309882")
        Try
            Dim c As cls2
            Dim o As Object

            c = New cls2()
            o = c

            'Earlybound give compiler error
            'c.P1(y:=1) = 1
            'c.P2(y:=1) = 1

            Console.Write("   1) ")
            Try
                o.P1(y:=1) = 1
                Failed()
            Catch ex As InvalidCastException
                Passed()
                Console.WriteLine("      " & ex.Message)
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("   2) ")
            Try
                o.P2(y:=1) = 1
                Failed()
            Catch ex As InvalidCastException
                Passed()
                Console.WriteLine("      " & ex.Message)
            Catch ex As Exception
                Failed(ex)
            End Try

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug309883
    Private m_i As Integer

    Private Class cls1
        Overridable Sub foo(ByVal ParamArray x() As Integer)
            m_i = 1
        End Sub
        Overridable Sub foo(ByVal x As Integer, ByVal ParamArray y() As Integer)
            m_i = 2
        End Sub
    End Class

    Sub test()
        Console.Write("Bug 309883: ")
        Try
            Dim c As cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New cls1()
            o = c

            c.foo(True, New Integer() {1, 1, 1})
            iEarly = m_i : m_i = -1

            o.foo(True, New Integer() {1, 1, 1})
            PassFail(iEarly = m_i)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module



Module Bug309884
    Private m_i As Integer

    Private Class cls1
        Overridable Sub foo(ByVal ParamArray x() As Integer)
            m_i = 1
        End Sub
        Overridable Sub foo(ByVal x() As Integer, ByVal ParamArray y() As Integer)
            m_i = 2
        End Sub
    End Class

    Sub test()
        Console.Write("Bug 309884: ")
        Try
            Dim c As cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New cls1()
            o = c

            c.foo(New Integer() {1, 1, 1}, 1)
            iEarly = m_i : m_i = -1

            o.foo(New Integer() {1, 1, 1}, 1)
            PassFail(iEarly = m_i)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module




Module Bug309894
    Private m_i As Integer

    Private Class cls1
        Property P1()
            Get
                m_i = 1
            End Get
            Set(ByVal Value)
                m_i = 2
            End Set
        End Property
        Property P1(ByVal ParamArray x As Integer()) As Integer
            Get
                m_i = 3
            End Get
            Set(ByVal Value As Integer)
                m_i = 4
            End Set
        End Property

        Sub foo()
            m_i = 5
        End Sub

        Sub foo(ByVal ParamArray x() As Integer)
            m_i = 6
        End Sub

    End Class

    Sub test()
        Console.WriteLine("Bug 309894: ")
        Try
            Dim c As cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New cls1()
            o = c

            Console.Write("   1) ")
            Try
                m_i = 0
                c.foo()
                iEarly = m_i : m_i = -1
                o.foo()
                PassFail(iEarly = m_i)
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("   2) ")
            Try
                m_i = 0
                c.P1() = 1
                iEarly = m_i : m_i = -1
                o.P1() = 1
                PassFail(iEarly = m_i)
            Catch ex As Exception
                Failed(ex)
            End Try

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module


Module Bug310044
    Private m_clr As c1.color

    Private Class c1
        Enum color
            unknown = -1
            red = 1
        End Enum

        Sub bar(ByRef x As color)
            m_clr = x
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 310044: ")
        Try
            Dim c As c1
            Dim o As Object
            Dim clrEarly As c1.color

            c = New c1()
            o = c

            c.bar(1)
            clrEarly = m_clr : m_clr = C1.Color.Unknown

            o.bar(1)
            PassFail(m_clr = clrEarly)

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub


End Module



Module Bug309805
    Private m_i As Integer

    Private Class cls1

        Sub foo(y As Object)
            m_i = 1
        End Sub

        Sub foo(ByVal ParamArray y As Object())
            m_i = 2
        End Sub

    End Class

    Sub test()
        Console.WriteLine("Bug 309805: ")
        Try
            Dim c As cls1
            Dim o As Object
            Dim iEarly As Integer

            c = New cls1()
            o = c

            Console.Write("   2) ")
            Try
                m_i = 0
                c.foo()
                iEarly = m_i : m_i = -1

                o.foo()
                PassFail(iEarly = m_i)
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("   2) ")
            Try
                m_i = 0
                c.foo(Nothing)
                iEarly = m_i : m_i = -1

                o.foo(Nothing)
                PassFail(iEarly = m_i)
            Catch ex As Exception
                Failed(ex)
            End Try

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub

End Module


Module Bug310146
    Private m_i As Integer

    Private Class cls2
        Overloads Function foo(ByVal ParamArray Arg()() As Integer)
            m_i = 1
        End Function

        Overloads Function foo(ByVal ParamArray Arg() As Integer)
            m_i = 2
        End Function
    End Class

    Sub Test()
        Console.Write("Bug 310146: ")
        Try
            Dim c As cls2
            Dim o As Object
            Dim iEarly As Integer

            c = New cls2()
            o = c

            'UNDONE: Enable when PaulV checks in earlybound fix
            'c.foo(New Integer() {}) 'Unexpected compile time error given here - BUG
            iEarly = m_i : m_i = -1

            o.foo(New Integer() {}) 'Unexpected run time error given here - BUG
            PassFail(m_i = 2)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module


Module Bug310322
    Private m_i As Integer

    Private Class Cls1
        Public Function foo(ByVal Arg1 As Integer)
            m_i = 1
        End Function
    End Class

    Private Class cls2
        Inherits Cls1

        Public Shadows Function Foo(ByVal Arg As Integer)
            m_i = 2
        End Function

        Public Shadows Function foo(ByVal Arg1 As String)
            m_i = 3
        End Function
    End Class

    Sub Test()

        Console.Write("Bug 310322: ")
        Try
            Dim c As cls2
            Dim o As Object
            Dim iEarly As Integer

            c = New cls2()
            o = c

            c.Foo(Arg1:=40)

            iEarly = m_i : m_i = -1
            o.Foo(Arg1:=40)

            PassFail(iEarly = m_i AndAlso m_i = 3)

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module

Module Bug310339
    Private m_i As Integer

    Private Class Cls1
        Public Function foo(ByVal Arg1 As Double)
            m_i = 1
        End Function
    End Class

    Private Class cls2
        Inherits Cls1

        Public Shadows Function Foo(ByVal ParamArray Ary As Long())
            m_i = 2
        End Function

    End Class

    Sub Test()
        Console.Write("Bug 310339: ")
        Try
            Dim c As cls2
            Dim o As Object
            Dim iEarly As Integer

            c = New cls2()
            o = c

            c.Foo(40)
            iEarly = m_i : m_i = -1

            o.Foo(40)
            PassFail(iEarly = m_i AndAlso m_i = 2)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module



Module Bug310340
    Private Class Cls1
        Public Sub Foo(ByVal Arg As Integer, ByVal Arg1 As Double)
            m_i = 1
        End Sub
    End Class

    Private Class cls2
        Inherits Cls1

        Public Overloads Sub Foo(ByVal Arg As Integer, ByVal ParamArray Ary As Long())
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 310339: ")
        Try
            Dim c As cls2
            Dim o As Object
            Dim iEarly As Integer

            c = New cls2()
            o = c

            c.Foo(40, 80)
            iEarly = m_i : m_i = -1

            o.Foo(40, 80)
            PassFail(iEarly = m_i AndAlso m_i = 2)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module



Module Bug310341
    Private Class Cls1
        Public Sub Foo(ByVal Arg As Integer, ByVal Arg1 As Double)
            m_i = 1
        End Sub
    End Class

    Private Class cls2
        Inherits Cls1

        Public Shadows Sub Foo(ByVal Arg As Integer, ByVal Arg1 As Double)
            m_i = 2
        End Sub

        Public Shadows Sub Foo(ByVal Arg As Integer, ByVal ParamArray Ary As Long())
            m_i = 3
        End Sub
    End Class


    Sub Test()
        Console.Write("Bug 310341: ")
        Try
            Dim c As cls2
            Dim o As Object
            Dim iEarly As Integer

            c = New cls2()
            o = c

            c.Foo(40, 80)
            iEarly = m_i : m_i = -1

            o.Foo(40, 80)
            PassFail(iEarly = m_i AndAlso m_i = 3)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug310333
    Private m_i As Integer

    Private Class Cls1
        Public Overloads Sub foo()
            m_i = 1
        End Sub
        Public Overloads Sub foo1(ByVal x As Integer)
            m_i = 2
        End Sub
    End Class

    Private Class cls2
        Inherits Cls1
        Public Overloads Sub foo(Optional ByVal x As Integer = 1)
            m_i = 3
        End Sub
        Public Overloads Sub foo1(ByVal x As Integer, Optional ByVal y As Integer = 1)
            m_i = 4
        End Sub
    End Class

    Sub Test()
        Console.WriteLine("Bug 310333: ")
        Try
            Dim c As cls2
            Dim o As Object
            Dim iEarly As Integer

            c = New cls2()
            o = c

            Console.Write("   1) ")
            Try
                c.foo()
                iEarly = m_i : m_i = -1
                o.foo()
                PassFail(iEarly = m_i)
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("   2) ")
            Try
                c.foo1(1)
                iEarly = m_i : m_i = -1
                o.foo1(1)
                PassFail(iEarly = m_i)
            Catch ex As Exception
                Failed(ex)
            End Try
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module



Module Bug310786
    Private m_i As Integer

    Private Class cls1
    End Class

    Private Class cls2
        Inherits cls1
    End Class

    Private Class cls4
        Inherits cls2
    End Class


    Private Class cls3
        Sub foo(ByVal ParamArray Arg1() As cls2)
            m_i = 1
        End Sub

        Sub foo(ByVal Arg1() As cls1)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 310786: ")
        Try
            Dim c As New cls3()
            Dim o As Object
            Dim cls4Ary() As cls4

            cls4Ary = New cls4() {}

            c = New cls3()
            o = c
            'c.foo(cls4Ary) 'BUGBUG: enable when early bound fixed
            o.foo(cls4Ary)  'this invokes Foo(cls()), but should be invoking Foo(cls2()) - BUG
            PassFail(m_i = 1)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module



Module Bug310713
    Private m_i As Integer

    Private Class cls2
        Overloads Sub foo(ByVal arg1 As Integer)
            m_i = 1
        End Sub

        Overloads Sub foo(ByVal arg2 As String)
            m_i = 2
        End Sub

        Overloads Sub foo(ByVal arg As Object)
            m_i = 3
        End Sub
    End Class

    Sub Test()

        Console.Write("Bug 310713: ")
        Try
            Dim c As New cls2()
            Dim o As Object = c

            'c.foo(Nothing) 'earlybound when uncommented gives the correct ambiguous resolution compile error
            o.foo(Nothing) 'this invokes cls2.Foo(Integer) instead of giving ambiguous exception
        Catch ex As Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module

Module Bug310716
    Private m_i As Integer

    Private Class cls2
        Overloads Sub foo(ByVal arg1 As String, Optional ByVal arg2 As Short = 40)
            m_i = 1
        End Sub

        Overloads Sub foo(ByVal arg1 As Object, Optional ByVal arg2 As String = "Hello")
            m_i = 2
        End Sub
    End Class

    Sub Test()

        Console.Write("Bug 310716: ")
        Try
            Dim c As New cls2()
            Dim o As Object = c

            'c.foo(Nothing, Nothing)  'early bound when commented gives correct compile error

            o.foo(Nothing, Nothing)
        Catch ex As Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module



Module Bug310784
    Private Class cls1
    End Class

    Private Class cls2
        Inherits cls1
    End Class

    Private Class cls4
        Inherits cls2
    End Class

    Private Class cls3
        Sub foo(ByVal ParamArray Arg1() As cls2)
            m_i = 1
        End Sub

        Sub foo(ByVal ParamArray Arg1() As cls1)
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 310784: ")
        Try
            Dim c As New cls3()
            Dim o As Object = c
            Dim cls4Ary() As cls4
            Dim iEarly As Integer

            cls4Ary = New cls4() {}

            c.foo(cls4Ary)
            iEarly = m_i : m_i = -1

            o.foo(cls4Ary)
            PassFail(iEarly = m_i)

        Catch ex as Exception
            Failed(ex)
        end try

    End Sub
End Module



Module Bug310851
    Private m_i As Integer

    Private Class base
        Public Overloads Sub foo(Optional ByVal y As String = "", Optional ByVal x As String = "")
            m_i = 1
        End Sub
        Public Overloads Sub too(Optional ByVal y As Object = "", Optional ByVal x As String = "")
            m_i = 2
        End Sub
    End Class
    
    Private Class derived
        Inherits base
        Public Overloads Sub foo(ByVal ParamArray x() As String)
            m_i = 3
        End Sub
        Public Overloads Sub too(ByVal ParamArray x() As String)
            m_i = 4
        End Sub
    End Class

    Sub Test()
        Console.WriteLine("Bug 310851: ")

        Console.Write("   1) ")
        Try
            Dim c As New derived()
            Dim o As Object = c
            Dim iEarly As Integer

            c.foo()
            iEarly = m_i : m_i = -1

            o.foo()
            PassFail(iEarly = m_i)

        Catch ex as Exception
            Failed(ex)
        End Try

        Console.Write("   2) ")
        Try
            Dim c As New derived()
            Dim o As Object = c
            Dim iEarly As Integer

            c.too()
            iEarly = m_i : m_i = -1

            o.too()
            PassFail(iEarly = m_i)

        Catch ex as Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug310901
    Private m_i As Integer

    Private Class cls2
        Public Overloads Sub Foo(ByVal arg1 As Integer, Optional ByVal Arg2 As Short = 20)
            m_i = 1
        End Sub

        Public Overloads Sub Foo(ByVal ParamArray Ary As Integer())
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 310901: ")

        Try
            Dim c As New Cls2
            Dim o As Object = c
            Dim iEarly As Integer

            c.foo(20)
            iEarly = m_i : m_i = -1

            o.foo(20)
            PassFail(iEarly = m_i)

        Catch ex as Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug298655
    Private m_Arg1 As Integer
    Private m_Arg2 As Integer
    Private m_Arg3 As Integer

    Private Class Cls1
        Public Sub Goo(ByVal Arg1 As Integer, Optional ByVal Arg2 As Integer = 24, Optional ByVal Arg3 As Integer = 20)
            m_Arg1 = Arg1
            m_Arg2 = Arg2
            m_Arg3 = Arg3
        End Sub
    End Class

    Sub Test()
        Console.Write("Bug 298655: ")
        Try
            Dim c As New Cls1()
            Dim o As Object = c
            Dim iEarly1, iEarly2,iEarly3 as integer

            c.goo(20, , 40)
            iEarly1 = m_Arg1
            iEarly2 = m_Arg2
            iEarly3 = m_Arg3

            m_Arg1 = -1 : m_Arg2 = -1 : m_Arg3 = -1

            o.goo(20, , 40)		'this is giving the unexpected exception - BUG
            PassFail(iEarly1 = m_Arg1 AndAlso iEarly2 = m_Arg2 AndAlso iEarly3 = m_Arg3)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug311040
    Private m_i As Integer

    Private Class cls2
    End Class

    Private Interface I1
    End Interface

    Private Class cls3
        Inherits cls2
        Implements I1
    End Class

    Private Class cls1
        Sub zoo(ByVal arg As I1)
            m_i = 1
        End Sub

        Sub zoo(ByVal arg As Object)
            m_i = 2
        End Sub

        Sub foo(ByVal arg As I1)
            m_i = 3
        End Sub

        Sub foo(ByVal ParamArray arg() As cls2)
            m_i = 4
        End Sub

    End Class

    Sub Test()

        Console.WriteLine("Bug 311040")

        Console.Write("   1) ")
        Try
            Dim c As New cls1()
            Dim o As Object = c
            'c.foo(New cls3()) 'Ambiguous compiler error
            o.foo(New cls3())
            Failed()
        Catch ex As Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.Write("   2) ")
        Try
            Dim c As New cls1()
            Dim o As Object = c
            Dim iEarly As Integer
            c.zoo(New cls3()) 'Ambiguous compiler error
            iEarly = m_i : m_i = -1
            o.zoo(New cls3())

            PassFail(iEarly = m_i)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


Module Bug311133
    Private m_i As Integer

    Private Class cls2
        Public Overloads Sub Foo(ByVal ParamArray Arg1 As Long())
            m_i = 1
        End Sub

        Public Overloads Sub foo(ByVal Arg1 As Double, Optional ByVal arg2 As Long = 20)
            m_i = 2
        End Sub
    End Class

    Sub Test()

        Console.Write("Bug 311133: ")

        Try
            Dim c As New cls2()
            Dim o As Object = c
            Dim iEarly As Integer

            c.Foo(20, 40)
            iEarly = m_i : m_i = -1

            o.foo(20, 40)
            PassFail(iEarly = m_i)

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

End Module


Module Bug311283
    Private m_i As Integer

    Private Class cls2
        Public Overloads Sub Foo(ByVal arg As Long, ByVal Ary As e1())
            m_i = 1
        End Sub

        Public Overloads Sub Foo(ByVal arg As Long, ByVal Ary As E())
            m_i = 2
        End Sub
    End Class

    Enum E
        r
    End Enum

    Enum e1
        one
    End Enum

    Sub test()
        Console.Write("Bug 311283: ")
        Try
            Dim c As New cls2()
            Dim o As Object = c
            Dim iEarly As Integer

            c.Foo(20, New E() {})
            iEarly = m_i : m_i = -1

            o.foo(20, New E() {})  'this should Bind to Foo(String, E()), but is not doing so - BUG
            PassFail(m_i = iEarly)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module




Module Bug311285
    Private m_i As Integer

    Private Class cls2
        Public Overloads Sub Foo(ByVal arg As Long, ByVal Ary As e1())
            m_i = 1
        End Sub

        Public Overloads Sub Foo(ByVal arg As String, ByVal Ary As E())
            m_i = 2
        End Sub

        Public Overloads Sub Foo(ByVal arg As String, ByVal Ary As Integer())
            m_i = 3
        End Sub
    End Class

    Enum E
        r
    End Enum

    Enum e1
        one
    End Enum

    Sub test()
        Console.Write("Bug 311285: ")
        Try
            Dim c As New cls2()
            Dim o As Object = c
            Dim iEarly As Integer

            'c.Foo(20, New E() {})
            iEarly = m_i : m_i = -1

            o.foo(20, New E() {})  'this should Bind to Foo(String, E()), but is not doing so - BUG
            PassFail(m_i = iEarly)
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module


Module Bug311223
    Private m_i As Integer
    Private Class cls2
        Public Sub Foo(ByVal arg As Integer, ByVal ParamArray Ary As Object())
            m_i = 1
        End Sub
        Public Sub Zoo(ByVal arg As Integer, ByVal ParamArray Ary As Integer())
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.WriteLine("Bug 311223")

        Console.Write("   1) ")
        Try
            Dim c As New cls2()
            Dim o As Object = c
            Dim iEarly As Integer

            c.foo(20, Type.Missing, 40)
            iEarly = m_i : m_i = -1
            o.foo(20, , 40)
            PassFail(m_i = iEarly)
        Catch ex As Exception
            Failed(ex)
        End Try

        Console.Write("   2) ")
        Try
            Dim c As New cls2()
            Dim o As Object = c
            Dim iEarly As Integer

            c.foo(20, Type.Missing, Type.Missing, 40)
            iEarly = m_i : m_i = -1
            o.foo(20, , , 40)
            PassFail(m_i = iEarly)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub
End Module


#if 0
Module Bug311188
    Private m_i As Integer

    Private Class cls1
        Sub foo(ByVal z As String, Optional ByVal y As Object = 1)
            m_i = 1
        End Sub
        Sub foo(ByVal x As Object, ByVal ParamArray y As Object())
            m_i = 2
        End Sub
    End Class

    Sub Test()
        Console.WriteLine("Bug 311188")

        Console.Write("   1) ")
        Try
            Dim c As New cls1()
            Dim o As Object = c

            o.foo("c", Nothing)
            Failed()
        Catch ex As Reflection.AmbiguousMatchException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module
#End If

Module Bug337350
  Class ClsBug337350
    Class Cls1
        Function Cls1Func1(ByVal Arg As Object) As Object
            Return Arg + 10
        End Function

        Public Cls1Field1 As Integer

        Sub Cls1Sub1(ByRef Arg as Integer)
            Arg = Arg + 20
        End Sub
    End Class

    Interface I1
        Property I1Prop1(ByVal Arg As Integer) As Integer
    End Interface

    Class Cls2
        Implements IReflect
        Implements I1

        Dim obj As New Cls1
        Dim I1Prop1Info As PropertyInfo = GetType(I1).GetMember("I1Prop1")(0)
        Dim Cls1Func1Info As MethodInfo = GetType(Cls1).GetMember("Cls1Func1")(0)
        Dim Cls1Field1Info As FieldInfo = GetType(Cls1).GetMember("Cls1Field1")(0)
        Dim Cls1Sub1Info As MethodInfo = GetType(Cls1).GetMember("Cls1Sub1")(0)

        Public Function GetField(ByVal name As String, ByVal bindingAttr As System.Reflection.BindingFlags) As System.Reflection.FieldInfo Implements System.Reflection.IReflect.GetField
        End Function

        Public Function GetFields(ByVal bindingAttr As System.Reflection.BindingFlags) As System.Reflection.FieldInfo() Implements System.Reflection.IReflect.GetFields
        End Function

        Public Function GetMember(ByVal name As String, ByVal bindingAttr As System.Reflection.BindingFlags) As System.Reflection.MemberInfo() Implements System.Reflection.IReflect.GetMember
            Dim m As MemberInfo
            Select Case name
                Case "Cls1Func1"
                    m = Cls1Func1Info
                Case "Cls1Field1"
                    m = Cls1Field1Info
                Case "I1Prop1"
                    m = I1Prop1Info
                Case "Cls1Sub1"
                    m = Cls1Sub1Info
            End Select

            If m Is Nothing Then
                Return New MemberInfo() {}
            Else
                Return New MemberInfo() {m}
            End If
        End Function

        Public Function GetMembers(ByVal bindingAttr As System.Reflection.BindingFlags) As System.Reflection.MemberInfo() Implements System.Reflection.IReflect.GetMembers

        End Function

        Public Overloads Function GetMethod(ByVal name As String, ByVal bindingAttr As System.Reflection.BindingFlags) As System.Reflection.MethodInfo Implements System.Reflection.IReflect.GetMethod

        End Function

        Public Overloads Function GetMethod1(ByVal name As String, ByVal bindingAttr As System.Reflection.BindingFlags, ByVal binder As System.Reflection.Binder, ByVal types() As System.Type, ByVal modifiers() As System.Reflection.ParameterModifier) As System.Reflection.MethodInfo Implements System.Reflection.IReflect.GetMethod

        End Function

        Public Function GetMethods(ByVal bindingAttr As System.Reflection.BindingFlags) As System.Reflection.MethodInfo() Implements System.Reflection.IReflect.GetMethods

        End Function

        Public Function GetProperties(ByVal bindingAttr As System.Reflection.BindingFlags) As System.Reflection.PropertyInfo() Implements System.Reflection.IReflect.GetProperties

        End Function

        Public Overloads Function GetProperty(ByVal name As String, ByVal bindingAttr As System.Reflection.BindingFlags) As System.Reflection.PropertyInfo Implements System.Reflection.IReflect.GetProperty

        End Function

        Public Overloads Function GetProperty1(ByVal name As String, ByVal bindingAttr As System.Reflection.BindingFlags, ByVal binder As System.Reflection.Binder, ByVal returnType As System.Type, ByVal types() As System.Type, ByVal modifiers() As System.Reflection.ParameterModifier) As System.Reflection.PropertyInfo Implements System.Reflection.IReflect.GetProperty

        End Function

        Public Function InvokeMember(ByVal name As String, ByVal invokeAttr As System.Reflection.BindingFlags, ByVal binder As System.Reflection.Binder, ByVal target As Object, ByVal args() As Object, ByVal modifiers() As System.Reflection.ParameterModifier, ByVal culture As System.Globalization.CultureInfo, ByVal namedParameters() As String) As Object Implements System.Reflection.IReflect.InvokeMember
            If Not (target Is Me) Then
                Exit Function
            End If

            Select Case name
                Case "Cls1Func1"
                    Dim m As MethodInfo = binder.BindToMethod(invokeAttr, New MethodBase() {Cls1Func1Info}, Nothing, Nothing, Nothing, Nothing, Nothing)
                    Return m.Invoke(obj, args)
                Case "Cls1Field1"
                    Dim f As FieldInfo = binder.BindToField(invokeAttr, New FieldInfo() {Cls1Field1Info}, Nothing, Nothing)
                    If invokeAttr And BindingFlags.GetField Then
                        Return f.GetValue(obj)
                    ElseIf invokeAttr And BindingFlags.SetField Then
                        f.SetValue(obj, args(0))
                    End If
                Case "Cls1Sub1"
                    Dim m As MethodInfo = binder.BindToMethod(invokeAttr, New MethodBase() {Cls1Sub1Info}, args, Nothing, Nothing, Nothing, Nothing)
                    Return m.Invoke(obj, args)
            End Select
        End Function

        Public Overloads ReadOnly Property UnderlyingSystemType() As System.Type Implements System.Reflection.IReflect.UnderlyingSystemType
            Get

            End Get
        End Property

        Private m_Prop1 As Integer
        Public Property Prop1(ByVal Arg As Integer) As Integer Implements I1.I1Prop1
            Get
                Return m_Prop1
            End Get
            Set(ByVal Value As Integer)
                m_Prop1 = Value
            End Set
        End Property
    End Class
  End Class

    Sub Test()
        Console.WriteLine("Bug 337350")

      Try
        Console.Write("   1) ")
        Dim t As Type = GetType(ClsBug337350)
        Dim obj As Object = t
        If (CObj(t).Name = "ClsBug337350") Then
           Passed()
        Else
           Failed()
        End If


        Console.Write("   2) ")
        obj = New ClsBug337350.Cls2
        If (obj.Cls1Func1(80) = 90) Then
            Passed()
        Else
            Failed()
        End If

        Console.Write("   3) ")
        obj.I1Prop1(2) = 200
        If  (obj.I1Prop1(2) = 200) Then
            Passed()
        Else
            Failed()
        End If

        Console.Write("   4) ")        
        obj.Cls1Field1 = 400
        If (obj.Cls1Field1 = 400) Then
           Passed()
        Else
           Failed()
        End If

        Console.Write("   5) ")
        Dim InOut As Integer = 500
        obj.Cls1Sub1(InOut)
        If (InOut = 520) Then
           Passed()
        Else
           Failed()
        End If

      Catch
          Failed()
      End Try
    End Sub
End Module

Module BugVSW32724

    class base
        public x as integer
    end class

    class derived1
        inherits base
        shadows public default readonly property x(byval i as integer) as integer
            get
                return 10
            end get
        end property
    end class

    class derived2
        inherits base
        shadows public default writeonly property x(byval i as integer) as integer
            set(value as integer)
                Console.Write("   2) ")
                If i = 10 AndAlso value = 3 Then Passed()
            end set
        end property
    end class

    sub Test
        Try
            Console.WriteLine("Bug VSW32724")
            dim x as object = new derived1
            Console.Write("   1) ")
            If x(10) = 10 Then Passed() Else Failed()
            x = new derived2
            x(10) = 3
        Catch ex As Exception
            Failed(ex)
        End Try
    end sub

End Module

Namespace BugVSW32699

Module module1
    Class c1

        Class base
        End Class
        Class derived
            Inherits base
        End Class

        Sub foo(ByVal ParamArray j() As derived)
            Console.WriteLine("2")
        End Sub

        Sub foo(ByVal ParamArray i() As base)
            Console.WriteLine("1")
        End Sub

    End Class

    Sub Test()
        Console.WriteLine("Bug VSW32699")
        Dim c As New c1()
        Dim o As Object = c
        Try
            o.foo()
            Failed()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Passed()
        End Try
    End Sub

End Module

End Namespace

Namespace BugVSW32809

Module Module1

    Structure s1

        Dim z As Integer

        Sub foo(ByVal i As Integer)
        End Sub

        Sub foo(ByVal i As Double)
        End Sub

        Sub loo(ByVal i As Double)
        End Sub

        Structure s2
            Dim z As Integer

            Sub goo()
                Dim o As Object = 5
                Try
                    foo(o)
                    Failed
                Catch ex As Exception
                    Console.Writeline(ex.message)
                    Passed
                end try
            End Sub
        End Structure

    End Structure

    Sub Test()
        Dim s As s1.s2
        s.goo()
    End Sub

End Module

End Namespace
