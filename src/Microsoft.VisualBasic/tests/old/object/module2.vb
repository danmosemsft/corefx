Option Strict Off

Imports System
Imports Microsoft.VisualBasic
Imports Conversion.Module2.VType
Imports System.Reflection

Namespace Conversion

    Module Module2

        Private ReadOnly SkipType As Boolean()
        Private ReadOnly ValueTable()() As Object

        Friend Enum VType
            t_emp
            t_bool
            t_i1
            t_ui1
            t_i2
            t_ui2
            t_i4
            t_ui4
            t_i8
            t_ui8
            t_dec
            t_r4
            t_r8
            t_char
            t_str
            t_date
            t_obj
            t_chararray
            t_maxNoEnum
            t_enumSByte = t_maxNoEnum
            t_enumByte
            t_enumShort
            t_enumUShort
            t_enumInteger
            t_enumUInteger
            t_enumLong
            t_enumULong
            t_max
        End Enum

        Enum Operators
            Add
            Subtract
            Divide
            IDivide
            Multiply
            [Mod]
            LessThan
            LessThanOrEqual
            Equal
            GreaterThan
            GreaterThanOrEqual
            [Or]
            [And]
            [Xor]
            [Not]
            Negate
            Plus
            Pow
            [AndAlso]
            [OrElse]
            ShiftLeft
            ShiftRight
        End Enum

        Class TestClass
            Implements IFormattable

            Public Function IFormattableToString( format As String, formatProvider As IFormatProvider ) As String Implements IFormattable.ToString
                Return "TestClass"
            End Function

            Overrides Function ToString() As String
                Return "TestClass"
            End Function
        End Class

        Enum SByteEnum As SByte
            Zero = 0
            One = 1
            MinValue = SByte.MinValue
            MaxValue = SByte.MaxValue
        End Enum

        Enum ByteEnum As Byte
            Zero = 0
            One = 1
            MinValue = Byte.MinValue 'Same as Zero
            MaxValue = Byte.MaxValue
        End Enum

        Enum ShortEnum As Short
            Zero = 0
            One = 1
            MinValue = Short.MinValue
            MaxValue = Short.MaxValue
        End Enum

        Enum UShortEnum As UShort
            Zero = 0
            One = 1
            MinValue = UShort.MinValue 'Same as Zero
            MaxValue = UShort.MaxValue
        End Enum

        Enum IntegerEnum As Integer
            Zero = 0
            One = 1
            MinValue = Integer.MinValue
            MaxValue = Integer.MaxValue
        End Enum

        Enum UIntegerEnum As UInteger
            Zero = 0
            One = 1
            MinValue = UInteger.MinValue 'Same as Zero
            MaxValue = UInteger.MaxValue
        End Enum

        Enum LongEnum As Long
            Zero = 0
            One = 1
            MinValue = Long.MinValue
            MaxValue = Long.MaxValue
        End Enum

        Enum ULongEnum As ULong
            Zero = 0
            One = 1
            MinValue = ULong.MinValue 'Same as Zero
            MaxValue = ULong.MaxValue
        End Enum


        Sub New()

            ValueTable = New Object(t_max - 1)() { _
                New Object() {Nothing}, _
                New Object() {False, True}, _
                New Object() {CSByte(0), CSByte(1), SByte.MinValue, SByte.MaxValue}, _
                New Object() {CByte(0), CByte(1), Byte.MaxValue}, _
                New Object() {CShort(0), CShort(1), Int16.MinValue, Int16.MaxValue}, _
                New Object() {CUShort(0), CUShort(1), UInt16.MaxValue}, _
                New Object() {CInt(0), CInt(1), Int32.MinValue, Int32.MaxValue}, _
                New Object() {CUInt(0), CUInt(1), UInt32.MaxValue}, _
                New Object() {CLng(0), CLng(1), Int64.MinValue, Int64.MaxValue}, _
                New Object() {CULng(0), CULng(1), UInt64.MaxValue}, _
                New Object() {CDec(0), CDec(1), Decimal.MinValue, Decimal.MaxValue}, _
                New Object() {CSng(0), CSng(1), Single.MinValue, Single.MaxValue}, _
                New Object() {CDbl(0), CDbl(1), Double.MinValue, Double.MaxValue}, _
                New Object() {ChrW(0), ChrW(1), Char.MaxValue}, _
                New Object() {CStr(0), CStr(1), "&H0", "&H1", CStr(Int32.MinValue), CStr(Int32.MaxValue)}, _
                New Object() {New DateTime(0), New DateTime(1, 1, 1, 12, 0, 0), DateTime.MinValue, DateTime.MaxValue}, _
                New Object() {New TestClass}, _
                New Object() {CType("255", Char())}, _
                New Object() {SByteEnum.One}, _
                New Object() {ByteEnum.One}, _
                New Object() {ShortEnum.One}, _
                New Object() {UShortEnum.One}, _
                New Object() {IntegerEnum.One}, _
                New Object() {UIntegerEnum.One}, _
                New Object() {LongEnum.One}, _
                New Object() {ULongEnum.One} _
            }

            SkipType = New Boolean(t_max - 1) {}

            SkipType(t_emp) = False
            SkipType(t_bool) = False
            SkipType(t_i1) = False
            SkipType(t_ui1) = False
            SkipType(t_i2) = False
            SkipType(t_ui2) = False
            SkipType(t_i4) = False
            SkipType(t_ui4) = False
            SkipType(t_i8) = False
            SkipType(t_ui8) = False
            SkipType(t_dec) = False
            SkipType(t_r4) = False
            SkipType(t_r8) = False
            SkipType(t_char) = False
            SkipType(t_str) = False
            SkipType(t_date) = False
            SkipType(t_obj) = False
            SkipType(t_chararray) = True
            SkipType(t_enumSByte) = False
            SkipType(t_enumByte) = False
            SkipType(t_enumShort) = False
            SkipType(t_enumUShort) = False
            SkipType(t_enumInteger) = False
            SkipType(t_enumUInteger) = False
            SkipType(t_enumLong) = False
            SkipType(t_enumULong) = False

        End Sub

        Private Function SkipUndefinedOperator(ByVal vartype As VType, op As Operators) As Boolean

            If vartype = t_char OrElse vartype = t_date Then

                Select Case op
                    Case Operators.LessThan, _
                         Operators.LessThanOrEqual, _
                         Operators.GreaterThan, _
                         Operators.GreaterThanOrEqual, _
                         Operators.Equal, _
                         Operators.Add
                        Return False
    
                    Case Else
                        Return True

                End Select
            End If

            if vartype = t_obj Then Return True

            Return False

        End Function


        Private Function OpSymbol(op As Operators) As String

            Select Case op
                Case Operators.Add
                    Return "+"

                Case Operators.Subtract
                    Return "-"

                Case Operators.Divide
                    Return "/"

                Case Operators.IDivide
                    Return "\"

                Case Operators.Multiply
                    Return "*"

                Case Operators.Mod
                    Return "Mod"

                Case Operators.LessThan
                    Return "<"

                Case Operators.LessThanOrEqual
                    Return "<="

                Case Operators.GreaterThan
                    Return ">"

                Case Operators.GreaterThanOrEqual
                    Return ">="

                Case Operators.Equal
                    Return "="

                Case Operators.Or
                    Return "Or"

                Case Operators.And
                    Return "And"

                Case Operators.Xor
                    Return "XOr"

                Case Operators.Not
                    Return "Not"

                Case Operators.Negate
                    Return "-"

                Case Operators.Plus
                    Return "+"

                Case Operators.Pow
                    Return "^"

                Case Operators.ShiftLeft
                    Return "<<"

                Case Operators.ShiftRight
                    Return ">>"

                Case Else
                    Return "UNKNOWN"

            End Select

        End Function

        Private Function VTypeFromTypeCode(ByVal typ As TypeCode) As VType

            Select Case typ

                Case TypeCode.Boolean
                    Return t_bool
                Case TypeCode.Byte
                    Return t_ui1
                Case TypeCode.Int16
                    Return t_i2
                Case TypeCode.Int32
                    Return t_i4
                Case TypeCode.Int64
                    Return t_i8
                Case TypeCode.Decimal
                    Return t_dec
                Case TypeCode.Single
                    Return t_r4
                Case TypeCode.Double
                    Return t_r8
                Case TypeCode.Char
                    Return t_char
                Case TypeCode.String
                    Return t_str
                Case TypeCode.DateTime
                    Return t_date
                Case Else
                    Return t_emp
            End Select

        End Function

        Private Function TypeCodeFromVType(ByVal vartyp As VType) As TypeCode

            Select Case vartyp

                Case t_bool
                    Return TypeCode.Boolean
                Case t_ui1
                    Return TypeCode.Byte
                Case t_i2
                    Return TypeCode.Int16
                Case t_i4
                    Return TypeCode.Int32
                Case t_i8
                    Return TypeCode.Int64
                Case t_dec
                    Return TypeCode.Decimal
                Case t_r4
                    Return TypeCode.Single
                Case t_r8
                    Return TypeCode.Double
                Case t_char
                    Return TypeCode.Char
                Case t_str
                    Return TypeCode.String
                Case t_date
                    Return TypeCode.DateTime
                Case t_enumByte
                    Return TypeCode.Byte
                Case t_enumShort
                    Return TypeCode.Int16
                Case t_enumInteger
                    Return TypeCode.Int32
                Case t_enumLong
                    Return TypeCode.Int64
                Case Else
                    Return TypeCode.Object
            End Select

        End Function

        Private Function VTypeName(ByVal vartyp As VType) As String

            Select Case vartyp

                Case t_emp
                    Return "Nothing"
                Case t_bool
                    Return "Boolean"
                Case t_i1
                    Return "SByte"
                Case t_ui1
                    Return "Byte"
                Case t_i2
                    Return "Short"
                Case t_ui2
                    Return "UShort"
                Case t_i4
                    Return "Integer"
                Case t_ui4
                    Return "UInteger"
                Case t_i8
                    Return "Long"
                Case t_ui8
                    Return "ULong"
                Case t_dec
                    Return "Decimal"
                Case t_r4
                    Return "Single"
                Case t_r8
                    Return "Double"
                Case t_char
                    Return "Char"
                Case t_str
                    Return "String"
                Case t_date
                    Return "DateTime"
                Case t_enumSByte
                    Return "enum(SByte)"
                Case t_enumByte
                    Return "enum(Byte)"
                Case t_enumShort
                    Return "enum(Short)"
                Case t_enumUShort
                    Return "enum(UShort)"
                Case t_enumInteger
                    Return "enum(Integer)"
                Case t_enumUInteger
                    Return "enum(UInteger)"
                Case t_enumLong
                    Return "enum(Long)"
                Case t_enumULong
                    Return "enum(ULong)"
                Case Else
                    Return "Object"
            End Select

        End Function


        Private Function TypeMismatch(o As Object, typ As VType) As Boolean
            If o Is Nothing Then
                Return typ <> t_emp
            Else
                Return VTypeFromTypeCode(System.Type.GetTypeCode(o.GetType()) = typ)
            End If
        End Function


        Private Sub VerifyTables()

            Dim LeftType As VType
            Dim o As Object
            Dim LeftValueIndex As Integer

            For LeftType = VType.t_emp To VType.t_obj

                For LeftValueIndex = 0 To ValueTable(LeftType).GetUpperBound(0)

                    o = ValueTable(LeftType)(LeftValueIndex)
                    if TypeMismatch(o, LeftType) Then
                        Console.WriteLine("Type mismatch in table: ValueType( " & VTypeName(LeftType) & ", " & LeftValueIndex & " )")
                    End If

                Next LeftValueIndex

            Next LeftType

        End Sub


        Sub MathTests()
            VerifyTables
            DoTwoOperandTest(Operators.Subtract)
            DoTwoOperandTest(Operators.Add)
            DoTwoOperandTest(Operators.Multiply)
            DoTwoOperandTest(Operators.Divide)
            DoTwoOperandTest(Operators.IDivide)
            DoTwoOperandTest(Operators.Mod)
            DoTwoOperandTest(Operators.LessThan)
            DoTwoOperandTest(Operators.LessThanOrEqual)
            DoTwoOperandTest(Operators.Equal)
            DoTwoOperandTest(Operators.GreaterThan)
            DoTwoOperandTest(Operators.GreaterThanOrEqual)

            DoTwoOperandTest(Operators.Pow)

            DoTwoOperandTest(Operators.Or)
            DoTwoOperandTest(Operators.And)
            DoTwoOperandTest(Operators.Xor)

            DoTwoOperandTest(Operators.ShiftLeft)
            DoTwoOperandTest(Operators.ShiftRight)

            DoOneOperandTest(Operators.Negate)
            DoOneOperandTest(Operators.Plus)
            DoOneOperandTest(Operators.Not)
        End Sub


        Sub DoTwoOperandTest(op As Operators)

            Dim LeftType, RightType As VType
            Dim LeftValue, RightValue, ResultValue As Object
            Dim LeftValueIndex, RightValueIndex As Integer
            Dim LeftValueTmp1, RightValueTmp1 As Object
            Dim conv As IConvertible

            Console.WriteLine("'****************************************************************")
            Console.WriteLine("'***  " & op.ToString())
            Console.WriteLine("'****************************************************************")

            Const MinType As VType = VType.t_emp
            Dim MaxType As VType

            Select Case op
                Case Operators.And, _
                     Operators.Or, _
                     Operators.Xor
                    MaxType = t_max - 1

                Case Else
                    MaxType = t_maxNoEnum - 1

            End Select

            Try
                For LeftType = MinType To MaxType

                    If SkipType(LeftType) Then Continue For

                    For RightType = MinType To MaxType

                        If SkipType(RightType) Then Continue For
                        Dim CompareOnce As Boolean = False

                        Console.WriteLine(VTypeName(LeftType) & " " & OpSymbol(op) & " " & VTypeName(RightType))

                        For LeftValueIndex = 0 To ValueTable(LeftType).GetUpperBound(0)

                            'For operators that aren't defined for the current type, just use the first value.
                            If SkipUndefinedOperator(LeftType, op) Then CompareOnce = True

                            LeftValue = ValueTable(LeftType)(LeftValueIndex)

                            If LeftType = t_str Then
                                LeftValueTmp1 = ChrW(34) & LeftValue & ChrW(34)
                            ElseIf LeftType = t_date Then
                                LeftValueTmp1 = "#" & DirectCast(LeftValue, DateTime).ToString("MM/dd/yyyy hh:mm:ss tt") & "#"
                            ElseIf LeftType = t_char Then
                                LeftValueTmp1 = "ChrW(" & AscW(LeftValue) & ")"
                            ElseIf LeftType = t_emp Then
                                LeftValueTmp1 = "Nothing"
                            Else
                                LeftValueTmp1 = LeftValue
                            End If

                            For RightValueIndex = 0 To ValueTable(RightType).GetUpperBound(0)

                                'For operators that aren't defined for the current type, just use the first value.
                                If SkipUndefinedOperator(RightType, op) Then CompareOnce = True

                                RightValue = ValueTable(RightType)(RightValueIndex)

                                Try
                                    If RightType = t_str Then
                                        RightValueTmp1 = ChrW(34) & RightValue & ChrW(34)
                                    ElseIf RightType = t_date Then
                                        RightValueTmp1 = "#" & DirectCast(RightValue, DateTime).ToString("MM/dd/yyyy hh:mm:ss tt") & "#"
                                    ElseIf RightType = t_char Then
                                        RightValueTmp1 = "ChrW(" & AscW(RightValue) & ")"
                                    ElseIf RightType = t_emp Then
                                        RightValueTmp1 = "Nothing"
                                    Else
                                        RightValueTmp1 = RightValue
                                    End If

                                    Select Case op

                                        Case Operators.Add
                                            ResultValue = (LeftValue + RightValue)

                                        Case Operators.Subtract
                                            ResultValue = (LeftValue - RightValue)

                                        Case Operators.Divide
                                            ResultValue = (LeftValue / RightValue)

                                        Case Operators.IDivide
                                            ResultValue = (LeftValue \ RightValue)

                                        Case Operators.Multiply
                                            ResultValue = (LeftValue * RightValue)

                                        Case Operators.Mod
                                            ResultValue = (LeftValue Mod RightValue)

                                        Case Operators.LessThan
                                            ResultValue = (LeftValue < RightValue)

                                        Case Operators.LessThanOrEqual
                                            ResultValue = (LeftValue <= RightValue)

                                        Case Operators.GreaterThan
                                            ResultValue = (LeftValue > RightValue)

                                        Case Operators.GreaterThanOrEqual
                                            ResultValue = (LeftValue >= RightValue)

                                        Case Operators.Equal
                                            ResultValue = (LeftValue = RightValue)

                                        Case Operators.Or
                                            ResultValue = (LeftValue Or RightValue)

                                        Case Operators.And
                                            ResultValue = (LeftValue And RightValue)

                                        Case Operators.Xor
                                            ResultValue = (LeftValue Xor RightValue)

                                        Case Operators.Pow
                                            ResultValue = (LeftValue ^ RightValue)

                                        Case Operators.ShiftLeft
                                            ResultValue = (LeftValue << RightValue)

                                        Case Operators.ShiftRight
                                            ResultValue = (LeftValue >> RightValue)

                                        Case Else
                                            Throw New ArgumentException
                                    End Select

                                    Console.WriteLine("    ( {0} {1} {2} ) = {3} '({4})", LeftValueTmp1, OpSymbol(op), RightValueTmp1, ResultValue, TypeName(ResultValue))

                                Catch ex As InvalidCastException
                                    Console.WriteLine("    ( {0} {1} {2} ) = {3}", LeftValueTmp1, OpSymbol(op), RightValueTmp1, ex.GetType().Name)

                                Catch ex As AmbiguousMatchException
                                    Console.WriteLine("    ( {0} {1} {2} ) = {3}", LeftValueTmp1, OpSymbol(op), RightValueTmp1, ex.GetType().Name)

                                Catch ex As DivideByZeroException
                                    Console.WriteLine("    ( {0} {1} {2} ) = {3}", LeftValueTmp1, OpSymbol(op), RightValueTmp1, ex.GetType().Name)

                                Catch ex As OverflowException
                                    Console.WriteLine("    ( {0} {1} {2} ) = {3}", LeftValueTmp1, OpSymbol(op), RightValueTmp1, ex.GetType().Name)

                                Catch ex As ArgumentException
                                    Console.WriteLine("    ( {0} {1} {2} ) = {3}", LeftValueTmp1, OpSymbol(op), RightValueTmp1, ex.GetType().Name)

                                Catch ex As Exception
                                    Console.WriteLine("    ( {0} {1} {2} ) = {3} {4}", LeftValueTmp1, OpSymbol(op), RightValueTmp1, ex.GetType().Name, ex.Message)
                                End Try

				If CompareOnce Then Goto NextType
                            Next RightValueIndex
                        Next LeftValueIndex
NextType:
                    Next RightType

                Next LeftType

            Catch ex As Exception
                'Failed(ex)
            End Try

        End Sub


        Sub DoOneOperandTest(op As Operators)

            Dim LeftType As VType
            Dim LeftValue, ResultValue As Object
            Dim LeftValueIndex As Integer
            Dim LeftValueTmp1 As Object
            Dim conv As IConvertible

            Console.WriteLine("'****************************************************************")
            Console.WriteLine("'***  " & op.ToString())
            Console.WriteLine("'****************************************************************")

            Const MinType As VType = VType.t_emp
            Dim MaxType As VType

            Select Case op
                Case Operators.Not
                    MaxType = t_max - 1

                Case Else
                    MaxType = t_maxNoEnum - 1

            End Select


            Try
                For LeftType = MinType To MaxType

                    If SkipType(LeftType) Then Continue For

                    Console.WriteLine(OpSymbol(op) & " " & VTypeName(LeftType))

                    For LeftValueIndex = 0 To ValueTable(LeftType).GetUpperBound(0)

                        'For operators that aren't defined for the current type, just use the first value.
                        If LeftValueIndex > 0 Andalso SkipUndefinedOperator(LeftType, op) Then Continue For

                        LeftValue = ValueTable(LeftType)(LeftValueIndex)

                        If LeftType = t_str Then
                            LeftValueTmp1 = ChrW(34) & LeftValue & ChrW(34)
                        ElseIf LeftType = t_date Then
                            LeftValueTmp1 = "#" & DirectCast(LeftValue, DateTime).ToString("MM/dd/yyyy hh:mm:ss tt") & "#"
                        ElseIf LeftType = t_char Then
                            LeftValueTmp1 = "ChrW(" & AscW(LeftValue) & ")"
                        ElseIf LeftType = t_emp Then
                            LeftValueTmp1 = "Nothing"
                        Else
                            LeftValueTmp1 = LeftValue
                        End If

                        Try
                            Select Case op
                                Case Operators.Negate
                                    ResultValue = - LeftValue

                                Case Operators.Plus
                                    ResultValue = + LeftValue

                                Case Operators.Not
                                    ResultValue = Not LeftValue

                                Case Else
                                    Throw New ArgumentException()
                            End Select

                            Console.WriteLine("    ( {0} {1} ) = {2} '({3})", OpSymbol(op), LeftValueTmp1, ResultValue, TypeName(ResultValue))

                        Catch ex As InvalidCastException
                            Console.WriteLine("    ( {0} {1} ) = {2}", OpSymbol(op), LeftValueTmp1, ex.GetType().Name)

                        Catch ex As AmbiguousMatchException
                            Console.WriteLine("    ( {0} {1} ) = {2}", OpSymbol(op), LeftValueTmp1, ex.GetType().Name)

                        Catch ex As DivideByZeroException
                            Console.WriteLine("    ( {0} {1} ) = {2}", LeftValueTmp1, OpSymbol(op), ex.GetType().Name)

                        Catch ex As OverflowException
                            Console.WriteLine("    ( {0} {1} ) = {2}", LeftValueTmp1, OpSymbol(op), ex.GetType().Name)

                        Catch ex As ArgumentException
                            Console.WriteLine("    ( {0} {1} ) = {2}", LeftValueTmp1, OpSymbol(op), ex.GetType().Name)

                        Catch ex As Exception
                            Console.WriteLine("    ( {0} {1} ) = {2} {3}", OpSymbol(op), LeftValueTmp1, ex.GetType().Name, ex.Message)
                        End Try

                    Next LeftValueIndex

                Next LeftType

            Catch ex As Exception
                'Failed(ex)
            End Try

        End Sub

    End Module

End Namespace
