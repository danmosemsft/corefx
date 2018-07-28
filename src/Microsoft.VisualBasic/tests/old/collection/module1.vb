Imports System
Imports System.Collections
Imports Microsoft.VisualBasic

Module Test

    Dim b As Boolean
    Dim i As Integer
    Dim s As String
    Dim TestNumber As Integer
    Dim TestCounter As Integer
    Dim ch As Char



    Sub Main()
        Console.WriteLine("Begin Tests")
        Collection_Add.Main()
        Collection_Count.Main()
        Collection_Item.Main()
        Collection_Remove.Main()
        Collection_IList()
        Collection_ICollection()
        'Regression tests
        Bug254032.Test()
        Console.WriteLine("End Tests")
    End Sub


    Sub Collection_ICollection()
        Dim c As New Collection()
        Dim intf As ICollection
        Dim i, j, index As Integer
        Dim bFailed As Boolean

        Console.WriteLine("----ICollection----")
        
        Try
            For i = 1 To 9
                c.Add("Data" & i, "Key" & i)
            Next i

            intf = c

            'Allocate a few extra slots
            Dim aryObject As Object() = New Object(intf.Count + 1) { }
            Dim aryString As String() = New String(intf.Count + 1) { }
            Dim strData As String() = New String(intf.Count) { }

            Console.Write("ICollection.Count : ")
            Try
                if intf.Count = 9 Then
                    Passed()
                Else
                    Failed()
                End If
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("ICollection.IsSynchronized : ")
            Try
                if intf.IsSynchronized Then
                    'VB Collection is not synchronized
                    Failed()
                Else
                    Passed()
                End If
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("ICollection.SyncRoot : ")
            Try
                if intf.SyncRoot Is c Then
                    Passed()
                Else
                    'SyncRoot should match collection ref
                    Failed()
                End If
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.WriteLine("ICollection.CopyTo( )")

            Console.Write("    ICollection.CopyTo() Null argument : ")
            Try
                intf.CopyTo(Nothing, 0)
            Catch ex As ArgumentNullException
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("    ICollection.CopyTo( New Object(1,1) { }, 0 ) : ")
            Try
                intf.CopyTo(New Object(1,1) { }, 0)
            Catch ex As ArgumentException
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("    ICollection.CopyTo( Object(), -1 ) : ")
            Try
                intf.CopyTo(aryObject, -1)
            Catch ex As ArgumentException
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("    ICollection.CopyTo( Object(), 5 ) : ")
            Try
                intf.CopyTo(aryObject, 5)
            Catch ex As ArgumentException
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

            bFailed = False
            Console.Write("    ICollection.CopyTo( Object(), 1 ) : ")
            Try
                intf.CopyTo(aryObject, 1)
                If (Not aryObject(0) Is Nothing) Then
                    Failed()
                Else
                    For i = 1 to 9 
                        If TypeOf aryObject(i) Is String Then
                            Dim str As String = DirectCast(aryObject(i), String)
                            if (str <> "Data" & i) Then
                                bFailed = True
                            End If
                        Else
                            bFailed = True
                        End If
                    Next i
                    If (Not aryObject(10) Is Nothing) Then
                        Failed()
                    End If
                    If bFailed Then
                        Failed()
                    Else
                        Passed()
                    End If
                End If
            Catch ex As Exception
                Failed(ex)
            End Try

            bFailed = False
            Console.Write("    ICollection.CopyTo( String(), 1 ) : ")
            Try
                intf.CopyTo(aryString, 1)
                If (Not aryString(0) Is Nothing) Then
                    Failed()
                Else
                    For i = 1 to 9 
                        if (aryString(i) <> "Data" & i) Then
                            bFailed = True
                        End If
                    Next i
                    If (Not aryString(10) Is Nothing) Then
                        Failed()
                    End If
                    If bFailed Then
                        Failed()
                        DumpIList(aryString)
                    Else
                        Passed()
                    End If
                End If
            Catch ex As Exception
                Failed(ex)
            End Try

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub


    Sub Collection_IList()
        Dim c As New Collection()
        Dim l As IList
        Dim i, j, index As Integer
        Dim bFailed As Boolean

        Console.WriteLine("----IList----")
        Try

            Console.WriteLine("Adding 9 1-based collection items")
            For i = 1 To 9
                c.Add("Data" & i, "Key" & i)
                Console.WriteLine("      Item(" & i & ")=" & CStr(c.Item(i)))
            Next i

            Console.WriteLine("Converting collection to a 0-based list")
            l = CType(c, IList)

            'IList.Insert
            Console.WriteLine("   Inserting After 2")
            l.Insert(2, "Inserted")
            For i = 0 To l.Count - 1
                Console.WriteLine("      Item(" & i & ")=" & CStr(l.Item(i)))
            Next i

            'IList.RemoveAt
            Console.Write("   IList.RemoveAt(5): ")
            Try
                l.RemoveAt(5)
                index = l.IndexOf("Data6")
                bFailed = Not (l.Count = 9 AndAlso index = 5 AndAlso (l.IndexOf("Data5") = -1))
                PassFail(Not bFailed)
                If bFailed Then
                    DumpIList(l)
                End If
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("   IList.RemoveAt() all items: ")
            Try
                For i = l.Count - 1 To 0 Step -1
                    l.RemoveAt(i)
                Next i
                If l.Count = 0 Then
                    Passed()
                Else
                    Failed()
                    DumpIList(l)
                End If
            Catch ex As Exception
                Failed(ex)
            End Try

            'IList.Clear
            Console.Write("   IList.Clear (empty list): ")
            Try
                l.Clear()
                PassFail(l.Count = 0)
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("   IList.Clear: ")
            Try
                l.Add("something to delete1")
                l.Add("something to delete2")
                l.Clear()
                PassFail(l.Count = 0)
            Catch ex As Exception
                Failed(ex)
            End Try

            'IList.Add
            Console.Write("   IList.Add: ")
            Try
                For i = 0 To 3
                    l.Add("IList.Add" & CStr(i))
                Next i
                PassFail(l.Count = 4)
            Catch ex As Exception
                Failed(ex)
            End Try

            Console.Write("   IList.Contains: ")
            Try
                bFailed = False
                For i = 0 To 3
                    If Not l.Contains("IList.Add" & CStr(i)) Then
                        bFailed = True
                    End If
                Next i
                If bFailed Then
                    Failed()
                    DumpIList(l)
                Else
                    Passed()
                End If
            Catch ex As Exception
                Failed(ex)
            End Try

            'IList.IndexOf
            Console.Write("   IList.IndexOf: ")
            Try
                Dim jSum As Integer = 0
                bFailed = False
                For i = 0 To 3
                    j = l.IndexOf("IList.Add" & CStr(i))
                    If j < 0 OrElse j > 3 Then
                        bFailed = True
                    End If
                    jSum += j
                Next i
                'Sum of index check
                If jSum <> 6 Then
                    bFailed = True
                End If

                If l.IndexOf("IList.Add39") <> -1 Then
                    bFailed = True
                End If
                If bFailed Then
                    Failed()
                    DumpIList(l)
                Else
                    Passed()
                End If
            Catch ex As Exception
                Failed(ex)
            End Try


            'IList.Remove
            Console.Write("   IList.Remove: ")
            Try
                l.Remove("IList.Add" & CStr(0))
                l.Remove("IList.Add" & CStr(3))
                l.Remove("IList.Add" & CStr(1))
                l.Remove("IList.Add" & CStr(2))
                If (l.Count = 0) Then
                    Passed()
                Else
                    DumpIList(l)
                End If
            Catch ex As Exception
                Failed(ex)
            End Try

            'IList.IsFixedSize
            Console.Write("   IList.IsFixedSize: ")
            Try
                PassFail(Not l.IsFixedSize)
            Catch ex As Exception
                Failed(ex)
            End Try

            'IList.IsReadOnly
            Console.Write("   IList.IsReadOnly: ")
            Try
                PassFail(Not l.IsReadOnly)
            Catch ex As Exception
                Failed(ex)
            End Try

            'IList.Item
            Console.Write("   IList.Item: ")
            Try
                Dim o As Object
                o = l.Item(0)
                Failed()
            Catch ex As ArgumentOutOfRangeException
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub


    Sub DumpIList(ByVal list As IList)
        Dim i As Integer
        For i = 0 To list.Count - 1
            Console.WriteLine("      Item({0}) {1}", i, CStr(list.Item(i)))
        Next i
    End Sub


    Sub Failed(ByVal ex As Exception)
        Console.WriteLine("FAILED !!!")
        If ex Is Nothing Then
            Console.WriteLine("NULL System.Exception")
        Else
            Console.WriteLine(ex.ToString())
        End If
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



    Sub NewApiTest(ByVal TestName As String)
        TestCounter = 0
        Console.WriteLine("----------------------------------")
        Console.WriteLine(TestName)
    End Sub



    Sub NewSubTest(ByVal SubTestName As String)
        Console.WriteLine(SubTestName)
        TestCounter += 1
    End Sub
End Module

Module Bug254032
    Sub Test()
        Console.WriteLine("*** Bug 254032")
        Try
            Dim c As New Collection()
            Dim o As Object

            c.Add("7"c)

            'Valid index
            Console.Write("   1) ")
            Try
                o = c.Item(1)
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

            'index out of range
            Console.Write("   2) ")
            Try
                Console.WriteLine(c.Item(2))
                Failed()
            Catch ex As IndexOutOfRangeException When Err.Number = 9
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

            'Char index
            Console.Write("   3) ")
            Try
                Console.WriteLine(c.Item("1"c))
            Catch ex As ArgumentException When Err.Number = 5
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

            'String index
            Console.Write("   4) ")
            Try
                Console.WriteLine(c.Item("1"))
            Catch ex As ArgumentException When Err.Number = 5
                Passed()
            Catch ex As Exception
                Failed(ex)
            End Try

        Catch e As Exception
            Console.WriteLine(e.ToString)
        End Try

    End Sub
End Module

