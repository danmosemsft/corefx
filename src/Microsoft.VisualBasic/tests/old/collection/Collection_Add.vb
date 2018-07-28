Option Strict Off

Imports System
Imports Microsoft.VisualBasic

Module Collection_Add


    Sub Main()
        Console.WriteLine( "----Add---" )
        Collection_Add001
        Collection_Add002
    End Sub


    Sub Collection_Add001()
        Dim MyCollection as New Collection
        Dim iData        as Integer
        Dim dData        as Double
        Dim sData        as String
        Dim iErrorStmt   as Integer

	    On Error GoTo ErrorHandler

    '   Test for correctly adding item to collection 
        Console.WriteLine( "Add Cor" )
        MyCollection.Add( 23, "Marks.1" )
        MyCollection.Add( "Redmond", "City.1" )
        MyCollection.Add( 72.5, "Fraction.1" )

        sData = MyCollection.Item("City.1")
        Console.WriteLine( sData )
    
        iData = MyCollection.Item("Marks.1")
        Console.WriteLine( iData )

        dData = MyCollection.Item("Fraction.1")
        Console.WriteLine( dData )

    '   Test for correctly adding item if key is missing 
        Console.WriteLine( "Add - Missing Key" )
        Call MyCollection.Add("No Key Data")
        sData = MyCollection.Item(4)
        Console.WriteLine( sData )

    '   Test for the duplicate key
        Console.WriteLine( "Add - Duplicate Key" )
        iErrorStmt = 1
        MyCollection.Add( "Seattle", "City.1" )
        sData = MyCollection.Item("City.1")
        Console.WriteLine( sData )

        On Error GoTo 0
        Exit Sub

    ErrorHandler :
        If iErrorStmt = 1 Then
                Console.WriteLine( "Duplicate Key" )
                Resume Next
        End If

        Console.WriteLine( "Unknown Error" )
        Resume Next
    End Sub



    Sub Collection_Add002()
        Console.WriteLine( "Test that the ForEach iterators get updated when adding items" )
        Dim m As New Collection()
        Dim item1 As Object
        Dim item2 As Object

        m.Add("a")
        m.Add("b")
        m.Add("c")
        m.Add("d")
        m.Add("e")
        m.Add("f")

        For Each item1 in m        
            Console.WriteLine( "1) " & item1 )
            For Each item2 in m
                Console.WriteLine( "2) " & item2 )
            Next item2
            m.Add("aa",,1)
            If m.Count>11 Then
                Exit For
            End If
        Next item1
    End Sub



End Module
