Option Strict Off

Imports System
Imports Microsoft.VisualBasic

Module Collection_Remove

'---------------------------------------------------------------------------
' Test       : Collection_Remove                                          [1]
'
'
' Component  : Language
'
' Major Area : EB
'
' Sub Area   : Language
'
' Test Area  : Collection
'
' Keyword    : Remove
'
'---------------------------------------------------------------------------
' Purpose    : Verify that Remove method is correctly removing item from collection
'              object.
'
' Scenarios  :   1. Test for existing index and key
'                2. Test for non existing index and key
'
' Abstract:     Removes an item from a collection
'---------------------------------------------------------------------------
' Category      : Functional
'
' Product       : COMMON
'
' Related Files : Start_Mod.bas
'
' Notes:
'---------------------------------------------------------------------------
' Revision History:
'
'---------------------------------------------------------------------------
'

    Sub Main()
        Console.WriteLine( "----Remove---" )
        Collection_Remove001
        Collection_Remove002
    End Sub


    Sub Collection_Remove001()
        Dim MyCollection as Collection
        Dim iData        as Integer
        Dim dData        as Double
        Dim sData        as String
        Dim iErrorStmt   as Integer

	    On Error GoTo ErrorHandler
        MyCollection = New Collection

    '   Test for an existing index and key
        Console.WriteLine( "Remove - Existent Index" )
        MyCollection.Add( 23, "Marks.1" )
        MyCollection.Add( "Redmond", "City.1" )
        MyCollection.Add( 72.5, "Fraction.1" )

        iErrorStmt = 1
        Call MyCollection.Remove("Marks.1")
        iData = MyCollection.Item("Marks.1")
        Console.WriteLine( MyCollection.Count() )

        iErrorStmt = 2
        MyCollection.Remove( "City.1" )
        iData = MyCollection.Item("City.1")
        Console.WriteLine( MyCollection.Count() )

        iErrorStmt = 3
        MyCollection.Remove( "Fraction.1" )
        iData = MyCollection.Item("Fraction.1")
        Console.WriteLine( MyCollection.Count() )

    '   Test for nonexistent index and key
        Console.WriteLine( "Remove - Non Existent Index" )
        MyCollection.Add( 23, "Marks.1" )
        MyCollection.Add( "Redmond", "City.1" )
        MyCollection.Add( 72.5, "Fraction.1" )

        iErrorStmt = 4
        MyCollection.Remove( "Non Existent Key" )

        iErrorStmt = 5
        MyCollection.Remove( 10 )

        On Error GoTo 0
        Exit Sub

    ErrorHandler :
        If iErrorStmt = 1 Then
                Console.WriteLine( "Marks.1 Removed" )
                Resume Next
        ElseIf iErrorStmt = 2 Then
                Console.WriteLine( "City.1 Removed" )
                Resume Next
        ElseIf iErrorStmt = 3 Then
                Console.WriteLine( "Fraction.1 Removed" )
                Resume Next
        ElseIf iErrorStmt = 4 Then
                Console.WriteLine( "Error Removing Non Existent Key" )
                Resume Next
        ElseIf iErrorStmt = 5 Then
                Console.WriteLine( "Error Removing Non Existent Index" )
                Resume Next
        End If

        Console.WriteLine( "Unknown Error" )
        Resume Next
    End Sub



    Sub Collection_Remove002()
        Console.WriteLine( "Test that the ForEach iterators get updated when removing items" )
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
            m.Remove(1)
        Next item1
    End Sub



End Module
