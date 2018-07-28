Option Strict Off

Imports System
Imports Microsoft.VisualBasic

Module Collection_Item

'---------------------------------------------------------------------------
' Test       : Collection_Item                                          [1]
'
' Component  : Language
'
' Major Area : EB
'
' Sub Area   : Language
'
' Test Area  : Collection
'
' Keyword    : Item
'
'---------------------------------------------------------------------------
' Purpose    : Verify that Item method is correctly returning item for correct
'              index or key and fails for their incorrect value
'
' Scenarios  :   1. Test for an existing index and key
'                2. Test for nonexistent index and key
'
' Abstract:     Returns the indexed member of a collection
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
        Console.WriteLine( "----Item---" )
        Collection_Item001
    End Sub


    Sub Collection_Item001()

        Dim MyCollection as New Collection
        Dim iData        as Integer
        Dim dData        as Double
        Dim sData        as String
        Dim iErrorStmt   as Integer

	    On Error GoTo ErrorHandler

    '   Test for an existing index and key
        Console.WriteLine( "Item - Existent Index" )
        MyCollection.Add( 23, "Marks.1" )
        MyCollection.Add( "Redmond", "City.1" )
        MyCollection.Add( 72.5, "Fraction.1" )

        sData = MyCollection.Item("City.1")
        Console.WriteLine( sData )
        sData = MyCollection.Item(2)
        Console.WriteLine( sData )
    
        iData = MyCollection.Item("Marks.1")
        Console.WriteLine( iData )
        iData = MyCollection.Item(1)
        Console.WriteLine( iData )

        dData = MyCollection.Item("Fraction.1")
        Console.WriteLine( dData )
        dData = MyCollection.Item(3)
        Console.WriteLine( dData )

    '   Test for nonexistent index and key
        Console.WriteLine( "Item - Non Existent Index" )
        iErrorStmt = 1
        sData = MyCollection.Item(9)

        iErrorStmt = 2
        sData = MyCollection.Item("Missing Key")

        On Error GoTo 0
        Exit Sub

    ErrorHandler :
        If iErrorStmt = 1 Then
                Console.WriteLine( "Missing Index" )
                Resume Next
        ElseIf iErrorStmt = 2 Then
                Console.WriteLine( "Missing Key" )
                Resume Next
        End If

        Console.WriteLine( "Unknown Error" )
        Resume Next

    End Sub

End Module

