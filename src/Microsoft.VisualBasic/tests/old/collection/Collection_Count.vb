Option Strict Off

Imports System
Imports Microsoft.VisualBasic

Module Collection_Count

'---------------------------------------------------------------------------
' Test       : Collection_Count                                          [1]
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
' Keyword    : Count
'
'---------------------------------------------------------------------------
' Purpose    : Verify that Count method is correctly returning number of items 
'              in collection
'
' Scenarios  :   1. Test for Count when no item has been added to collection
'                2. Test for Count when items have been added to collection
'
' Abstract:     Returns the number of items in a collection
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
        Console.WriteLine( "----Count---" )
        Collection_Count001
    End Sub


    Sub Collection_Count001()

        Dim MyCollection as New Collection
        Dim iData        as Integer
        Dim dData        as Double
        Dim sData        as String

	    On Error Resume Next

    '   Test for Count when no item has been added to collection
        Console.WriteLine( "Count - No Item" ) 
        Console.WriteLine( MyCollection.Count() )


    '   Test for Count when items have been added to collection
        Console.WriteLine( "Count - With Items" ) 

        MyCollection.Add( 23, "Marks.1" )
        Console.WriteLine( MyCollection.Count() )

        MyCollection.Add( "Redmond", "City.1")
        Console.WriteLine( MyCollection.Count() )

        MyCollection.Add( 72.5, "Fraction.1" )
        Console.WriteLine(MyCollection.Count())

        On Error GoTo 0

    End Sub

End Module
