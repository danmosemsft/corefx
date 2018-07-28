Option Strict Off

Imports System
Imports System.Text
Imports Microsoft.VisualBasic
Imports System.Security
Imports System.Security.Permissions

Module Test

    Dim s As String
    Dim obj As Object
	Dim i As Integer



    Sub Main()
        'Don't run NamedParamTest, just ensure compile
        'NamedParamTest

        Console.WriteLine( "Begin Test" )

        RegressionTests

        CommandTests
        BeepTests
        EnvironTests
        CallByNameTests
        CreateObjectTests
        GetObjectTests
        ChooseTests
        IIFTests
        PartitionTests
        SwitchTests
        RegistryTests
        'Disable ShellTests as shell() require VisualStudio and does not work in IA64
        'ShellTests
        Console.WriteLine( "End Test" )
    End Sub

    Sub RegressionTests
        Console.WriteLine("REGRESSION TESTS")
        Bug245610.Test
    End Sub

    Sub Failed(ByVal ex As Exception)
        If ex Is Nothing Then
            Console.WriteLine("NULL System.Exception")
        Else
            Console.WriteLine(ex.GetType().FullName)
        End If
        Console.WriteLine(ex.StackTrace)
        Console.WriteLine("FAILED !!!")
    End Sub



    Sub Failed()
        Console.WriteLine("FAILED !!!")
    End Sub



    Sub Passed()
        Console.WriteLine("passed")
    End Sub



    Function PassFail(ByVal bPassed As Boolean) As String
        If bPassed Then
            Console.WriteLine( "passed" )
        Else
            Console.WriteLine( "FAILED !!!" )
        End If
    End Function



'    Sub ShellTests
'        Dim ProcessId As Integer
'        Dim Path As String
'
'        Console.WriteLine( "*** Shell tests" )
'        ProcessId = Shell("vbcrun-ShellTest.exe", AppWinStyle.MinimizedNoFocus, True)
'    End Sub



    Sub CommandTests()
        Console.WriteLine( "*** Command tests" )

        Try
            Console.WriteLine( Command )
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub EnvironTests()
        Console.WriteLine( "*** Environ tests" )

        Try
            'UNDONE
            s = "PATH"
            Console.Write( "1) " )
            PassFail(Environ(s) <> "" )

            Console.Write( "2) " )
            PassFail(Environ(1) <> "" )
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write( "3) " )
            s = Environ(-1)
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write( "4) " )
            Environ(0)
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write( "5) " ) 
            Environ(256)
            Failed
        Catch aex As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub BeepTests	    
        Console.WriteLine( "*** Beep tests" )
        Try
            Beep
            Console.WriteLine( "Did you hear that?" )
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub


    Private Class Class1

        Public m_i As Integer

        Public Function MethodFoo() As Integer
            Return 1
        End Function

        Public Property PropFoo() As Integer
            Get
                Return 1
            End Get
            Set
                m_i = Value        
            End Set
        End Property
    
    End Class

    Sub CallByNameTests()
        Dim c As New Class1
        Dim o As Object
        Dim i As Integer

        Console.WriteLine( "*** CallByName tests" )
        Try
            Console.Write( "    1) " )
            o = CallByName(c, "MethodFoo", CallType.Method)
            PassFail(o = 1)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write( "    2) " )
            o = CallByName(c, "PropFoo", CallType.Get)
            PassFail(o = 1)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write( "    3) " )
            CallByName(c, "PropFoo", CallType.Let, 3)
            PassFail(c.m_i = 3)
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.Write( "    4) " )
            CallByName(c, "PropFoo", CallType.Set, 4)
            PassFail(c.m_i = 4)
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub

 

    Sub CreateObjectTests()
        Dim o As Object
        Console.WriteLine( "*** CreateObject tests" )
        Try
            Console.Write( " 1: " )
            o = createobject("Scripting.FileSystemObject") 'RAID 145631
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try

        CreateObjectErrors
    End Sub
 


    Sub CreateObjectErrors()
        Dim o As Object
        On Error Resume Next

        'RAID 145735  This takes 3.5 minutes to timeout on NT4 build 1381 SP6
        'RAID 200052  This leaks memory  in OLE Automation on NT5
        'o = CreateObject("Scripting.FileSystemObject", "SDFSDF") 
        'Console.Write( " 2: " )
        'PassFail( Err.Number = 462 )
    End Sub



    Sub GetObjectTests()
        Dim o As Object
        Console.WriteLine( "*** GetObject tests" )

        Try
            Console.Write( "1: " )
            o = GetObject("", "Scripting.FileSystemObject") 'RAID 145631
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try

        GetObjectErrors
    End Sub



    Sub GetObjectErrors()
        Dim o As Object
        On Error Resume Next

        o = GetObject("", "blah.blah") 'RAID 78296
        Console.Write( "2: ")
        PassFail( Err.Number = 429 )

        o = GetObject("asdfasdf", "Scripting.FileSystemObject") 'RAID 78296
        Console.Write( "3: ")
        PassFail( Err.Number = 432 )

        o = GetObject("BogusClass.Blah") 'RAID 109939
        Console.Write( "4: ")
        PassFail( Err.Number = 429 )
    End Sub



    Sub ChooseTests()
        Console.WriteLine( "*** Choose tests" )
        Try
	        Dim ii As Integer
	        Dim X1 As Object
	        Dim X2 As Object
	        Dim X3 As Object
	        Dim X4 As Object
	        Dim X5 As Object
	        Dim X6 As Object

            X1 = "Choice1"
            X2 = "Choice2"
            X3 = "Choice3"
            X4 = "Choice4"
            X5 = "Choice5"	
            X6 = "Choice6"

	        '0 and Last should return Null
            For ii = 0 To 7
                Console.Write( CStr(ii) & "): " )
                Console.WriteLine( CStr(Choose(ii, X1, X2, X3, X4, X5, X6)) )
            Next ii

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub IIFTests()
        Console.WriteLine( "*** IIF tests" )

        Try
            'TODO: do an object test
            Console.Write( "1) " )
            PassFail( CBool(IIF(True, True, False )) )

            Console.Write( "2) ")
            PassFail( CBool(IIF(False, False, True )) )

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub PartitionTests()
        Console.WriteLine( "*** Partition tests" )

        Try
            Console.Write("1) ")
            Console.WriteLine( Partition(0, 1, 100, 10) )

            Console.Write("2) ")
            Console.WriteLine( Partition(1, 1, 100, 10) )

            Console.Write("3) ")
            Console.WriteLine( Partition(15, 1, 100, 10) )

            Console.Write("4) ")
            Console.WriteLine( Partition(25, 1, 100, 10) )

            Console.Write("5) ")
            Console.WriteLine( Partition(35, 1, 100, 10) )

            Console.Write("6) ")
            Console.WriteLine( Partition(45, 1, 100, 10) )

            Console.Write("7) ")
            Console.WriteLine( Partition(5001, 1, 10000, 100) )

        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub SwitchTests()
        Console.WriteLine( "*** Switch tests" )

        Try
            For i = 1 to 4
                Console.WriteLine( Switch( i = 1, "red", i = 2, "blue", i = 3, "green" ) )
            Next i

            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub RegistryTests()
        Console.WriteLine( "*** Registry Tests" )

        Try
            Console.WriteLine( "Calling SaveSetting for 9 sections" )
	        For i = 1 to 9
		        SaveSetting("MyApp1", "MySection" & i, "MyKey" & i, "MyValue" & i)
	        Next i

            Console.WriteLine( "Calling GetSetting on 9 sections" )
	        For i=1 to 9
		        Console.WriteLine( GetSetting("MyApp1", "MySection" & i, "MyKey" & i ) )
	        Next i

            Console.WriteLine( "Calling DeleteSetting on MyKey1" )
            DeleteSetting("MyApp1", "MySection1", "MyKey1")
	        Console.WriteLine( GetSetting("MyApp1", "MySection1", "MyKey1", "MyKey1 was deleted" ) )

            Console.WriteLine( "Calling DeleteSetting on 3 sections" )
	        For i=1 to 3
		        DeleteSetting("MyApp1", "MySection" & i)
		        Console.WriteLine( GetSetting("MyApp1", "MySection" & i, "MyKey", "MySection" & i & " was deleted" ) )
	        Next i

            Console.WriteLine( "Calling DeleteSetting on 1 application" )
            DeleteSetting("MyApp1")
	        For i=4 to 9
		        Console.WriteLine( GetSetting("MyApp1", "MySection" & i, "MyKey", "MySection" & i & " was deleted" ) )
	        Next i

            Console.WriteLine( "Testing GetAllSettings" )
            SaveSetting( "Bob", "mom", "Dad", "Test" )
            SaveSetting( "Bob", "mom", "Dad2", "Test2" )
            obj = GetAllSettings("bob", "mom")
            Console.WriteLine( CStr(obj(0,0)) & " " & CStr(obj(0,1)) )
            Console.WriteLine( CStr(obj(1,0)) & " " & CStr(obj(1,1)) )

            obj = GetAllSettings("aaaaSDFSDF", "SDF")
            Console.Write( "GetAllSettings of nonexistant key returns Nothing: " ) 
            PassFail(IsNothing(obj))

            Console.Write( "Testing DeleteSettings: " ) 
            DeleteSetting("Bob")
            obj = GetAllSettings("bob", "mom")
            PassFail(IsNothing(obj))
        
        Catch ex As Exception
            Failed(ex)
        End Try

        RegistrySubTest1
        RegistrySubTest2
    End Sub



    Sub RegistrySubTest1()
        Console.Write( "DeleteSetting of nonexistant app: " )
        Try
            DeleteSetting("asd","")
            Failed
        Catch ae As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub RegistrySubTest2()
        Console.Write( "DeleteSetting of nonexistant section: " )

        Try
            DeleteSetting("asd","asd", "")
            Failed
        Catch ae As ArgumentException
            Passed
        Catch ex As Exception
            Failed(ex)
        End Try
    End Sub



    Sub NamedParamTest()
#if 0
        ' NOTE:  ***** DO NOT MODIFY THIS FUNCTION *****
        '       (unless you really know what you're doing)
        '
        ' This function is intended to just to ensure compilation
        ' If someone changes a name in the runtime, this should break
        ' we will rely on the test team to make sure named params
        ' return results as expected
        '
        Dim av(2) As Object
        Dim obj As Object

        VBA.Beep() 'No params
        obj = VBA.CallByName(Object := obj, ProcName := "proc", CallType := 0, Args := av)
        'Choose uses a ParamArray, which named parameters are invalid
        'obj = VBA.Choose(Index := 1, Choice := obj)
        obj = VBA.Command() 'No params
        obj = VBA.Command$() 'No params
        obj = VBA.CreateObject(Class := "class", ServerName := "server")
        VBA.DeleteSetting(AppName := "app", Section := "section", Key := "key")
        obj = VBA.Environ(Expression := obj)
        obj = VBA.Environ$(Expression := obj)
        obj = VBA.GetAllSettings(AppName := "app", Section := "section")
        obj = VBA.GetObject(PathName := obj, Class := obj)
        obj = VBA.GetSetting(AppName := "app", Section := "section", Key := "key", Default := obj)
        obj = VBA.IIF(Expression := obj, TruePart := True, FalsePart := False)
        obj = VBA.Partition(Number := obj, Start := obj, Stop := obj, Interval := obj)
        VBA.SaveSetting(AppName := "app", Section := "section", Key := "key", Setting := obj)
        'Switch uses a ParamArray, which named parameters are invalid
        'obj = VBA.Switch(VarExpr := obj)
#end if
    End Sub


End Module

Module Bug245610
    Sub Test()
        Console.Write("Bug 245610: ")
        Dim s As String
        Try
            s = Command()
            Call (New EnvironmentPermission(EnvironmentPermissionAccess.Read, "Path")).Deny()
            Try
                s = Command()
                Failed()
            Catch ex As SecurityException
                Passed()
            End Try
        Catch ex As Exception
            Failed(ex)
        End Try

    End Sub
End Module


