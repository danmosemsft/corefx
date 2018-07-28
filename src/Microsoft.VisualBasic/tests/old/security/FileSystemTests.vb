Imports System.Security
Imports System.Security.Permissions

Module FileSystem

    Sub DoTests()
        Console.WriteLine("FileSystem Security Tests")

        TestChDir()
        TestChDrive()
        TestCurDir()
        'TestDirFunction() 'UNDONE: bugs to fix before enabling
        TestMkDir()
        TestRmDir()

    End Sub


    Sub TestChDir()
        Console.WriteLine("*** Testing ChDir ***")

        Console.WriteLine("  authorized")
        Try
            ChDir(TestDir)
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Dim sec As SecurityPermission
        sec = New SecurityPermission(SecurityPermissionFlag.UnmanagedCode)
        sec.Deny()

        Console.WriteLine("  unauthorized")
        Try
            ChDir(TestDir)
        Catch ex As SecurityException
            Passed()
        Catch ex As Exception
            Failed(ex)
        Finally
            CodeAccessPermission.RevertDeny()
        End Try

    End Sub


    Sub TestChDrive()
        Console.WriteLine("*** Testing ChDrive ***")

        Console.WriteLine("  authorized")
        Try
            ChDrive(TestDir.Chars(0))
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Dim sec As SecurityPermission
        sec = New SecurityPermission(SecurityPermissionFlag.UnmanagedCode)
        sec.Deny()

        Console.WriteLine("  unauthorized")
        Try
            ChDrive(TestDir.Chars(0))
        Catch ex As SecurityException
            Passed()
        Catch ex As Exception
            Failed(ex)
        Finally
            CodeAccessPermission.RevertDeny()
        End Try

    End Sub

    Sub TestCurDir()
        Console.WriteLine("*** Testing CurDir ***")
        
        Dim s As String

        Console.WriteLine("  authorized")
        Try
            'Test only that no exception is thrown
            s = CurDir()
            s = CurDir(TestDir.Chars(0))
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Dim sec As New FileIOPermission( FileIOPermissionAccess.PathDiscovery, TestDir & "\.")
        sec.Deny()

        Console.WriteLine("  unauthorized")
        Try
            Console.WriteLine("   CurDir()")
            s = CurDir()
            Failed()
        Catch ex As SecurityException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Try
            Console.WriteLine("   CurDir(Char)")
            s = CurDir(TestDir.Chars(0))
            Failed()
        Catch ex As SecurityException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        CodeAccessPermission.RevertDeny()

    End Sub


    Sub TestDirFunction()
        Console.WriteLine("*** Testing Dir ***")
        
        Dim s As String
        Dim DenyPath As String

        DenyPath = TestDir & "\."

        Console.WriteLine("  authorized")
        Try
            'Test only that no exception is thrown
            s = Dir("*.*")
            s = Dir()
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Dim sec As New FileIOPermission( FileIOPermissionAccess.PathDiscovery, DenyPath)
        sec.Deny()

        Console.WriteLine("  unauthorized")
        Try
            Console.WriteLine("   Dir(" & ChrW(34) & "*.*" & ChrW(34) & ")")
            s = Dir(TestDir & "\*.*")
            Failed()
        Catch ex As SecurityException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
        CodeAccessPermission.RevertDeny()

        'Setup for Dir() call
        s = Dir(TestDir & "\*.*")
        Try
            Console.WriteLine("   Dir()")
            sec = New FileIOPermission(FileIOPermissionAccess.PathDiscovery, DenyPath)
            sec.Deny()
            s = Dir()
            Failed()
        Catch ex As SecurityException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
        CodeAccessPermission.RevertDeny()

        'Setup for Dir() call
        s = Dir(TestDir & "\*.*")
        Try
            Console.WriteLine("   Dir() scenario 2")
            sec = New FileIOPermission(FileIOPermissionAccess.PathDiscovery, DenyPath)
            sec.Deny()
            s = Dir()
            Failed()
        Catch ex As SecurityException
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try
        CodeAccessPermission.RevertDeny()

    End Sub


    Sub TestMkDir()
        Console.WriteLine("*** Testing MkDir ***")

        Dim NewDir As String = TestDir & "\MkDirTest"

        Console.WriteLine("  authorized")
        Try
            Try 
                'Cleanup old dir
                RmDir(NewDir)
            Catch
                'ignore
            End Try
            MkDir(NewDir)
            RmDir(NewDir)
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        Dim sec As FileIOPermission
        sec = New FileIOPermission(FileIOPermissionAccess.Write, NewDir & "\.")
        sec.Deny()

        Console.WriteLine("  unauthorized")
        Try
            MkDir(NewDir)
            Failed()
        Catch ex As SecurityException
            Passed()
        Catch ex As Exception
            Failed(ex)
        Finally
            CodeAccessPermission.RevertDeny()
        End Try

    End Sub

    Sub TestRmDir()
        Console.WriteLine("*** Testing RmDir ***")

        Dim NewDir As String = TestDir & "\RmDirTest"

        Console.WriteLine("  authorized")
        Try
            Try
                'Make something to remove
                MkDir(NewDir)
            Catch
                'ignore
            End Try
            RmDir(NewDir)
            Passed()
        Catch ex As Exception
            Failed(ex)
        End Try

        'Setup for next test
        Try
            MkDir(NewDir)
        Catch
            'ignore (assuming already existed)
        End Try

        Dim sec As FileIOPermission
        sec = New FileIOPermission(FileIOPermissionAccess.Write, NewDir & "\.")
        sec.Deny()

        Console.WriteLine("  unauthorized")
        Try
            RmDir(NewDir)
            Failed()
        Catch ex As SecurityException
            Passed()
        Catch ex As Exception
            Failed(ex)
        Finally
            CodeAccessPermission.RevertDeny()
        End Try

    End Sub

End Module
