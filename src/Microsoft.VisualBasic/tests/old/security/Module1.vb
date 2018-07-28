Imports System
Imports Microsoft.VisualBasic

Module Module1

    Dim m_lBugID As Long
    Dim m_TestDir As String

    Public ReadOnly Property TestDir As String
        Get
            Return m_TestDir
        End Get
    End Property

    Sub New()

        m_TestDir = Environ("TESTDIR") & "\temp"
        Try
            MkDir(m_TestDir)
        Catch
            'probably already created
        End Try

    End Sub

    Sub Main()
        Console.WriteLine("Begin Tests")

        FileSystem.DoTests()

        Console.WriteLine("End Tests")
    End Sub


    Sub ErrorCheck(ByVal iErrExpected As Long)
        If iErrExpected = Err.Number Then
            Console.WriteLine("passed")
        Else
            Console.WriteLine("FAILED Error: " & CStr(Err.Number) & "  Expected " & CStr(iErrExpected))
        End If
        m_lBugID = 0
    End Sub



    Sub SetBugID(ByVal lID As Long, Optional ByVal lExceptedErr As Long = 0)
        If m_lBugID <> 0 Then
            Console.WriteLine("fail (unexpected exception): " & lID)
        End If
        m_lBugID = lID
        Console.Write("Bug" &  CStr(m_lBugID) & ": " )
    End Sub


    Sub Failed(ByVal ex As Exception)
        If ex Is Nothing Then
            Console.WriteLine("    NULL System.Exception")
        Else
            Console.WriteLine("    " & ex.GetType().FullName)
        End If
        Console.WriteLine("    " & ex.Message)
        Console.WriteLine("    " & ex.StackTrace)
        Console.WriteLine("    FAILED !!!")
    End Sub


    Sub Failed()
        Console.WriteLine("    FAILED !!!")
    End Sub


    Sub Passed()
        Console.WriteLine("    passed")
    End Sub


    Sub PassFail(ByVal bPassed As Boolean)
        If bPassed Then
            Console.WriteLine("    passed")
        Else
            Console.WriteLine("    FAILED !!!")
        End If
        m_lBugId = 0
    End Sub

End Module
