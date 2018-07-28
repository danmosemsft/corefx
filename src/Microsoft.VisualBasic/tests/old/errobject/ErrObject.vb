option strict off

Imports Microsoft.VisualBasic
Imports System

Public Module Test


    Sub Main()

        TestFirstThing
        TestInitialDescription
        TestDescriptionSetting
        TestOverwrite
        TestInvalidSettings
        TestFieldsAfterRaise
        TestNumberSettings
        TestRaiseSettings
        TestSettingProperties
        PrintAllErrVals

        TestNoExit
        PrintAllErrVals
 
        TestNoExitNoError
        PrintAllErrVals

        TestOnErrorResumeNext
        PrintAllErrVals

        TestGoto 
        PrintAllErrVals
   
        TestExitSub
        PrintAllErrVals
    
        TestExitSubNoError
        PrintAllErrVals

        TestExitFunction
        PrintAllErrVals

        TestResumeNext
        TestResumeNext2
    
        On Error Resume Next
        TestOnErrorGoto0
        PrintAllErrVals

        TestMultipleCatches
      
        TestTryCatch

        TestLineNumber   

    End Sub


    Public x As Object

    Sub PrintAllErrVals
        Console.WriteLine("Num:" & Hex(Err.Number) & _
                  " Descrip:" & Err.Description & _
                  " Context:" & Err.HelpContext & _
                  " File:" & Err.HelpFile & _
                  " Source:" & Err.Source)
    End Sub


    Sub TestSettingProperties()
        Err.Clear

        Err.Number = 10
        PrintAllErrVals
        Err.Clear
    
        Err.Number = 5
        Err.Description = "hi"
        Err.HelpContext = 3
        Err.HelpFile = "test.txt"
        Err.Source = "blah"
        PrintAllErrVals
    
        Err.Number = 6
        PrintAllErrVals
    
        Err.Clear
        PrintAllErrVals
    
        On Error Resume Next
        Err.Raise(5, "blah2", "hi2", "test2.txt", 3)
        PrintAllErrVals
        On Error Resume Next

        PrintAllErrVals
        Err.Number = 5
        Err.Description = "this is a description"
    
    End Sub                    'this will not clear the err object

    Sub TestNoExit()

        On Error GoTo lbl_test
        x.foo
    
    lbl_test:
        PrintAllErrVals
    
    End Sub                    'this clears the err object, but only if an exception occurs


    Sub TestNoExitNoError()

        On Error GoTo lbl_test
        Err.Number = 5
        Err.Description = "this should still exist"
    
    lbl_test:             'this clears the err object, but only if an exception occurs
        PrintAllErrVals
    
    End Sub


    Sub TestGoto()

        On Error GoTo lbl_test
        x.foo

    lbl_exit:
        PrintAllErrVals
        Err.Number = 5
        Err.Description = "this is a description"
        GoTo lbl_end
    
    lbl_test:             'this clears the err object, but only if an exception occurs
        PrintAllErrVals
        GoTo lbl_exit
    
    lbl_end:
        PrintAllErrVals
    End Sub


    Sub TestOnErrorResumeNext()

        On Error Resume Next
        x.foo

    End Sub             'the err object should not be cleared upon ending the sub

    Sub TestExitSub()
    
        On Error GoTo lbl_test
        x.foo

    lbl_test:           'this clears the err object, but only if an exception occurs
        Exit Sub

    End Sub

    Sub TestExitSubNoError()
    
        On Error GoTo lbl_test
        Err.Number = 5
        Err.Description = "this should still exist"

    lbl_test:           'this clears the err object, but only if an exception occurs
        Exit Sub

    End Sub


    Function TestExitFunction() As Integer

        On Error GoTo lbl_test
        x.foo

    lbl_test:           'this clears the err object, but only if an exception occurs
        Exit Function

    End Function


    Sub TestResumeNext()

        On Error GoTo lbl_test
        x.foo
    
        PrintAllErrVals
        Exit Sub
    
    lbl_test:
        Resume Next     'this clears the err object
    
    End Sub

    Sub ThrowOnZero(n As Integer)
        If n = 0 Then
            Err.Raise(5, "test.txt", "this is the descrip", "help.txt", 3)
        End If
    End Sub

    Sub TestResumeNext2()

        Dim n As Integer
        n = 0
    
        On Error GoTo lbl_test
        ThrowOnZero(n)
        PrintAllErrVals
    
        ThrowOnZero(0)
        PrintAllErrVals
        Exit Sub
    
    lbl_test:
        n = 1
        PrintAllErrVals
        On Error GoTo lbl_test2    'this clears the err object
        PrintAllErrVals
        Err.Number = 5
        Err.Description = "blah"
        PrintAllErrVals
        Resume                     'this clears the err object

    lbl_test2:
        PrintAllErrVals
        Resume Next                'this clears the err object
    
    End Sub


    Sub TestOnErrorGoto0()

        On Error GoTo 0
        x.foo

    End Sub                        'this does not clear the err object


    Sub TestFieldsAfterRaise()

        On Error Resume Next
        Err.Raise(5)
        PrintAllErrVals
        Err.Raise(6)
        PrintAllErrVals
        Error 10              'the error statement is equivalent to Raise with all params specified
        PrintAllErrVals
        Dim s As String
        Err.Description = ""
        Err.Raise(9)
        PrintAllErrVals
    
        Err.Number = 0
        Err.Raise(9)
        PrintAllErrVals

    End Sub


    'Raise(val, source, descrip, file, context):
    '    If Number = 0 Then
    '       SetAllFieldsAccordingToNumber(val)
    '    Else
    '        Number = Val
    '    End If
    '
    '    If source is supplied, set the souce
    '    If descrip is supplied, set the descrip
    '    If file is supplied, set the file
    '    If context is supplied, set the context
    '
    Sub TestRaiseSettings()
    
        On Error Resume Next
        Err.Raise(6)
        PrintAllErrVals
        Err.Clear
    
        Err.Raise(7)
        PrintAllErrVals
        Err.Clear
    
        Err.Raise(8, "bob", "smith", "blah", 42)
        PrintAllErrVals
        Err.Clear
    
        Err.Raise(9, "bob2", , "blah2", 43)
        PrintAllErrVals
        Err.Clear
    
    End Sub


    Sub TestNumberSettings()
        On Error Resume Next
        Err.Number = 65536
        PrintAllErrVals
        On Error GoTo 0
    
        Err.Number = &H800A00FF
        PrintAllErrVals
    
        Err.Number = &H800B0001
        PrintAllErrVals
    
        Err.Number = &H80100001
        PrintAllErrVals
    
        On Error Resume Next
        Err.Raise(&H80010001)  'BUG BUG:  this doesn't set the automation error description
        PrintAllErrVals
        On Error GoTo 0
    
        On Error Resume Next
        Err.Raise(&H80030002)
        PrintAllErrVals
        On Error GoTo 0
    
        On Error Resume Next
        Err.Raise(&H80004001)
        PrintAllErrVals
        On Error GoTo 0
    
    
    #If 0 Then
        Dim l As Long
        For l = &H80000000 To &HFFFF
            Err.Number = l
            If l <> Err.Number And ((l And SCODE_FACILITY) <> FACILITY_CONTROL) Then
                Debug.Print "ouch: " & Hex(l) & " " & Hex(Err.Number)
            End If
            If l Mod &HFFFFFF = 0 Then
                Debug.Print Hex(l)
            End If
        Next
    #End If
    

    #If 0 Then
        Dim l As Long
        For l = 0 To &HFFFFFFF
            Err.Number = l
            If l <> Err.Number And ((l And SCODE_FACILITY) <> FACILITY_CONTROL) Then
                Debug.Print "ouch: " & Hex(l) & " " & Hex(Err.Number)
            End If
            If l Mod &HFFFFFF = 0 Then
                Debug.Print Hex(l)
            End If
        Next
    #End If


    End Sub

    Class Class1
    End Class

    Sub TestInvalidSettings()
        Dim x as New Class1
            
        On Error Resume Next
        Err.Raise(10, , "asdf", "asdf", "asdf")
        PrintAllErrVals
        Err.Raise(10, x,"asdf", "asdf", 0)
        PrintAllErrVals

        Error 0
        PrintAllErrVals
        Err.Raise(0)
        PrintAllErrVals

    End Sub



    Sub TestOverwrite()
    
        On Error Resume Next
    
        Err.Description = "hello"
        Err.Number = 5
        Err.HelpFile = "bob"
        PrintAllErrVals
    
        x.foo
        PrintAllErrVals

    End Sub


    Sub TestInitialDescription()
        On Error Resume Next
        PrintAllErrVals
        Err.Raise(16)
        PrintAllErrVals
    End Sub


    Sub TestDescriptionSetting
        On Error Resume Next
        Err.Description = "hello.  this should not disappear"
        PrintAllErrVals
        Err.Raise(3)
        PrintAllErrVals
    End Sub


    'NOTE:  This function has to be called first because it tests
    '       functionality of the Err object's constructor
    Sub TestFirstThing
        Try
            Err.Raise(9)
        Catch e As Exception
            Console.WriteLine(e.Message)
            Console.WriteLine(err.Description)
        End Try
    End Sub


    Function GetErrNum() As Integer
        GetErrNum = Err.Number
        Console.WriteLine("Num:" & GetErrNum)
    End Function

    Sub TestMultipleCatches

        Dim i As Short = 1

        Try
            Try
                Err.Raise(i)

            Catch When GetErrNum() = i - 1
                Console.WriteLine("shouldn't get here: " & Err.Number)
            Catch When GetErrNum() = i
                Console.WriteLine("SUCCESS " & Err.Number)
            End Try
            
            Catch
                Console.WriteLine("shouldn't get here: " & Err.Number)
        End Try

    End Sub

End Module


Public Module TryCatchTest

    Sub PassFail(i as integer, b as boolean)
        Console.Write(i & ": ")
        If b Then console.writeline("pass") else console.writeline("****FAIL")
    End Sub


    Sub Test2
        Try
            Throw New Exception
        Catch
            Exit Sub
        End Try
    End Sub


    Sub Test3
        Try
            Throw New Exception
        Catch
            Try
                Exit Sub
            Catch
            End Try
        End Try
    End Sub


    Sub Test4
        Try
            Throw New Exception
        Catch
            Select Case 1
            Case 1:
                Try
                    Exit Sub
                Catch
                End Try
            End Select
        End Try
    End Sub


    Sub Test5
        Try
            Throw New Exception
        Catch
            Select Case 1
            Case 1:
                Try
                    Throw New Exception
                Catch
                    Exit Sub
                End Try
            End Select
        End Try
    End Sub


    Sub Test1

        Try
            Throw New Exception
        Catch
            PassFail(1, Err.Number = 5)
        End Try
        PassFail(2, Err.Number = 0)


        Try
           Throw New Exception
        Catch
           Err.Number = 42
           PassFail(3, Err.Number = 42)
        End Try
        PassFail(4, Err.Number = 0)


        Try
            Throw New Exception
        Catch
            PassFail(5, Err.Number = 5)
        Finally
            PassFail(6, Err.Number = 0)
        End Try
        PassFail(7, Err.Number = 0)
        

        Try
            Err.Number = 42
        Catch
        End Try
        PassFail(8, Err.Number = 42)


        Try
            Err.Number = 42
        Catch
        Finally
        End Try
        PassFail(9, Err.Number = 42)

        
        Try
            Err.Number = 42
            Goto Label1
        Catch
        End Try
Label1:
        PassFail(10, Err.Number = 42)

        
        Try
            Throw New Exception
        Catch
            PassFail(11, Err.Number = 5)
            Goto Label2
        End Try
Label2:
        PassFail(12, Err.Number = 0)


        Try
            Throw New Exception
Label3:
        Catch
            PassFail(13, Err.Number = 5)
            Goto Label3
        End Try
        PassFail(14, Err.Number = 0)


        Try
            Throw New Exception
        Catch
            Try
                Throw New Exception
            Catch
            End Try
            PassFail(15, Err.Number = 0) 
        End Try
        PassFail(16, Err.Number = 0)


#if BUG_217056
        Try
            Throw New Exception
Label4:
            PassFail(17, Err.Number = 0)
        Catch
            Try
                Throw New Exception
            Catch
                Goto Label4
            End Try
        End Try
#end if


        Try
            Throw New Exception
        Catch
            Try
                Throw New Exception
Label5:
                PassFail(18, Err.Number = 0)
                Err.Number = 42
                Exit Try
            Catch
                Goto Label5
            End Try
            PassFail(19, Err.Number = 42)
            Exit Try
        End Try
        PassFail(20, Err.Number = 0)


        Try
            Throw New Exception
        Catch
            Try
                Err.Number = 42
                Goto Label6
            Catch
            End Try
        End Try
Label6:
        PassFail(21, Err.Number = 0)



        Try
            Throw New Exception
        Catch
            Try
                Throw New Exception
            Catch
                Goto Label7
            End Try
        End Try
Label7:
        PassFail(22, Err.Number = 0)


        Dim i as integer = 2
        Select Case i
            Case 1:
            Case 2:
                Try
                    Throw New Exception
                Catch
                    Exit Select
                End Try
        End Select
        PassFail(23, Err.Number = 0)

        
        Try
            Throw New Exception
        Catch
            Try
                Exit Try
            Catch
            Finally
                PassFail(24, Err.Number = 5)
            End Try
        End Try
        PassFail(25, Err.Number = 0)


        Do
            Try
                Throw New Exception
            Catch
                Exit Do
            End Try
        Loop While True
        PassFail(26, Err.Number = 0)

try
catch
end try

        Test2
        PassFail(27, Err.Number = 0)


        Test3
        PassFail(28, Err.Number = 0)


        Test4
        PassFail(29, Err.Number = 0)


        Test5
        PassFail(30, Err.Number = 0)


    End Sub


    Sub TestTryCatch
        Console.WriteLine("*** TestTryCatch ***")
        Test1
    End Sub

End Module

Public Module LineNumberTest

    Sub SubLineNumber
        42: Err.Raise(1)
    End Sub
    
    Sub CaptureSubLineNumber
        On Error Goto Blah
        9: SubLineNumber
        exit sub
    blah:
        Console.WriteLine("CaptureSubLineNumber: " & CStr(Erl = 9))
        Resume Next
    End Sub

    Sub ResumeNextLineNumber
        On Error Resume Next
        56: Err.Raise(1)
    End Sub

    Sub TryCatchLineNumber
        Try
            5: Throw New Exception
        Catch
            Console.WriteLine("TryCatchLineNumber: " & CStr(Erl = 5))
        End Try
    End Sub

    Sub TestLineNumber
        Console.WriteLine("*** TestLineNumber ***")

        CaptureSubLineNumber

        ResumeNextLineNumber
        Console.WriteLine("ResumeNextLineNumber: " & CStr(Erl = 56))

        TryCatchLineNumber
    End Sub

End Module

