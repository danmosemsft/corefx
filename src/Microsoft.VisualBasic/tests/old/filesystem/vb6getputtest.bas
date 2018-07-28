Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
    VB6CompatGetTest
End Sub

Sub VB6CompatGetTest()
    'Data types to test
    Dim i As Integer
    Dim l As Long
    
    Dim byt As Byte
    Dim dt As Date
    Dim s As String
    Dim b As Boolean
    Dim dbl As Double
    Dim sng As Single
    Dim v As Variant
    Dim dec As Variant
    Dim FileNumber As Integer
    
    FileNumber = FreeFile
    
    Debug.Print ("*** VB6 Get compatibility test")
    
    Open "f:\GetPutTest-vb6.out" For Random Access Write As #1
    
    Debug.Print ("Short: ")
    i = 1234
    Put #FileNumber, , i
    
    Debug.Print ("Integer: ")
    l = 56789
    Put #FileNumber, , l
    
    Debug.Print ("Date: ")
    dt = #6/28/1999#
    Put #FileNumber, , dt
    
    Debug.Print ("Byte: ")
    byt = 129
    Put #FileNumber, , byt
    
    Debug.Print ("Double: ")
    dbl = 1234567.7654321
    Put #FileNumber, , dbl
    
    Debug.Print ("Single: ")
    sng = CSng(1234.567)
    Put #FileNumber, , sng
    
    Debug.Print ("Boolean: ")
    b = True
    Put #FileNumber, , b
    
    Debug.Print ("String: ")
    s = "A 16 char string"
    Put #FileNumber, , s
    
    'VB6 cannot read/write a decimal except inside a variant
    
    Debug.Print ("VARIANT(Short): ")
    v = CInt(1234)
    Put #FileNumber, , v
    
    Debug.Print ("VARIANT(Integer): ")
    v = CLng(56789)
    Put #FileNumber, , v
    
    Debug.Print ("VARIANT(Date): ")
    v = #6/28/1999#
    Put #FileNumber, , v
    
    Debug.Print ("VARIANT(Byte): ")
    v = CByte(129)
    Put #FileNumber, , v
    
    Debug.Print ("VARIANT(Double): ")
    v = 1234567.7654321
    Put #FileNumber, , v
    
    Debug.Print ("VARIANT(Single): ")
    v = CSng(1234.567)
    Put #FileNumber, , v
    
    Debug.Print ("VARIANT(Boolean): ")
    v = True
    Put #FileNumber, , v
    
    Debug.Print ("VARIANT(String): ")
    v = "A 16 char string"
    Put #FileNumber, , v
    
    Debug.Print ("VARIANT(Decimal) positive: ")
    v = CDec(12345.67)
    Put #FileNumber, , v
    
    Debug.Print ("VARIANT(Decimal) negative: ")
    v = CDec(-12345.67)
    Put #FileNumber, , v
    
    Debug.Print ("VARIANT(DBNull): ")
    v = Null
    Put #FileNumber, , v
    
    Close #FileNumber
    Debug.Print ("*** End VB6 Get compatibility test")
End Sub

Sub PassFail(b As Boolean)
    If b Then
        Debug.Print "passed"
    Else
        Debug.Print "FAILED!!!"
    End If
End Sub

