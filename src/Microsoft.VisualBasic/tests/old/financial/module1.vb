Option Compare Binary

Imports System
Imports Microsoft.VisualBasic

Module Test

    Sub Main
    
        'Don't run NamedParamTest, just ensure compile
        'NamedParamTest

        DDB_Test.Main
        FV_Test.Main
        IPmt_Test.Main
        IRR_Test.Main
        NPer_Test.Main
        NPV_Test.Main
        PPmt_Test.Main
        PV_Test.Main
        Rate_Test.Main
        Pmt_Test.Main
        SLN_Test.Main
        SYD_Test.Main
        MIRR_Test.Main
    End Sub


    Sub NamedParamTest()
        ' NOTE:  ***** DO NOT MODIFY THIS FUNCTION *****
        '       (unless you really know what you're doing)
        '
        ' This function is intended to just to ensure compilation
        ' If someone changes a name in the runtime, this should break
        ' we will rely on the test team to make sure named params
        ' return results as expected
        '
        Dim v As object
        Dim ai() As Double

        v = Microsoft.VisualBasic.Financial.DDB(Cost := 1, Salvage := 1, Life := 1, Period := 1, Factor := 1)
        v = Microsoft.VisualBasic.Financial.FV(Rate := 1, NPer := 1, Pmt := 1, PV := 1, Due := 1)
        v = Microsoft.VisualBasic.Financial.IPmt(Rate := 1, Per := 1, NPer := 1, PV := 1, FV := 1, Due := 1)
        v = Microsoft.VisualBasic.Financial.IRR(ValueArray := ai, Guess := 1)
        v = Microsoft.VisualBasic.Financial.MIRR(ValueArray := ai, FinanceRate := 1, ReinvestRate := 1)
        v = Microsoft.VisualBasic.Financial.NPer(Rate := 1, Pmt := 1, PV := 1, FV := 1, Due := 1)
        v = Microsoft.VisualBasic.Financial.NPV(Rate := 1, ValueArray := ai)
        v = Microsoft.VisualBasic.Financial.Pmt(Rate := 1, NPer := 1, PV := 1, FV := 1, Due := 1)
        v = Microsoft.VisualBasic.Financial.PPmt(Rate := 1, Per := 1, NPer := 1, PV := 1, FV := 1, Due := 1)
        v = Microsoft.VisualBasic.Financial.PV(Rate := 1, NPer := 1, Pmt := 1, FV := 1, Due := 1)
        v = Microsoft.VisualBasic.Financial.Rate(NPer := 1, Pmt := 1, PV := 1, FV := 1, Due := 1, Guess := 1)
        v = Microsoft.VisualBasic.Financial.SLN(Cost := 1, Salvage := 1, Life := 1)
        v = Microsoft.VisualBasic.Financial.SYD(Cost := 1, Salvage := 1, Life := 1, Period := 1)
    End Sub

End Module
