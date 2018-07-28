Begin Test
REGRESSION TESTS
Bug 245610: passed
*** Command tests
CommandLineArgs
*** Beep tests
Did you hear that?
passed
*** Environ tests
1) passed
2) passed
3) passed
4) passed
5) passed
*** CallByName tests
    1) passed
    2) passed
    3) passed
    4) passed
*** CreateObject tests
 1: passed
*** GetObject tests
1: passed
2: passed
3: passed
4: passed
*** Choose tests
0): 
1): Choice1
2): Choice2
3): Choice3
4): Choice4
5): Choice5
6): Choice6
7): 
*** IIF tests
1) passed
2) passed
*** Partition tests
1)    :  0
2)   1: 10
3)  11: 20
4)  21: 30
5)  31: 40
6)  41: 50
7)  5001: 5100
*** Switch tests
red
blue
green

passed
*** Registry Tests
Calling SaveSetting for 9 sections
Calling GetSetting on 9 sections
MyValue1
MyValue2
MyValue3
MyValue4
MyValue5
MyValue6
MyValue7
MyValue8
MyValue9
Calling DeleteSetting on MyKey1
MyKey1 was deleted
Calling DeleteSetting on 3 sections
MySection1 was deleted
MySection2 was deleted
MySection3 was deleted
Calling DeleteSetting on 1 application
MySection4 was deleted
MySection5 was deleted
MySection6 was deleted
MySection7 was deleted
MySection8 was deleted
MySection9 was deleted
Testing GetAllSettings
Dad Test
Dad2 Test2
GetAllSettings of nonexistant key returns Nothing: passed
Testing DeleteSettings: passed
DeleteSetting of nonexistant app: passed
DeleteSetting of nonexistant section: passed
End Test
Shell test passed
