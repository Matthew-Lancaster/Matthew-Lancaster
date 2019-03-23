VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3012
   ClientLeft      =   60
   ClientTop       =   528
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3012
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim E1 As Double

'-------------------------------------
'CALC OUR SPEED FROM KNOWN MESSUREMENT

E1 = TimeSerial(2, 25, 46) ' TIME TAKEN
M1 = 5.9 'MILES

MPH = E1 / M1 '1.71570935342122E-02
MPH = 1.71570935342122E-02
Debug.Print MPH
'-------------------------------------


'-------------------------------------
M2 = 4.2 ' MILES NEW ROUTE
M2 = 1.3 ' MILES NEW ROUTE

'M2 AT MPH =

E2 = M2 * MPH ' CALC THE DOUBLE NUMBER
Dim E3 As Date ' COPY THE DOUBLE CALC NUMBER INTO A DATE DECLARE VAR
E3 = E2 ' 01:43:46


Debug.Print E3
'-------------------------------------
' --- RESULT 00:32:07 TO WALK MEET CLARE AND FEMALE
' --- AT LAGOON PROM BY BIG BEACH CAFE 01 APR 2016
'-------------------------------------



'-----------------------------------
'------------ ACCOUNT EXTRACTION ---
'-----------------------------------
'Fri-01-Jul-2016 09:52:19 AM
'LAST WORKING MAYBE FIRST -- 15 APR
'-----------------------------------
'VALUE OF HALF HOUSE PAID £240,000 AND OTHER SETTLEMENT WAS HALF MINUS £10,000
'HALF VALUE = £120,000
Debug.Print "-----------------------------------"
Debug.Print "------------ ACCOUNT EXTRACTION ---"
Debug.Print "-----------------------------------"
Debug.Print Format(Now, "DDD-DD-MMM-YYYY HH:MM:SS Am/Pm")
Debug.Print "---------------------------"


Dim AD2 As Double
Dim AD3 As Double
Dim AD8 As Double
Dim AD4 As Currency

AD2 = (230 / 240)
Debug.Print "DIFFERENCE IN % BETWEEN 240K AND 10K ----"
Debug.Print "SUM £230K DIVEDE £240 (230 / 240) ------- =" + Str(AD2)

'NOT USED ---------------------------------------------------------------------------------------------
'Debug.Print "£240K DIVEDE OFFSET MINUS £10K IS £230K LESS BY SECOND PARTY =%" + Str(AD2)
'Debug.Print "£240K DIVEDE OFFSET MINUS £10K IS £230K LESS BY SECOND PARTY * 100% = %" + Str(AD2 * 100)
'------------------------------------------------------------------------------------------------------
' --- THE 4.166 RESULT
Debug.Print "INVERT OPPOSITE RESULT - 100 = %" + Str(100 - (AD2 * 100)) + " INVERT"
AD8 = 100 - (AD2 * 100) ' RESULT TO KEEP FINAL 4.166
Debug.Print "REVERT OPPOSITE RESULT ----- = %" + Str(AD2 * 100) + " NOT INVERT"
AD3 = 120 / AD2

Debug.Print "HALF £240K -- DIVEDE BY 2 = £120 AND THE OFFSET PREVIOUS RESULT % =" + Str(AD3)
'AD4 HAS A DIMENSION DIM DECLARE CURRENCY FORMAT
'THAT USED TO CONVERT AD3 DOUBLE PRECION VARIABLE TO THIS AD4 CURRENCY VARIABLE BY TRANSFER ONE INTO OTHER
AD4 = AD3
Debug.Print "HALF £240K -- DIVEDE BY 2 = £120 AND THE OFFSET PREVIOUS RESULT % IN £ CURRENCY ROUNDED =" + Str(AD4)
Debug.Print "HALF £240K -- DIVEDE BY 2 = £120 AND THE OFFSET PREVIOUS RESULT % IN £ CURRENCY ROUNDED = " + Format(AD4, "0.00")

'NOT USED --------------------------------------------------------------------
'Debug.Print "RESULT % 4.16 OF 10K £ =" + Str(10000/4.16)
'Debug.Print "RESULT TOTAL BY HALF 80+80 = 160 * % .958333 =" + Str(160 * AD2)
'-----------------------------------------------------------------------------

Debug.Print "RESULT TOTAL BY HALF AT 80K * % .958333 =" + Str(80 * AD2)
AD5 = 100 * (50 / (AD2 * 100))

Debug.Print "RESULT -- 50% 50% OFFSET SLIDE SWING IS"
TEXT2 = " 50% OFFSET SWING"
Debug.Print "RESULT -- CHANGE OFFSET AT 50% 50*(RESULT .958333*100) =" + Str(AD5) + TEXT2
Debug.Print "OR INVERT FROM BACK 100"
AD7 = 100 - (100 * (50 / (AD2 * 100)))
Debug.Print "RESULT -- CHANGE OFFSET AT 50% 50*(RESULT .958333*100) =" + Str(AD7) + TEXT2 + " INVERT"
TEXT1 = "ADD BOTH BACK TOGETHER TO CHECK MATCH PAIR COMPARE RESULT VERIFY ="
Debug.Print TEXT1 + Str((AD5) + (AD7))
TEXT3 = "MYSELF PAID FULL SHARE ON THIRD 80K AND EXTRA TO HALF 40K = 120K OF 240K TOTAL"
Debug.Print TEXT3
TEXT3 = "DIFFERENCE OF OFFSET 10K ON 240K IS --" + Str(AD8) + "%"
Debug.Print TEXT3

TEXT3 = "--------------------"
Debug.Print TEXT3
TEXT3 = "THINK 10K DIFFERENCE - SAGE SEER ACCOUNT EXTRACTOR"
Debug.Print TEXT3
TEXT3 = "10K DIFFERENCE OF % IS --" + Str(AD8) + "£"
Debug.Print TEXT3

TEXT3 = "--------------------"
Debug.Print TEXT3
TEXT3 = "SHARE COST OF 10K DIFFERENCE DIVEDE BY 2"
Debug.Print TEXT3
TEXT3 = "10K SHORTFALL WE BOTH GAIN HALF OF"
Debug.Print TEXT3

TEXT3 = "--------------------"
Debug.Print TEXT3
TEXT3 = "NOT A PERFECT EXTRACTOR ACCOUNTING AS LEFT WITH QUESTION -- 50% QUESTION -- OR -- MONEY REMAINING QUESTION"
Debug.Print TEXT3

TEXT3 = "10K DIFFERENCE OF % IS --" + Str(AD8 / 2) + "£ SHARED BY 2"
Debug.Print TEXT3

TEXT3 = "OR DIFFERENCE"
Debug.Print TEXT3
TEXT3 = "5K DIFFERENCE OF % IS --" + Str(5 * AD8) + "£"
Debug.Print TEXT3
TEXT3 = "THINK REAL DIFFERENCE IS 10K THOUGH AS I PAID UPTO THE FULL HALF 80+40=120"
Debug.Print TEXT3
TEXT3 = "THE OFFSET 10K IS NOT MY PARTY TO BE PARTIAL CONTRIBUTE TOWARD"
Debug.Print TEXT3
TEXT3 = "FOLLOW CONCLUSION RESULT 50 50% OWNER VALUE OR OFFSET MONEY REMAINING"
Debug.Print TEXT3
TEXT3 = "ANSWER 1 OF 2 .. 50 50 = "
Debug.Print TEXT3
Debug.Print "--" + Str(AD5) + " OR" + Str(AD7) + TEXT2 + " PROPERTY OWNERSHIP"
Debug.Print "YES 50 50% SOUND EASIER APPROACH ANSWER -- BUT -- AND THEN RAISE QUESTION"
TEXT3 = "ANSWER 2 OF 2 .. MONEY REMAIN UN-ACCOUNTED FOR"
Debug.Print TEXT3
Debug.Print "--" + " " + Format(AD4, "0.00") + " WITH 120K£ MINUS = " + Format(AD4 - 120, "0.00")
TEST4 = "MONEY TO BE GAIN " + Format(AD4 - 120, "0.00")
TEST5 = " -- OR -- MONEY TO BE LOSS " + Format(10 - (AD4 - 120), "0.00") + " "
'USE MID$ MANIPULATE TO CHOP ROUNDED FIGURE OFF END AS NOT SHOW TRUE TO THE WAY CALC WAS DONE
'CALC USED FULL FIGURE NOT ROUNDED IN DOUBLE PRECISION DIMENSION DECLARE VARIABLE
TEST7 = "WHEN DIFFERENCE % WORKING IS " + Mid(Format(AD8, "0.000000"), 1, 7)
Debug.Print TEST4 + TEST5 + TEST7
'-----------------------------
TEXT3 = "--------------------"
Debug.Print TEXT3
Debug.Print "ACCOUNT ACCURACY PROBLEM HERE -------------------- 01 OF 02"
Debug.Print "IS IT -- MONEY TO BE LOSS 120K£ WITH DIFFERENCE PERCENT"
Debug.Print "ACCOUNT ACCURACY PROBLEM HERE -------------------- 02 OF 02"
Debug.Print "OR IS IT -- -- %%"


'-----------------------------
TEXT3 = "--------------------"
Debug.Print TEXT3
TEXT3 = "BALANCE THE FIGURE -- NOT A EXPERT WILL EXTRACTOR ACCOUNTANT"
Debug.Print TEXT3
TEXT3 = "HERE IS THE CODE WORKING OUT AND THE CODE RESULT OUTPUT"
Debug.Print TEXT3
TEXT3 = "GUESS VISUAL BASIC MAKES A SUPER CALC WHEN EDITABLE Formula AND HISTORY OF WORKING"
Debug.Print TEXT3
TEXT3 = "SHARE OF THE SEGMENT SECTION PIE SLICE SECTOR SELECTOR FAIR GENEROUS Genius GEM EQUAL"
Debug.Print TEXT3
TEXT3 = "DIVEDE -- MULTIPLY"
Debug.Print TEXT3


End


End Sub
