Attribute VB_Name = "module1"
Public Sub MyTest1(): 'Open Verification
Sheets(SheetName).Cells(36, col).Select
If TEST_SETUP("1", "OPEN ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 0, 0: If CalErrLog = True Then Exit Sub

    Bprint DevInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
lp = 36
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
''Bprint DevInst, "CAL:VAL ORES,100000000", 10
Bprint DevInst, "CAL? ON", 10
wait 10000
MyPBAR PbarStart, 1, 3000
Bprint DevInst, "CAL:VAL?", 10
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
Bprint DevInst, "*CLS", 10
''MsgBox MyTempValue
''Bprint DevInst, "CAL:REC", 2000
''MsgBox MyTempValue
ilocal DevInst: iclose DevInst
PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0

End Sub

Public Sub MyTest12(): 'OHM GAIN
heets(SheetName).Cells(210, col).Select
If TEST_SETUP("5", "OHM GAIN ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 1, 0: If CalErrLog = True Then Exit Sub
    Bprint DevInst, "*RST", 10
    Bprint CalInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
TestAgain:
For lp = 210 To 216
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select
Bprint CalInst, "OUT " & UUTRANGE & " Ohm", 10: wait 500
''Bprint DevInst, "CAL:VAL ORES,100000000", 10
Bprint CalInst, "OPER", 500: wait 500
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
Bprint DevInst, "CAL? ON", 10
wait 10000
MyPBAR PbarStart, 7, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
Bprint CalInst, "STBY", 100
''MsgBox MyTempValue
If lp = 216 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
If lp = 216 Then If MsgBox("OHM Gain Adjustment Completed, Do you want to save the pass Adjustment?", vbYesNo + vbExclamation, "Test Confirmation") = vbNo Then lp = 210: GoTo TestAgain
Else
GoTo SaveAdj
End If
Next lp
SaveAdj:
''Bprint DevInst, "CAL:REC", 2000
ilocal DevInst: iclose DevInst
ilocal CalInst: iclose CalInst: Bprint CalInst, "STBY", 100
End Sub

Public Sub MyTest1old(): 'Open Verification
Sheets(SheetName).Cells(36, col).Select
If TEST_SETUP("1", "OPEN VERIFICATION") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 0, 0: If CalErrLog = True Then Exit Sub
       'RESET EQUIPMENTS
      
    Bprint DevInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
lp = 36
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select

    If lp = 36 Then
    Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
    End If
Bprint DevInst, "CAL? ON", 100
''Bprint DevInst, "CAL:VAL?", 1000
''MyTempValue = getdata(CInt(DevInst))
MyPBAR PbarStart, 1, 3000
MyTempValue = getdata(DevInst)
Sheets(SheetName).Cells(lp, col).value = MyTempValue
Bprint DevInst, "CAL:REC", 10
Bprint DevInst, "*RST", 100: Bprint DevInst, "*CLS", 100
ilocal DevInst
iclose DevInst
End Sub
Public Sub MyTest1new(): 'Open Verification
Sheets(SheetName).Cells(36, col).Select
If TEST_SETUP("1", "OPEN VERIFICATION") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 0, 0: If CalErrLog = True Then Exit Sub
       'RESET EQUIPMENTS
      
    Bprint DevInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
lp = 36
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select
    ''If lp = 36 Then
    Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 1000
    ''End If
Bprint DevInst, "CAL? ON", 1000
''Bprint DevInst, "CAL:VAL?", 1000
MyPBAR PbarStart, 1, 3000
''MyTempValue = getdata(DevInst)
''Sheets(SheetName).Cells(lp, col).value = MyTempValue
''Bprint DevInst, "CAL:REC", 10
''Bprint DevInst, "*RST", 100: Bprint DevInst, "*CLS", 1000
PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0

ilocal DevInst
iclose DevInst
End Sub

Public Sub MyTest2OHM():  'ACV / DCV/ OHM ZERO
Sheets(SheetName).Cells(67, col).Select
If TEST_SETUP("2", "OHM ZERO ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 0, 0: If CalErrLog = True Then Exit Sub
       'RESET EQUIPMENTS
    Bprint DevInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
For lp = 67 To 72
MsgBox lp
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
Bprint DevInst, "CAL? ON", 1000
wait 10000
''Bprint DevInst, "CAL:VAL?", 1000
MyPBAR PbarStart, 6, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, 7).value = MyTempValue
If lp = 72 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
If lp = 72 Then Bprint DevInst, "CAL:REC", 2000
Next lp
ilocal DevInst: iclose DevInst
End Sub

Public Sub MyTest2DCV():  'ACV / DCV/ OHM ZERO
Sheets(SheetName).Cells(56, col).Select
If TEST_SETUP("2", "DCV ZERO") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 0, 0: If CalErrLog = True Then Exit Sub
       'RESET EQUIPMENTS
    Bprint DevInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
For lp = 56 To 61
MsgBox lp
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
Bprint DevInst, "CAL? ON", 1000
wait 10000
''Bprint DevInst, "CAL:VAL?", 1000
MyPBAR PbarStart, 6, 3000
''Bprint DevInst, "CAL:VAL?", 10
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
''wait 500
''Bprint DevInst, "*CLS", 10
If lp = 61 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
If lp = 61 Then Bprint DevInst, "CAL:REC", 2000
Next lp
ilocal DevInst: iclose DevInst
End Sub
Public Sub MyTest2new(): 'Open Verification
Sheets(SheetName).Cells(42, col).Select
If TEST_SETUP("1", "ZERO OFFSET VERIFICATION") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 0, 0: If CalErrLog = True Then Exit Sub
       'RESET EQUIPMENTS
    Bprint DevInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
For lp = 42 To 72
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select
TextBack:
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 2000
MyPBAR PbarStart, 10, 3000
Bprint DevInst, "CAL? ON", 1000: wait 2500
Bprint DevInst, "*CLS", 1000: wait 2000
If lp = 51 Then lp = 56: GoTo TextBack
If lp = 61 Then lp = 67: GoTo TextBack
''If lp = 51 Then lp = 56: GoTo textback
''If lp = 61 Then lp = 67: GoTo textback
If lp = 72 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
Next lp
''Bprint DevInst, "*RST", 100 '' Bprint DevInst, "*CLS", 100
ilocal DevInst
iclose DevInst


End Sub
    
Public Sub MyTest3DCV(): 'REAR OHM / DCV ZERO
Sheets(SheetName).Cells(87, col).Select
If TEST_SETUP("2", "REAR DCV ZERO ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 0, 0: If CalErrLog = True Then Exit Sub
       'RESET EQUIPMENTS
    Bprint DevInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
For lp = 87 To 88
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 100
Bprint DevInst, "CAL? ON", 1000
wait 10000
''Bprint DevInst, "CAL:VAL?", 1000
MyPBAR PbarStart, 2, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
''If lp = 81 Then lp = 87: GoTo TextBack
If lp = 88 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
If lp = 88 Then Bprint DevInst, "CAL:REC", 2000
Next lp
ilocal DevInst: iclose DevInst
End Sub

Public Sub MyTest10(): 'LOW IAC GAIN
Sheets(SheetName).Cells(186, col).Select
If TEST_SETUP("4", "LOW IAC GAIN ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 1, 0: If CalErrLog = True Then Exit Sub
    Bprint DevInst, "*RST", 10
    Bprint CalInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
TestAgain:
For lp = 186 To 191
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
    UUTFREQ = Sheets(SheetName).Cells(lp, 7).value
Sheets(SheetName).Cells(lp, col).Select
Bprint CalInst, "OUT " & UUTRANGE & " A," & UUTFREQ & " hz", 10
wait 500
Bprint CalInst, "OPER", 500
wait 500
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
Bprint DevInst, "CAL? ON", 10
wait 10000
MyPBAR PbarStart, 6, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
Bprint CalInst, "STBY", 100
''MsgBox MyTempValue
If lp = 191 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
If lp = 191 Then If MsgBox("HI IAC GAIN Adjustment Completed, Do you want to save the pass Adjustment?", vbYesNo + vbExclamation, "Test Confirmation") = vbNo Then lp = 186: GoTo TestAgain
Else
GoTo SaveAdj
End If
Next lp
SaveAdj:
''Bprint DevInst, "CAL:REC", 2000
ilocal DevInst: iclose DevInst
ilocal CalInst: iclose CalInst: Bprint CalInst, "STBY", 100
End Sub
Public Sub MyTest11(): 'LOW IDC GAIN
Sheets(SheetName).Cells(197, col).Select
If TEST_SETUP("4", "LOW IDC GAIN ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 1, 0: If CalErrLog = True Then Exit Sub
    Bprint DevInst, "*RST", 10
    Bprint CalInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
TestAgain:
For lp = 197 To 204
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select

Bprint CalInst, "OUT " & UUTRANGE & " A", 10: wait 500
''Bprint DevInst, "CAL:VAL ORES,100000000", 10
Bprint CalInst, "OPER", 500: wait 500
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
Bprint DevInst, "CAL? ON", 10
wait 10000
MyPBAR PbarStart, 8, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
Bprint CalInst, "STBY", 100
''MsgBox MyTempValue
If lp = 204 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
If lp = 204 Then If MsgBox("LOW IDC GAIN Adjustment Completed, Do you want to save the pass Adjustment?", vbYesNo + vbExclamation, "Test Confirmation") = vbNo Then lp = 197: GoTo TestAgain
Else
GoTo SaveAdj
End If
Next lp
SaveAdj:
''Bprint DevInst, "CAL:REC", 2000
ilocal DevInst: iclose DevInst
ilocal CalInst: iclose CalInst: Bprint CalInst, "STBY", 100
End Sub


Public Sub MyTest5LIN(): 'LINEARITY
Sheets(SheetName).Cells(121, col).Select
If TEST_SETUP("3", "LINEARITY ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 1, 0: If CalErrLog = True Then Exit Sub
    Bprint DevInst, "*RST", 10
    Bprint CalInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
    
For lp = 121 To 124
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
    UUTFREQ = Sheets(SheetName).Cells(lp, 7).value
Sheets(SheetName).Cells(lp, col).Select
Bprint CalInst, "OUT " & UUTRANGE & " V," & UUTFREQ & " hz", 10
wait 500
Bprint CalInst, "OPER", 500
wait 1000
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
Bprint DevInst, "CAL? ON", 10
wait 10000
MyPBAR PbarStart, 4, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
Bprint CalInst, "STBY", 300
''MsgBox MyTempValue
If lp = 124 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
If lp = 124 Then If MsgBox("LINEARITY Adjustment Completed, Do you want to save the pass Adjustment?", vbYesNo + vbExclamation, "Test Confirmation") = vbNo Then lp = 121: GoTo TestAgain
Else
GoTo SaveAdj
End If
Next lp
SaveAdj:
Bprint DevInst, "CAL:REC", 2000
ilocal DevInst: iclose DevInst
ilocal CalInst: iclose CalInst: Bprint CalInst, "STBY", 100
End Sub

Public Sub MyTest9(): 'HI IAC GAIN
Sheets(SheetName).Cells(177, col).Select
If TEST_SETUP("5", "HI IAC GAIN ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 1, 0: If CalErrLog = True Then Exit Sub
    Bprint DevInst, "*RST", 10
    Bprint CalInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
TestAgain:
For lp = 177 To 180
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
    UUTFREQ = Sheets(SheetName).Cells(lp, 7).value
Sheets(SheetName).Cells(lp, col).Select
Bprint CalInst, "OUT " & UUTRANGE & " A," & UUTFREQ & " hz", 10
wait 500
''Bprint DevInst, "CAL:VAL ORES,100000000", 10
Bprint CalInst, "OPER", 500
wait 500
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
Bprint DevInst, "CAL? ON", 10
wait 10000
MyPBAR PbarStart, 4, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
Bprint CalInst, "STBY", 100
''MsgBox MyTempValue
If lp = 180 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
If lp = 180 Then If MsgBox("HI IAC GAIN Adjustment Completed, Do you want to save the pass Adjustment?", vbYesNo + vbExclamation, "Test Confirmation") = vbNo Then lp = 177: GoTo TestAgain
Else
GoTo SaveAdj
End If
Next lp
SaveAdj:
''Bprint DevInst, "CAL:REC", 2000
ilocal DevInst: iclose DevInst
ilocal CalInst: iclose CalInst: Bprint CalInst, "STBY", 100
End Sub

Public Sub MyTestOpen(): 'Open Verification
If TEST_SETUP("1", "ZERO OFFSET VERIFICATION") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 0, 0: If CalErrLog = True Then Exit Sub
''If Form1.InstName.Text = "8845A DIGITAL MULTIMETER" Then Exit Sub
Bprint DevInst, "*RST", 10: wait 2000
''Bprint DevInst, "CAL:SEC:STAT OFF, FLUKE884X", 1000
Bprint DevInst, "CAL:SEC:STAT?", 10
MyTempValue = getdata(CInt(DevInst))
''MsgBox MyTempValue
If MyTempValue = 1 Then Bprint DevInst, "CAL:SEC:STAT OFF, FLUKE884X", 2000
Bprint DevInst, "CAL:SEC:STAT?", 10
MyTempValue = getdata(CInt(DevInst))
    If Form1.TestOpen.value = 0 Then
        ActLock 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
    Else
        ActLock 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1
    End If
Form1.TestOpen.value = 0
Form1.TestOpen.Enabled = False

End Sub
'*****************************************************************************************************
'*****************************************************************************************************
Public Sub MyTest2():  'ACV / DCV/ OHM ZERO
Sheets(SheetName).Cells(42, col).Select
If TEST_SETUP("2", "ACV ZERO ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 0, 0: If CalErrLog = True Then Exit Sub
       'RESET EQUIPMENTS
    Bprint DevInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
For lp = 42 To 51
MsgBox lp
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
Bprint DevInst, "CAL? ON", 1000
wait 10000
MyPBAR PbarStart, 10, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
If lp = 51 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
If lp = 51 Then Bprint DevInst, "CAL:REC", 2000
Next lp

ilocal DevInst: iclose DevInst
End Sub

'*****************************************************************************************************
'*****************************************************************************************************
Public Sub MyTest3(): 'REAR OHM / DCV ZERO
Sheets(SheetName).Cells(78, col).Select
If TEST_SETUP("2", "REAR OHM ZERO ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 0, 0: If CalErrLog = True Then Exit Sub
       'RESET EQUIPMENTS
    Bprint DevInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
For lp = 78 To 81
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 100
Bprint DevInst, "CAL? ON", 1000
wait 10000
''Bprint DevInst, "CAL:VAL?", 1000
MyPBAR PbarStart, 4, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
''If lp = 81 Then lp = 87: GoTo TextBack
If lp = 81 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
If lp = 81 Then Bprint DevInst, "CAL:REC", 2000
Next lp
ilocal DevInst: iclose DevInst
End Sub
'*****************************************************************************************************
Public Sub MyTest4(): 'LOW I ZERO
Sheets(SheetName).Cells(94, col).Select
If TEST_SETUP("2", "LOW I ZERO ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 0, 0: If CalErrLog = True Then Exit Sub
    Bprint DevInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
    
For lp = 94 To 105
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
''Bprint CalInst, "VOLT " & UUTRANGE, 10
''Bprint CalInst, ""
''Bprint DevInst, "CAL:VAL ORES,100000000", 10
Bprint DevInst, "CAL? ON", 10
wait 10000
MyPBAR PbarStart, 12, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
''MsgBox MyTempValue
''Bprint DevInst, "CAL:REC", 2000
''MsgBox MyTempValue
If lp = 105 Then Bprint DevInst, "CAL:REC", 2000
Next lp
ilocal DevInst: iclose DevInst
PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
End Sub

'*****************************************************************************************************
Public Sub MyTest5(): 'HI I ZERO
Sheets(SheetName).Cells(111, col).Select
If TEST_SETUP("2", "HIGH I ZERO ADJUSTMENT ") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 0, 0: If CalErrLog = True Then Exit Sub
    Bprint DevInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
    
For lp = 111 To 116
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
''Bprint CalInst, "OUT " & UUTRANGE, 10
''Bprint CalInst, ""
''Bprint DevInst, "CAL:VAL ORES,100000000", 10
Bprint DevInst, "CAL? ON", 10
wait 10000
MyPBAR PbarStart, 6, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
''MsgBox MyTempValue
''Bprint DevInst, "CAL:REC", 2000
''MsgBox MyTempValue
If lp = 116 Then Bprint DevInst, "CAL:REC", 2000
Next lp
ilocal DevInst: iclose DevInst
PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
End Sub
'*****************************************************************************************************
Public Sub MyTest6(): 'ACV GAIN
Sheets(SheetName).Cells(132, col).Select
If TEST_SETUP("3", "ACV GAIN ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 1, 0: If CalErrLog = True Then Exit Sub
    Bprint DevInst, "*RST", 10
    Bprint CalInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
TestAgain:
For lp = 132 To 147
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
    UUTFREQ = Sheets(SheetName).Cells(lp, 7).value
    UUTRANGEAC = Sheets(SheetName).Cells(lp, 8).value
Sheets(SheetName).Cells(lp, col).Select
If lp = 147 Then Bprint CalInst, "OUT " & UUTRANGEAC & " V, " & UUTFREQ & " hz", 10
Else
Bprint CalInst, "OUT " & UUTRANGE & " V, " & UUTFREQ & " hz", 10
End If
''Bprint DevInst, "CAL:VAL ORES,100000000", 10
Bprint CalInst, "OPER", 500
wait 500
Bprint DevInst, "CAL:VAL " & UUTCODE, 10
Bprint DevInst, "CAL? ON", 10
wait 10000
MyPBAR PbarStart, 16, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
Bprint CalInst, "STBY", 100
''MsgBox MyTempValue
If lp = 147 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
If lp = 147 Then If MsgBox("ACV GAIN Adjustment Completed, Do you want to save the pass Adjustment?", vbYesNo + vbExclamation, "Test Confirmation") = vbNo Then lp = 132:   GoTo TestAgain:
Else
GoTo SaveAdj
End If
Next lp
SaveAdj:
''Bprint DevInst, "CAL:REC", 2000
ilocal DevInst: iclose DevInst
ilocal CalInst: iclose CalInst: Bprint CalInst, "STBY", 100
End Sub
'*****************************************************************************************************
Public Sub MyTest7(): 'VDC GAIN VERIFICATION (2)
Sheets(SheetName).Cells(153, col).Select
If TEST_SETUP("3", "VDC GAIN ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 1, 0: If CalErrLog = True Then Exit Sub
    Bprint DevInst, "*RST", 10
    Bprint CalInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
TestAgain:
For lp = 153 To 162
''lp = 157
MsgBox lp
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select
MsgBox "CALIBRATOR:" & "OUT " & UUTRANGE & " V", 10
Bprint CalInst, "OUT " & UUTRANGE & " V", 10
wait 500
''Bprint DevInst, "CAL:VAL ORES,100000000", 10
Bprint CalInst, "OPER", 500
wait 2000
MsgBox "8845A:" & "CAL:VAL " & UUTCODE & "," & UUTRANGE, 10
Bprint DevInst, "CAL:VAL " & UUTCODE & "," & UUTRANGE, 10
Bprint DevInst, "CAL? ON", 10
wait 10000
MyPBAR PbarStart, 10, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
Bprint CalInst, "STBY", 100
''MsgBox MyTempValue
If lp = 162 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
If lp = 162 Then If MsgBox("VDC GAIN Adjustment Completed, Do you want to save the pass Adjustment?", vbYesNo + vbExclamation, "Test Confirmation") = vbNo Then lp = 153: GoTo TestAgain

Next lp
SaveAdj:
''Bprint DevInst, "CAL:REC", 2000
ilocal DevInst: iclose DevInst
ilocal CalInst: iclose CalInst: Bprint CalInst, "STBY", 100

End Sub
'*****************************************************************************************************
Public Sub MyTest8(): 'HI IDC GAIN
Sheets(SheetName).Cells(168, col).Select
If TEST_SETUP("5", "HI IDC GAIN ADJUSTMENT") = True Then: MainSkipFlag = True: Exit Sub
Bopen_All 1, 1, 0: If CalErrLog = True Then Exit Sub
    Bprint DevInst, "*RST", 10
    Bprint CalInst, "*RST", 10: wait 2000 'UUT
    PROGRESS2.Show 0: wait 1000
TestAgain:
For lp = 168 To 171
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
Sheets(SheetName).Cells(lp, col).Select
Bprint CalInst, "OUT " & UUTRANGE & " A", 10
wait 500
''Bprint DevInst, "CAL:VAL ORES,100000000", 10
Bprint CalInst, "OPER", 500
wait 500
Bprint DevInst, "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
Bprint DevInst, "CAL? ON", 10
wait 10000
MyPBAR PbarStart, 4, 3000
MyTempValue = getdata(CInt(DevInst))
Sheets(SheetName).Cells(lp, col).value = MyTempValue
Bprint CalInst, "STBY", 100
''MsgBox MyTempValue
If lp = 171 Then PROGRESS2.Hide: PROGRESS2.ProgressBar2.value = 0
If lp = 171 Then If MsgBox("HI IDC GAIN Adjustment Completed, Do you want to save the pass Adjustment?", vbYesNo + vbExclamation, "Test Confirmation") = vbNo Then lp = 168: GoTo TestAgain
Else
GoTo SaveAdj
End If
Next lp
SaveAdj:
''Bprint DevInst, "CAL:REC", 2000
ilocal DevInst: iclose DevInst
ilocal CalInst: iclose CalInst: Bprint CalInst, "STBY", 100
End Sub

