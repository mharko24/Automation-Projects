Attribute VB_Name = "Module1"
Global DevInst As Long
Global CalInst As Long
Global CalInst2 As Long
Global CalInst3 As Long
Global CalInst4 As Long
Global CalInst5 As Long
Global CalInst6 As Long
Global CalInst7 As Long
Global CalInst8 As Long
Global CalInst9 As Long
Global CalInst10 As Long
Global CalInst11 As Long
Global CalInst12 As Long

Global y As String * 25
Global z As Long
Global col As Long
Global mydata As Single
Global ADDRESSid As Integer
Global lp As Integer
Global lp2 As Integer
Global Myprog As Long
Global PbarStart As Long
Global PbarStop As Long
Global PbarVAL As Long
Global DA As Variant
Const Giga = 1000000000
Const Mega = 1000000
Const Kilo = 1000
Const Milli = 0.001
Const Micro = 0.00001

Global PowerLP As Single
Global PowerLEVEL As Single
Global My128MHz As Single
Global MyPMeter As Single
Global ATT300HZ As Single
Global ATT3KHZ As Single
Global Atten As Single
Global RefLevel As Single

Global FormDet As Integer
Global CalErrLog As Boolean
Global TestExit As Boolean
Global GENSETexit As Boolean
Global Test17_verify As Boolean
Global OpticalPower As Single
Global OpticPow1 As Single
Global sHuntval As Double
'Global SheetName As String

Global MyZeroOffset As Integer
Global SheetName As String
Global CommercialCheck As Integer
Global CommercialLp As Integer
'Global GPIBflag As Boolean

Global MyTempValue
Global MyTempValue2
Global DMMvalue
Global PMaccuracy

Global NewFocus_WAV1
Global NewFocus_WAV2
Global NewFocus_WAVjump
Global NewFocus_LEV1
Global NewFocus_LEV2
Global NewFocus_LEVjump
Global DEVIATION1
Global DEVIATION2
'*******************OPTICAL VARIABLES
    Global Atten500 As Single
    Global Atten300 As Single
    Global Atten200 As Single
    Global Atten100 As Single
    Global Atten60 As Single
    Global Atten20 As Single
    'Global Atten As Single
Global TEST_ALL_ELECTRICAL As Boolean
Global MainSkipFlag As Boolean

Global my86120_wave1 As Single
Global my86120_pow1 As Single

Global my86120_wave2 As Single
Global my86120_pow2 As Single

Global Mreciv50MHZ
Global Mreciv12GHZ
Global Mreciv20GHZ
Global Smodule50MHZ
Global Smodule12GHZ
Global Smodule20GHZ
Global MainFormTrig As Integer
Global MyFREQ As String
Global DMMflag As Boolean
'============for form load front=====================
Public Const SWP_NOMOVE = 2
      Public Const SWP_NOSIZE = 1
      Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
      Public Const HWND_TOPMOST = -1
      Public Const HWND_NOTOPMOST = -2
 Declare Function SetWindowPos Lib "user32" _
            (ByVal hwnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) As Long


Sub t()
Dim x As Integer
Dim Mytime
Mytime = Date
MsgBox Mytime
End Sub
Sub ghjkl()
'PROGRESS1.Visible ': PROGRESS1.ProgressBar1.Enabled = True
PROGRESS1.ProgressBar1.Min = 0
PROGRESS1.ProgressBar1.Max = 1000
Dim t As Integer
For t = PROGRESS1.ProgressBar1.Min To PROGRESS1.ProgressBar1.Max Step 1
If t = 1 Then PROGRESS1.Show 0
PROGRESS1.ProgressBar1.value = t
wait 20
Next t
PROGRESS1.Hide
End Sub
Sub klkl()
ghjkl
wait 2000
wait 3000
wait 4000
End Sub

Function Imprenta(id As Integer, mydelay As Long)
Dim segment As Long
Dim x As Long
segment = mydelay / 200
For x = 20 To mydelay Step 20
If x = 0 Then PROGRESS1.ProgressBar1.value = Myprog + 20
PROGRESS1.ProgressBar1.value = PROGRESS1.ProgressBar1.value + 20
wait 10 ': If x Mod 320 = 20 Then PROGRESS1.ProgLbl.Caption = "Configuring Equipments" Else PROGRESS1.ProgLbl.Caption = ""
Next x
Myprog = PROGRESS1.ProgressBar1.value
End Function


Sub ttttt()
Dim CalInst2 As Integer
Dim DevInst As Integer
PROGRESS1.ProgressBar1.Min = 0
PROGRESS1.ProgressBar1.Max = 7200
PROGRESS1.Show 0

Imprenta CalInst2, 200
Imprenta CalInst2, 200
Imprenta CalInst2, 200
Imprenta CalInst2, 2000
Imprenta CalInst2, 200
Imprenta CalInst2, 3000
Imprenta CalInst2, 200

Imprenta DevInst, 200
Imprenta DevInst, 200
Imprenta DevInst, 200
Imprenta DevInst, 200
Imprenta DevInst, 200
Imprenta DevInst, 200
PROGRESS1.Hide: End
'MsgBox "finished"
End Sub

Sub HP3458()
Dim x As Integer
Dim HP As Long
Dim SA As String
BOpen HP, 12
wait 1000
'PROGRESS1.ProgressBar1.Min = 0
'PROGRESS1.ProgressBar1.Max = 27200
Bprint HP, "RESET"
'Bprint HP, "FUNC DCV, 10, 0.00001", 400, False

For x = 0 To 20
'If x = 0 Then PROGRESS1.Show 0
Bprint HP, "FUNC ACV", 400, False
Bprint HP, "ARANGE", 400, False
Bprint HP, "TRIG AUTO", 400, False
vb_itermchr CalInst, Asc(" ")
'iread HP, y, 100, &O0, z
SA = getdataCOUNTER(CInt(HP), 1, 16) * 1000000
'Bprint HP, "RESET"
'PROGRESS1.ProgressBar1.value = x
ilocal HP
'wait 200
Next x
End Sub


Sub Dev86100_INIT(MyTrigger As String, Optional MyCH1 As Integer, Optional MyCH2 As Integer, _
    Optional MyCH3 As Integer, Optional MyCH4 As Integer, Optional AllChScale As String, _
    Optional AllChOffs As String, Optional TimeBase_Range As String, _
    Optional AllChBW As String, Optional TimeBase_Delay As String, _
    Optional AverageCnt As Integer, Optional AverageOn As Integer, _
    Optional MyMeasureSource As Integer, Optional MyMeasurePoint As String)
    'Dev86100_INITIALIZE "FPANEL", 0, 1, 0, 0, "60E-3", "100e-12", "26E-9", 64, 1, "MEASURE:PERIOD CHANNEL2"
    Bprint DevInst, "*RST", 1000
    Bprint DevInst, "*CLS", 200
    Bprint DevInst, "SYSTEM:HEADER OFF", 200
    Bprint DevInst, "SYST:MODE OSC", 200
    Bprint DevInst, "TRIG:SOUR " & MyTrigger, 200
    Bprint DevInst, "CHAN1:DISP " & MyCH1, 200
    Bprint DevInst, "CHAN2:DISP " & MyCH2, 200
    Bprint DevInst, "CHAN3:DISP " & MyCH3, 200
    Bprint DevInst, "CHAN4:DISP " & MyCH4, 200
    For lp = 1 To 4
    If lp = 1 And MyCH1 = 0 Then
        GoTo Skip
    ElseIf lp = 1 And MyCH1 = 1 Then GoTo proceed
    End If
    
    If lp = 2 And MyCH2 = 0 Then
        GoTo Skip
    ElseIf lp = 2 And MyCH2 = 1 Then GoTo proceed
    End If
    
    If lp = 3 And MyCH3 = 0 Then
        GoTo Skip
        ElseIf lp = 3 And MyCH3 = 1 Then GoTo proceed
    End If
    
    If lp = 4 And MyCH4 = 0 Then
        GoTo Skip
        ElseIf lp = 4 And MyCH4 = 1 Then GoTo proceed
    End If
proceed:
        Bprint DevInst, "CHAN" & lp & ":BAND " & AllChBW, 200
        Bprint DevInst, "CHAN" & lp & ":SCAL " & AllChScale, 200
        Bprint DevInst, "CHAN" & lp & ":OFFS " & AllChOffs, 200
Skip:
    Next lp
    If TimeBase_Range <> "" Then Bprint DevInst, "TIMEBASE:RANGE " & TimeBase_Range, 200
    If TimeBase_Delay <> "" Then Bprint DevInst, "TIMEBASE:DELAY " & TimeBase_Delay, 200
    If AverageCnt <> 0 Then Bprint DevInst, "ACQ:COUN " & AverageCnt, 200
    If AverageOn <> 0 Then Bprint DevInst, "ACQ:AVER " & AverageOn, 200
    If MyMeasureSource <> 0 Then Bprint DevInst, "MEASURE:SOURCE CHANNEL" & MyMeasureSource, 200
    If MyMeasurePoint <> "" Then Bprint DevInst, MyMeasurePoint, 200
    Bprint DevInst, "MEAS:SEND ON", 200
    Bprint DevInst, "MEAS:ANN ON", 200
    
    
'Dev86100_INIT "FRUN", 0, 1, 0, 0, "1E-3", "0", , "LOW", , 16, 1, 2, "MEASURE:AVERAGE DISPLAY,CHANNEL2"
    
'Bprint DevInst, "TRIG:SOUR FRUN", 300
'Bprint DevInst, "ACQ:COUN 16", 300
'Bprint DevInst, "ACQ:AVER 1", 300
'Bprint DevInst, "CHAN2:DISP 1", 300
'Bprint DevInst, "CHAN1:DISP 0", 300
'Bprint DevInst, "CHAN2:BAND LOW", 300
'Bprint DevInst, "CHAN2:OFFS 0", 300
'Bprint DevInst, "CHAN2:SCAL 1E-3", 300
'Bprint DevInst, "MEASURE:SOURCE CHANNEL2", 500
'Bprint DevInst, "MEASURE:VAVERAGE DISPLAY,CHANNEL2", 500
'Bprint DevInst, "MEAS:SEND ON", 500
'Bprint DevInst, "MEAS:ANN ON", 500

ilocal DevInst
    
End Sub

Sub Cal8133_INIT(MyTrigger As String, Optional MyTrigCnt As Integer, Optional MyFunction_Wform As String, _
    Optional MyAmplitude_V As Single, Optional MyOffset_V As Single, Optional MyOutput_ONOFF As String)
    'Cal8133_INIT "EXT", 1, SQU, 2.5, 0, "OFF"
    Bprint CalInst2, "*RST", 1000
    Bprint CalInst2, "*CLS", 500
    Bprint CalInst2, "TRIG:SOUR " & MyTrigger, 300
    Bprint CalInst2, "TRIG:ECOUNT " & MyTrigCnt, 300
    For lp = 1 To 2
        Bprint CalInst2, "FUNC" & lp & ":SHAP " & MyFunction_Wform, 200
        Bprint CalInst2, "VOLT" & lp & ":LEV:AMPL " & MyAmplitude_V & "V", 200
        Bprint CalInst2, "VOLT" & lp & ":LEV:OFFS " & MyOffset_V & "V", 200
        Bprint CalInst2, "OUTPUT" & lp & ":STAT " & MyOutput_ONOFF, 200
    Next lp
End Sub

Sub Cal8360_INIT(MyMode As String, MyFREQ As Single, MyFreqUnit As String, _
    MyPowLevel As Single, MyPowerStat As String)
    'Cal8360_INIT "CW", 19.98, "GHZ", 0, "OFF"
    Bprint CalInst, "*RST", 1000
    Bprint CalInst, "SYST:LANG SCPI", 1000
    Bprint CalInst, "FREQ:MODE " & MyMode, 100
    Bprint CalInst, "FREQ:CW " & MyFREQ & MyFreqUnit, 100
    Bprint CalInst, "POW:LEV " & MyPowerLev & " DBM", 100
    Bprint CalInst, "POW:STAT " & MyPowerStat, 100
End Sub

Sub Cal8340_INIT(MyMode As String, MyFREQ As Single, MyFreqUnit As String, MyPowLevel As Single, _
    MyPowerStat As Integer)
    'Cal8340_INIT "CW", 19.98, "GZ", 0, 0
    Bprint CalInst, "IP", 1000
    Bprint CalInst, MyMode & MyFREQ & MyFreqUnit, 100
    Bprint CalInst, "PL" & MyPowerLev & "DB", 100
    Bprint CalInst, "RF" & MyPowerStat, 100
End Sub

Sub Cal5500_INIT(MyVolt1 As String, Optional MyVolt2 As String, Optional MyCmd As String)
'Cal5700_INIT 1, 0.5

    Bprint DevInst, "*RST", 2000
    Bprint DevInst, "*CLS", 500
    Bprint DevInst, "OUT " & MyVolt1, 200
    'Bprint DevInst, "RANGELCK ON", 500
    If MyVolt2 <> "" Then Bprint DevInst, "OUT " & MyVolt2, 200
    Bprint DevInst, "OPER", 100
End Sub
Sub Cal4284_INIT(Optional oninFreq As String)
'Cal5700_INIT 1, 0.5

    Bprint CalInst4, "*RST;*CLS", 1000
'    Bprint CalInst4, "CORR:OPEN: STAT ON", 500
'    Bprint CalInst4, "CORR:SHOR: STAT ON", 200
    Bprint CalInst4, "FREQ " & oninFreq, 200
If TEST_SETUP("9A", "CAPACITANCE ACCURACY") = True Then: Exit Sub
PROGRESS2.Show
Bprint CalInst4, "CORR:OPEN", 100
MyPBAR PbarStart, 1, 900000: wait 8000
PROGRESS2.Hide
PROGRESS2.ProgressBar1.value = 0
If TEST_SETUP("9B", "CAPACITANCE ACCURACY") = True Then: Exit Sub
Bprint CalInst4, "CORR:SHOR", 100
PROGRESS2.Show 0
MyPBAR PbarStart, 1, 900000: wait 8000
PROGRESS2.Hide
PROGRESS2.ProgressBar1.value = 0

End Sub



Sub Ilocal_All(id1 As Long, Optional id2 As Long, Optional id3 As Long, Optional id4 As Long, _
    Optional id5 As Long, Optional id6 As Long, Optional id7 As Long, Optional id8 As Long, _
    Optional id9 As Long, Optional id10 As Long, Optional id11 As Long, Optional id12 As Long)
    If id1 <> 0 Then ilocal id1: If id2 <> 0 Then ilocal id2
    If id3 <> 0 Then ilocal id3: If id4 <> 0 Then ilocal id4
    If id5 <> 0 Then ilocal id5: If id6 <> 0 Then ilocal id6
    If id7 <> 0 Then ilocal id7: If id8 <> 0 Then ilocal id8
    If id9 <> 0 Then ilocal id9: If id8 <> 0 Then ilocal id10
    If id11 <> 0 Then ilocal id11: If id8 <> 0 Then ilocal id12
End Sub

Sub Iclose_All(id1 As Long, Optional id2 As Long, Optional id3 As Long, Optional id4 As Long, _
    Optional id5 As Long, Optional id6 As Long, Optional id7 As Long, Optional id8 As Long, _
    Optional id9 As Long, Optional id10 As Long, Optional id11 As Long, Optional id12 As Long)
    If id1 <> "" Then iclose id1: If id2 <> "" Then iclose id2
    If id3 <> "" Then iclose id3: If id4 <> "" Then iclose id4
    If id5 <> "" Then iclose id5: If id6 <> "" Then iclose id6
    If id7 <> "" Then iclose id7: If id8 <> "" Then iclose id8
    If id9 <> "" Then iclose id9: If id10 <> "" Then iclose id10
    If id11 <> "" Then iclose id11: If id12 <> "" Then iclose id12
End Sub

Sub Preset_All(id1 As Long, Preset1 As String, Optional id2 As Long, Optional Preset2 As String, _
 Optional id3 As Long, Optional Preset3 As String, Optional id4 As Long, Optional Preset4 As String, _
 Optional id5 As Long, Optional Preset5 As String, Optional id6 As Long, Optional Preset6 As String, _
 Optional id7 As Long, Optional Preset7 As String, Optional id8 As Long, Optional Preset8 As String)
    If id1 <> "" Then Bprint id1, Preset1: If id2 <> "" Then Bprint id2, Preset2
    If id3 <> "" Then Bprint id3, Preset3: If id4 <> "" Then Bprint id4, Preset4
    If id5 <> "" Then Bprint id5, Preset5: If id6 <> "" Then Bprint id6, Preset6
    If id7 <> "" Then Bprint id7, Preset7: If id8 <> "" Then Bprint id8, Preset8
End Sub

Sub Bopen_All(Dev34410 As Long, Optional Cinst_5520 As Long, Optional Cinst2_3458 As Long, Optional Cinst5_8902 As Long, _
 Optional Cinst6 As Long, Optional Cinst7 As Long, Optional Cinst8 As Long)
    If Dev34410 = 1 Then BOpen DevInst, CLng(Form1.InstAdd.Text): If CalErrLog = True Then Exit Sub
    If Cinst_5520 = 1 Then BOpen CalInst, CLng(Form1.CalAdd_5520.Text): If CalErrLog = True Then Exit Sub
    If Cinst2_3458 = 1 Then BOpen CalInst2, CLng(Form1.CalAdd_3458.Text): If CalErrLog = True Then Exit Sub
End Sub

Function TEST_SETUP(SetupNum As String, MyCaption As String) As Boolean
TEST_SETUP = False
Form1.Hide: MyPIC.Image1.Picture = LoadPicture("C:\vee_user\hp\34401a_4to1\34401a_4to1_SETUP" & SetupNum & ".jpg"): MyPIC.Caption = "MP AUTOCAL SOFTWARE -" & MyCaption
MyPIC.Label1.Caption = "Connect equipments same as above." & Chr(10) & _
                         "When connection is ""OK"" click to start measurement."
If MyZeroOffset = 1 Then
   MyPIC.BACK.Visible = False: MyPIC.OK.Left = 3720
   Else
   MyPIC.BACK.Visible = True: MyPIC.OK.Left = 1990
End If
MyPIC.Show 1: If TestExit = True Then TestExit = False: TEST_SETUP = True: Exit Function
If TestExit = True Then TestExit = False: TEST_SETUP = True: Exit Function
End Function

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
         As Long

         If Topmost = True Then 'Make the window topmost
            SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
               0, FLAGS)
         Else
            SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
               0, 0, FLAGS)
            SetTopMostWindow = False
         End If
End Function

