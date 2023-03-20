VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form PROGRESS1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP AUTOCAL SOFTWARE"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ProgressBar1 
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   495
         Left            =   -30
         TabIndex        =   1
         Top             =   -30
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "PROGRESS1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
PROGRESS1.Visible = True
'If lp = 0 Then lp = 1: GoTo Skip
PROGRESS1.Caption = "MP AUTOCAL SOFTWARE EXITING"
PROGRESS1.ProgressBar2.Min = 0: PROGRESS1.ProgressBar2.Max = 5000
For x = PROGRESS1.ProgressBar2.Min To PROGRESS1.ProgressBar2.Max Step 1
PROGRESS1.ProgressBar2.value = x: wait 0
Next x
Skip:
End
End Sub

Private Sub Form_Terminate()
PROGRESS1.Visible = True
If lp = 0 Then lp = 1: GoTo Skip
PROGRESS1.Caption = "MP AUTOCAL SOFTWARE EXITING"
PROGRESS1.ProgressBar2.Min = 0: PROGRESS1.ProgressBar2.Max = 3000
For x = PROGRESS1.ProgressBar2.Min To PROGRESS1.ProgressBar2.Max Step 1
PROGRESS1.ProgressBar2.value = x: wait 0
Next x
Skip:
'End
End Sub

