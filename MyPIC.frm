VERSION 5.00
Begin VB.Form MyPIC 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MPC AUTOCAL"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9960
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BACK 
      BackColor       =   &H008080FF&
      Caption         =   "&BACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   2400
   End
   Begin VB.CommandButton OK 
      BackColor       =   &H008080FF&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1990
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   2400
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   5340
      Left            =   600
      Picture         =   "MyPIC.frx":0000
      Top             =   960
      Width           =   9300
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TEST CAPTION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      TabIndex        =   2
      Top             =   6360
      Width           =   8775
   End
   Begin VB.Image Image 
      Height          =   735
      Left            =   240
      Picture         =   "MyPIC.frx":30A56
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "MyPIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BACK_Click()
MyPIC.Hide
Form1.Show
TestExit = True
End Sub

Private Sub Form_Terminate()
Unload MyPIC
Form1.Show 1
End Sub

Private Sub OK_Click()
MyPIC.Hide
If GENSETexit = True Then Form1.Show: GENSETexit = False
If MainFormTrig = 1 Then Form1.Show
'TestExit = True
End Sub
