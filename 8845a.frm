VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MPC AUTOCAL SOFTWARE - 8845A DIGITAL MULTIMETER"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Command5"
      Height          =   375
      Left            =   8400
      TabIndex        =   47
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CheckBox Test5LIN 
      Caption         =   "LINEARITY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   46
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CheckBox Test2 
      Caption         =   "ACV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   39
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   4800
      TabIndex        =   37
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CheckBox Test12 
      Caption         =   "OHM IDC GAIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   35
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CheckBox Test9 
      Caption         =   "LOW IDC GAIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   2280
      TabIndex        =   34
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CheckBox Test10 
      Caption         =   "HI IAC GAIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   360
      TabIndex        =   33
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CheckBox Test11 
      Caption         =   "LOW IAC GAIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   2280
      TabIndex        =   32
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   $"8845a.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   $"8845a.frx":001F
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "EQUIPMENTS REQUIRED"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "GENERAL SETUP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Test Item List:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   4215
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   375
         Left            =   2520
         TabIndex        =   45
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Caption         =   "Rear Panel Zero"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   40
         Top             =   1560
         Width           =   3975
         Begin VB.CheckBox Test3DCV 
            Caption         =   "DCV"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1560
            TabIndex        =   42
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox Test3 
            Caption         =   "OHM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   41
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Front Panel Zero"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   3975
         Begin VB.CheckBox Test2OHM 
            Caption         =   "OHM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2880
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox Test2DCV 
            Caption         =   "DCV"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1560
            TabIndex        =   43
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CheckBox TestOpen 
         Caption         =   "UNLOCK/LOCK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   720
         TabIndex        =   36
         Top             =   4920
         Width           =   1935
      End
      Begin VB.CheckBox Test10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PHASE ACCURACY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   6000
         Width           =   2775
      End
      Begin VB.CheckBox Test11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TEMPERATURE ACCURACY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   6360
         Width           =   2655
      End
      Begin VB.CheckBox Test9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CAPACITANCE ACCURACY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   5640
         Width           =   2655
      End
      Begin VB.CheckBox Test8 
         Caption         =   "HI IDC GAIN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   28
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CheckBox Test7 
         Caption         =   "VDC GAIN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   27
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox TestAll 
         Caption         =   " TEST ALL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CheckBox Test6 
         Caption         =   "ACV GAIN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox Test5 
         Caption         =   "HI I ZERO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   17
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox Test4 
         Caption         =   "LOW I ZERO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CheckBox Test1 
         Caption         =   "OPEN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4440
      TabIndex        =   0
      Top             =   480
      Width           =   5295
      Begin VB.ComboBox CalName_8902 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "8845a.frx":003F
         Left            =   240
         List            =   "8845a.frx":0046
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2880
         Width           =   3855
      End
      Begin VB.ComboBox CalAdd_8902 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "8845a.frx":0063
         Left            =   4200
         List            =   "8845a.frx":00C1
         OLEDragMode     =   1  'Automatic
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox CalName_4284 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "8845a.frx":013D
         Left            =   240
         List            =   "8845a.frx":0144
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2520
         Width           =   3855
      End
      Begin VB.ComboBox CalAdd_4284 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "8845a.frx":0159
         Left            =   4200
         List            =   "8845a.frx":01B7
         OLEDragMode     =   1  'Automatic
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox CalName_5700 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "8845a.frx":0233
         Left            =   240
         List            =   "8845a.frx":0243
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   3855
      End
      Begin VB.ComboBox CalAdd_5700 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "8845a.frx":028F
         Left            =   4200
         List            =   "8845a.frx":02ED
         OLEDragMode     =   1  'Automatic
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox CalName_3458 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "8845a.frx":0369
         Left            =   240
         List            =   "8845a.frx":0373
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1800
         Width           =   3855
      End
      Begin VB.ComboBox CalAdd_3458 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "8845a.frx":03AC
         Left            =   4200
         List            =   "8845a.frx":040A
         OLEDragMode     =   1  'Automatic
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   855
      End
      Begin VB.ComboBox CalName_5790 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "8845a.frx":0486
         Left            =   240
         List            =   "8845a.frx":048D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   3855
      End
      Begin VB.ComboBox CalAdd_5790 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "8845a.frx":04AF
         Left            =   4200
         List            =   "8845a.frx":050D
         OLEDragMode     =   1  'Automatic
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2160
         Width           =   855
      End
      Begin VB.ComboBox InstName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "8845a.frx":0589
         Left            =   240
         List            =   "8845a.frx":0590
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.ComboBox InstAdd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "8845a.frx":05AE
         Left            =   4200
         List            =   "8845a.frx":060C
         OLEDragMode     =   1  'Automatic
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Click Equipments Required to see all the devices needed."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Top             =   3840
      Width           =   4935
   End
   Begin VB.Label Label8 
      Caption         =   "Select in Test Item List the parameter to be evaluated."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4560
      TabIndex        =   24
      Top             =   4560
      Width           =   4575
   End
   Begin VB.Label Label7 
      Caption         =   "Set the proper GPIB address of instruments to be used."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   4200
      Width           =   4695
   End
   Begin VB.Label Label5 
      Caption         =   "IMPORTANT NOTICE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   120
      Picture         =   "8845a.frx":0688
      Top             =   0
      Width           =   2745
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Form1.Hide: MyPIC.Image1.Picture = LoadPicture("C:\temp\hp\34401a_vb\34411a_vb_GEN_SETUP.jpg"): MyPIC.Caption = "MPC AUTOCAL GENERAL SETUP"
Form1.Hide: MyPIC.Image1.Picture = LoadPicture("C:\vee_user\fluke\8845a adj\8845A_Adjust_EQ.jpg"): MyPIC.Caption = "MPC AUTOCAL GENERAL SETUP"

MyPIC.Label1.Caption = "This is the general setup GPIB connection." & Chr(10) & _
       "Click ""OK"" to continue."
MyPIC.BACK.Visible = False
MyPIC.OK.Left = 3800
MainFormTrig = 1
MyPIC.Show 1
MyPIC.BACK.Visible = True
MyPIC.OK.Left = 1995
End Sub

Private Sub Command2_Click()
col = 5: MEASURE_BA
End Sub

Private Sub Command3_Click()
Form1.Hide: MyPIC.Image1.Picture = LoadPicture("C:\vee_user\fluke\8845a adj\8845A_Adjust_EQ.jpg"): MyPIC.Caption = "MPC AUTOCAL EQUIPMENT LIST"
MyPIC.Label1.Caption = "" & Chr(10) & _
       "This is the standards needed." & Chr(10) & _
       "Click ""OK"" to continue."
MyPIC.BACK.Visible = False
MyPIC.OK.Left = 3800
MainFormTrig = 1
MyPIC.Show 1
MyPIC.BACK.Visible = True
MyPIC.OK.Left = 1995
'83483_GEN_INSTRUMENTS.jpg
'MsgBox Form1.InstAdd.Text
End Sub

Private Sub Command4_Click()
MainFormTrig = 0
col = 3: MEASURE_BA
End Sub

Private Sub Command5_Click()
''If InstName.Text = "8846A DIGITAL MULTIMETER" Then MsgBox "8846A DIGITAL MULTIMETER"

''If InstName.Text = "8845A DIGITAL MULTIMETER" Then MsgBox "8845A DIGITAL MULTIMETER"
lp = 129
    UUTCODE = Sheets(SheetName).Cells(129, 2).value
    UUTRANGE = Sheets(SheetName).Cells(129, 1).value
    UUTFREQ = Sheets(SheetName).Cells(129, 7).value
''Sheets(SheetName).Cells(, 7).Select
MsgBox "OUT " & UUTRANGE & " V, " & UUTFREQ & " hz", 10
''Bprint DevInst, "CAL:VAL " & UUTCODE & " V, " & UUTRANGE & " hz", 10
End Sub

Private Sub Command6_Click()
lp = 157
    UUTCODE = Sheets(SheetName).Cells(lp, 2).value
    UUTRANGE = Sheets(SheetName).Cells(lp, 1).value
''Sheets(SheetName).Cells(lp, col).Select
wait 500
''Bprint DevInst, "CAL:VAL ORES,100000000", 10

wait 2000
MsgBox "CAL:VAL " & UUTCODE & "," & UUTRANGE & "CAL:VAL GVDC,10 "
''MsgBox "CAL:VAL " & UUTCODE & ", " & UUTRANGE, 10
End Sub

Private Sub Command7_Click()
''lefttt = Sheets(SheetName).Cells(4, 32).value
Sheets(SheetName).Cells(32, 5).NumberFormat = NumFormat1
 Sheets(SheetName).Cells(32, 4).NumberFormat = NumFormat2
NumFormat1 = NumFormat2
''Sheets(SheetName).Cells(32, 5).NumberFormat = Sheets(SheetName).Cells(32, 4).NumberFormat
''MsgBox UUTCODE
''Sheets(SheetName).Cells(4, 32).NumberFormat = Sheets(SheetName).Cells(5, 32).NumberFormat
 
End Sub

Private Sub Form_Activate()
If FormDet >= 1 Then GoTo FormDetSKIP
        Set Sheet = GetObject("", "Excel.Application")
        Set Sheet = Sheet.Workbooks.Open("C:\vee_user\fluke\8845a adj\Fluke_8845A_adj.xlsm", 0, False)
        Sheet.Application.Visible = 1
        Sheet.Windows(1).Visible = 1
        Application.WindowState = xlMaximized 'maximize Excel
Form1.Show
'Sheet.Sheets(SheetName).Select
'Sheet.Sheets(SheetName).Cells(1, 1).Select
lp = 0
'Dim x As Integer
'Dim lp As Integer
ActLock 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
'InstAdd.ListIndex = 6
'InstName.ListIndex = 0
'CalAdd_5700.ListIndex = 3
'CalName_5700.ListIndex = 0
'CalAdd_3458.ListIndex = 21
'CalName_3458.ListIndex = 0
'CalAdd_83640.ListIndex = 24
'CalName_83640.ListIndex = 0
'CalAdd_8657.ListIndex = 18
'CalName_8657.ListIndex = 0
'CalAdd_8902.ListIndex = 13
'CalName_8902.ListIndex = 0
'CalAdd_8156.ListIndex = 4
'CalName_8156.ListIndex = 0
'CalAdd_8163.ListIndex = 27
'CalName_8163.ListIndex = 0
'CalAdd_6427A.ListIndex = 1
'CalName_6427A.ListIndex = 0
'CalAdd_6427B.ListIndex = 3
'CalName_6427B.ListIndex = 0
'CalAdd_86120.ListIndex = 19
'CalName_86120.ListIndex = 0

ADDRESSid = ADDRESSid + 1: If ADDRESSid > 1 Then GoTo Skip
Skip:
FormDet = FormDet + 1
    For x = 1 To 3000
        If x Mod 2 = 1 Then Form1.Label5.Visible = True: wait 700
        If x Mod 2 = 0 Then Form1.Label5.Visible = False: wait 400
    Next x
FormDetSKIP:
TestExit = False
End Sub
Private Sub Form_Load()
SheetName = "AUTOMATION_SHEET"
'If MsgBox("SELECT WHAT TYPE OF DATA YOU NEED TO CREATE." & Chr(10) & _
'          "YES = COMMERCIAL DATA." & Chr(10) & _
'          " NO = NON-COMMERCIAL.", vbYesNo + vbInformation, "MPC AUTOCAL 3458A") = vbNo Then
'    SheetName = "DATA_UNC"
'    CommercialCheck = 2
'    Else
'    SheetName = "DATA_NO_UNC"
'    CommercialCheck = 0
'End If

Dim lR As Long
lR = SetTopMostWindow(Form1.hwnd, True)
InstAdd.ListIndex = 9
InstName.ListIndex = 0
CalAdd_5700.ListIndex = 3
CalName_5700.ListIndex = 0
'CalAdd_3458.ListIndex = 21
'CalName_3458.ListIndex = 0
'CalAdd_5520.ListIndex = 3
'CalName_5520.ListIndex = 2
'CalAdd_4284.ListIndex = 18
'CalName_4284.ListIndex = 0
'CalAdd_8902.ListIndex = 13
'CalName_8902.ListIndex = 0
'CalAdd_8156.ListIndex = 4
'CalName_8156.ListIndex = 0
'CalAdd_8163.ListIndex = 27
'CalName_8163.ListIndex = 0
'CalAdd_6427A.ListIndex = 1
'CalName_6427A.ListIndex = 0
'CalAdd_6427B.ListIndex = 3
'CalName_6427B.ListIndex = 0
'CalAdd_86120.ListIndex = 19
'CalName_86120.ListIndex = 0
End Sub

Private Sub Form_Terminate()
Unload Me
PROGRESS1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unload Me
PROGRESS1.Show
End Sub

'Private Sub InstAdd_GotFocus()
''InstAdd.Locked = False
'End Sub
'Private Sub InstAdd_LostFocus()
''InstAdd.Locked = True
'End Sub
'Private Sub CalAdd_3458_GotFocus()
'CalAdd_3458.Locked = False
'End Sub
'Private Sub CalName_3458_LostFocus()
'CalAdd_3458.Locked = True
'End Sub
'Private Sub CalAdd_5520A_GotFocus()
'CalAdd_5500.Locked = False
'End Sub
'Private Sub CalAdd_5520A_LostFocus()
'CalAdd_5500.Locked = True
'End Sub
'Private Sub CalAdd_6427A_GotFocus()
'CalAdd_6427A.Locked = False
'End Sub
'Private Sub Caladd_6427A_LostFocus()
'CalAdd_6427A.Locked = True
'End Sub
'Private Sub Caladd_6427B_GotFocus()
'CalAdd_6427B.Locked = False
'End Sub
'Private Sub Caladd_6427B_LostFocus()
'CalAdd_6427B.Locked = True
'End Sub
'Private Sub CalAdd_8156_GotFocus()
'CalAdd_8156.Locked = False
'End Sub
'Private Sub CalAdd_8156_LostFocus()
'CalAdd_8156.Locked = True
'End Sub
'Private Sub CalAdd_8163_GotFocus()
'CalAdd_8163.Locked = False
'End Sub
'Private Sub CalAdd_8163_LostFocus()
'CalAdd_8163.Locked = True
'End Sub
'Private Sub CalAdd_83640_GotFocus()
'CalAdd_83640.Locked = False
'End Sub
'Private Sub CalAdd_83640_LostFocus()
'CalAdd_83640.Locked = True
'End Sub
'Private Sub CalAdd_86120_GotFocus()
'CalAdd_86120.Locked = False
'End Sub
'Private Sub CalAdd_86120_LostFocus()
'CalAdd_86120.Locked = True
'End Sub
'Private Sub Caladd_8657_GotFocus()
'CalAdd_8657.Locked = False
'End Sub
'Private Sub CalAdd_8657_LostFocus()
'CalAdd_8657.Locked = True
'End Sub
'Private Sub Caladd_8902_GotFocus()
'CalAdd_8902.Locked = False
'End Sub
'Private Sub CalAdd_8902_LostFocus()
'CalAdd_8902.Locked = True
'End Sub
Private Sub Label8_Click()

End Sub
Private Sub TEST_DC_ELEC_Click()
If TEST_DC_ELEC.value = 1 Then
Test2.value = 1: Test3.value = 1: Test4.value = 1: Test5.value = 1: Test6.value = 1
Test7.value = 1: Test8.value = 1: Test9.value = 1: Test10.value = 1: Test11.value = 1
End If
End Sub

Private Sub Test1_Click()
If Form1.Test1 = 0 Then Form1.TestAll = 0
End Sub

Private Sub Test2_Click()
If Form1.Test2 = 0 Then Form1.TestAll = 0
End Sub
Private Sub Test3_Click()
If Form1.Test3 = 0 Then Form1.TestAll = 0
End Sub
Private Sub Test4_Click()
If Form1.Test4 = 0 Then Form1.TestAll = 0
End Sub
Private Sub Test5_Click()
If Form1.Test5 = 0 Then Form1.TestAll = 0
End Sub
Private Sub Test6_Click()
If Form1.Test6 = 0 Then Form1.TestAll = 0
End Sub
Private Sub Test7_Click()
If Form1.Test7 = 0 Then Form1.TestAll = 0
End Sub
Private Sub Test8_Click()
If Form1.Test8 = 0 Then Form1.TestAll = 0
End Sub




Private Sub TestAll_Click()
If Form1.TestAll = 1 Then
    Form1.Test1 = 1
    Form1.Test2 = 1
    Form1.Test3 = 1
    Form1.Test4 = 1
    Form1.Test5 = 1
    Form1.Test6 = 1
    Form1.Test7 = 1
    Form1.Test8 = 1
    Form1.Test9 = 1
    Form1.Test10 = 1
    Form1.Test11 = 1
End If
End Sub

Private Sub TestOpen_Click()
''If TestOpen.value = 0 Then
''ActLock 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
''Test1.Enabled = False
''Test2.Enabled = False
''Test3.Enabled = False
''Test4.Enabled = False
''Test5.Enabled = False
''Test6.Enabled = False
''Test7.Enabled = False
''Test8.Enabled = False
''Test9(1).Enabled = False
''Test10(1).Enabled = False
''Test11(1).Enabled = False
''Test12.Enabled = False
''Test13.Enabled = False
''TestAll.Enabled = False
''Else
''''ActLock 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1
''Test1.Enabled = True
''Test2.Enabled = True
''Test3.Enabled = True
''Test4.Enabled = True
''Test5.Enabled = True
''Test6.Enabled = True
''Test7.Enabled = True
''Test8.Enabled = True
''Test9(1).Enabled = True
''Test10(1).Enabled = True
''Test11(1).Enabled = True
''Test12.Enabled = True
''Test13.Enabled = True
''TestAll.Enabled = True

End Sub
