VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Download Calc"
   ClientHeight    =   1380
   ClientLeft      =   1380
   ClientTop       =   2085
   ClientWidth     =   7725
   Icon            =   "Calc.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1380
   ScaleWidth      =   7725
   Begin VB.CommandButton Command2 
      Caption         =   "&About"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "Calc.frx":0442
      Left            =   4680
      List            =   "Calc.frx":0452
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2280
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Calculate"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Created by Martin Porter (Jibajab)"
      Height          =   285
      Left            =   4290
      TabIndex        =   21
      Top             =   1080
      Width           =   3420
   End
   Begin VB.Label Label11 
      Caption         =   "Email me at: Jibajab@hotmail.com"
      Height          =   255
      Left            =   4320
      TabIndex        =   20
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label10 
      Caption         =   "Time in hours/mins"
      Height          =   375
      Left            =   6840
      TabIndex        =   18
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Time in Secs"
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Bps"
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Modem Speed"
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Bits"
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "File Size"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Kb"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "File Size"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "File Size"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Mb"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = Text1.Text * 1024 * 1024
Text3.Text = Text2.Text * 8
Text5.Text = Text3.Text / (Combo1.Text * 1000)
Text8.Text = Int(Text5.Text / 60)
Text7.Text = Text8.Text - 60 * Int(Text8.Text / 60)
Text6.Text = Int(Text8.Text / 60)
MsgBox ("This is a average time only!")
End Sub

Private Sub Command2_Click()
frmAbout.Show

End Sub

Private Sub Form_Load()
MsgBox ("This program was created by Martin Porter")
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MsgBox "Thank you for using my software" & vbCrLf & "please rate this at www.planet-source-code.com/vb", vbOKOnly, "Download Manger"


End Sub
