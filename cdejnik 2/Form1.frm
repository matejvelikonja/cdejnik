VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   2160
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Programi"
      Height          =   855
      Left            =   4320
      TabIndex        =   1
      Top             =   6360
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Igre"
      Height          =   855
      Left            =   4320
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "By: Ruco™"
      BeginProperty Font 
         Name            =   "OzHandicraft BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8760
      TabIndex        =   4
      Top             =   2400
      Width           =   825
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1800
      Left            =   480
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2040
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "CDejnik"
      BeginProperty Font 
         Name            =   "Staccato222 BT"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   8775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show MODAL
Unload Me
End Sub

Private Sub Command2_Click()
Form1.Visible = False
End Sub

Private Sub Form_Click()
Command1_Click
End Sub

Private Sub Form_Load()
Label4.Caption = Label2.Caption
Label5.Caption = Label2.Caption
End Sub

Private Sub Label2_Change()
Label4.Caption = Label2.Caption
Label5.Caption = Label2.Caption
End Sub

Private Sub Timer1_Timer()
Label4.Caption = Label2.Caption
Label5.Caption = Label2.Caption
If Label2.Caption = "1" Then
Command1_Click
Else
Label2.Caption = Label2.Caption - 1
End If
End Sub
