VERSION 5.00
Begin VB.Form about 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2295
   ClientLeft      =   1665
   ClientTop       =   3690
   ClientWidth     =   6765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   2295
   ScaleWidth      =   6765
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "V redu"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1560
      Picture         =   "about.frx":0000
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2002 Ruco"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label verzija 
      BackStyle       =   0  'Transparent
      Caption         =   "verzija"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Avgust, 2002"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Programer: Ruco"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   1440
      Picture         =   "about.frx":0442
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Beep
verzija(1).Caption = "Verzija: " & App.Major & "." & App.Minor & App.Revision
End Sub
