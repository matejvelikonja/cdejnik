VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Popravi vnos"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Preklièi"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Popravi"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "ID:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Zvrst:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ocena:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ime:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Št. CDjev:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
MsgBox ("Vnesite vsa polja!")
Else
    With Form2.Data1.Recordset
        .Edit
            !ID = Text1.Text
            !Ime = Text2.Text
            !Ocena = Text3.Text
            !zvrst = Text4.Text
            !St_CDjev = Text5.Text
        .Update
        .Bookmark = .LastModified
    End With
    Command2_Click
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = Form2.Data1.Recordset!ID
Text2.Text = Form2.Data1.Recordset!Ime
Text3.Text = Form2.Data1.Recordset!Ocena
Text4.Text = Form2.Data1.Recordset!zvrst
Text5.Text = Form2.Data1.Recordset!St_CDjev
End Sub
