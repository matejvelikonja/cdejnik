VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dodaj vnos"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Preklièi"
      Height          =   495
      Left            =   2400
      MaskColor       =   &H000000FF&
      Picture         =   "Form3.frx":030A
      TabIndex        =   11
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dodaj"
      Height          =   495
      Left            =   120
      MaskColor       =   &H000000FF&
      Picture         =   "Form3.frx":0614
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Št. CDjev:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ime:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ocena:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Zvrst:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "ID:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
MsgBox ("Vnesite vsa polja!")
Else
    With Form2.Data1.Recordset
        .AddNew
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
Form2.Data1.Recordset.MoveLast
Text1.Text = Form2.Data1.Recordset!ID + 1
End Sub
