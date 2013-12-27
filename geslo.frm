VERSION 5.00
Begin VB.Form geslo 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2085
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Izhod"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Potrdi"
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   10
      PasswordChar    =   "?"
      TabIndex        =   1
      ToolTipText     =   "Sem vpišite geslo"
      Top             =   720
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vtipkajte pravilno geslo:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "geslo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
gesko = GetSetting("CDejnik", "Nastavitve", "PGeslo", "")
If Text1.Text = gesko Then
    Unload Me
    zacetno.Show modal
Else
    MsgBox "Geslo ni pravilno", vbCritical
    Text1.Text = ""
    Exit Sub
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Activate()
If App.PrevInstance = True Then
    MsgBox "Program je že pognan!", vbInformation
    End
Else
    geslic = GetSetting("CDejnik", "Nastavitve", "Geslo", "0")
    
    If geslic = 0 Then
        Unload Me
        zacetno.Show modal
    End If
End If
End Sub

