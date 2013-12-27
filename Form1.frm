VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4095
   ClientLeft      =   15
   ClientTop       =   3255
   ClientWidth     =   1215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   1215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Igre"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   2880
   End
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "Izhod"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "O programu"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Nastavitve"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ostalo"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Filmi"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Programi"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If (Dir(App.Path & "\Files\cdejnik.mdb") <> "") Then
           With Form1
            .Show modal
            .igra.DatabaseName = App.Path & "\Files\cdejnik.mdb"
            .igra.Refresh
            .Caption = "CDejnik " & App.Major & "." & App.Minor & " (" & .igra.RecordSource & ")"
           End With
           
           With main
            .Command2.Enabled = False
            .Command3.Enabled = False
            .Command4.Enabled = False
           End With
           
        Else
            MsgBox "Datoteke z zbirko ni mogoèe najti.", vbExclamation
            Exit Sub
        End If
End Sub

Private Sub Command2_Click()
If (Dir(App.Path & "\Files\cdejnik.mdb") <> "") Then
           With programi
            .Show modal
            .progi.DatabaseName = App.Path & "\Files\cdejnik.mdb"
            .progi.Refresh
            .Caption = "CDejnik " & App.Major & "." & App.Minor & " (" & .progi.RecordSource & ")"
           End With
           
           With main
            .Command1.Enabled = False
            .Command3.Enabled = False
            .Command4.Enabled = False
           End With
           
        Else
            MsgBox "Datoteke z zbirko ni mogoèe najti.", vbExclamation
            Exit Sub
        End If
End Sub

Private Sub Command3_Click()
If (Dir(App.Path & "\Files\cdejnik.mdb") <> "") Then
           With film
            .Show modal
            .filmi.DatabaseName = App.Path & "\Files\cdejnik.mdb"
            .filmi.Refresh
            .Caption = "CDejnik " & App.Major & "." & App.Minor & " (" & .filmi.RecordSource & ")"
           End With
           
           With main
            .Command2.Enabled = False
            .Command1.Enabled = False
            .Command4.Enabled = False
           End With
           
        Else
            MsgBox "Datoteke z zbirko ni mogoèe najti.", vbExclamation
            Exit Sub
        End If
End Sub

Private Sub Command4_Click()
If (Dir(App.Path & "\Files\cdejnik.mdb") <> "") Then
           With ostalo
            .Show modal
            .ostane.DatabaseName = App.Path & "\Files\cdejnik.mdb"
            .ostane.Refresh
            .Caption = "CDejnik " & App.Major & "." & App.Minor & " (" & .ostane.RecordSource & ")"
           End With
           
           With main
            .Command2.Enabled = False
            .Command3.Enabled = False
            .Command1.Enabled = False
           End With
           
        Else
            MsgBox "Datoteke z zbirko ni mogoèe najti.", vbExclamation
            Exit Sub
        End If
End Sub

Private Sub Command5_Click()
nastavitve.Show modal
End Sub

Private Sub Command6_Click()
about.Show modal
End Sub

Private Sub Command7_Click()
SaveSetting "CDejnik", "Nastavitve", "Zadnjic", (Date & " ob " & Time)
Timer2.Enabled = True
End Sub

Private Sub Form_Activate()
Timer1.Enabled = True
ChDir App.Path
End Sub

Private Sub Form_Load()
main.Left = 0
main.Top = 0

pot = App.Path & "\Pics"

If (Dir(pot, vbDirectory) <> "") Then 'obstaja pics?
    l = 0
Else
    MkDir (pot) 'ustvari mapo
    
    If (Dir(App.Path & "\nopic.cd3") <> "") Then
    FileCopy App.Path & "\nopic.cd3", pot & "\nopic.jpg" 'kopiraj
    Kill App.Path & "\nopic.cd3" 'unièi staro
    End If
    
    If (Dir(App.Path & "\Readme.cd3") <> "") Then
    FileCopy App.Path & "\Readme.cd3", pot & "\Readme.txt"
    End If
    
End If

potka = App.Path & "\Files"

If (Dir(potka, vbDirectory) <> "") Then 'obstaja pics?
    l = 1
Else
    MkDir (potka) 'ustvari mapo
    
    If (Dir(App.Path & "\cdejnik.cd3") <> "") Then
    FileCopy App.Path & "\cdejnik.cd3", potka & "\cdejnik.mdb" 'kopiraj
    Kill App.Path & "\cdejnik.cd3" 'unièi staro
    End If
    
    If (Dir(App.Path & "\Readme.cd3") <> "") Then
    FileCopy App.Path & "\Readme.cd3", potka & "\Readme.txt"
    Kill App.Path & "\Readme.cd3"
    End If
    
    If (Dir(App.Path & "\users.cd3") <> "") Then
    FileCopy App.Path & "\users.cd3", potka & "\users.mdb" 'kopiraj
    Kill App.Path & "\users.cd3" 'unièi staro
    End If
    
    If (Dir(App.Path & "\uvod.cd3") <> "") Then
    FileCopy App.Path & "\uvod.cd3", potka & "\uvod.swf"
    Kill App.Path & "\uvod.cd3"
    End If
End If

End Sub

Private Sub Timer1_Timer()
If main.Top < 3240 Then
main.Top = main.Top + 60
Else
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If main.Top = 0 Then
End
Else
main.Top = main.Top - 60
End If
End Sub
