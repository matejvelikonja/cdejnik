VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form popravi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Popravi vnos"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox lep 
      Height          =   285
      Left            =   4560
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   6
      ToolTipText     =   "Vnesite ime vnosa [max. 30 znakov]"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1920
      MaxLength       =   25
      TabIndex        =   5
      ToolTipText     =   "Sem vnesite zvrst [max. 25 znakov]"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   4
      ToolTipText     =   "Sem vnesite število zgošèenk, ki jih vaš vnos vsebuje [max. 2 znaka]"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Prebrskaj"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      ToolTipText     =   "Poišèi sliko."
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Popravi"
      Default         =   -1  'True
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Preklièi"
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   2880
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog odpri 
      Left            =   4560
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Izberite sliko"
      Filter          =   "Slike|*.jpg|"
   End
   Begin VB.Label plot 
      Height          =   135
      Left            =   4320
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label preverip 
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   480
      Stretch         =   -1  'True
      ToolTipText     =   "Dvoklik za izbris slike. POZOR !!!! Izbrisalo bo tudi sliko na disku !!!"
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ID:"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Ime:"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Zvrst:"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Število CDejev:"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Slika:"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "popravi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function preveri(str As String, ch As String) As String
Dim a As Integer

  preveri = ""

For a = 1 To Len(str)
  If Not Mid(str, a, 1) = ch Then
   preveri = preveri & Mid(str, a, 1)
  End If
Next

End Function

Private Sub kontrola()
'preveri, ali lahko uporabi znake za ime slike
lep.Text = preveri(lep.Text, ":")
lep.Text = preveri(lep.Text, "*")
lep.Text = preveri(lep.Text, "/")

lep.Text = preveri(lep.Text, "\")
lep.Text = preveri(lep.Text, "|")
lep.Text = preveri(lep.Text, "?")

lep.Text = preveri(lep.Text, "<")
lep.Text = preveri(lep.Text, ">")
lep.Text = preveri(lep.Text, """")
End Sub

Private Sub Command1_Click()
With odpri
    .Flags = cd10fnfilemustexist
    .FileName = ""
    .DialogTitle = "Odpri sliko"
    .CancelError = True
On Error GoTo napaka
    .ShowOpen
    pot = .FileName
    
Text5.Text = pot
napaka:
    End With
End Sub

Private Sub Command2_Click()
    If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
    MsgBox ("Vnesite vsa polja!")
    Else
        If Text5.Text = "" Then Text5.Text = "Slike ni"
            If (Dir(Text5.Text) <> "") Then
                
                        kontrola
                        y = App.Path & "\Pics\" & Text1.Text & lep.Text & ".jpg"
                        x = Text5.Text
                If plot.Caption <> Text5.Text Then
                    If Dir(plot.Caption) <> "" Then
                        Kill (plot.Caption) 'izbrise staro datoteko
                    End If
                        If x <> y Then
                            FileCopy x, y
                            Text5.Text = y
                        End If
                End If
            ElseIf Text5.Text = "Slike ni" Then
                    vili = 1
            Else
                        MsgBox "Napaka v datoteki!", vbCritical, napaka
                        Exit Sub
            End If
            If preverip.Caption = "film" Then 'filmi
                With film.filmi.Recordset
                    .Edit
                        !id = Text1.Text
                        !ime = Text2.Text
                        !zvrst = Text3.Text
                        !St_CDejev = Text4.Text
                        !slika = Text5.Text
                    .Update
                    .Bookmark = .LastModified
                End With
                If Text5.Text = "Slike ni" Then
                    film.slika.Picture = LoadPicture(App.Path & "\Pics\nopic.jpg")
                Else
                    film.slika.Picture = LoadPicture(Text5.Text) 'naloži sliko
                End If
            ElseIf preverip.Caption = "igra" Then 'igre
                With Form1.igra.Recordset
                    .Edit
                        !id = Text1.Text
                        !ime = Text2.Text
                        !zvrst = Text3.Text
                        !St_CDejev = Text4.Text
                        !slika = Text5.Text
                    .Update
                    .Bookmark = .LastModified
                End With
                If Text5.Text = "Slike ni" Then
                    Form1.slika.Picture = LoadPicture(App.Path & "\Pics\nopic.jpg")
                Else
                    Form1.slika.Picture = LoadPicture(Text5.Text) 'naloži sliko
                End If
            ElseIf preverip.Caption = "progi" Then 'programi
                With programi.progi.Recordset
                    .Edit
                        !id = Text1.Text
                        !ime = Text2.Text
                        !zvrst = Text3.Text
                        !St_CDejev = Text4.Text
                        !slika = Text5.Text
                    .Update
                    .Bookmark = .LastModified
                End With
                If Text5.Text = "Slike ni" Then
                    programi.slika.Picture = LoadPicture(App.Path & "\Pics\nopic.jpg")
                Else
                    programi.slika.Picture = LoadPicture(Text5.Text) 'naloži sliko
                End If
            ElseIf preverip.Caption = "ostanek" Then 'ostalo
                With ostalo.ostane.Recordset
                    .Edit
                        !id = Text1.Text
                        !ime = Text2.Text
                        !zvrst = Text3.Text
                        !St_CDejev = Text4.Text
                        !slika = Text5.Text
                    .Update
                    .Bookmark = .LastModified
                End With
                If Text5.Text = "Slike ni" Then
                    ostalo.slika.Picture = LoadPicture(App.Path & "\Pics\nopic.jpg")
                Else
                    ostalo.slika.Picture = LoadPicture(Text5.Text) 'naloži sliko
                End If
            End If
                Command3_Click
    End If
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If preverip.Caption = "film" Then 'izpolni okna
            Text1.Text = film.filmi.Recordset!id
            Text2.Text = film.filmi.Recordset!ime
            Text3.Text = film.filmi.Recordset!zvrst
            Text4.Text = film.filmi.Recordset!St_CDejev
            Text5.Text = film.filmi.Recordset!slika
    ElseIf preverip.Caption = "igra" Then
            Text1.Text = Form1.igra.Recordset!id
            Text2.Text = Form1.igra.Recordset!ime
            Text3.Text = Form1.igra.Recordset!zvrst
            Text4.Text = Form1.igra.Recordset!St_CDejev
            Text5.Text = Form1.igra.Recordset!slika
    ElseIf preverip.Caption = "progi" Then
            Text1.Text = programi.progi.Recordset!id
            Text2.Text = programi.progi.Recordset!ime
            Text3.Text = programi.progi.Recordset!zvrst
            Text4.Text = programi.progi.Recordset!St_CDejev
            Text5.Text = programi.progi.Recordset!slika
    ElseIf preverip.Caption = "ostanek" Then
            Text1.Text = ostalo.ostane.Recordset!id
            Text2.Text = ostalo.ostane.Recordset!ime
            Text3.Text = ostalo.ostane.Recordset!zvrst
            Text4.Text = ostalo.ostane.Recordset!St_CDejev
            Text5.Text = ostalo.ostane.Recordset!slika
    End If
    plot.Caption = Text5.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
'omogoèi nazaj okno
    If preverip.Caption = "film" Then
        film.Enabled = True
    ElseIf preverip.Caption = "igra" Then
        Form1.Enabled = True
    ElseIf preverip.Caption = "progi" Then
        programi.Enabled = True
    ElseIf preverip.Caption = "ostanek" Then
        ostalo.Enabled = True
    End If
End Sub

Private Sub Image1_DblClick()
pot = Text5.Text
If (Dir(pot) <> "") Then
    Kill (pot)
End If
Text5.Text = ""
End Sub


Private Sub Text2_Change()
lep.Text = Text2.Text
End Sub

Private Sub Text4_Change()
Static vnos
If Not IsNumeric(Text4.Text) Then
    Text4.Text = vnos
    Beep
Else
    vnos = Text4.Text
End If
End Sub

Private Sub Text5_Change()
If (Dir(Text5.Text) <> "") Then
Image1.Picture = LoadPicture(Text5.Text)
Else
Image1.Picture = LoadPicture(App.Path & "\Pics\nopic.jpg")
End If
End Sub
