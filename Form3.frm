VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form nov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nov vnos"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5475
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lep 
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog odpri 
      Left            =   4440
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Izberite sliko"
      Filter          =   "Slike|*.jpg|"
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Preklièi"
      Height          =   615
      Left            =   2880
      TabIndex        =   12
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Dodaj"
      Default         =   -1  'True
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Prebrskaj"
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      ToolTipText     =   "Poišèi sliko"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "Pot do datoteke s slike. Slika bo prekopirana v imenik CDejnika."
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   7
      ToolTipText     =   "Sem vnesite število zgošèenk, ki jih ima vaš vnos [max. 2 znaka]"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1920
      MaxLength       =   25
      TabIndex        =   6
      ToolTipText     =   "Sem vnesite zvrst vnosa [max. 25 znakov)"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   5
      ToolTipText     =   "Sem vnesite ime vnosa [max. 30 znakov]"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   480
      Stretch         =   -1  'True
      ToolTipText     =   "Dvoklik za izbris poti."
      Top             =   1920
      Width           =   915
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Število CDejev:"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Zvrst:"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Ime:"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ID:"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "nov"
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
Image1.Picture = LoadPicture(pot)
napaka:
    End With
End Sub

Private Sub Command2_Click()
    If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then 'preveri èe so polja polna
    MsgBox ("Vnesite vsa polja!")
        Else
        If Text5.Text = "" Then 'ali je slika?
           Text5.Text = "Slike ni"
        Else
            If (Dir(Text5.Text) <> "") Then
                        kontrola
                        y = App.Path & "\Pics\" & Text1.Text & lep.Text & ".jpg"
                        x = Text5.Text
                        FileCopy x, y
                        Text5.Text = y
            Else
                        MsgBox "Napaka v datoteki!", vbCritical, napaka
                        Exit Sub
            End If
        End If
         If nov.Caption = "Dodaj nov film" Then 'doda film
            With film.filmi.Recordset
                .AddNew
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
                
          ElseIf nov.Caption = "Dodaj novo igro" Then 'doda igro
            With Form1.igra.Recordset
                .AddNew
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
                
           ElseIf nov.Caption = "Dodaj nov program" Then 'doda program
            With programi.progi.Recordset
                .AddNew
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
                
            ElseIf nov.Caption = "Dodaj drugo" Then 'doda ostalo
            With ostalo.ostane.Recordset
                .AddNew
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
        Command3_Click 'gre ven
    End If
End Sub

Private Sub Command3_Click()
'izhod
    Unload Me
End Sub

Private Sub Form_Paint()
'prišteje id
    If nov.Caption = "Dodaj nov film" Then 'filmi
        If film.filmi.Recordset.EOF = False And film.filmi.Recordset.BOF = False Then
            film.filmi.Recordset.MoveLast
            Text1.Text = film.filmi.Recordset!id + 1
        Else
            Text1.Text = "1"
        End If
    ElseIf nov.Caption = "Dodaj novo igro" Then 'igre
        If Form1.igra.Recordset.EOF = False And Form1.igra.Recordset.BOF = False Then
            Form1.igra.Recordset.MoveLast
            Text1.Text = Form1.igra.Recordset!id + 1
        Else
            Text1.Text = "1"
        End If
    ElseIf nov.Caption = "Dodaj nov program" Then 'programi
        If programi.progi.Recordset.EOF = False And programi.progi.Recordset.BOF = False Then
            programi.progi.Recordset.MoveLast
            Text1.Text = programi.progi.Recordset!id + 1
        Else
            Text1.Text = "1"
        End If
    ElseIf nov.Caption = "Dodaj drugo" Then 'drugo
        If ostalo.ostane.Recordset.EOF = False And ostalo.ostane.Recordset.BOF = False Then
            ostalo.ostane.Recordset.MoveLast
            Text1.Text = ostalo.ostane.Recordset!id + 1
        Else
            Text1.Text = "1"
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'omogoèi nazaj okno
    If nov.Caption = "Dodaj nov film" Then
        film.Enabled = True
    ElseIf nov.Caption = "Dodaj novo igro" Then
        Form1.Enabled = True
    ElseIf nov.Caption = "Dodaj nov program" Then
        programi.Enabled = True
    ElseIf nov.Caption = "Dodaj drugo" Then
        ostalo.Enabled = True
    End If
End Sub

Private Sub Image1_DblClick()
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

