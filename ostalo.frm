VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ostalo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDejnik 3.0 (Ostalo)"
   ClientHeight    =   9540
   ClientLeft      =   1695
   ClientTop       =   1440
   ClientWidth     =   13335
   Icon            =   "ostalo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   13335
   Begin MSDBGrid.DBGrid seznamigre 
      Bindings        =   "ostalo.frx":0442
      Height          =   6855
      Left            =   5760
      OleObjectBlob   =   "ostalo.frx":0457
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
   Begin VB.TextBox lep 
      Height          =   285
      Left            =   360
      TabIndex        =   33
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informacije"
      Height          =   3855
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   5415
      Begin VB.TextBox cdigre 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "St_CDejev"
         DataSource      =   "ostane"
         Height          =   255
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "cd"
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox zvrstigre 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "Zvrst"
         DataSource      =   "ostane"
         Height          =   255
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "zvrst"
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox imeigre 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "Ime"
         DataSource      =   "ostane"
         Height          =   255
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "ime"
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox IDigre 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "ID"
         DataSource      =   "ostane"
         Height          =   255
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "id"
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox opombeigre 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "Opombe"
         DataSource      =   "ostane"
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   19
         Text            =   "ostalo.frx":0E2D
         ToolTipText     =   "Sem vnesite opombe"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Image slika 
         Height          =   2490
         Left            =   2040
         Stretch         =   -1  'True
         ToolTipText     =   "Slika"
         Top             =   1200
         Width           =   3120
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Slika:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   29
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Število CDejev:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Zvrst:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   15
         Left            =   720
         TabIndex        =   27
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ime:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   14
         Left            =   840
         TabIndex        =   26
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ID:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   13
         Left            =   720
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Opombe:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Posoja"
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   5415
      Begin VB.TextBox dnis 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CommandButton vrni 
         Caption         =   "Vrni"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2520
         TabIndex        =   14
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton sposodi 
         Caption         =   "Sposodi"
         Enabled         =   0   'False
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox spodne 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "Sposojeno"
         DataSource      =   "ostane"
         Height          =   255
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "21/02/2002"
         Top             =   960
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox sposojevalec 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "Sposojevalec"
         DataSource      =   "ostane"
         Height          =   255
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "kdo?"
         Top             =   720
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox sposojeno 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "da/ne"
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sposojeno:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   35
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sposojeno od dne:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sposojevalec:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sposojeno:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton tiskaj 
      Caption         =   "Natisni"
      Height          =   1095
      Left            =   6720
      TabIndex        =   8
      ToolTipText     =   "Natisni seznam"
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton izbrisi 
      Caption         =   "Izbriši vnos"
      Height          =   1095
      Left            =   4560
      TabIndex        =   7
      ToolTipText     =   "Izbriši trenutni vnos"
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton popravig 
      Caption         =   "Popravi vnos"
      Height          =   1095
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "Popravi že obstojeèi vnos"
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton vnos 
      Caption         =   "Nov vnos"
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Dodaj nov vnos"
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Konèaj"
      Height          =   1095
      Left            =   11040
      TabIndex        =   4
      ToolTipText     =   "Konèaj"
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton iskaj 
      Caption         =   "Natanènejše iskanje"
      Height          =   1095
      Left            =   8880
      TabIndex        =   3
      ToolTipText     =   "Išèi"
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   7320
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4800
      Top             =   7320
   End
   Begin VB.Timer Timer3 
      Interval        =   105
      Left            =   4200
      Top             =   7320
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   6360
      Top             =   7320
   End
   Begin VB.CommandButton isci 
      Caption         =   "Išèi"
      Default         =   -1  'True
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      ToolTipText     =   "Sem vpišite iskalni niz"
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Data ostane 
      Caption         =   "ostalo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Ostalo"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar bar 
      Height          =   300
      Left            =   9360
      TabIndex        =   30
      ToolTipText     =   "Pregled baze"
      Top             =   9240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar vrstica 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   31
      Top             =   9240
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2647
            MinWidth        =   2647
            Object.ToolTipText     =   "Število vseh vnosov, ki jih imate v zbirki."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3000
            MinWidth        =   3000
            Object.ToolTipText     =   "Trenutni zapis"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2293
            MinWidth        =   2293
            Object.ToolTipText     =   "Trenutna ura"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "Današnji datum"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Išèi:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   32
      Top             =   6720
      Width           =   615
   End
End
Attribute VB_Name = "ostalo"
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
Timer1.Enabled = True
End Sub


Sub timer()
Timer2.Enabled = True
Timer3.Enabled = True
End Sub


Private Sub status()
If ostane.Recordset.EOF = False And ostane.Recordset.BOF = False Then
    trenutna = ostane.Recordset.AbsolutePosition

    ostane.Recordset.MoveLast
    With vrstica
        .Panels(1).Text = "Št. zapisov: " & ostane.Recordset.RecordCount
        .Panels(2).Text = "Ste na zapisu št.: " & (trenutna + 1)
        .Panels(3).Text = "Ura je " & Time
        .Panels(4).Text = "Danes smo " & Date
    End With
    
    ostane.Recordset.AbsolutePosition = trenutna
End If
End Sub

Private Sub Form_Activate()
status
IDigre_Change
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ostane.Recordset.EOF = True Then
        If KeyCode = vbKeyUp Then ostane.Recordset.MovePrevious
        If KeyCode = vbKeyDown Then ostane.Recordset.MoveFirst
        
    ElseIf ostane.Recordset.BOF = True Then
        If KeyCode = vbKeyDown Then ostane.Recordset.MoveNext
        If KeyCode = vbKeyUp Then ostane.Recordset.MoveLast
        
    ElseIf ostane.Recordset.EOF = False And ostane.Recordset.BOF = False Then
        If KeyCode = vbKeyUp Then
            ostane.Recordset.MovePrevious
        ElseIf KeyCode = vbKeyDown Then
            ostane.Recordset.MoveNext
        End If
    End If
End Sub

Private Sub Form_Load()
If (Dir(App.Path & "\Files\cdejnik.mdb") <> "") Then
           With ostalo
            .ostane.DatabaseName = App.Path & "\Files\cdejnik.mdb"
            .ostane.Refresh
           End With
        Else
            MsgBox "Datoteke z zbirko ni mogoèe najti.", vbExclamation
            Unload Me
        End If
sposojeno.Text = ""
timer
End Sub

Private Sub IDigre_Change()
' sposojeno da ne
     If ostane.Recordset.EOF = False And ostane.Recordset.BOF = False Then
            If ostane.Recordset!sposojevalec <> "" Then
                sposojeno.Text = "Da"
                vrni.Enabled = True
                sposodi.Enabled = False
                
                dni = DateDiff("d", (ostane.Recordset!sposojeno), Date) 'število sposojenih dni
                If dni = 1 Then
                    dnis.Text = dni & " dan"
                ElseIf dni = 2 Then
                    dnis.Text = dni & " dneva"
                Else
                    dnis.Text = dni & " dni"
                End If
                
            Else
                sposojeno.Text = "Ne"
                vrni.Enabled = False
                sposodi.Enabled = True
            End If
           
        timer 'gre na timer
        vrstica.Panels(2).Text = "Ste na zapisu št.: " & (ostane.Recordset.AbsolutePosition + 1)
        bar.Value = ((ostane.Recordset.AbsolutePosition + 1) / ostane.Recordset.RecordCount) * 100
     Else
        vrstica.Panels(2).Text = "Ste na zapisu št.: 0"
        bar.Value = 0
     End If
End Sub

Private Sub isci_Click()
If ostane.Recordset.EOF = False And ostane.Recordset.BOF = False Then
'išèi
    x = ostane.Recordset.AbsolutePosition
    If Text1.Text = "" Then
        MsgBox "Izpolnite polje ", vbExclamation
        Exit Sub
    End If
    
        niz = UCase(Text1.Text)
        ostane.Recordset.MoveFirst
        Do Until ostane.Recordset.EOF = True
            If UCase(ostane.Recordset!ime) Like "*" & niz & "*" Then
                MsgBox "Najden vnos " & ostane.Recordset!ime, vbInformation
                Exit Sub
            Else
                ostane.Recordset.MoveNext
            End If
        Loop
        
        If MsgBox("Nobenega vnosa ni bilo najdenega.", vbInformation) = vbOK Then
            ostane.Recordset.AbsolutePosition = x
            Exit Sub
        End If
Else
    MsgBox "Zbirka je prazna", vbCritical, "Opozorilo"
End If
End Sub

Private Sub iskaj_Click()
If ostane.Recordset.EOF = False And ostane.Recordset.BOF = False Then
    iski.Show modal
    iski.filmii.DatabaseName = App.Path & "\Files\cdejnik.mdb"
    iski.filmii.Refresh
    iski.preverip.Caption = "ostanek"
    Me.Enabled = False
Else
    MsgBox "Zbirka je prazna", vbCritical
End If
End Sub

Private Sub izbrisi_Click()
    If ostane.Recordset.EOF = True Or ostane.Recordset.BOF = True Then
            MsgBox "Izberite vnos!", vbCritical
    Else
            If MsgBox("Ste preprièani, da želite izbrisati " & ostane.Recordset!ime & "?", vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then
                pot = ostane.Recordset!slika
                If (Dir(pot) <> "") Then
                    Kill (pot)
                End If
                With ostane.Recordset
                .Delete
                .MoveLast
                End With
            Else
                Exit Sub
            End If
            timer
            status
    End If
End Sub

Private Sub popravig_Click()
    If ostane.Recordset.EOF = True Or ostane.Recordset.BOF = True Then
        MsgBox "Izberite vnos!", vbCritical
    Else
        popravi.Show modal
        popravi.Caption = "Popravi vnos: " & ostane.Recordset!ime
        popravi.preverip.Caption = "ostanek"
        Me.Enabled = False
    End If
End Sub


Private Sub sposodi_Click()
   ChDir App.Path
   If ostane.Recordset.EOF = True Or ostane.Recordset.BOF = True Then
        MsgBox ("Izberite vnos!")
    Else
        posodio.Show modal
        posodio.Caption = "Posodi: " & ostane.Recordset!ime
        posodio.preveri.Caption = "ostanek"
        Me.Enabled = False
    End If
End Sub

Private Sub sposojeno_Change()
 If sposojeno.Text = "Da" Then
        sposojevalec.Visible = True
        spodne.Visible = True
        dnis.Visible = True
        Label1(7).Visible = True
        Label1(8).Visible = True
        Label1(1).Visible = True
    Else
        sposojevalec.Visible = False
        spodne.Visible = False
        dnis.Visible = False
        Label1(7).Visible = False
        Label1(8).Visible = False
        Label1(1).Visible = False
    End If
End Sub


Private Sub Timer1_Timer()
If Me.Height <> 510 Or Me.Width <> 1005 Then
    Me.Height = Me.Height - 505
    Me.Width = Me.Width - 505
Else
    Unload Me
End If
End Sub

Private Sub Timer2_Timer()
    If ostane.Recordset.EOF = False And ostane.Recordset.BOF = False Then 'veljaven zapis
    
        slikca = ostane.Recordset!slika 'definicija
        nislike = App.Path & "\Pics\nopic.jpg"
        
            If slikca = "Slike ni" Then 'slike je?
                    If (Dir(nislike) <> "") Then
                        slika.Picture = LoadPicture(nislike) 'ni slike
                    End If
            Else
                    If (Dir(slikca) <> "") Then 'obstaja datoteka?
                        slika.Picture = LoadPicture(slikca) 'slika je
                    Else
                    lep.Text = ostane.Recordset!ime
                    kontrola
                    y = App.Path & "\Pics\" & ostane.Recordset!id & lep.Text & ".jpg"
                        If Dir(y) <> "" Then
                            With ostane.Recordset
                            .Edit
                                !slika = y
                            .Update
                            .Bookmark = .LastModified
                            End With
                            slika.Picture = LoadPicture(ostane.Recordset!slika)
                         Else
                            MsgBox "Pot do slike je neveljavna", vbExclamation
                            Timer2.Enabled = False
                         End If
                    End If
            End If
        
    Else
        If ostane.Recordset.EOF = False And ostane.Recordset.BOF = False Then
            If (Dir(nislike) <> "") Then
                slika.Picture = LoadPicture(nislike)
            End If
        End If
    
    End If
End Sub

Private Sub Timer3_Timer()
Timer2.Enabled = False
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
With vrstica
.Panels(3).Text = "Ura je " & Time
.Panels(4).Text = "Danes smo " & Date
End With
End Sub

Private Sub tiskaj_Click()
If ostane.Recordset.EOF = True And ostane.Recordset.BOF = True Then
    MsgBox "Zbirka je prazna", vbExclamation, "Tiskanje nemogoèe"
Else
    If Err.Number Then
        MsgBox Err.Description, vbCritical
        Exit Sub
    Else
            With natisni
                .Show modal
                .Data1.DatabaseName = ostane.DatabaseName
                .Data1.Refresh
                .Data1.RecordSource = "Select Ime, ID, Zvrst , St_CDejev From Ostalo order by ime"
                .Data1.Refresh
                .x.Caption = "ostanek"
            End With
            Me.Enabled = False
    End If
End If
End Sub

Private Sub vnos_Click()
    'prikaže okno vnos
    nov.Show modal
    nov.Caption = "Dodaj drugo"
    Me.Enabled = False
End Sub

Private Sub vrni_Click()
x = ostane.Recordset!ime
With ostane.Recordset
        .Edit
            !sposojevalec = Null
            !sposojeno = Null
        .Update
        .Bookmark = .LastModified
    End With
IDigre_Change
MsgBox "Uspešno vrnjen " & x & ".", vbInformation
End Sub
Private Sub Form_Unload(Cancel As Integer)
   With main
    .Command1.Enabled = True
    .Command2.Enabled = True
    .Command3.Enabled = True
    .Command4.Enabled = True
   End With
End Sub

