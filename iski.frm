VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form iski 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Iskanje"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11835
   Icon            =   "iski.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid tabela 
      Bindings        =   "iski.frx":0442
      Height          =   5415
      Left            =   4200
      OleObjectBlob   =   "iski.frx":0457
      TabIndex        =   11
      Top             =   120
      Width           =   7095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Potrdi"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Sem vpišite iskani niz"
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Išèi"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "Išèi"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Išèi po:"
      Height          =   2655
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   2295
      Begin VB.OptionButton sposojeno 
         Caption         =   "Sposojenih vnosih"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Izpiše vnose, ki jih imate izposojene"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.OptionButton id 
         Caption         =   "IDju"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Išèite po ID številki"
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton sposojevalec 
         Caption         =   "Sposojevalcu"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Išèite kdo ima vaše vnose"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.OptionButton opomba 
         Caption         =   "Opombi"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Išèite po opombah"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton zvrst 
         Caption         =   "Zvrsti"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Išèite po zvrsteh"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton ime 
         Caption         =   "Imenu"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Išèite po imenu"
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Data filmii 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Nazaj"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label preverip 
      Caption         =   "Label2"
      Height          =   495
      Left            =   3360
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Vpiši iskani niz:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "iski"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
Sub Movie()
niz = UCase(Text1.Text)
'išèi
If Text1.Text = "" And sposojeno.Value = False Then
    MsgBox "Izpolnite polje ", vbExclamation
    Exit Sub
Else
    tabela.Visible = True
    
    If ime.Value = True Then 'išèi po imenu
            filmii.RecordSource = "Select Ime, ID, Zvrst From Filmi order by ime"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!ime) Like "*" & niz & "*" Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf id.Value = True Then 'išèi po idju
        filmii.RecordSource = "Select ID, Ime, Zvrst From Filmi order by ID"
        filmii.Refresh
        filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!id) Like niz Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime & " z IDjem " & filmii.Recordset!id, vbInformation
                    Exit Sub
                Else
                    filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf zvrst.Value = True Then 'išèi po zvrsti
            filmii.RecordSource = "Select Zvrst, ID, Ime From Filmi order by Zvrst"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!zvrst) Like "*" & niz & "*" Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime & " zvrsti " & filmii.Recordset!zvrst, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf zvrst.Value = True Then 'išèi po opombi
            filmii.RecordSource = "Select Opombe, ID, Ime, Zvrst From Filmi order by Opombe"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!opombe) Like "*" & niz & "*" Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime & " z opombo " & niz, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf sposojevalec.Value = True Then 'išèi po sposojevalcu
            filmii.RecordSource = "Select Sposojevalec, ID, Ime, Zvrst From Filmi order by Sposojevalec"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!sposojevalec) Like "*" & niz & "*" Then
                    MsgBox "Vnos " & filmii.Recordset!ime & " ima " & filmii.Recordset!sposojevalec, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf sposojeno.Value = True Then 'sposojeno?
            tip = ""
            filmii.RecordSource = "Select ID, Ime, Sposojevalec, Sposojeno As [Sposojeno od:] from filmi Where Sposojevalec <> 'tip' order by Sposojevalec"
            filmii.Refresh
    End If
End If
End Sub

Sub igra()
niz = UCase(Text1.Text)
'išèi
If Text1.Text = "" And sposojeno.Value = False Then
    MsgBox "Izpolnite polje ", vbExclamation
    Exit Sub
Else
    tabela.Visible = True
    
    If ime.Value = True Then 'išèi po imenu
            filmii.RecordSource = "Select Ime, ID, Zvrst From Igre order by ime"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!ime) Like "*" & niz & "*" Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf id.Value = True Then 'išèi po idju
        filmii.RecordSource = "Select ID, Ime, Zvrst From Igre order by ID"
        filmii.Refresh
        filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!id) Like niz Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime & " z IDjem " & filmii.Recordset!id, vbInformation
                    Exit Sub
                Else
                    filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf zvrst.Value = True Then 'išèi po zvrsti
            filmii.RecordSource = "Select Zvrst, ID, Ime From Igre order by Zvrst"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!zvrst) Like "*" & niz & "*" Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime & " zvrsti " & filmii.Recordset!zvrst, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf zvrst.Value = True Then 'išèi po opombi
            filmii.RecordSource = "Select Opombe, ID, Ime, Zvrst From Igre order by Opombe"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!opombe) Like "*" & niz & "*" Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime & " z opombo " & niz, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf sposojevalec.Value = True Then 'išèi po sposojevalcu
            filmii.RecordSource = "Select Sposojevalec, ID, Ime, Zvrst From Igre order by Sposojevalec"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!sposojevalec) Like "*" & niz & "*" Then
                    MsgBox "Vnos " & filmii.Recordset!ime & " ima " & filmii.Recordset!sposojevalec, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf sposojeno.Value = True Then 'sposojeno?
            tip = ""
            filmii.RecordSource = "Select ID, Ime, Sposojevalec, Sposojeno As [Sposojeno od:] from Igre Where Sposojevalec <> 'tip' order by Sposojevalec"
            filmii.Refresh
    End If
End If
End Sub

Sub program()
niz = UCase(Text1.Text)
'išèi
If Text1.Text = "" And sposojeno.Value = False Then
    MsgBox "Izpolnite polje ", vbExclamation
    Exit Sub
Else
    tabela.Visible = True
    
    If ime.Value = True Then 'išèi po imenu
            filmii.RecordSource = "Select Ime, ID, Zvrst From Programi order by ime"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!ime) Like "*" & niz & "*" Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf id.Value = True Then 'išèi po idju
        filmii.RecordSource = "Select ID, Ime, Zvrst From Programi order by ID"
        filmii.Refresh
        filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!id) Like niz Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime & " z IDjem " & filmii.Recordset!id, vbInformation
                    Exit Sub
                Else
                    filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf zvrst.Value = True Then 'išèi po zvrsti
            filmii.RecordSource = "Select Zvrst, ID, Ime From Programi order by Zvrst"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!zvrst) Like "*" & niz & "*" Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime & " zvrsti " & filmii.Recordset!zvrst, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf zvrst.Value = True Then 'išèi po opombi
            filmii.RecordSource = "Select Opombe, ID, Ime, Zvrst From Programi order by Opombe"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!opombe) Like "*" & niz & "*" Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime & " z opombo " & niz, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf sposojevalec.Value = True Then 'išèi po sposojevalcu
            filmii.RecordSource = "Select Sposojevalec, ID, Ime, Zvrst From Programi order by Sposojevalec"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!sposojevalec) Like "*" & niz & "*" Then
                    MsgBox "Vnos " & filmii.Recordset!ime & " ima " & filmii.Recordset!sposojevalec, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf sposojeno.Value = True Then 'sposojeno?
            tip = ""
            filmii.RecordSource = "Select ID, Ime, Sposojevalec, Sposojeno As [Sposojeno od:] from Programi Where Sposojevalec <> 'tip' order by Sposojevalec"
            filmii.Refresh
    End If
End If
End Sub
Sub drugo()
niz = UCase(Text1.Text)
'išèi
If Text1.Text = "" And sposojeno.Value = False Then
    MsgBox "Izpolnite polje ", vbExclamation
    Exit Sub
Else
    tabela.Visible = True
    
    If ime.Value = True Then 'išèi po imenu
            filmii.RecordSource = "Select Ime, ID, Zvrst From Ostalo order by ime"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!ime) Like "*" & niz & "*" Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf id.Value = True Then 'išèi po idju
        filmii.RecordSource = "Select ID, Ime, Zvrst From Ostalo order by ID"
        filmii.Refresh
        filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!id) Like niz Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime & " z IDjem " & filmii.Recordset!id, vbInformation
                    Exit Sub
                Else
                    filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf zvrst.Value = True Then 'išèi po zvrsti
            filmii.RecordSource = "Select Zvrst, ID, Ime From Ostalo order by Zvrst"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!zvrst) Like "*" & niz & "*" Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime & " zvrsti " & filmii.Recordset!zvrst, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf zvrst.Value = True Then 'išèi po opombi
            filmii.RecordSource = "Select Opombe, ID, Ime, Zvrst From Ostalo order by Opombe"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!opombe) Like "*" & niz & "*" Then
                    MsgBox "Najden vnos " & filmii.Recordset!ime & " z opombo " & niz, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf sposojevalec.Value = True Then 'išèi po sposojevalcu
            filmii.RecordSource = "Select Sposojevalec, ID, Ime, Zvrst From Ostalo order by Sposojevalec"
            filmii.Refresh
            filmii.Recordset.MoveFirst
            Do Until filmii.Recordset.EOF = True
                If UCase(filmii.Recordset!sposojevalec) Like "*" & niz & "*" Then
                    MsgBox "Vnos " & filmii.Recordset!ime & " ima " & filmii.Recordset!sposojevalec, vbInformation
                    Exit Sub
                Else
                   filmii.Recordset.MoveNext
                End If
            Loop
                nenajdem
                
    ElseIf sposojeno.Value = True Then 'sposojeno?
            tip = ""
            filmii.RecordSource = "Select ID, Ime, Sposojevalec, Sposojeno As [Sposojeno od:] from Ostalo Where Sposojevalec <> 'tip' order by Sposojevalec"
            filmii.Refresh
    End If
End If
End Sub

Private Sub Command2_Click()
'iskanje in sortiranje
If preverip.Caption = "film" Then
    Movie
ElseIf preverip.Caption = "igra" Then
    igra
ElseIf preverip.Caption = "progi" Then
    program
ElseIf preverip.Caption = "ostanek" Then
    drugo
End If
Command5.Enabled = True
End Sub
Private Sub Command5_Click()
'potrdi
If filmii.Recordset.EOF = False And filmii.Recordset.BOF = False Then
    If preverip.Caption = "film" Then
        Do Until film.filmi.Recordset.EOF = True
            If film.filmi.Recordset!id = filmii.Recordset!id Then
                Exit Do
            Else
                film.filmi.Recordset.MoveNext
            End If
        Loop
    ElseIf preverip.Caption = "progi" Then
        Do Until programi.progi.Recordset.EOF = True
            If programi.progi.Recordset!id = filmii.Recordset!id Then
                Exit Do
            Else
                programi.progi.Recordset.MoveNext
            End If
        Loop
    ElseIf preverip.Caption = "igra" Then
        Do Until Form1.igra.Recordset.EOF = True
            If Form1.igra.Recordset!id = filmii.Recordset!id Then
                Exit Do
            Else
                Form1.igra.Recordset.MoveNext
            End If
        Loop
    ElseIf preverip.Caption = "ostanek" Then
        Do Until ostalo.ostane.Recordset.EOF = True
            If ostalo.ostane.Recordset!id = filmii.Recordset!id Then
                Exit Do
            Else
                ostalo.ostane.Recordset.MoveNext
            End If
        Loop
    End If
    Command1_Click
End If
End Sub

Private Sub Form_Load()
tabela.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'omogoèi nazaj okno
    If preverip.Caption = "igra" Then
        Form1.Enabled = True
    ElseIf preverip.Caption = "film" Then
        film.Enabled = True
    ElseIf preverip.Caption = "progi" Then
        programi.Enabled = True
    ElseIf preverip.Caption = "ostanek" Then
        ostalo.Enabled = True
    End If
End Sub

Private Sub id_Click()
Text1.Enabled = True
Text1.Text = ""
End Sub

Private Sub ime_Click()
id_Click
End Sub

Private Sub opomba_Click()
id_Click
End Sub

Private Sub sposojeno_Click()
Text1.Enabled = False
End Sub

Private Sub sposojevalec_Click()
id_Click
End Sub

Private Sub Text1_Change()
 Static vnos
If id.Value = True Then
    If Not IsNumeric(Text1.Text) Then
        Text1.Text = vnos
        Beep
    Else
        vnos = Text1.Text
    End If
ElseIf zvrst.Value = True Then
    If IsNumeric(Text1.Text) Then
        Text1.Text = vnos
        Beep
    Else
        vnos = Text1.Text
    End If
ElseIf sposojevalec.Value = True Then
    If IsNumeric(Text1.Text) Then
        Text1.Text = vnos
        Beep
    Else
        vnos = Text1.Text
    End If
End If
End Sub

Sub nenajdem()
MsgBox "Nobenega vnosa ni bilo najdenega.", vbInformation
End Sub

Private Sub zvrst_Click()
id_Click
End Sub
