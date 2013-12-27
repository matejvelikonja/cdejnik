VERSION 5.00
Begin VB.Form nastavitve 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nastavitve"
   ClientHeight    =   6885
   ClientLeft      =   1530
   ClientTop       =   2265
   ClientWidth     =   8715
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   8715
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   5880
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Matej\My Documents\Racunalnik\CDejnik 3.0\Files\cdejnik.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Ostalo"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Matej\My Documents\Racunalnik\CDejnik 3.0\Files\cdejnik.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Programi"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Matej\My Documents\Racunalnik\CDejnik 3.0\Files\cdejnik.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Igre"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Matej\My Documents\Racunalnik\CDejnik 3.0\Files\cdejnik.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Filmi"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Informacije"
      Height          =   2415
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   8415
      Begin VB.Label ura 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   35
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Danes smo:"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Program je bil prviè pognan:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Zadnjiè je bil uporabljen:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Skupaj"
         Height          =   255
         Index           =   7
         Left            =   6360
         TabIndex        =   29
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5520
         TabIndex        =   28
         Top             =   1920
         Width           =   795
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         X1              =   5040
         X2              =   6960
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "ostalih vnosov."
         Height          =   255
         Index           =   6
         Left            =   6360
         TabIndex        =   27
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "programov,"
         Height          =   255
         Index           =   5
         Left            =   6360
         TabIndex        =   26
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "iger,"
         Height          =   255
         Index           =   4
         Left            =   6360
         TabIndex        =   25
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "filmov,"
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5520
         TabIndex        =   23
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5520
         TabIndex        =   22
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5520
         TabIndex        =   21
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5520
         TabIndex        =   20
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "V bazi imate:"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Program ste pognali:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label podatki 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Baza podatkov je velika:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Animacije"
      Height          =   2055
      Left            =   4440
      TabIndex        =   12
      Top             =   120
      Width           =   4095
      Begin VB.OptionButton odlicno 
         Caption         =   "Odlièno"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   960
         Value           =   -1  'True
         Width           =   1000
      End
      Begin VB.OptionButton slabo 
         Caption         =   "Slabo"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   1440
         Width           =   1000
      End
      Begin VB.OptionButton srednje 
         Caption         =   "Srednje"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   1200
         Width           =   1000
      End
      Begin VB.CheckBox panimacija 
         Caption         =   "Igraj animacije"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Izklopi/Vklopi animacije"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Kvaliteta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Varovanje"
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4095
      Begin VB.TextBox geslo1 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   10
         PasswordChar    =   "?"
         TabIndex        =   2
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox geslo 
         BackColor       =   &H80000004&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   10
         PasswordChar    =   "?"
         TabIndex        =   1
         Top             =   840
         Width           =   2295
      End
      Begin VB.CheckBox pgeslo 
         Caption         =   "Uporabi geslo"
         Height          =   375
         Left            =   480
         TabIndex        =   0
         ToolTipText     =   "Izkljuèite/vklopite geslo"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Potrdi geslo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Vpiši geslo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Preklièi"
      Height          =   735
      Left            =   5160
      TabIndex        =   8
      ToolTipText     =   "Ne uporabite sprememb"
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Potrdi"
      Default         =   -1  'True
      Height          =   735
      Left            =   1080
      TabIndex        =   7
      ToolTipText     =   "Potrdite spremembe"
      Top             =   5760
      Width           =   2415
   End
End
Attribute VB_Name = "nastavitve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public baza

Private Sub Command1_Click()
SaveSetting "CDejnik", "Nastavitve", "Geslo", pgeslo.Value 'shranjevanje gesla
If pgeslo.Value = 1 Then
    If geslo.Text = geslo1.Text Then
            If geslo.Text <> "" Then
                SaveSetting "CDejnik", "Nastavitve", "PGeslo", geslo.Text
            Else
                MsgBox "Geslo ne more biti prazen niz znakov", vbExclamation
                Exit Sub
            End If
    Else
        MsgBox "Geslo morate vtipkati obakrat enako!", vbExclamation
        Exit Sub
    End If
End If

SaveSetting "CDejnik", "Nastavitve", "Animacija", panimacija.Value
If panimacija.Value = 1 Then
    If odlicno.Value = True Then
        SaveSetting "CDejnik", "Nastavitve", "Kakovost animacije", "Odlièno"
    ElseIf srednje.Value = True Then
        SaveSetting "CDejnik", "Nastavitve", "Kakovost animacije", "Srednje"
    ElseIf slabo.Value = True Then
        SaveSetting "CDejnik", "Nastavitve", "Kakovost animacije", "Slabo"
    Else
        MsgBox "Unknown error", vbExclamation
        End
    End If
End If

Command2_Click

End Sub

Private Sub pgeslo_Click()
If pgeslo.Value = Checked Then
    geslo.Enabled = True
    geslo1.Enabled = True
    geslo.BackColor = &H80000004
    geslo1.BackColor = &H80000004
ElseIf pgeslo.Value = Unchecked Then
    geslo.Enabled = False
    geslo1.Enabled = False
    geslo.BackColor = &H8000000F
    geslo1.BackColor = &H8000000F
End If
End Sub

Private Sub panimacija_Click()
If panimacija.Value = 1 Then
    If GetSetting("CDejnik", "Nastavitve", "Kakovost animacije", "Odlièno") = "Odlièno" Then
        odlicno.Value = True
    ElseIf GetSetting("CDejnik", "Nastavitve", "Kakovost animacije", "Odlièno") = "Srednje" Then
        srednje.Value = True
    ElseIf GetSetting("CDejnik", "Nastavitve", "Kakovost animacije", "Odlièno") = "Slabo" Then
        slabo.Value = True
    Else
        MsgBox "Unknown error", vbExclamation
        End
    End If
End If

If panimacija.Value = Checked Then
    odlicno.Enabled = True
    srednje.Enabled = True
    slabo.Enabled = True
ElseIf panimacija.Value = Unchecked Then
    odlicno.Enabled = False
    srednje.Enabled = False
    slabo.Enabled = False
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Activate()
Dim gesko As Integer
Dim anka As Integer
gesko = GetSetting("CDejnik", "Nastavitve", "Geslo", "0")
pgeslo.Value = gesko
If pgeslo.Value = 1 Then
    geslo.Text = GetSetting("CDejnik", "Nastavitve", "PGeslo", "")
    geslo1.Text = geslo.Text
End If
pgeslo_Click

anka = GetSetting("CDejnik", "Nastavitve", "Animacija", "1")
panimacija.Value = anka
panimacija_Click

Label4.Caption = GetSetting("CDejnik", "Nastavitve", "Krat", "1") & " krat"
Label10.Caption = GetSetting("CDejnik", "Nastavitve", "Zadnjic", (Date & " ob " & Time))
Label11.Caption = GetSetting("CDejnik", "Nastavitve", "Prvic", (Date & " ob " & Time))
Timer1_Timer
End Sub

Private Sub Form_Load()

Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim stevilkaCD1 As Integer
Dim stevilkaCD2 As Integer
Dim stevilkaCD3 As Integer
Dim stevilkaCD4 As Integer

If Dir(App.Path & "\Files\cdejnik.mdb") <> "" Then 'koliko zasede
        Data1.DatabaseName = App.Path & "\Files\cdejnik.mdb"
        Data1.Refresh
        Data2.DatabaseName = App.Path & "\Files\cdejnik.mdb"
        Data2.Refresh
        Data3.DatabaseName = App.Path & "\Files\cdejnik.mdb"
        Data3.Refresh
        Data4.DatabaseName = App.Path & "\Files\cdejnik.mdb"
        Data4.Refresh
    If (FileLen(App.Path & "\Files\cdejnik.mdb") / 1024) >= 1024 Then
        podatki.Caption = (FileLen(App.Path & "\Files\cdejnik.mdb") / 1024 / 1024) & " MB"
    Else
        podatki.Caption = (FileLen(App.Path & "\Files\cdejnik.mdb") / 1024) & " KB"
    End If
Else
    podatki.Caption = "Podatki niso dosegljivi"
    Command1.Enabled = False
    Exit Sub
End If

If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then 'sešteje filme
Data1.Recordset.MoveLast
Label5.Caption = Data1.Recordset.RecordCount

Data1.Recordset.MoveFirst 'skupaj sešteje vse CDje
    Do Until Data1.Recordset.EOF = True
        stevilkaCD1 = Data1.Recordset!St_CDejev + stevilkaCD1
        Data1.Recordset.MoveNext
    Loop
    Label5.ToolTipText = "To je skupaj " & stevilkaCD1 & " CDjev."
Else
Label5.Caption = "0"
End If

If Data2.Recordset.EOF = False And Data2.Recordset.BOF = False Then 'sešteje igre
Data2.Recordset.MoveLast
Label6.Caption = Data2.Recordset.RecordCount
Data2.Recordset.MoveFirst 'skupaj sešteje vse CDje
    Do Until Data2.Recordset.EOF = True
        stevilkaCD2 = Data2.Recordset!St_CDejev + stevilkaCD2
        Data2.Recordset.MoveNext
    Loop
    Label6.ToolTipText = "To je skupaj " & stevilkaCD2 & " CDjev."
Else
Label6.Caption = "0"
End If

If Data3.Recordset.EOF = False And Data3.Recordset.BOF = False Then 'sešteje progame
Data3.Recordset.MoveLast
Label7.Caption = Data3.Recordset.RecordCount
Data3.Recordset.MoveFirst 'skupaj sešteje vse CDje
    Do Until Data3.Recordset.EOF = True
        stevilkaCD3 = Data3.Recordset!St_CDejev + stevilkaCD3
        Data3.Recordset.MoveNext
    Loop
    Label7.ToolTipText = "To je skupaj " & stevilkaCD3 & " CDjev."
Else
Label7.Caption = "0"
End If

If Data4.Recordset.EOF = False And Data4.Recordset.BOF = False Then 'sešteje ostalo
Data4.Recordset.MoveLast
Label8.Caption = Data4.Recordset.RecordCount
Data4.Recordset.MoveFirst 'skupaj sešteje vse CDje
    Do Until Data4.Recordset.EOF = True
        stevilkaCD4 = Data4.Recordset!St_CDejev + stevilkaCD4
        Data4.Recordset.MoveNext
    Loop
    Label8.ToolTipText = "To je skupaj " & stevilkaCD4 & " CDjev."
Else
Label8.Caption = "0"
End If

a = Label6.Caption
b = Label7.Caption
c = Label8.Caption
d = Label5.Caption
Label9.Caption = a + b + c + d
Label9.ToolTipText = "To je skupaj " & (stevilkaCD1 + stevilkaCD2 + stevilkaCD3 + stevilkaCD4) & " CDjev"
End Sub

Private Sub Timer1_Timer()
ura.Caption = Date & " ob " & Time
End Sub
