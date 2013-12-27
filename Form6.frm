VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form natisni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Natisni"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6510
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog save 
      Left            =   5880
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Shrani v datoteko"
      Filter          =   "Internetna stran (*.html)|*.html|"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Natisni"
      Default         =   -1  'True
      Height          =   615
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Natisni"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CheckBox datoteka 
      Caption         =   "Izvozi v datoteko"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      ToolTipText     =   "Izvozite seznam v html"
      Top             =   960
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Preklièi"
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label krompir 
      Height          =   615
      Left            =   6000
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label x 
      Caption         =   "Label1"
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label gonilnik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      ToolTipText     =   "Uporabljen gonilnik"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Gonilnik:"
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   8
      Top             =   600
      Width           =   735
   End
   Begin VB.Label kvaliteta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      ToolTipText     =   "Kvaliteta tiskanja v pikah na palec"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label vrata 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      ToolTipText     =   "Tiskalnik je prikljuèen na"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label tiskalnik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "Ime tiskalnika"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Kvaliteta:"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Na portu:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Tiskalnik:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "natisni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim vendar
ime
vendar = krompir.Caption
If datoteka.Value = Unchecked Then 'tisklanik
    With Printer
        .FontSize = "8"
        .Font = "Arial"
        Printer.Print Tab(146); "Dne: " & Date
        .FontBold = True
        Printer.Print Tab(140); "CDejnik " & App.Major & "." & App.Minor
        Printer.Print
        .Font = "Courier New"
        .FontBold = True
        .FontSize = "20"
        Printer.Print Tab(19); "-- "; UCase(vendar); " --"
        Printer.Line (500, 1700)-(14000, 1700)
        Printer.Line (500, 1705)-(14000, 1705)
        Printer.Line (500, 1710)-(14000, 1710)
        .Font = "Impact"
        .FontSize = "14"
        .FontItalic = True
        .FontBold = False
        Printer.Print Tab(5); "ID"; Tab(15); "Ime"; Tab(45); "Število CDejev"; Tab(69); "Zvrst"
        Printer.Line (500, 2450)-(14000, 2450)
        Data1.Recordset.MoveFirst
        .Font = "Times New Roman"
        .FontBold = False
        .FontItalic = False
        .FontSize = "13"
        Do Until Data1.Recordset.EOF = True
            Printer.Print Tab(5); Data1.Recordset!id; Tab(15); Data1.Recordset!ime; Tab(55); Data1.Recordset!St_CDejev; Tab(76); Data1.Recordset!zvrst
            Data1.Recordset.MoveNext
        Loop
        .EndDoc
    End With
    
ElseIf datoteka.Value = Checked Then 'v datoteko izvozi
            With save
            .Flags = cd10fnfilemustexist
            .FileName = ""
            .DialogTitle = "Shrani v datoteko"
            .CancelError = True
        On Error GoTo napaka
            .ShowSave
            pot = .FileName
            Open pot For Output As #1
                Print #1, "<html>"
                Print #1, "<head>"
                Print #1, "<title>Seznam ("; vendar; ")</title>"
                Print #1, "</head>"
                Print #1, "<body bgcolor=""#ffffff"" marginleft="; 0; ">"
                Print #1, "<h1 align=""center"">"; UCase(vendar); "</h1>"
                Print #1,
                Print #1,
                Print #1, "<table border=""0"" align=""center"">"
                Print #1, "<tr>"
                Print #1, "<td><b>ID</td>"
                Print #1, "<td width=""20""></td>"
                Print #1, "<td><b>IME</td>"
                Print #1, "<td width=""20""></td>"
                Print #1, "<td><b>ŠTEVILO CDEJEV</td>"
                Print #1, "<td width=""20""></td>"
                Print #1, "<td><b>ZVRST</td>"
                Print #1, "</tr>"
                    Do Until Data1.Recordset.EOF = True
                        Print #1, "<tr>"
                        Print #1, "<td>"; Data1.Recordset!id; "</td>"
                        Print #1, "<td width=""20""></td>"
                        Print #1, "<td>"; Data1.Recordset!ime; "</td>"
                        Print #1, "<td width=""20""></td>"
                        Print #1, "<td align=""center"">"; Data1.Recordset!St_CDejev; "</td>"
                        Print #1, "<td width=""20""></td>"
                        Print #1, "<td>"; Data1.Recordset!zvrst; "</td>"
                        Print #1, "</tr>"
                        Print #1,
                        Data1.Recordset.MoveNext
                    Loop
                Print #1, "</table>"
                Print #1,
                Print #1,
                Print #1,
                Print #1,
                Print #1, "<h4 align=""right"">CDejnik " & App.Major & "." & App.Minor; "</h4>"
                Print #1, "</body>"
                Print #1, "</html>"
                Print #1,
                Print #1,
                Print #1,
                Print #1,
                Close #1
napaka:
            End With
            
End If
MsgBox "Konèano", vbInformation, "Obvestilo"
Command2_Click
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
tiskalnik.Caption = Printer.DeviceName
vrata.Caption = Printer.Port
kvaliteta.Caption = Printer.PrintQuality & " dpi"
gonilnik.Caption = Printer.DriverName
End Sub

Private Sub Form_Click()
Form_Activate
End Sub

Private Sub Form_Unload(Cancel As Integer)
'omogoèi nazaj okno
    If x.Caption = "film" Then
        film.Enabled = True
    ElseIf x.Caption = "igra" Then
        Form1.Enabled = True
    ElseIf x.Caption = "progi" Then
        programi.Enabled = True
    ElseIf x.Caption = "ostanek" Then
        ostalo.Enabled = True
    End If
End Sub

Sub ime()
With krompir
  If x.Caption = "film" Then
        .Caption = "Filmi"
    ElseIf x.Caption = "igra" Then
        .Caption = "Igre"
    ElseIf x.Caption = "progi" Then
        .Caption = "Programi"
    ElseIf x.Caption = "ostanek" Then
        .Caption = "Ostalo"
    End If
End With
End Sub
