VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form posodio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Posodi"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "posodi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   240
      Top             =   1440
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Imena"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Dodaj..."
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "Dodaj v seznam novega sposojevalca"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Preklièi"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      ToolTipText     =   "Nesposodi"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Posodi"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Sposodi"
      Top             =   1920
      Width           =   1695
   End
   Begin MSMask.MaskEdBox datum 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Vpišite datum, katerega ste posodili."
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##.##.####"
      PromptChar      =   " "
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "posodi.frx":0442
      Left            =   1800
      List            =   "posodi.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo1"
      ToolTipText     =   "Izberite sposojevalca"
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "DODANO !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label preveri 
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sposojeno dne:"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sposojevalec:"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "posodio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'preveri èe so polja vnesena
    If Combo1.Text = "" Or Not IsDate(datum.FormattedText) Then
        MsgBox ("Vnesite vsa polja!")
    Else
    
    If preveri.Caption = "film" Then 'filmi
        With film.filmi.Recordset
            .Edit
            'posodi
                !sposojevalec = Combo1.Text
                !sposojeno = datum.FormattedText
            .Update
            .Bookmark = .LastModified
        End With
        'isto kot v film idigre_Change
        film.sposojeno.Text = "Da"
        film.vrni.Enabled = True
        film.sposodi.Enabled = False
    ElseIf preveri.Caption = "igra" Then 'igre
        With Form1.igra.Recordset
            .Edit
            'posodi
                !sposojevalec = Combo1.Text
                !sposojeno = datum.FormattedText
            .Update
            .Bookmark = .LastModified
        End With
        'isto kot v film idigre_Change
        Form1.sposojeno.Text = "Da"
        Form1.vrni.Enabled = True
        Form1.sposodi.Enabled = False
    ElseIf preveri.Caption = "progi" Then 'programi
        With programi.progi.Recordset
            .Edit
            'posodi
                !sposojevalec = Combo1.Text
                !sposojeno = datum.FormattedText
            .Update
            .Bookmark = .LastModified
        End With
        'isto kot v film idigre_Change
        programi.sposojeno.Text = "Da"
        programi.vrni.Enabled = True
        programi.sposodi.Enabled = False
    ElseIf preveri.Caption = "ostanek" Then 'ostanek
        With ostalo.ostane.Recordset
            .Edit
            'posodi
                !sposojevalec = Combo1.Text
                !sposojeno = datum.FormattedText
            .Update
            .Bookmark = .LastModified
        End With
        'isto kot v film idigre_Change
        ostalo.sposojeno.Text = "Da"
        ostalo.vrni.Enabled = True
        ostalo.sposodi.Enabled = False
    End If
    Command2_Click
    End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If (Dir(App.Path & "\Files\users.mdb") <> "") Then
    
        With Data1.Recordset
     If .EOF = False And .BOF = False Then
            .MoveFirst
            
            Do Until .EOF = True 'preveri èe obstaja
                    If Combo1.Text = Data1.Recordset!ime Then
                    MsgBox "To osebo že imate", vbExclamation, "Obvestilo"
                    Exit Sub
                Else
                    .MoveNext
                End If
            Loop
      End If
      
        If .EOF = False And .BOF = False Then
            .MoveLast
        End If
        .AddNew
          !ime = Combo1.Text
        .Update
        .Bookmark = .LastModified
    End With
    Label3.Visible = True
    Timer1.Enabled = True
Else
    MsgBox "Datoteka z osebami ne obstaja!", vbExclamation, "Napaka"
    Exit Sub
End If
End Sub

Private Sub nalozi()
With Data1
    If .Recordset.EOF = True Or .Recordset.BOF = True Then
         Combo1.Text = ""
    Else
        .Recordset.MoveFirst
        Do Until .Recordset.EOF = True
            Combo1.AddItem .Recordset!ime
            .Recordset.MoveNext
        Loop
        .Recordset.MoveFirst
        Combo1.Text = .Recordset!ime
    End If
End With
End Sub


Private Sub Form_Load()
If (Dir(App.Path & "\Files\users.mdb") <> "") Then
    Data1.DatabaseName = App.Path & "\Files\users.mdb"
    Data1.Refresh
    nalozi
Else
    MsgBox "Datoteke ni mogoèe najti!", vbExclamation
    Command3.Enabled = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'izhod
    If preveri.Caption = "film" Then 'omogoèi obrazec
    film.Enabled = True
    ElseIf preveri.Caption = "igra" Then
    Form1.Enabled = True
    ElseIf preveri.Caption = "progi" Then
    programi.Enabled = True
    ElseIf preveri.Caption = "ostanek" Then
    ostalo.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
Label3.Visible = False
Timer1.Enabled = False
End Sub
