VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seznam programov"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10380
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form5.frx":0000
      Height          =   3735
      Left            =   120
      OleObjectBlob   =   "Form5.frx":0010
      TabIndex        =   0
      Top             =   240
      Width           =   9975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dodaj vnos"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Popravi vnos"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Izbriši vnos"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Izhod"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\AProjekti\Program\seznam.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "igre"
      Top             =   4320
      Visible         =   0   'False
      Width           =   9855
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Unload Me
Form1.Visible = True
End Sub
