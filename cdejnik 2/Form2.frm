VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seznam CDejev"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10380
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBGrid1 
      Align           =   1  'Align Top
      Bindings        =   "Form2.frx":030A
      Height          =   4095
      Left            =   0
      OleObjectBlob   =   "Form2.frx":031A
      TabIndex        =   0
      Top             =   0
      Width           =   10380
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Najveèja ocena"
      Height          =   855
      Left            =   4560
      Picture         =   "Form2.frx":0CCC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "seznam.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "igre"
      Top             =   4080
      Visible         =   0   'False
      Width           =   9855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Izhod"
      Height          =   855
      Left            =   6000
      Picture         =   "Form2.frx":10AB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Izbriši vnos"
      Height          =   855
      Left            =   3120
      Picture         =   "Form2.frx":13B5
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Popravi vnos"
      Height          =   855
      Left            =   1680
      Picture         =   "Form2.frx":16BF
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dodaj vnos"
      Height          =   855
      Left            =   240
      Picture         =   "Form2.frx":19C9
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Show MODAL
End Sub

Private Sub Command2_Click()
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
MsgBox ("Izberite polje!!!!!!!!",vbExclamation,"Napaka")
Else
Form4.Show MODAL
End If
End Sub

Private Sub Command3_Click()
With Data1.Recordset
.Delete
.MoveFirst
End With
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
x = 0
y = 0
Data1.Recordset.MoveFirst
Do Until Data1.Recordset.EOF = True
x = Data1.Recordset!Ocena
If x > y Then
y = x
End If
Data1.Recordset.MoveNext
Loop
 MsgBox (y)
End Sub
