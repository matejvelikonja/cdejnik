VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form zacetno 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5880
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   18000
      Left            =   4200
      Top             =   4800
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Preskoèi"
      Default         =   -1  'True
      Height          =   375
      Left            =   7805
      TabIndex        =   1
      ToolTipText     =   "Preskoèi animacijo."
      Top             =   3880
      Width           =   975
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash uvod 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10095
      _cx             =   4212110
      _cy             =   4204702
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
End
Attribute VB_Name = "zacetno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
main.Show modal
Unload Me
End Sub


Private Sub Form_Activate()
If GetSetting("CDejnik", "Nastavitve", "Animacija", "1") = 0 Then
    Command1_Click
Else
    If GetSetting("CDejnik", "Nastavitve", "Kakovost animacije", "Odlièno") = "Odlièno" Then
        uvod.Quality2 = "High"
    ElseIf GetSetting("CDejnik", "Nastavitve", "Kakovost animacije", "Odlièno") = "Srednje" Then
        uvod.Quality2 = "Medium"
    ElseIf GetSetting("CDejnik", "Nastavitve", "Kakovost animacije", "Odlièno") = "Slabo" Then
        uvod.Quality2 = "Low"
    End If
    Form_Load
End If

Dim krat As Integer
krat = GetSetting("CDejnik", "Nastavitve", "Krat", "0")
SaveSetting "CDejnik", "Nastavitve", "Krat", krat + 1
If GetSetting("CDejnik", "Nastavitve", "Prvic", (Date & " ob " & Time)) = (Date & " ob " & Time) Then
    SaveSetting "CDejnik", "Nastavitve", "Prvic", (Date & " ob " & Time)
End If
End Sub

Private Sub Form_Load()
If (Dir(App.Path & "\uvod.cd3") <> "") Then
    uvod.Movie = App.Path & "\uvod.cd3"
    uvod.Playing = True
ElseIf (Dir(App.Path & "\Files\uvod.swf") <> "") Then
    uvod.Movie = App.Path & "\Files\uvod.swf"
    uvod.Playing = True
Else
    Timer1.Interval = 100
End If
End Sub

Private Sub Timer1_Timer()
Command1_Click
End Sub

