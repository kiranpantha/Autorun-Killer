VERSION 5.00
Begin VB.Form home 
   BorderStyle     =   0  'None
   Caption         =   "Autrun Killer[Kiransoft Nepal]"
   ClientHeight    =   5010
   ClientLeft      =   5400
   ClientTop       =   2475
   ClientWidth     =   5400
   Icon            =   "home.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "home.frx":0442
   ScaleHeight     =   5010
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer threat 
      Left            =   4800
      Top             =   1560
   End
   Begin VB.Timer text 
      Interval        =   100
      Left            =   4800
      Top             =   1920
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   3480
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7560
      Top             =   1200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   5
      X1              =   0
      X2              =   5400
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "No manual scan needed for autorun killer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   5
      X1              =   0
      X2              =   5400
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      FillColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   0
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label autolist 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No autorun files have been detected in any of the drive in your computer's harddisk and removeable disks "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   5175
   End
   Begin VB.Label movem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"home.frx":3F67E
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dlist(26), a, movet

Private Sub Form_Load()
tray.stat = "Running Autorun killer"
End Sub

Private Sub Form_Unload(Cancel As Integer)
tray.stat = "Autorun killer is still running Hidden"
Cancel = 0
End Sub

Private Sub text_Timer()
movet = "Autorun Kiiller is continiously protecting your System from the external autostart files created by viruses and also helps to prevent your system from external autostart files."
movem = Left$(movet, a)
a = a + 1
If Len(movem) = Len(movet) Then Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
If WindowState = 1 Then Caption = "Still Running" Else Caption = "Autorun Killer [Kiransoft Nepal]"
On Error GoTo errorlist
For i = 1 To 26
dlist(i) = Chr$(65 + i) & ":"
Next i
For i = 1 To 26
If Dir(dlist(i), vbVolume) <> "" Then
SetAttr dlist(i) & "/autorun.inf", vbNormal
If Dir(dlist(i) & "/autorun.inf") <> "" Or Dir(dlist(i) & "/autorun.exe") <> "" Or Dir(dlist(i) & "/Autorun.inf") <> "" Or Dir(dlist(i) & "/Autorun.exe") <> "" Then
MsgBox "Autorun Killer" & Chr(10) & Chr(10) & "An autorun file is detected in " & dlist(i) & Chr(10) & "Autorun killer has detected it" & Chr(10) & Chr(10) & "Action Taken: Auto-Delete" & Chr(10) & "Priority: Normal", vbInformation, "Autorun detected"
Call autorun(dlist(i), "/autorun.inf")
End If
End If
Next i
errorlist:
Call FileErrors(Err.Number)
End Sub

Private Sub Timer2_Timer()
Timer2.Interval = 100
If Me.Height > 100 Then
Me.Height = Me.Height - 100
Else
Unload Me
End If
End Sub
