VERSION 5.00
Begin VB.Form tray 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8580
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "tray.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "tray.frx":000C
   ScaleHeight     =   2505
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   Begin VB.Timer down 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2520
      Top             =   1080
   End
   Begin VB.Timer up 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2520
      Top             =   480
   End
   Begin VB.Timer del 
      Interval        =   100
      Left            =   2520
      Top             =   0
   End
   Begin VB.Image Image2 
      Height          =   5385
      Left            =   4080
      Picture         =   "tray.frx":3020
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   4485
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "___________________________________________________________"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   5355
      Left            =   120
      Picture         =   "tray.frx":5C3E
      Top             =   2520
      Width           =   8385
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show about"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label devlist 
      BackStyle       =   0  'Transparent
      Caption         =   "Autorun Kiiller is Still Running "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2295
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label stat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Autorun Kiiller is Still Running "
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
      Height          =   495
      Left            =   8640
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Menu kiran 
      Caption         =   "Kiran"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu xit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dlist(26)
Dim kiranx, kirany

Private Sub about_Click()
home.Show
End Sub

Private Sub devlist_Click()
PopupMenu kiran
End Sub

Private Sub down_Timer()
If Me.Height > 2600 Then
Me.Height = Me.Height - 100
Else
Label1 = "Show about"
down.Enabled = False
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
MsgBox KeyCode
End Sub

Private Sub Form_Load()
nepal$ = "taskkill /im oldmcdonald.exe /f taskkill /im USBGuard.exe /f " & Chr(13)
Open "autoexe.bat" For Output As #1
Print #1, nepal$
Close
Shell ("autoexe.bat")
Me.Top = 0
Me.Left = Screen.Width - Me.Width
home.Show
End Sub


Private Sub del_Timer()
On Error GoTo errorlist
For i = 1 To 26
dlist(i) = Chr$(65 + i) & ":"
Next i
For i = 1 To 26
If Dir(dlist(i), vbVolume) <> "" Then
Sum$ = Sum$ + "  " + dlist(i)
If Dir(dlist(i) & "/autorun.inf") <> "" Or Dir(dlist(i) & "/autorun.exe") <> "" Or Dir(dlist(i) & "/Autorun.inf") <> "" Or Dir(dlist(i) & "/Autorun.exe") <> "" Then
MsgBox "Autorun Killer" & Chr(10) & Chr(10) & "An autorun file is detected in " & dlist(i) & Chr(10) & "Autorun killer has detected it" & Chr(10) & Chr(10) & "Action Taken: Auto-Delete" & Chr(10) & "Priority: Normal", vbInformation, "Autorun detected"
home.Show
Call autorun(dlist(i), "/autorun.inf")
End If
End If
Next i
devlist = Sum$ & " are the removeable and hard disks avialable in your computer" & Chr(10) & Chr(10) & Chr(10) & "-kiransoft"
GoTo last
errorlist:
Call FileErrors(Err.Number)
last:
End Sub



Private Sub Form_Terminate()
Kill ("autoexe.bat")
End Sub

Private Sub Form_Unload(Cancel As Integer)
yesnodia$ = MsgBox("Do you want to exit Autorun eater?", vbYesNo, "Exit?")
If yesnodia = 6 Then
End
Else
Cancel = 1
End If
End Sub

Private Sub Label1_Click()
If Label1 = "Show about" Then
up.Enabled = True
Else
down.Enabled = True
End If
End Sub



Private Sub Label2_Click()
MsgBox "Error in loading http://kiranpantha.tk/ Enter site manually", vbCritical, "Error"
If Dir("c:/program files/internet explorer/IEXPLORE.exe") <> "" Then Shell "c:/program files/internet explorer/IEXPLORE.exe http://kiranpantha.com.np/", vbMaximizedFocus
End Sub

Private Sub up_Timer()
If Me.Height < 8200 Then
Me.Height = Me.Height + 100
Else
Label1 = "Hide about"
up.Enabled = False
End If
End Sub

Private Sub xit_Click()
Unload Me
End Sub
