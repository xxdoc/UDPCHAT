VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lan Messenger"
   ClientHeight    =   6405
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "chatwindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "chatwindow.frx":13498
   ScaleHeight     =   6405
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox combofont 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   4440
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   10560
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdplay 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      Picture         =   "chatwindow.frx":86935
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Play"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton cmdpause 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      Picture         =   "chatwindow.frx":87137
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Pause"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton cmdstop 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      Picture         =   "chatwindow.frx":881C0
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Stop"
      Top             =   5640
      Width           =   615
   End
   Begin MCI.MMControl mm 
      Height          =   495
      Left            =   7920
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   873
      _Version        =   393216
      PauseEnabled    =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdopen 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      Picture         =   "chatwindow.frx":884A6
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Open Music Files"
      Top             =   5640
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Interval        =   40000
      Left            =   6000
      Top             =   3840
   End
   Begin VB.CommandButton cmdbuzz 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Picture         =   "chatwindow.frx":88B5A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "BUZZ !!"
      Top             =   4320
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5400
      Top             =   2880
   End
   Begin VB.TextBox txtdisp 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton cmdconnect 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Connect To"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   2
      ToolTipText     =   "Lets you connect to the required user."
      Top             =   3120
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock peera 
      Left            =   7200
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton cmdsend 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   1
      ToolTipText     =   "Sends the message."
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox txtsend 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   6255
   End
   Begin VB.Label lblfont 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Font:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   24
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label lblnowplaying 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   22
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label lbltype 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   45
   End
   Begin VB.Label lblremoteportd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8520
      TabIndex        =   14
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label lblremoteipd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8520
      TabIndex        =   13
      Top             =   2160
      Width           =   45
   End
   Begin VB.Label lbllocalhostnamed 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8520
      TabIndex        =   12
      Top             =   1800
      Width           =   45
   End
   Begin VB.Label lbllocalportd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8520
      TabIndex        =   11
      Top             =   1440
      Width           =   45
   End
   Begin VB.Label lbllocalipd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8520
      TabIndex        =   10
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label lbllocalip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local IP:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7485
      TabIndex        =   9
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label lbllocalport 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local Port:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7440
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lbllocalhostname 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local Host Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6915
      TabIndex        =   7
      Top             =   1800
      Width           =   1380
   End
   Begin VB.Label lblremoteip 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remote IP:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7305
      TabIndex        =   6
      Top             =   2160
      Width           =   870
   End
   Begin VB.Label Lblremoteport 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remote Port:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7260
      TabIndex        =   5
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label lbltime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9135
      TabIndex        =   4
      Top             =   120
      Width           =   105
   End
   Begin VB.Menu file_menu 
      Caption         =   "File"
      Begin VB.Menu new_chat 
         Caption         =   "New Chat"
         Shortcut        =   ^N
      End
      Begin VB.Menu save_chat 
         Caption         =   "Save Chat"
         Shortcut        =   ^S
      End
      Begin VB.Menu readchat_history 
         Caption         =   "Read Chat History"
         Shortcut        =   ^R
      End
      Begin VB.Menu exit_form 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sname, prevname, nowplaying As String
Dim fs, fss, fso, songname As New FileSystemObject
Dim fldr As Folder
Dim r As Double
Dim rr As Integer

Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub cmdbuzz_Click()
On Error Resume Next
peera.SendData "ffffbuzz"
txtdisp.Text = txtdisp.Text & "BUZZ !! " & vbNewLine
txtdisp.SelStart = Len(txtdisp.Text)
Beep
cmdbuzz.Enabled = False
End Sub

Private Sub cmdconnect_Click()
On Error Resume Next
On Error GoTo herror_lport
exec1: peera.LocalPort = InputBox("Enter the Local port number", "LAN Messenger")
exec2: peera.RemoteHost = InputBox("Enter the Remote IP address or name of the remote machine to which you want to connect.", "LAN Messenger")
On Error GoTo herror_rport
peera.RemotePort = InputBox("Enter the Remote Port address of the Remote computer", "LAN Messenger")
peera.Bind
peera.SendData "ffffonline"
Form1.Caption = "Chat with " & peera.RemoteHost
lbllocalipd.Caption = peera.LocalIP
lbllocalportd.Caption = peera.LocalPort
lblremoteipd.Caption = peera.RemoteHost
lblremoteportd.Caption = peera.RemotePort
lbllocalhostnamed.Caption = peera.LocalHostName
cmdconnect.Visible = False
txtsend.Enabled = True
txtsend.SetFocus
cmdbuzz.Visible = True
Exit Sub
herror_lport:
rr = MsgBox("Please enter valid Localport.Press to YES enter and NO to exit", vbYesNo, "LAN Messenger")
If rr = 6 Then
Resume exec1
Else
End
End If
herror_rport:
rr = MsgBox("Please enter a valid Remote IP and Remote Port", vbYesNo, "LAN Messenger")
If rr = 6 Then
Resume exec2
Else
End
End If
End Sub

Private Sub cmdopen_Click()
On Error Resume Next
cd1.ShowOpen
sname = cd1.FileName
If (prevname <> sname) Then
mm.Command = "close"
End If
mm.FileName = sname
mm.Command = "open"
mm.Command = "play"
prevname = sname
Set songname = CreateObject("Scripting.FileSystemObject")
nowplaying = songname.GetFileName(sname)
lblnowplaying.Caption = nowplaying
End Sub

Private Sub cmdpause_Click()
mm.Command = "pause"
End Sub

Private Sub cmdplay_Click()
mm.Command = "open"
mm.Command = "play"
End Sub

Private Sub cmdsend_Click()
On Error Resume Next
Dim sendmsg As String
If txtsend.Text <> "" Then
sendmsg = peera.LocalHostName & " : " & txtsend.Text
peera.SendData sendmsg
peera.SendData "ffffreceived"
txtdisp.Text = txtdisp.Text & sendmsg & vbNewLine
txtdisp.SelStart = Len(txtdisp.Text)
txtsend.Text = ""
txtsend.SetFocus
Beep
End If
End Sub

Private Sub cmdstop_Click()
mm.Command = "close"
mm.Command = "stop"
End Sub

Private Sub combofont_Click()
Dim temp As String
temp = combofont.Text
txtdisp.Font = temp
End Sub

Private Sub exit_form_Click()
End
End Sub

Private Sub Form_Load()
Form1.Show
cmdconnect.Visible = True
cmdsend.Enabled = False
txtsend.Enabled = False
cmdbuzz.Visible = False
combofont.Clear
combofont.AddItem ("Arial")
combofont.AddItem ("Book Antiqua")
combofont.AddItem ("Bradley Hand ITC")
combofont.AddItem ("Calibri")
combofont.AddItem ("Cambria")
combofont.AddItem ("Century Gothic")
combofont.AddItem ("Comic Sans MS")
combofont.AddItem ("Consolas")
combofont.AddItem ("Kristen ITC")
combofont.AddItem ("Monotype Corsiva")
combofont.AddItem ("Segoe Print")
combofont.AddItem ("Segoe Script")
combofont.AddItem ("Times New Roman")

End Sub


Private Sub Form_Unload(Cancel As Integer)
If peera.RemoteHostIP <> "" Then
peera.SendData "ffffoffline"
End If
End Sub

Private Sub new_chat_Click()
On Error Resume Next
Dim newchat As New Form1
Set newchat = New Form1
newchat.Show
End Sub

Private Sub peera_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim rmsg As String
Dim c As Integer
peera.GetData rmsg
If rmsg = "ffffonline" Then
lbltype.Caption = peera.RemoteHostIP & " is online. "
txtsend.Enabled = True
cmdsend.Enabled = True
cmdbuzz.Enabled = True
combofont.Enabled = True
Timer2.Enabled = True
End If

If rmsg = "ffffoffline" Then
lbltype.Caption = peera.RemoteHostIP & "is offline. "
txtsend.Enabled = False
cmdsend.Enabled = False
cmdbuzz.Enabled = False
combofont.Enabled = False
Timer2.Enabled = False
End If
If rmsg = "fffftyping" Then
lbltype.Caption = peera.RemoteHostIP & " is Typing..."
End If
If rmsg <> "fffftyping" And rmsg <> "ffffreceived" And rmsg <> "ffffbuzz" And rmsg <> "ffffoffline" And rmsg <> "ffffonline" And rmsg <> "" Then
txtdisp.Text = txtdisp.Text & rmsg & vbNewLine
txtdisp.SelStart = Len(txtdisp.Text)
Beep
End If
If rmsg = "ffffreceived" Then
lbltype.Caption = peera.RemoteHostIP & " has entered."
End If
c = InStr(rmsg, "ffffbuzz")
If c <> 0 Or rmsg = "ffffbuzz" Then
txtdisp.Text = txtdisp.Text & "BUZZ !!" & vbNewLine
txtdisp.SelStart = Len(txtdisp.Text)
Beep
Form1.WindowState = 1
Form1.WindowState = 0
End If
End Sub



Private Sub peera_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox ("Not Connected to the remote machine !" & Description)
End Sub

Private Sub readchat_history_Click()
On Error Resume Next
Dim foltrue As Boolean
foltrue = fso.FolderExists("D:\myudpchatlog")
If foltrue = True Then
r = Shell("explorer D:\myudpchatlog", vbMaximizedFocus)
Else
Beep
rr = MsgBox("You do not have any chat logs !", vbOKOnly, "Lan Messenger")
End If
End Sub

Private Sub save_chat_Click()
On Error Resume Next
Dim savechat, fname, foldername As String
foldername = "D:\myudpchatlog"
Set fso = CreateObject("scripting.filesystemobject")
If fso.FolderExists(foldername) Then
Else
Set fldr = fso.CreateFolder(foldername)
End If
again: fname = InputBox("Enter the file name to be saved")
fname = "D:\myudpchatlog\" & fname & ".txt"
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FileExists(fname) Then
rr = MsgBox("File already exists! Please enter a different filename!", vbOKOnly, "Lan Messenger")
GoTo again
Else
Set a = fs.CreateTextFile(fname, True)
  savechat = txtdisp.Text
  a.WriteLine (savechat)
  a.WriteLine (vbCrLf & "Chat log saved on " & Date & " at " & Time)
  End If
  a.Close
End Sub

'Private Sub Timer1_Timer()
'lbltime.Caption = "Time : " & Time
'End Sub

Private Sub Timer2_Timer()
cmdbuzz.Enabled = True
End Sub

Private Sub txtsend_Change()
On Error Resume Next
If txtsend.Text <> "" Then
cmdsend.Enabled = True
peera.SendData "fffftyping"
End If
If txtsend.Text = "" Then
cmdsend.Enabled = False
End If
If txtsend.Text = "" Then
peera.SendData "ffffreceived"
End If
End Sub

Private Sub txtsend_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Call cmdsend_Click
End If

End Sub

