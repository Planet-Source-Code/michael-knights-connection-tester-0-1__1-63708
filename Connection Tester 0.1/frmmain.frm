VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connection Tester 0.1 by Michael Knights"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Socket 
      Left            =   6000
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Connection"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   5040
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
      Begin VB.Timer tmrstate 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   840
         Top             =   240
      End
      Begin VB.TextBox txtPort 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdDissconnect 
         BackColor       =   &H8000000A&
         Caption         =   "Dissconnect"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtHost 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lbl2 
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lbl1 
         Caption         =   "Host:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Send Data"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   6735
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtSend 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Connection Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4815
      Begin VB.TextBox txtinfo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Connection Type "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      Begin VB.CommandButton cmdUPD 
         Caption         =   "UPD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdTCP 
         BackColor       =   &H8000000A&
         Caption         =   "TCP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Incomming Data"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4815
      Begin VB.TextBox txtdata 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Label Label1 
      Caption         =   "www.mksoftware.co.uk "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5640
      Width           =   6735
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright Michael Knights - 2005


'Note - i do not like documenting my work
'so i will just label what each sub does.



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''' Socket '''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'If there is a socket error disaply it to the user
MsgBox "Socket Error!!! - " & Description, vbInformation, App.EXEName & " Error"
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
'socket data arrival - gets the data and displays in the incomming data textbox
On Error GoTo ERR_DA
Dim DaTa As String
Socket.GetData DaTa
txtdata = txtdata & DaTa & vbNewLine
Exit Sub
ERR_DA:
MsgBox "DataArrival Error!!!", vbInformation, App.EXEName & " Error"
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''' Text Boxes '''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Private Sub txtdata_Change()
'if text gets change auto scroll : Note - Im Soo Great Ant I!
txtdata.SelStart = Len(txtdata.Text)
End Sub

Private Sub txtHost_Change()
'on text change update variable
On Error GoTo ERR_U
ConnectionHost = txtHost.Text
Exit Sub
ERR_U:
MsgBox "Unknown Error!!!", vbInformation, App.EXEName & " Error"
End Sub

Private Sub txtport_Change()
'on text change update variable
On Error GoTo ERR_U
ConnectionPort = txtPort.Text
Exit Sub
ERR_U:
MsgBox "Unknown Error!!!", vbInformation, App.EXEName & " Error"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''' Command Buttons '''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdSend_Click()
'send data to the remote server
On Error GoTo ERR_SND
Socket.SendData txtSend.Text
Exit Sub
ERR_SND:
MsgBox "Error Sending Data!!!", vbInformation, App.EXEName & " Error"
End Sub

Private Sub cmdconnect_Click()
'connect to the server useing the users settings
On Error GoTo ERR_U
Socket.Close
cmdDissconnect.BackColor = &H8000000F
cmdConnect.BackColor = &H8000000A
If ConnectionType = "TCP" Then
Socket.Protocol = sckTCPProtocol
Socket.RemoteHost = "0"
Socket.LocalPort = "0"
Socket.RemotePort = "0"
Socket.Connect ConnectionHost, ConnectionPort
ElseIf ConnectionType = "UPD" Then
Socket.Protocol = sckUDPProtocol
Socket.RemoteHost = ConnectionHost
Socket.LocalPort = ConnectionPort
Socket.RemotePort = ConnectionPort
Socket.Connect ConnectionHost, ConnectionPort
Else
MsgBox "Could Not Deturman connection Type", vbCritical, App.EXEName & " Erorr"
Exit Sub
End If

tmrstate.Enabled = True
Exit Sub
ERR_U:
MsgBox "Unknown Error!!!", vbInformation, App.EXEName & " Error"
End Sub

Private Sub cmdDissconnect_Click()
'disconnect from the server
On Error GoTo ERR_U
cmdDissconnect.BackColor = &H8000000A
cmdConnect.BackColor = &H8000000F
Socket.Close
tmrstate.Enabled = False
txtinfo.Text = ""
Exit Sub
ERR_U:
MsgBox "Unknown Error!!!", vbInformation, App.EXEName & " Error"
End Sub

Private Sub cmdTCP_Click()
'change the setting for a TCP connection
On Error GoTo ERR_U
frmmain.cmdTCP.BackColor = &H8000000A
frmmain.cmdUPD.BackColor = &H8000000F
ConnectionType = "TCP"
Exit Sub
ERR_U:
MsgBox "Unknown Error!!!", vbInformation, App.EXEName & " Error"
End Sub

Private Sub cmdUPD_Click()
'change the setting for a UPD connection
On Error GoTo ERR_U
frmmain.cmdUPD.BackColor = &H8000000A
frmmain.cmdTCP.BackColor = &H8000000F
ConnectionType = "UPD"
Exit Sub
ERR_U:
MsgBox "Unknown Error!!!", vbInformation, App.EXEName & " Error"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''' Form ''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Load()
' load settings
On Error GoTo ERR_L
LoadSettings
Exit Sub
ERR_L:
MsgBox "Unable To Load Settings.", vbInformation, App.EXEName & " Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
' save settings
On Error GoTo ERR_S
SaveSettings
Exit Sub
ERR_S:
MsgBox "Unable To Save Settings.", vbInformation, App.EXEName & " Error"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''' Timers '''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub tmrstate_Timer()
'display the winsock status into a text box - WOW
On Error GoTo ERR_SATE
txtinfo = "Host: " & Socket.RemoteHost & " Port: " & Socket.RemotePort & vbNewLine & _
"Winsock State:" & Socket.State
Exit Sub
ERR_SATE:
MsgBox "Unable To display Corrent State!!!", vbInformation, App.EXEName & " Error"
End Sub





' ::: www.mksoftware.co.uk :::
