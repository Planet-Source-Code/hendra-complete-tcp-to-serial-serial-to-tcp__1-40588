VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTCPToCOM 
   Caption         =   "TCP/IP To Serial Communication"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   Icon            =   "frmTCPToCOM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLog 
      Height          =   1335
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   33
      Top             =   3600
      Width           =   7290
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   7395
      TabIndex        =   29
      Top             =   5040
      Width           =   7455
      Begin VB.Label lblHakCipta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (c) 2002 R.Hendra Suryanegara"
         Height          =   195
         Left            =   75
         TabIndex        =   30
         Top             =   30
         Width           =   2970
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3465
      Left            =   3825
      TabIndex        =   15
      Top             =   0
      Width           =   3540
      Begin VB.CheckBox checkAutoActive 
         Caption         =   "A&uto Active"
         Height          =   240
         Left            =   1425
         TabIndex        =   32
         Top             =   1425
         Width           =   1785
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   30
         ScaleHeight     =   735
         ScaleWidth      =   3495
         TabIndex        =   22
         Top             =   2700
         Width           =   3495
         Begin VB.CheckBox checkActive 
            Caption         =   "&Active"
            DownPicture     =   "frmTCPToCOM.frx":058A
            Height          =   615
            Left            =   2640
            Picture         =   "frmTCPToCOM.frx":06D4
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   90
            Width           =   765
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   $"frmTCPToCOM.frx":081E
            ForeColor       =   &H00FFFFFF&
            Height          =   720
            Left            =   90
            TabIndex        =   24
            Top             =   75
            Width           =   1290
         End
      End
      Begin VB.TextBox txtServerName 
         Height          =   315
         Left            =   1425
         TabIndex        =   20
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox txtIPAddress 
         Height          =   315
         Left            =   1425
         TabIndex        =   17
         Top             =   675
         Width           =   1815
      End
      Begin VB.TextBox txtPort 
         Height          =   315
         Left            =   1425
         TabIndex        =   16
         Top             =   1050
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transmit :"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   2175
         Width           =   690
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receive :"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   1950
         Width           =   690
      End
      Begin VB.Label lblByteOut 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblByteOut"
         Height          =   195
         Left            =   975
         TabIndex        =   26
         Top             =   2175
         Width           =   720
      End
      Begin VB.Label lblByteIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblByteIn"
         Height          =   240
         Left            =   975
         TabIndex        =   25
         Top             =   1950
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name :"
         Height          =   195
         Left            =   -450
         TabIndex        =   21
         Top             =   375
         Width           =   1755
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address :"
         Height          =   195
         Left            =   -450
         TabIndex        =   19
         Top             =   750
         Width           =   1755
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Port :"
         Height          =   195
         Left            =   -450
         TabIndex        =   18
         Top             =   1050
         Width           =   1755
      End
   End
   Begin VB.Timer Timer1 
      Left            =   375
      Top             =   2850
   End
   Begin MSWinsockLib.Winsock SockServer 
      Index           =   0
      Left            =   225
      Top             =   2100
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SockListen 
      Left            =   225
      Top             =   1650
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   150
      Top             =   825
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   1
      RThreshold      =   1
      SThreshold      =   1
      InputMode       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      Height          =   3465
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CheckBox checkNull 
         Caption         =   "Null Discard"
         Height          =   240
         Left            =   1425
         TabIndex        =   31
         Top             =   3150
         Width           =   1815
      End
      Begin VB.ComboBox cboCOMPort 
         Height          =   315
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   300
         Width           =   1890
      End
      Begin VB.ComboBox cboSpeed 
         Height          =   315
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   675
         Width           =   1890
      End
      Begin VB.ComboBox cboDataBits 
         Height          =   315
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1050
         Width           =   1890
      End
      Begin VB.ComboBox cboParity 
         Height          =   315
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1425
         Width           =   1890
      End
      Begin VB.ComboBox cboStopBits 
         Height          =   315
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1800
         Width           =   1890
      End
      Begin VB.ComboBox cboHandshaking 
         Height          =   315
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2175
         Width           =   1890
      End
      Begin VB.CheckBox checkRTS 
         Caption         =   "RTS Enable"
         Height          =   240
         Left            =   1425
         TabIndex        =   2
         Top             =   2550
         Width           =   1665
      End
      Begin VB.CheckBox checkDTR 
         Caption         =   "DTR Enable"
         Height          =   240
         Left            =   1425
         TabIndex        =   1
         Top             =   2850
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COM Port :"
         Height          =   195
         Left            =   -450
         TabIndex        =   14
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Speed :"
         Height          =   195
         Left            =   -450
         TabIndex        =   13
         Top             =   675
         Width           =   1755
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data bits :"
         Height          =   195
         Left            =   -450
         TabIndex        =   12
         Top             =   1050
         Width           =   1755
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Parity :"
         Height          =   195
         Left            =   -450
         TabIndex        =   11
         Top             =   1425
         Width           =   1755
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stop bits :"
         Height          =   195
         Left            =   -450
         TabIndex        =   10
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Handshaking :"
         Height          =   195
         Left            =   -450
         TabIndex        =   9
         Top             =   2175
         Width           =   1755
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "MenuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuKeluar 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmTCPToCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MaxConnection As Integer
Dim nByteOut As Long
Dim nByteIn As Long

Private Sub checkActive_Click()
On Local Error GoTo Trap
    If checkActive.Value = vbChecked Then
        txtPort.Enabled = False
        cboCOMPort.Enabled = False
        cboSpeed.Enabled = False
        cboDataBits.Enabled = False
        cboParity.Enabled = False
        cboStopBits.Enabled = False
        cboHandshaking.Enabled = False
        checkRTS.Enabled = False
        checkDTR.Enabled = False
        checkNull.Enabled = False
        txtPort.Enabled = False
        txtServerName.Enabled = False
        txtIPAddress.Enabled = False
        checkAutoActive.Enabled = False
        checkAutoActive.Enabled = False
        
        SockListen.LocalPort = txtPort.Text
        SockListen.Listen
        lblInfo.Caption = "Listening..." & vbCrLf & txtIPAddress.Text & ":" & txtPort.Text & vbCrLf & SocketCount & " User Connected..."
        
        MSComm1.CommPort = cboCOMPort.ListIndex + 1
        MSComm1.Settings = cboSpeed.Text & "," & Left(cboParity.Text, 1) & "," & cboDataBits.Text & "," & cboStopBits.Text
        MSComm1.InputLen = 0
        MSComm1.SThreshold = 1
        MSComm1.SThreshold = 1
        If checkRTS.Value = vbChecked Then MSComm1.RTSEnable = True Else MSComm1.RTSEnable = False
        If checkDTR.Value = vbChecked Then MSComm1.DTREnable = True Else MSComm1.DTREnable = False
        If checkNull.Value = vbChecked Then MSComm1.NullDiscard = True Else MSComm1.NullDiscard = False
        MSComm1.Handshaking = cboHandshaking.ListIndex
        MSComm1.PortOpen = True
        SimpanSeting
    Else
        txtPort.Enabled = True
        cboCOMPort.Enabled = True
        cboSpeed.Enabled = True
        cboDataBits.Enabled = True
        cboParity.Enabled = True
        cboStopBits.Enabled = True
        cboHandshaking.Enabled = True
        checkRTS.Enabled = True
        checkDTR.Enabled = True
        checkNull.Enabled = True
        txtPort.Enabled = True
        'txtServerName.Enabled = True
        'txtIPAddress.Enabled = True
        checkAutoActive.Enabled = True
        checkAutoActive.Enabled = True
        lblInfo.Caption = ""
        MSComm1.PortOpen = False
        SockListen.Close
        CloseAllSocket
    End If
    Exit Sub
    
Trap:
    MsgBox Err.Description, vbCritical, "Error"
    Resume Next
    Exit Sub
End Sub

Sub SimpanSeting()
    SaveSetting "R.Hendra Suryanegara", "TCP2COM", "ComPort", cboCOMPort.Text
    SaveSetting "R.Hendra Suryanegara", "TCP2COM", "Speed", cboSpeed.Text
    SaveSetting "R.Hendra Suryanegara", "TCP2COM", "DataBits", cboDataBits.Text
    SaveSetting "R.Hendra Suryanegara", "TCP2COM", "Parity", cboParity.Text
    SaveSetting "R.Hendra Suryanegara", "TCP2COM", "StopBits", cboStopBits.Text
    SaveSetting "R.Hendra Suryanegara", "TCP2COM", "Handshaking", cboHandshaking.Text
    SaveSetting "R.Hendra Suryanegara", "TCP2COM", "RTS", checkRTS.Value
    SaveSetting "R.Hendra Suryanegara", "TCP2COM", "DTR", checkDTR.Value
    SaveSetting "R.Hendra Suryanegara", "TCP2COM", "NullDiscard", checkNull.Value
    SaveSetting "R.Hendra Suryanegara", "TCP2COM", "Port", txtPort.Text
    SaveSetting "R.Hendra Suryanegara", "TCP2COM", "AutoActive", checkAutoActive.Value
End Sub

Private Sub Form_Load()
    MaxConnection = 10
    BuildSocket
    
    lblInfo.Caption = ""
    lblByteIn.Caption = ""
    lblByteOut.Caption = ""
    lblHakCipta.Caption = "Copyright (c) 2002 R.Hendra Suryanegara"
    txtServerName.Text = SockListen.LocalHostName
    txtIPAddress.Text = SockListen.LocalIP
    txtServerName.Enabled = False
    txtIPAddress.Enabled = False
        
    For Index = 1 To 8
        cboCOMPort.AddItem Index
    Next
    
    cboSpeed.AddItem "110"
    cboSpeed.AddItem "300"
    cboSpeed.AddItem "600"
    cboSpeed.AddItem "1200"
    cboSpeed.AddItem "2400"
    cboSpeed.AddItem "4800"
    cboSpeed.AddItem "9600"
    cboSpeed.AddItem "14400"
    cboSpeed.AddItem "19200"
    cboSpeed.AddItem "28800"
    cboSpeed.AddItem "38400"
    cboSpeed.AddItem "57600"
    cboSpeed.AddItem "115200"
    cboSpeed.AddItem "128000"
    cboSpeed.AddItem "256000"
    
    cboDataBits.AddItem "4"
    cboDataBits.AddItem "5"
    cboDataBits.AddItem "6"
    cboDataBits.AddItem "7"
    cboDataBits.AddItem "8"
    
    cboParity.AddItem "Even"
    cboParity.AddItem "Odd"
    cboParity.AddItem "None"
    cboParity.AddItem "Mark"
    cboParity.AddItem "Space"
    
    cboStopBits.AddItem "1"
    cboStopBits.AddItem "1.5"
    cboStopBits.AddItem "2"
    
    cboHandshaking.AddItem "None"       '0
    cboHandshaking.AddItem "XonXoff"    '1
    cboHandshaking.AddItem "RTS"        '2
    cboHandshaking.AddItem "RTSXonXoff" '3
    
    txtPort.Text = GetSetting("R.Hendra Suryanegara", "TCP2COM", "Port", "2002")
    cboCOMPort.Text = GetSetting("R.Hendra Suryanegara", "TCP2COM", "ComPort", "1")
    cboSpeed.Text = GetSetting("R.Hendra Suryanegara", "TCP2COM", "Speed", "19200")
    cboDataBits.Text = GetSetting("R.Hendra Suryanegara", "TCP2COM", "DataBits", "8")
    cboParity.Text = GetSetting("R.Hendra Suryanegara", "TCP2COM", "Parity", "None")
    cboStopBits.Text = GetSetting("R.Hendra Suryanegara", "TCP2COM", "StopBits", "1")
    cboHandshaking.Text = GetSetting("R.Hendra Suryanegara", "TCP2COM", "Handshaking", "XonXoff")
    checkRTS.Value = GetSetting("R.Hendra Suryanegara", "TCP2COM", "RTS", "0")
    checkDTR.Value = GetSetting("R.Hendra Suryanegara", "TCP2COM", "DTR", "1")
    checkNull.Value = GetSetting("R.Hendra Suryanegara", "TCP2COM", "NullDiscard", "1")
    txtPort.Text = GetSetting("R.Hendra Suryanegara", "TCP2COM", "Port", "2002")
    checkAutoActive.Value = GetSetting("R.Hendra Suryanegara", "TCP2COM", "AutoActive", "0")
    If checkAutoActive.Value = vbChecked Then
        checkActive.Value = vbChecked
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuKeluar_Click()
    Form_Unload 1
End Sub

Private Sub mnuRestore_Click()
    Me.WindowState = vbNormal
    Me.Show
End Sub

Private Sub MSComm1_OnComm()
    Dim EVMsg$
    Dim ERMsg$
    Select Case MSComm1.CommEvent
        Case comEvReceive
            Dim Buffer As String
            Buffer = MSComm1.Input
            
            Buffer = (StrConv(Buffer, vbUnicode))
            SendToAll Buffer
        
        Case comEvSend
        Case comEvCTS
            EVMsg$ = "Change in CTS Detected"
        Case comEvDSR
            EVMsg$ = "Change in DSR Detected"
        Case comEvCD
            EVMsg$ = "Change in CD Detected"
        Case comEvRing
            EVMsg$ = "The Phone is Ringing"
        Case comEvEOF
            EVMsg$ = "End of File Detected"

        ' Error messages.
        Case comBreak
            ERMsg$ = "Break Received"
        Case comCDTO
            ERMsg$ = "Carrier Detect Timeout"
        Case comCTSTO
            ERMsg$ = "CTS Timeout"
        Case comDCB
            ERMsg$ = "Error retrieving DCB"
        Case comDSRTO
            ERMsg$ = "DSR Timeout"
        Case comFrame
            ERMsg$ = "Framing Error"
        Case comOverrun
            ERMsg$ = "Overrun Error"
        Case comRxOver
            ERMsg$ = "Receive Buffer Overflow"
        Case comRxParity
            ERMsg$ = "Parity Error"
        Case comTxFull
            ERMsg$ = "Transmit Buffer Full"
        Case Else
            ERMsg$ = "Unknown error or event"
    End Select
    
    If Len(EVMsg$) Then
        AddLog EVMsg$
    ElseIf Len(ERMsg$) Then
        AddLog ERMsg$
    End If
End Sub

Private Sub SockListen_ConnectionRequest(ByVal requestID As Long)
    Dim Index As Integer
    Index = FindFreeSocket
    If Index <> -1 Then
        SockServer(Index).Close
        SockServer(Index).Accept requestID
        AddLog SockServer(Index).RemoteHostIP & " connected"
        lblInfo.Caption = "Listening..." & vbCrLf & txtIPAddress.Text & ":" & txtPort.Text & vbCrLf & SocketCount & " User Connected..."
    End If
    Exit Sub
    
Trap:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub SockServer_Close(Index As Integer)
    AddLog SockServer(Index).RemoteHostIP & " disconnected"
    SockServer(Index).Close
    lblInfo.Caption = "Listening..." & vbCrLf & txtIPAddress.Text & ":" & txtPort.Text & vbCrLf & SocketCount & " User Connected..."
End Sub

Private Sub SockServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim IncomingData As String
    SockServer(Index).GetData IncomingData
    MSComm1.Output = IncomingData
    
    nByteIn = nByteIn + bytesTotal
    lblByteIn = Trim(Format(nByteIn, "###,###,###,###")) & " bytes"
End Sub


Private Sub SockServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddLog SockServer(Index).RemoteHostIP & " error : " & Description
    SockServer(Index).Close
    lblInfo.Caption = "Listening..." & vbCrLf & txtIPAddress.Text & ":" & txtPort.Text & vbCrLf & SocketCount & " User Connected..."
End Sub

Private Sub SockServer_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    nByteOut = nByteOut + bytesSent
    lblByteOut = Trim(Format(nByteOut, "###,###,###,###")) & " bytes"
End Sub


Sub BuildSocket()
    Dim Index As Integer
    For Index = 1 To MaxConnection
        Load SockServer(Index)
    Next
End Sub

Function FindFreeSocket() As Integer
    Dim Index As Integer
    FindFreeSocket = -1
    For Index = 0 To MaxConnection
        If SockServer(Index).State <> sckConnected Then
           FindFreeSocket = Index
           Exit Function
        End If
    Next
End Function

Function SocketCount() As Integer
    Dim Index As Integer
    SocketCount = 0
    For Index = 0 To MaxConnection
        If SockServer(Index).State = sckConnected Then
           SocketCount = SocketCount + 1
        End If
    Next
End Function

Sub SendToAll(strData As String)
    Dim Index As Integer
    For Index = 0 To MaxConnection
        If SockServer(Index).State = sckConnected Then
           SockServer(Index).SendData strData
           DoEvents
        End If
    Next
End Sub


Sub AddLog(strText As String)
    txtLog.SelText = strText & vbCrLf
End Sub

Sub CloseAllSocket()
    Dim Index As Integer
    For Index = 0 To MaxConnection
        SockServer(Index).Close
    Next
End Sub



