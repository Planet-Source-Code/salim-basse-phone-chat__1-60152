VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Dialer"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Select modem port"
      Height          =   1695
      Left            =   4680
      TabIndex        =   10
      Top             =   4080
      Width           =   1815
      Begin VB.OptionButton Option5 
         Caption         =   "port 3"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "port 2"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "port 1"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "pulse"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "tone"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   3960
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Call Wait"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      ToolTipText     =   "Message Received"
      Top             =   2160
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Message to Send"
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton CmdDial 
      Caption         =   "&Dial"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox TxtNumber 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Telephone No To Dial"
      Top             =   240
      Width           =   3615
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   2
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   360
      TabIndex        =   5
      ToolTipText     =   "Status"
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

DefInt A-Z
Dim p As Integer



Private Sub CmdDial_Click()
    Dim Number$
   
    CmdDial.Enabled = False
    Number$ = Trim(TxtNumber.Text)
    If Number$ = "" Then
    CmdDial.Enabled = True
    Exit Sub
    End If
    If portOk Then
      Dial Number$
    Else
      MsgBox "Modem Not Present...."
      MSComm1.PortOpen = False
    End If

    CmdDial.Enabled = True
End Sub
Private Sub Dial(Number$)
    Dim DialString$, FromModem$, dummy, n
    Label1.Caption = "Dialing ...."
    Text2.Text = ""
    If Option1 Then
    DialString$ = "ATDT" + Number$ + Chr$(13)
    Else
     DialString$ = "ATDp" + Number$ + Chr$(13)
     End If
    On Error Resume Next
    If Err Then
       MsgBox "COM3: not available. Change the CommPort property to another port."
       Exit Sub
    End If
    
    MSComm1.InBufferCount = 0
    
    MSComm1.Output = DialString$
   
    Do
       dummy = DoEvents()
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
             If InStr(FromModem$, "CONNECT") Then
                Beep
                Label1.Caption = "Connected ...."
                MsgBox "Please pick up the phone and either press Enter or click OK"
                Exit Do
            End If
       End If
        
       
          Exit Do
       
    Loop
End Sub
Private Function portOk() As Boolean
   Dim n As Single, s As String
   If Option3 Then
   p = 1
   ElseIf Option4 Then p = 2
   Else: p = 3
   End If
   
   
   MSComm1.CommPort = p
   MSComm1.Settings = "1200,N,8,1"
   portOk = False
   
   With MSComm1
      .PortOpen = True
      If Err = 0 Then
        .Output = "ATV1Q0" & Chr$(13)
        n = Timer
        While Timer - n < 1
          DoEvents
        Wend
        s = s & .Input
        If InStr(s, "OK" & vbCrLf) <> 0 Then
          portOk = True
          .Output = "ATZ" & Chr(13)
          n = Timer
          While Timer - n < 1
            DoEvents
          Wend

          .Output = "ATX1S36=3" & Chr(13)
          n = Timer
          While Timer - n < 1
            DoEvents
          Wend
          .Output = "AT+FCLASS=0" & Chr(13)
          n = Timer
          While Timer - n < 1
            DoEvents
          Wend
         
         s = s & .Input
        End If
      End If
  End With
End Function
Private Sub Command1_Click()
If Text1.Text = "" Then
  MsgBox "Enter the Message First", vbInformation, "Dialer"
  Text1.SetFocus
Else
  If MSComm1.PortOpen = True Then
    MSComm1.Output = Trim(Text1.Text) + vbCrLf
    Text2.Text = Text2.Text + ">>" + Trim(Text1.Text) + vbCrLf
  End If
End If
End Sub
Private Sub Command2_Click()
 
  If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
  Unload Me
End Sub

Private Sub Command3_Click()
Dim n As Single
If MSComm1.PortOpen = False Then
   MSComm1.PortOpen = True
   Text2.Text = "Port Open..."
Else
   Text2.Text = "Port Open..."
End If
Label1.Caption = "Port Open"

MSComm1.DTREnable = True 'Enables waiting for call

MSComm1.Output = "AT+FCLASS=0" & Chr(13)
n = Timer
While Timer - n < 1
  DoEvents
Wend

MSComm1.Output = "ATE1S0=1S36=3" & vbCrLf 'Wait for call
   
Text2.Text = Text2.Text + vbCrLf + "Waiting For Connection..."
Label1.Caption = "Waiting For Connection"
Text2.Text = ""

End Sub


Private Sub Option1_Validate(Cancel As Boolean)
Option2.Value = False
End Sub

Private Sub Option2_Validate(Cancel As Boolean)
Option1.Value = False
End Sub

Private Sub Option3_Validate(Cancel As Boolean)
Option4.Value = False
Option5.Value = False
p = 1
End Sub

Private Sub Option4_Validate(Cancel As Boolean)
Option3.Value = False
Option5.Value = False
p = 2
End Sub

Private Sub Option5_Validate(Cancel As Boolean)
Option4.Value = False
Option3.Value = False
p = 3
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

Case 13
Command1_Click
End Select

End Sub



Private Sub TxtNumber_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

Case 13
CmdDial_Click
End Select


End Sub

