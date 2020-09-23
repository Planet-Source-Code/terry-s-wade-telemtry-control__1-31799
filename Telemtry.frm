VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   Caption         =   "Telemtry Control"
   ClientHeight    =   2460
   ClientLeft      =   165
   ClientTop       =   645
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Telemtry.frx":0000
      Left            =   2760
      List            =   "Telemtry.frx":000A
      TabIndex        =   4
      Text            =   "Select Port"
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdReplyON 
      BackColor       =   &H0000FF00&
      Caption         =   "REPLY"
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2640
      Top             =   1560
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2640
      Top             =   600
   End
   Begin VB.CommandButton cmdON 
      BackColor       =   &H0000FF00&
      Caption         =   "ON"
      Height          =   855
      Left            =   360
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   120
   End
   Begin VB.CommandButton cmdReplyOff 
      BackColor       =   &H0000FF00&
      Caption         =   "REPLY"
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOFF 
      BackColor       =   &H0000FF00&
      Caption         =   "OFF"
      Height          =   855
      Left            =   360
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   5160
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      InBufferSize    =   512
      RThreshold      =   1
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      InBufferSize    =   512
      RThreshold      =   1
   End
   Begin VB.Label lblError 
      BackColor       =   &H00FFFF80&
      Caption         =   "PORT ERROR may be in use by other device"
      Height          =   735
      Left            =   2880
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Menu mnuFilet 
      Caption         =   "&File"
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu Info 
         Caption         =   "Info"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comFLAG As Boolean   ' true is COM1, False is COM2




Private Sub About_Click()
MsgBox ("Version 1.0  14FEB2002")
End Sub

Private Sub cmdON_Click()
If Combo1.Text <> "Select Port" Then

On Error Resume Next

'Sends "00"
Select Case comFLAG
  Case Is = True
    MSComm1.Output = "OO"
  Case Is = False
    MSComm2.Output = "OO"
End Select
DoEvents
  
cmdON.BackColor = &HFF&
cmdON.Enabled = False

Timer3.Enabled = True
cmdON.Enabled = True


lblError.Visible = False

Else
cmdON.Caption = "ERROR Select Port"
End If


End Sub




Public Sub cmdOFF_Click()

If Combo1.Text <> "Select Port" Then

On Error Resume Next

'Sends "FF"
Select Case comFLAG
  Case Is = True
    MSComm1.Output = "FF"
  Case Is = False
    MSComm2.Output = "FF"
End Select
DoEvents

cmdOFF.BackColor = &HFF&
cmdOFF.Enabled = False

Timer1.Enabled = True
cmdOFF.Enabled = True


lblError.Visible = False


Else

cmdOFF.Caption = "ERROR Select Port"

End If

End Sub







Private Sub cmdReply_Click()

End Sub

Public Sub Combo1_Click()
On Error GoTo ERR_Port
If Combo1.Text = "ComPort1" Then
  comFLAG = True
  lblError.Visible = False
  MSComm1.PortOpen = True
  cmdON.Visible = True
  cmdOFF.Visible = True
  
End If



If Combo1.Text = "ComPort2" Then
  comFLAG = False   ' for com2
lblError.Visible = False
MSComm2.PortOpen = True
test = MSComm2.CommEvent

cmdON.Visible = True
  cmdOFF.Visible = True

End If

  If Combo1.Text = "????" Then   ' This is a dummy "if" to prevent other then an error from entering
ERR_Port:
   lblError.Visible = True
   
   cmdON.Visible = True
  cmdOFF.Visible = True
  End If



End Sub



Private Sub Exit_Click()
On Error Resume Next

MSComm1.PortOpen = False
MSComm2.PortOpen = False


End

End Sub

Public Sub Form_Load()
Dim OnFlag As Boolean

OnFlag = True

End Sub

Private Sub mnuEdit_Click()

End Sub




Private Sub Info_Click()
MsgBox (" Only Com1 and Com2 Ports are supported, 9600,n,8,1    Rx,Tx and GND  is all that is needed to operate.  OFF sends 'FF' and expects 'FF' in return, ON sends 'OO' ")

End Sub

Sub MSComm1_OnComm()
 
 

 Select Case MSComm1.CommEvent
        ' Event messages.
        Case comEvReceive
        
        Data = MSComm1.Input
        
        
          If (Data = "FF") Then
             Timer1.Enabled = False
             Timer2.Enabled = False
             
             cmdOFF.BackColor = &HFF00&
             cmdOFF.Caption = "OFF"
             
             cmdReplyOff.Visible = True ' show Reply box
             DoEvents
             For index3 = 1 To 10000
             For index4 = 1 To 2000
             Next index4
             Next index3
             
             cmdReplyOff.Visible = False
                        
          End If
        
        
        If (Data = "OO") Then
             Timer3.Enabled = False
             Timer4.Enabled = False
             
             cmdON.BackColor = &HFF00&
             cmdON.Caption = "ON"
             
             cmdReplyON.Visible = True ' show Reply box
             DoEvents
             For index3 = 1 To 10000
             For index4 = 1 To 2000
             Next index4
             Next index3
        
         cmdReplyON.Visible = False
                        
          End If
        
        
        
        Case comEvSend
        ' do nothing
        
   End Select



End Sub

Private Sub MSComm2_OnComm()
 

 Select Case MSComm2.CommEvent
        ' Event messages.
        Case comEvReceive
        
        Data = MSComm2.Input
        
        
          If (Data = "FF") Then
             Timer1.Enabled = False
             Timer2.Enabled = False
             
             cmdOFF.BackColor = &HFF00&
             cmdOFF.Caption = "OFF"
             
             cmdReplyOff.Visible = True ' show Reply box
             DoEvents
             For index3 = 1 To 10000
             For index4 = 1 To 2000
             Next index4
             Next index3
             
             cmdReplyOff.Visible = False
                        
          End If
        
        
        If (Data = "OO") Then
             Timer3.Enabled = False
             Timer4.Enabled = False
             
             cmdON.BackColor = &HFF00&
             cmdON.Caption = "ON"
             
             cmdReplyON.Visible = True ' show Reply box
             DoEvents
             For index3 = 1 To 10000
             For index4 = 1 To 2000
             Next index4
             Next index3
        
         cmdReplyON.Visible = False
                        
          End If
        
        
        
        Case comEvSend
        ' do nothing
        
   End Select


 

End Sub





















Private Sub Timer1_Timer()
Timer1.Enabled = False

cmdOFF.Caption = "Timed Out"
DoEvents
Timer2.Enabled = True



End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
cmdOFF.Caption = "OFF"
cmdOFF.BackColor = &HFF00&
DoEvents
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False

cmdON.Caption = "Timed Out"
DoEvents
Timer4.Enabled = True

End Sub

Private Sub Timer4_Timer()
Timer4.Enabled = False
cmdON.Caption = "ON"
cmdON.BackColor = &HFF00&
DoEvents
End Sub



