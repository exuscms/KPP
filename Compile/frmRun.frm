VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmRun 
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9735
   Icon            =   "frmRun.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9735
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtSource 
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox usrpb 
      Height          =   615
      Index           =   0
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pb 
      BackColor       =   &H80000007&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtText 
      Height          =   4455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   1
      Top             =   0
      Width           =   7575
   End
   Begin VB.TextBox cnts 
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   3000
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   3360
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UserInput, InputAt As Boolean
Public strRecivedData1, strRecivedData2, asciis As String
Public Dis1, Dis2 As Boolean
Dim PropBag As New PropertyBag

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
asciis = KeyCode
End Sub

Private Sub Form_Load()

    Dim BeginPos As Long
    Dim varTemp As Variant
    Dim byteArr() As Byte
    
    Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
        Get #1, LOF(1) - 3, BeginPos    'get the start position of data

        Seek #1, BeginPos               'seek to data start
        Get #1, , varTemp               'get property bag contents
        
        byteArr = varTemp
        PropBag.Contents = byteArr      'load property bag
    
        PropBag.WriteProperty "LOF", LOF(1) 'a few extra props
        PropBag.WriteProperty "BeginPos", BeginPos
    Close #1
        
    
    'password protection
    'I know that this is not tight, but just for a demo
    With PropBag
        txtSource.Text = .ReadProperty("Source")
    End With
    
    Dim Run As New clsParser
    Run.Parser txtSource.Text
    Run.Desp = False
End Sub

Private Sub Form_Paint()
    If Dis1 = True Then
        PaintDesktop frmRun.hdc
    End If
End Sub

Private Sub Form_Resize()
    If Not Me.WindowState = 1 Then
        txtText.Width = Me.ScaleWidth
        txtText.Height = Me.ScaleHeight
        pb.Width = Me.ScaleWidth
        pb.Height = Me.ScaleHeight
    End If
End Sub

Private Sub pb_Paint()
    If Dis2 = True Then
        PaintDesktop frmRun.pb.hdc
    End If
End Sub

Private Sub Timer_Timer()
Timer.Enabled = False
End Sub

Private Sub txtText_Change()
txtText.SelStart = Len(txtText.Text)
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
Dim Content As String
    If txtText.Locked = False Then
        txtText.SelStart = Len(txtText.Text)
        If KeyAscii = vbKeyReturn Then
            If InputAt = 0 Then
                KeyAscii = 0
                Exit Sub
            Else
                txtText.Locked = True
            End If
            ElseIf KeyAscii = 8 Then
        If InputAt <= 0 Then
            KeyAscii = 0
            Exit Sub
        Else
            InputAt = InputAt - 1
            UserInput = Mid(UserInput, 1, Len(UserInput) - 1)
        End If
        Else
            UserInput = UserInput & Chr(KeyAscii)
            If Password = True Then KeyAscii = 42
        End If
        Else
            KeyAscii = 0
    End If
End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
asciis = KeyCode
End Sub

Private Sub wskClient_Close()
wskClient.Close
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
wskClient.GetData strRecivedData2
End Sub

Private Sub wskServer_Close(Index As Integer)
frmRun.cnts.Text = frmRun.cnts.Text - 1
frmRun.wskServer(Index).Close
End Sub

Private Sub wskServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
frmRun.cnts.Text = frmRun.cnts.Text + 1
SocketCount = SocketCount + 1
Load frmRun.wskServer(SocketCount)
frmRun.wskServer(SocketCount).Accept requestID
End Sub

Private Sub wskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
frmRun.wskServer(Index).GetData strRecivedData1
End Sub
