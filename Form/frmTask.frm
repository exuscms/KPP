VERSION 5.00
Begin VB.Form frmTask 
   BorderStyle     =   1  '단일 고정
   Caption         =   "핸들 관리자"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5535
   Icon            =   "frmTask.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   5535
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '없음
      Height          =   5895
      Left            =   0
      ScaleHeight     =   5895
      ScaleWidth      =   975
      TabIndex        =   12
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "다시읽기"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   5520
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "종료"
      Height          =   285
      Left            =   4680
      TabIndex        =   10
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "복구"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "보이기"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   8
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "숨키기"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "최소화"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdMaximize 
      Caption         =   "최대화"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox lstTask 
      Height          =   4380
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label labtitle 
      AutoSize        =   -1  'True
      Caption         =   "이름 : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   510
   End
   Begin VB.Label labhandle 
      AutoSize        =   -1  'True
      Caption         =   "핸들 :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHide_Click()
frmMain.txtText.Text = frmMain.txtText.Text & vbCrLf & "윈도우 제어(""" & Text2.Text & """, ""0"")"
End Sub

Private Sub cmdMaximize_Click()
frmMain.txtText.Text = frmMain.txtText.Text & vbCrLf & "윈도우 제어(""" & Text2.Text & """, ""3"")"
End Sub

Private Sub cmdMinimize_Click()
frmMain.txtText.Text = frmMain.txtText.Text & vbCrLf & "윈도우 제어(""" & Text2.Text & """, ""6"")"
End Sub

Private Sub cmdRestore_Click()
frmMain.txtText.Text = frmMain.txtText.Text & vbCrLf & "윈도우 제어(""" & Text2.Text & """, ""9"")"
End Sub

Private Sub cmdShow_Click()
frmMain.txtText.Text = frmMain.txtText.Text & vbCrLf & "윈도우 제어(""" & Text2.Text & """, ""5"")"
End Sub

Private Sub Command1_Click()
frmMain.txtText.Text = frmMain.txtText.Text & vbCrLf & "함수 process"
frmMain.txtText.Text = frmMain.txtText.Text & vbCrLf & "process ~ 프로세스핸들(""" & Text2.Text & """,""&H1F0FFF"",""a"")"
frmMain.txtText.Text = frmMain.txtText.Text & vbCrLf & "강제종료(process)"
End Sub

Private Sub Command2_Click()
lstTask.Clear
svar = EnumWindows(AddressOf getalltopwindows, 0)
End Sub

Private Sub Form_Load()
lstTask.Clear
svar = EnumWindows(AddressOf getalltopwindows, 0)
End Sub

Private Sub lstTask_Click()
For X = 0 To lstTask.ListCount - 1
    If lstTask.Selected(X) = True Then
        Text1.Text = lstTask.List(X)
        Text2.Text = lstTask.ItemData(X)
    End If
Next X
End Sub
