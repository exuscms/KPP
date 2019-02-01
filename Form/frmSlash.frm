VERSION 5.00
Begin VB.Form frmSlash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  '고정 도구 창
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '없음
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808080&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmSlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Starts As Integer
Dim Eps As Boolean
Dim TitleName As String
Private Sub Form_Load()
Starts = "0"
Eps = True
Label4 = "프로그램을 시작하는중..."
Label2.Caption = App.CompanyName & " / " & App.Comments
Label3.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Eps = False Then End
End Sub

Private Sub Timer1_Timer()
Starts = Starts + 1
    If Starts = "1" Then
        Label4 = "연결된 확장자의 파일을 확인중..."
    ElseIf Starts = "2" Then
        If Command <> "" Then
            Label4 = "연결된 확장자의 파일이 확인됨"
            Dim a
            a = FreeFile()
            Open (Command) For Input As #1
            frmMain.txtText.Text = ""
            Do While Not EOF(1)
            Line Input #1, a
            If EOF(1) = True Then
                frmMain.txtText.Text = frmMain.txtText.Text & a
            Else
                frmMain.txtText.Text = frmMain.txtText.Text & a & vbCrLf
            End If
            Loop
            Close
            frmMain.Caption = "KPP - " & (Command)
            frmMain.filenames = (Command)
            frmMain.fileSave = False
            Else
            Label4 = "연결된 확장자의 파일이 확인되지 않음"
            SetDefExt "KPP", "KPP 파일", ".kpp", App.Path & "\" & App.EXEName & ".exe"
            fileSave = False
            filenames = "새파일.kpp"
            frmMain.Caption = "KPP - 새파일.kpp"
        End If
        ElseIf Starts = "3" Then
            If Eps = True Then
                frmMain.Show
                Unload Me
            Else
                If Starts = "40" Then End
            End If
    End If
End Sub
Public Sub SetDefExt(AppName As String, Description As String, Extension As String, AppPath As String)
Dim ret As Long
Dim lphKey As Long
Dim FilePath As String
ret = RegCreateKey&(HKEY_CLASSES_ROOT, AppName, lphKey)
ret = RegSetValue&(lphKey&, Empty, REG_SZ, Description, 0&)
ret = RegCreateKey&(HKEY_CLASSES_ROOT, Extension, lphKey)
ret = RegSetValue&(lphKey, Empty, REG_SZ, AppName, 0&)
ret = RegCreateKey&(HKEY_CLASSES_ROOT, AppName, lphKey)
ret = RegSetValue&(lphKey, "shell\open\command", REG_SZ, AppPath & " %1", MAX_PATH)
    If Not ret = 0 Then
        Label4 = "관리자권한으로 실행하십시요"
        Eps = False
        Exit Sub
    End If
End Sub
