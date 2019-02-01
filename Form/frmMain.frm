VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C000&
   Caption         =   "KPP"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   14655
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":2DA2
   ScaleHeight     =   9015
   ScaleWidth      =   14655
   StartUpPosition =   2  '화면 가운데
   Begin MSComDlg.CommonDialog cdCompile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox ptop 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   1080
      ScaleHeight     =   4575
      ScaleWidth      =   11760
      TabIndex        =   0
      Top             =   1080
      Width           =   11760
      Begin VB.TextBox txtText 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         Height          =   375
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   2
         Top             =   0
         Width           =   120
      End
      Begin RichTextLib.RichTextBox txtDebug 
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   12632256
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":38C4
      End
      Begin VB.Image pright 
         Height          =   10
         Left            =   11640
         Picture         =   "frmMain.frx":3961
         Stretch         =   -1  'True
         Top             =   0
         Width           =   120
      End
      Begin VB.Image pleft 
         Height          =   10
         Left            =   0
         Picture         =   "frmMain.frx":39D3
         Stretch         =   -1  'True
         Top             =   0
         Width           =   105
      End
   End
   Begin MSComDlg.CommonDialog Cd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lab_create 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "torinoyume@naver.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2025
   End
   Begin VB.Image picbag 
      Height          =   135
      Left            =   0
      Picture         =   "frmMain.frx":3A45
      Stretch         =   -1  'True
      Top             =   0
      Width           =   135
   End
   Begin VB.Menu mnuFille 
      Caption         =   "파일(&F)"
      Begin VB.Menu mnuNew 
         Caption         =   "새 프로젝트(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "프로젝트 열기(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "프로젝트 저장(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "프로젝트를 다른이름으로 저장(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuLIne2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompile 
         Caption         =   "새파일.exe 만들기(&K)..."
      End
      Begin VB.Menu mnuLIne3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "프로그램 종료(&Q)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "설정(&S)"
      Begin VB.Menu mnuConfig 
         Caption         =   "설정(&S)"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "보기(&V)"
      Begin VB.Menu mnuDebug 
         Caption         =   "디버그(&D)"
         Checked         =   -1  'True
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "실행(&R)"
      Begin VB.Menu mnuParser 
         Caption         =   "시작(&S)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "도구(&T)"
      Begin VB.Menu mnuHandle 
         Caption         =   "핸들 관리자(&H)"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuExplorer 
         Caption         =   "파일 탐색기(&E)"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "도움말(&H)"
      Begin VB.Menu mnuFunction 
         Caption         =   "명령어(&F)"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fileSave As Boolean
Public filenames As String

Private Sub Form_Load()
filenames = "새파일"
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Not Me.WindowState = 1 Then
        If mnuDebug.Checked = True Then
            txtDebug.Visible = True
            ptop.Top = 0
            pright.Left = 11640
            pright.Height = Me.ScaleHeight
            pleft.Height = Me.ScaleHeight
            ptop.Left = (Me.ScaleWidth - ptop.Width) / 2
            ptop.Height = Me.ScaleWidth
            picbag.Width = Me.ScaleWidth
            picbag.Height = Me.ScaleHeight
            txtDebug.Left = 105
            txtDebug.Height = 1000
            txtText.Width = 11520
            txtText.Height = Me.ScaleHeight - txtDebug.Height
            txtText.Left = 105
            txtDebug.Width = txtText.Width + 10
            txtDebug.Top = Me.ScaleHeight - (Me.ScaleHeight - txtText.Height)
        Else
            txtDebug.Visible = False
            ptop.Top = 0
            pright.Height = Me.ScaleHeight
            pright.Left = 11640
            pleft.Height = Me.ScaleHeight
            ptop.Left = (Me.ScaleWidth - ptop.Width) / 2
            ptop.Height = Me.ScaleWidth
            picbag.Width = Me.ScaleWidth
            picbag.Height = Me.ScaleHeight
            txtText.Left = 105
            txtText.Height = Me.ScaleHeight
            txtText.Width = 11520
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Question
    If fileSave = True Then Question = MsgBox("변경된 내용을 " & filenames & "에 저장하시겠습니까?", vbYesNoCancel, "종료")
    If Question = 6 Then mnuSaveas_Click
    If Question = 7 Then End
    If Question = 2 Then Cancel = 1
End Sub

Private Sub mnuCompile_Click()

    On Local Error GoTo errTrap

    Dim BeginPos As Long
    Dim cpContent As New PropertyBag
    Dim varTemp As Variant
    
    Cd.Filter = "EXE 파일(*.kpp)|*.kpp"
    Cd.DialogTitle = "프로젝트 저장"
    Cd.FileName = App.Path & "\" & filenames & ".exe"
    Cd.ShowSave
    
    If Cd.FileName <> "" Then
        With cpContent
            .WriteProperty "Source", txtText.Text
        End With
        
        FileCopy App.Path & "\Compile.dll", Cd.FileName
        
        Open Cd.FileName For Binary As #1
            BeginPos = LOF(1)   'the point were we add extra data
            varTemp = cpContent.Contents
                    
            Seek #1, LOF(1)
            Put #1, , varTemp   'write data
            Put #1, , BeginPos  'write starting point of extra data
        
        Close #1
    
        MsgBox "컴파일 완료", vbInformation

    End If
    Exit Sub

errTrap:
    'to err is electronic
    Msg = "There was an error during compilation" & vbCrLf
    Msg = Msg & vbCrLf & Err.Description
    MsgBox Msg, vbCritical, "Error"
End Sub

Private Sub mnuConfig_Click()
frmSetup.Show
End Sub

Private Sub mnuDebug_Click()
    If mnuDebug.Checked = True Then
        mnuDebug.Checked = False
        Form_Resize
    Else
        mnuDebug.Checked = True
        Form_Resize
    End If
End Sub

Private Sub mnuExplorer_Click()
Cd.Filter = "MP3 파일(*.mp3)|*.mp3|모든 파일(*.*)|*.*"
Cd.DialogTitle = "프로젝트 열기"
Cd.ShowOpen
    If Not Cd.FileName = "" And Not Cd.flags = "0" Then txtText.Text = txtText.Text & Cd.FileName
End Sub

Private Sub mnuFunction_Click()
frmHelp.Show
End Sub

Private Sub mnuHandle_Click()
frmTask.Show
End Sub

Private Sub mnuNew_Click()
Dim Question
    If fileSave = True Then Question = MsgBox("변경된 내용을 " & filenames & "에 저장하시겠습니까?", vbYesNoCancel, "새 파일")
    If Question = 6 Then mnuSaveas_Click
    If Question = 7 Then
        Me.Caption = "KPP - " & filenames
        txtText = ""
        fileSave = False
    End If
        If fileSave = False Then
        txtText = ""
        filenames = "새파일.kpp"
        Me.Caption = "KPP - " & filenames
        fileSave = False
    End If
End Sub

Private Sub mnuOpen_Click()
Dim Question
    If fileSave = True Then Question = MsgBox("저장하지 않고 바로 열겠습니까?", vbYesNoCancel, "저장")
    If Question = 6 Then GoTo Opens
Opens:
        Dim a As String, b As String, f As Integer
        Cd.Filter = "KPP 파일(*.kpp)|*.kpp|모든 파일(*.*)|*.*"
        Cd.DialogTitle = "프로젝트 열기"
        Cd.ShowOpen
    If Not Cd.FileName = "" And Not Cd.flags = "0" Then
        a = FreeFile()
        Open Cd.FileName For Input As #1
        txtText.Text = ""
        Do While Not EOF(1)
        Line Input #1, a
    If EOF(1) = True Then txtText.Text = txtText.Text & a
    If EOF(1) = False Then txtText.Text = txtText.Text & a & vbCrLf
        Loop
        Close
        filenames = Cd.FileTitle
        Me.Caption = "KPP - " & filenames
        mnuCompile.Caption = Mid(filenames, 1, Len(filenames) - 4) & ".exe 만들기(&K)..."
        fileSave = False
    End If
End Sub

Public Sub mnuParser_Click()
Dim Run As New clsParser
frmRun.Show
txtDebug.Text = ""
Run.Parser txtText.Text
Run.Desp = False
Debugs "프로그램 시작"
End Sub
Public Sub Debugs(txtDebugs As String)
    If txtDebug.Text = "" Then
        txtDebug.Text = ">> " & txtDebugs
    Else
        txtDebug.Text = txtDebug.Text & vbCrLf & ">> " & txtDebugs
    End If
End Sub

Private Sub mnuQuit_Click()
End
End Sub

Private Sub mnuSave_Click()
On Error GoTo ass
    If Not filenames = "" Then
        Open filenames For Output As #1
        Print #1, txtText.Text
        Close #1
        Cd.FileName = filenames
        Me.Caption = "KPP - " & Cd.FileName
        fileSave = False
        Exit Sub
    Else
        mnuSaveas_Click
    End If
Exit Sub
ass:
mnuSaveas_Click
End Sub

Public Sub mnuSaveas_Click()
Cd.Filter = "KPP 파일(*.kpp)|*.kpp|모든 파일(*.*)|*.*"
Cd.DialogTitle = "프로젝트 저장"
Cd.ShowSave
    If Not Cd.FileName = "" And Not Cd.flags = "0" Then
        Open Cd.FileName For Output As #1
        Print #1, txtText.Text
        Close #1
        Me.Caption = "KPP - " & Cd.FileName
        filenames = Cd.FileTitle
        fileSave = False
        Exit Sub
    End If
End Sub

Private Sub txtText_Change()
    If fileSave = False Then
        fileSave = True
        Me.Caption = Me.Caption & " *"
    End If
End Sub
