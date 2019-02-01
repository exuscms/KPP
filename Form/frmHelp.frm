VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "도움말"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7365
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   7365
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtHelp 
      Height          =   3375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox leftpb 
      Appearance      =   0  '평면
      BackColor       =   &H00808080&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   855
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   1
      Top             =   300
      Width           =   2175
   End
   Begin VB.ComboBox Cb 
      Appearance      =   0  '평면
      Height          =   300
      Left            =   840
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cb_Click()
    If Cb.ListIndex = "0" Then
        If Not Dir(App.Path & "\Doc\프로그램 소개.txt") = "" Then
            a = FreeFile()
            Open App.Path & "\Doc\프로그램 소개.txt" For Input As #1
            txtText.Text = ""
            Do While Not EOF(1)
            Line Input #1, a
            txtText.Text = txtText.Text & a & vbCrLf
            Loop
            Close
        End If
        ElseIf Cb.ListIndex = "1" Then
            If Not Dir(App.Path & "\Doc\기본명령어.txt") = "" Then
                a = FreeFile()
                Open App.Path & "\Doc\기본명령어.txt" For Input As #1
                txtText.Text = ""
                Do While Not EOF(1)
                Line Input #1, a
                txtText.Text = txtText.Text & a & vbCrLf
                Loop
                Close
            End If
            ElseIf Cb.ListIndex = "2" Then
                If Not Dir(App.Path & "\Doc\시스템명령어.txt") = "" Then
                a = FreeFile()
                Open App.Path & "\Doc\시스템명령어.txt" For Input As #1
                txtText.Text = ""
                Do While Not EOF(1)
                Line Input #1, a
                txtText.Text = txtText.Text & a & vbCrLf
                Loop
                Close
            End If
        ElseIf Cb.ListIndex = "3" Then
            If Not Dir(App.Path & "\Doc\문자열명령어.txt") = "" Then
                a = FreeFile()
                Open App.Path & "\Doc\문자열명령어.txt" For Input As #1
                txtText.Text = ""
                Do While Not EOF(1)
                Line Input #1, a
                txtText.Text = txtText.Text & a & vbCrLf
                Loop
                Close
            End If
            ElseIf Cb.ListIndex = "4" Then
                If Not Dir(App.Path & "\Doc\TCP,IP명령어.txt") = "" Then
                a = FreeFile()
                Open App.Path & "\Doc\TCP,IP명령어.txt" For Input As #1
                txtText.Text = ""
                Do While Not EOF(1)
                Line Input #1, a
                txtText.Text = txtText.Text & a & vbCrLf
                Loop
            Close
        End If
    End If
End Sub

Private Sub Form_Load()
Cb.AddItem "1.프로그램 제거 밑 주의사항", "0"
Cb.AddItem "2.기본 명령어", "1"
Cb.AddItem "3.시스템 명령어", "2"
Cb.AddItem "4.문자열 명령어", "3"
Cb.AddItem "5.TCP/IP 명령어", "4"
Cb.ListIndex = "0"
End Sub

Private Sub Form_Resize()
    If Not Me.WindowState = 1 Then
        leftpb.Height = Me.ScaleHeight
        Cb.Width = Me.ScaleWidth - leftpb.Width
        txtText.Width = Me.ScaleWidth - leftpb.Width
        txtText.Height = Me.ScaleHeight - Cb.Height
    End If
End Sub

