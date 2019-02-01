VERSION 5.00
Begin VB.Form frmAnalys 
   Caption         =   "분석 도구"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13995
   Icon            =   "frmAnalys.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   13995
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox ptop 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   0
      ScaleHeight     =   4575
      ScaleWidth      =   11760
      TabIndex        =   0
      Top             =   0
      Width           =   11760
      Begin VB.ListBox lstAnalys 
         Appearance      =   0  '평면
         Height          =   750
         Left            =   105
         TabIndex        =   1
         Top             =   0
         Width           =   11535
      End
      Begin VB.Image pleft 
         Height          =   4575
         Left            =   0
         Picture         =   "frmAnalys.frx":2DA2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   105
      End
      Begin VB.Image pright 
         Height          =   4575
         Left            =   11640
         Picture         =   "frmAnalys.frx":2E14
         Stretch         =   -1  'True
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Image picbag 
      Height          =   495
      Left            =   0
      Picture         =   "frmAnalys.frx":2E86
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmAnalys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
On Error Resume Next
    If Not Me.WindowState = 1 Then
        ptop.Top = 0
        pleft.Height = Me.ScaleHeight
        pright.Height = Me.ScaleHeight
        ptop.Left = (Me.ScaleWidth - ptop.Width) / 2
        ptop.Height = Me.ScaleWidth
        picbag.Width = Me.ScaleWidth
        picbag.Height = Me.ScaleHeight
        lstAnalys.Height = pleft.Height
        lstAnalys.Top = (Me.ScaleHeight - lstAnalys.Height) / 2
    End If
End Sub
