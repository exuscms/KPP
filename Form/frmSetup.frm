VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetup 
   BorderStyle     =   1  '단일 고정
   Caption         =   "설정"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3255
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   3255
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   240
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   600
      Width           =   2775
      Begin VB.TextBox txtinput 
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CheckBox ck6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "입력에 위의 값 삽입"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   7215
      End
      Begin VB.CheckBox ck5 
         BackColor       =   &H00FFFFFF&
         Caption         =   """~ ""연산 나타내기"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   7215
      End
      Begin VB.CheckBox ck4 
         BackColor       =   &H00FFFFFF&
         Caption         =   """= ""연산 나타내기"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   7215
      End
      Begin VB.CheckBox ck3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "주석 나타내기"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   7215
      End
      Begin VB.CheckBox ck2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "산수연산 변화 나타내기"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   7215
      End
      Begin VB.CheckBox ck1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "값 동일화 변화 나타내기"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   7215
      End
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5953
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "분석도구"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
