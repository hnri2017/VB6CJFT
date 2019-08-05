VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInterval 
   Caption         =   "时间选择"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   9555
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   495
      Left            =   7200
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "每次间隔"
      Height          =   1215
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.OptionButton Option1 
         Caption         =   "无"
         Height          =   255
         Index           =   0
         Left            =   5200
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "每N天"
         Height          =   255
         Index           =   5
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "每年"
         Height          =   255
         Index           =   4
         Left            =   3200
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "每月"
         Height          =   255
         Index           =   3
         Left            =   2200
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "每周"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "每天"
         Height          =   255
         Index           =   1
         Left            =   200
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   100532226
         CurrentDate     =   43680.8125
      End
   End
End
Attribute VB_Name = "frmInterval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox Format(Me.DTPicker1.Value, "yyyy-MM-dd dddd HH:mm:ss")
End Sub

Private Sub Form_Load()
    Me.DTPicker1.Format = dtpCustom
    Me.DTPicker1.Value = "19:00:00"
    Me.DTPicker1.UpDown = True
End Sub

Private Sub Option1_Click(Index As Integer)
    Me.DTPicker1.Enabled = IIf(Index = 0, False, True)
    Me.DTPicker1.UpDown = IIf(Index = 1, True, False)
    Select Case Index
        Case 0
            Me.DTPicker1.CustomFormat = ""
        Case 1
            Me.DTPicker1.CustomFormat = "HH:mm:ss"
        Case 2
            Me.DTPicker1.CustomFormat = "dddd HH:mm:ss"
        Case 3
            Me.DTPicker1.CustomFormat = "d日 HH:mm:ss"
        Case 4
            Me.DTPicker1.CustomFormat = "M月d日 HH:mm:ss"
        Case 5
            Me.DTPicker1.CustomFormat = "每d天 HH:mm:ss"
        Case Else
            MsgBox "未定义单选按钮"
    End Select
End Sub
