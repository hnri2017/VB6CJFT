VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInterval 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ʱ��ѡ��"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   8475
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "ÿ�μ��"
      Height          =   1215
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.OptionButton Option1 
         Caption         =   "��"
         Height          =   255
         Index           =   0
         Left            =   5200
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ÿN��"
         Height          =   255
         Index           =   5
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ÿ��"
         Height          =   255
         Index           =   4
         Left            =   3200
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ÿ��"
         Height          =   255
         Index           =   3
         Left            =   2200
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ÿ��"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ÿ��"
         Height          =   255
         Index           =   1
         Left            =   200
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   400
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   714
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   102891522
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
    'ȷ��
    Dim K As Long
    
    If MsgBox("ȷ�����浱ǰ������", vbQuestion + vbOKCancel, "����") = vbCancel Then Exit Sub
    
    For K = Me.Option1.LBound To Me.Option1.UBound
        If Me.Option1.Item(K).Value Then
            With gVar
                .ParaBackupInterval = K
                If K = 5 Then
                    .ParaBackupTime = Format(Date, "yyyy-MM-dd ") & Format(Me.DTPicker1.Value, "HH:mm:ss")
                    .ParaBackupIntervalDays = Me.DTPicker1.Day
                Else
                    .ParaBackupTime = Format(Me.DTPicker1.Value, "yyyy-MM-dd HH:mm:ss")
                End If
                Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyServerBackInterval, .ParaBackupInterval) '����Ƶ��
                Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyServerBackTime, .ParaBackupTime) '����ȷ��ʱ��
                Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyServerBackIntervalDays, .ParaBackupIntervalDays) 'ÿN��
            End With
            
            Dim frmOP As Form
            For Each frmOP In Forms
                If LCase(frmOP.Name) = LCase(gWind.CommandBars1.Actions(gID.toolOptions).Key) Then
                    frmOP.Command1.Value = True '���¼��ز���
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
End Sub

Private Sub Form_Load()
    Me.DTPicker1.Format = dtpCustom
    Me.DTPicker1.UpDown = True
    Me.DTPicker1.Value = gVar.ParaBackupTime ' Date & " 19:00:00"
    If gVar.ParaBackupInterval = 5 Then
        Dim dateTemp As Date
        If Me.DTPicker1.Day <> gVar.ParaBackupIntervalDays Then
            Me.DTPicker1.Day = gVar.ParaBackupIntervalDays
        End If
    End If
    Me.Option1.Item(gVar.ParaBackupInterval).Value = True
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
            Me.DTPicker1.CustomFormat = "d�� HH:mm:ss"
        Case 4
            Me.DTPicker1.CustomFormat = "M��d�� HH:mm:ss"
        Case 5
            Me.DTPicker1.CustomFormat = "ÿd�� HH:mm:ss"
        Case Else
            MsgBox "δ���嵥ѡ��ť"
    End Select
End Sub
