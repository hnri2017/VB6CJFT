VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmBK 
   Caption         =   "�ļ�����"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   10050
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command5 
      Caption         =   "�����ԭ�ļ�"
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�����ԭλ��"
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2050
      Width           =   6255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   370
      Width           =   6255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ԭ�ļ�λ��"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����ļ�λ��"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmBK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrBackup As String
Dim mstrRestore As String
Dim mstrDtBK As String
Dim mstrDtRS As String

Private Sub Command1_Click()
    '����
    
    Call EnabledControl(Me, False)
    If Not BackupFile(mstrBackup, mstrRestore) Then
        MsgBox "����ʧ��", vbCritical, "����"
    End If
    Call EnabledControl(Me, True)
    
End Sub

Private Sub Command2_Click()
    '��ԭ
    Call EnabledControl(Me, False)
    If Not RestoreFile(Me.Text2.Text, mstrBackup) Then
        MsgBox "��ԭʧ��", vbCritical, "����"
    End If
    Call EnabledControl(Me, True)
End Sub

Private Sub Command3_Click()
    mstrBackup = BrowseForFolder(Me, Me.Text1.Text)
    Me.Text1.Text = mstrBackup
End Sub

Private Sub Command4_Click()
    mstrRestore = BrowseForFolder(Me, Me.Text2.Text)
    Me.Text2.Text = mstrRestore
End Sub

Private Sub Command5_Click()
    With Me.CommonDialog1
        .Filter = "�����ļ�(.bak)|*.bak"
        .Flags = cdlOFNFileMustExist
        .ShowOpen
        Me.Text2.Text = .FileName
    End With
End Sub

Private Sub Form_Load()
    mstrDtBK = App.Path & "\store"
    mstrDtRS = App.Path & "\data"
End Sub
