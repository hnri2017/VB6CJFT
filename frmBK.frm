VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmBK 
   Caption         =   "�ļ�����"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   11070
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command10 
      Caption         =   "Test"
      Height          =   495
      Left            =   5160
      TabIndex        =   15
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "���"
      Height          =   375
      Left            =   8400
      TabIndex        =   14
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "���"
      Height          =   375
      Left            =   8400
      TabIndex        =   13
      Top             =   3180
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ѹ��"
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   3645
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��ѹ"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   1080
      TabIndex        =   8
      Top             =   3240
      Width           =   7335
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1080
      TabIndex        =   7
      Top             =   4965
      Width           =   7335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�����ԭ�ļ�"
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�����ԭλ��"
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1455
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
      Top             =   1440
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ѹ��ԴĿ¼"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   3285
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��ѹԴ�ļ�"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   5010
      Width           =   900
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
Dim mstrData As String
Dim mstrStore As String

Private Sub Command1_Click()
    '����
    
    Call EnabledControl(Me, False)
    If Not FilePackage(mstrBackup, mstrRestore) Then
        MsgBox "����ʧ��", vbCritical, "����"
    End If
    Call EnabledControl(Me, True)
End Sub

Private Sub Command10_Click()
'    MsgBox DriveFreeSpace(App.Path & "\ffc.exe")
'    MsgBox DriveTotalSize(App.Path & "\ffc.exe")
'    MsgBox FolderSize(App.Path)
'    MsgBox FolderPathBuild("e:\a\b\c\")
'    MsgBox FolderPathBuild("\\192.168.12.100\��֮��\��������\Ф�λ�\aa\bb\")
'    MsgBox FolderPathBuild("\\192.168.12.120\��֮��\��������\Ф�λ�\aa\bb\")
    MsgBox DriveLetter("\\192.168.12.100\��֮��\��������\Ф�λ�")
    MsgBox DriveLetter("c:")
End Sub

Private Sub Command2_Click()
    '��ԭ
    Call EnabledControl(Me, False)
    If Not FileUnpack(Me.Text2.Text, Me.Text1.Text) Then
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

Private Sub Command6_Click()
    If FileExtract(Me.Text3.Text, Me.Text4.Text) Then
        MsgBox "��ѹ�ļ����", vbInformation, "��ʾ"
    Else
        MsgBox "��ѹ�ļ�ʧ��", vbExclamation, "����"
    End If
End Sub

Private Sub Command7_Click()
    If FileCompress(Me.Text4.Text, Me.Text3.Text, 0) Then
        MsgBox "�ļ�ѹ�����", vbInformation, "��ʾ"
    Else
        MsgBox "�ļ�ѹ��ʧ��", vbExclamation, "����"
    End If
End Sub

Private Sub Command8_Click()
    Me.Text4.Text = BrowseForFolder(Me, Me.Text4.Text)
End Sub

Private Sub Command9_Click()
    With Me.CommonDialog1
        .Filter = "ѹ���ļ�(.rar)|*.rar"
        .Flags = cdlOFNFileMustExist
        .ShowOpen
        Me.Text3.Text = .FileName
    End With
End Sub

Private Sub Form_Load()
    mstrBackup = App.Path & "\store"
    mstrRestore = App.Path & "\data"
    mstrData = App.Path & "\data"
    mstrStore = App.Path & "\store"
    Me.Text1.Text = mstrStore
    Me.Text2.Text = mstrData
    Me.Text3.Text = mstrData
    Me.Text4.Text = mstrStore
End Sub
