VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmFileRestore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "服务端文件还原"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8415
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   3240
      Width           =   1000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "开始还原"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   1000
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1965
      Width           =   6335
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   6335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "选择文件"
      Height          =   375
      Left            =   7280
      TabIndex        =   1
      Top             =   180
      Width           =   1000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "选择位置"
      Height          =   375
      Left            =   7280
      TabIndex        =   4
      Top             =   1920
      Width           =   1000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7320
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "还原位置"
      Height          =   180
      Left            =   200
      TabIndex        =   7
      Top             =   2010
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "还原文件"
      Height          =   180
      Left            =   200
      TabIndex        =   3
      Top             =   285
      Width           =   720
   End
End
Attribute VB_Name = "frmFileRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim strFile As String
    
    With Me.CommonDialog1
        .DialogTitle = "选择还原的源文件"
        .Filter = "备份文件(.bak)|*.bak"
        .Flags = cdlOFNFileMustExist
        .InitDir = gVar.ParaBackupPath
        .ShowOpen
        strFile = Trim(.FileName)
    End With
    If Len(strFile) > 0 Then
        Me.Text1.Text = strFile
    End If
End Sub

Private Sub Command2_Click()
    Dim strFolder As String
    
    strFolder = Trim(BrowseForFolder(Me, Me.Text2.Text))
    If Len(strFolder) > 0 Then
        Me.Text2.Text = strFolder
    End If
End Sub

Private Sub Command3_Click()
    '还原
    Dim strFile As String, strFolder As String
    
    strFile = Me.Text1.Text
    strFolder = Me.Text2.Text
    
    If Len(strFile) = 0 Then
        MsgBox "请先选择一个还原的源文件！", vbExclamation, "提醒"
        Exit Sub
    End If
    If Not FileExist(strFile) Then
        MsgBox "还原的源文件不存在！", vbExclamation, "提醒"
        Exit Sub
    End If
    
    If Len(strFolder) = 0 Then
        MsgBox "请先选择一个还原的位置！", vbExclamation, "提醒"
        Exit Sub
    End If
    If Not FolderExist(strFolder, True) Then
        MsgBox "还原的位置不存在！", vbExclamation, "提醒"
        Exit Sub
    End If
    
    If MsgBox("是否立即进行还原？", vbQuestion + vbYesNo, "询问") = vbNo Then Exit Sub
    
    If FileRestoreCP(strFile, strFolder) Then
        MsgBox "还原成功！", vbInformation, "提示"
    Else
        MsgBox "还原过程异常，还原失败！", vbCritical, "警告"
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Text2.Text = gVar.FolderNameStore
End Sub
