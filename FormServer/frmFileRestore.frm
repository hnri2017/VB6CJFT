VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmFileRestore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������ļ���ԭ"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8415
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command4 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   3240
      Width           =   1000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ʼ��ԭ"
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
      Caption         =   "ѡ���ļ�"
      Height          =   375
      Left            =   7280
      TabIndex        =   1
      Top             =   180
      Width           =   1000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ѡ��λ��"
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
      Caption         =   "��ԭλ��"
      Height          =   180
      Left            =   200
      TabIndex        =   7
      Top             =   2010
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ԭ�ļ�"
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
        .DialogTitle = "ѡ��ԭ��Դ�ļ�"
        .Filter = "�����ļ�(.bak)|*.bak"
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
    '��ԭ
    Dim strFile As String, strFolder As String
    
    strFile = Me.Text1.Text
    strFolder = Me.Text2.Text
    
    If Len(strFile) = 0 Then
        MsgBox "����ѡ��һ����ԭ��Դ�ļ���", vbExclamation, "����"
        Exit Sub
    End If
    If Not FileExist(strFile) Then
        MsgBox "��ԭ��Դ�ļ������ڣ�", vbExclamation, "����"
        Exit Sub
    End If
    
    If Len(strFolder) = 0 Then
        MsgBox "����ѡ��һ����ԭ��λ�ã�", vbExclamation, "����"
        Exit Sub
    End If
    If Not FolderExist(strFolder, True) Then
        MsgBox "��ԭ��λ�ò����ڣ�", vbExclamation, "����"
        Exit Sub
    End If
    
    If MsgBox("�Ƿ��������л�ԭ��", vbQuestion + vbYesNo, "ѯ��") = vbNo Then Exit Sub
    
    If FileRestoreCP(strFile, strFolder) Then
        MsgBox "��ԭ�ɹ���", vbInformation, "��ʾ"
    Else
        MsgBox "��ԭ�����쳣����ԭʧ�ܣ�", vbCritical, "����"
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Text2.Text = gVar.FolderNameStore
End Sub
