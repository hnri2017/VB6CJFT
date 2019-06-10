VERSION 5.00
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSysFile 
   Caption         =   "�ļ�����"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   9945
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   120
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�ϴ�"
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   6255
   End
   Begin FlexCell.Grid Grid1 
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7646
      Cols            =   5
      GridColor       =   12632256
      Rows            =   30
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�ļ�·����"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   300
      Width           =   900
   End
End
Attribute VB_Name = "frmSysFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub msLoadFileList(Optional ByVal blnLD As Boolean = True)
    '�����ļ���Ϣ�����
    Dim strSQL As String
    Dim rsFile As ADODB.Recordset
    Dim K As Long, C As Long
    
    strSQL = "SELECT FileID ,FileClassify ,FileExtension ,FileOldName ,FileSaveName ," & _
             "FileSize ,FileSaveLocation ,FileUploadMen ,FileUploadTime FROM tb_FT_Lib_File "
    Set rsFile = gfBackRecordset(strSQL)
    If rsFile.State = adStateOpen Then
        C = rsFile.RecordCount
        If C > 0 Then
            With Grid1
                .AutoRedraw = False
                .Rows = C + 1
                If C < 20 Then .Rows = 21
                While Not rsFile.EOF
                    K = K + 1
                    Grid1.Cell(K, 1).Text = rsFile.Fields("FileID") & ""
                    Grid1.Cell(K, 2).Text = rsFile.Fields("FileSaveName") & ""
                    Grid1.Cell(K, 3).Text = rsFile.Fields("FileSaveLocation") & ""
                    Grid1.Cell(K, 4).Text = ""
                    Grid1.Cell(K, 5).Text = rsFile.Fields("FileClassify") & ""
                    Grid1.Cell(K, 6).Text = rsFile.Fields("FileExtension") & ""
                    Grid1.Cell(K, 7).Text = rsFile.Fields("FileSize") & ""
                    Grid1.Cell(K, 8).Text = rsFile.Fields("FileUploadMen") & ""
                    Grid1.Cell(K, 9).Text = rsFile.Fields("FileUploadTime") & ""
                    Grid1.Cell(K, 10).Text = "��"
                    Grid1.Cell(K, 11).Text = rsFile.Fields("FileOldName") & ""
                    rsFile.MoveNext
                Wend
                .Range(1, 10, K, 10).ForeColor = vbBlue
                .ReadOnly = True
                .AutoRedraw = True
                .Refresh
            End With
        End If
        rsFile.Close
    End If
    Set rsFile = Nothing
End Sub

Private Sub Command1_Click()
    '���
    
    With CommonDialog1
        .DialogTitle = "ѡ��һ���ļ�"
        .Flags = cdlOFNFileMustExist
        .ShowOpen
        Text1.Text = .FileName
    End With
End Sub

Private Sub Command2_Click()
    '�ϴ�
    Const conLngSize As Long = 524288000 '500MB=500*1024*1024=524288000(B)
    Dim strFilePath As String, strFileName As String, strExtension As String
    Dim strSaveName As String, strSaveLocation As String, strSQL As String
    Dim lngSize As Long, K As Long
    Dim rsFile As ADODB.Recordset

    strFilePath = Trim(Text1.Text)
    If Len(strFilePath) = 0 Then
        MsgBox "����ѡ��һ���ļ���", vbExclamation, "��ʾ"
        Exit Sub
    End If
    
    If MsgBox("ȷ��Ҫ�ϴ���ѡ�ļ���", vbQuestion + vbOKCancel, "����") = vbCancel Then Exit Sub
    
    If Not gfFileExist(strFilePath) Then
        MsgBox "�ļ������ڣ���ȷ�ϻ�����ѡ��", vbCritical, "����"
        Exit Sub
    End If
    
    lngSize = FileLen(strFilePath)  '��ȡ�ļ���С����λ�ֽ�
    If lngSize > conLngSize Then
        MsgBox "��ѡ�ļ���С���ܳ���500MB��", vbCritical, "����"
        Exit Sub
    End If
    
    strFileName = Mid(strFilePath, InStrRev(strFilePath, "\") + 1)  '��ȡ����·�����ļ���
    strExtension = Mid(strFilePath, InStrRev(strFilePath, ".") + 1) '��ȡ�ļ�����չ��
    For K = 1 To 30
        strSaveName = strSaveName & gfBackOneChar(udUpperCase) '�����ļ��ڷ���˱����õ��ļ�����30�������ĸ
    Next
    strSaveLocation = gVar.FolderStore  '�����ļ��ڷ���˵Ĵ洢λ�á�ע�ⲻ��·��
    
    
End Sub

Private Sub Form_Load()
    '���ڼ��أ��������
    
    Text1.Text = ""
    Text1.Locked = True
    Text1.Font.Size = 11
    With Grid1
        .AutoRedraw = False
        .Rows = 16
        .Cols = 12
        .Cell(0, 0).Text = "���"
        .Cell(0, 1).Text = "�ļ�ID"
        .Cell(0, 2).Text = "�洢����"
        .Cell(0, 3).Text = "�洢λ��"
        .Cell(0, 4).Text = "����λ��"
        .Cell(0, 5).Text = "�ļ�����"
        .Cell(0, 6).Text = "��չ��"
        .Cell(0, 7).Text = "�ļ���С"
        .Cell(0, 8).Text = "�ϴ���"
        .Cell(0, 9).Text = "�ϴ�����"
        .Cell(0, 10).Text = "�鿴"
        .Cell(0, 11).Text = "�ļ���"
        .Cell(0, 7).Comment = "��λ���ֽ�(B)"
        .Column(0).Width = 40
        .Column(1).Width = 0
        .Column(2).Width = 0
        .Column(3).Width = 0
        .Column(4).Width = 0
        .Column(5).Width = 100
        .Column(6).Width = 50
        .Column(7).Width = 70
        .Column(8).Width = 50
        .Column(9).Width = 120
        .Column(10).Width = 50
        .ExtendLastCol = True
        .rowHeight(0) = 30
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellRightCenter
        .Column(8).Alignment = cellCenterCenter
        .Column(9).Alignment = cellCenterCenter
        .Column(10).Alignment = cellCenterCenter
        .Column(10).CellType = cellHyperLink
        .Column(9).FormatString = gVar.Formaty_M_dH_m_s
        .DisplayRowIndex = True
        .AllowUserSort = True
        .AllowUserResizing = True
        .Appearance = Flat
        .BackColorBkg = Me.BackColor
        .BackColorFixed = RGB(121, 151, 219)
        .BackColor2 = RGB(250, 235, 215)
        .BackColorFixedSel = vbYellow
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
    End With
    Call msLoadFileList(True)
End Sub

Private Sub Form_Resize()
    '���ڴ�С�ĸı�
    Const conLngW As Long = 10650 + 720
    Const conLngH As Long = 5000 + 720
    Dim lngW As Long, lngH As Long
    
    On Error Resume Next
    
    lngW = Me.Width
    lngH = Me.Height
    If lngW > conLngW Then
        Grid1.Width = lngW - 900
    Else
        Grid1.Width = 10000
    End If
    If lngH > conLngH Then
        Grid1.Height = lngH - 1600
    Else
        Grid1.Height = 5000
    End If
End Sub

Private Sub Grid1_HyperLinkClick(ByVal Row As Long, ByVal Col As Long, URL As String, Changed As Boolean)
    '�鿴�ļ�
    
    URL = ""
    Changed = True
    If Row > 0 And Col = 10 Then
        Debug.Print Grid1.Cell(Row, 11).Text, Grid1.Cell(Row, 4).Text
    End If
End Sub
