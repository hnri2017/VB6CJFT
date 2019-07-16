VERSION 5.00
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmErrorLog 
   Caption         =   "������־�鿴"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14745
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   14745
   StartUpPosition =   1  '����������
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   7500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   300
      Left            =   9000
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9000
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FlexCell.Grid Grid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   555
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5953
      Cols            =   5
      GridColor       =   12632256
      Rows            =   30
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��־�ļ�·����"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   150
      Width           =   1260
   End
   Begin VB.Label Label2 
      ForeColor       =   &H00FF00FF&
      Height          =   180
      Left            =   10080
      TabIndex        =   3
      Top             =   180
      Width           =   3180
   End
   Begin VB.Menu mnuExport 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuExportExcel 
         Caption         =   "������Excel"
      End
   End
End
Attribute VB_Name = "frmErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrFile As String  '��־�ļ�·��
Private Const mconRows As Long = 50 '�����С����

Private Sub mGridSet()
    With Me.Grid1
        .AutoRedraw = False
        .Appearance = 0
        .FixedCols = 1
        .FixedRows = 1
        .Cols = 5
        .Rows = mconRows + 1
        .BackColorBkg = Me.BackColor
        .BackColorFixed = RGB(121, 151, 219)
        .BackColor2 = RGB(250, 235, 215)
        .BackColorFixedSel = vbYellow
        .DisplayRowIndex = True
        .AllowUserResizing = True
        .AllowUserSort = True
        .ExtendLastCol = True
        
        .Cell(0, 0).Text = "���"
        .Cell(0, 1).Text = "�쳣��¼ʱ��"
        .Cell(0, 2).Text = "�쳣����"
        .Cell(0, 3).Text = "�쳣����"
        .Cell(0, 4).Text = "�쳣����"
        .Range(0, 0, 0, .Cols - 1).WrapText = True
        .Range(0, 0, 0, .Cols - 1).FontBold = True
        
        .RowHeight(0) = 40
        .Column(0).Width = 50
        .Column(1).Width = 150
        .Column(2).Width = 160
        .Column(3).Width = 110
        .Column(4).Width = 300
        .Column(1).Alignment = cellCenterCenter
        
        .AutoRedraw = True
        .Refresh
    End With
End Sub

Private Sub mOpenLog()
    Dim intNum As Integer
    Dim strLine As String, arrStr() As String, strSep As String
    Dim L As Long, U As Long, K As Long, Rs As Long, Cs As Long
    Dim sngTime As Single
    
    On Error Resume Next
    
    If Not gfFileExist(mstrFile) Then Exit Sub
    If FileLen(mstrFile) = 0 Then Exit Sub
    
    intNum = FreeFile
    strSep = vbTab & vbTab
    sngTime = Timer
    Me.MousePointer = 13
    
    Open mstrFile For Input As #intNum
    With Me.Grid1
        .AutoRedraw = False
        While Not EOF(intNum)
            Rs = Rs + 1
            Line Input #intNum, strLine
            arrStr = Split(strLine, strSep)
            L = LBound(arrStr)
            U = UBound(arrStr)
            Cs = U - L + 2
            If .Cols < Cs Then .Cols = Cs
            If .Rows < Rs + 1 Then .Rows = Rs + 1
            For K = L To U
                .Cell(Rs, K + 1).Text = arrStr(K)
            Next
        Wend
        If Rs <= mconRows Then
            .Rows = mconRows + 1
            If Rs < mconRows Then .Range(Rs + 1, 1, mconRows, .Cols - 1).ClearText
        Else
            .Rows = Rs + 1
        End If
        .AutoRedraw = True
        .Refresh
    End With
    
    Close #intNum
    Me.Label2.Caption = "��ʱ" & Format(Timer - sngTime, "0.000") & "��"
    Me.Text1.Text = mstrFile
    Me.MousePointer = 0
    
    If Err.Number Then
        Call gsAlarmAndLog("������־��ȡ�쳣")
    End If
End Sub

Private Sub Command1_Click()
    Dim strFile As String, strPrefix As String, strExtension As String
    Dim strOpen As String, blnOpen As Boolean
    
    Me.Label2.Caption = "��ʱ��"
    strPrefix = Mid(gVar.FileNameErrLog, InStrRev(gVar.FileNameErrLog, "\") + 1, InStrRev(gVar.FileNameErrLog, ".") - InStrRev(gVar.FileNameErrLog, "\") - 1)
    strExtension = Mid(gVar.FileNameErrLog, InStrRev(gVar.FileNameErrLog, "."))
    With Me.CommonDialog1
        .DialogTitle = "ѡ����־�ļ�"
        .Filter = "��־(" & strExtension & ")|" & strPrefix & "*" & strExtension
        .Flags = cdlOFNFileMustExist
        .InitDir = gVar.FolderData
        .ShowOpen
        strFile = .FileName
    End With
    
    If Len(strFile) > 0 Then
        strOpen = Mid(strFile, InStrRev(strFile, "\") + 1)
        If LCase(Right(strOpen, 4)) = LCase(strExtension) Then
            If LCase(Left(strOpen, Len(strPrefix))) = LCase(strPrefix) Then
                If gfFileExist(strFile) Then
                    mstrFile = strFile
                    Call mOpenLog
                    blnOpen = True
                End If
            End If
        End If
        If Not blnOpen Then
            MsgBox "��ѡ��־�ļ������ϴ�Ҫ��", vbExclamation, "����"
        End If
    End If
End Sub

Private Sub Form_Load()
    Set Me.Icon = gWind.Icon
    Call mGridSet
    mstrFile = gVar.FileNameErrLog
    Call mOpenLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.Grid1.Move 0, 550, Me.ScaleWidth, Me.ScaleHeight - Me.Grid1.Top
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    KeyCode = 0 '���ΰ���
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 3 Then KeyAscii = 0  '����Ctrl+C����������
End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuExport
    End If
End Sub

Private Sub mnuExportExcel_Click()
    Dim strFile As String
    
    If gfFileRepair(gVar.FolderNameTemp, True) Then
       strFile = gVar.FolderNameTemp & "ErrLog" & Format(Now, gVar.Formatymdhms) & ".xls"
        Me.Grid1.ExportToExcel strFile
        Call gfFileOpen(strFile)
    End If
End Sub

