VERSION 5.00
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmLoginLog 
   Caption         =   "用户登陆日志查看"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   15885
   StartUpPosition =   1  '所有者中心
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9000
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "浏览"
      Height          =   300
      Left            =   9000
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   7500
   End
   Begin FlexCell.Grid Grid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   550
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5953
      Cols            =   5
      GridColor       =   12632256
      Rows            =   30
   End
   Begin VB.Label Label2 
      Caption         =   "用时123.321秒"
      ForeColor       =   &H00FF00FF&
      Height          =   180
      Left            =   10080
      TabIndex        =   4
      Top             =   180
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日志文件路径："
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   150
      Width           =   1260
   End
   Begin VB.Menu mnuExport 
      Caption         =   "导出"
      Visible         =   0   'False
      Begin VB.Menu mnuExportExcel 
         Caption         =   "导出至Excel"
      End
   End
End
Attribute VB_Name = "frmLoginLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrFile As String  '日志文件路径
Private Const mconRows As Long = 50 '表格最小行数

Private Sub mGridSet()
    With Me.Grid1
        .AutoRedraw = False
        .Appearance = 0
        .FixedCols = 1
        .FixedRows = 1
        .Cols = 8
        .Rows = mconRows + 1
        .BackColorBkg = Me.BackColor
        .BackColorFixed = RGB(121, 151, 219)
        .BackColor2 = RGB(250, 235, 215)
        .BackColorFixedSel = vbYellow
        .DisplayRowIndex = True
        .AllowUserResizing = True
        .AllowUserSort = True
        .ExtendLastCol = True
        
        .Cell(0, 0).Text = "序号"
        .Cell(0, 1).Text = "连接用户IP地址"
        .Cell(0, 2).Text = "连接用户计算机名称"
        .Cell(0, 3).Text = "连接用户登陆账号"
        .Cell(0, 4).Text = "连接用户姓名"
        .Cell(0, 5).Text = "连接建立时间"
        .Cell(0, 6).Text = "索引号"
        .Cell(0, 7).Text = "申请号"
        .Range(0, 0, 0, .Cols - 1).WrapText = True
        .Range(0, 0, 0, .Cols - 1).FontBold = True
        
        .RowHeight(0) = 40
        .Column(0).Width = 50
        .Column(1).Width = 120
        .Column(2).Width = 130
        .Column(3).Width = 130
        .Column(4).Width = 120
        .Column(5).Width = 130
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        
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
    Me.Label2.Caption = "用时" & Format(Timer - sngTime, "0.000") & "秒"
    Me.Text1.Text = mstrFile
    Me.MousePointer = 0
    
    If Err.Number Then
        Call gsAlarmAndLog("登陆日志读取异常")
    End If
End Sub

Private Sub Command1_Click()
    Dim strFile As String, strPrefix As String, strExtension As String
    Dim strOpen As String, blnOpen As Boolean
    
    Me.Label2.Caption = "用时…"
    strPrefix = Mid(gVar.FileNameLoginLog, InStrRev(gVar.FileNameLoginLog, "\") + 1, InStrRev(gVar.FileNameLoginLog, ".") - InStrRev(gVar.FileNameLoginLog, "\") - 1)
    strExtension = Mid(gVar.FileNameLoginLog, InStrRev(gVar.FileNameLoginLog, "."))
    With Me.CommonDialog1
        .DialogTitle = "选择日志文件"
        .Filter = "日志(" & strExtension & ")|" & strPrefix & "*" & strExtension
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
            MsgBox "所选日志文件不符合打开要求！", vbExclamation, "警告"
        End If
    End If
End Sub

Private Sub Form_Load()
    Set Me.Icon = gWind.Icon
    Call mGridSet
    mstrFile = gVar.FileNameLoginLog
    Call mOpenLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.Grid1.Move 0, 550, Me.ScaleWidth, Me.ScaleHeight - Me.Grid1.Top
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    KeyCode = 0 '屏蔽按键
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 3 Then KeyAscii = 0  '除了Ctrl+C，其余屏蔽
End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuExport
    End If
End Sub

Private Sub mnuExportExcel_Click()
    Dim strFile As String
    
    If gfFileRepair(gVar.FolderNameTemp, True) Then
       strFile = gVar.FolderNameTemp & "LoginLog" & Format(Now, gVar.Formatymdhms) & ".xls"
        Me.Grid1.ExportToExcel strFile
        Call gfFileOpen(strFile)
    End If
End Sub
