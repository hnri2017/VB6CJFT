VERSION 5.00
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选项"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "UnVisible"
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin FlexCell.Grid Grid1 
      Height          =   3255
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5741
      Cols            =   5
      GridColor       =   12632256
      Rows            =   30
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mconstrTip As String = "选择一个文件夹" '选择一个文件夹点击打开即可


Private Function mfCheckFolder(ByVal strFolder As String) As String
    '对选择的文件夹名检查
    Dim strCheck As String, strEnd As String
    
    strCheck = Trim(strFolder)
    If Len(strCheck) = 0 Then
        strCheck = gVar.FolderNameBackup
    ElseIf LCase(strFolder) = LCase(mconstrTip) Then    '未选择路径，点了取消
        strCheck = ""
    Else
        strEnd = Mid(strCheck, InStrRev(strCheck, "\") + 1)
        If LCase(strEnd) = LCase(mconstrTip) Then
            strCheck = Left(strCheck, InStrRev(strCheck, "\"))
        End If
        If Not gfFolderRepair(strCheck) Then
            strCheck = gVar.FolderNameBackup
        End If
    End If
    mfCheckFolder = strCheck
End Function

Private Sub msLoadParameter(Optional ByVal blnLoad As Boolean = True)
    Dim lngRow As Long  '表格行号记录
    
    If Not blnLoad Then Exit Sub
    
    '从公共变量或注册表中加载配置信息
    With Me.Grid1
        '窗口控制参数
        lngRow = 2
        .Cell(lngRow, 1).Text = gVar.ParaBlnWindowCloseMin   '关闭时最小化
        .Cell(lngRow, 5).Text = gVar.ParaBlnWindowMinHide    '最小化时隐藏
        .Cell(lngRow + 1, 1).Text = gVar.ParaBlnWindowStartMinS '启动时最小化
        
        '服务端参数
        lngRow = lngRow + 4
        .Cell(lngRow, 3).Text = gVar.TCPSetPort  '侦听端口
        .Cell(lngRow, 5).Text = gVar.ParaBlnAutoStartupAtBoot   '开机自动启动
        
        '数据库服务器参数
        lngRow = lngRow + 4
        .Cell(lngRow, 3).Text = gVar.ConSource   '服务器名称/IP
        .Cell(lngRow, 7).Text = gVar.ConDatabase '数据库名
        .Cell(lngRow + 2, 3).Text = gVar.ConUserID '登陆名
        Text1.Text = gVar.ConPassword       '登陆密码
        .Cell(lngRow + 2, 7).Text = String(Len(gVar.ConPassword), "*") '登陆密码*号显示
        
        '客户端控制参数
        lngRow = lngRow + 5
        .Cell(lngRow, 1).Text = gVar.ParaBlnLimitClientConnect '限制客户端连接
        .Cell(lngRow, 7).Text = gVar.ParaLimitClientConnectTime '限制客户端连接时长
        .Cell(lngRow + 1, 3).Text = gVar.TCPConnectMax '限制客户端连接数
        
        '服务端文件备份参数
        lngRow = lngRow + 4
        .Cell(lngRow, 3).Text = gVar.ParaBackupPath '备份路径
        .Cell(lngRow + 1, 3).Text = ShowBackupTimeInfo(gVar.ParaBackupInterval, gVar.ParaBackupTime) & _
                 "，下次备份时间:" & ShowBackupNextTime(gVar.ParaBackupInterval, gVar.ParaBackupTime)
    End With
End Sub

Private Sub msSaveParameter(Optional ByVal blnSave As Boolean = True)
    Dim lngRow As Long  '表格行号记录
    Dim tempVal
    
    If Not blnSave Then Exit Sub
    
    '参数值更新至公共变量
    With Grid1
        '窗口控制参数
        lngRow = 2
        gVar.ParaBlnWindowCloseMin = .Cell(lngRow, 1).Text   '关闭时最小化
        gVar.ParaBlnWindowMinHide = .Cell(lngRow, 5).Text    '最小化时隐藏
        gVar.ParaBlnWindowStartMinS = .Cell(lngRow + 1, 1).Text  '启动时最小化
        
        '服务端参数
        lngRow = lngRow + 4
        tempVal = Val(.Cell(lngRow, 3).Text)                 '侦听端口
        gVar.TCPSetPort = IIf(tempVal < 10000, gVar.TCPDefaultPort, tempVal)
        gVar.ParaBlnAutoStartupAtBoot = .Cell(lngRow, 5).Text    '开机自动启动
        
        '数据库服务器参数
        lngRow = lngRow + 4
        gVar.ConSource = gfCheckIP(Trim(.Cell(lngRow, 3).Text))    '服务器名称/IP
        gVar.ConDatabase = Trim(.Cell(lngRow, 7).Text)   '数据库名
        gVar.ConUserID = Trim(.Cell(lngRow + 2, 3).Text)  '登陆名
        gVar.ConPassword = Text1.Text               '登陆密码
        
        '客户端控制参数
        lngRow = lngRow + 5
        gVar.ParaBlnLimitClientConnect = .Cell(lngRow, 1).Text '限制客户端连接
        tempVal = Val(.Cell(lngRow, 7).Text)
        gVar.ParaLimitClientConnectTime = IIf(tempVal < 1 Or tempVal > 60, 30, tempVal) '限制客户端连接时长
        tempVal = Val(.Cell(lngRow + 1, 3).Text)
        gVar.TCPConnectMax = IIf(tempVal < 1 Or tempVal > 20, 2, tempVal) '限制客户端连接数
        
        '服务端文件备份参数
        lngRow = lngRow + 4
        gVar.ParaBackupPath = mfCheckFolder(.Cell(lngRow, 3).Text) '备份路径
        Rem gVar.ParaBackupInterval
        Rem gVar.ParaBackupTime
        
    End With
    
    '参数值通过公用变量保存进注册表中
    With gVar
        '窗口控制参数
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowCloseMin, IIf(.ParaBlnWindowCloseMin, 1, 0))    '关闭时最小化
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowMinHide, IIf(.ParaBlnWindowMinHide, 1, 0))  '最小化时隐藏
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowStartMinS, IIf(.ParaBlnWindowStartMinS, 1, 0)) '启动时最小化
        
        '服务端参数
        Call SaveSetting(.RegAppName, .RegSectionTCP, .RegKeyTCPPort, .TCPSetPort)  '侦听端口
        If .ParaBlnAutoStartupAtBoot Then   '注册表中添加启动项
            .ParaBlnAutoStartupAtBoot = gfStartUpSet(True, RegWrite)
        Else    '注册表中删除启动项
            Call gfStartUpSet(True, RegDelete)
        End If
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaAutoStartupAtBoot, IIf(.ParaBlnAutoStartupAtBoot, 1, 0)) '开机自动启动
        
        '数据库服务器参数
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyDBServerIP, .ConSource)
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyDBServerDatabase, EncryptString(.ConDatabase, .EncryptKey)) '数据库名
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyDBServerAccount, EncryptString(.ConUserID, .EncryptKey)) '登陆名
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyDBServerPassword, EncryptString(.ConPassword, .EncryptKey)) '登陆密码
        
        '客户端控制参数
        Call SaveSetting(.RegAppName, .RegSectionTCP, .RegKeyParaLimitClientConnect, IIf(.ParaBlnLimitClientConnect, 1, 0)) '限制客户端连接
        Call SaveSetting(.RegAppName, .RegSectionTCP, .RegKeyParaLimitClientConnectTime, .ParaLimitClientConnectTime) '限制客户端连接时长
        Call SaveSetting(.RegAppName, .RegSectionTCP, .RegKeyParaLimitClientConnectNumber, .TCPConnectMax) '限制客户端连接数
        
        '服务端文件备份参数
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyServerBackStore, .ParaBackupPath) '备份路径
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyServerBackInterval, .ParaBackupInterval) '备份频率
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyServerBackTime, .ParaBackupTime) '备份确切时间
        Call SaveSetting(.RegAppName, .RegSectionDBServer, .RegKeyServerBackIntervalDays, .ParaBackupIntervalDays) '每N天
    End With
    
    Call msLoadParameter(True)  '窗口重新加载一次保存后的值
    
    If MsgBox("参数保存完成！是否现在退出窗口？", vbInformation + vbYesNo, "提示") = vbYes Then Unload Me
    
End Sub


Private Sub Command1_Click()
    Call msLoadParameter(True)  '加载参数值
End Sub

Private Sub Form_Load()
    Dim strFile As String
    Dim K As Long, lngSum As Long
    
    Me.Icon = LoadPicture("")
    Me.Command1.Visible = False
    strFile = gVar.FolderNameBin & "OptionWindowServer.cel"
    If Not gfFileExist(strFile) Then
        MsgBox "以下配置文件加载失败，请解决后再重新打开窗口。" & vbCrLf & strFile, vbCritical, "异常提示"
        Exit Sub
    End If
    With Grid1
        .AutoRedraw = False
        .OpenFile (strFile) '加载模板
        
        .Appearance = Flat
        .Column(0).Width = 0
        .RowHeight(0) = 0
        .ExtendLastCol = True   '扩展最后一列
        .GridColor = vbWhite    '网格线的颜色
        .BorderColor = Me.BackColor '边框的颜色
        .BackColorBkg = Me.BackColor    '空白区域的背景色
        .ReadOnlyFocusRect = Solid  '锁定（只读）单元格所显示的虚框样式
        .DisplayFocusRect = False   '活动单元格是否显示一个虚框
        .SelectionMode = cellSelectionNone  '表格的选择模式
        
        Me.Command1.Value = True  '加载参数值
        
        For K = 0 To .Rows - 1  '计算表格的实际高度
            lngSum = lngSum + .RowHeight(K) * 15    'FC此属性值单位为像素，转成VB的缇要*15.
        Next
        .Height = lngSum    '设置表格高度
        Me.Height = .Top + lngSum + 220 '设置窗口高度
        
        .AutoRedraw = True
        .Refresh
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Grid1.Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - 240
End Sub

Private Sub Grid1_ButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strPath As String
    Dim lngRow As Long, lngCol As Long
    
    On Error Resume Next
    
    lngRow = 19 '浏览按键所在行号
    lngCol = 3  '浏览按键所在列号
    
    If Row = lngRow And Col = lngCol Then    '备份文件保存路径选择
'''        With CommonDialog1   '权宜之计用法。舍弃了
'''            .DialogTitle = "备份路径选择"
'''            .Flags = cdlOFNPathMustExist  '路径必须存在且有效 cdlOFNCreatePrompt=cdlOFNFileMustExist + cdlOFNPathMustExist
'''            .InitDir = IIf(Len(Grid1.Cell(lngRow, lngCol).Text) > 0, Grid1.Cell(lngRow, lngCol).Text, gVar.FolderNameBackup)
'''            .FileName = mconstrTip
'''            .ShowOpen
'''            strPath = mfCheckFolder(.FileName)
'''            If Len(strPath) > 0 Then
'''                If Not Right(strPath, 1) = "\" Then strPath = strPath & "\"
'''                Grid1.Cell(lngRow, lngCol).Text = strPath
'''            End If
'''        End With
        
        strPath = BrowseForFolder(Me, Grid1.Cell(lngRow, lngCol).Text)
        If Len(strPath) > 0 Then
            If Not Right(strPath, 1) = "\" Then strPath = strPath & "\"
            Grid1.Cell(lngRow, lngCol).Text = strPath
        End If
        
    ElseIf Row = 20 And Col = 3 Then    '备份频率设置
        Call gsOpenTheWindow("frmInterval", vbModal, vbNormal)
    
    End If
End Sub

Private Sub Grid1_Click()
    With Grid1.ActiveCell
        If .Row = 12 And .Col = 7 Then  '密码单元格借用TextBox控件处理成星号*
            Text1.Move .Left * 15 + 100, .Top * 15 + 100, .Width * 15, .Height * 15
            With Text1
                .Visible = True
                .ZOrder
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
        End If
    End With
End Sub

Private Sub Grid1_HyperLinkClick(ByVal Row As Long, ByVal Col As Long, URL As String, Changed As Boolean)
    '超链接响应
    
    URL = ""
    Changed = True
    If Row = (Grid1.Rows - 1) Then
        If Col = 3 Then '保存
            If MsgBox("确定保存所有参数值吗？", vbQuestion + vbOKCancel, "保存询问") = vbOK Then Call msSaveParameter(True)
        ElseIf Col = 7 Then '退出
            Unload Me
        End If
    ElseIf Row = 20 And Col = 8 Then    '立即备份
        If MsgBox("是否现在进行文件备份？", vbQuestion + vbYesNo, "备份提示") = vbYes Then
            Call FileBackupCP(gVar.FolderNameStore, gVar.ParaBackupPath)
            Debug.Print "手动备份一次：" & Now
        End If
    End If
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    Dim intRow As Integer, intCol As Integer

    intRow = Grid1.ActiveCell.Row
    intCol = Grid1.ActiveCell.Col
    If (intRow = 19 Or intRow = 20) And intCol = 3 Then '屏蔽输入：备份路径,备份频率
        KeyCode = 0
    End If
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer, intCol As Integer

    intRow = Grid1.ActiveCell.Row
    intCol = Grid1.ActiveCell.Col
    If (intRow = 19 Or intRow = 20) And intCol = 3 Then  '屏蔽输入：备份路径,备份频率
        KeyAscii = 0
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, 65 To 90, 97 To 122  '0-9,A-Z,a-z
'            Debug.Print KeyAscii & ":" & Chr(KeyAscii)
        Case Else
            KeyAscii = 0    '密码：限制字母数字以外的输入
    End Select
End Sub

Private Sub Text1_LostFocus()
    Grid1.Cell(10, 7).Text = String(Len(Text1.Text), "*")   '表格只显示等数量的*号
    Text1.Visible = False
End Sub
