VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSysUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5835
   Icon            =   "frmSysUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5835
   StartUpPosition =   3  '窗口缺省
   Begin FrameFileUpdate.LabelProgressBar LabelProgressBar1 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   2640
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   4335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   1440
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   180
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   180
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   1995
   End
End
Attribute VB_Name = "frmSysUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnHide As Boolean     '更新窗口有隐藏打开模式与显示打开模式
Dim mblnCheckStart As Boolean   '已开始检查标识
Dim mblnUpdateFinish As Boolean     '更新完成标识




Private Function mfCheckUpdate() As Boolean
    '检查更新
    Dim strFileLoc As String, strFileNet As String, strVerLoc As String, strVerNet As String
    
    strFileLoc = gVar.AppPath & gVar.EXENameOfClient
    If Not gfDirFile(strFileLoc) Then Exit Function
    strVerLoc = Trim(gfBackVersion(strFileLoc))
    If Len(strVerLoc) = 0 Then Exit Function
    
    If Me.Winsock1.Item(1).State <> 7 Then Exit Function
    Call msSetText("正在联网验证版本中……", vbBlue)
    Call gfSendInfo(gVar.PTVersionOfClient & strVerLoc, Winsock1.Item(1))
    
End Function

Private Function mfConnect(Optional ByVal blnCon As Boolean = True) As Boolean
    Dim strIP As String, strPort As String
    Static lngCount As Long
            
    lngCount = lngCount + 1
    If lngCount = 2 Then
        Call msSetText("版本检测失败！无法连接服务器。" & vbCrLf & _
                       "请确认服务器已启动，并重新运行更新程序！", vbRed)
        Exit Function    '尝试百次后不再连接了
    End If
    
    With Me.Winsock1.Item(1)
        If Label1(1).Caption = gVar.ClientStateDisConnected Then
            If .State <> 0 Then .Close
            .RemoteHost = gVar.TCPSetIP
            .RemotePort = gVar.TCPSetPort
            .Connect
            If .State = 7 Then gVar.TCPStateConnected = True
        End If
    End With
End Function

Private Function mfShellSetup(ByVal strFile As String) As Boolean
    '关闭客户端程序，执行更新安装包
    
    Dim strClient As String
    
    If MsgBox("是否立即执行更新程序？", vbQuestion + vbYesNo, "安装询问") = vbYes Then
        If gfCloseApp(gVar.EXENameOfClient) Then   '关闭客户端exe
            If gfShellExecute(strFile) Then     '运行安装包
                Unload Me
            End If
        Else
            MsgBox "请确认已关闭客户端程序，并重新运行更新程序！", vbInformation, "警告"
        End If
    Else
        Call Winsock1_Close(1)
        Unload Me
    End If
End Function

Private Sub msLoadParameter(Optional ByVal blnLoad As Boolean = True)
    '从注册表中加载参数值至公用变量中
    
    If Not blnLoad Then Exit Sub
    
    On Error Resume Next    '加/解密函数过程可能有异常
    With gVar
        .TCPDefaultIP = Me.Winsock1.Item(0).LocalIP '本机IP地址
        .TCPSetIP = gfCheckIP(GetSetting(.RegAppName, .RegSectionTCP, .RegKeyTCPIP, .TCPDefaultIP)) '要连接服务端IP地址
        .TCPSetPort = gfGetRegNumericValue(.RegAppName, .RegSectionTCP, .RegKeyTCPPort, , .TCPDefaultPort, 10000, 65535) '要连接的服务器端口
    End With
End Sub

Private Sub msSetLabel(ByVal strCaption As String, ByVal BackColor As Long)
    Me.Label1.Item(1).Caption = strCaption
    Me.Label1.Item(1).BackColor = BackColor
End Sub

Private Sub msSetText(ByVal strTxt As String, ByVal ForeColor As Long)
    Me.Text1.Text = strTxt
    Me.Text1.ForeColor = ForeColor
End Sub


Private Sub Form_Load()
    
    Dim strCmd As String, arrCmd() As String
    
    Label1.Item(0).Caption = ""
    Text1.BackColor = Me.BackColor
    Timer1.Interval = 1000
    Timer1.Enabled = True

    ReDim gArr(1)
    
    Call Main
    Call msLoadParameter(True)
    
    '检测是否传入命令行参数进来，没有则退出程序
    strCmd = Command
    If Len(strCmd) = 0 Then
        GoTo LineUnload '禁止直接启动更新程序，必须带命令参数
    Else
        arrCmd = Split(strCmd, gVar.CmdLineSeparator)
        
        If UCase(arrCmd(0)) <> UCase(gVar.EXENameOfClient) Then
            GoTo LineUnload    '命令参数中第一串字符固定为exe文件名，不是则认为非法启动更新程序，不准执行
        End If
        
        If UBound(arrCmd) > 0 Then  '判断命令参数中是否带否隐藏窗口命令
            If LCase(arrCmd(1)) = LCase(gVar.CmdLineParaOfHide) Then
                mblnHide = True
                Me.Hide
            End If
        End If
    End If
    
    Call msSetLabel(gVar.ClientStateDisConnected, vbRed)
    Call mfConnect(True)
    
    Exit Sub
    
LineUnload:
    Unload Me   '此行以下除End Sub不可再跟任何有效代码
End Sub

Private Sub Timer1_Timer()
    Const conConn As Byte = 1       '连接状态检测间隔conConn秒
    Const conState As Byte = 5      '连接服务器检测间隔conState秒
    
    Static byteConn As Byte
    Static byteState As Byte
    Static byteDotCount As Byte
    
    byteConn = byteConn + 1
    byteState = byteState + 1
    
    If byteConn >= conConn Then
        If Me.Winsock1.Item(1).State = 7 Then
            Call msSetLabel(gVar.ClientStateConnected, vbGreen)
            gVar.TCPStateConnected = True
            If Not mblnCheckStart And gArr(1).Connected Then
                mblnCheckStart = True
                Call mfCheckUpdate
            End If
        Else
            Call msSetLabel(gVar.ClientStateDisConnected, vbRed)
            gVar.TCPStateConnected = False
        End If
        byteConn = 0    '复位静态变量
    End If
    
    If byteState >= conState Then
        If Me.Winsock1.Item(1).State <> 7 Then
            If Not mblnUpdateFinish Then Call mfConnect
        End If
        byteState = 0   '复位静态变量
    End If
    
    If gArr(1).FileTransmitState Then
        byteDotCount = byteDotCount + 1
        If byteDotCount > 6 Then byteDotCount = 1
        Me.Label1.Item(0).Caption = "更新下载中" & String(byteDotCount, "·")
    End If
End Sub

Private Sub Winsock1_Close(Index As Integer)
    '传输被关闭
    If UBound(gArr) = 1 Then
        gArr(1) = gArr(0)
        Rem Debug.Print "Winsock1_Close trigger all time ?"
    End If
    
    If mblnCheckStart Then
        Call msSetText("服务器连接中断！版本更新检测失败！", vbRed)
        mblnCheckStart = False
    End If
    Label1.Item(0).Caption = ""
End Sub


Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    '接收服务器端传来信息或文件
    
    Dim strGet As String    '接收字符信息
    Dim byteGet() As Byte   '接收文件
    
    With gArr(Index)
        If Not .FileTransmitState Then
            '字符信息传输状态↓
            
            Me.Winsock1.Item(Index).GetData strGet
            
            If InStr(strGet, gVar.PTClientConfirm) Then
                Call gfSendInfo(gVar.PTClientIsTrue, Me.Winsock1.Item(Index))
                .Connected = True
                Call gfSendClientInfo("UpdatePC", "Update", "UpdateProgram", Me.Winsock1.Item(Index))
                
            End If
            
            If Not gfRestoreInfo(strGet, Me.Winsock1.Item(Index)) Then
                
            End If
            
            If InStr(strGet, gVar.PTVersionNeedUpdate) > 0 Then
                Dim strVer As String
                
                strVer = Mid(strGet, Len(gVar.PTVersionNeedUpdate) + 1)
                Call msSetText("发现新版：" & strVer, vbBlue)
            End If
            
            If InStr(strGet, gVar.PTVersionNotUpdate) > 0 Then
                Dim strNot As String
                
                If Len(strGet) = Len(gVar.PTVersionNotUpdate) Then
                    strNot = "您当前的版本已是最新版本，不需要更新。"
                    Call msSetText(strNot, vbGreen)
                    If mblnHide Then Unload Me  '隐藏模式打开更新窗口时，无更新则直接退出
                Else
                    strNot = Mid(strGet, Len(gVar.PTVersionNotUpdate) + 1)
                    strNot = "版本检测异常：" & strNot
                    Call msSetText(strNot, vbMagenta)
                End If
                
                mblnUpdateFinish = True
            End If
            Debug.Print "Get Server Info:" & strGet, bytesTotal
            '字符信息传输状态↑
            
        Else
            '文件传输状态↓
            
            If .FileNumber = 0 Then
                .FileNumber = FreeFile
                Open .FilePath For Binary As #.FileNumber
                
                LabelProgressBar1.Min = 0
                LabelProgressBar1.Max = .FileSizeTotal
                LabelProgressBar1.Value = 0
            End If
            
            ReDim byteGet(bytesTotal - 1)
            Me.Winsock1.Item(Index).GetData byteGet, vbArray + vbByte
            Put #.FileNumber, , byteGet
            .FileSizeCompleted = .FileSizeCompleted + bytesTotal
            LabelProgressBar1.Value = .FileSizeCompleted
            
            If .FileSizeCompleted >= .FileSizeTotal Then
                Dim strSetupFile As String
                
                strSetupFile = .FilePath
                Close #.FileNumber
                Call gfSendInfo(gVar.PTFileEnd, Winsock1.Item(Index))
                gArr(Index) = gArr(0)
                Label1.Item(0).Caption = "下载完成！"
                
                Call mfShellSetup(strSetupFile)
                
                Debug.Print "Received Over"
            End If
            
            '文件传输状态↑
        End If
    End With
    
End Sub


Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Index <> 0 Then
        If gArr(Index).FileTransmitState Then   '异常时清空文件传输信息
            Close #gArr(Index).FileNumber
            gArr(Index) = gArr(0)
        End If
    End If
End Sub
