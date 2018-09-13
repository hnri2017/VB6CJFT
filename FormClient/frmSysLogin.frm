VERSION 5.00
Begin VB.Form frmSysLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   Icon            =   "frmSysLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5295
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "登陆"
      Default         =   -1  'True
      Height          =   375
      Left            =   1410
      TabIndex        =   2
      Top             =   2160
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1770
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎登陆系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   2
      Left            =   1215
      TabIndex        =   7
      Top             =   240
      Width           =   2715
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "用户名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   915
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "密  码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   1500
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   6330
      Left            =   120
      Picture         =   "frmSysLogin.frx":000C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   12000
   End
End
Attribute VB_Name = "frmSysLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub msLoadUserInfo(Optional ByVal blnLoad As Boolean = True)
    '加载登陆过的用户信息
    
    Dim strReg As String, arrUser() As String, strLast As String, strPWDde As String
    Dim K As Long, C As Long
    
    '加载用户名列表
    strReg = GetSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserList, "") '获取列表
    If Len(strReg) = 0 Then Exit Sub    '没有保存的用户名则退出
    arrUser = Split(strReg, gVar.CmdLineSeparator) '分解列表
    C = UBound(arrUser)
    Combo1.Clear
    For K = 0 To C
        Combo1.AddItem Trim(arrUser(K)) '将每个用户名加载进下拉列表中
    Next
    
    '加载最近登陆的用户名
    strLast = GetSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserLast, "")
    Combo1.Text = Trim(strLast)
    
    '如果勾选了记住密码，则自动填充对应密码
    If gVar.ParaBlnRememberUserPassword And Len(strLast) > 0 Then
        strPWDde = GetSetting(gVar.RegAppName, gVar.RegSectionUser, strLast, "")
        If Len(strPWDde) > 0 Then
            On Error Resume Next    '密文异常时可能报错
            Text1.Text = DecryptString(strPWDde, gVar.EncryptKey)
            If Err.Number <> 0 Then
                Call gsAlarmAndLog("密码被破坏警报")
            End If
        End If
    End If

End Sub

Private Sub msSaveUserInfo(Optional ByVal blnSave As Boolean = True)
    '保存登陆过的用户信息
    Dim strCurUser As String, strList As String, strCombo As String
    Dim K As Long, C As Long
    
    '用户列表处理
    If gVar.ParaBlnRememberUserList Then '保存用户列表
        strCurUser = Trim(Combo1.Text)
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserLast, strCurUser) '记录最近登陆过的用户名
        strList = strCurUser '当前登陆用户总是排在列表第一位
        C = Combo1.ListCount
        If C > 0 Then '下拉中有其他用户名
            strCurUser = LCase(strCurUser)
            C = C - 1
            For K = 0 To C '生成新顺序的用户名列表
                strCombo = LCase(Trim(Combo1.List(K)))
                If strCombo <> strCurUser Then
                    strList = strList & gVar.CmdLineSeparator & strCombo '两个名之间用分隔符gVar.CmdLineSeparator
                End If
            Next
        End If
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserList, strList) '保存用户列表到注册表中
    Else '清除注册表中的用户
        If gfGetSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserLast) Then
            Call gsDeleteSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserLast, "最新登陆用户名删除异常") '删除最近登陆用户
        End If
        If gfGetSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserList) Then
            Call gsDeleteSetting(gVar.RegAppName, gVar.RegSectionUser, gVar.RegKeyUserList, "用户名记录列表删除异常") '删除用户列表
        End If
    End If
    
    '记住密码处理
    strCurUser = Trim(Combo1.Text)
    If gVar.ParaBlnRememberUserPassword Then '加密保存密码
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionUser, strCurUser, EncryptString(Text1.Text, gVar.EncryptKey))
    Else '删除密码
        If gfGetSetting(gVar.RegAppName, gVar.RegSectionUser, strCurUser) Then
            Call gsDeleteSetting(gVar.RegAppName, gVar.RegSectionUser, strCurUser, "用户" & strCurUser & "记住密码删除异常")
        End If
    End If
    
End Sub

Private Sub Command1_Click()
    '登陆系统
    Dim strName As String, strPWD As String
    
    strName = Trim(Combo1.Text)
    strPWD = Text1.Text
    Call msSaveUserInfo(True)
    
    gWind.Show
    gVar.UserLoginName = strName
    gVar.UserFullName = strPWD
    Call gfSendClientInfo(gVar.UserComputerName, gVar.UserLoginName, gVar.UserFullName, gWind.Winsock1.Item(1))
    gVar.ShowMainWindow = True
    Unload Me
End Sub

Private Sub Command2_Click()
    gVar.UnloadFromLogin = True
End Sub

Private Sub Form_Load()
    '加载窗体
    
    If gVar.ParaBlnRememberUserList Then
        Call msLoadUserInfo(True) '加载用户列表
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Image1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gVar.UnloadFromLogin = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Label3.FontUnderline Then    '复原样式
        Label3.FontUnderline = False '去除下划线
        Label3.ForeColor = vbBlack  '字体黑色
    End If
End Sub

Private Sub Label3_Click()
    '弹出设置窗口
    Call gsOpenTheWindow("frmOption", vbModal, vbNormal)
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '通API函数鼠标指针变手指
    
    Dim hHandCursor As Long
    
    Label3.FontUnderline = True '显示下划线
    Label3.ForeColor = vbRed '字体红色
    hHandCursor = LoadCursor(0, IDC_HAND) '调用API载入光标
    Call SetCursor(hHandCursor) '调用API使指针变手指状
End Sub
