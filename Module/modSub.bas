Attribute VB_Name = "modSub"
Option Explicit


Public Sub Main(Optional ByVal blnLoad As Boolean = True)
    
    Dim strTemp As String
    
    '主窗体CommandBars的ID值初始化
    With gID
        .Sys = 1000
        
        .SysLoginOut = 1101
        .SysLoginAgain = 1102
        .SysAuthChangePassword = 1103
        .SysAuthDepartment = 1104
        .SysAuthUser = 1105
        .SysAuthRole = 1106
        .SysAuthFunc = 1107
        .SysAuthLog = 1108
        .SysFileManage = 1121
        
        .SysExportMain = 1200
        .SysExportToCSV = 1201
        .SysExportToExcel = 1202
        .SysExportToHTML = 1203
        .SysExportToPDF = 1204
        .SysExportToXML = 1205
        .SysExportToText = 1206
        .SysExportToWord = 1207
        
        .SysPrintMain = 1300
        .SysPrint = 1303
        .SysPrintPageSet = 1301
        .SysPrintPreview = 1302
                
        .SysSearch = 1400
        .SysSearch1Label = 1401
        .SysSearch2TextBox = 1402
        .SysSearch3Button = 1403
        .SysSearch4ListBoxCaption = 1404
        .SysSearch4ListBoxFormID = 1405
        .SysSearch5Go = 1406
        
        
        .Wnd = 2000
        
        .WndResetLayout = 2050
        .TabWorkspacePopupMenu = 2051
        .WndThemeSkinSet = 2052
        .WndOpenListCaption = 2053
        
        .WndOpenListID = XTP_ID_WINDOWLIST '=35000
        .WndToolBarCustomize = XTP_ID_CUSTOMIZE '=35001
        .WndToolBarList = XTP_ID_TOOLBARLIST '=59392
        
        .WndThemeCommandBars = 2100
        .WndThemeCommandBarsOffice2000 = 2101
        .WndThemeCommandBarsOffice2003 = 2102
        .WndThemeCommandBarsOfficeXp = 2103
        .WndThemeCommandBarsResource = 2104
        .WndThemeCommandBarsRibbon = 2105
        .WndThemeCommandBarsVS2008 = 2106
        .WndThemeCommandBarsVS2010 = 2107
        .WndThemeCommandBarsVS6 = 2108
        .WndThemeCommandBarsWhidbey = 2109
        .WndThemeCommandBarsWinXP = 2110

        .WndThemeTaskPanel = 2200
        .WndThemeTaskPanelListView = 2201
        .WndThemeTaskPanelListViewOffice2003 = 2202
        .WndThemeTaskPanelListViewOfficeXP = 2203
        .WndThemeTaskPanelNativeWinXP = 2204
        .WndThemeTaskPanelNativeWinXPPlain = 2205
        .WndThemeTaskPanelOffice2000 = 2206
        .WndThemeTaskPanelOffice2000Plain = 2207
        .WndThemeTaskPanelOffice2003 = 2208
        .WndThemeTaskPanelOffice2003Plain = 2209
        .WndThemeTaskPanelOfficeXPPlain = 2210
        .WndThemeTaskPanelResource = 2211
        .WndThemeTaskPanelShortcutBarOffice2003 = 2212
        .WndThemeTaskPanelToolbox = 2213
        .WndThemeTaskPanelToolboxWhidbey = 2214
        .WndThemeTaskPanelVisualStudio2010 = 2215
        
        .WndSon = 2300
        .WndSonCloseAll = 2301
        .WndSonCloseCurrent = 2302
        .WndSonCloseLeft = 2303
        .WndSonCloseOther = 2304
        .WndSonCloseRight = 2305
        .WndSonVbAllBack = 2306
        .WndSonVbAllMin = 2307
        .WndSonVbArrangeIcons = 2308
        .WndSonVbCascade = 2309
        .WndSonVbTileHorizontal = 2310
        .WndSonVbTileVertical = 2311
        
        
        .WndThemeSkin = 2400
        .WndThemeSkinCodejock = 2401
        .WndThemeSkinOffice2007 = 2402
        .WndThemeSkinOffice2010 = 2403
        .WndThemeSkinVista = 2404
        .WndThemeSkinWinXPLuna = 2405
        .WndThemeSkinWinXPRoyale = 2406
        .WndThemeSkinZune = 2407
               
        
        .Help = 3000
        .HelpAbout = 3101
        .HelpDocument = 3102
        .HelpUpdate = 3103
        
        
        .Tool = 4000
        .toolErrorLog = 4101
        .toolLoginLog = 4102
        .toolOptions = 4103
        
        
        '''***请将所有菜单栏中的【菜单】的CommandBrs的ID值设置在20000以下*******************
        
        
        .Pane = 21000
        
        .PaneNavi = 21102
        
        .PanePopupMenuNavi = 21103
        .PanePopupMenuNaviAutoFoldOther = 21104
        .PanePopupMenuNaviExpandALL = 21105
        .PanePopupMenuNaviFoldALL = 21106
        
        
        .StatusBarPane = 22000
        
        .StatusBarPaneConnectButton = 22101
        .StatusBarPaneConnectState = 22102
        .StatusBarPaneProgress = 22103
        .StatusBarPaneProgressText = 22104
        .StatusBarPaneServerButton = 22105
        .StatusBarPaneServerState = 22106
        .StatusBarPaneTime = 22107
        .StatusBarPaneUserInfo = 22108
        .StatusBarPaneIP = 22109
        .StatusBarPanePort = 22110
        .StatusBarPaneReStartButton = 22111
        
        .IconPopupMenu = 23000
        .IconPopupMenuMaxWindow = 23101
        .IconPopupMenuMinWindow = 23102
        .IconPopupMenuShowWindow = 23103
        
    End With
    
    '公用变量值初始化
    With gVar
        
        .TCPConnectMax = 20 '单位个
        .TCPDefaultIP = "127.0.0.1"
        .TCPDefaultPort = 19898
        .TCPWaitTime = 3    '单位秒
                
        .UpdateAccount = "UpdatePC"
        .UpdatePCName = "Update"
        .UpdateUserName = "UpdateProgram"
        
        .FTChunkSize = 5734
        .FTWaitTime = 3     '单位秒
        
        .EncryptKey = "[FT]"    '密钥
        
        .ServerButtonClose = "关闭服务"
        .ServerButtonStart = "开启服务"
        .ServerStateError = "异常"
        .ServerStateNotStarted = "未启动"
        .ServerStateStarted = "已启动"
        
        .ClientStateConnected = "已连接"
        .ClientStateDisConnected = "未连接"
        .ClientStateConnectError = "连接异常"
        .ClientButtonConnectToServer = "建立连接"
        .ClientButtonDisConnectFromServer = "断开连接"
        
        .PTFileName = "<FileName>"
        .PTFileSize = "<FileSize>"
        .PTFileFolder = "<FileFolder>"
        .PTFileStart = "<FileStart>"
        .PTFileEnd = "<FileEnd>"
        .PTFileSend = "<FileSend>"
        .PTFileReceive = "<FileReceive>"
        .PTFileExist = "<FileExist>"
        .PTFileNoExist = "<FileNoExist>"
        
        .PTVersionNeedUpdate = "<VersionNeedUpdate>"
        .PTVersionNotUpdate = "<VersionNotUpdate>"
        .PTVersionOfClient = "<VersionOfClient>"
        
        .PTClientConfirm = "<ClientConfirm>"
        .PTClientIsTrue = "<ClientIsTrue>"
        
        .PTConnectIsFull = "ConnectIsFull"
        .PTConnectTimeOut = "ConnectTimeOut"
        
        .PTClientUserComputerName = "<ClientUserComputerName>"
        .PTClientUserFullName = "<ClientUserFullName>"
        .PTClientUserLoginName = "<ClientUserLoginName>"
        
        .PTDBDatabase = "<DBDatabase>"
        .PTDBDataSource = "<DBDataSource>"
        .PTDBPassword = "<DBPassword>"
        .PTDBUserID = "<DBUserID>"
        
        .EXENameOfClient = "FFC.exe"
        .EXENameOfServer = "FFS.exe"
        .EXENameOfSetup = "setup-fc.exe" '"客户端更新/安装程序"
        .EXENameOfUpdate = "FFU.exe"
        
        .CmdLineParaOfHide = "Hide"
        .CmdLineSeparator = " / "
        
        .RegAppName = "FF"
        .RegKeyTCPIP = "IP"
        .RegKeyTCPPort = "Port"
        .RegSectionTCP = "TCP"
        
        .RegSectionSkin = "SkinFile"
        .RegKeySkinRes = "SkinRes"
        .RegKeySkinIni = "SkinIni"
        .RegKeySkinSvrRes = "SkinSvrRes"
        .RegKeySkinSvrIni = "SkinSvrIni"
        
        .RegSectionDBServer = "Server"
        .RegKeyDBServerAccount = "ServerAccount"
        .RegKeyDBServerDatabase = "ServerDatabase"
        .RegKeyDBServerIP = "ServerIP"
        .RegKeyDBServerPassword = "ServerPassword"
        .RegKeyServerBackStore = "BackupStore"
        .RegKeyServerBackInterval = "BackupInterval"
        .RegKeyServerBackTime = "BackupTime"
        
        .RegSectionUser = "UserInfo"
        .RegKeyUserLast = "LastLoginUser"
        .RegKeyUserList = "LoginUserList"
        
        .RegKeyCommandBars = "FF"
        .RegKeyCBSClientSetting = "ClientSetting"
        .RegKeyCBSServerSetting = "ServerSetting"
        .RegKeyDockingPane = .RegKeyCommandBars
        .RegKeyDockPaneClientSetting = "ClientSetting"
        .RegKeyDockPaneServerSetting = "ServerSetting"
        
        .RegSectionSettings = "Settings"
        .RegKeyServerWindowHeight = "ServerWindowHeight"
        .RegKeyServerWindowLeft = "ServerWindowLeft"
        .RegKeyServerWindowTop = "ServerWindowTop"
        .RegKeyServerWindowWidth = "ServerWindowWidth"
        .RegKeyServerWindowStateMax = "ServerWindowStateMax"
        .RegKeyServerCommandbarsTheme = "ServercbsTheme"
        
        .RegKeyClientWindowHeight = "ClientWindowHeight"
        .RegKeyClientWindowLeft = "ClientWindowLeft"
        .RegKeyClientWindowTop = "ClientWindowTop"
        .RegKeyClientWindowWidth = "ClientWindowWidth"
        .RegKeyClientWindowStateMax = "ClientWindowStateMax"
        .RegKeyClientCommandbarsTheme = "ClientcbsTheme"
        .RegKeyClientTaskPanelAutoFold = "ClientTPAutoFold"
        .RegKeyClientTaskPanelTheme = "ClientTPTheme"
        
        .RegTrailPath = "SoftWare\Common\Section"   'HKEY_CURRENT_USER\SoftWare\……
        .RegTrailKey = "Key"
        .TrailPeriod = 15
        
        .RegKeyParaWindowMinHide = "WindowMinHide"
        .RegKeyParaWindowCloseMin = "WindowCloseMin"
        .RegKeyParaWindowStartMinS = "WindowStartMinS"
        .RegKeyParaWindowStartMinC = "WindowStartMinC"
        .RegKeyParaAutoReStartServer = "AutoReStartServer"
        .RegKeyParaAutoStartupAtBoot = "AutoStartupAtBoot"
        .RegKeyParaLimitClientConnect = "LimitClientConnect"
        .RegKeyParaLimitClientConnectTime = "LimitClientConnectTime"
        .RegKeyParaLimitClientConnectNumber = "LimitClientConnectNumber"
        
        .RegKeyParaUserAutoLogin = "UserAutoLogin"
        .RegKeyParaRememberUserList = "RememberUserList"
        .RegKeyParaRememberUserPassword = "RememberUserPassword "
        
        .AppPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
        
        .FolderBin = "Bin"
        .FolderData = "Data"
        .FolderTemp = "Temp"
        .FolderBackup = "Backup"
        .FolderStore = "Store"
        .FolderNameBin = .AppPath & .FolderBin & "\"
        .FolderNameData = .AppPath & .FolderData & "\"
        .FolderNameTemp = .AppPath & .FolderTemp & "\"
        .FolderNameBackup = "D:\Backup\Store\"
        .FolderNameStore = .AppPath & "Store\"
        
        .FileNameErrLog = .FolderNameData & "ErrorRecord.LOG"
        .FileNameSkin = ""
        .FileNameSkinIni = ""
        .FileNameLoginLog = .FolderNameData & "LoginLog.LOG"
        
        .AccountAdmin = "Admin"     '两个特殊用户
        .AccountSystem = "System"   '两个特殊用户
        
        .FuncButton = "按钮"
        .FuncControl = "其它"
        .FuncForm = "窗口"
        .FuncMainMenu = "主菜单"
        
        .Formaty_M_dH_m_s = "yyyy-MM-dd HH:mm:ss"   '时间格式
        .Formatymdhms = "yyyyMMddHHmmss"
        
        .WindowHeight = 8700
        .WindowWidth = 15800
        
'''        '''*****在注册表中保存服务器地址、访问的账号与密码****
'''        '转移至Server端主窗体中加载函数中了
'''        .ConSource = gfCheckIP(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerIP, ""))
'''        .ConDatabase = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerDatabase))     '暂仅限连接SQLServer2008 OR 2012 数据库
'''        .ConUserID = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerAccount))
'''        .ConPassword = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerPassword))
'''        .ConString = "Provider=SQLOLEDB;Persist Security Info=False;Data Source=" & .ConSource & _
'''                    ";UID=" & .ConUserID & ";PWD=" & .ConPassword & _
'''                    ";DataBase=" & .ConDatabase & ";"   '''在64位系统上Data Source中间要空格隔开才能建立连接
        
    End With
    
End Sub


Public Sub gsAlarmAndLog(Optional ByVal strErr As String, Optional ByVal blnMsgBox As Boolean = True, _
        Optional ByVal MsgButton As VbMsgBoxStyle = vbCritical)
    '系统异常提示并写下异常日志
    
    Dim strMsg As String, strWrt As String
    
    strMsg = "异常代号：" & Err.Number & vbCrLf & "异常描述：" & Err.Description
    If blnMsgBox Then MsgBox strMsg, MsgButton, strErr
    strWrt = Err.Number & vbTab & vbTab & Err.Description
    Call gsFileWrite(gVar.FileNameErrLog, strErr & vbTab & vbTab & strWrt)
    
End Sub

Public Sub gsAlarmAndLogEx(Optional ByVal strErrDescription As String, Optional ByVal strErrTitle As String, _
        Optional ByVal blnMsgBox As Boolean = True, Optional ByVal MsgButton As VbMsgBoxStyle = vbCritical)
    '自定义异常提示并写下日志
    
    Err.Clear
    If Err.Number = 0 Then Err.Number = vbObjectError + 100001 '固定一个自定义异常号码
    If Len(Err.Description) = 0 Then Err.Description = strErrDescription
    Call gsAlarmAndLog(strErrTitle, blnMsgBox, MsgButton)
    
End Sub

Public Sub gsDeleteSetting(ByVal AppName As String, ByVal Section As String, ByVal Key As String, ByVal strMsg As String)
    '调用系统函数删除注册信息
    
    On Error Resume Next
    Call DeleteSetting(AppName, Section, Key) '不存在时可能异常
    If Err.Number <> 0 Then
        Call gsAlarmAndLog(strMsg, False)
    End If
End Sub

Public Sub gsFileProgress(ByRef ProgressBar As XtremeCommandBars.StatusBarProgressPane, _
                          ByRef ProgressTxt As XtremeCommandBars.StatusBarPane, _
                          Optional ByVal pType As genumFileProgressValue = ftZero, _
                          Optional ByVal pValue As Long = 0, _
                          Optional ByVal pMax As Long = 100, Optional ByVal pMin As Long = 0)
    '主窗口状态栏上的文件传输进度条与比率设置
    ''调用模式:Call gsFileProgress(gWind.CommandBars1.StatusBar.FindPane(gID.StatusBarPaneProgress), _
    '                            gWind.CommandBars1.StatusBar.FindPane(gID.StatusBarPaneProgressText), _
    '                            ftZero, 0, lngFileLen, 0)
    With ProgressBar
        If pType = ftZero Then
            .Max = pMax  '设置最大值
            .Min = pMin  '设置最小值
            .Value = pMin   '当前值为最小值
        ElseIf pType = ftOver Then
            .Value = pMax
        Else
            If pValue > .Max Then
                .Value = .Max   '当前值不能大于进度条的最大值
            Else
                .Value = pValue '正常进度值
            End If
        End If
        ProgressTxt.Text = CStr(Int(.Value / (.Max - .Min) * 100)) & "%"   '数值百分比
    End With
End Sub

Public Sub gsFileWrite(ByVal strFile As String, ByVal strContent As String, _
    Optional ByVal OpenMode As genumFileOpenType = udAppend, _
    Optional ByVal WriteMode As genumFileWriteType = udPrint)
    '将指定内容以指定的方式写入指定文件中
    
    Dim intNum As Integer
    Dim strTime As String
    
    If Not gfFileRepair(strFile) Then Exit Sub
    intNum = FreeFile
    
    On Error Resume Next
    
    Select Case OpenMode
        Case udBinary
            Open strFile For Binary As #intNum
        Case udInput
            Open strFile For Input As #intNum
        Case udOutput
            Open strFile For Output As #intNum
        Case Else   '其余皆当作udAppend
            Open strFile For Append As #intNum
    End Select
    
    strTime = Format(Now, gVar.Formaty_M_dH_m_s)
    Select Case WriteMode
        Case udWrite
            Write #intNum, strTime, strContent
        Case udPut
            Put #intNum, , strTime & vbTab & strContent
        Case Else   '其余皆当作udPrint
            Print #intNum, strTime & vbTab & vbTab & strContent
    End Select
    
    Close #intNum
    
End Sub


Public Sub gsFormScrollBar(ByRef frmCur As Form, ByRef ctlMv As Control, _
    ByRef Hsb As HScrollBar, ByRef Vsb As VScrollBar, _
    Optional ByVal lngMW As Long = 12000, _
    Optional ByVal lngMH As Long = 9000, _
    Optional ByVal lngHV As Long = 255)
    
    'frmCur：滚动条所在的窗体
    'ctlMv：窗体中的控件（除滚动条以外）都在此容器控件中
    'Hsb：窗体frmCur中水平滚动条控件
    'Vsb：窗体frmCur中垂直滚动条控件
    'lngMW：窗体不出现滚动条的宽度
    'lngMH：窗体不出现滚动条的高度
    'lngHV：滚动条的窄边宽度或高度。
    '***注意注意注意：滚动条控件需最后添加至窗体中，且不能放在容器控件ctlMv中*******
    
    Dim lngW As Long
    Dim lngH As Long
    Dim lngSW As Long
    Dim lngSH As Long
    Dim lngMin As Long
    
    lngW = frmCur.Width
    lngH = frmCur.Height
    lngSW = frmCur.ScaleWidth
    lngSH = frmCur.ScaleHeight
    lngMin = -120
    
    On Error Resume Next
    
    If lngW >= lngMW Then
        Hsb.Visible = False
        ctlMv.Left = -lngMin
    Else
        With Hsb
            .Move 0, lngSH - lngHV, lngSW, lngHV
            .Min = lngMin
            .Max = lngMW - lngW + lngHV
            .SmallChange = 10
            .LargeChange = 50
            .Visible = True
        End With
    End If
    
    If lngH >= lngMH Then
        Vsb.Visible = False
        ctlMv.Top = -lngMin
    Else
        With Vsb
            .Move lngSW - lngHV, 0, lngHV, IIf(Hsb.Visible, lngSH - lngHV, lngSH)
            .Min = lngMin
            .Max = lngMH - lngH + lngHV
            .SmallChange = 10
            .LargeChange = 50
            .Visible = True
        End With
    End If
    
'    '在窗体中添加窗口控件ctlMove，将所有其它控件放入此容器中，然
'    '后添加名称分别为Hsb\Vsb的水平\垂直滚动条在窗体中，最好留到最后放入窗体中
'    '然后在窗体中添加以下事件调用即可
'Private Sub Form_Resize()
'    Call gsFormScrollBar(Me, Me.ctlMove, Me.Hsb, Me.Vsb, 12000, 9000)  '注意长、宽的修改
'End Sub
'Private Sub Hsb_Change()
'    ctlMove.Left = -Hsb.Value
'End Sub
'
'Private Sub Hsb_Scroll()
'    Call Hsb_Change    '当滑动滚动条中的滑块时会同时更新对应内容，以下同。
'End Sub
'
'Private Sub Vsb_Change()
'    ctlMove.Top = -Vsb.Value
'End Sub
'
'Private Sub Vsb_Scroll()
'    Call Vsb_Change
'End Sub

End Sub

Public Sub gsFormSizeLoad(ByRef frmLoad As Form, Optional blnServer As Boolean = True)
    '从注册表中加载窗口的位置与大小信息
    Dim Left As Long, Top As Long, Width As Long, Height As Long
    Dim blnStateMax As Boolean
    
    blnStateMax = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, IIf(blnServer, gVar.RegKeyServerWindowStateMax, gVar.RegKeyClientWindowStateMax), 1))
    If blnStateMax Then
        frmLoad.WindowState = vbMaximized
    Else
        If blnServer Then
            Left = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyServerWindowLeft, 0))
            Top = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyServerWindowTop, 0))
            Width = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyServerWindowWidth, gVar.WindowWidth))
            Height = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyServerWindowHeight, gVar.WindowHeight))
        Else
            Left = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientWindowLeft, 0))
            Top = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientWindowTop, 0))
            Width = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientWindowWidth, gVar.WindowWidth))
            Height = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientWindowHeight, gVar.WindowHeight))
        End If
        If Width <= 0 Or Width > Screen.Width Then Width = gVar.WindowWidth
        If Height <= 0 Or Height > Screen.Height Then Height = gVar.WindowHeight
        If Left <= 0 Or Left > Screen.Width Then
            If Width < Screen.Width Then
                Left = (Screen.Width - Width) / 2
            Else
                Left = 0
            End If
        End If
        If Top <= 0 Or Left > Screen.Height Then
            If Height < Screen.Height Then
                Top = (Screen.Height - Height) / 2
            Else
                Top = 0
            End If
        End If
        If frmLoad.WindowState = vbNormal Then frmLoad.Move Left, Top, Width, Height
    End If
End Sub

Public Sub gsFormSizeSave(ByRef frmSave As Form, Optional ByVal blnServer As Boolean = True)
    '保存窗口的位置与大小信息至注册表中
    Dim Left As Long, Top As Long, Width As Long, Height As Long
    Dim blnStateMax As Boolean
    
    If frmSave.WindowState = vbMaximized Then blnStateMax = True
    
    If blnStateMax Then
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, IIf(blnServer, gVar.RegKeyServerWindowStateMax, gVar.RegKeyClientWindowStateMax), 1)
    Else
        With frmSave
            Left = .Left
            Top = .Top
            Width = .Width
            Height = .Height
            If Left < 0 Or Left > Screen.Width Then Left = 0
            If Top < 0 Or Top > Screen.Height Then Top = 0
            If Width < gVar.WindowWidth Or Width > Screen.Width Then Width = gVar.WindowWidth
            If Height < gVar.WindowHeight Or Height > Screen.Height Then Height = gVar.WindowHeight
        End With
    
        If blnServer Then
            Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyServerWindowStateMax, 0)
            Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyServerWindowLeft, CStr(Left))
            Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyServerWindowTop, CStr(Top))
            Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyServerWindowWidth, CStr(Width))
            Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyServerWindowHeight, CStr(Height))
        Else
            Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientWindowStateMax, 0)
            Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientWindowLeft, CStr(Left))
            Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientWindowTop, CStr(Top))
            Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientWindowWidth, CStr(Width))
            Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientWindowHeight, CStr(Height))
        End If
    End If
End Sub

Public Sub gsGridPageSet(ByRef gridControl As Control)
    '打印页面设置
    
    Dim blnFlexCell As Boolean
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    If blnFlexCell Then
'''        frmSysPageSet.Show vbModal   '内容较多暂不设置
        gridControl.PrintDialog
    Else
        GoTo LineBreak
    End If
        
    Exit Sub

LineBreak:
    MsgBox "页面设置检测异常，请重试！", vbExclamation
    
End Sub

Public Sub gsGridPrint(ByRef printGrid As Control)
    '打印表格内容
    
    Call gsGridPrintPreview(printGrid)
    
End Sub

Public Sub gsGridPrintPreview(ByRef gridControl As Control)
    '预览表格内容
    
    Dim blnFlexCell As Boolean
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    If blnFlexCell Then
        With gridControl
            With .PageSetup
                .PrintFixedColumn = True
                .PrintFixedRow = True
                .PrintGridlines = True
                .Footer = "第 &P 页 共 &N 页"
                .FooterAlignment = cellCenter
            End With
            .PrintPreview
        End With
    Else
        GoTo LineBreak
    End If
        
    Exit Sub

LineBreak:
    MsgBox "预览页面检测异常，请重试！", vbExclamation
    
End Sub

Public Sub gsGridToExcel(ByRef gridControl As Control, Optional ByVal TimeCol As Long = -1, Optional ByVal TimeStyle As String = "yyyy-MM-dd HH:mm:ss")  '导出至Excel
    '将表格控件中的内容导出至Excel中
    '参数TimeCol：为控件中的时间列的列号，TimeStyle设定格式
    '最好引用Excel对象。运行时电脑上应有MSOFFICE软件。
    
'    Dim xlsOut As Excel.Application    '用这个申明好编程但要引用，编完后改为Object
    Dim xlsOut As Object
'    Dim sheetOut As Excel.Worksheet
    Dim sheetOut  As Object
    Dim blnFlexCell As Boolean
    Dim R As Long, C As Long, I As Long, J As Long
    
    If gridControl Is Nothing Then Exit Sub
    
    On Error Resume Next
    Screen.MousePointer = 13
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    Set xlsOut = CreateObject("Excel.Application")
    xlsOut.Workbooks.Add
    Set sheetOut = xlsOut.ActiveSheet
    
    With gridControl
        R = .Rows
        C = .Cols
        '表格内容复制到Excel中
        If blnFlexCell Then
            For I = 0 To R - 1
                For J = 0 To C - 1
                    sheetOut.Cells(I + 1, J + 1) = .Cell(I, J).Text
                Next
            Next
        Else
            For I = 0 To R - 1
                For J = 0 To C - 1
                    sheetOut.Cells(I + 1, J + 1) = .TextMatrix(I, J)
                Next
            Next
        End If
    End With
    
    With sheetOut
        If TimeCol > -1 Then
            .Columns(TimeCol + 1).NumberFormatLocal = TimeStyle
        End If
        .Range(.Cells(1, 1), .Cells(1, C)).Font.Bold = True '加粗显示(第一行默认标题行)
        .Range(.Cells(1, 1), .Cells(1, C)).Font.Size = 12   '第一行12号字大小
        .Range(.Cells(2, 1), .Cells(R, C)).Font.Size = 10   '第二行以后10号字大小
        .Range(.Cells(1, 1), .Cells(R, C)).HorizontalAlignment = -4108  'xlCenter= -4108(&HFFFFEFF4)   '居中显示
        .Range(.Cells(1, 1), .Cells(R, C)).Borders.Weight = 2   'xlThin=2  '单元格显示黑色线宽
        .Columns.EntireColumn.AutoFit   '自动列宽
        .Rows(1).RowHeight = 23 '第一行行高
    End With
    
    xlsOut.Visible = True   '显示Excel文档
    
    Set sheetOut = Nothing
    Set xlsOut = Nothing
    Screen.MousePointer = 0
    
End Sub

Public Sub gsGridExportTo(ByRef gridControl As FlexCell.Grid, ByVal ExportID As Long, _
        Optional ByVal blnOpenFile As Boolean = True, Optional ByVal blnExportFixedRow As Boolean = True, _
        Optional ByVal blnExportFixedCol As Boolean = True)
    '将FlexcellGrid表格控件中的内容导出为CSV、Excel、HTML、PD、XMLF文件
    
    Dim strFileName As String, strMsg As String, strFileType As String
    Dim K As Long
    Dim blnOK As Boolean
    
    If gridControl Is Nothing Then Exit Sub
    
    For K = 1 To 8
        strFileName = strFileName & gfBackOneChar(udNumber + udUpperCase) '文件名中的8个随机字符，不含小写字母
    Next
    If TypeOf gridControl Is FlexCell.Grid Then '确定提示内容
        Select Case ExportID
            Case gID.SysExportToCSV
                strMsg = "CSV": strFileType = ".csv"
            Case gID.SysExportToHTML
                strMsg = "HTML": strFileType = ".html"
            Case gID.SysExportToPDF
                strMsg = "PDF": strFileType = ".pdf"
            Case gID.SysExportToXML
                strMsg = "XML": strFileType = ".xml"
            Case Else
                strMsg = "Excel": strFileType = ".xls"
        End Select
        
        If Not gfFileRepair(gVar.FolderNameTemp, True) Then
            Call gsAlarmAndLogEx(gVar.FolderNameTemp & "  文件夹创建失败！无法缓存文件！", "导出警告")
            Exit Sub
        End If
        strFileName = gVar.FolderNameTemp & Format(Now, gVar.Formatymdhms & "_") & strFileName & strFileType
        If MsgBox("确定导出当前表格内容为" & strMsg & "文件吗？", vbQuestion + vbOKCancel, "导出询问") = vbCancel Then Exit Sub
        
        Select Case ExportID
            Case gID.SysExportToCSV
                blnOK = gridControl.ExportToCSV(strFileName, blnExportFixedRow, blnExportFixedCol)
            Case gID.SysExportToHTML
                blnOK = gridControl.ExportToHTML(strFileName)
            Case gID.SysExportToPDF
                blnOK = gridControl.ExportToPDF(strFileName)
            Case gID.SysExportToXML
                blnOK = gridControl.ExportToXML(strFileName)
            Case Else
                blnOK = gridControl.ExportToExcel(strFileName, blnExportFixedRow, blnExportFixedCol)
        End Select
        If blnOK Then
            If blnOpenFile Then Call gfFileOpen(strFileName)    '打开文件
        End If
    End If

End Sub

Public Sub gsGridToText(ByRef gridControl As Control)
    '将传入的表格控件中的内容导出为文本文件
    
    Dim strFileName As String
    Dim blnFlexCell As Boolean
    Dim intFree As Integer
    Dim R As Long, C As Long, I As Long, J As Long
    Dim strTxt As String
    
    If gridControl Is Nothing Then Exit Sub
    
    For I = 1 To 8
        strFileName = strFileName & gfBackOneChar(udNumber + udUpperCase) '文件名中的8个随机字符，不含小写字母
    Next
    strFileName = gVar.FolderNameTemp & Format(Now, gVar.Formatymdhms & "_") & strFileName & ".txt"
    If Not gfFileRepair(strFileName) Then
        Call gsAlarmAndLogEx("创建" & strFileName & "文件失败！", "文件生成警告")
        Exit Sub
    End If
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    intFree = FreeFile
    Open strFileName For Output As #intFree
    With gridControl
        R = .Rows - 1
        C = .Cols - 1
        If blnFlexCell Then
            For I = 0 To R
                strTxt = ""
                For J = 0 To C
                    strTxt = strTxt & .Cell(I, J).Text & vbTab
                Next
                Print #intFree, strTxt
            Next
        End If
    End With
    
    Close
    
    Call gfFileOpen(strFileName)    '打开
    
End Sub


Public Sub gsGridToWord(ByRef gridControl As Control)
    '将指定表格中的内容导出至Word文档中
    
'    Dim wordApp As Word.Application
    Dim wordApp As Object
'    Dim docOut As Word.Document
    Dim docOut As Object
'    Dim tbOut As Word.Table
    Dim tbOut As Object
    Dim lngRows As Long, lngCols As Long
    Dim I As Long, J As Long
    Dim blnFlexCell As Boolean
    Dim strFileName As String
    
    If gridControl Is Nothing Then Exit Sub
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    lngRows = gridControl.Rows
    lngCols = gridControl.Cols
    
    On Error Resume Next
        
    Set wordApp = CreateObject("Word.Application")
    Set docOut = wordApp.Documents.Add()
    
    If blnFlexCell Then
        If gridControl.PageSetup.Orientation = cellLandscape Then
            docOut.Range.PageSetup.Orientation = wdOrientLandscape '表格预览为横向则设置纸张为横向
        End If
        Set tbOut = docOut.Tables.Add(docOut.Range, lngRows, lngCols, True)
        
        For I = 0 To lngRows - 1
            For J = 0 To lngCols - 1
                tbOut.Cell(I + 1, J + 1).Range.Text = gridControl.Cell(I, J).Text
            Next
            If Len(gridControl.Cell(I, 0).Text) = 0 Then tbOut.Cell(I + 1, 1).Range.Text = I
        Next
    Else
        For I = 0 To lngRows - 1
            For J = 0 To lngCols - 1
                tbOut.Cell(I + 1, J + 1).Range.Text = gridControl.TextMatrix(I, J)
            Next
        Next
    End If
    tbOut.Rows(1).Range.Bold = True             '第一行内容加粗
    tbOut.Range.ParagraphFormat.Alignment = 1   '表格内容居中显示
    Call tbOut.AutoFitBehavior(1)               '根据内容自动调整列宽
    
    For I = 1 To 8
        strFileName = strFileName & gfBackOneChar(udNumber + udUpperCase) '文件名中的8个随机字符，不含小写字母
    Next
    strFileName = gVar.FolderNameTemp & Format(Now, gVar.Formatymdhms & "_") & strFileName & ".doc"
    If gfFileRepair(strFileName) Then
        docOut.SaveAs strFileName   '另存为
    Else
        Call gsAlarmAndLogEx("创建" & strFileName & "文件失败！", "文件生成警告")
    End If
    
    wordApp.Visible = True  '显示文档
    wordApp.Activate    '顶层显示
    
    Set tbOut = Nothing
    Set docOut = Nothing
    Set wordApp = Nothing
    
End Sub

Public Sub gsLoadSkin(ByRef frmCur As Form, ByRef skFRM As XtremeSkinFramework.SkinFramework, _
    Optional ByVal lngResource As genumSkinResChoose = -1, Optional ByVal blnFromReg As Boolean = False, _
    Optional ByVal strRes As String = "", Optional ByVal strIni As String = "", _
    Optional ByVal blnServer As Boolean = False)
    '加载主题
    
    If blnFromReg Then  '如果从注册表中获取资源文件，则按注册表中值修改lngResource的值
        If blnServer Then   '服务端程序主题设置
            strRes = GetSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinSvrRes, "")
            strIni = GetSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinSvrIni, "")
        Else    '客户端程序主题设置
            strRes = GetSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinRes, "")
            strIni = GetSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinIni, "")
        End If
    Else
        If lngResource >= sNone And lngResource <= sMSO10 Then
            Select Case lngResource '选择窗口风格资源文件
                Case sMSO7
                    strRes = gVar.FolderNameBin & "cjstylesO7.cjstyles"
                    strIni = "NormalBlue.ini"   'NormalBlue LightBlue NormalBlack NormalSilver NormalAqua
                Case sMSO10
                    strRes = gVar.FolderNameBin & "cjstylesO10.cjstyles"
                    strIni = "NormalBlue.ini"   'NormalBlue NormalBlack NormalSilver
                Case sMSVst
                    strRes = gVar.FolderNameBin & "cjstylesOvst.cjstyles"
                    strIni = "NormalBlue.ini"   'NormalBlue NormalBlack NormalSilver NormalBlack2
                Case Else
                    strRes = gVar.FolderNameBin & "cjstyles.cjstyles"
                    strIni = "NormalBlue.ini"
            End Select
        End If
    End If
    
    With skFRM
        .LoadSkin strRes, strIni
'''        .ApplyOptions = .ApplyOptions Or xtpSkinApplyMetrics Or xtpSkinApplyMenus   '全部应用
        .ApplyOptions = xtpSkinApplyMenus Or xtpSkinApplyColors Or xtpSkinApplyMetrics  '如果添加xtpSkinApplyFrame，鼠标滚轮不能控制FC表格滚动条
        .ApplyWindow frmCur.hwnd
    End With
    
    If blnServer Then   '将设置的主题信息保存进注册表
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinSvrRes, strRes)
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinSvrIni, strIni)
    Else
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinRes, strRes)
        Call SaveSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinIni, strIni)
    End If
End Sub

Public Sub gsLogAdd(ByRef frmCur As Form, Optional ByVal LogType As genumLogType = udSelect, _
    Optional ByVal strTable As String = "", Optional ByVal strContent As String = "")
    '添加操作日志
    
    Dim strType As String
    Dim strSQL As String
    Dim rsLog As ADODB.Recordset
    
    strType = gfBackLogType(LogType)
    
    strSQL = "EXEC sp_FT_Sys_LogAdd '" & strType & "','" & frmCur.Name & "," & frmCur.Caption & "','" & strTable & _
             "','" & strContent & "','" & gVar.UserLoginName & "," & gVar.UserFullName & "','" & gVar.UserLoginIP & "','" & gVar.UserComputerName & "'"
'Debug.Print strSQL
    Set rsLog = gfBackRecordset(strSQL, , adLockOptimistic)
    If rsLog.State = adStateOpen Then rsLog.Close
    Set rsLog = Nothing
    
End Sub

Public Sub gsNodeCheckCascade(ByRef nodeCheck As MSComctlLib.Node, Optional ByVal blnCheck As Boolean)
    '结点的Checked属性级联变化
    
    If blnCheck Then    '=False时不处理
        Call gsNodeCheckUp(nodeCheck)
    End If
    
    Call gsNodeCheckDown(nodeCheck, blnCheck)
    
End Sub

Public Sub gsNodeCheckDown(ByRef nodeCheck As MSComctlLib.Node, Optional ByVal blnCheck As Boolean)
    '不/勾选结点的所有子结点
    
    Dim nodeSon As MSComctlLib.Node
    Dim C As Long, K As Long
    
    C = nodeCheck.Children
    If C > 0 Then
        For K = 1 To C
            If K = 1 Then
                Set nodeSon = nodeCheck.Child
            Else
                Set nodeSon = nodeSon.Next
            End If
            If nodeSon.Checked <> blnCheck Then nodeSon.Checked = blnCheck
            If nodeSon.Children > 0 Then
                Call gsNodeCheckDown(nodeSon, blnCheck)
            End If
        Next
    End If
    
End Sub

Public Sub gsNodeCheckUp(ByRef nodeCheck As MSComctlLib.Node, Optional ByVal blnCheck As Boolean = True)
    '勾选结点的所有父结点
    
    Dim nodeDad As MSComctlLib.Node
    
    If Not nodeCheck.Parent Is Nothing Then
        Set nodeDad = nodeCheck.Parent
        If Not nodeDad.Checked Then nodeDad.Checked = blnCheck
        If Not nodeDad.Parent Is Nothing Then
            Call gsNodeCheckUp(nodeDad)
        End If
    End If
    
End Sub


Public Sub gsOpenTheWindow(ByVal strFormName As String, _
    Optional ByVal OpenMode As FormShowConstants = vbModeless, _
    Optional ByVal FormWndState As FormWindowStateConstants = vbMaximized, _
    Optional ByVal UseMainIcon As Boolean = True)
    '以指定窗口模式OpenMode与窗口FormWndState状态来打开指定窗体strFormName
    
    Dim frmOpen As Form
    Dim C As Long
    
    strFormName = LCase(strFormName)
    If gfFormLoad(strFormName) Then '窗体已存在
        For C = 0 To Forms.Count - 1
            If LCase(Forms(C).Name) = strFormName Then
                Set frmOpen = Forms(C)  '引用该窗体
                Exit For
            End If
        Next
    Else    '窗体不存在
        Set frmOpen = Forms.Add(strFormName)    '新建该窗体
    End If
    
    If UseMainIcon Then
        If frmOpen.Icon Is Nothing Then
            Set frmOpen.Icon = gWind.Icon   '使用主窗体图标
        End If
    End If
    frmOpen.WindowState = FormWndState
    frmOpen.Show OpenMode               '此句放最后，不能放上句前面，否则退出程序时MDI窗体不能完全关闭，可能因为CommandBars控件的原因。
        
End Sub

Public Sub gsSaveCommandbarsTheme(ByRef cbsBars As XtremeCommandBars.CommandBars, Optional ByVal blnServer As Boolean = True)
    '保存CommandBars的风格主题
    Dim lngID As Long
    
    For lngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        If cbsBars.Actions(lngID).Checked Then Exit For
    Next
    If lngID > gID.WndThemeCommandBarsWinXP Then lngID = gID.WndThemeCommandBarsRibbon
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, IIf(blnServer, gVar.RegKeyServerCommandbarsTheme, gVar.RegKeyClientCommandbarsTheme), lngID)
    
End Sub

Public Sub gsStartProgressBar(ByVal CurVal As Long, Optional ByVal MinVal As Long = 0, Optional ByVal MaxVal As Long = 100)
    '主窗体状态栏中的进度条显示进度、百分值
    
    Dim cbsBars As XtremeCommandBars.CommandBars
    Dim paneBar As XtremeCommandBars.StatusBarProgressPane
    Dim PaneTxt As XtremeCommandBars.StatusBarPane
    
    Set cbsBars = gWind.CommandBars1
    Set paneBar = cbsBars.StatusBar.FindPane(gID.StatusBarPaneProgress)
    Set PaneTxt = cbsBars.StatusBar.FindPane(gID.StatusBarPaneProgressText)
    With paneBar
        .Min = MinVal
        .Max = MaxVal
        .Value = CurVal
    End With
    PaneTxt.Text = CStr(CurVal / MaxVal * 100) & "%"
    
    Set paneBar = Nothing
    Set PaneTxt = Nothing
    Set cbsBars = Nothing
End Sub

Public Sub gsThemeCommandBar(ByVal CID As Long, ByRef cbsBars As XtremeCommandBars.CommandBars)
    'CommandBars风格设置
    Dim lngTheme As Long, lngID As Long
    Dim blnChangeSkin As Boolean
    
    Select Case CID
        Case gID.WndThemeCommandBarsOffice2000
            lngTheme = xtpThemeOffice2000
        Case gID.WndThemeCommandBarsOffice2003
            lngTheme = xtpThemeOffice2003
            blnChangeSkin = True
        Case gID.WndThemeCommandBarsOfficeXp
            lngTheme = xtpThemeOfficeXP
        Case gID.WndThemeCommandBarsResource
            lngTheme = xtpThemeResource
            blnChangeSkin = True
        Case gID.WndThemeCommandBarsRibbon
            lngTheme = xtpThemeRibbon: blnChangeSkin = True
        Case gID.WndThemeCommandBarsVS2008
            lngTheme = xtpThemeVisualStudio2008
        Case gID.WndThemeCommandBarsVS2010
            lngTheme = xtpThemeVisualStudio2010
        Case gID.WndThemeCommandBarsVS6
            lngTheme = xtpThemeVisualStudio6
        Case gID.WndThemeCommandBarsWhidbey
            lngTheme = xtpThemeWhidbey
        Case Else   'gID.WndThemeCommandBarsWinXP
            lngTheme = xtpThemeNativeWinXP
            CID = gID.WndThemeCommandBarsWinXP
    End Select
    
    cbsBars.VisualTheme = lngTheme
    
    For lngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        cbsBars.Actions(lngID).Checked = False
    Next
    cbsBars.Actions(CID).Checked = True
    
'''    If blnChangeSkin Then   '更改对应窗口主题使颜色统一。注释于2019.5.28会影响加载主题
'''        Call gsLoadSkin(gWind, gWind.SkinFramework1, sMSO7)
'''    Else
'''        Call gsLoadSkin(gWind, gWind.SkinFramework1, sMSVst)
'''    End If
    
End Sub

Public Sub gsUnCheckedAction(ByVal strFormName As String, cbsBars As XtremeCommandBars.CommandBars)
    '当窗口关闭时，去掉主窗体中CommandBars控件中被勾选的对应Action
    
    Dim actionCur As XtremeCommandBars.CommandBarAction
    
    strFormName = LCase(strFormName)
    For Each actionCur In cbsBars.Actions
        If Len(actionCur.Key) > 0 Then
            If LCase(actionCur.Key) = strFormName Then
                actionCur.Checked = False
                Exit For
            End If
        End If
    Next
    
End Sub


