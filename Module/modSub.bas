Attribute VB_Name = "modSub"
Option Explicit


Public Sub Main()
    
    Dim strTemp As String
    
    '������CommandBars��IDֵ��ʼ��
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
        
        .SysExportToExcel = 1201
        .SysExportToPDF = 1202
        .SysExportToText = 1203
        .SysExportToWord = 1204
        .SysExportToXML = 1205
        
        .SysPrint = 1303
        .SysPrintPageSet = 1301
        .SysPrintPreview = 1302
                
        
        .Wnd = 2000
        
        .WndResetLayout = 2050
        
        .TabWorkspacePopupMenu = 2051
        
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
        
        .WndThemeSkinSet = 2450
        
        
        .Help = 3000
        .HelpAbout = 3101
        .HelpDocument = 3102
        .HelpUpdate = 3103
        
        
        
        '''***�뽫���в˵����еġ��˵�����CommandBrs��IDֵ������20000����*******************
        
        
        .Pane = 21000
        
        .PaneIDFirst = 21101
        .PaneTitleFirst = 21102
        
        .PanePopupMenu = 21103
        .PanePopupMenuAutoFoldOther = 21104
        .PanePopupMenuExpandALL = 21105
        .PanePopupMenuFoldALL = 21106
        
        
        .StatusBarPane = 22000
        
        .StatusBarPaneConnectButton = 22101
        .StatusBarPaneConnectState = 22102
        .StatusBarPaneProgress = 22103
        .StatusBarPaneProgressText = 22104
        .StatusBarPaneServerButton = 22105
        .StatusBarPaneServerState = 22106
        .StatusBarPaneTime = 22107
        .StatusBarPaneUserInfo = 22108
        
        .IconPopupMenu = 23000
        .IconPopupMenuMaxWindow = 23101
        .IconPopupMenuMinWindow = 23102
        .IconPopupMenuShowWindow = 23103
        
    End With
    
    '���ñ���ֵ��ʼ��
    With gVar
        
        .TCPSetConnectMax = 20
        .TCPSetIP = "127.0.0.1"
        .TCPSetPort = 9898
        
        .FTChunkSize = 5734
        .FTWaitTime = 5
        
        .ServerClose = "�رշ���"
        .ServerError = "�쳣"
        .ServerNotStarted = "δ����"
        .ServerStart = "��������"
        .ServerStarted = "������"
        
        .StateConnected = "������"
        .StateDisConnected = "δ����"
        .StateConnectError = "�����쳣"
        .StateConnectToServer = "��������"
        .StateDisConnectFromServer = "�Ͽ�����"
        
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
        .PTWaitTime = 2
        
        .EXENameOfClient = "FFC.exe"
        .EXENameOfServer = "FFS.exe"
        .EXENameOfSetup = "FFSetup.exe"
        .EXENameOfUpdate = "FFU.exe"
        
        .CmdLineParaOfHide = "Hide"
        .CmdLineSeparator = " / "
        
        .RegAppName = "FF"
        .RegKeyTCPIP = "IP"
        .RegKeyTCPPort = "Port"
        .RegSectionTCP = "TCP"
        
        .RegSectionSkin = "SkinFile"
        .RegKeySkinFile = "SkinRes"
        
        .RegSectionServer = "Server"
        .RegKeyServerAccount = "ServerAccount"
        .RegKeyServerIP = "ServerIP"
        .RegKeyServerPassword = "ServerPassword"
        
        .RegSectionUser = "UserInfo"
        .RegKeyUserLast = "LastLoginUser"
        .RegKeyUserList = "LoginUserList"
        
        .RegSectionSettings = "Settings"
        .RegKeyCommandBars = "cbs"
        .RegKeyWindowHeight = "WindowHeight"
        .RegKeyWindowLeft = "WindowLeft"
        .RegKeyWindowTop = "WindowTop"
        .RegKeyWindowWidth = "WindowWidth"
        .RegKeyCommandbarsTheme = "cbsTheme"
        
        .RegTrailPath = "SoftWare\Common\Section"   'HKEY_CURRENT_USER\SoftWare\����
        .RegTrailKey = "Key"
        .TrailPeriod = 15
        
        .RegKeyParaWindowMinHide = "WindowMinHide"
        
        .AppPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
        
        .FolderNameBin = .AppPath & "Bin\"
        .FolderNameData = .AppPath & "Data\"
        .FolderNameTemp = .AppPath & "Temp\"
        
        .FileNameErrLog = .FolderNameData & "ErrorRecord.LOG"
        .FileNameSkin = ""
        .FileNameSkinIni = ""
        
        .AccountAdmin = "Admin"     '���������û�
        .AccountSystem = "System"   '���������û�
        
        .FuncButton = "��ť"
        .FuncControl = "����"
        .FuncForm = "����"
        .FuncMainMenu = "���˵�"
        
        .WindowHeight = 8700
        .WindowWidth = 15800
        
        '''*****��ע����б����������ַ�����ʵ��˺�������****
        strTemp = GetSetting(.RegAppName, .RegSectionServer, .RegKeyServerIP)
        .ConSource = gfCheckIP(strTemp)
        
        strTemp = GetSetting(.RegAppName, .RegSectionServer, .RegKeyServerAccount, "")
        If Len(strTemp) > 0 Then strTemp = gfDecryptSimple(strTemp)
        .ConUserID = strTemp
        
        strTemp = GetSetting(.RegAppName, .RegSectionServer, .RegKeyServerPassword, "")
        If Len(strTemp) > 0 Then strTemp = gfDecryptSimple(strTemp)
        .ConPassword = strTemp
        
        .ConDatabase = "db_Test"    '�ݽ�������SQLServer2008 OR 2012 ���ݿ�
        .ConString = "Provider=SQLOLEDB;Persist Security Info=False;Data Source=" & .ConSource & _
                    ";UID=" & .ConUserID & ";PWD=" & .ConPassword & _
                    ";DataBase=" & .ConDatabase & ";"   '''��64λϵͳ��Data Source�м�Ҫ�ո�������ܽ�������
        
    End With
    
End Sub


Public Sub gsAlarmAndLog(Optional ByVal strErr As String, Optional ByVal blnMsgBox As Boolean = True, Optional ByVal MsgButton As VbMsgBoxStyle = vbCritical)
    '�쳣��ʾ��д���쳣��־
    
    Dim strMsg As String
    
    strMsg = "�쳣���ţ�" & Err.Number & vbCrLf & "�쳣������" & Err.Description
    If blnMsgBox Then MsgBox strMsg, MsgButton, strErr
    Call gsFileWrite(gVar.FileNameErrLog, strErr & vbTab & Replace(strMsg, vbCrLf, vbTab))
    
End Sub


Public Sub gsFileWrite(ByVal strFile As String, ByVal strContent As String, _
    Optional ByVal OpenMode As genumFileOpenType = udAppend, _
    Optional ByVal WriteMode As genumFileWriteType = udPrint)
    '��ָ��������ָ���ķ�ʽд��ָ���ļ���
    
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
        Case Else   '����Ե���udAppend
            Open strFile For Append As #intNum
    End Select
    
    strTime = Format(Now, "yyyy-MM-dd hh:mm:ss")
    Select Case WriteMode
        Case udWrite
            Write #intNum, strTime, strContent
        Case udPut
            Put #intNum, , strTime & vbTab & strContent
        Case Else   '����Ե���udPrint
            Print #intNum, strTime, strContent
    End Select
    
    Close #intNum
    
End Sub


Public Sub gsFormScrollBar(ByRef frmCur As Form, ByRef ctlMv As Control, _
    ByRef Hsb As HScrollBar, ByRef Vsb As VScrollBar, _
    Optional ByVal lngMW As Long = 12000, _
    Optional ByVal lngMH As Long = 9000, _
    Optional ByVal lngHV As Long = 255)
    
    'frmCur�����������ڵĴ���
    'ctlMv�������еĿؼ��������������⣩���ڴ������ؼ���
    'Hsb������frmCur��ˮƽ�������ؼ�
    'Vsb������frmCur�д�ֱ�������ؼ�
    'lngMW�����岻���ֹ������Ŀ��
    'lngMH�����岻���ֹ������ĸ߶�
    'lngHV����������խ�߿�Ȼ�߶ȡ�
    '***ע��ע��ע�⣺�������ؼ����������������У��Ҳ��ܷ��������ؼ�ctlMv��*******
    
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
    
'    '�ڴ�������Ӵ��ڿؼ�ctlMove�������������ؼ�����������У�Ȼ
'    '��������Ʒֱ�ΪHsb\Vsb��ˮƽ\��ֱ�������ڴ����У�������������봰����
'    'Ȼ���ڴ�������������¼����ü���
'Private Sub Form_Resize()
'    Call gsFormScrollBar(Me, Me.ctlMove, Me.Hsb, Me.Vsb, 12000, 9000)  'ע�ⳤ������޸�
'End Sub
'Private Sub Hsb_Change()
'    ctlMove.Left = -Hsb.Value
'End Sub
'
'Private Sub Hsb_Scroll()
'    Call Hsb_Change    '�������������еĻ���ʱ��ͬʱ���¶�Ӧ���ݣ�����ͬ��
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

Public Sub gsFormSizeLoad(ByRef frmLoad As Form)
    '��ע����м��ش��ڵ�λ�����С��Ϣ
    Dim Left As Long, Top As Long, Width As Long, Height As Long
    
    Left = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowLeft, 0))
    If Left < 0 Or Left > Screen.Width Then Left = 0
    Top = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowTop, 0))
    If Top < 0 Or Left > Screen.Height Then Top = 0
    Width = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowWidth, gVar.WindowWidth))
    If Width <= 0 Or Width > Screen.Width Then Width = gVar.WindowWidth
    Height = Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowHeight, gVar.WindowHeight))
    If Height <= 0 Or Height > Screen.Height Then Height = gVar.WindowHeight
    frmLoad.Move Left, Top, Width, Height
    
End Sub

Public Sub gsFormSizeSave(ByRef frmSave As Form)
    '���洰�ڵ�λ�����С��Ϣ��ע�����
    Dim Left As Long, Top As Long, Width As Long, Height As Long
    
    With frmSave
        Left = .Left
        Top = .Top
        Width = .Width
        Height = .Height
        If Left < 0 Or Left > Screen.Width Then Left = 0
        If Top < 0 Or Top > Screen.Height Then Top = 0
        If Width > Screen.Width Then Width = gVar.WindowWidth
        If Height > Screen.Height Then Height = gVar.WindowHeight
    End With
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowLeft, CStr(Left))
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowTop, CStr(Top))
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowWidth, CStr(Width))
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyWindowHeight, CStr(Height))
    
End Sub

Public Sub gsGridPageSet()
    '��ӡҳ������
    
    Dim gridControl As Control
    Dim blnFlexCell As Boolean
    
    If Screen.ActiveForm Is Nothing Then GoTo LineBreak
    If Screen.ActiveForm.ActiveControl Is Nothing Then GoTo LineBreak
    
    Set gridControl = Screen.ActiveForm.ActiveControl
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    If blnFlexCell Then
'''        frmSysPageSet.Show vbModal   '���ݽ϶��ݲ�����
        gridControl.PrintDialog
    Else
        GoTo LineBreak
    End If
        
    Exit Sub

LineBreak:
    MsgBox "ҳ�����ü���쳣�������ԣ�", vbExclamation
    
End Sub

Public Sub gsGridPrint()
    '��ӡ�������
    
    Call gsGridPrintPreview
    
End Sub

Public Sub gsGridPrintPreview()
    'Ԥ���������
    
    Dim gridControl As Control
    Dim blnFlexCell As Boolean
    
    If Screen.ActiveForm Is Nothing Then GoTo LineBreak
    If Screen.ActiveForm.ActiveControl Is Nothing Then GoTo LineBreak
    
    Set gridControl = Screen.ActiveForm.ActiveControl
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True

    If blnFlexCell Then
        With gridControl
            With .PageSetup
                .PrintFixedColumn = True
                .PrintFixedRow = True
                .PrintGridlines = True
                .Footer = "�� &P ҳ �� &N ҳ"
                .FooterAlignment = cellCenter
            End With
            .PrintPreview
        End With
    Else
        GoTo LineBreak
    End If
        
    Exit Sub

LineBreak:
    MsgBox "Ԥ��ҳ�����쳣�������ԣ�", vbExclamation
    
End Sub

Public Sub gsGridToExcel(ByRef gridControl As Control, Optional ByVal TimeCol As Long = -1, Optional ByVal TimeStyle As String = "yyyy-MM-dd HH:mm:ss")  '������Excel
    '�����ؼ��е����ݵ�����Excel��
    '����TimeCol��Ϊ�ؼ��е�ʱ���е��кţ�TimeStyle�趨��ʽ
    '�������Excel��������ʱ������Ӧ��MSOFFICE�����
    
'    Dim xlsOut As Excel.Application    '����������ñ�̵�Ҫ���ã�������ΪObject
    Dim xlsOut As Object
'    Dim sheetOut As Excel.Worksheet
    Dim sheetOut  As Object
    Dim blnFlexCell As Boolean
    Dim R As Long, C As Long, I As Long, J As Long
    
    On Error Resume Next
    Screen.MousePointer = 13
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    Set xlsOut = CreateObject("Excel.Application")
    xlsOut.Workbooks.Add
    Set sheetOut = xlsOut.ActiveSheet
    
    With gridControl
        R = .Rows
        C = .Cols
        '������ݸ��Ƶ�Excel��
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
        .Range(.Cells(1, 1), .Cells(1, C)).Font.Bold = True '�Ӵ���ʾ(��һ��Ĭ�ϱ�����)
        .Range(.Cells(1, 1), .Cells(1, C)).Font.Size = 12   '��һ��12���ִ�С
        .Range(.Cells(2, 1), .Cells(R, C)).Font.Size = 10   '�ڶ����Ժ�10���ִ�С
        .Range(.Cells(1, 1), .Cells(R, C)).HorizontalAlignment = -4108  'xlCenter= -4108(&HFFFFEFF4)   '������ʾ
        .Range(.Cells(1, 1), .Cells(R, C)).Borders.Weight = 2   'xlThin=2  '��Ԫ����ʾ��ɫ�߿�
        .Columns.EntireColumn.AutoFit   '�Զ��п�
        .Rows(1).RowHeight = 23 '��һ���и�
    End With
    
    xlsOut.Visible = True   '��ʾExcel�ĵ�
    
    Set sheetOut = Nothing
    Set xlsOut = Nothing
    Screen.MousePointer = 0
    
End Sub


Public Sub gsGridToText(ByRef gridControl As Control)
    '������ı��ؼ��е����ݵ���Ϊ�ı��ļ�
    
    Dim strFileName As String
    Dim blnFlexCell As Boolean
    Dim intFree As Integer
    Dim R As Long, C As Long, I As Long, J As Long
    Dim strTxt As String
    
    For I = 1 To 8
        strFileName = strFileName & gfBackOneChar(udNumber + udUpperCase) '�ļ����е�8������ַ�������Сд��ĸ
    Next
    strFileName = gVar.FolderNameData & Format(Now, "yyyyMMddHHmmss_") & strFileName & ".txt"
    If Not gfFileRepair(strFileName) Then
        MsgBox "�����ļ�ʧ�ܣ������ԣ�", vbExclamation, "�ļ����ɾ���"
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
    
    Call gfFileOpen(strFileName)    '��
    
End Sub


Public Sub gsGridToWord(ByRef gridControl As Control)
    '��ָ������е����ݵ�����Word�ĵ���
    
'    Dim wordApp As Word.Application
    Dim wordApp As Object
'    Dim docOut As Word.Document
    Dim docOut As Object
'    Dim tbOut As Word.Table
    Dim tbOut As Object
    Dim lngRows As Long, lngCols As Long
    Dim I As Long, J As Long
    Dim blnFlexCell As Boolean
    
    lngRows = gridControl.Rows
    lngCols = gridControl.Cols
    
    On Error Resume Next
'    Set wordApp = New Word.Application
    Set wordApp = CreateObject("Word.Application")
    Set docOut = wordApp.Documents.Add()
    Set tbOut = docOut.Tables.Add(docOut.Range, lngRows, lngCols, True)
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    If blnFlexCell Then
        For I = 0 To lngRows - 1
            For J = 0 To lngCols - 1
                tbOut.Cell(I + 1, J + 1).Range.Text = gridControl.Cell(I, J).Text
            Next
        Next
    Else
        For I = 0 To lngRows - 1
            For J = 0 To lngCols - 1
                tbOut.Cell(I + 1, J + 1).Range.Text = gridControl.TextMatrix(I, J)
            Next
        Next
    End If
    tbOut.Rows(1).Range.Bold = True             '��һ�����ݼӴ�
    tbOut.Range.ParagraphFormat.Alignment = 1   '������ݾ�����ʾ
    Call tbOut.AutoFitBehavior(1)               '���������Զ������п�
    
    wordApp.Visible = True
    
    Set tbOut = Nothing
    Set docOut = Nothing
    Set wordApp = Nothing
    
End Sub

Public Sub gsLoadAuthority(ByRef frmCur As Form, ByRef ctlCur As Control)
    '���ش����еĿ���Ȩ��
    
    Dim strUser As String, strForm As String, strCtlName As String
    
    strUser = LCase(gVar.UserLoginName)
    strForm = LCase(frmCur.Name)
    strCtlName = LCase(ctlCur.Name)
    
    If strUser = LCase(gVar.AccountAdmin) Or strUser = LCase(gVar.AccountSystem) Then Exit Sub
    ctlCur.Enabled = False
    
    With gVar.rsURF
        If .State = adStateOpen Then
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If strForm = LCase(.Fields("FuncFormName")) Then
                        If strCtlName = LCase(.Fields("FuncName")) Then
                            ctlCur.Enabled = True
                            Exit Do
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
    
End Sub

Public Sub gsLoadSkin(ByRef frmCur As Form, ByRef skFRM As XtremeSkinFramework.SkinFramework, _
    Optional ByVal lngResource As genumSkinResChoose, Optional ByVal blnFromReg As Boolean)
    '��������
    Dim lngReg As Long, strRes As String, strIni As String
    
    lngReg = GetSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinFile, 0)
    If blnFromReg Then  '�����ע����л�ȡ��Դ�ļ�����ע�����ֵ�޸�lngResource��ֵ
        If lngReg > sMSO10 Or lngReg < sNone Then lngReg = sNone
        lngResource = lngReg
    End If
    
    Select Case lngResource 'ѡ�񴰿ڷ����Դ�ļ�
        Case sMSO7
            strRes = gVar.FolderNameBin & "cjstylesO7.dll"
            strIni = "NormalBlue.ini"   'NormalBlue LightBlue NormalBlack NormalSilver NormalAqua
        Case sMSO10
            strRes = gVar.FolderNameBin & "cjstylesO10.dll"
            strIni = "NormalBlue.ini"   'NormalBlue NormalBlack NormalSilver
        Case sMSVst
            strRes = gVar.FolderNameBin & "cjstylesOvst.dll"
            strIni = "NormalBlue.ini"   'NormalBlue NormalBlack NormalSilver NormalBlack2
        Case Else
    End Select
    
    With skFRM
        .LoadSkin strRes, strIni
'''        .ApplyOptions = .ApplyOptions Or xtpSkinApplyMetrics Or xtpSkinApplyMenus   'ȫ��Ӧ��
        .ApplyOptions = xtpSkinApplyMenus Or xtpSkinApplyColors Or xtpSkinApplyMetrics  '������xtpSkinApplyFrame�������ֲ��ܿ���FC��������
        .ApplyWindow frmCur.hwnd
    End With
    
    If lngReg <> lngResource Then Call SaveSetting(gVar.RegAppName, gVar.RegSectionSkin, gVar.RegKeySkinFile, lngResource)
    
End Sub

Public Sub gsLogAdd(ByRef frmCur As Form, Optional ByVal LogType As genumLogType = udSelect, _
    Optional ByVal strTable As String = "", Optional ByVal strContent As String = "")
    '��Ӳ�����־
    
    Dim strType As String
    Dim strSQL As String
    Dim rsLog As ADODB.Recordset
    
    strType = gfBackLogType(LogType)
    
    strSQL = "EXEC sp_Test_Sys_LogAdd '" & strType & "','" & frmCur.Name & "," & frmCur.Caption & "','" & strTable & _
             "','" & strContent & "','" & gVar.UserLoginName & "," & gVar.UserFullName & "','" & gVar.UserLoginIP & "','" & gVar.UserComputerName & "'"
'Debug.Print strSQL
    Set rsLog = gfBackRecordset(strSQL, , adLockOptimistic)
    If rsLog.State = adStateOpen Then rsLog.Close
    Set rsLog = Nothing
    
End Sub

Public Sub gsNodeCheckCascade(ByRef nodeCheck As MSComctlLib.Node, Optional ByVal blnCheck As Boolean)
    '����Checked���Լ����仯
    
    If blnCheck Then    '=Falseʱ������
        Call gsNodeCheckUp(nodeCheck)
    End If
    
    Call gsNodeCheckDown(nodeCheck, blnCheck)
    
End Sub

Public Sub gsNodeCheckDown(ByRef nodeCheck As MSComctlLib.Node, Optional ByVal blnCheck As Boolean)
    '��/��ѡ���������ӽ��
    
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
    '��ѡ�������и����
    
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
    Optional ByVal FormWndState As FormWindowStateConstants = vbMaximized)
    '��ָ������ģʽOpenMode�봰��FormWndState״̬����ָ������strFormName
    
    Dim frmOpen As Form
    Dim C As Long
    
    strFormName = LCase(strFormName)
    If gfFormLoad(strFormName) Then
        For C = 0 To Forms.Count - 1
            If LCase(Forms(C).Name) = strFormName Then
                Set frmOpen = Forms(C)
                Exit For
            End If
        Next
    Else
        Set frmOpen = Forms.Add(strFormName)
    End If
    
    frmOpen.WindowState = FormWndState
    frmOpen.Show OpenMode               '�˾����󣬲��ܷ��Ͼ�ǰ�棬�����˳�����ʱMDI���岻����ȫ�رգ�������ΪCommandBars�ؼ���ԭ��
        
End Sub

Public Sub gsSaveCommandbarsTheme(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '����CommandBars�ķ������
    Dim lngID As Long
    
    For lngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        If cbsBars.Actions(lngID).Checked Then Exit For
    Next
    If lngID > gID.WndThemeCommandBarsWinXP Then lngID = gID.WndThemeCommandBarsRibbon
    Call SaveSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyCommandbarsTheme, lngID)
    
End Sub

Public Sub gsStartUpSet(Optional ByVal blnSet As Boolean = True)
    
    '��������������
    Dim strReg As String, strCur As String
    Dim blnReg As Boolean
    
    If Not blnSet Then Exit Sub
    strCur = Chr(34) & gVar.AppPath & App.EXEName & ".exe" & Chr(34) & "-s"
    blnReg = gfRegOperate(HKEY_LOCAL_MACHINE, HKEY_USER_RUN, App.EXEName, REG_SZ, strReg, RegRead)
    If blnReg Then
        If LCase(strCur) <> LCase(strReg) Then
            blnReg = False
'''Debug.Print LCase(strCur),LCase(strReg)
        End If
    End If
    If Not blnReg Then
        blnReg = gfRegOperate(HKEY_LOCAL_MACHINE, HKEY_USER_RUN, App.EXEName, REG_SZ, strCur, RegWrite)
        If Not blnReg Then
            '��¼���ÿ����Զ�����ʧ��
            Call gsAlarmAndLog("���ÿ����Զ�����ʧ�ܣ�")
        End If
    End If
    
End Sub

Public Sub gsThemeCommandBar(ByVal CID As Long, ByRef cbsBars As XtremeCommandBars.CommandBars)
    'CommandBars�������
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
    End Select
    
    cbsBars.VisualTheme = lngTheme
    
    For lngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        cbsBars.Actions(lngID).Checked = False
    Next
    cbsBars.Actions(CID).Checked = True
    
    If blnChangeSkin Then   '���Ķ�Ӧ��������ʹ��ɫͳһ
        Call gsLoadSkin(gWind, gWind.SkinFramework1, sMSO7)
    Else
        Call gsLoadSkin(gWind, gWind.SkinFramework1, sMSVst)
    End If
    
End Sub

Public Sub gsUnCheckedAction(ByVal strFormName As String, cbsBars As XtremeCommandBars.CommandBars)
    '�����ڹر�ʱ��ȥ����������CommandBars�ؼ��б���ѡ�Ķ�ӦAction
    
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


