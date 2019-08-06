VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.CommandBars.v15.3.1.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form frmSysMain 
   Caption         =   "Main服务端"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9315
   Icon            =   "frmSysMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   9315
   StartUpPosition =   2  '屏幕中心
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   720
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Left            =   1440
      Top             =   3840
   End
   Begin FlexCell.Grid Grid1 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5106
      Cols            =   5
      GridColor       =   12632256
      Rows            =   30
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   68
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":0442
            Key             =   "cNativeWinXP"
            Object.Tag             =   "2110"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":0799
            Key             =   "SysPDF"
            Object.Tag             =   "1204"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":0E6D
            Key             =   "SysXML"
            Object.Tag             =   "1205"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":142E
            Key             =   "cOffice2000"
            Object.Tag             =   "2101"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1725
            Key             =   "cOffice2003"
            Object.Tag             =   "2102"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1CB9
            Key             =   "cOfficeXP"
            Object.Tag             =   "2103"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1FE5
            Key             =   "cResource"
            Object.Tag             =   "2104"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":254C
            Key             =   "cRibbon"
            Object.Tag             =   "2105"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":298D
            Key             =   "cVisualStudio6.0"
            Object.Tag             =   "2108"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":2CD0
            Key             =   "cVisualStudio2008"
            Object.Tag             =   "2106"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":31ED
            Key             =   "cVisualStudio2010"
            Object.Tag             =   "2107"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":3604
            Key             =   "cWhidbey"
            Object.Tag             =   "2109"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":3A4F
            Key             =   "tListView"
            Object.Tag             =   "841"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":3D0D
            Key             =   "tListViewOffice2003"
            Object.Tag             =   "842"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":400D
            Key             =   "tListViewOfficeXP"
            Object.Tag             =   "843"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":42CA
            Key             =   "tNativeWinXP"
            Object.Tag             =   "844"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":4876
            Key             =   "tNativeWinXPPlain"
            Object.Tag             =   "845"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":4CF7
            Key             =   "tOffice2000"
            Object.Tag             =   "846"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":5106
            Key             =   "tOffice2000Plain"
            Object.Tag             =   "847"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":551B
            Key             =   "tOffice2003"
            Object.Tag             =   "848"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":5825
            Key             =   "tOffice2003Plain"
            Object.Tag             =   "849"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":5B28
            Key             =   "tOfficeXPPlain"
            Object.Tag             =   "850"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":5DE4
            Key             =   "tResource"
            Object.Tag             =   "851"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":60DE
            Key             =   "tShortcutBarOffice2003"
            Object.Tag             =   "852"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":63FE
            Key             =   "tToolbox"
            Object.Tag             =   "853"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":66BB
            Key             =   "tToolboxWhidbey"
            Object.Tag             =   "854"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":6A78
            Key             =   "tVisualStudio2010"
            Object.Tag             =   "855"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":6E40
            Key             =   "sCodejock"
            Object.Tag             =   "871"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":7E92
            Key             =   "sOffice2007"
            Object.Tag             =   "872"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":8EE4
            Key             =   "sOffice2010"
            Object.Tag             =   "873"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":9F36
            Key             =   "sOrangina"
            Object.Tag             =   "878"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":AF88
            Key             =   "sVista"
            Object.Tag             =   "874"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":BFDA
            Key             =   "sWinXPLuna"
            Object.Tag             =   "875"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":D02C
            Key             =   "sWinXPRoyale"
            Object.Tag             =   "876"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":E07E
            Key             =   "sZune"
            Object.Tag             =   "877"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":F0D0
            Key             =   ""
            Object.Tag             =   "901"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":F76A
            Key             =   "SysWord"
            Object.Tag             =   "1207"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":10444
            Key             =   "SysText"
            Object.Tag             =   "1206"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1111E
            Key             =   "SysExcel"
            Object.Tag             =   "1202"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":11DF8
            Key             =   "SysSearch"
            Object.Tag             =   "113"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":11F0A
            Key             =   "SysPageSet"
            Object.Tag             =   "1301"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1265C
            Key             =   "SysPreview"
            Object.Tag             =   "1302"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":132AE
            Key             =   "SysPrint"
            Object.Tag             =   "1303"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":13F00
            Key             =   "SysGo"
            Object.Tag             =   "116"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":14BDA
            Key             =   "SysLoginOut"
            Object.Tag             =   "1101"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":158B4
            Key             =   "SysLoginAgain"
            Object.Tag             =   "1102"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1658E
            Key             =   "SysCompany"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":171E0
            Key             =   "SysDepartment"
            Object.Tag             =   "104"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":17E32
            Key             =   "threemen"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":18A84
            Key             =   "SysUser"
            Object.Tag             =   "105"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":196D6
            Key             =   "man"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1A328
            Key             =   "woman"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1AF7A
            Key             =   "SysPassword"
            Object.Tag             =   "102"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1BBCC
            Key             =   ""
            Object.Tag             =   "902"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1BF1E
            Key             =   "themes"
            Object.Tag             =   "801"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1CB70
            Key             =   "SelectedMen"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1D7C2
            Key             =   "unknown"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1D8CC
            Key             =   "SysLog"
            Object.Tag             =   "106"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1E51E
            Key             =   "SysRole"
            Object.Tag             =   "107"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1F170
            Key             =   "RoleSelect"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":1FDC2
            Key             =   "SysFunc"
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":20A14
            Key             =   "FuncHead"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":21666
            Key             =   "FuncSelect"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":222B8
            Key             =   "FuncControl"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":22F0A
            Key             =   "FuncButton"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":2361C
            Key             =   "FuncForm"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":2426E
            Key             =   "FuncMainMenu"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMain.frx":25DC0
            Key             =   "themeSet"
            Object.Tag             =   "802"
         EndProperty
      EndProperty
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   2040
      Top             =   3840
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   2640
      Top             =   3840
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSysMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlngID As Long  '循环变量ID
Dim WithEvents mXtrStatusBar As XtremeCommandBars.StatusBar  '状态栏控件
Attribute mXtrStatusBar.VB_VarHelpID = -1
Dim mcbsPopupIcon As XtremeCommandBars.CommandBar    '托盘图标Pupup菜单
Dim CheckConnectTime As Long
Dim ConfirmTime() As Long
Dim ConfirmOK() As Boolean
Dim CountTime() As Long


Private Sub msAddAction(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建CommandBars的Action
    
    Dim cbsAction As XtremeCommandBars.CommandBarAction
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    
    Set cbsActions = cbsBars.Actions
    cbsBars.EnableActions   '启用CommandBars的Actions集合
    
'    cbsActions.Add "Id", "Caption", "TooltipText", "DescriptionText", "Category"   '范例
    With cbsActions
        .Add gID.Sys, "系统", "", "", "系统"
        
        .Add gID.SysLoginOut, "退出", "", "", ""
        .Add gID.SysLoginAgain, "重启", "", "", ""
        
        .Add gID.SysExportToCSV, "导出至CSV", "", "", ""
        .Add gID.SysExportToExcel, "导出至Excel", "", "", ""
        .Add gID.SysExportToHTML, "导出至HTML", "", "", ""
        .Add gID.SysExportToPDF, "导出至PDF", "", "", ""
        .Add gID.SysExportToText, "导出至txt", "", "", ""
        .Add gID.SysExportToWord, "导出至Word", "", "", ""
        .Add gID.SysExportToXML, "导出至XML", "", "", ""
        
        .Add gID.SysPrint, "打印", "", "", ""
        .Add gID.SysPrintPageSet, "打印页面设置", "", "", ""
        .Add gID.SysPrintPreview, "打印预览", "", "", ""
        
        .Add gID.Wnd, "窗口", "", "", "窗口"
        
        .Add gID.WndThemeSkinSet, "窗口主题设置", "", "", "frmSysThemeSet"
        .Add gID.WndResetLayout, "重置窗口布局", "", "", ""
        
        .Add gID.WndThemeCommandBars, "工具栏主题", "", "", ""
        .Add gID.WndThemeCommandBarsOffice2000, "Office2000", "", "", ""
        .Add gID.WndThemeCommandBarsOffice2003, "Office2003", "", "", ""
        .Add gID.WndThemeCommandBarsOfficeXp, "OfficeXP", "", "", ""
        .Add gID.WndThemeCommandBarsResource, "Resource", "", "", ""
        .Add gID.WndThemeCommandBarsRibbon, "Ribbon", "", "", ""
        .Add gID.WndThemeCommandBarsVS2008, "VisualStudio2008", "", "", ""
        .Add gID.WndThemeCommandBarsVS2010, "VisualStudio2010", "", "", ""
        .Add gID.WndThemeCommandBarsVS6, "VisualStudio6", "", "", ""
        .Add gID.WndThemeCommandBarsWhidbey, "Whidbey", "", "", ""
        .Add gID.WndThemeCommandBarsWinXP, "WinXP", "", "", ""
        
        .Add gID.Help, "帮助", "", "", "帮助"
        .Add gID.HelpAbout, "关于…", "", "", ""
        
        .Add gID.StatusBarPane, "状态栏", "", "", ""
        .Add gID.StatusBarPaneProgress, "进度条", "", "", ""
        .Add gID.StatusBarPaneProgressText, "进度条百分比值", "", "", ""
        .Add gID.StatusBarPaneServerButton, "服务开启/断开按钮", "", "", ""
        .Add gID.StatusBarPaneServerState, "服务状态", "", "", ""
        .Add gID.StatusBarPaneTime, "系统时间", "", "", ""
        .Add gID.StatusBarPaneIP, "本机IP地址", "", "", ""
        .Add gID.StatusBarPanePort, "侦听端口", "", "", ""
        .Add gID.StatusBarPaneReStartButton, "服务自动/手动重启模式切换按钮", "", "", ""
        
        .Add gID.IconPopupMenu, "托盘图标菜单", "", "", ""
        .Add gID.IconPopupMenuMaxWindow, "最大化窗口", "", "", ""
        .Add gID.IconPopupMenuMinWindow, "最小化窗口", "", "", ""
        .Add gID.IconPopupMenuShowWindow, "显示窗口", "", "", ""
        
        .Add gID.Tool, "工具", "", "", "工具"
        .Add gID.toolErrorLog, "系统异常日志", "", "", "frmErrorLog"
        .Add gID.toolLoginLog, "用户登陆日志", "", "", "frmLoginLog"
        .Add gID.toolOptions, "选项", "", "", "frmOption"
        
        
'        .Add gID, "", "", "", ""
        
    End With
    
    '填充cbsActions的其它属性ToolTipText、DescriptionText、Key、Category
    For Each cbsAction In cbsActions
        With cbsAction
            If .ID < 20000 Then
                .ToolTipText = .Caption
                .DescriptionText = .ToolTipText
                .Key = .Category    '为菜单时有特殊用，创建Action时窗体名保存在Category中
                .Category = cbsActions((.ID \ 1000) * 1000).Category
            End If
        End With
    Next
    
    '风格系列的cbsActions的两个属性的描述补充
    With cbsActions
        For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            .Action(mlngID).DescriptionText = .Action(gID.WndThemeCommandBars).Caption & "设置为：" & .Action(mlngID).DescriptionText
            .Action(mlngID).ToolTipText = .Action(mlngID).DescriptionText
        Next
    End With
    
    Set cbsAction = Nothing
    Set cbsActions = Nothing
End Sub

Private Sub msAddConnectToGrid(ByRef sckAdd As MSWinsockLib.Winsock)
    '添加连接信息至表格中
    
    Dim K As Long, C As Long
    
    With Me.Grid1
        .AutoRedraw = False
        C = .Rows - 1
        For K = 1 To C  '寻找可用空行
            If Len(.Cell(K, 1).Text) = 0 Then Exit For
        Next
        If K = C + 1 Then   '表格已满，添加新行
            .Rows = .Rows + 1
        End If
        .Cell(K, 1).Text = sckAdd.RemoteHostIP
        .Cell(K, 2).Text = sckAdd.RemoteHost
        .Cell(K, 5).Text = Format(Now, gVar.Formaty_M_dH_m_s)
        .Cell(K, 6).Text = sckAdd.Index
        .Cell(K, 7).Text = sckAdd.Tag
        .AutoRedraw = True
        .Refresh
    End With
    
End Sub


Private Sub msAddDesignerControls(ByRef cbsBars As XtremeCommandBars.CommandBars)
    'CommandBars自定义对话框中内容项的添加
    
    Dim cbsControls As XtremeCommandBars.CommandBarControls
    Dim cbsAction As XtremeCommandBars.CommandBarAction
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    
    Set cbsActions = cbsBars.Actions
    Set cbsControls = cbsBars.DesignerControls
    For Each cbsAction In cbsActions
        If cbsAction.ID < 20000 Then
            cbsControls.Add xtpControlButton, cbsAction.ID, ""
        End If
    Next
    
    Set cbsControls = Nothing
    Set cbsAction = Nothing
    Set cbsActions = Nothing
End Sub

Private Sub msAddKeyBindings(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建快捷键
    
    With cbsBars.KeyBindings
        .AddShortcut gID.SysLoginOut, "F10"
    End With
    
End Sub

Private Sub msAddMenu(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建菜单栏
    
    Dim cbsMenuBar As XtremeCommandBars.MenuBar
    Dim cbsMenuMain As XtremeCommandBars.CommandBarPopup
    Dim cbsMenuCtrl As XtremeCommandBars.CommandBarControl
    
    Set cbsMenuBar = cbsBars.ActiveMenuBar
    cbsMenuBar.ShowGripper = False  '不显示可拖动的那个点点标记
    cbsMenuBar.EnableDocking xtpFlagStretched     '菜单栏独占一行且不能主动拖动
    
    '系统主菜单
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Sys, "")
    With cbsMenuMain.CommandBar.Controls
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysExportToCSV, "")
        cbsMenuCtrl.BeginGroup = True
        For mlngID = gID.SysExportToExcel To gID.SysExportToWord
            .Add xtpControlButton, mlngID, ""
        Next
        
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysPrintPageSet, "")
        cbsMenuCtrl.BeginGroup = True
        .Add xtpControlButton, gID.SysPrintPreview, ""
        .Add xtpControlButton, gID.SysPrint, ""
        
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysLoginAgain, "")
        cbsMenuCtrl.BeginGroup = True
        .Add xtpControlButton, gID.SysLoginOut, ""
        
    End With
    
    '窗口主菜单
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Wnd, "")
    With cbsMenuMain.CommandBar.Controls
        '窗口主题设置
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.WndThemeSkinSet, "")
        cbsMenuCtrl.BeginGroup = True
        
        '重置布局
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.WndResetLayout, "")
        cbsMenuCtrl.BeginGroup = True
        
        '特殊ID35001自定义工具栏
        Set cbsMenuCtrl = .Add(xtpControlButton, XTP_ID_CUSTOMIZE, "自定义工具栏...")
        cbsMenuCtrl.BeginGroup = True
    
        '特殊ID59392工具栏列表
        Set cbsMenuCtrl = .Add(xtpControlPopup, 0, "工具栏列表")
        cbsMenuCtrl.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, ""
        
        'CommandBars工具栏主题子菜单
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndThemeCommandBars, "")
        With cbsMenuCtrl.CommandBar.Controls
            For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
                .Add xtpControlButton, mlngID, ""
            Next
        End With
    End With
    
    '工具菜单
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Tool, "")
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlButton, gID.toolErrorLog, "")
    cbsMenuCtrl.BeginGroup = True
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlButton, gID.toolLoginLog, "")
    cbsMenuCtrl.BeginGroup = True
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlButton, gID.toolOptions, "")
    cbsMenuCtrl.BeginGroup = True
    
    '帮助主菜单
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Help, "")
    cbsMenuMain.CommandBar.Controls.Add xtpControlButton, gID.HelpAbout, ""
    
    Set cbsMenuBar = Nothing
    Set cbsMenuMain = Nothing
    Set cbsMenuCtrl = Nothing
End Sub

Private Sub msAddXtrStatusBar(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建状态栏
    
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    Dim BarPane As XtremeCommandBars.StatusBarPane
    
    Set cbsActions = cbsBars.Actions
    Set mXtrStatusBar = cbsBars.StatusBar
    With mXtrStatusBar
        .AddPane 0      '系统Pane，显示CommandBarActions的Description
        .SetPaneStyle 0, SBPS_STRETCH
        .SetPaneText 0, "Hello"
        .IdleText = "Hello"
        
        .AddPane gID.StatusBarPaneIP
        .SetPaneText gID.StatusBarPaneIP, Me.Winsock1.Item(0).LocalIP  'gVar.TCPSetIP
        .FindPane(gID.StatusBarPaneIP).Width = 90
        
        .AddPane gID.StatusBarPanePort
        .SetPaneText gID.StatusBarPanePort, gVar.TCPSetPort
        .FindPane(gID.StatusBarPanePort).Width = 60
        
        .AddPane gID.StatusBarPaneReStartButton
        .SetPaneText gID.StatusBarPaneReStartButton, IIf(gVar.ParaBlnAutoReStartServer, "自", "手") & "动重启服务模式"
        .FindPane(gID.StatusBarPaneReStartButton).Width = 120
        .FindPane(gID.StatusBarPaneReStartButton).BackgroundColor = vbCyan
        .FindPane(gID.StatusBarPaneReStartButton).Button = True
        
        .AddPane gID.StatusBarPaneServerState
        .FindPane(gID.StatusBarPaneServerState).Text = gVar.ServerStateNotStarted
        .FindPane(gID.StatusBarPaneServerState).Width = 60
        
        .AddPane gID.StatusBarPaneServerButton
        .FindPane(gID.StatusBarPaneServerButton).Text = gVar.ServerButtonStart
        .FindPane(gID.StatusBarPaneServerButton).Width = 60
        .FindPane(gID.StatusBarPaneServerButton).Button = True
        
        .AddProgressPane gID.StatusBarPaneProgress
                
        .AddPane gID.StatusBarPaneProgressText
        .SetPaneText gID.StatusBarPaneProgressText, "0%"
        .FindPane(gID.StatusBarPaneProgressText).Width = 60
        
        .AddPane 59137  'CapsLock键的状态
        .AddPane 59138  'NumLK键的状态
        .AddPane 59139  'ScrLK键的状态
        .FindPane(0).Caption = "Idle Text"
        .FindPane(59137).Caption = "Caps Lock键状态"
        .FindPane(59138).Caption = "Num LocK键状态"
        .FindPane(59139).Caption = "Scroll LocK键状态"
        
        .Visible = True
        .EnableCustomization True
    End With
    
    For Each BarPane In mXtrStatusBar     '设置Caption、ToolTip、Alignment属性
        If Len(BarPane.Caption) = 0 Then BarPane.Caption = cbsActions(BarPane.ID).Caption
        BarPane.ToolTip = BarPane.Caption
        If BarPane.ID <> 0 Then BarPane.Alignment = xtpAlignmentCenter
    Next
    
    Set cbsActions = Nothing
    Set BarPane = Nothing
End Sub

Private Sub msAddPopupMenu(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建托盘图标右键弹出式菜单
        
    Set mcbsPopupIcon = cbsBars.Add(cbsBars.Actions(gID.IconPopupMenu).Caption, xtpBarPopup)
    With mcbsPopupIcon.Controls
        .Add xtpControlButton, gID.IconPopupMenuMaxWindow, ""
        .Add xtpControlButton, gID.IconPopupMenuMinWindow, ""
        .Add xtpControlButton, gID.IconPopupMenuShowWindow, ""
        .Add xtpControlButton, gID.SysLoginAgain, ""
        .Add xtpControlButton, gID.SysLoginOut, ""
    End With
End Sub

Private Sub msAddToolBar(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '创建工具栏
    
    Dim cbsBar As XtremeCommandBars.CommandBar
    Dim cbsCtr As XtremeCommandBars.CommandBarControl
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    
    Set cbsActions = cbsBars.Actions
    
    '系统操作工具栏
    Set cbsBar = cbsBars.Add(cbsActions(gID.Sys).Caption, xtpBarTop)
    With cbsBar.Controls
        For mlngID = gID.SysLoginOut To gID.SysLoginAgain
            Set cbsCtr = .Add(xtpControlButton, mlngID, "")
            cbsCtr.BeginGroup = True
        Next
        For mlngID = gID.SysExportToExcel To gID.SysExportToWord
            If mlngID <> gID.SysExportToHTML Then
                Set cbsCtr = .Add(xtpControlButton, mlngID, "")
                cbsCtr.BeginGroup = True
            End If
        Next
        For mlngID = gID.SysPrintPageSet To gID.SysPrint
            Set cbsCtr = .Add(xtpControlButton, mlngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    '工具栏主题
    Set cbsBar = cbsBars.Add(cbsActions(gID.WndThemeCommandBars).Caption, xtpBarTop)
    With cbsBar.Controls
        For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            Set cbsCtr = .Add(xtpControlButton, mlngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    Set cbsBar = Nothing
    Set cbsCtr = Nothing
    Set cbsActions = Nothing
End Sub

Private Sub msCloseAllConnect(Optional ByVal blnClose As Boolean = True, _
                              Optional ByVal blnCloseListen As Boolean = True)
    '关闭所有客户端连接
    Dim sckDel As MSWinsockLib.Winsock
    
    If Not blnClose Then Exit Sub   '不执行关闭
    
    For Each sckDel In Me.Winsock1
        If sckDel.Index = 0 Then
            If blnCloseListen Then
                sckDel.Close     '关闭侦听
            End If
        Else
            If sckDel.State <> 0 Then sckDel.Close  '先关闭连接
            gArr(sckDel.Index) = gArr(0)    '清空对应数组信息
            Unload sckDel   '卸载对应控件
        End If
    Next
            
    With Me.Grid1
        .AutoRedraw = False
        .ReadOnly = False   '不取消锁定好像清除不了表格中内容。但可写入？
        .Range(1, 1, .Rows - 1, .Cols - 1).ClearText
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
    End With
    Debug.Print "msCloseAllConnect(" & blnClose & "," & blnCloseListen & ")"
    Set sckDel = Nothing
End Sub

Private Sub msGetClientInfo(ByVal strInfo As String, ByVal Index As Long)
    '将客户端发来的信息：计算机名、用户登陆系统名、用户姓名填入表格中
    '注意：客户端发来信息顺序固定为计算机名--用户登陆系统名--用户姓名
    Dim strPC As String, strLogin As String, strFull As String
    Dim LocPC As Long, LocLogin As Long, LocFull As Long, lngTemp As Long, C As Long
    
    With gVar
        LocPC = InStr(strInfo, .PTClientUserComputerName)
        LocLogin = InStr(strInfo, .PTClientUserLoginName)
        LocFull = InStr(strInfo, .PTClientUserFullName)
        
        If LocPC > 0 And LocLogin > LocPC Then
            lngTemp = LocPC + Len(.PTClientUserComputerName)
            strPC = Mid(strInfo, lngTemp, LocLogin - lngTemp)
        End If
        
        If LocLogin > LocPC And LocFull > LocLogin Then
            lngTemp = LocLogin + Len(.PTClientUserLoginName)
            strLogin = Mid(strInfo, lngTemp, LocFull - lngTemp)
        End If
        
        If LocFull > LocLogin Then
            lngTemp = LocFull + Len(.PTClientUserFullName)
            strFull = Mid(strInfo, lngTemp)
        End If
    End With
    
    With Me.Grid1
        .AutoRedraw = False
        .ReadOnly = False
        C = .Rows - 1
        For lngTemp = 1 To C
            If Val(.Cell(lngTemp, 6).Text) = Index Then
                .Cell(lngTemp, 2).Text = strPC
                .Cell(lngTemp, 3).Text = strLogin
                .Cell(lngTemp, 4).Text = strFull
                If LCase(strLogin) <> LCase(gVar.UpdateAccount) Then
                    Call msWriteLoginInfoLog(.Cell(lngTemp, 1).Text, .Cell(lngTemp, 2).Text, _
                        .Cell(lngTemp, 3).Text, .Cell(lngTemp, 4).Text, .Cell(lngTemp, 5).Text, _
                        .Cell(lngTemp, 6).Text, .Cell(lngTemp, 7).Text)
                End If
                Exit For
            End If
        Next
        .AutoRedraw = True
        .ReadOnly = True
        .Refresh
    End With
End Sub

Private Sub msGridSet(ByRef gridSet As FlexCell.Grid)
    With gridSet
        .AutoRedraw = False
        .Appearance = Flat
        .BackColorBkg = Me.BackColor
        .DisplayRowIndex = True
        .ExtendLastCol = True
        .ReadOnly = True    '禁止表格编辑
        
        .Cols = 9
        .Rows = 2
        .Cell(0, 0).Text = "序号"
        .Cell(0, 1).Text = "连接用户IP地址"
        .Cell(0, 2).Text = "连接用户计算机名称"
        .Cell(0, 3).Text = "连接用户登陆账号"
        .Cell(0, 4).Text = "连接用户姓名"
        .Cell(0, 5).Text = "连接建立时间"
        .Cell(0, 6).Text = "索引号" '"Index"
        .Cell(0, 7).Text = "申请号" '"RequestID"
        .Cell(0, 8).Text = "连接时长"
        .Column(1).Width = 120
        .Column(2).Width = 130
        .Column(3).Width = 130
        .Column(4).Width = 120
        .Column(5).Width = 120
        .RowHeight(0) = 40
        .Range(0, 0, 0, .Cols - 1).WrapText = True

        .AutoRedraw = True
        .Refresh
    End With
End Sub

Private Sub msLeftClick(ByVal CID As Long, ByRef cbsBars As XtremeCommandBars.CommandBars)
    'CommandBars单击命令响应公共过程
    
    Dim strKey As String
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    
    Set cbsActions = cbsBars.Actions
    With gID
        Select Case CID
            Case .WndThemeCommandBarsOffice2000 To .WndThemeCommandBarsWinXP
                Call gsThemeCommandBar(CID, cbsBars)
            Case .WndResetLayout
                Call msResetLayout(cbsBars)
                
            Case .SysLoginAgain
                If MsgBox("确定重新启动服务端程序吗？", vbQuestion + vbOKCancel, "重启主程序询问") = vbOK Then
                    gVar.CloseWindow = True
                    Unload Me
                    Me.Show
                End If
            Case .SysLoginOut
                If MsgBox("确定退出服务端程序吗？", vbQuestion + vbOKCancel, "关闭主程序询问") = vbOK Then
                    gVar.CloseWindow = True
                    Unload Me
                End If
                
            Case .IconPopupMenuMaxWindow
                Me.WindowState = vbMaximized
                Me.Show
            Case .IconPopupMenuMinWindow
                Me.WindowState = vbMinimized
            Case .IconPopupMenuShowWindow
                Me.WindowState = vbNormal
                Me.Show
                
            Case .HelpAbout
                Dim strAbout As String
                strAbout = "名称：" & App.Title & vbCrLf & _
                           "版本：" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
                           "版权所有：XMH"
                MsgBox strAbout, vbInformation, "关于" & App.Title
                
            Case .SysExportToCSV To .SysExportToWord, .SysPrintPageSet To .SysPrint
                If Me.ActiveControl Is Nothing Then Exit Sub
                If Not (TypeOf Me.ActiveControl Is FlexCell.Grid) Then Exit Sub
                If Not cbsActions(CID).Enabled Then Exit Sub
                Select Case CID
                    Case .SysExportToCSV To .SysExportToXML
                        Call gsGridExportTo(Me.ActiveControl, CID)
                    Case .SysExportToText
                        If MsgBox("是否将当前表格内容导出至txt文本文档？", vbQuestion + vbYesNo, "询问") = vbYes Then Call gsGridToText(Me.ActiveControl)
                    Case .SysExportToWord
                        If MsgBox("是否将当前表格内容导出至Word文档？", vbQuestion + vbYesNo, "询问") = vbYes Then Call gsGridToWord(Me.ActiveControl)
                        
                    Case .SysPrint
                        If MsgBox("确定打印当前表格内容吗？", vbQuestion + vbOKCancel, "打印询问") = vbOK Then Call gsGridPrint(Me.ActiveControl)
                    Case .SysPrintPreview
                        Call gsGridPrintPreview(Me.ActiveControl)
                    Case .SysPrintPageSet
                        Call gsGridPageSet(Me.ActiveControl)
                End Select
                
            Case Else
                strKey = LCase(cbsActions.Action(CID).Key)
                If Left(strKey, 3) = "frm" Then
                    If cbsActions.Action(CID).Enabled Then
'''                        Select Case strKey
'''                            Case LCase(cbsActions(gID.toolOptions).Key)
'''                                Call gsOpenTheWindow(strKey, vbModal, vbNormal)
'''                            Case Else
'''                                Call gsOpenTheWindow(strKey)
'''                        End Select
                        Select Case CID
                            Case .toolOptions, .WndThemeSkinSet
                                Call gsOpenTheWindow(strKey, vbModal, vbNormal)
                            Case Else
                                Call gsOpenTheWindow(strKey)
                        End Select
                    End If
                Else
                    MsgBox "【" & cbsActions(CID).Caption & "】命令未定义！", vbExclamation, "命令警告"
                End If
        End Select
    End With
    
    Set cbsActions = Nothing
End Sub

Private Sub msLoadParameter(Optional ByVal blnLoad As Boolean = True)
    '从注册表中加载参数值至公用变量中
    Dim tempVal
    
    If Not blnLoad Then Exit Sub
    
    On Error Resume Next    '加/解密函数过程可能有异常
    With gVar
        .ParaBlnWindowCloseMin = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowCloseMin, 1))    '关闭时最小化
        .ParaBlnWindowMinHide = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowMinHide, 1))  '最小化时隐藏
        .ParaBlnWindowStartMinS = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowStartMinS, 1)) '启动时最小化
        
        .TCPDefaultIP = Me.Winsock1.Item(0).LocalIP '本机IP地址
        .TCPSetIP = gVar.TCPDefaultIP   '服务端使用本机IP地址
        .ParaBlnAutoReStartServer = Val(GetSetting(.RegAppName, .RegSectionTCP, .RegKeyParaAutoReStartServer, 1))   '手动/自动重启服务模式
        .TCPSetPort = gfGetRegNumericValue(.RegAppName, .RegSectionTCP, .RegKeyTCPPort, , gVar.TCPDefaultPort, 10000, 65535) '侦听端口
        .ParaBlnAutoStartupAtBoot = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaAutoStartupAtBoot, 0))  '开机自动启动
        
        .ConSource = gfCheckIP(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerIP, .TCPSetIP))   '服务器名称/IP
        .ConDatabase = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerDatabase, EncryptString("dbTest", .EncryptKey)), .EncryptKey)    '数据库名
        .ConUserID = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerAccount, EncryptString("123", .EncryptKey)), .EncryptKey)  '登陆名
        .ConPassword = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerPassword, EncryptString("888888", .EncryptKey)), .EncryptKey)    '登陆密码
        
        .ParaBlnLimitClientConnect = Val(GetSetting(.RegAppName, .RegSectionTCP, .RegKeyParaLimitClientConnect, 0)) '限制客户端连接
        .ParaLimitClientConnectTime = gfGetRegNumericValue(.RegAppName, .RegSectionTCP, .RegKeyParaLimitClientConnectTime, True, 30, 1, 60) '限制客户端连接时长
        .TCPConnectMax = gfGetRegNumericValue(.RegAppName, .RegSectionTCP, .RegKeyParaLimitClientConnectNumber, True, 2, 1) '限制客户端连接数
        .ParaBackupPath = gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyServerBackStore, .FolderNameBackup)  '备份路径
        .ParaBackupInterval = gfGetRegNumericValue(.RegAppName, .RegSectionDBServer, .RegKeyServerBackInterval, True, 0, 0, 5) '备份频率
        .ParaBackupTime = gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyServerBackTime, "1989-12-30 19:00:00", True) '备份时间
    End With
End Sub

Private Sub msResetLayout(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '重置窗口布局：CommandBars与Dockingpane控件重置
    
    Dim cBar As XtremeCommandBars.CommandBar
    Dim L As Long, T As Long, R As Long, b As Long

    For Each cBar In cbsBars
    Debug.Print cBar.BarID, cBar.Title, cBar.Type
        cBar.Reset
        cBar.Visible = True
    Next
    
    For mlngID = 2 To cbsBars.Count
        cbsBars.GetClientRect L, T, R, b
        cbsBars.DockToolBar cbsBars(mlngID), 0, b, xtpBarTop
    Next
    
    Set cBar = Nothing
End Sub

Private Sub msRestart()
    '重启程序
    Dim strCMD As String
    
    strCMD = gVar.AppPath & "FFSR.exe " & gVar.EXENameOfServer & gVar.CmdLineSeparator & "close"
    Call gfShell(strCMD)
End Sub

Private Sub msSetServerState(ByVal colorSet As Long)
    '设置状态栏中服务端的状态
    
    Dim paneState As XtremeCommandBars.StatusBarPane
    Dim paneButton As XtremeCommandBars.StatusBarPane
    
    Set paneState = mXtrStatusBar.FindPane(gID.StatusBarPaneServerState)
    Set paneButton = mXtrStatusBar.FindPane(gID.StatusBarPaneServerButton)
    If colorSet = vbGreen Then
        paneState.BackgroundColor = vbGreen
        paneState.Text = gVar.ServerStateStarted
        paneButton.Text = gVar.ServerButtonClose
        paneButton.TextColor = vbMagenta 'vbRed
    ElseIf colorSet = vbRed Then
        paneState.BackgroundColor = vbRed
        paneState.Text = gVar.ServerStateError
        paneButton.Text = gVar.ServerButtonStart
        paneButton.TextColor = vbBlue ' vbGreen
    Else
        paneState.BackgroundColor = vbYellow
        paneState.Text = gVar.ServerStateNotStarted
        paneButton.Text = gVar.ServerButtonStart
        paneButton.TextColor = vbBlue ' vbGreen
    End If
    
    Set paneState = Nothing
    Set paneButton = Nothing
End Sub

Private Sub msStartConfirm(ByVal Index As Integer)
    '防非客户端来连接服务器，启动对应计时器进行反馈检查
    Dim tmrConfirm As VB.Timer
    Dim blnExist As Boolean
    
    If Index = 0 Then
        MsgBox "计时器Index值非法传入！", vbCritical, "防意外建立连接警报"
        Exit Sub
    End If
    
    For Each tmrConfirm In Me.Timer1    '防意外，检查是否已存在该计时器
        If tmrConfirm.Index = Index Then
            blnExist = True
            Exit For
        End If
    Next
    
    With Me.Timer1
        If Not blnExist Then Load .Item(Index) '不存在指定Index的控件时才加载
        .Item(Index).Interval = 1000    '计时器间隔
        .Item(Index).Enabled = True '激活计时
    End With
    
    Set tmrConfirm = Nothing
End Sub

Private Sub msStartServer(ByRef sckCon As MSWinsockLib.Winsock)
    '开启服务
    
    With sckCon
        If .State <> 0 Then .Close  '先关闭
        .LocalPort = gVar.TCPSetPort
        .Listen
    End With
End Sub

Private Sub msVersionCS(ByVal strVer As String, ByRef sckVer As MSWinsockLib.Winsock)
    '处理客户端发来的版本信息
    Dim strVC As String, strVS As String, strCompare As String
    Dim strNetFile As String, strSetupFile As String
    
    strNetFile = gVar.AppPath & gVar.EXENameOfClient    '保存在服务端的对比用的客户端软件exe文件
    strVS = gfBackVersion(strNetFile)   '获取服务端对比用的客户端版本号
    strVC = Mid(strVer, Len(gVar.PTVersionOfClient) + 1)    '截取客户端发来的版本号
    
    strCompare = gfVersionCompare(strVC, strVS) '比较 客户端发来的 与服务端对比用的 版本号
    If strCompare = "0" Then    '没有新版本，不用更新
        Call gfSendInfo(gVar.PTVersionNotUpdate, sckVer)
    ElseIf strCompare = "1" Then    '有新版，要更新
        Call gfSendInfo(gVar.PTVersionNeedUpdate & strVS, sckVer)   '通知客户端需要更新
        
        '组织并发送更新文件安装程序的信息给客户端
        gArr(sckVer.Index) = gArr(0)
        strSetupFile = gVar.AppPath & gVar.EXENameOfSetup
        If Not gfFileExist(strSetupFile) Then
            Call gsAlarmAndLog("要发送的更新程序不存在", False)
            Exit Sub
        End If
        With gArr(sckVer.Index)
            .FileFolder = "" ' gVar.FolderTemp
            .FileName = gVar.EXENameOfSetup
            .FilePath = strSetupFile
            .FileSizeTotal = FileLen(.FilePath)
        End With
        If sckVer.State = 7 Then
            If gfSendInfo(gfFileInfoJoin(sckVer.Index, ftSend), sckVer) Then '先发送文件的信息
                Debug.Print "Server:已发送更新程序的文件信息"
            End If
        End If
        
    Else    '版本检测异常处理
        Call gfSendInfo(gVar.PTVersionNotUpdate & strCompare, sckVer)
        Call gsAlarmAndLogEx("客户端版本：" & strVC & ",服务端版本：" & strVS, "Server端版本检测异常", False)
    End If
End Sub

Private Sub msWriteLoginInfoLog(ByVal strIP As String, ByVal strPC As String, ByVal strAccount As String, _
    ByVal strUserName As String, ByVal strTime As String, ByVal strIndex As String, ByVal strApplyID As String)
    '记录用户的登陆日志，文件名保存在gVar.FileNameLoginLog中
    
    Const conSize As Long = 1000000 '固定日志文件大小，超过则按日期存储
    Dim strNewFile As String, strSep As String
    Dim intNum As Integer
    
    If Not gfFileRepair(gVar.FolderData, True) Then Exit Sub    '判断日志目录是否存在
    If Not gfFileRepair(gVar.FileNameLoginLog) Then Exit Sub    '判断日志文件是否存在
    
    If FileLen(gVar.FileNameLoginLog) > conSize Then    '日志文件太大时存档
        strNewFile = Left(gVar.FileNameLoginLog, InStrRev(gVar.FileNameLoginLog, ".") - 1) & _
            Format(Now, gVar.Formatymdhms) & Mid(gVar.FileNameLoginLog, InStrRev(gVar.FileNameLoginLog, "."))
        Debug.Print strNewFile  '生成按日期保存的文件名
        Close
        If Not gfFileRename(gVar.FileNameLoginLog, strNewFile) Then Exit Sub    '改名存储
    End If
    
    intNum = FreeFile
    strSep = vbTab & vbTab
    On Error Resume Next
    
    Open gVar.FileNameLoginLog For Append As intNum
    Print #intNum, strIP & strSep & strPC & strSep & strAccount & _
        strSep & strUserName & strSep & strTime & strSep & strIndex & strSep & strApplyID
    Close
    
End Sub

Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '命令单击事件
    Call msLeftClick(Control.ID, Me.CommandBars1)
End Sub

Private Sub CommandBars1_Resize()
    '调整窗口布局
    
    Dim L As Long, T As Long, R As Long, b As Long
    
    On Error Resume Next
    Me.CommandBars1.GetClientRect L, T, R, b
    Grid1.Move L, T, R - L, b - T
    
End Sub

Private Sub CommandBars1_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    'CommandBars控件的Action状态的切换
    
    Dim blnFC As Boolean
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs控件Actions集合的引用
    
    Set cbsActions = Me.CommandBars1.Actions
    If Screen.ActiveControl Is Nothing Then
        blnFC = False
    Else
        blnFC = TypeOf Screen.ActiveControl Is FlexCell.Grid
    End If
    With gID
        For mlngID = .SysExportToCSV To .SysExportToWord
            cbsActions(mlngID).Enabled = blnFC
        Next
        For mlngID = .SysPrintPageSet To .SysPrint
            cbsActions(mlngID).Enabled = blnFC
        Next
    End With
End Sub

Private Sub Form_Load()
    '窗体加载
    
    Dim cbsBars As XtremeCommandBars.CommandBars
    
    '打开多个应用程序检查
    If App.PrevInstance Then
        MsgBox "不可同时打开多个应用程序！", vbCritical, "警报"
        End
    End If
    
    Timer1.Item(0).Interval = 1000  '计时器循环时间
    Call Main   '初始化全局公用变量
    Set gWind = Me  '指定主窗体给全局引用对象
    XtremeCommandBars.CommandBarsGlobalSettings.App = App '一个默认设置
    Set cbsBars = Me.CommandBars1
    
    Call msLoadParameter(True)  '加载配置参数
    Call msAddAction(cbsBars)   '创建Actions集合
    Call msAddMenu(cbsBars)     '创建菜单栏
    Call msAddToolBar(cbsBars)  '创建工具栏
    Call msAddPopupMenu(cbsBars)    '创建托盘图标的菜单
    Call msAddXtrStatusBar(cbsBars) '创建状态栏
    Call msAddKeyBindings(cbsBars)  '添加快捷键,放到LoadCommandBars方法后面才能生效？？？
    Call msAddDesignerControls(cbsBars) 'CommandBars自定义对话框中使用的
    
    cbsBars.AddImageList ImageList1         '使CommandBars控件匹配ImageList控件中图标
    cbsBars.EnableCustomization True        '允许CommandBars自定义，此属性最好放在所有CommandBars设定之后
    cbsBars.Options.UpdatePeriod = 250      '更改CommandBars的Update事件的执行周期，默认100ms
    
    Call gsLoadSkin(Me, Me.SkinFramework1, sMSO7, True, , , True) '加载窗口主题
    
    '加载工具栏主题
    Call gsThemeCommandBar(Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyServerCommandbarsTheme, gID.WndThemeCommandBarsRibbon)), cbsBars)
    
    '注册表信息加载-CommandBars设置
    Call cbsBars.LoadCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegKeyCBSServerSetting)

    Call gsFormSizeLoad(Me) '注册表信息加载-窗口位置大小
    
    
    
    
    '检查是否为试用版******************************
    
    
    Call msGridSet(Grid1)  '表格设置
    Call gfNotifyIconAdd(Me)    '添加托盘图标
    If gVar.ParaBlnWindowStartMinS Then
        Me.WindowState = vbMinimized '启动时最小化
    End If
    
    Set cbsBars = Nothing   '销毁使用完的对象
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '响应托盘图标的菜单
    Dim sngMsg As Single
    
    If Y <> 0 Then Exit Sub    '似乎此句可限制住鼠标一定是在托盘图标上，不是在窗体上
    sngMsg = X / Screen.TwipsPerPixelX
    Select Case sngMsg
        Case WM_RBUTTONUP
            mcbsPopupIcon.ShowPopup  '右键弹出Popup菜单

        Case WM_LBUTTONDBLCLK   '左键双击托盘图标时 窗口最显示/最小化 切换
            With Me
                If .WindowState = vbMinimized Then
                    .WindowState = vbNormal
                    .Show
                    .SetFocus
                Else
                    .WindowState = vbMinimized
                End If
            End With
        Case Else
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '判断是否真正要关闭窗口
    
    If gVar.ParaBlnWindowCloseMin Then
        If Not gVar.CloseWindow Then
            Cancel = True
            Me.WindowState = vbMinimized
        End If
        gVar.CloseWindow = False
    Else
        If Not gVar.CloseWindow Then
            If MsgBox("是否最小化窗口？", vbQuestion + vbYesNo, "关闭或最小化") = vbYes Then
                Cancel = True
                Me.WindowState = vbMinimized
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    '窗口最小化提示
    If Me.Visible And Me.WindowState = vbMinimized Then
        If gVar.ParaBlnWindowMinHide Then
            Me.Hide
            Call gfNotifyIconBalloon(Me, "最小化到系统托盘图标啦", "最小化提示")
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '卸载窗体时保存信息
    Dim resetNotifyIconData As gtypeNOTIFYICONDATA
    
    On Error Resume Next
    
    '保存注册表信息-CommandBars设置
    Call Me.CommandBars1.SaveCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegKeyCBSServerSetting)
    
    Call gsFormSizeSave(Me) '保存注册表信息-窗口位置大小
    Call gsSaveCommandbarsTheme(Me.CommandBars1)   '保存CommandBars的风格主题
    
    gVar.CloseWindow = False    '清除关闭窗口状态
    Call SkinFramework1.LoadSkin("", "")    '清空皮肤
    Set mXtrStatusBar = Nothing  '清除状态栏
    Set mcbsPopupIcon = Nothing '清除Popup菜单
    Call gfNotifyIconDelete(Me) '删除托盘图标
    gNotifyIconData = resetNotifyIconData   '清空托盘气泡信息。否则重启程序时会自动弹出？而且只能放上句删除托盘图标语句的后面?
    ReDim gArr(0)
    Set gWind = Nothing '清除全局窗体引用
    
End Sub

Private Sub mXtrStatusBar_PaneClick(ByVal Pane As XtremeCommandBars.StatusBarPane)
    '状态栏按钮事件
    Dim strMsg As String
    
    If Pane.ID = gID.StatusBarPaneServerButton Then '断开/开启服务
        If Pane.Text = gVar.ServerButtonClose Then strMsg = "关闭后会断开所有用户的连接。"
        If MsgBox("是否" & Pane.Text & "？" & strMsg, vbQuestion + vbYesNo, "重启/断开服务询问") = vbNo Then Exit Sub
        If Pane.Text = gVar.ServerButtonClose Then     '关闭服务
            Pane.Text = gVar.ServerButtonStart
            Call msCloseAllConnect(True, True)
        ElseIf Pane.Text = gVar.ServerButtonStart Then     '开启服务
            Pane.Text = gVar.ServerButtonClose
            Call msStartServer(Me.Winsock1.Item(0))
        End If
        
    ElseIf Pane.ID = gID.StatusBarPaneReStartButton Then    '手动/自动重启服务模式
        strMsg = "是否切换成" & IIf(gVar.ParaBlnAutoReStartServer, "手", "自") & "动重启服务模式？"
        If MsgBox(strMsg, vbQuestion + vbYesNo, "模式切换询问") = vbYes Then
            gVar.ParaBlnAutoReStartServer = Not gVar.ParaBlnAutoReStartServer
            mXtrStatusBar.FindPane(gID.StatusBarPaneReStartButton).Text = IIf(gVar.ParaBlnAutoReStartServer, "自", "手") & "动重启服务模式"
            Call SaveSetting(gVar.RegAppName, gVar.RegSectionTCP, gVar.RegKeyParaAutoReStartServer, IIf(gVar.ParaBlnAutoReStartServer, 1, 0))
        End If
        
    End If
End Sub

Private Sub Timer1_Timer(Index As Integer)
    'Index=0的计时器间隔1秒。Timer1的Index值 与 Winsock1的Index对应
    
    Dim sckClose As MSWinsockLib.Winsock, sckCheck As MSWinsockLib.Winsock, tmrUld As VB.Timer
    Dim timeOut As Long, lngRows As Long
'    Static CheckConnectTime As Long
'    Static ConfirmTime() As Long
'    Static ConfirmOK() As Boolean
'    Static CountTime() As Long
        
    On Error Resume Next
    If Index = 0 Then
        If Me.Winsock1.Item(Index).State = 2 Then  '侦听正常状态
            Call msSetServerState(vbGreen)
        Else    '其它状态
            If Me.Winsock1.Item(Index).State = 9 Then  '异常状态
                Call msSetServerState(vbRed)
            Else    '关闭等
                Call msSetServerState(vbYellow)
            End If
            If gVar.ParaBlnAutoReStartServer Then   '若勾选了自动开启服务则重启服务
                Call msStartServer(Me.Winsock1.Item(0))
            End If
        End If
        
        '当客户端非正常关闭时，连接不会自动断开，此处每隔一段时间检查一次所有连接的状态，不等于7则关闭掉连接
        CheckConnectTime = CheckConnectTime + 1
        If CheckConnectTime > 5 Then    '每隔N秒检查一次
            For Each sckCheck In Me.Winsock1
                If sckCheck.Index <> 0 Then
                    If sckCheck.State <> 7 Then
                        For Each tmrUld In Me.Timer1
                            If tmrUld.Index = sckCheck.Index Then
                                Unload tmrUld 'Me.Timer1.Item(sckCheck.Index)
                                Exit For
                            End If
                        Next
                        Debug.Print "CheckConnect:" & sckCheck.Index
                        Call Winsock1_Close(sckCheck.Index)
                    End If
                End If
            Next
            CheckConnectTime = 0
        End If
        '''index=0计时器为服务端自身用
        
        '刷新每个连接的时长
        With Me.Grid1
            lngRows = .Rows - 1
            If lngRows > 0 Then
                For mlngID = 1 To lngRows
                    If Len(.Cell(mlngID, 1).Text) = 0 Or Not IsDate(.Cell(mlngID, 5).Text) Then
                        Exit For
                    Else
                        .Cell(mlngID, 8).Text = Format(Now - CDate(.Cell(mlngID, 5).Text), "HH:mm:ss")
                    End If
                Next
            End If
        End With
    Else
        '''index>0为各个客户端连接用
        
        timeOut = gVar.ParaLimitClientConnectTime * 60  '计算允许连接时长的总秒数
'''        ReDim Preserve ConfirmTime(Me.Timer1.UBound)    '需要每次都这样？
'''        ReDim Preserve ConfirmOK(Me.Timer1.UBound)
'''        ReDim Preserve CountTime(Me.Timer1.UBound)
        
        If Not ConfirmOK(Index) Then ConfirmTime(Index) = ConfirmTime(Index) + 1
        If gVar.ParaBlnLimitClientConnect Then CountTime(Index) = CountTime(Index) + 1
        
        If ConfirmTime(Index) > gVar.TCPWaitTime Then   '确认是客户端发来的连接
            If Not gArr(Index).Connected Then   '不是客户端则关闭
                For Each sckClose In Me.Winsock1
                    If sckClose.Index = Index Then
                        Call Winsock1_Close(Index) '关闭连接
                        Exit For
                    End If
                Next
            End If
            ConfirmTime(Index) = 0  '确认计时器清零
            ConfirmOK(Index) = True '确认标志
            If Not gVar.ParaBlnLimitClientConnect Then  '若没有选择限制客户端连接功能
                ConfirmOK(Index) = False
                CountTime(Index) = 0    '限制连接计时器清零
                For Each tmrUld In Me.Timer1
                    If tmrUld.Index = Index Then
                        Unload tmrUld   'Unload Me.Timer1.Item(Index) '确认完后卸载掉对应计时器控件
                        Exit For
                    End If
                Next
            End If
        End If
            
        If CountTime(Index) > timeOut Then   '计时已满则关闭对应客户端连接
            '若在文件传输状态下则等待传输2分钟后直接关闭
            If (Not gArr(Index).FileTransmitState) Or (CountTime(Index) - timeOut > 120) Then
                CountTime(Index) = 0    '清空计时
                If gArr(Index).FileTransmitState Then   '清空文件传输信息
                    Close
                    gArr(Index) = gArr(0)
                End If
                Unload Me.Timer1.Item(Index)
                For Each sckCheck In Me.Winsock1
                    If sckCheck.Index = Index Then '确认控件是否存在
                        If sckCheck.State = 7 Then '确认控件状态是否对
                            Call gfSendInfo(gVar.PTConnectTimeOut, Me.Winsock1.Item(Index)) '发送连接时间已到信息给客户端
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        
    End If
    
    Set sckClose = Nothing
    Set sckCheck = Nothing
    Set tmrUld = Nothing
    
    If Err.Number > 0 Then
        Call msRestart
    End If
End Sub

Private Sub Winsock1_Close(Index As Integer)
    '关闭连接
    Dim K As Long, C As Long
    Dim strIP As String, strRequestID As String
    Dim tmDel As VB.Timer
    
    On Error Resume Next
    If Index = 0 Then
        Call msCloseAllConnect(True, False)  '关闭侦听控件则关闭所有连接
    Else
        With Me.Grid1
            .AutoRedraw = False
            .ReadOnly = False   '先取消锁定
            C = .Rows - 1
            For K = 1 To C
                strIP = Trim(.Cell(K, 1).Text)
                strRequestID = Trim(.Cell(K, 7).Text)
                If Len(strIP) > 0 Then
                    If (strIP = Me.Winsock1.Item(Index).RemoteHostIP) And _
                            (strRequestID = Me.Winsock1.Item(Index).Tag) Then   '有可能同一IP登陆多个客户端，故以RequestID区分
                        .RemoveItem K   '移除对应行信息
                        .AddItem ""     '末尾添加一行维持表格行数不变
                        Debug.Print K & ",Winsock_Close:" & Index, Me.Winsock1.Item(Index).Tag, Me.Winsock1.Item(Index).RemoteHostIP
                        For Each tmDel In Me.Timer1 '卸载对应计时器
                            If tmDel.Index <> 0 Then
                                If tmDel.Index = Index Then '确认控件是否存在
                                    Unload tmDel
                                    If Index <= UBound(CountTime) Then CountTime(Index) = 0
                                    Exit For
                                End If
                            End If
                        Next
                        Unload Me.Winsock1.Item(Index)  '卸载断开的客户端的连接控件
                        gArr(Index) = gArr(0)   '清除数组
                        Close   '关闭所有打开的文件
                        Exit For
                    End If
                End If
            Next
            .AutoRedraw = True
            .ReadOnly = True
            .Refresh
        End With
    End If
    
    Set tmDel = Nothing
    
    If Err.Number > 0 Then
        Call msRestart
    End If
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    '建立连接
    
    Dim sckNew As MSWinsockLib.Winsock
    Dim K As Long
    Dim blnFull As Boolean
    
    On Error Resume Next
    
    If Index <> 0 Then Exit Sub '仅0元素的控件在侦听能接收连接申请
    
    '查找Winsock1中可用的最小的Index值
    For Each sckNew In Winsock1
        If sckNew.Index = K Then
            K = K + 1
        Else    '说明Index断号不连续了
            Exit For
        End If
    Next
    Set sckNew = Nothing
    
    If Me.Winsock1.Count > gVar.TCPConnectMax Then
        Call gsAlarmAndLog("客户端发送的申请连接数已超过最大值" & CStr(gVar.TCPConnectMax), False)
        blnFull = True
    End If
    
    With Me.Winsock1
        If K = .Count Then ReDim Preserve gArr(K)   '重置数组大小
        gArr(K) = gArr(0)   '初始化数组元素K。可能被之前建立过的连接使用过
        
        Load .Item(K)   '生成Winsock1(K)控件
        .Item(K).Accept requestID   '指定客户端申请的连接给新生成的Winsock1(K)控件
        .Item(K).Tag = requestID    '存储该申请号，另作它用
        Debug.Print "Load winsock1(" & K & ")"
        Call msAddConnectToGrid(.Item(K)) '将连接添加进表格中
        
        ReDim Preserve ConfirmTime(.UBound) '数组元素值的变化在Timer1控件中，初始化权且放此
        ReDim Preserve ConfirmOK(.UBound)
        ReDim Preserve CountTime(.UBound)
    End With
    
    If blnFull Then '若连接数已满，则发信通知客户端。该客户端自动关闭连接，客户端同时触发关闭事件.
        Call gfSendInfo(gVar.PTConnectIsFull, Me.Winsock1.Item(K))
        Exit Sub
    End If
    
    Call gfSendInfo(gVar.PTClientConfirm, Me.Winsock1.Item(K)) '发送客户端确认信息，若规定时间内返回确认信息则连接正常，否则断开连接。
    Call msStartConfirm(K)  '激活 返回确认信息 计时器

    If Err.Number > 0 Then
        Call msRestart
    End If
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    '接收信息
    Dim strGet As String
    Dim byteGet() As Byte
    
    On Error Resume Next
    
    With gArr(Index)
        If Not .FileTransmitState Then
            '字符信息传输状态
            
            Me.Winsock1.Item(Index).GetData strGet  '先控件接收信息
            
            Call gfRestoreInfo(strGet, Me.Winsock1.Item(Index))  '获取发来的是不是文件信息
            
            If InStr(strGet, gVar.PTClientIsTrue) Then '客户端发回的连接确认信息
                .Connected = True   '设置状态以供计时器判断
                Call gfSendInfo(gfDatabaseInfoJoin(True), Me.Winsock1.Item(Index))
                
            ElseIf InStr(strGet, gVar.PTVersionOfClient) > 0 Then '接收到客户端版本信息
                Call msVersionCS(strGet, Me.Winsock1.Item(Index))
            
            ElseIf InStr(strGet, gVar.PTClientUserComputerName) > 0 Then '客户端发来的计算机名、用户名等信息
                Call msGetClientInfo(strGet, Index)
                
            ElseIf InStr(strGet, gVar.PTFileStart) > 0 Then '要开始发送文件给客户端
                Call gfSendFile(.FilePath, Me.Winsock1.Item(Index))
                
            End If
            Debug.Print "Server GetInfo:" & strGet, bytesTotal
            
        '字符信息传输状态
        Else
        '文件传输状态
            
            If .FileNumber = 0 Then
                .FileNumber = FreeFile '生成操作文件号
                Open .FilePath For Binary As #.FileNumber '以二进制方式打开文件
            End If
            
            ReDim byteGet(bytesTotal - 1)   '确定字节数组大小
            Me.Winsock1.Item(Index).GetData byteGet, vbArray + vbByte    '接收文件
            Put #.FileNumber, , byteGet '写入文件
            .FileSizeCompleted = .FileSizeCompleted + bytesTotal    '统计已传输文件大小的总量
            
            If .FileSizeCompleted >= .FileSizeTotal Then    '若接收完成
                Close #.FileNumber  '关闭操作文件号
                Call gfSendInfo(gVar.PTFileEnd, Me.Winsock1.Item(Index)) '发送接收完的信号
                gArr(Index) = gArr(0)   '清空文件信息
                Debug.Print "Server Received Over"
            End If
            
        End If
    End With
    
    If Err.Number > 0 Then
        Call msRestart
    End If
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '异常处理
    
    If Index = 0 Then
        Call msCloseAllConnect(True, True)  '侦听控件异常则关闭所有连接。可能没什么用
    Else
        If gArr(Index).FileTransmitState Then   '异常时清空文件传输信息
            Close
            gArr(Index) = gArr(0)
        End If
    End If
    Debug.Print "ServerWinsockError:" & Index & "--" & Err.Number & "  " & Err.Description
End Sub

Private Sub Winsock1_SendComplete(Index As Integer)
    '发送完处理
    
    If Index = 0 Then Exit Sub  '0元素只负责侦听不传输
    With gArr(Index)
        If .FileTransmitState Then
            If .FileSizeCompleted < .FileSizeTotal Then '未发送完则继续发送
                Call gfSendFile(.FilePath, Me.Winsock1.Item(Index))
            Else    '发送则完清空传输信息
                gArr(Index) = gArr(0)
                Debug.Print "Server Send File Over"
            End If
        End If
    End With
End Sub
