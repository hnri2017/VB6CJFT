VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.CommandBars.v15.3.1.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.MDIForm frmSysMain 
   BackColor       =   &H8000000C&
   Caption         =   "FFC"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10155
   Icon            =   "frmSysMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Index           =   1
      Left            =   2400
      Top             =   2880
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   1680
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   2880
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
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   3480
      Top             =   2880
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   3000
      Top             =   2880
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

Dim mlngID As Long  'ѭ������ID
Dim WithEvents mXtrStatusBar As XtremeCommandBars.StatusBar  '״̬���ؼ�
Attribute mXtrStatusBar.VB_VarHelpID = -1
Dim mcbsPopupIcon As XtremeCommandBars.CommandBar    '����ͼ��Pupup�˵�



Private Sub msAddAction(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '����CommandBars��Action
    
    Dim cbsAction As XtremeCommandBars.CommandBarAction
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
    Set cbsActions = cbsBars.Actions
    cbsBars.EnableActions   '����CommandBars��Actions����
    
'    cbsActions.Add "Id", "Caption", "TooltipText", "DescriptionText", "Category"   '����
    With cbsActions
        .Add gID.Sys, "ϵͳ", "", "", "ϵͳ"
        
        .Add gID.SysLoginOut, "�˳�", "", "", ""
        .Add gID.SysLoginAgain, "����", "", "", ""
        
        .Add gID.SysExportToCSV, "������CSV", "", "", ""
        .Add gID.SysExportToExcel, "������Excel", "", "", ""
        .Add gID.SysExportToHTML, "������HTML", "", "", ""
        .Add gID.SysExportToPDF, "������PDF", "", "", ""
        .Add gID.SysExportToText, "������txt", "", "", ""
        .Add gID.SysExportToWord, "������Word", "", "", ""
        .Add gID.SysExportToXML, "������XML", "", "", ""
        
        .Add gID.SysPrint, "��ӡ", "", "", ""
        .Add gID.SysPrintPageSet, "��ӡҳ������", "", "", ""
        .Add gID.SysPrintPreview, "��ӡԤ��", "", "", ""
        
        .Add gID.Wnd, "����", "", "", "����"
        
        .Add gID.WndResetLayout, "���ô��ڲ���", "", "", ""
        
        .Add gID.WndThemeCommandBars, "����������", "", "", ""
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
        
        .Add gID.Help, "����", "", "", "����"
        .Add gID.HelpAbout, "���ڡ�", "", "", ""
        
        .Add gID.StatusBarPane, "״̬��", "", "", ""
        .Add gID.StatusBarPaneProgress, "������", "", "", ""
        .Add gID.StatusBarPaneUserInfo, "�û���Ϣ", "", "", ""
        .Add gID.StatusBarPaneTime, "����ʱ��", "", "", ""
        .Add gID.StatusBarPaneProgressText, "�������ٷֱ�ֵ", "", "", ""
        .Add gID.StatusBarPaneServerButton, "������/�Ͽ���ť", "", "", ""
        .Add gID.StatusBarPaneServerState, "����״̬", "", "", ""
        .Add gID.StatusBarPaneTime, "ϵͳʱ��", "", "", ""
        .Add gID.StatusBarPaneIP, "����IP��ַ", "", "", ""
        .Add gID.StatusBarPanePort, "���ӷ������˿�", "", "", ""
        .Add gID.StatusBarPaneConnectState, "���ӷ�����״̬", "", "", ""
        .Add gID.StatusBarPaneConnectButton, "��������������Ӱ�ť", "", "", ""
        .Add gID.StatusBarPaneReStartButton, "�����Զ�/�ֶ�����ģʽ�л���ť", "", "", ""
        
        .Add gID.IconPopupMenu, "����ͼ��˵�", "", "", ""
        .Add gID.IconPopupMenuMaxWindow, "��󻯴���", "", "", ""
        .Add gID.IconPopupMenuMinWindow, "��С������", "", "", ""
        .Add gID.IconPopupMenuShowWindow, "��ʾ����", "", "", ""
        
        .Add gID.Tool, "����", "", "", "����"
        .Add gID.toolOptions, "ѡ��", "", "", "frmOption"
        
        
'        .Add gID, "", "", "", ""
        
    End With
    
    '���cbsActions����������ToolTipText��DescriptionText��Key��Category
    For Each cbsAction In cbsActions
        With cbsAction
            If .ID < 20000 Then
                .ToolTipText = .Caption
                .DescriptionText = .ToolTipText
                .Key = .Category    'Ϊ�˵�ʱ�������ã�����Actionʱ������������Category��
                .Category = cbsActions((.ID \ 1000) * 1000).Category
            End If
        End With
    Next
    
    '���ϵ�е�cbsActions���������Ե���������
    With cbsActions
        For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            .Action(mlngID).DescriptionText = .Action(gID.WndThemeCommandBars).Caption & "����Ϊ��" & .Action(mlngID).DescriptionText
            .Action(mlngID).ToolTipText = .Action(mlngID).DescriptionText
        Next
    End With
    
    Set cbsAction = Nothing
    Set cbsActions = Nothing
End Sub

Private Sub msAddDesignerControls(ByRef cbsBars As XtremeCommandBars.CommandBars)
    'CommandBars�Զ���Ի���������������
    
    Dim cbsControls As XtremeCommandBars.CommandBarControls
    Dim cbsAction As XtremeCommandBars.CommandBarAction
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
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
    '������ݼ�
    
    With cbsBars.KeyBindings
        .AddShortcut gID.SysLoginOut, "F10"
    End With
    
End Sub

Private Sub msAddMenu(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '�����˵���
    
    Dim cbsMenuBar As XtremeCommandBars.MenuBar
    Dim cbsMenuMain As XtremeCommandBars.CommandBarPopup
    Dim cbsMenuCtrl As XtremeCommandBars.CommandBarControl
    
    Set cbsMenuBar = cbsBars.ActiveMenuBar
    cbsMenuBar.ShowGripper = False  '����ʾ���϶����Ǹ������
    cbsMenuBar.EnableDocking xtpFlagStretched     '�˵�����ռһ���Ҳ��������϶�
    
    'ϵͳ���˵�
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
    
    '�������˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Wnd, "")
    With cbsMenuMain.CommandBar.Controls
        '���ò���
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.WndResetLayout, "")
        cbsMenuCtrl.BeginGroup = True
        
        '����ID35001�Զ��幤����
        Set cbsMenuCtrl = .Add(xtpControlButton, XTP_ID_CUSTOMIZE, "�Զ��幤����...")
        cbsMenuCtrl.BeginGroup = True
    
        '����ID59392�������б�
        Set cbsMenuCtrl = .Add(xtpControlPopup, 0, "�������б�")
        cbsMenuCtrl.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, ""
        
        'CommandBars�����������Ӳ˵�
        Set cbsMenuCtrl = .Add(xtpControlPopup, gID.WndThemeCommandBars, "")
        With cbsMenuCtrl.CommandBar.Controls
            For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
                .Add xtpControlButton, mlngID, ""
            Next
        End With
    End With
    
    '���߲˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Tool, "")
    cbsMenuMain.CommandBar.Controls.Add xtpControlButton, gID.toolOptions, ""
    
    '�������˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Help, "")
    cbsMenuMain.CommandBar.Controls.Add xtpControlButton, gID.HelpAbout, ""
    
    Set cbsMenuBar = Nothing
    Set cbsMenuMain = Nothing
    Set cbsMenuCtrl = Nothing
End Sub

Private Sub msAddPopupMenu(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '��������ͼ���Ҽ�����ʽ�˵�
        
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
    '����������
    
    Dim cbsBar As XtremeCommandBars.CommandBar
    Dim cbsCtr As XtremeCommandBars.CommandBarControl
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
    Set cbsActions = cbsBars.Actions
    
    'ϵͳ����������
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
    
    '����������
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

Private Sub msAddXtrStatusBar(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '����״̬��
    
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    Dim BarPane As XtremeCommandBars.StatusBarPane
    
    Set cbsActions = cbsBars.Actions
    Set mXtrStatusBar = cbsBars.StatusBar
    With mXtrStatusBar
        .AddPane 0      'ϵͳPane����ʾCommandBarActions��Description
        .SetPaneStyle 0, SBPS_STRETCH
        .SetPaneText 0, "Hello"
        .IdleText = "Hello"
        
        .AddPane gID.StatusBarPaneUserInfo
        .SetPaneText gID.StatusBarPaneUserInfo, "�л�����"
        .FindPane(gID.StatusBarPaneUserInfo).Width = 60
        
        .AddPane gID.StatusBarPaneIP
        .SetPaneText gID.StatusBarPaneIP, Me.Winsock1.Item(1).LocalIP  'gVar.TCPSetIP
        .FindPane(gID.StatusBarPaneIP).Width = 90
        
        .AddPane gID.StatusBarPanePort
        .SetPaneText gID.StatusBarPanePort, gVar.TCPSetPort
        .FindPane(gID.StatusBarPanePort).Width = 60
        
        .AddPane gID.StatusBarPaneConnectState
        .SetPaneText gID.StatusBarPaneConnectState, gVar.ClientStateDisConnected
        .FindPane(gID.StatusBarPaneConnectState).Width = 60
        
'''        .AddPane gID.StatusBarPaneReStartButton
'''        .SetPaneText gID.StatusBarPaneReStartButton, IIf(gVar.ParaBlnAutoReStartServer, "��", "��") & "����������ģʽ"
'''        .FindPane(gID.StatusBarPaneReStartButton).Width = 120
'''        .FindPane(gID.StatusBarPaneReStartButton).BackgroundColor = vbCyan
'''        .FindPane(gID.StatusBarPaneReStartButton).Button = True
        
'''        .AddPane gID.StatusBarPaneServerState
'''        .FindPane(gID.StatusBarPaneServerState).Text = gVar.ServerStateNotStarted
'''        .FindPane(gID.StatusBarPaneServerState).Width = 60
        
'''        .AddPane gID.StatusBarPaneServerButton
'''        .FindPane(gID.StatusBarPaneServerButton).Text = gVar.ServerButtonStart
'''        .FindPane(gID.StatusBarPaneServerButton).Width = 60
'''        .FindPane(gID.StatusBarPaneServerButton).Button = True
        
        .AddProgressPane gID.StatusBarPaneProgress
                
        .AddPane gID.StatusBarPaneProgressText
        .SetPaneText gID.StatusBarPaneProgressText, "0%"
        .FindPane(gID.StatusBarPaneProgressText).Width = 60
        
        .AddPane 59137  'CapsLock����״̬
        .AddPane 59138  'NumLK����״̬
        .AddPane 59139  'ScrLK����״̬
        .FindPane(0).Caption = "Idle Text"
        .FindPane(59137).Caption = "Caps Lock��״̬"
        .FindPane(59138).Caption = "Num LocK��״̬"
        .FindPane(59139).Caption = "Scroll LocK��״̬"
        
        .Visible = True
        .EnableCustomization True
    End With
    
    For Each BarPane In mXtrStatusBar     '����Caption��ToolTip��Alignment����
        If Len(BarPane.Caption) = 0 Then BarPane.Caption = cbsActions(BarPane.ID).Caption
        BarPane.ToolTip = BarPane.Caption
        If BarPane.ID <> 0 Then BarPane.Alignment = xtpAlignmentCenter
    Next
    
    Set cbsActions = Nothing
    Set BarPane = Nothing
End Sub

Private Sub msConnectToServer(ByRef sckCon As MSWinsockLib.Winsock, Optional ByVal blnConnect As Boolean = False)
    '�����������������
    
    If Not blnConnect Then Exit Sub
    With sckCon
        If .State <> 0 Then .Close
        .RemoteHost = gVar.TCPSetIP
        .RemotePort = gVar.TCPSetPort
        .Connect
    End With
End Sub

Private Sub msLeftClick(ByVal CID As Long, ByRef cbsBars As XtremeCommandBars.CommandBars)
    'CommandBars����������Ӧ��������
    
    Dim strKey As String
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
    Set cbsActions = cbsBars.Actions
    With gID
        Select Case CID
            Case .WndThemeCommandBarsOffice2000 To .WndThemeCommandBarsWinXP
                Call gsThemeCommandBar(CID, cbsBars)
            Case .WndResetLayout
                Call msResetLayout(cbsBars)
                
            Case .SysLoginAgain
                If MsgBox("ȷ�����������ͻ��˳�����", vbQuestion + vbOKCancel, "����������ѯ��") = vbOK Then
                    Call msUnloadMe(True)
                    Load Me
                End If
            Case .SysLoginOut
                If MsgBox("ȷ���˳��ͻ��˳�����", vbQuestion + vbOKCancel, "�ر�������ѯ��") = vbOK Then
                    Call msUnloadMe(True)
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
                strAbout = "���ƣ�" & App.Title & vbCrLf & _
                           "�汾��" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
                           "��Ȩ���У�XMH"
                MsgBox strAbout, vbInformation, "����" & App.Title
            
            Case .SysExportToCSV To .SysExportToWord, .SysPrintPageSet To .SysPrint
                If Screen.ActiveControl Is Nothing Then Exit Sub
                If Not (TypeOf Screen.ActiveControl Is FlexCell.Grid) Then Exit Sub
                Select Case CID
                    Case .SysExportToCSV To .SysExportToXML
                        Call gsGridExportTo(Screen.ActiveControl, CID)
                    Case .SysExportToText
                        If MsgBox("�Ƿ񽫵�ǰ������ݵ�����txt�ı��ĵ���", vbQuestion + vbYesNo, "ѯ��") = vbYes Then Call gsGridToText(Screen.ActiveControl)
                    Case .SysExportToWord
                        If MsgBox("�Ƿ񽫵�ǰ������ݵ�����Word�ĵ���", vbQuestion + vbYesNo, "ѯ��") = vbYes Then Call gsGridToWord(Screen.ActiveControl)
                        
                    Case .SysPrint
                        If MsgBox("ȷ����ӡ��ǰ���������", vbQuestion + vbOKCancel, "��ӡѯ��") = vbOK Then Call gsGridPrint
                    Case .SysPrintPreview
                        Call gsGridPrintPreview
                    Case .SysPrintPageSet
                        Call gsGridPageSet
                End Select
                
            Case Else
                strKey = LCase(cbsActions.Action(CID).Key)
                If Left(strKey, 3) = "frm" Then
                    If cbsActions.Action(CID).Enabled Then
                        Select Case CID
                            Case .toolOptions
                                Call gsOpenTheWindow(strKey, vbModal, vbNormal)
                            Case Else
                                Call gsOpenTheWindow(strKey)
                        End Select
                    End If
                Else
                    MsgBox "��" & cbsActions(CID).Caption & "������δ���壡", vbExclamation, "�����"
                End If
        End Select
    End With
    
    Set cbsActions = Nothing
End Sub

Private Sub msLoadParameter(Optional ByVal blnLoad As Boolean = True)
    '��ע����м��ز���ֵ�����ñ�����
    Dim tempVal
    
    If Not blnLoad Then Exit Sub
    
    On Error Resume Next    '��/���ܺ������̿������쳣
    With gVar
        .ParaBlnWindowCloseMin = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowCloseMin, 1))    '�ر�ʱ��С��
        .ParaBlnWindowMinHide = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowMinHide, 1))  '��С��ʱ����
        
        .TCPDefaultIP = Me.Winsock1.Item(0).LocalIP '����IP��ַ
        .TCPSetIP = gfCheckIP(GetSetting(.RegAppName, .RegSectionTCP, .RegKeyTCPIP, .TCPDefaultIP)) 'Ҫ���ӷ����IP��ַ
        .TCPSetPort = gfGetRegNumericValue(.RegAppName, .RegSectionTCP, .RegKeyTCPPort, , .TCPDefaultPort, 10000, 65535) 'Ҫ���ӵķ������˿�
        
        .ParaBlnAutoReStartServer = Val(GetSetting(.RegAppName, .RegSectionTCP, .RegKeyParaAutoReStartServer, 1))   '�ֶ�/�Զ���������ģʽ
        .ParaBlnAutoStartupAtBoot = Val(GetSetting(.RegAppName, .RegSectionSettings, .RegKeyParaAutoStartupAtBoot, 0))  '�����Զ�����
        
        .UserComputerName = VBA.Environ("ComputerName")
        .UserLoginName = VBA.Environ("UserName") '"XiaoMing"
        .UserFullName = "С��"
        
'''        '�ɷ���˷��������ͻ���
'''        .ConSource = gfCheckIP(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerIP, .TCPSetIP))   '����������/IP
'''        .ConDatabase = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerDatabase, EncryptString("dbTest", .EncryptKey)), .EncryptKey)    '���ݿ���
'''        .ConUserID = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerAccount, EncryptString("123", .EncryptKey)), .EncryptKey)  '��½��
'''        .ConPassword = DecryptString(gfGetRegStringValue(.RegAppName, .RegSectionDBServer, .RegKeyDBServerPassword, EncryptString("888888", .EncryptKey)), .EncryptKey)    '��½����
        
        
    End With
End Sub

Private Sub msResetLayout(ByRef cbsBars As XtremeCommandBars.CommandBars)
    '���ô��ڲ��֣�CommandBars��Dockingpane�ؼ�����
    
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

Private Sub msSetClientState(ByVal ColorSet As Long)
    '����״̬���пͻ������ӷ���˵�״̬
    
    Dim paneState As XtremeCommandBars.StatusBarPane
    
    Set paneState = mXtrStatusBar.FindPane(gID.StatusBarPaneConnectState)
    With paneState
        If ColorSet = vbGreen Then  '������
            .Text = gVar.ClientStateConnected
            .BackgroundColor = ColorSet
            gVar.TCPStateConnected = True   '�ñ�����¼״̬
        Else    '����
            gVar.TCPStateConnected = False
            If ColorSet = vbRed Then    '�����쳣
                .Text = gVar.ClientStateConnectError
                .BackgroundColor = ColorSet
            Else    'δ���ӵ�
                .Text = gVar.ClientStateDisConnected
                .BackgroundColor = vbYellow
            End If
        End If
    End With
    Set paneState = Nothing
    
End Sub

Private Sub msUnloadMe(Optional ByVal blnUnload As Boolean = True)
    'ж�ش���
    If Not blnUnload Then Exit Sub
    gVar.CloseWindow = True
    Unload Me
End Sub


Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '������¼�
    Call msLeftClick(Control.ID, Me.CommandBars1)
End Sub

Private Sub CommandBars1_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    'CommandBars�ؼ���Action״̬���л�
    
    Dim blnFC As Boolean    '�ж��Ƿ�ΪFC���
    Dim cbsActions As XtremeCommandBars.CommandBarActions  'cbs�ؼ�Actions���ϵ�����
    
    Set cbsActions = Me.CommandBars1.Actions
    If Screen.ActiveControl Is Nothing Then
        blnFC = False
    Else
        blnFC = TypeOf Screen.ActiveControl Is FlexCell.Grid    '��ǰ��ؼ���FC���
    End If
    With gID
        For mlngID = .SysExportToCSV To .SysExportToWord
            cbsActions(mlngID).Enabled = blnFC  '��ؼ���FC����򼤻��ӦAction������ʹ�䲻����
        Next
        For mlngID = .SysPrintPageSet To .SysPrint
            cbsActions(mlngID).Enabled = blnFC
        Next
    End With
End Sub

Private Sub MDIForm_Load()
    '�������
    
    Dim cbsBars As XtremeCommandBars.CommandBars
    Dim strUpdate As String
    
    ReDim gArr(1)   '��ʼ�����顣�ͻ���ͳһʹ���±�1������Timer1�ؼ���Winsocket�ؼ�
    Timer1.Item(1).Interval = 1000  '��ʱ��ѭ��ʱ��
    Set gWind = Me  'ָ���������ȫ�����ö���
    
    XtremeCommandBars.CommandBarsGlobalSettings.App = App 'һ��Ĭ������
    Set cbsBars = Me.CommandBars1
    
    Call Main   '��ʼ��ȫ�ֹ��ñ���
    Call msLoadParameter(True)  '�������ò���
    Call msAddAction(cbsBars)   '����Actions����
    Call msAddMenu(cbsBars)     '�����˵���
    Call msAddToolBar(cbsBars)  '����������
    Call msAddPopupMenu(cbsBars)    '��������ͼ��Ĳ˵�
    Call msAddXtrStatusBar(cbsBars) '����״̬��
    Call msAddKeyBindings(cbsBars)  '��ӿ�ݼ�,�ŵ�LoadCommandBars�������������Ч������
    Call msAddDesignerControls(cbsBars) 'CommandBars�Զ���Ի�����ʹ�õ�
    
    cbsBars.AddImageList ImageList1         'ʹCommandBars�ؼ�ƥ��ImageList�ؼ���ͼ��
    cbsBars.EnableCustomization True        '����CommandBars�Զ��壬��������÷�������CommandBars�趨֮��
    cbsBars.Options.UpdatePeriod = 250      '����CommandBars��Update�¼���ִ�����ڣ�Ĭ��100ms
    
    Call gsLoadSkin(Me, Me.SkinFramework1, sMSO7, True)  '���ش�������
    
    '���ع���������
    Call gsThemeCommandBar(Val(GetSetting(gVar.RegAppName, gVar.RegSectionSettings, gVar.RegKeyClientCommandbarsTheme, gID.WndThemeCommandBarsRibbon)), cbsBars)
    
    'ע�����Ϣ����-CommandBars����
    Call cbsBars.LoadCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegKeyCBSClientSetting)
    
    Set cbsBars = Nothing   '����ʹ����Ķ���
    
    Call gsFormSizeLoad(Me, False) 'ע�����Ϣ����-����λ�ô�С
    
    Call msConnectToServer(Me.Winsock1.Item(1), True)      '��������������
    
    strUpdate = gVar.AppPath & gVar.EXENameOfUpdate & " " & gVar.EXENameOfClient & _
            gVar.CmdLineSeparator & gVar.CmdLineParaOfHide      '��ʽ�򿪸��¼�����
    If Not gfShell(strUpdate) Then
        Rem MsgBox "���³��������쳣��", vbExclamation, "����"
        Rem Debug.Print strUpdate
        Call gsAlarmAndLog("���³��������쳣", False)
    End If
    
    '����Ƿ�Ϊ���ð�*******************************
    '==============================================
    
    
    '==============================================
    
    Call gfNotifyIconAdd(Me)    '�������ͼ��
    Me.Hide
    
    If LCase(App.EXEName & ".exe") <> LCase(gVar.EXENameOfClient) Then
        MsgBox "���������޸Ŀ�ִ�е�Ӧ�ó����ļ�����", vbCritical, "���ؾ���"
        Call msUnloadMe(True)    '��ֹexe�ļ�������
    End If
    
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '��Ӧ����ͼ��������Ҽ����������̲˵�
    Dim sngMsg As Single
    
    If Y <> 0 Then Exit Sub    '�ƺ��˾������ס���һ����������ͼ���ϣ������ڴ�����
    sngMsg = X / Screen.TwipsPerPixelX
    Select Case sngMsg
        Case WM_RBUTTONUP
            mcbsPopupIcon.ShowPopup  '�Ҽ�����Popup�˵�
        Case WM_LBUTTONDBLCLK   '���˫������ͼ��ʱ ��������ʾ/��С�� �л�
            Rem If Button <> vbLeftButton Then Exit Sub '���ڲ���ԭ���ż���Զ���С���ˣ��ƺ��˾������ס
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

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '�ж��Ƿ�����Ҫ�رմ���
    
    If gVar.ParaBlnWindowCloseMin Then
        If Not gVar.CloseWindow Then
            Cancel = True
            Me.WindowState = vbMinimized
        End If
        gVar.CloseWindow = False
    Else
        If Not gVar.CloseWindow Then
            If MsgBox("�Ƿ���С�����ڣ�", vbQuestion + vbYesNo, "�رջ���С��") = vbYes Then
                Cancel = True
                Me.WindowState = vbMinimized
            End If
        End If
    End If
End Sub

Private Sub MDIForm_Resize()
    '������С����ʾ
    If Me.Visible And Me.WindowState = vbMinimized Then
        If gVar.ParaBlnWindowMinHide Then
            Me.Hide
            Call gfNotifyIconBalloon(Me, "��С����ϵͳ����ͼ����", "��С����ʾ")
        End If
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'ж�ش���ʱ������Ϣ
    Dim resetNotifyIconData As gtypeNOTIFYICONDATA
    
    '����ע�����Ϣ-CommandBars����
    Call Me.CommandBars1.SaveCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegKeyCBSClientSetting)
    
    Call gsFormSizeSave(Me, False) '����ע�����Ϣ-����λ�ô�С
    Call gsSaveCommandbarsTheme(Me.CommandBars1, False)   '����CommandBars�ķ������
    
    gVar.CloseWindow = False    '���״̬-�رմ���
    gVar.ClientLoginShow = False '���״̬-��ʾ��½����
    Set gVar.rsURF = Nothing    '���Ȩ����Ϣ
    Call SkinFramework1.LoadSkin("", "")    '���Ƥ��
    Set mXtrStatusBar = Nothing  '���״̬��
    Set mcbsPopupIcon = Nothing '���Popup�˵�
    Call gfNotifyIconDelete(Me) 'ɾ������ͼ��
    gNotifyIconData = resetNotifyIconData   '�������������Ϣ��������������ʱ���Զ�����������ֻ�ܷ��Ͼ�ɾ������ͼ�����ĺ���?
    gArr(1) = gArr(0)
'    Erase gArr
    Set gWind = Nothing '���ȫ�ִ�������
    
End Sub

Private Sub Timer1_Timer(Index As Integer)
    Const conCon As Byte = 1    '����״̬�����conConn��
    Static byteCon As Byte
    Static byteChk As Byte
    
    byteCon = byteCon + 1
    byteChk = byteChk + 1
    
    If byteCon >= conCon Then
        With Me.Winsock1.Item(1)
            If .State = 7 Then  '������
                Call msSetClientState(vbGreen)
                If Not gVar.ClientLoginShow Then '������½����
                    If gVar.TCPStateConnected Then
                        If gArr(Index).Connected Then
                            frmSysLogin.Show vbModeless, Me
                            gVar.ClientLoginShow = True
                        End If
                    End If
                End If
            ElseIf .State = 9 Then  '�����쳣
                Call msSetClientState(vbRed)
            Else    'δ���ӵ�
                Call msSetClientState(vbYellow)
            End If
        End With
        byteCon = 0 '���㾲̬�ۻ�����
    End If
    
    If byteChk > (gVar.TCPWaitTime + 2) Then  '��Ϊ��������Ҳ�ǵȴ�gVar.TCPWaitTime�ŶϿ����ӣ������ӳ�һ��
        byteChk = 0
        If Not gVar.TCPStateConnected Then
            Call gsAlarmAndLogEx("���������������ʧ�ܣ���ȷ�Ϸ���˳���������", "���Ӿ�ʾ")
            Call msUnloadMe(True)
        End If
    End If
    
End Sub

Private Sub Winsock1_Close(Index As Integer)
    '���ӹر�ʱ��մ�����Ϣ
    If UBound(gArr) = 1 Then gArr(1) = gArr(0)
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strGet As String
    Dim byteGet() As Byte
    
    With gArr(Index)
        If Not .FileTransmitState Then
            '�ַ���Ϣ����״̬��
            
            Me.Winsock1.Item(Index).GetData strGet  '�����ַ�
            
            If InStr(strGet, gVar.PTClientConfirm) > 0 Then '�յ�Ҫ�ظ������ȷ�����ӵ���Ϣ
                Call gfSendInfo(gVar.PTClientIsTrue, Me.Winsock1.Item(Index))
                gArr(Index).Connected = True
            ElseIf InStr(strGet, gVar.PTConnectIsFull) Then '�յ�����˷���������������
                Me.Timer1.Item(Index).Enabled = False
                MsgBox "�ͻ������������������ޣ��������û��˳������ԣ�", vbCritical, "��������������"
                Call msUnloadMe(True)
                End
                
            ElseIf InStr(strGet, gVar.PTConnectTimeOut) Then '��������ʱ���ѵ�
                Me.Timer1.Item(Index).Enabled = False
                MsgBox "���������������ʱ���ѵ��������µ�½��", vbExclamation, "����ʱ��������ʾ"
                Call msUnloadMe(True)
                Load Me
                
            ElseIf InStr(strGet, gVar.PTFileStart) > 0 Then '���Է����ļ���������˵�״̬
                Call gfSendFile(.FilePath, Me.Winsock1.Item(Index)) '�����ļ��������
                Call gsFormEnable(Me, False)    '��ֹ�ͻ����ٲ���
                
            ElseIf InStr(strGet, gVar.PTFileExist) > 0 Then '����˷����ͻ�����Ҫ���ļ�����
                Dim strSize As String, lngInstrSize As Long
                
                lngInstrSize = InStr(strGet, gVar.PTFileSize) '��ȡ�ͻ�����Ҫ�Ĵ����ڷ���˵��ļ��Ĵ�С
                If lngInstrSize > 0 Then
                    strSize = Mid(strGet, lngInstrSize + Len(gVar.PTFileSize))
                    If IsNumeric(strSize) Then
                        .FileSizeTotal = Val(strSize)
                        Call gfSendInfo(gVar.PTFileSend, Me.Winsock1.Item(Index)) '֪ͨ����˿��Է��͹�����
                        '��ʼ��������
                        
                        .FileTransmitState = True
                        Call gsFormEnable(Me, False)    '��ֹ�ͻ����ٲ���
                    End If
                End If
                
            ElseIf InStr(strGet, gVar.PTFileNoExist) > 0 Then   '
                MsgBox "��Ҫ�ļ�<" & .FileName & ">�ڷ���˲����ڣ�", vbExclamation, "�ļ�����"
                gArr(Index) = gArr(0)
                
            End If
            
            Debug.Print "Client GetInfo:" & strGet, bytesTotal
            '�ַ���Ϣ����״̬��
        Else
            '�ļ�����״̬��
            
            If .FileNumber = 0 Then '�����ļ���
                .FileNumber = FreeFile
                Open .FilePath For Binary As #.FileNumber
            End If
            
            ReDim byteGet(bytesTotal - 1)   '�ض��������С
            Me.Winsock1.Item(Index).GetData byteGet, vbArray + vbByte   '�����ļ���Ϣ����������
            Put #.FileNumber, , byteGet '������ļ���
            .FileSizeCompleted = .FileSizeCompleted + bytesTotal    '��¼�Ѵ����С
            '���½�����
            
            If .FileSizeCompleted >= .FileSizeTotal Then    '������ɺ��һЩ����
                Close #.FileNumber
                Call gsFormEnable(Me, True) '����ͻ��˵�����
                gArr(Index) = gArr(0)
                Call gfSendInfo(gVar.PTFileEnd, Me.Winsock1.Item(Index)) '���ͽ�����־
                Debug.Print "Client Received Over"
            End If
            
            '�ļ�����״̬��
        End If
    End With
    
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '�����쳣����
    
    If Index <> 0 Then
        If gArr(Index).FileTransmitState Then   '�쳣ʱ����ļ�������Ϣ
            Debug.Print "ClientWinsockError:" & Index & "--" & Err.Number & "  " & Err.Description
            Close
            gArr(Index) = gArr(0)
            Call gsFormEnable(Me, True)
            Call gsAlarmAndLog("����������ӷ����쳣", False)
        End If
    End If
End Sub

Private Sub Winsock1_SendComplete(Index As Integer)
    '�����괦��
    
    If Index = 0 Then Exit Sub
    With gArr(Index)
        If .FileTransmitState Then
            If .FileSizeCompleted < .FileSizeTotal Then
                Call gfSendFile(.FilePath, Me.Winsock1.Item(Index))
            Else
                gArr(Index) = gArr(0)
                Call gsFormEnable(Me, True)
                Debug.Print "Client Send File Over"
            End If
        End If
    End With
End Sub
