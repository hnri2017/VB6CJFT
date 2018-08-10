Attribute VB_Name = "modDeclare"
Option Explicit


'''���Ҵ��ڣ�������Ϣ
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'''ʹ�� ShellExecute ���ļ���ִ�г���
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'hWnd������ָ�������ھ�������������ù��̳��ִ���ʱ��������ΪWindows��Ϣ���ڵĸ�����
'Operation������ָ��Ҫ���еĲ���������:
'''edit �ñ༭���� lpFile ָ�����ĵ������ lpFile �����ĵ������ʧ��;
'''explore ��� lpFile ָ�����ļ���
'''find ���� lpDirectory ָ����Ŀ¼
'''open �� lpFile �ļ���lpFile �������ļ����ļ���
'''print ��ӡ lpFile����� lpFile �����ĵ�������ʧ��
'''properties ��ʾ����
'''runas �����Թ���ԱȨ�����У������Թ���ԱȨ������ĳ��exe
'''NULL ִ��Ĭ�ϡ�open������
'FileName������ָ��Ҫ�򿪵��ļ�����Ҫִ�еĳ����ļ�����Ҫ������ļ�����
'Parameters����FileName������һ����ִ�г�����˲���ָ�������в���������˲���ӦΪnil��PChar(0)
'Directory������ָ��Ĭ��Ŀ¼
'ShowCmd����FileName������һ����ִ�г�����˲���ָ�����򴰿ڵĳ�ʼ��ʾ��ʽ������˲���Ӧ����Ϊ0
'��ShellExecute�������óɹ����򷵻�ֵΪ��ִ�г����ʵ�������������ֵС��32�����ʾ���ִ���,��������:
Public Const NO_ERROR = 0   'ϵͳ�ڴ����Դ����
Public Const ERROR_FILE_NOT_FOUND = 2&  '�Ҳ���ָ�����ļ�
Public Const ERROR_PATH_NOT_FOUND = 3&  '�Ҳ���ָ��·��
Public Const ERROR_BAD_FORMAT = 11&     '.exe�ļ���Ч
Public Const SE_ERR_ACCESSDENIED = 5    '�ܾ�����ָ���ļ�
Public Const SE_ERR_ASSOCINCOMPLETE = 27    '�ļ���������Ч������
Public Const SE_ERR_DDEBUSY = 30    'DDE�������ڴ�����DDE�����޷����
Public Const SE_ERR_DDEFAIL = 29    'DDE����ʧ��
Public Const SE_ERR_DDETIMEOUT = 28 '����ʱ���޷����DDE��������
Public Const SE_ERR_DLLNOTFOUND = 32    'δ�ҵ�ָ��dll
Public Const SE_ERR_FNF = 2         'δ�ҵ�ָ���ļ�
Public Const SE_ERR_NOASSOC = 31    'δ�ҵ�������ļ���չ��������Ӧ�ó��򣬱����ӡ���ɴ�ӡ���ļ���
Public Const SE_ERR_OOM = 8         '�ڴ治�㣬�޷���ɲ���
Public Const SE_ERR_PNF = 3         'δ�ҵ�ָ��·��
Public Const SE_ERR_SHARE = 26      '����������ͻ
'ShellExecute����nShowCmd���õĳ���ShowWindow() Commands
Public Const SW_HIDE = 0        '���ش��ڣ��״̬����һ������
Public Const SW_SHOWNORMAL = 1  '��SW_RESTORE��ͬ
Public Const SW_NORMAL = 1      '
Public Const SW_SHOWMINIMIZED = 2   '��С�����ڣ������伤��
Public Const SW_SHOWMAXIMIZED = 3   'SHOWMAXIMIZED ��󻯴��ڣ������伤��
Public Const SW_MAXIMIZE = 3        '
Public Const SW_SHOWNOACTIVATE = 4  '������Ĵ�С��λ����ʾһ�����ڣ�ͬʱ���ı�����
Public Const SW_SHOW = 5            '�õ�ǰ�Ĵ�С��λ����ʾһ�����ڣ�ͬʱ�������״̬
Public Const SW_MINIMIZE = 6        '��С�����ڣ��״̬����һ������
Public Const SW_SHOWMINNOACTIVE = 7 '��С��һ�����ڣ�ͬʱ���ı�����
Public Const SW_SHOWNA = 8          '�õ�ǰ�Ĵ�С��λ����ʾһ�����ڣ����ı�����
Public Const SW_RESTORE = 9         '��ԭ���Ĵ�С��λ����ʾһ�����ڣ�ͬʱ�������״̬
Public Const SW_SHOWDEFAULT = 10    '
Public Const SW_MAX = 10            '

'''ע�������API������
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Const HKEY_USER_RUN As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"  '���������Զ�����ע����Ӽ�λ��

Public Enum genumRegRootDirectory   'ע�������
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
End Enum

Public Enum genumRegDataType    'ע���ֵ����
    REG_SZ = 1          ' Unicode nul terminated string
    REG_EXPAND_SZ = 2   ' Unicode nul terminated string
    REG_BINARY = 3      ' Free form binary
    REG_DWORD = 4       ' 32-bit number
End Enum

Public Enum genumRegOperateType 'ע�����������
    RegRead = 1
    RegWrite = 2
    RegDelete = 3
End Enum

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)  '������ͣ���У����룩


'''����API����Shell_NotifyIcon��һ�ѳ�����ö�١��ṹ�嶼�й�����
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
    lpData As gtypeNOTIFYICONDATA) As Long

Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const IMAGE_ICON = 1

Public Const NIF_ICON = &H2     'hIcon��Ա������
Public Const NIF_INFO = &H10    'ʹ��������ʾ ������ͨ����ʾ��
Public Const NIF_MESSAGE = &H1  'uCallbackMessage��Ա������
Public Const NIF_STATE = &H8    'dwState��dwStateMask��Ա������
Public Const NIF_TIP = &H4      'szTip��Ա������

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIM_SETFOCUS = &H3
Public Const NIM_SETVERSION = &H4
Public Const NIM_VERSION = &H5

Public Const WM_USER As Long = &H400
Public Const NIN_BALLOONSHOW = (WM_USER + 2)
Public Const NIN_BALLOONHIDE = (WM_USER + 3)
Public Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Public Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

Public Const NOTIFYICON_VERSION = 3 'ʹ��Windows2000�����һ������ֵ0��ʾʹ��Windows95���

Public Const NIS_HIDDEN = &H1       'ͼ������
Public Const NIS_SHAREDICON = &H2   'ͼ�깲��

Public Const WM_NOTIFY As Long = &H4E
Public Const WM_COMMAND As Long = &H111
Public Const WM_CLOSE As Long = &H10

Public Const WM_MOUSEMOVE As Long = &H200
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RBUTTONDBLCLK As Long = &H206
Public Const WM_MBUTTONDOWN As Long = &H207
Public Const WM_MBUTTONUP As Long = &H208
Public Const WM_MBUTTONDBLCLK As Long = &H209

Public Type gtypeNOTIFYICONDATA
    cbSize As Long  '�ṹ��С���ֽڣ�
    hwnd As Long    '������Ϣ�Ĵ��ھ��
    uID As Long     '����ͼ��ı�ʶ��
    uFlags As Long  '�˳�Ա������Щ������Ա������
    uCallbackMessage As Long        'Ӧ�ó��������Ϣ��ʾ
    hIcon As Long                   '����ͼ����
    szTip As String * 128   '��ʾ��Ϣ,��֪Ϊ�Σ����Ȳ���128����������

    dwState As Long         'ͼ��״̬
    dwStateMask As Long     'ָ��dwState��Ա����Щλ���Ա����û����
    szInfo As String * 256      '������ʾ��Ϣ
    uTimeoutOrVersion As Long  '������ʾ��ʧʱ���汾
    szInfoTitle As String * 64  '������ʾ����
    dwInfoFlags As Long         '��������ʾ������һ��ͼ��
End Type

Public Enum genumNotifyIconFlag
    NIIF_NONE = &H0     'û��ͼ��
    NIIF_INFO = &H1     '��Ϣͼ��
    NIIF_WARNING = &H2  '����ͼ��
    NIIF_ERROR = &H3    '����ͼ��
    NIIF_GUID = &H5     'Version6.0����
    NIIF_ICON_MASK = &HF    'Version6.0����
    NIIF_NOSOUND = &H10     'Version6.0��ֹ������Ӧ����
End Enum

Public Enum genumNotifyIconMouseEvent  '����¼�
    MouseMove = &H200
    LeftUp = &H202
    LeftDown = &H201
    LeftDbClick = &H203
    RightUp = &H205
    RightDown = &H204
    RightDbClick = &H206
    MiddleUp = &H208
    MiddleDown = &H207
    MiddleDbClick = &H209
    BalloonClick = (WM_USER + 5)
End Enum

Public gNotifyIconData As gtypeNOTIFYICONDATA


Public Enum genumFileTransimitType    '�ļ���������ö��
    ftSend = 1      '����
    ftReceive = 2   '����
End Enum

Public Enum genumSkinResChoose  '��������Դ�ļ�ѡ��
    sNone = 0   '��
    sMSVst = 1  'MicrosoftVista���
    sMSO7 = 2   'MicrosoftOffice2007���
    sMSO10 = 3   'MicrosoftOffice2010���
End Enum

Public Const gconAscAdd As Integer = 5      '�򵥼ӽ������ַ�ת��������
Public Const gconAddLenStart As Integer = 10    '�������Ŀ�ʼ���ֵ��ַ�����
Public Const gconSumLen As Integer = 60     '���ĵ����ַ���
Public Const gconMaxPWD As Integer = 20     '���������ַ���

'''�Զ��幫�ó���
Public Type gtypeCommonVariant
    TCPSetIP As String     'IP��ַ
    TCPSetPort As Long     '�˿�
    TCPSetConnectMax As Long  '���������
    
    TCPStateConnected As Boolean     '���ӳɹ���ʶ
    TCPStateServerStarted As Boolean '������������ʶ
    
    FTChunkSize As Long   '�ļ�����ʱ�ķֿ��С
    FTWaitTime As Long    'ÿ���ļ�����ʱ�ĵȴ�ʱ�䣬��λ��
        
    ServerStart As String       '������״̬����������
    ServerClose As String       '�رշ���
    ServerError As String       '�쳣
    ServerStarted As String     '������
    ServerNotStarted As String  'δ����
    
    StateConnected As String             '�ͻ���״̬��������
    StateDisConnected As String          'δ����
    StateConnectError As String          '�����쳣
    StateConnectToServer As String       '��������
    StateDisConnectFromServer As String  '�Ͽ�����
    
    PTFileName As String    'Э�飺�ļ�����ʶ
    PTFileSize As String    'Э�飺�ļ���С��ʶ
    PTFileFolder As String  'Э�飺�ļ�Ҫ������ļ�������ʶ
    PTFileStart As String   'Э�飺�ļ���ʼ�����ʶ
    PTFileEnd As String     'Э�飺�ļ����������ʶ
    PTFileSend As String    'Э�飺�ļ����ͱ�ʶ
    PTFileReceive As String 'Э�飺�ļ����ձ�ʶ
    PTFileExist As String  'Э�飺�ļ����ڱ�ʶ
    PTFileNoExist As String    'Э�飺�ļ������ڱ�ʶ
    
    PTVersionOfClient As String     'Э�飺�ͻ��˰汾��
    PTVersionNotUpdate As String    'Э�飺����Ҫ����
    PTVersionNeedUpdate As String   'Э�飺��Ҫ����
    
    PTClientConfirm As String   'Э�飺�ͻ���ȷ��
    PTClientIsTrue As String    'Э�飺�ͻ��˸�����˵�ȷ��
    PTWaitTime As Long
    
    EXENameOfClient As String   '�ͻ��˳���exe�ļ���
    EXENameOfUpdate As String   '���¶˳���exe�ļ���
    EXENameOfServer As String   '����˳���exe�ļ���
    EXENameOfSetup As String    '���°�װ��exe�ļ���
    
    CmdLineSeparator As String  '�����м����
    CmdLineParaOfHide As String '�����в���֮����
    
    '''SaveSetting(appname, section, key, setting)�����в���������
    '''GetSetting(appname, section, key[, default])
    
    RegAppName As String        'SaveSettin OR GetSetting������AppNameֵ
    RegSectionTCP As String     '����section_TCPֵ
    RegKeyTCPIP As String       '����key_IPֵ
    RegKeyTCPPort As String     '����key_portֵ
    
    RegSectionSkin As String    '����section_Skin
    RegKeySkinFile As String    '����Key_SkinFile
    
    RegSectionServer As String  '���ݿ��������Ϣ��
    RegKeyServerIP As String    '���ݿ������IP
    RegKeyServerAccount As String   '���ݿ�����������˺�
    RegKeyServerPassword As String  '���ݿ��������������
        
    RegSectionUser As String    'Section_�û���Ϣ
    RegKeyUserLast As String    '����½�û���
    RegKeyUserList As String    '������½�����û����б�
    
    RegSectionSettings As String    'Section_Settings��
    RegKeyCommandBars As String 'SaveCommandBars����RegistryKey
    RegKeyWindowLeft As String  'Key_����Leftֵ
    RegKeyWindowTop As String   '
    RegKeyWindowWidth As String '
    RegKeyWindowHeight As String    '
    RegKeyCommandbarsTheme As String    '
    
    RegTrailPath As String  'ע�����HKEY_CURRENT_USER��SOFTWARE·��
    RegTrailKey As String   '������Ϣ-Keyֵ
    TrailPeriod As Long     '����������
    
    AppPath As String           'App·����ȷ������ַ�Ϊ"\"
    FolderNameTemp As String    '�ļ������ƣ�Temp
    FolderNameData As String    '�ļ������ƣ�Data
    FolderNameBin As String     '�ļ������ƣ�Bin
    
    FileNameErrLog As String    '�����¼��־�ļ���ȫ·��
    FileNameSkin As String      '������Դ�ļ���
    FileNameSkinIni As String   '���������ļ���
    
    UserAutoID As String    '�û���ʶID
    UserLoginName As String '�û���½��
    UserNickName As String  '�û��ǳ�
    UserFullName As String  '�û�����
    UserPassword As String  '�û�����
    UserDepartment As String    '�û����ڲ���
    UserLoginIP As String       '�û���½����IP
    UserComputerName As String  '�û���½��������
    
    rsURF As New ADODB.Recordset '�����û�������Ȩ����Ϣ
    
    AccountAdmin As String         '�ر��˺ţ�ϵͳ����Ա
    AccountSystem As String        '�ر��˺ţ�ϵͳ����Ա
    
    ConSource As String      '�������ݿ���������ƻ�IP��ַ
    ConUserID As String      '�������ݿ��û���
    ConPassword As String    '�������ݿ�����
    ConDatabase As String    '���ӵ����ݿ���
    ConString As String      '���ݿ������ַ���ȫ��
    
    FuncButton As String    '������𣺰�ť
    FuncForm As String      '������𣺴���
    FuncControl As String   '������������ؼ�
    FuncMainMenu As String  '����������˵�
    
    WindowWidth As Long     '����Ĭ�Ͽ���
    WindowHeight As Long    '����Ĭ�ϸ߶�
    
End Type

Public Type gtypeFileTransmitVariant    '�Զ����ļ��������
    Connected As Boolean        'ȷ������״̬
    FileNumber As Integer       '�ļ�����ʱ�򿪵��ļ���
    FilePath As String          '�ļ�������ȫ·��
    FileName As String          '���ļ���������·��
    FileFolder As String        '�ļ��洢λ�õ��ļ������ƣ��ݲ�֧������·����Ĭ�϶���App.Path��
    FileSizeTotal As Long       '�ļ��ܴ�С
    FileSizeCompleted As Long   '�ļ��Ѵ����С
    FileTransmitState As Boolean    '�Ƿ��ڴ����ļ�
End Type

Public gVar As gtypeCommonVariant
Public gArr() As gtypeFileTransmitVariant

'''CommandBars��ID����
Public Type gtypeCommandBarID
    
    Sys As Long             'ģ��-ϵͳ
    
    SysLoginOut As Long     '�˳�ϵͳ
    SysLoginAgain As Long   '���µ�½
    SysAuthChangePassword As Long   '�����޸�
    SysAuthDepartment As Long       '���Ź���
    SysAuthUser As Long     '�û�����
    SysAuthLog As Long      '��־����
    SysAuthRole As Long     '��ɫ����
    SysAuthFunc As Long     '���ܹ���
    
    SysPrint As Long        '��ӡ
    SysPrintPageSet As Long '��ӡҳ������
    SysPrintPreview As Long '��ӡԤ��
    SysExportToExcel As Long    '������Excel
    SysExportToWord As Long '������Word
    SysExportToText As Long '�������ı�
    SysExportToXML As Long  '����ΪXML�ĵ�
    SysExportToPDF As Long  '����ΪPDF
    
    SysSearch As Long           '������
    SysSearch1Label As Long     '
    SysSearch2TextBox As Long   '
    SysSearch3Button As Long    '
    SysSearch4ListBoxCaption As Long    '
    SysSearch4ListBoxFormID As Long '
    SysSearch5Go As Long    '
    
    
    TestWindow As Long  'ģ��-����
    
    TestWindowFirst As Long '
    TestWindowSecond As Long
    TestWindowThird As Long
    TestWindowThour As Long
    TestWindowMDB As Long
    
    
    Help As Long        'ģ��-����
    
    HelpAbout As Long   '����
    HelpDocument As Long    '�����ĵ�
    HelpUpdate As Long  '������
    
    
    Wnd As Long 'ģ��-���ڿ���
    
    WndResetLayout As Long  '���ڲ�������
    
    TabWorkspacePopupMenu As Long   '���ǩ�Ҽ��˵�ģ��
    
    WndThemeCommandBars As Long '����-CommandBars
    WndThemeCommandBarsOffice2000 As Long
    WndThemeCommandBarsOfficeXp As Long
    WndThemeCommandBarsOffice2003 As Long
    WndThemeCommandBarsWinXP As Long
    WndThemeCommandBarsWhidbey As Long
    WndThemeCommandBarsResource As Long
    WndThemeCommandBarsRibbon As Long
    WndThemeCommandBarsVS2008 As Long
    WndThemeCommandBarsVS6 As Long
    WndThemeCommandBarsVS2010 As Long
    
    WndThemeTaskPanel As Long   '����-TaskPanel
    WndThemeTaskPanelOffice2000 As Long
    WndThemeTaskPanelOffice2003 As Long
    WndThemeTaskPanelNativeWinXP As Long
    WndThemeTaskPanelOffice2000Plain As Long
    WndThemeTaskPanelOfficeXPPlain As Long
    WndThemeTaskPanelOffice2003Plain As Long
    WndThemeTaskPanelNativeWinXPPlain As Long
    WndThemeTaskPanelToolbox As Long
    WndThemeTaskPanelToolboxWhidbey As Long
    WndThemeTaskPanelListView As Long
    WndThemeTaskPanelListViewOfficeXP As Long
    WndThemeTaskPanelListViewOffice2003 As Long
    WndThemeTaskPanelShortcutBarOffice2003 As Long
    WndThemeTaskPanelResource As Long
    WndThemeTaskPanelVisualStudio2010 As Long
    
    WndThemeSkin As Long    '����-SkinFrameWork
    WndThemeSkinCodejock As Long
    WndThemeSkinOffice2007 As Long
    WndThemeSkinOffice2010 As Long
    WndThemeSkinVista As Long
    WndThemeSkinWinXPRoyale As Long
    WndThemeSkinWinXPLuna As Long
    WndThemeSkinZune As Long
    WndThemeSkinSet As Long
    
    WndSon As Long  '�Ӵ��ڿ���
    WndSonVbCascade As Long
    WndSonVbTileHorizontal As Long
    WndSonVbTileVertical As Long
    WndSonVbArrangeIcons As Long
    WndSonVbAllMin As Long
    WndSonVbAllBack As Long
    WndSonCloseAll As Long
    WndSonCloseCurrent As Long
    WndSonCloseLeft As Long
    WndSonCloseRight As Long
    WndSonCloseOther As Long
    
        
    Pane As Long   'ģ��--�������
    
    PaneIDFirst As Long     '���ID
    PaneTitleFirst As Long  '������
    
    PanePopupMenu As Long   '��嵯��ʽ�˵�ģ��
    PanePopupMenuExpandALL As Long  'չ������
    PanePopupMenuAutoFoldOther As Long  '�Զ��۵�����
    PanePopupMenuFoldALL As Long    '�۵�����
    
    
    StatusBarPane As Long               'ģ��-״̬�����
    
    StatusBarPaneProgress As Long       '״̬���н�����
    StatusBarPaneProgressText As Long   '״̬���н��Ȱٷ�ֵ
    StatusBarPaneUserInfo As Long       '״̬�����û���Ϣ
    StatusBarPaneTime As Long           '״̬����ʱ��
    StatusBarPaneConnectState As Long   '״̬��������״̬-Client
    StatusBarPaneConnectButton As Long  '״̬�������Ӱ�ť-Client
    StatusBarPaneServerState As Long    '״̬���з���������״̬-Server
    StatusBarPaneServerButton As Long   '״̬���з���������/�Ͽ�����ť-Server
    
    
End Type

Public Type gtypeValueAndErr    '���ڷ��ز���ֵ�Ĺ��̣�˳�㷵���쳣����
    Result As Boolean
    ErrNum As Long
End Type

Public Enum genumFileOpenType   '���ļ���ʽ
    udAppend    '��˳���ͷ��ʣ����ַ�׷�ӵ��ļ�
    udBinary    '�Զ����Ʒ���
    udInput     '��˳���ͷ��ʣ����ļ������ַ�
    udOutput    '��˳���ͷ��ʣ����ļ�����ַ�
    udRandom    '�������ʽ
End Enum

Public Enum genumFileWriteType  'д���ļ���ʽ
    udPut       '��Get����.For Binary��Random.
    udWrite     '��Input����
    udPrint     '��Line Input �� Input����
End Enum

Public Enum genumCharType   '�����ַ�����
    udUpperCase = 4     '����д��ĸ
    udLowerCase = 1     '��Сд��ĸ
    udNumber = 2        '������
    udUpperLowerNum = 7 '��д��Сд������
End Enum

Public Enum genumLogType    '������־��������ɾ���ġ���
    udSelect        '������ѯ
    udInsert
    udDelete
    udUpdate
    udSelectBatch   '�����ѯ
    udInsertBatch
    udDeleteBatch
    udUpdateBatch
End Enum


Public gID As gtypeCommandBarID '�������е�ȫ��CommandBars��ID����
Public gWind As Form            'ȫ������������



