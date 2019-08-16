Attribute VB_Name = "modDeclare"
Option Explicit


'''鼠标光标变手指状
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long  'SetCursor确定光标形状
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, _
        ByVal lpCursorName As String) As Long   'LoadCursor载入指定光标资源
Public Const IDC_HAND = "#32649"


'''查找窗口，发送信息
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'''使用 ShellExecute 打开文件或执行程序
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'hWnd：用于指定父窗口句柄。当函数调用过程出现错误时，它将作为Windows消息窗口的父窗口
'Operation：用于指定要进行的操作。其中:
'''edit 用编辑器打开 lpFile 指定的文档，如果 lpFile 不是文档，则会失败;
'''explore 浏览 lpFile 指定的文件夹
'''find 搜索 lpDirectory 指定的目录
'''open 打开 lpFile 文件，lpFile 可以是文件或文件夹
'''print 打印 lpFile，如果 lpFile 不是文档，则函数失败
'''properties 显示属性
'''runas 请求以管理员权限运行，比如以管理员权限运行某个exe
'''NULL 执行默认”open”动作
'FileName：用于指定要打开的文件名、要执行的程序文件名或要浏览的文件夹名
'Parameters：若FileName参数是一个可执行程序，则此参数指定命令行参数，否则此参数应为nil或PChar(0)
'Directory：用于指定默认目录
'ShowCmd：若FileName参数是一个可执行程序，则此参数指定程序窗口的初始显示方式，否则此参数应设置为0
'若ShellExecute函数调用成功，则返回值为被执行程序的实例句柄。若返回值小于32，则表示出现错误,错误如下:
Public Const NO_ERROR = 0   '系统内存或资源不足
Public Const ERROR_FILE_NOT_FOUND = 2&  '找不到指定的文件
Public Const ERROR_PATH_NOT_FOUND = 3&  '找不到指定路径
Public Const ERROR_BAD_FORMAT = 11&     '.exe文件无效
Public Const SE_ERR_ACCESSDENIED = 5    '拒绝访问指定文件
Public Const SE_ERR_ASSOCINCOMPLETE = 27    '文件名关联无效或不完整
Public Const SE_ERR_DDEBUSY = 30    'DDE事务正在处理，DDE事务无法完成
Public Const SE_ERR_DDEFAIL = 29    'DDE事务失败
Public Const SE_ERR_DDETIMEOUT = 28 '请求超时，无法完成DDE事务请求
Public Const SE_ERR_DLLNOTFOUND = 32    '未找到指定dll
Public Const SE_ERR_FNF = 2         '未找到指定文件
Public Const SE_ERR_NOASSOC = 31    '未找到与给的文件拓展名关联的应用程序，比如打印不可打印的文件等
Public Const SE_ERR_OOM = 8         '内存不足，无法完成操作
Public Const SE_ERR_PNF = 3         '未找到指定路径
Public Const SE_ERR_SHARE = 26      '发生共享冲突
'ShellExecute参数nShowCmd所用的常量ShowWindow() Commands
Public Const SW_HIDE = 0        '隐藏窗口，活动状态给令一个窗口
Public Const SW_SHOWNORMAL = 1  '与SW_RESTORE相同
Public Const SW_NORMAL = 1      '
Public Const SW_SHOWMINIMIZED = 2   '最小化窗口，并将其激活
Public Const SW_SHOWMAXIMIZED = 3   'SHOWMAXIMIZED 最大化窗口，并将其激活
Public Const SW_MAXIMIZE = 3        '
Public Const SW_SHOWNOACTIVATE = 4  '用最近的大小和位置显示一个窗口，同时不改变活动窗口
Public Const SW_SHOW = 5            '用当前的大小和位置显示一个窗口，同时令其进入活动状态
Public Const SW_MINIMIZE = 6        '最小化窗口，活动状态给令一个窗口
Public Const SW_SHOWMINNOACTIVE = 7 '最小化一个窗口，同时不改变活动窗口
Public Const SW_SHOWNA = 8          '用当前的大小和位置显示一个窗口，不改变活动窗口
Public Const SW_RESTORE = 9         '用原来的大小和位置显示一个窗口，同时令其进入活动状态
Public Const SW_SHOWDEFAULT = 10    '
Public Const SW_MAX = 10            '

'''注册表操作API与类型
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Const HKEY_USER_RUN As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"  '软件开机自动启动注册表子键位置

Public Enum genumRegRootDirectory   '注册表根键
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
End Enum

Public Enum genumRegDataType    '注册表值类型
    REG_SZ = 1          ' Unicode nul terminated string
    REG_EXPAND_SZ = 2   ' Unicode nul terminated string
    REG_BINARY = 3      ' Free form binary
    REG_DWORD = 4       ' 32-bit number
End Enum

Public Enum genumRegOperateType '注册表操作类型
    RegRead = 1
    RegWrite = 2
    RegDelete = 3
End Enum

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)  '程序暂停运行（毫秒）

'返回电脑信息API
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Enum genumComputerInfoType   '要返回的电脑上的信息类别
    ciComputerName
    ciUserName
End Enum


'''以下API函数Shell_NotifyIcon与一堆常量、枚举、结构体都有关托盘
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
    lpData As gtypeNOTIFYICONDATA) As Long

Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const IMAGE_ICON = 1

Public Const NIF_ICON = &H2     'hIcon成员起作用
Public Const NIF_INFO = &H10    '使用气球提示 代替普通的提示框
Public Const NIF_MESSAGE = &H1  'uCallbackMessage成员起作用
Public Const NIF_STATE = &H8    'dwState和dwStateMask成员起作用
Public Const NIF_TIP = &H4      'szTip成员起作用

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

Public Const NOTIFYICON_VERSION = 3 '使用Windows2000风格，另一个常量值0表示使用Windows95风格

Public Const NIS_HIDDEN = &H1       '图标隐藏
Public Const NIS_SHAREDICON = &H2   '图标共享

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
    cbSize As Long  '结构大小（字节）
    hwnd As Long    '处理消息的窗口句柄
    uID As Long     '托盘图标的标识符
    uFlags As Long  '此成员表明哪些其他成员起作用
    uCallbackMessage As Long        '应用程序定义的消息标示
    hIcon As Long                   '托盘图标句柄
    szTip As String * 128   '提示信息,不知为何，长度不设128弹不出气泡

    dwState As Long         '图标状态
    dwStateMask As Long     '指明dwState成员的哪些位可以被设置或访问
    szInfo As String * 256      '气球提示信息
    uTimeoutOrVersion As Long  '气球提示消失时间或版本
    szInfoTitle As String * 64  '气球提示标题
    dwInfoFlags As Long         '给气球提示框增加一个图标
End Type

Public Enum genumNotifyIconFlag
    NIIF_NONE = &H0     '没有图标
    NIIF_INFO = &H1     '信息图标
    NIIF_WARNING = &H2  '警告图标
    NIIF_ERROR = &H3    '错误图标
    NIIF_GUID = &H5     'Version6.0保留
    NIIF_ICON_MASK = &HF    'Version6.0保留
    NIIF_NOSOUND = &H10     'Version6.0禁止播放相应声音
End Enum

Public Enum genumNotifyIconMouseEvent  '鼠标事件
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

'''判断文件是否被打开用的API函数与常量
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As String, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const INVALID_HANDLE_VALUE = -1



Public Enum genumFileTransimitType    '文件传输类型枚举
    ftSend = 1      '发送
    ftReceive = 2   '接收
End Enum

Public Enum genumFileProgressValue  '文件传输进度设置枚举
    ftZero  '将进度设置为0
    ftOver  '将进度设置为100(满了，100%)
    ftRate  '将进度设置为当前传入的进度值
End Enum

Public Enum genumSkinResChoose  '窗体风格资源文件选择
    sNone = 0   '无
    sMSVst = 1  'MicrosoftVista风格
    sMSO7 = 2   'MicrosoftOffice2007风格
    sMSO10 = 3   'MicrosoftOffice2010风格
End Enum

Public Const gconAscAdd As Integer = 5      '简单加解密中字符转化的增量
Public Const gconAddLenStart As Integer = 10    '加在密文开始部分的字符个数
Public Const gconSumLen As Integer = 60     '密文的总字符数
Public Const gconMaxPWD As Integer = 20     '密码的最大字符数

'''自定义公用常量
Public Type gtypeCommonVariant
    TCPSetIP As String     '设置IP地址
    TCPSetPort As Long     '设置端口号
    TCPConnectMax As Long    '最大连接数
    TCPDefaultIP As String      '默认IP地址
    TCPDefaultPort As Long      '默认端口号
    TCPWaitTime As Long     '连接确认等待时间
    
    TCPStateConnected As Boolean     '客户端连接服务端成功标识
    TCPStateServerStarted As Boolean '服务器启动标识
    
    UpdatePCName As String  '启动更新程序的电脑名
    UpdateAccount As String '启动更新程序的账号
    UpdateUserName As String    '启动更新程序的用户名
    
    FTChunkSize As Long   '文件传输时的分块大小
    FTWaitTime As Long    '每段文件传输时的等待时间，单位秒
    FTIsOver As Boolean     '文件传输结束状态：False没传输完,True传输完毕.
    
    EncryptKey As String    '加密解密的密钥
        
    ServerButtonStart As String       '服务器状态：启动服务
    ServerButtonClose As String       '关闭服务
    ServerStateError As String       '异常
    ServerStateStarted As String     '已启动
    ServerStateNotStarted As String  '未启动
    
    ClientStateConnected As String             '客户端状态：已连接
    ClientStateDisConnected As String          '未连接
    ClientStateConnectError As String          '连接异常
    ClientButtonConnectToServer As String       '建立连接
    ClientButtonDisConnectFromServer As String  '断开连接
    
    PTFileName As String    '协议：文件名标识
    PTFileSize As String    '协议：文件大小标识
    PTFileFolder As String  '协议：文件要保存的文件夹名标识
    PTFileStart As String   '协议：文件开始传输标识
    PTFileEnd As String     '协议：文件结束传输标识
    PTFileSend As String    '协议：文件发送标识
    PTFileReceive As String '协议：文件接收标识
    PTFileExist As String  '协议：文件存在标识
    PTFileNoExist As String    '协议：文件不存在标识
    
    PTVersionOfClient As String     '协议：客户端版本号
    PTVersionNotUpdate As String    '协议：不需要更新
    PTVersionNeedUpdate As String   '协议：需要更新
    
    PTClientConfirm As String   '协议：客户端确认
    PTClientIsTrue As String    '协议：客户端给服务端的确认
    
    PTDBDataSource As String    '协议：数据库服务器地址
    PTDBDatabase As String      '协议：数据库名
    PTDBUserID As String        '协议：数据库访问账号
    PTDBPassword As String      '协议：数据库访问密码
    
    PTConnectIsFull As String   '协议：连接数已满
    PTConnectTimeOut As String  '协议：连续连接时间到
    
    PTClientUserComputerName As String  '协议：客户端计算机名
    PTClientUserLoginName As String '协议：客户端用户登陆名
    PTClientUserFullName As String  '协议：客户端用户姓名
    
    EXENameOfClient As String   '客户端程序exe文件名
    EXENameOfUpdate As String   '更新端程序exe文件名
    EXENameOfServer As String   '服务端程序exe文件名
    EXENameOfSetup As String    '更新安装包exe文件名
    
    CmdLineSeparator As String  '命令行间隔符
    CmdLineParaOfHide As String '命令行参数之隐藏
    
    '''SaveSetting(appname, section, key, setting)函数中参数的设置
    '''GetSetting(appname, section, key[, default])
    
    RegAppName As String        'SaveSettin OR GetSetting函数中AppName值
    RegSectionTCP As String     '参数section_TCP值
    RegKeyTCPIP As String       '参数key_IP值
    RegKeyTCPPort As String     '参数key_port值
    
    RegSectionSkin As String    '参数section_Skin
    RegKeySkinRes As String    '参数Key_SkinRes
    RegKeySkinIni As String    '参数Key_SkinIni
    RegKeySkinSvrRes As String    '参数Key_SkinSvrRes
    RegKeySkinSvrIni As String    '参数Key_SkinSvrIni
    
    
    RegSectionDBServer As String  '数据库服务器信息块
    RegKeyDBServerIP As String    '数据库服务器IP
    RegKeyDBServerDatabase As String    '数据库名
    RegKeyDBServerAccount As String   '数据库服务器连接账号
    RegKeyDBServerPassword As String  '数据库服务器连接密码
    RegKeyServerBackStore As String     '服务器端资料文件备份路径
    RegKeyServerBackInterval As String  '备份频率
    RegKeyServerBackTime As String      '备份时间
    RegKeyServerBackIntervalDays As String  '每N天备份间隔天数
        
    RegSectionUser As String    'Section_用户信息
    RegKeyUserLast As String    '最后登陆用户名
    RegKeyUserList As String    '曾经登陆过年用户名列表
    
    RegKeyCommandBars As String 'SaveCommandBars参数RegistryKey
    RegKeyCBSServerSetting As String    'Server上的CBS控件注册信息保存Key值
    RegKeyCBSClientSetting As String    'Client上的CBS控件注册信息保存Key值
    RegKeyDockingPane As String
    RegKeyDockPaneServerSetting As String
    RegKeyDockPaneClientSetting As String
    
    RegSectionSettings As String    'Section_Settings区
    RegKeyServerWindowLeft As String  'Server上Key_窗口Left值
    RegKeyServerWindowTop As String   '
    RegKeyServerWindowWidth As String '
    RegKeyServerWindowHeight As String    '
    RegKeyServerWindowStateMax As String
    RegKeyServerCommandbarsTheme As String    '
    
    
    RegKeyClientWindowLeft As String  'Client上Key_窗口Left值
    RegKeyClientWindowTop As String   '
    RegKeyClientWindowWidth As String '
    RegKeyClientWindowHeight As String    '
    RegKeyClientWindowStateMax As String
    RegKeyClientCommandbarsTheme As String    '
    RegKeyClientTaskPanelTheme As String '导航菜单主题
    RegKeyClientTaskPanelAutoFold As String '导航菜单自动折叠
    
    
    RegTrailPath As String  '注册表中HKEY_CURRENT_USER下SOFTWARE路径
    RegTrailKey As String   '试用信息-Key值
    TrailPeriod As Long     '试用期周期
    
    RegKeyParaWindowMinHide As String   '参数Key-窗口最小化隐藏
    RegKeyParaWindowCloseMin As String  '参数Key-窗口点击关闭时默认最小化
    RegKeyParaWindowStartMinS As String  '软件启动时自动最小化
    RegKeyParaWindowStartMinC As String  '软件启动时自动最小化
    RegKeyParaAutoReStartServer As String   '服务端是否自动重启服务
    RegKeyParaAutoStartupAtBoot As String   '开机自动启动
    RegKeyParaLimitClientConnect As String  '限制客户端连接
    RegKeyParaLimitClientConnectTime As String '限制客户端连接时长
    RegKeyParaLimitClientConnectNumber As String '限制客户端连接数
    
    RegKeyParaRememberUserList As String  '记住用户名
    RegKeyParaRememberUserPassword As String  '记住密码
    RegKeyParaUserAutoLogin As String '自动登陆
    
    AppPath As String           'App路径，确保最后字符为"\"
    FolderNameTemp As String    '文件夹名称：Temp的全路径
    FolderNameData As String    '文件夹名称：Data的全路径
    FolderNameBin As String     '文件夹名称：Bin的全路径
    FolderNameBackup As String  '文件夹名称：Backup的全路径
    FolderNameStore As String   '文件夹名称：Store的全路径
    FolderBin As String     '文件夹名称：Bin
    FolderData As String    '文件夹名称：Data
    FolderTemp As String    '文件夹名称：Temp
    FolderStore As String   '文件夹名称：Store
    FolderBackup As String  '文件夹名称：Backup
    
    
    FileNameErrLog As String    '错误记录日志文件的全路径
    FileNameSkin As String      '主题资源文件名
    FileNameSkinIni As String   '主题配置文件名
    FileNameLoginLog As String  '登陆日志文件名
    
    UserAutoID As String    '用户标识ID
    UserLoginName As String '用户登陆名
    UserNickName As String  '用户昵称
    UserFullName As String  '用户姓名
    UserPassword As String  '用户密码
    UserDepartment As String    '用户所在部门
    UserLoginIP As String       '用户登陆电脑IP
    UserComputerName As String  '用户登陆电脑名称
    
    rsURF As New ADODB.Recordset '保存用户的所有权限信息
    
    AccountAdmin As String         '特别账号：系统管理员
    AccountSystem As String        '特别账号：系统管理员
    
    ConSource As String      '连接数据库服务器名称或IP地址
    ConUserID As String      '连接数据库用户名
    ConPassword As String    '连接数据库密码
    ConDatabase As String    '连接的数据库名
    ConString As String      '数据库连接字符串全称
    
    FuncButton As String    '功能类别：按钮
    FuncForm As String      '功能类别：窗口
    FuncControl As String   '功能类别：其它控件
    FuncMainMenu As String  '功能类别：主菜单
    
    Formaty_M_dH_m_s As String  '时间格式yyyy-MM-dd HH:mm:ss
    Formatymdhms As String       '时间格式yyyyMMddHHmmss
    
    WindowWidth As Long     '窗口默认宽度
    WindowHeight As Long    '窗口默认高度
    
    CloseWindow As Boolean '是否真正关闭窗口。或者说是否点击了窗口右上角的关闭按钮
    ClientLoginShow As Boolean '显示客户端登陆窗口
    ClientReLoad As Boolean   '接收到服务端发来的重启客户端标志
    ShowMainWindow As Boolean '客户端成功登陆后显示过主窗体标志
    UpdateRunOver As Boolean   '更新程序是否运行完成
    UnloadFromLogin As Boolean '从登陆窗口传过来的关闭程序指令
    RestoreDBInfoOver As Boolean '接收完数据库连接信息
    ClientLoginCheckOver As Boolean '登陆检验完成
    ClientCancelAutoLogin As Boolean '登陆界面中手动临时取消自动登陆
    
    ParaBlnWindowStateMaxClient As Boolean '窗口上次关闭时是否最大化
    ParaBlnWindowStateMaxServer As Boolean
    ParaBlnWindowMinHide As Boolean '主窗口最小化时是否隐藏
    ParaBlnWindowCloseMin As Boolean    '主窗口点击关闭按钮时最小化
    ParaBlnWindowStartMinS As Boolean    '窗口启动时自动最小化
    ParaBlnWindowStartMinC As Boolean    '窗口启动时自动最小化
    ParaBlnAutoReStartServer As Boolean '服务端程序断开服务时自动重新开启服务
    ParaBlnAutoStartupAtBoot As Boolean '开机自动启动
    ParaBlnLimitClientConnect As Boolean '限制客户端连接时间
    ParaLimitClientConnectTime As Long  '限制客户端最大连续连接时长是多少
    
    ParaBlnRememberUserList As Boolean  '记住用户名
    ParaBlnRememberUserPassword As Boolean  '记住密码
    ParaBlnUserAutoLogin As Boolean '自动登陆
    
    ParaBackupPath As String   '备份路径
    ParaBackupInterval As String    '备份频率
    ParaBackupTime As String        '备份时间
    ParaBackupIntervalDays As String '每N天的间隔天数
    
    FTUploadOrDownload As Boolean   '上传(True)还是下载(False)状态
    FTUploadFilePath As String      '上传文件的各项信息
    FTUploadFileNameOld As String
    FTUploadFileNameNew As String
    FTUploadFileSize As Long
    FTUploadFileFolder As String
    FTUploadFileClassify As String
    FTUploadFileExtension As String
    
    FTDownloadFilePath As String      '下载文件的各项信息
    FTDownloadFileNameOld As String
    FTDownloadFileNameNew As String
    FTDownloadFileSize As Long
    FTDownloadFileFolder As String
    FTDownloadFileClassify As String
    FTDownloadFileExtension As String
    
End Type

Public Type gtypeFileTransmitVariant    '自定义文件传输变量
    Connected As Boolean        '确认连接状态
    FileNumber As Integer       '文件传输时打开的文件号
    FilePath As String          '文件名，含全路径
    FileName As String          '仅文件名，不含路径
    FileFolder As String        '文件存储位置的文件夹名称，暂不支持其它路径，默认定在App.Path下
    FileSizeTotal As Long       '文件总大小
    FileSizeCompleted As Long   '文件已传输大小
    FileTransmitState As Boolean    '是否在传输文件
    FileClassify As String      '传送文件的存放类别
    FileOldName As String       '原文件名
    FileExtension As String     '原文件名的扩展名
    FileTransmitNotOver As Boolean  '文件传输结束标识
    FileTransmitError As Boolean    '传输是否异常结束
End Type

Public gVar As gtypeCommonVariant
Public gArr() As gtypeFileTransmitVariant

'''CommandBars的ID集合
Public Type gtypeCommandBarID
    
    Sys As Long             '模块-系统
    
    SysLoginOut As Long     '退出系统
    SysLoginAgain As Long   '重新登陆
    SysAuthChangePassword As Long   '密码修改
    SysAuthDepartment As Long       '部门管理
    SysAuthUser As Long     '用户管理
    SysAuthLog As Long      '日志管理
    SysAuthRole As Long     '角色管理
    SysAuthFunc As Long     '功能管理
    
    SysFileManage As Long   '文件管理
    
    SysPrintMain As Long
    SysPrint As Long        '打印
    SysPrintPageSet As Long '打印页面设置
    SysPrintPreview As Long '打印预览
    SysExportMain As Long
    SysExportToExcel As Long    '导出至Excel
    SysExportToWord As Long '导出至Word
    SysExportToText As Long '导出至文本
    SysExportToXML As Long  '导出为XML文档
    SysExportToPDF As Long  '导出为PDF
    SysExportToCSV As Long  '导出为CSV文件
    SysExportToHTML As Long '导出为HTML
    
    SysSearch As Long   '窗口检索工具栏
    SysSearch1Label As Long
    SysSearch2TextBox As Long
    SysSearch3Button As Long
    SysSearch4ListBoxCaption As Long
    SysSearch4ListBoxFormID As Long
    SysSearch5Go As Long
    
    Help As Long        '模块-帮助
    
    HelpAbout As Long   '关于
    HelpDocument As Long    '帮助文档
    HelpUpdate As Long  '检查更新
    
    
    Wnd As Long '模块-窗口控制
    
    WndThemeSkinSet As Long '窗口主题设置
    WndResetLayout As Long  '窗口布局重置
    
    WndOpenListCaption As Long '已打开窗口列表
    WndOpenListID As Long
    
    WndToolBarCustomize As Long '自定义工具栏
    WndToolBarList As Long '工具栏列表
    
    WndThemeCommandBars As Long '工具栏主题-CommandBars
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
    
    WndThemeTaskPanel As Long   '任务面板(导航菜单)主题-TaskPanel
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
    
    WndThemeSkin As Long    '系统皮肤主题-SkinFrameWork
    WndThemeSkinCodejock As Long
    WndThemeSkinOffice2007 As Long
    WndThemeSkinOffice2010 As Long
    WndThemeSkinVista As Long
    WndThemeSkinWinXPRoyale As Long
    WndThemeSkinWinXPLuna As Long
    WndThemeSkinZune As Long
    
    WndSon As Long  '子窗口控制
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
    
    
    Tool As Long    '模块--工具
    
    toolErrorLog As Long  '错误日志查看
    toolLoginLog As Long '登陆日志查看
    toolOptions As Long  '选项
    
    
'''**************************************************************'''

    
    Pane As Long   '模块--浮动面板
    
    PaneNavi As Long '导航菜单
    
    PanePopupMenuNavi As Long   '导航菜单任务面板弹出式菜单模块
    PanePopupMenuNaviExpandALL As Long  '展开所有
    PanePopupMenuNaviAutoFoldOther As Long  '自动折叠其它
    PanePopupMenuNaviFoldALL As Long    '折叠所有
    
    TabWorkspacePopupMenu As Long   '多标签右键菜单模块
    
    StatusBarPane As Long               '模块-状态栏面板
    
    StatusBarPaneProgress As Long       '状态栏中进度条
    StatusBarPaneProgressText As Long   '状态栏中进度百分值
    StatusBarPaneUserInfo As Long       '状态栏中用户信息
    StatusBarPaneTime As Long           '状态栏中时间
    StatusBarPaneConnectState As Long   '状态栏中连接状态-Client
    StatusBarPaneConnectButton As Long  '状态栏中连接按钮-Client
    StatusBarPaneServerState As Long    '状态栏中服务器服务状态-Server
    StatusBarPaneServerButton As Long   '状态栏中服务器开启/断开服务按钮-Server
    StatusBarPaneIP As Long     '状态栏中IP
    StatusBarPanePort As Long   '状态栏中端口
    StatusBarPaneReStartButton As Long  '状态栏中 自动重启服务按钮
    
    IconPopupMenu As Long           '托盘图标菜单
    IconPopupMenuShowWindow As Long '显示窗口
    IconPopupMenuMinWindow As Long  '窗口最小化
    IconPopupMenuMaxWindow As Long  '窗口最大化
    
End Type

Public Type gtypeValueAndErr    '用于返回布尔值的过程，顺便返回异常代号
    Result As Boolean
    ErrNum As Long
End Type

Public Enum genumFileOpenType   '打开文件方式
    udAppend    '以顺序型访问，把字符追加到文件
    udBinary    '以二进制访问
    udInput     '以顺序型访问，从文件输入字符
    udOutput    '以顺序型访问，向文件输出字符
    udRandom    '以随机方式
End Enum

Public Enum genumFileWriteType  '写入文件方式
    udPut       '用Get读出.For Binary、Random.
    udWrite     '用Input读出
    udPrint     '用Line Input 或 Input读出
End Enum

Public Enum genumCharType   '返回字符类型
    udUpperCase = 4     '仅大写字母
    udLowerCase = 1     '仅小写字母
    udNumber = 2        '仅数字
    udUpperLowerNum = 7 '大写、小写、数字
End Enum

Public Enum genumLogType    '操作日志类型增、删、改、查
    udSelect        '单个查询
    udInsert
    udDelete
    udUpdate
    udSelectBatch   '多个查询
    udInsertBatch
    udDeleteBatch
    udUpdateBatch
End Enum

Public Enum genumGridExportType 'Flexcell Grid导出文件的类型
    fcCSV
    fcExcel
    fcHTML
    fcPDF
    fcXML
End Enum

Public Enum genumNumber '顺序数字
    eZero = 0
    eOne = 1
    eTwo = 2
    eThree = 3
    eFour = 4
    eFive = 5
End Enum


Public gID As gtypeCommandBarID '主窗体中的全局CommandBars的ID变量
Public gWind As Form            '全局主窗体引用




