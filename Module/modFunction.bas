Attribute VB_Name = "modFunction"
Option Explicit


'浏览目录所使用的常量、API、Type、变量等。功能：定位到当前文件夹，而且选定它
Private Const BIF_RETURNONLYFSDIRS = 1  '仅仅返回文件系统的目录
Private Const BIF_DONTGOBELOWDOMAIN = 2 '在树形视窗中，不包含域名底下的网络目录结构
Private Const BIF_STATUSTEXT = &H4&     '在对话框中包含一个状态区域
Private Const BIF_RETURNFSANCESTORS = 8 '返回文件系统的一个节点
Private Const BIF_EDITBOX = &H10& ' 16  '浏览对话框中包含一个编辑框
Private Const BIF_VALIDATE = &H20& '32  '当没有BIF_EDITBOX标志位时，该标志位被忽略
Private Const BIF_NEWDIALOGSTYLE = &H40& '64    '支持新建文件夹功能
Private Const MAX_PATH = 260

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSelectION = (WM_USER + 102)

Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private m_CurrentDirectory As String   'The current directory

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Function BrowseForFolder(ByRef Owner As Form, _
                                Optional ByVal StartDir As String = "", _
                                Optional ByVal Title As String = "请选择一个文件夹：") As String
    '打开浏览目录窗口，并返回文件夹路径
    Dim lpIDList As Long
    Dim szTitle As String
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    
    m_CurrentDirectory = StartDir & vbNullChar

    szTitle = Title
    With tBrowseInfo
        .hWndOwner = Owner.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT + BIF_RETURNFSANCESTORS _
                 + BIF_EDITBOX + BIF_VALIDATE + BIF_NEWDIALOGSTYLE  '=1+2+4+8+16+32+64=112
        .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)  'get address of function.
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseForFolder = sBuffer
    Else
        BrowseForFolder = ""
    End If
  
End Function

Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    Dim lpIDList As Long
    Dim Ret As Long
    Dim sBuffer As String
    
    On Error Resume Next
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            Call SendMessage(hwnd, BFFM_SETSelectION, 1, m_CurrentDirectory)
        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH)
            Ret = SHGetPathFromIDList(lp, sBuffer)
            If Ret = 1 Then
                Call SendMessage(hwnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
            End If
        End Select
    BrowseCallbackProc = 0
End Function

Private Function GetAddressofFunction(ByVal AddOf As Long) As Long
    GetAddressofFunction = AddOf
End Function

Public Function gfAsciiAdd(ByVal strIn As String) As String
    '返回传入字符的Ascii码值加N后 对应的字符。
    '与gAsciiSub过程互逆
    '注意1：暂时设定支持字母和数字。
    '注意2：输入的字符对应的ASCII值不能超过122即小写字母z。
    '注意3：字符增量N值大于0且不能超过5。
    
    Dim intASC As Integer
    
    If Len(strIn) = 0 Then Exit Function
    
    If gconAscAdd > 5 Or gconAscAdd = 0 Then
        MsgBox "字符增量大于0且不能超过5！", vbExclamation, "字符转化增量警告"
        Exit Function
    End If
    
    intASC = Asc(Left(strIn, 1))
    Select Case intASC
        Case 48 To 57, 65 To 90, 97 To 122
            
            intASC = intASC + gconAscAdd
            Select Case intASC
                Case 48 To 57, 65 To 90, 97 To 122
                    '在些区间表示正常转化
                Case 58 To 64
                    intASC = intASC + 7     '7= - 57 + 64
                Case 91 To 96
                    intASC = intASC + 6     '6= - 90 + 96
                Case 123 To 127
                    intASC = intASC - 75    '-75= - 122 + 47
            End Select
            gfAsciiAdd = Chr(intASC)
            
        Case Else
            MsgBox "非法字符转化【" & strIn & "】！" & vbCrLf & "暂不支持数字和字母以外的字符！", vbExclamation, "不支持字符警告"
    End Select
    
End Function

Public Function gfAsciiSub(ByVal strIn As String) As String
    '返回传入字符的Ascii码值减N后 对应的字符。
    '与gAsciiAdd过程互逆
    '注意1：暂时设定只支持字母和数字。
    '注意2：输入的字符对应的ASCII值不能超过127。
    '注意3：字符增量N大于0且不能超过5。
    
    Dim intSub As Integer
    
    If Len(strIn) = 0 Then Exit Function
    
    If gconAscAdd > 5 Or gconAscAdd = 0 Then
        MsgBox "字符增量大于0且不能超过5！", vbExclamation, "字符转化增量警告"
        Exit Function
    End If
    
    intSub = Asc(Left(strIn, 1))
    Select Case intSub
        Case 48 To 57, 65 To 90, 97 To 122
            
            intSub = intSub - gconAscAdd
            Select Case intSub
                Case 48 To 57, 65 To 90, 97 To 122
                    '在些区间表示正常转化
                Case 43 To 47
                    intSub = intSub + 75    '=122-(47-intSub)
                Case 58 To 64
                    intSub = intSub - 7     '=57-(64-intSub)
                Case 91 To 96
                    intSub = intSub - 6     '=90-(96-intSub)
            End Select
            gfAsciiSub = Chr(intSub)
            
        Case Else
            MsgBox "非法字符转化【" & strIn & "】！" & vbCrLf & "暂不支持数字和字母以外的字符！", vbExclamation, "不支持字符警告"
    End Select
    
End Function

Public Function gfBackComputerInfo(Optional ByVal cType As genumComputerInfoType = ciComputerName, _
        Optional ByVal UseDefault As Boolean = True, Optional ByVal DefaultValue As String = "Null") As String
    '返回指定的电脑上的信息
    
    Dim strBack As String, strBuffer As String * 255
    
    If cType = ciComputerName Then  '计算机名称
        strBack = VBA.Environ("ComputerName")   '直接VBA函数获取
        If Len(strBack) = 0 Then
            Call GetComputerName(strBuffer, 255) '若获取失败则用API函数再获取一次
            strBack = strBuffer
        End If
    ElseIf cType = ciUserName Then  '计算机当前用户名
        strBack = VBA.Environ("UserName")
        If Len(strBack) = 0 Then
            Call GetUserName(strBuffer, 255)
            strBack = strBuffer
        End If
    End If
    
    If Len(strBack) = 0 Then  '如果为空时是否使用默认值
        If UseDefault Then strBack = DefaultValue
    End If
    gfBackComputerInfo = strBack
    
End Function


Public Function gfBackConnection(ByVal strCon As String, _
        Optional ByVal CursorLocation As CursorLocationEnum = adUseClient) As ADODB.Connection
    '返回数据库连接
       
    On Error GoTo LineERR
    
    Set gfBackConnection = New ADODB.Connection
    gfBackConnection.CursorLocation = CursorLocation
    gfBackConnection.ConnectionString = gVar.ConString
    gfBackConnection.CommandTimeout = 5
    gfBackConnection.Open
    
    Exit Function
    
LineERR:
    Call gsAlarmAndLog("数据库连接异常")
    
End Function


Public Function gfBackRecordset(ByVal cnSQL As String, _
                Optional ByVal cnCursorType As CursorTypeEnum = adOpenStatic, _
                Optional ByVal cnLockType As LockTypeEnum = adLockReadOnly, _
                Optional ByVal CursorLocation As CursorLocationEnum = adUseClient) As ADODB.Recordset
    '返回指定SQL查询语句的记录集
    
    Dim cnBack As ADODB.Connection
    
    On Error GoTo LineERR

    Set gfBackRecordset = New ADODB.Recordset
    Set cnBack = gfBackConnection(gVar.ConString, CursorLocation)
    If cnBack.State = adStateClosed Then Exit Function
    gfBackRecordset.CursorLocation = CursorLocation
    gfBackRecordset.Open cnSQL, cnBack, cnCursorType, cnLockType
    
    Exit Function

LineERR:
    Call gsAlarmAndLog("返回记录集异常")

End Function


Public Function gfBackLogType(Optional ByVal strType As genumLogType = udSelect) As String
    '返回日志操作类型
    Select Case strType
        Case udDelete
            gfBackLogType = "Delete"
        Case udDeleteBatch
            gfBackLogType = "DeleteBatch"
        Case udInsert
            gfBackLogType = "Insert"
        Case udInsertBatch
            gfBackLogType = "InsertBatch"
        Case udSelectBatch
            gfBackLogType = "SelectBatch"
        Case udUpdate
            gfBackLogType = "Update"
        Case udUpdateBatch
            gfBackLogType = "UpdateBatch"
        Case Else
            gfBackLogType = "Select"
    End Select
End Function


Public Function gfBackOneChar(Optional ByVal CharType As genumCharType = udUpperLowerNum) As String
    '随机返回一个字符（字母或数字）
    '48-57:0-9
    '65-90:A-Z
    '97-122:a-z
    
    Dim intRd  As Integer

    If (CharType > udUpperLowerNum) Or (CharType < udLowerCase) Then CharType = udUpperLowerNum
    
    Randomize
    Do
        intRd = CInt((74 * Rnd) + 48)
        If (CharType Or udNumber) = CharType Then
            If (intRd > 47 And intRd < 58) Then Exit Do
        End If
        If (CharType Or udUpperCase) = CharType Then
            If (intRd > 64 And intRd < 91) Then Exit Do
        End If
        If (CharType Or udLowerCase) = CharType Then
            If (intRd > 96 And intRd < 123) Then Exit Do
        End If
    Loop
    
    gfBackOneChar = Chr(intRd)
    
End Function


Public Function DecryptStringSimple(ByVal strIn As String) As String
    '解密输入的字符串密文为明文
    '密文长度限定为gconSumLen位
    
    Dim strVar As String    '中间变量
    Dim strPt As String     '明文
    Dim strMid As String    '截取输入字符串中的每一个字符
    Dim intMid As Integer, K As Integer, C As Integer, R As Integer   '变量
    
    strIn = Trim(strIn) '去空格
    C = Len(strIn)
    If C <> gconSumLen Then GoTo LineBreak
    
    '一、获取密文中填充的无用字符个数、明文的长度
    R = Val(Mid(strIn, 2, 1))       '截取密文的第二位，其值即密文第gconAddLenStart+1位后填充的无用随机数个数
    If R < 1 Then GoTo LineBreak
    
    intMid = Val(Left(strIn, 1))    '截取密文的第一位，计算填充字符个数的值的 个位上的数字
    C = IIf(intMid < (gconAddLenStart - 2), intMid, gconAddLenStart - 2)  '通过第一位的数值计算出填充数值的十位上的数字所在位置
    K = Val(Mid(strIn, C + 2 + 1, 1))   '截取填充数值的十位上的数字
    C = Val(CStr(K) & CStr(intMid))     '得出真正的 填充字符 总数值
    If (C < (gconSumLen - gconMaxPWD)) Or (C > (gconSumLen - 1)) Then GoTo LineBreak
    
    C = gconSumLen - C  '得出明文的长度
    C = C * 2           '因为明文中插入了相同个数的随机字符
    
    '二、删除加在密文前面的gconAddLenStart+ 1 + R 个字符 和 加在密文最后的字符
    strVar = Mid(strIn, gconAddLenStart + 1 + R + 1, C)
    If Len(strVar) <> C Then GoTo LineBreak
    
    '三、解密剩下的strVar字符
    For K = 1 To C Step 2
        strPt = strPt & gfAsciiSub(Mid(strVar, K, 1))
    Next
    If Len(strPt) <> C / 2 Then GoTo LineBreak
    
    DecryptStringSimple = strPt  '将解密好的密文返回给函数的调用者
    
    Exit Function
    
LineBreak:
'    Err.Clear
'    Err.Number = vbObjectError + 100001
'    Err.Description = "密文[" & strIn & "]被破坏，无法解密！"
'    Call gsAlarmAndLog("密文警告", False)
    Call gsAlarmAndLogEx("密文[" & strIn & "]被破坏，无法解密！", "密文警告", False)
End Function

Public Function EncryptStringSimple(ByVal strIn As String) As String
    '将传入的字符串(明文)进行简单加密，生成密文并返回给调用者
    '明文长度<=20个字符，且只能是大写或小写字母、数字，否则转化时会报错
    
    Dim strEt As String     '密文
    Dim strMid As String    '截取输入字符串中的每一个字符
    Dim strTen As String    '密文的前10个字符
    Dim K As Integer, J As Integer, R As Integer  '变量
    Dim C As Integer        '明文的字符个数
    Dim intFill As Integer  '填充字符数
    Dim intRightNum As Integer      'strFill 个位上的数字
    Dim intAddLenEnd As Integer     '加在最后的字符数量

    C = Len(Trim(strIn))
    If C = 0 Then
        MsgBox "传入字符不能为空字符，且不能有空格！", vbCritical, "空字符警报"
        Exit Function
    End If
    strIn = Left(strIn, gconMaxPWD) '截取前gconMaxPWD(20)字符
    C = Len(strIn)  '重新获取字符个数。重要！
    
    '一、将字符串中的每个字符的ASCII值前进N位并插入一个随机字符得到一新字符串
    For K = 1 To C
        strEt = strEt & gfAsciiAdd(Mid(strIn, K, 1)) & gfBackOneChar(udUpperLowerNum)
    Next
    If Len(strEt) <> (C * 2) Then
        MsgBox "输入字符不规范，只能是数字或字母！", vbCritical, "字符警报"
        Exit Function
    End If
    
    '二、在转化后的字符串strEt前面总是加入gconAddLenStart个字符
    '   在这gconAddLenStart个字符中包含明文的长度信息gconSumLen-C
    '   然后将gconSumLen-C的值的 个位与十位调换位置
    '   然后在strTen的第二位插入原strTen后应填充的随机数字个数
    intFill = gconSumLen - C        '计算去除明文个数后要填充的总字符个数
    intRightNum = intFill Mod 10    '获取个位上的数字
    strTen = CStr(intRightNum)      '将个位上的数字放在strTen的第一位,也即密文的第一位
    
    '根据strTen的第一位的值计算在其后插入的随机数字的个数
    J = IIf(intRightNum < (gconAddLenStart - 2), intRightNum, gconAddLenStart - 2)
    For K = 1 To J
        strTen = strTen & gfBackOneChar(udNumber)
    Next
    strTen = strTen & CStr(Int(intFill / 10))   '并上intFill的十位上的数字
    
    Do
        R = gfBackOneChar(udNumber)     '获取一个1~9中的随机数字
        If R > 0 Then Exit Do
    Loop
    strTen = Left(strTen, 1) & CStr(R) & Right(strTen, Len(strTen) - 1)
    
    '若strTen的长度不够gconAddLenStart位，则填充随机数字,再在strTen后面并上随机R个数字
    J = (gconAddLenStart - 2 - J) + R
    For K = 1 To J
        strTen = strTen & gfBackOneChar(udNumber)
    Next
    strEt = strTen & strEt
    
    '三、在strEt后追加intAddLenEnd个随机字符凑成gconSumLen个字符的最终密文
    intAddLenEnd = gconSumLen - (C * 2) - gconAddLenStart - R - 1
    If intAddLenEnd > 0 Then
        For K = 1 To intAddLenEnd
            strEt = strEt & gfBackOneChar(udUpperLowerNum)
        Next
    End If
    
    EncryptStringSimple = strEt  '最后将strEt赋给函数的返回值
    
End Function

Public Function gfFileCopy(ByVal strOld As String, ByVal strNew As String, Optional ByVal blnDelOld As Boolean = False) As Boolean
    '复制文件
    
    On Error GoTo LineERR
    
    FileCopy strOld, strNew
    gfFileCopy = True
    If blnDelOld Then
        Kill strOld
    End If
    Exit Function
LineERR:
    Call gsAlarmAndLog("文件复制异常")
End Function


Public Function gfFileExist(ByVal strPath As String) As Boolean
    '判断文件、文件目录 是否存在

    Dim strBack As String
        
    On Error GoTo LineERR
    
    If Len(strPath) > 0 Then    '空字符串不算
        strBack = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem)
        If Len(strBack) > 0 Then gfFileExist = True
    End If
  
    Exit Function
    
LineERR:
    Call gsAlarmAndLog("判断文件异常")
    
End Function


Public Function gfFileExistEx(ByVal strPath As String) As gtypeValueAndErr
    '另一种返回值方式：来判断文件、文件目录 是否存在
    '专供后面的过程gfFileRepair调用
    
    Dim strBack As String
    
    On Error GoTo LineERR
    
    If Len(strPath) > 0 Then    '空字符串不算
        strBack = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem)
        If Len(strBack) > 0 Then
            gfFileExistEx.Result = True
        Else
            gfFileExistEx.ErrNum = -1   '不存在，也没异常
        End If
    End If
    
    Exit Function
    
LineERR:
    gfFileExistEx.ErrNum = Err.Number   '异常了，也当作不存在了
    Call gsAlarmAndLog("文件判断返回异常")
    
End Function

Public Function gfFileIsRun(ByVal pFile As String) As Boolean
    '判断文件是否被打开(在运行)
    Dim Ret As Long
    
    Ret = CreateFile(pFile, GENERIC_READ Or GENERIC_WRITE, 0&, vbNullString, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    gfFileIsRun = (Ret = INVALID_HANDLE_VALUE)
    CloseHandle Ret
    '经小部分测试，似乎没用，只能判断可执行文件？
End Function


Public Function gfFileOpen(ByVal strFilePath As String) As gtypeValueAndErr
    '打开指定全路径的文件
    
    Dim lngRet As Long
    Dim strDir As String
    
    On Error GoTo LineERR
    
    If gfFileExist(strFilePath) Then
        
        lngRet = ShellExecute(GetDesktopWindow, "open", strFilePath, vbNullString, vbNullString, vbNormalFocus)
        If lngRet = SE_ERR_NOASSOC Then     '没有关联的程序
             strDir = Space(260)
             lngRet = GetSystemDirectory(strDir, Len(strDir))
             strDir = Left(strDir, lngRet)
             
            '显示打开方式窗口
            Call ShellExecute(GetDesktopWindow, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & strFilePath, strDir, vbNormalFocus)
            gfFileOpen.ErrNum = -1   '不成功，也没异常
        Else
            gfFileOpen.Result = True
        End If
        
    End If
    
    Exit Function
    
LineERR:
    gfFileOpen.ErrNum = Err.Number
    Call gsAlarmAndLog("文件打开异常")
    
End Function

Public Function gfFileRename(ByVal strOld As String, ByVal strNew As String) As Boolean
    '重命名文件或文件名
    
    On Error GoTo LineERR
    
    Close
    Name strOld As strNew
    Close
    gfFileRename = True
    Exit Function
LineERR:
    Close
    Call gsAlarmAndLog("文件/文件夹重命名异常", False)
End Function


Public Function gfFileReNameEx(ByVal strOld As String, ByVal strNew As String) As Boolean
    '重命名文件或文件名。先删除存在的新文件名的文件
    
    On Error GoTo LineERR
    
    If gfFileExist(strNew) Then
        Kill strNew '新文件存在则先删除
    End If
    
    Name strOld As strNew
    gfFileReNameEx = True
    
    Exit Function
LineERR:
    Call gsAlarmAndLog("文件/文件夹重命名异常", False)
End Function


Public Function gfFileRepair(ByVal strFile As String, Optional ByVal blnFolder As Boolean) As Boolean
    '如果 文件/文件夹 不存在 则创建
    '前提是路径的上层目录可访问
    '参数blnFolder指明传入的路径strFile是文件夹则为True，默认是文件False
    
    Dim strTemp As String
    Dim typBack As gtypeValueAndErr
    Dim lngLoc As Long
    
    If Right(strFile, 1) = "\" Then
        strFile = Left(strFile, Len(strFile) - 1)   '去掉最末的"\"
    End If
    strTemp = strFile
    If Len(strTemp) = 0 Then Exit Function          '防止传入空字符串
    
    On Error GoTo LineERR

    typBack = gfFileExistEx(strTemp)    '判断是否存在
    If Not typBack.Result Then          '文件不存在
        If typBack.ErrNum = -1 Then     '且无异常
            
            lngLoc = InStrRev(strTemp, "\") '判断是否有上层目录
            If lngLoc > 0 Then              '有上层目录则递归
                strTemp = Left(strTemp, lngLoc - 1) '得出上层目录的具体路径
                Call gfFileRepair(strTemp, True)    '递归调用自身，以保证上层目录存在
            End If

            If blnFolder Then                   '传入参数是文件夹
                MkDir strFile                   '则创建文件夹
            Else                                '传入参数是文件
                Close                           '则创建文件
                Open strFile For Random As #1
                Close
            End If
            
            gfFileRepair = True '创建成功返回True
            
        End If
        
    Else
        gfFileRepair = True '路径完整直接返回True
    End If

LineERR:
    Close
End Function

Public Function gfFolderRepair(ByVal strFile As String) As Boolean
    '如果 文件夹 不存在 则创建
    '前提是路径的上层目录可访问
    
    Dim strTemp As String, strDir As String
    Dim fsObject As Scripting.FileSystemObject
    Dim lngLoc As Long
    
    On Error GoTo LineERR
    
    strTemp = Trim(strFile)
    If Len(strTemp) = 0 Then GoTo LineERR   '防止传入空字符串
    
    Set fsObject = New Scripting.FileSystemObject   '实例化文件对象
    If fsObject.FolderExists(strTemp) Then    '判断文件夹是否存在
        gfFolderRepair = True '存在直接返回True
    Else    '文件夹不存在
        lngLoc = InStrRev(strTemp, "\") '判断是否有上层目录。目前不处理\\192.168.2.2这种路径
        If lngLoc > 0 Then              '有上层目录则递归
            strDir = Left(strTemp, lngLoc - 1) '得出上层目录的具体路径
            Call gfFolderRepair(strDir)        '递归调用自身，以保证上层目录存在
        End If
        fsObject.CreateFolder (strTemp) '上层目录确保存在后则创建该文件夹
        gfFolderRepair = True           '创建成功同时返回True
    End If
LineERR:
    Set fsObject = Nothing
    If Err.Number > 0 Then
        Call gsAlarmAndLog("文件夹路径[" & strTemp & "]异常！", False)
        Err.Clear
    End If
End Function


Public Function gfFormLoad(ByVal strFormName As String) As Boolean
    '判断指定窗口是否被加载了
    
    Dim frmLoad As Form
    
    strFormName = LCase(strFormName)
    For Each frmLoad In Forms
        If LCase(frmLoad.Name) = strFormName Then
            gfFormLoad = True
            Exit Function
        End If
    Next
    
End Function

Public Function gfGetRegStringValue(ByVal AppName As String, ByVal Section As String, ByVal Key As String, _
        Optional ByVal Default As String = "abc", Optional ByVal BackDefault As Boolean = True) As String
    '使GetSetting函数返回的字符串值不为空
    Dim strGet As String
    
    strGet = GetSetting(AppName, Section, Key, Default)
    If BackDefault Then
        If Len(Trim(strGet)) = 0 Then strGet = Default    '当获取值为空字符时也返回默认值
    End If
    gfGetRegStringValue = strGet
    
End Function

Public Function gfGetRegNumericValue(ByVal AppName As String, ByVal Section As String, _
        ByVal Key As String, Optional ByVal inMinMax As Boolean = True, Optional ByVal Default As Long = 1, _
        Optional ByVal nMin As Long = 1, Optional ByVal nMax As Long = 10) As Long
    '使GetSetting函数返回整形数值,，但这个值不能超出最小与最大值，超出以最小值返回
    Dim lngGet As Long
    
    lngGet = Val(GetSetting(AppName, Section, Key, Default))
    If inMinMax Then
        If lngGet < nMin Or lngGet > nMax Then lngGet = Default
    End If
    gfGetRegNumericValue = lngGet
    
End Function

Public Function gfGetSetting(ByVal AppName As String, ByVal Section As String, ByVal Key As String, Optional ByVal strNO As String = "*&^%$#@!") As Boolean
    '判断注册项是否存在
    
    Dim strGet As String
    
    strGet = GetSetting(AppName, Section, Key, strNO)
    If strGet <> strNO Then gfGetSetting = True
End Function

Public Function gfLoadAuthority(ByRef frmCur As Form, ByRef ctlCur As Control) As Boolean
    '加载窗口中的控制权限
    
    Dim strUser As String, strForm As String, strCtlName As String
    
    strUser = LCase(gVar.UserLoginName)
    strForm = LCase(frmCur.Name)
    strCtlName = LCase(ctlCur.Name)
    
    If strUser = LCase(gVar.AccountAdmin) Or strUser = LCase(gVar.AccountSystem) Then Exit Function
    ctlCur.Enabled = False
    
    With gVar.rsURF
        If .State = adStateOpen Then
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If strForm = LCase(.Fields("FuncFormName")) Then
                        If strCtlName = LCase(.Fields("FuncName")) Then
                            ctlCur.Enabled = True
                            gfLoadAuthority = True
                            Exit Do
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
    
End Function

Public Function gfIsTreeViewChild(ByRef nodeDad As MSComctlLib.Node, ByVal strKey As String) As Boolean
    '判断传入Key值是不是自己的子结点
    
    Dim I As Long, C As Long
    Dim nodeSon As MSComctlLib.Node
    
    C = nodeDad.Children
    If C = 0 Then Exit Function

    For I = 1 To C
        If I = 1 Then
            Set nodeSon = nodeDad.Child
        Else
            Set nodeSon = nodeSon.Next
        End If

'Debug.Print nodeSon.Text & "--" & nodeSon.Key

        If nodeSon.Key = strKey Then
            gfIsTreeViewChild = True
            Exit Function
        End If
        If nodeSon.Children > 0 Then
            If gfIsTreeViewChild(nodeSon, strKey) Then
                gfIsTreeViewChild = True
                Exit Function
            End If
        End If
    Next

End Function


Public Function gfStringCheck(ByVal strIn As String) As String
    '''敏感字符检测
    
    Dim arrStr As Variant
    Dim I As Long
    
    arrStr = Array(";", "--", "'", "//", "/*", "*/", "select", "update", _
                   "delete", "insert", "alter", "drop", "create")
    strIn = LCase(strIn)
    For I = LBound(arrStr) To UBound(arrStr)
        If InStr(strIn, arrStr(I)) > 0 Then
            gfStringCheck = arrStr(I)
            Exit Function
        End If
    Next

End Function

Public Function ShowBackupTimeInfo(ByVal BKInterval As Long, ByVal BKDate As Date) As String
    '转化备份频率与备份时间
    Dim strShow As String, strNext As String, strTime As String
    
    strTime = Format(BKDate, "HH:mm:ss")
    Select Case BKInterval
        Case 0
            strShow = "无"
        Case 1
            strShow = Format(BKDate, "每天") & strTime
        Case 2
            strShow = "每周" & WeekdayName(Weekday(BKDate)) & strTime
        Case 3
            strShow = Format(BKDate, "每月d日") & strTime
        Case 4
            strShow = Format(BKDate, "每年M月d日") & strTime
        Case 5
            strShow = "每" & gVar.ParaBackupIntervalDays & "天" & strTime
        Case Else
            strShow = "未定义"
    End Select
    ShowBackupTimeInfo = strShow
End Function

Public Function ShowBackupNextTime(ByVal BKInterval As Long, ByVal BKDate As Date) As String
    '转化备份频率与备份时间
    Dim strShow As String
    Dim NowTime As Date, BackTime As Date, NextDay As Date, ThisYear As Date
    Dim NowWeek As Long, BKWeek As Long, modDay As Long
    Dim NowDay As Long, BKDay As Long, NowMonth As Long, BKMonth As Long
    
    NowTime = Time
    BackTime = Format(BKDate, "HH:mm:ss")
    NowWeek = Weekday(Date)
    BKWeek = Weekday(BKDate)
    NowDay = Day(Date)
    BKDay = Day(BKDate)
    NowMonth = Month(Date)
    BKMonth = Month(BKDate)
    ThisYear = CDate(Format(Date, "yyyy-") & Format(BKDate, "MM-dd"))
    
    Select Case BKInterval
        Case 1 To 5
            Select Case BKInterval
                Case 1  '每天
                    If NowTime <= BackTime Then
                        NextDay = Date
                    Else
                        NextDay = Date + 1
                    End If
                Case 2  '每周
                    If (NowWeek < BKWeek) Or (NowWeek = BKWeek And NowTime <= BackTime) Then
                        NextDay = Date + (BKWeek - NowWeek)
                    Else
                        NextDay = Date + (7 - NowWeek + BKWeek)
                    End If
                Case 3  '每月
                    If (NowDay < BKDay) Or (NowDay = BKDay And NowTime <= BackTime) Then
                        NextDay = Date + (BKDay - NowDay)
                    Else
                        NextDay = DateAdd("m", DateDiff("m", BKDate, DateAdd("m", 1, Date)), BKDate)
                    End If
                Case 4  '每年
                    If (Date < ThisYear) Or (Date = ThisYear And NowTime <= BackTime) Then
                        NextDay = ThisYear
                    Else
                        NextDay = DateAdd("yyyy", 1, ThisYear)
                    End If
                Case 5  '每N天
                    If Now < BKDate Then
                        NextDay = BKDate
                    Else
                        modDay = (Now - BKDate) Mod gVar.ParaBackupIntervalDays
                        If modDay = 0 And NowTime > BackTime Then
                            modDay = modDay + gVar.ParaBackupIntervalDays
                        End If
                        NextDay = Date + modDay
                    End If
            End Select
            strShow = Format(NextDay, "yyyy-MM-dd ") & Format(BKDate, "HH:mm:ss")
        Case Else
            strShow = "无"
    End Select
    ShowBackupNextTime = strShow
End Function

