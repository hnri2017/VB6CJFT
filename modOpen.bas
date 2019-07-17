Attribute VB_Name = "modOpen"
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

'
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
    Dim ret As Long
    Dim sBuffer As String
    
    On Error Resume Next
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            Call SendMessage(hwnd, BFFM_SETSelectION, 1, m_CurrentDirectory)
        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH)
            ret = SHGetPathFromIDList(lp, sBuffer)
            If ret = 1 Then
                Call SendMessage(hwnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
            End If
        End Select
    BrowseCallbackProc = 0
End Function

Private Function GetAddressofFunction(ByVal AddOf As Long) As Long
    GetAddressofFunction = AddOf
End Function

Public Function DirFile(ByVal strPath As String) As Boolean
    '判断文件是否存在
    Dim strDir As String, strMid As String
    
    On Error Resume Next
    strDir = Dir(strPath)
    If Len(strDir) > 0 Then
        strMid = Mid(strPath, InStrRev(strPath, "\") + 1)
        If LCase(strMid) = LCase(strDir) Then
            DirFile = True
        End If
    End If
    If Err.Number Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
    End If
End Function

Public Function DirFolder(ByVal strPath As String) As Boolean
    '判断文件夹是否存在
    Dim strFod As String, strGet As String
    
    On Error Resume Next
    If Len(Trim(strPath)) > 0 Then  '以防传入的空字符串路径
        If Right(strPath, 1) = "\" Then
            If InStr(strPath, "\") <> InStrRev(strPath, "\") Then   '以防传入的是根目录
                strPath = Left(strPath, Len(strPath) - 1) '踢除末尾多余的"\"
            End If
        End If
    End If
    strFod = Dir(strPath, vbDirectory)
    If Len(strFod) > 0 Then '说明有返回值
        If strFod <> "." And strFod <> ".." Then    '若是空路径则返回"."
            If InStr(strPath, "\") = InStrRev(strPath, "\") Then    '以防传入的是根目录如"D:\"
                strGet = strPath
            Else
                strGet = Left(strPath, Len(strPath) - Len(strFod)) & strFod '正常情况下strFod值+上层目录=strPath
            End If
            If GetAttr(strGet) And vbDirectory = vbDirectory Then   '如果是文件夹或者存在的根目录
                DirFolder = True
            End If
        End If
    End If
    If Err.Number Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
    End If
End Function

Public Sub EnabledControl(ByRef frmEN As Form, Optional ByVal blnEN As Boolean = True)
    Dim ctlEn As VB.Control
    On Error Resume Next
    For Each ctlEn In frmEN.Controls
        ctlEn.Enabled = blnEN
    Next
    Screen.MousePointer = IIf(blnEN, 0, 13)
End Sub

Public Function BackupFile(ByVal strFolderSrc As String, ByVal strFolderDes As String, Optional ByVal blnSingleFile As Boolean = False) As Boolean
    '备份
    Dim strFS As String, strFD As String
    Dim strFBK As String, strFind As String, strDir As String, strGet As String, strPre As String
    Dim bytFile() As Byte, intSrc As Integer, intDes As Integer, lngSize As Long
    Dim strFR As String, intFR As Integer
    
    If Not (DirFolder(strFolderSrc) And DirFolder(strFolderDes)) Then
        MsgBox "备份的源路径或目的路径不存在", vbCritical, "警告"
        Exit Function
    End If
    
    On Error GoTo LineErr
    strFD = IIf(Right(strFolderDes, 1) = "\", strFolderDes, strFolderDes & "\")
    strFS = IIf(Right(strFolderSrc, 1) = "\", strFolderSrc, strFolderSrc & "\")
    strPre = "fbk" & Format(Now, "yyyy-MM-dd-HH-mm-ss")
    strFBK = strFD & strPre & ".bak"
    strFind = strFS & "*.*"
    
    strFR = strFD & strPre & ".fst"
    intFR = FreeFile
    Open strFR For Output As #intFR
    
    intDes = FreeFile
    Open strFBK For Binary As #intDes
    strDir = Dir(strFind)
    Do While Not Len(strDir) = 0
        DoEvents    '单个文件太大时容易内存溢出报错,大约三四百M的时候?
        intSrc = FreeFile
        strGet = strFS & strDir
        lngSize = FileLen(strGet)
        ReDim bytFile(lngSize - 1)
        Open strGet For Binary As #intSrc
        Get #intSrc, , bytFile
        Put #intDes, , bytFile
        Close intSrc
        Print #intFR, strDir & vbTab & CStr(lngSize)
        strDir = Dir
    Loop
    Close intSrc
    Close intDes
    Close intFR
    BackupFile = True
    
LineErr:
    Close   '关闭所有
    If Err.Number Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
    End If
End Function

Public Function RestoreFile(ByVal strFileSrc As String, ByVal strFolderDes As String) As Boolean
    '还原
    Dim strFS As String, strFD As String
    Dim strFI As String, strLine As String, strArr() As String, strFBK As String
    Dim bytFile() As Byte, intFI As Integer, intSrc As Integer, intDes As Integer, lngSize As Long

    strFI = Left(strFileSrc, InStrRev(strFileSrc, ".")) & "fst" '恢复文件的配置文件
    If Not (DirFolder(strFolderDes) And DirFile(strFileSrc) And DirFile(strFI)) Then
        MsgBox "还原的源文件或还原位置不存在", vbCritical, "警告"
        Exit Function
    End If
    
    On Error GoTo LineErr
    strFS = Left(strFileSrc, InStrRev(strFileSrc, "\"))
    strFD = IIf(Right(strFolderDes, 1) = "\", strFolderDes, strFolderDes & "\")
    
    intFI = FreeFile
    Open strFI For Input As #intFI
    intSrc = FreeFile
    Open strFileSrc For Binary As #intSrc
    While Not EOF(intFI)
        Line Input #intFI, strLine
        strArr = Split(strLine, vbTab)
        If UBound(strArr) <> 1 Then GoTo LineErr
        If Not IsNumeric(strArr(1)) Then GoTo LineErr
        strFBK = strFD & strArr(0)
        ReDim bytFile(strArr(1) - 1)
        Get #intSrc, , bytFile
        intDes = FreeFile
        Open strFBK For Binary As #intDes
        Put #intDes, , bytFile
        Close intDes
    Wend
    Close intFI
    Close intSrc
    Close intDes
    RestoreFile = True
    
LineErr:
    Close   '关闭所有
    If Err.Number Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
    End If
End Function
