Attribute VB_Name = "modFunction"
Option Explicit



'---------------------------------------------------------------------------------------
'MsgBox�Զ��˳����API�����
Private Declare Function SetWindowTextA Lib "user32" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
Private Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long
Private Const TIME_PERIODIC As Long = 1 'program for continuous periodic event
Private Const TIME_ONESHOT As Long = 0  'program timer for single event
Private Const DelayTime As Long = 500   'API����timeSetEvent�ļ�ʱ�����ʱ�����

Rem Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const WM_GETTEXT As Long = &HD&
Private Const WM_SETTEXT As Long = &HC&
Private Const WM_CLOSE As Long = &H10&

Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Private TimeID As Long      '���ض�ý���ʱ�������ʶ
Private Dlghwnd As Long     '�Ի�����
Private Dlgtexthwnd As Long '�Ի�����ʾ�ı����
Private MediaCount As Double    '����ʱ���ۼ���
Private MsgBoxCloseTime As Long     '���öԻ���ر�ʱ��
Private MsgBoxPromptText As String  '���öԻ�����ʾ�ı�
Private MsgBoxTitleText As String   '���öԻ��򴰿ڱ����ı�

Private Declare Function MessageBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hwnd As Long, _
    ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, _
    ByVal wlange As Long, ByVal dwTimeout As Long) As Long
'---------------------------------------------------------------------------------------

'��ȡһ�����жϽ��̵��˳�����.�����ʾ�ɹ������ʾʧ��
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
'��һ���Ѵ��ڵĽ��̶��󣬲����ؽ��̵ľ��
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'�ر�һ���ں˶������а����ļ����ļ�ӳ�䡢���̡��̡߳���ȫ��ͬ�������
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const CREATE_NEW_CONSOLE = &H10
Private Const Process_query_infomation = &H400  '��ȡ���̵����ơ��˳�������ȼ�����Ϣ
Private Const Still_Active = &H103

'�õ���ǰƽ̨�Ͳ���ϵͳ�йصİ汾��Ϣ.�����ʾ�ɹ������ʾʧ��
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long '��ʼ��Ϊ�ṹ�Ĵ�С
    dwMajorVersion As Long      'ϵͳ���汾��
    dwMinorVersion As Long      'ϵͳ�ΰ汾��
    dwBuildNumber As Long       'ϵͳ������
    dwPlatformId As Long        'ϵͳ֧�ֵ�ƽ̨
    szCSDVersion As String * 128    '
End Type

'�õ���ǰwindowϵͳĿ������·����
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'---------------------------------------------------------------------------------------

'���Ŀ¼��ʹ�õĳ�����API��Type�������ȡ����ܣ���λ����ǰ�ļ��У�����ѡ����
Private Const BIF_RETURNONLYFSDIRS = 1  '���������ļ�ϵͳ��Ŀ¼
Private Const BIF_DONTGOBELOWDOMAIN = 2 '�������Ӵ��У��������������µ�����Ŀ¼�ṹ
Private Const BIF_STATUSTEXT = &H4&     '�ڶԻ����а���һ��״̬����
Private Const BIF_RETURNFSANCESTORS = 8 '�����ļ�ϵͳ��һ���ڵ�
Private Const BIF_EDITBOX = &H10& ' 16  '����Ի����а���һ���༭��
Private Const BIF_VALIDATE = &H20& '32  '��û��BIF_EDITBOX��־λʱ���ñ�־λ������
Private Const BIF_NEWDIALOGSTYLE = &H40& '64    '֧���½��ļ��й���
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
'-----------------------------------------------------------------------------------
'---------------------------------------------------------------------------

Private Const mconStrKey As String = "ftkey"        '������Կ
Private Const mconStrBKbak As String = ".bak"       '�����ļ�����չ��
Private Const mconStrBKrst As String = ".rst"       '���������ļ���չ��
Private Const mconStrRar As String = ".rar"         'ѹ���ļ���չ��
Private Const mconLngSizeCompress As Long = 100     'ѹ���ļ��־��С����λMB

'---------------------------------------------------------------------------


Public Function BrowseForFolder(ByRef Owner As Form, _
                                Optional ByVal StartDir As String = "", _
                                Optional ByVal Title As String = "��ѡ��һ���ļ��У�") As String
    '�����Ŀ¼���ڣ��������ļ���·��
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

'--------------------------------------------------------------------------

Public Function DriveFreeSpaceMB(ByVal strPath As String) As Long
    '���ش���ʣ����ÿռ�,��λMB
    Dim strDir As String
    Dim objFSO As Object, objDrv As Object
    
    On Error Resume Next
    
    strDir = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem + vbVolume)
    If Len(strDir) > 0 Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objDrv = objFSO.GetDrive(objFSO.GetDriveName(strPath))
        DriveFreeSpaceMB = objDrv.FreeSpace / 1024 / 1024
    End If
    
    If Err.Number Then
        Call gsAlarmAndLog(Err.Number & "--" & Err.Description, False)
        DriveFreeSpaceMB = -1
    End If
    Set objDrv = Nothing
    Set objFSO = Nothing
End Function

Public Function DriveLetter(ByVal strPath As String) As String
    '���ش��̵ķ�����ĸ
    Dim strDir As String
    Dim objFSO As Object, objDrv As Object
    
    On Error Resume Next
    
    strDir = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem + vbVolume)
    If Len(strDir) > 0 Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objDrv = objFSO.GetDrive(objFSO.GetDriveName(strPath))
        DriveLetter = objDrv.DriveLetter
    End If
    
    If Err.Number Then
        Call gsAlarmAndLog(Err.Number & "--" & Err.Description, False)
    End If
    Set objDrv = Nothing
    Set objFSO = Nothing
End Function

Public Function DriveTotalSizeMB(ByVal strPath As String) As Long
    '���ش����ܿռ��С,��λMB
    Dim strDir As String
    Dim objFSO As Object, objDrv As Object
    
    On Error Resume Next
    
    strDir = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem + vbVolume)
    If Len(strDir) > 0 Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objDrv = objFSO.GetDrive(objFSO.GetDriveName(strPath))
        DriveTotalSizeMB = objDrv.TotalSize / 1024 / 1024
    End If
    
    If Err.Number Then
        Call gsAlarmAndLog(Err.Number & "--" & Err.Description, False)
        DriveTotalSizeMB = -1
    End If
    Set objDrv = Nothing
    Set objFSO = Nothing
End Function

Public Function DriveVolumeName(ByVal strPath As String) As String
    '���ش��̾����
    Dim strDir As String
    Dim objFSO As Object, objDrv As Object
    
    On Error Resume Next
    
    strDir = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem + vbVolume)
    If Len(strDir) > 0 Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objDrv = objFSO.GetDrive(objFSO.GetDriveName(strPath))
        DriveVolumeName = objDrv.VolumeName
    End If
    
    If Err.Number Then
        Call gsAlarmAndLog(Err.Number & "--" & Err.Description, False)
    End If
    Set objDrv = Nothing
    Set objFSO = Nothing
End Function

Public Sub EnabledControl(ByRef frmEN As Form, Optional ByVal blnEN As Boolean = True)
    Dim ctlEn As VB.Control
    On Error Resume Next
    For Each ctlEn In frmEN.Controls
        ctlEn.Enabled = blnEN
    Next
    Screen.MousePointer = IIf(blnEN, 0, 13)
End Sub
'--------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------

Public Function FileBackupCP(ByVal strSrcFolder As String, ByVal strDesFolder As String) As Boolean
    '�ļ�����,��ѹ���ٴ��
    Dim lngSizeSrc As Long, lngSizeDes As Long
    Dim strPathTemp As String, strSrc As String, strDes As String, strMsg As String
    
    If Not FolderExist(strSrcFolder) Then
        strMsg = "Ҫ���ݵ�Ŀ¼������"
        GoTo LineEnd
    End If
    If Not FolderPathBuild(strDesFolder) Then
        strMsg = "���ɵı����ļ��ı���λ�ò�����"
        GoTo LineEnd
    End If
    
    lngSizeSrc = FolderSizeMB(strSrcFolder)
    lngSizeDes = DriveFreeSpaceMB(strDesFolder)
    If lngSizeSrc = -1 Or lngSizeDes = -1 Then
        strMsg = "�ļ��д�С��ȡ�쳣"
        GoTo LineEnd
    End If
    If lngSizeDes < lngSizeSrc * 2 Then
        strMsg = "�����ļ�����λ�õĿռ䲻��"
        GoTo LineEnd
    End If
    
    strSrc = IIf(Right(strSrcFolder, 1) = "\", strSrcFolder, strSrcFolder & "\")
    strDes = IIf(Right(strDesFolder, 1) = "\", strDesFolder, strDesFolder & "\")
    strPathTemp = strDes & "Temp" & Format(Now, "yyyyMMddHHmmss")
    If FolderExist(strPathTemp, True) Then
        Call FolderDelete(strPathTemp)    'ɾ��Ҳ�������ʱ�ļ���
    End If
    If Not FolderPathBuild(strPathTemp) Then
        strMsg = "��ʱ�ļ����쳣"
        GoTo LineEnd
    End If
    
    If Not FileCompress(strSrc, strPathTemp, , True) Then
        strMsg = "�ļ�ѹ���쳣"
        GoTo LineEnd
    End If
    
    If FilePackage(strPathTemp, strDes) Then
        Call FolderDelete(strPathTemp)
    Else
        strMsg = "�ļ�����쳣"
        GoTo LineEnd
    End If
    
    FileBackupCP = True
LineEnd:
    If Len(strMsg) > 0 Then
        Call gsAlarmAndLog(strMsg, False)
        Call FolderDelete(strPathTemp)
    End If
End Function

Public Function FileCompress(ByVal strSrcFolder As String, ByVal strDesFolder As String, _
            Optional ByVal MSize As Long = mconLngSizeCompress, _
            Optional ByVal HideBack As Boolean = True) As Boolean
    'ѹ���ļ�
    Dim strWinRAR As String, strSrc As String, strDes As String
    Dim strSize As String, strHide As String, strCommand As String, strMsg As String
    
    strWinRAR = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "WinRAR.exe"
    If Not FileExist(strWinRAR) Then  'ѹ�������Ƿ����
        strMsg = "WinRARѹ��Ӧ�ó��򲻴���"
        GoTo LineEnd
    End If
    If Not FolderExist(strSrcFolder) Then 'Դ�ļ���Ŀ���ļ��Ƿ����
        strMsg = "��ѹ�����ļ�Ŀ¼������"
        GoTo LineEnd
    End If
    If Not FolderExist(strDesFolder) Then
        strMsg = "����ѹ���ļ���Ŀ¼������"
        GoTo LineEnd
    End If
    If FolderNotNull(strSrcFolder) = 0 Then 'ԴĿ¼�Ƿ�Ϊ��
        strMsg = "��ѹ�����ļ�Ŀ¼�޿�ѹ���ļ�"
        GoTo LineEnd
    End If
    
    strSrc = IIf(Right(strSrcFolder, 1) = "\", strSrcFolder, strSrcFolder & "\")    '��ʽ: D:\temp\��'\'����Ҫ����
    strDes = IIf(Right(strDesFolder, 1) = "\", strDesFolder, strDesFolder & "\") & "FC_" & Format(Now, "yyyy_MM_DD_HH_mm_ss") & mconStrRar
    If MSize < 0 Then MSize = mconLngSizeCompress
    If MSize <> 0 Then strSize = "-v" & MSize & "M" 'ָ����С�ķ־�ѹ��
    If HideBack Then strHide = " -ibck"            'ѹ�����ȴ�����С������������
    '����ѹ��shell���'-k�����ļ���-v50M ��50M�־�-r ��ͬ���ļ��У�-ep1 ·���в����������ļ���
    strCommand = strWinRAR & " a " & strSize & strHide & " -y -s -k -r -ep1 " & strDes & " " & strSrc
    If ShellWait(strCommand) Then
        FileCompress = True '�������ѹ�����̱��ж�ȡ��Ҳ�Ƿ���True��
    End If
LineEnd:
    If Len(strMsg) > 0 Then
        Call gsAlarmAndLog(strMsg, False)
    End If
End Function

Public Function FileExist(ByVal strPath As String, Optional ByVal blnSetNormal As Boolean = False) As Boolean
    '�ж� [�ļ�] �Ƿ����
    Dim strDir As String, strMid As String
    
    On Error Resume Next
    strDir = Dir(strPath, vbHidden) '�������򷵻ز���·�����ļ���
    If Len(strDir) > 0 Then
        strMid = Mid(strPath, InStrRev(strPath, "\") + 1)   '��ò������ļ���·�����ļ���
        If LCase(strMid) = LCase(strDir) Then   '������ʾ�ļ�����
            If blnSetNormal Then    '��Ҫǿ�Ƹı��ļ����ԣ���ɾ��ֻ��ȡ����������
                If GetAttr(strPath) <> vbNormal Then
                    SetAttr strPath, vbNormal   'ȥ����������ԣ�ǿ�Ƹĳɳ��������ļ�
                End If
            End If
            FileExist = True
        End If
    End If
    If Err.Number Then
        Call gsAlarmAndLog(Err.Number & "--" & Err.Description, False)
    End If
End Function

Public Function FileExtract(ByVal strSrcFile As String, ByVal strDesFolder As String, _
                    Optional ByVal HideBack As Boolean = True) As Boolean
    '��ѹ�ļ�
    
    Dim strWinRAR As String, strSrc As String, strDes As String
    Dim strHide As String, strCommand As String, strMsg As String
    
    strWinRAR = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "WinRAR.exe"
    If Not FileExist(strWinRAR) Then  'ѹ�������Ƿ����
        strMsg = "WinRARѹ��Ӧ�ó��򲻴���"
        GoTo LineEnd
    End If
    If Not FileExist(strSrcFile) Then 'Դ�ļ���Ŀ���ļ��Ƿ����
        strMsg = "����ѹ���ļ�������"
        GoTo LineEnd
    End If
    If Not FolderExist(strDesFolder) Then
        strMsg = "��ѹ����ļ����Ŀ¼������"
        GoTo LineEnd
    End If
        
    strSrc = strSrcFile
    strDes = IIf(Right(strDesFolder, 1) = "\", strDesFolder, strDesFolder & "\")
    
    '����ѹ��shell����
    If HideBack Then strHide = " -ibck "
    strCommand = strWinRAR & " x -y " & strHide & strSrc & " " & strDes '-y������ѯ���Զ���Ӧ��
    If ShellWait(strCommand) Then
        FileExtract = True
    End If
LineEnd:
    If Len(strMsg) > 0 Then
        Call gsAlarmAndLog(strMsg, False)
    End If
End Function

Public Function FilePackage(ByVal strFolderSrc As String, ByVal strFolderDes As String, _
    Optional ByVal blnEncrypt As Boolean = True, _
    Optional ByVal strKey As String = mconStrKey) As Boolean
    '��ɢ�ļ������һ���ļ�
    Dim strFS As String, strFD As String
    Dim strFBK As String, strFind As String, strDir As String, strGet As String, strPre As String
    Dim bytFile() As Byte, intSrc As Integer, intDes As Integer, lngSize As Long, bytEncrypt() As Byte
    Dim strFR As String, intFR As Integer, strSize As String, strMsg As String
    
    If Not (FolderExist(strFolderSrc) And FolderExist(strFolderDes)) Then
        strMsg = "Դ·����Ŀ��·��������"
        GoTo LineEnd
    End If
    If LCase(strFolderSrc) = LCase(strFolderDes) Then
        strMsg = "Դ·����Ŀ��·��������ͬ"
        GoTo LineEnd
    End If
    
    On Error GoTo LineErr
    strFD = IIf(Right(strFolderDes, 1) = "\", strFolderDes, strFolderDes & "\")
    strFS = IIf(Right(strFolderSrc, 1) = "\", strFolderSrc, strFolderSrc & "\")
    strPre = "fbk" & Format(Now, "yyyy-MM-dd-HH-mm-ss")
    strFBK = strFD & strPre & mconStrBKbak
    strFind = strFS & "*.*"
    
    strFR = strFD & strPre & mconStrBKrst
    intFR = FreeFile
    Open strFR For Output As #intFR
    
    intDes = FreeFile
    Open strFBK For Binary As #intDes
    strDir = Dir(strFind)
    Do While Not Len(strDir) = 0
        DoEvents
        intSrc = FreeFile
        strGet = strFS & strDir
        lngSize = FileLen(strGet)
        strSize = CStr(lngSize)
        ReDim bytFile(lngSize - 1)
        Open strGet For Binary As #intSrc
        Get #intSrc, , bytFile '������ʱ�����ļ���Լ����380MBʱ�ڴ��������
        If blnEncrypt Then
            Rem ReDim bytEncrypt(lngSize - 1)
            bytEncrypt = EncryptByte(bytFile, strKey)   '����ʱ�����ļ���Լ����185MBʱ�ڴ��������
            bytFile = bytEncrypt    '����ǰ��Byte�����ά�����С���
            strDir = EncryptString(strDir, strKey)
            strSize = EncryptString(strSize, strKey)
        End If
        Put #intDes, , bytFile
        Close intSrc
        Print #intFR, strDir & vbTab & strSize
        strDir = Dir
    Loop
    Close intSrc
    Close intDes
    Close intFR
    FilePackage = True
    
LineErr:
    Close   '�ر����д򿪵��ļ�
    If Err.Number Then
        Call gsAlarmAndLog(Err.Number & "--" & Err.Description, False)
        If FileExist(strFBK) Then Kill strFBK 'ɾ���ļ�
        If FileExist(strFR) Then Kill strFR
    End If
LineEnd:
    If Len(strMsg) > 0 Then
        Call gsAlarmAndLog(strMsg, False)
    End If
End Function

Public Function FileRestoreCP(ByVal strFileSrc As String, ByVal strFolderDes As String) As Boolean
    '��ԭ�ļ�
    Dim lngSizeSrc As Long, lngSizeDes As Long
    Dim strDes As String, strPathTemp As String, strPathFile As String, strMsg As String
    
    On Error Resume Next
    If Not FileExist(strFileSrc) Then
        strMsg = "��ԭ��Դ�ļ�������"
        GoTo LineEnd
    End If
    If Not FolderPathBuild(strFolderDes) Then
        strMsg = "�ļ���ԭλ�ò�����"
        GoTo LineEnd
    End If
    
    lngSizeSrc = FileLen(strFileSrc) / 1024 / 1024
    lngSizeDes = DriveFreeSpaceMB(strFolderDes)
    If lngSizeDes < lngSizeSrc * 2 Then
        strMsg = "��ԭλ�ÿռ䲻��"
        GoTo LineEnd
    End If
    
    strDes = IIf(Right(strFolderDes, 1) = "\", strFolderDes, strFolderDes & "\")
    strPathTemp = strDes & "Temp" & Format(Now, "yyyyMMddHHmmss")
    If FolderExist(strPathTemp) Then
        Call FolderDelete(strPathTemp)
    End If
    If Not FolderPathBuild(strPathTemp) Then
        strMsg = "��ʱ�ļ��д����쳣"
        GoTo LineEnd
    End If
    
    If Not FileUnpack(strFileSrc, strPathTemp) Then
        strMsg = "����ļ���ԭ�쳣"
        GoTo LineEnd
    End If
    
    strPathFile = Dir(strPathTemp & "\*" & mconStrRar)
    If InStr(strPathFile, mconStrRar) = 0 Then
        strMsg = "��ѹ�ļ������쳣"
        GoTo LineEnd
    Else
        strPathFile = strPathTemp & "\" & strPathFile
    End If

    If FileExtract(strPathFile, strDes) Then
        FileRestoreCP = True
        Call FolderDelete(strPathTemp)
    Else
        strMsg = "�ļ���ѹ�쳣"
        GoTo LineEnd
    End If
LineEnd:
    If Len(strMsg) > 0 Then
        Call gsAlarmAndLog(strMsg, False)
        Call FolderDelete(strPathTemp)
    End If
End Function

Public Function FileUnpack(ByVal strFileSrc As String, ByVal strFolderDes As String, _
    Optional ByVal blnDecrypt As Boolean = True, _
    Optional ByVal strKey As String = mconStrKey) As Boolean
    '����ð���Դ�ļ���ԭ����ɢ�ļ�
    Dim strFS As String, strFD As String, strMsg As String
    Dim strFI As String, strLine As String, strArr() As String, strFBK As String, bytDecrypt() As Byte
    Dim bytFile() As Byte, intFI As Integer, intSrc As Integer, intDes As Integer, lngSize As Long

    strFI = Left(strFileSrc, InStrRev(strFileSrc, ".") - 1) & mconStrBKrst '�ָ��ļ��������ļ�
    If Not (FolderExist(strFolderDes) And FileExist(strFileSrc) And FileExist(strFI)) Then
        strMsg = "��ԭ��Դ�ļ���ԭλ�ò�����"
        GoTo LineEnd
    End If
    If LCase(Mid(strFileSrc, InStrRev(strFileSrc, "."))) <> LCase(mconStrBKbak) Then
        strMsg = "��ԭ��Դ�ļ���ʽ����"
        GoTo LineEnd
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
        If blnDecrypt Then
            strArr(0) = DecryptString(strArr(0), strKey)
            strArr(1) = DecryptString(strArr(1), strKey)
        End If
        If Not IsNumeric(strArr(1)) Then GoTo LineErr
        strFBK = strFD & strArr(0)
        ReDim bytFile(strArr(1) - 1)
        Get #intSrc, , bytFile
        If blnDecrypt Then
            bytDecrypt = DecryptByte(bytFile, strKey)
            bytFile = bytDecrypt
        End If
        intDes = FreeFile
        Open strFBK For Binary As #intDes
        Put #intDes, , bytFile
        Close intDes
    Wend
    Close intFI
    Close intSrc
    Close intDes
    FileUnpack = True
    
LineErr:
    Close   '�ر����д򿪵��ļ�
    If Err.Number Then
        Call gsAlarmAndLog(Err.Number & "--" & Err.Description, False)
    End If
LineEnd:
    If Len(strMsg) > 0 Then
        Call gsAlarmAndLog(strMsg, False)
    End If
End Function

Public Function FolderDelete(ByVal strFolderPath As String) As Boolean
    'ɾ��ָ���ļ���
    Dim objFSO As Object, objFod As Object
    
    On Error Resume Next
    
    If FolderExist(strFolderPath) Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFod = objFSO.GetFolder(strFolderPath)
        objFod.Delete True
        DoEvents
        FolderDelete = True
    End If
    
    If Err.Number Then
        Call gsAlarmAndLog(Err.Number & "--" & Err.Description, False)
    End If
    Set objFod = Nothing
    Set objFSO = Nothing
End Function

Public Function FolderExist(ByVal strPath As String, Optional ByVal blnSetNormal As Boolean = False) As Boolean
    '�ж� [�ļ���] �Ƿ����
    Dim strFod As String, strGet As String
    
    On Error Resume Next
    If Len(Trim(strPath)) > 0 Then  '�Է�������ַ���·��
        If Right(strPath, 1) = "\" Then
            If InStr(strPath, "\") <> InStrRev(strPath, "\") Then   '�Է�������Ǹ�Ŀ¼
                strPath = Left(strPath, Len(strPath) - 1) '�Ǹ�Ŀ¼���߳�ĩβ�����"\"
            End If
        End If
    End If
    strFod = Dir(strPath, vbDirectory + vbHidden)
    If Len(strFod) > 0 Then '˵���з���ֵ
        If strFod <> "." And strFod <> ".." Then    '���ǿ�·���򷵻�"."
            If InStr(strPath, "\") = InStrRev(strPath, "\") Then    '�Է�������Ǹ�Ŀ¼��"D:\"
                strGet = strPath
            Else
                strGet = Left(strPath, Len(strPath) - Len(strFod)) & strFod '���������strFodֵ+�ϲ�Ŀ¼=strPath
            End If
            If GetAttr(strGet) And vbDirectory = vbDirectory Then   '������ļ��л��ߴ��ڵĸ�Ŀ¼
                If blnSetNormal Then    '��Ҫǿ�Ƹı��ļ����ԣ���ɾ��ֻ��ȡ����������
                    If GetAttr(strPath) <> vbNormal Then
                        SetAttr strPath, vbNormal   'ȥ����������ԣ�ǿ�Ƹĳɳ��������ļ�
                    End If
                End If
                FolderExist = True
            End If
        End If
    End If
    If Err.Number Then
        Call gsAlarmAndLog(Err.Number & "--" & Err.Description, False)
    End If
End Function

Public Function FolderNotNull(ByVal strFolderPath As String) As Boolean
    '����ļ����Ƿ�Ϊ��Ŀ¼
    Dim objFSO As Object, objFolder As Object, objFiles As Object
    
    On Error Resume Next
    
    If FolderExist(strFolderPath) Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFolder = objFSO.GetFolder(strFolderPath)
        Set objFiles = objFolder.Files
        If objFiles.Count > 0 Then  '�ļ�����
            FolderNotNull = True
        ElseIf objFolder.SubFolders.Count > 0 Then  '�ļ��и���
            FolderNotNull = True
        End If
    End If
    
    If Err.Number Then
        Call gsAlarmAndLog(Err.Number & "--" & Err.Description, False)
    End If
    Set objFiles = Nothing
    Set objFolder = Nothing
    Set objFSO = Nothing
End Function

Public Function FolderPathBuild(ByVal strFolderPath As String) As Boolean
    '�ļ���·��������ʱ�½�
    Dim strFod As String, strParentFolder As String, strNew As String
    
    On Error GoTo LineErr
    
    If FolderExist(strFolderPath) Then
        FolderPathBuild = True
    Else
        strFod = IIf(Right(strFolderPath, 1) = "\", Left(strFolderPath, Len(strFolderPath) - 1), strFolderPath)
        strParentFolder = Left(strFod, InStrRev(strFod, "\") - 1)   '��ȡ��һ���ļ���·��
        If InStr(strParentFolder, "\") = 0 Then
            strParentFolder = strParentFolder & "\" '��ֹ��Ŀ¼��C:�ں���FolderExist�з���False
        End If
        If FolderPathBuild(strParentFolder) Then    '�ݹ����
            MkDir strFod
            FolderPathBuild = True
        End If
    End If
    
LineErr:
    If Err.Number Then
        Call gsAlarmAndLog(Err.Number & "--" & Err.Description, False)
    End If
End Function

Public Function FolderSizeMB(ByVal strFolderPath As String) As Long
    '�����ļ��еĴ�С����λMB
    Dim objFSO As Object, objFod As Object
    
    On Error Resume Next
    
    If FolderExist(strFolderPath) Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFod = objFSO.GetFolder(strFolderPath)
        FolderSizeMB = objFod.Size / 1024 / 1024
    End If
    
    If Err.Number Then
        Call gsAlarmAndLog(Err.Number & "--" & Err.Description, False)
        FolderSizeMB = -1
    End If
    Set objFod = Nothing
    Set objFSO = Nothing
End Function



Public Function gfAsciiAdd(ByVal strIn As String) As String
    '���ش����ַ���Ascii��ֵ��N�� ��Ӧ���ַ���
    '��gAsciiSub���̻���
    'ע��1����ʱ�趨֧����ĸ�����֡�
    'ע��2��������ַ���Ӧ��ASCIIֵ���ܳ���122��Сд��ĸz��
    'ע��3���ַ�����Nֵ����0�Ҳ��ܳ���5��
    
    Dim intASC As Integer
    
    If Len(strIn) = 0 Then Exit Function
    
    If gconAscAdd > 5 Or gconAscAdd = 0 Then
        MsgBox "�ַ���������0�Ҳ��ܳ���5��", vbExclamation, "�ַ�ת����������"
        Exit Function
    End If
    
    intASC = Asc(Left(strIn, 1))
    Select Case intASC
        Case 48 To 57, 65 To 90, 97 To 122
            
            intASC = intASC + gconAscAdd
            Select Case intASC
                Case 48 To 57, 65 To 90, 97 To 122
                    '��Щ�����ʾ����ת��
                Case 58 To 64
                    intASC = intASC + 7     '7= - 57 + 64
                Case 91 To 96
                    intASC = intASC + 6     '6= - 90 + 96
                Case 123 To 127
                    intASC = intASC - 75    '-75= - 122 + 47
            End Select
            gfAsciiAdd = Chr(intASC)
            
        Case Else
            MsgBox "�Ƿ��ַ�ת����" & strIn & "����" & vbCrLf & "�ݲ�֧�����ֺ���ĸ������ַ���", vbExclamation, "��֧���ַ�����"
    End Select
    
End Function

Public Function gfAsciiSub(ByVal strIn As String) As String
    '���ش����ַ���Ascii��ֵ��N�� ��Ӧ���ַ���
    '��gAsciiAdd���̻���
    'ע��1����ʱ�趨ֻ֧����ĸ�����֡�
    'ע��2��������ַ���Ӧ��ASCIIֵ���ܳ���127��
    'ע��3���ַ�����N����0�Ҳ��ܳ���5��
    
    Dim intSub As Integer
    
    If Len(strIn) = 0 Then Exit Function
    
    If gconAscAdd > 5 Or gconAscAdd = 0 Then
        MsgBox "�ַ���������0�Ҳ��ܳ���5��", vbExclamation, "�ַ�ת����������"
        Exit Function
    End If
    
    intSub = Asc(Left(strIn, 1))
    Select Case intSub
        Case 48 To 57, 65 To 90, 97 To 122
            
            intSub = intSub - gconAscAdd
            Select Case intSub
                Case 48 To 57, 65 To 90, 97 To 122
                    '��Щ�����ʾ����ת��
                Case 43 To 47
                    intSub = intSub + 75    '=122-(47-intSub)
                Case 58 To 64
                    intSub = intSub - 7     '=57-(64-intSub)
                Case 91 To 96
                    intSub = intSub - 6     '=90-(96-intSub)
            End Select
            gfAsciiSub = Chr(intSub)
            
        Case Else
            MsgBox "�Ƿ��ַ�ת����" & strIn & "����" & vbCrLf & "�ݲ�֧�����ֺ���ĸ������ַ���", vbExclamation, "��֧���ַ�����"
    End Select
    
End Function

Public Function gfBackComputerInfo(Optional ByVal cType As genumComputerInfoType = ciComputerName, _
        Optional ByVal UseDefault As Boolean = True, Optional ByVal DefaultValue As String = "Null") As String
    '����ָ���ĵ����ϵ���Ϣ
    
    Dim strBack As String, strBuffer As String * 255
    
    If cType = ciComputerName Then  '���������
        strBack = VBA.Environ("ComputerName")   'ֱ��VBA������ȡ
        If Len(strBack) = 0 Then
            Call GetComputerName(strBuffer, 255) '����ȡʧ������API�����ٻ�ȡһ��
            strBack = strBuffer
        End If
    ElseIf cType = ciUserName Then  '�������ǰ�û���
        strBack = VBA.Environ("UserName")
        If Len(strBack) = 0 Then
            Call GetUserName(strBuffer, 255)
            strBack = strBuffer
        End If
    End If
    
    If Len(strBack) = 0 Then  '���Ϊ��ʱ�Ƿ�ʹ��Ĭ��ֵ
        If UseDefault Then strBack = DefaultValue
    End If
    gfBackComputerInfo = strBack
    
End Function


Public Function gfBackConnection(ByVal strCon As String, _
        Optional ByVal CursorLocation As CursorLocationEnum = adUseClient) As ADODB.Connection
    '�������ݿ�����
       
    On Error GoTo LineErr
    
    Set gfBackConnection = New ADODB.Connection
    gfBackConnection.CursorLocation = CursorLocation
    gfBackConnection.ConnectionString = gVar.ConString
    gfBackConnection.CommandTimeout = 5
    gfBackConnection.Open
    
    Exit Function
    
LineErr:
    Call gsAlarmAndLog("���ݿ������쳣")
    
End Function


Public Function gfBackRecordset(ByVal cnSQL As String, _
                Optional ByVal cnCursorType As CursorTypeEnum = adOpenStatic, _
                Optional ByVal cnLockType As LockTypeEnum = adLockReadOnly, _
                Optional ByVal CursorLocation As CursorLocationEnum = adUseClient) As ADODB.Recordset
    '����ָ��SQL��ѯ���ļ�¼��
    
    Dim cnBack As ADODB.Connection
    
    On Error GoTo LineErr

    Set gfBackRecordset = New ADODB.Recordset
    Set cnBack = gfBackConnection(gVar.ConString, CursorLocation)
    If cnBack.State = adStateClosed Then Exit Function
    gfBackRecordset.CursorLocation = CursorLocation
    gfBackRecordset.Open cnSQL, cnBack, cnCursorType, cnLockType
    
    Exit Function

LineErr:
    Call gsAlarmAndLog("���ؼ�¼���쳣")

End Function


Public Function gfBackLogType(Optional ByVal strType As genumLogType = udSelect) As String
    '������־��������
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
    '�������һ���ַ�����ĸ�����֣�
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
    '����������ַ�������Ϊ����
    '���ĳ����޶�ΪgconSumLenλ
    
    Dim strVar As String    '�м����
    Dim strPt As String     '����
    Dim strMid As String    '��ȡ�����ַ����е�ÿһ���ַ�
    Dim intMid As Integer, K As Integer, C As Integer, R As Integer   '����
    
    strIn = Trim(strIn) 'ȥ�ո�
    C = Len(strIn)
    If C <> gconSumLen Then GoTo LineBreak
    
    'һ����ȡ���������������ַ����������ĵĳ���
    R = Val(Mid(strIn, 2, 1))       '��ȡ���ĵĵڶ�λ����ֵ�����ĵ�gconAddLenStart+1λ�������������������
    If R < 1 Then GoTo LineBreak
    
    intMid = Val(Left(strIn, 1))    '��ȡ���ĵĵ�һλ����������ַ�������ֵ�� ��λ�ϵ�����
    C = IIf(intMid < (gconAddLenStart - 2), intMid, gconAddLenStart - 2)  'ͨ����һλ����ֵ����������ֵ��ʮλ�ϵ���������λ��
    K = Val(Mid(strIn, C + 2 + 1, 1))   '��ȡ�����ֵ��ʮλ�ϵ�����
    C = Val(CStr(K) & CStr(intMid))     '�ó������� ����ַ� ����ֵ
    If (C < (gconSumLen - gconMaxPWD)) Or (C > (gconSumLen - 1)) Then GoTo LineBreak
    
    C = gconSumLen - C  '�ó����ĵĳ���
    C = C * 2           '��Ϊ�����в�������ͬ����������ַ�
    
    '����ɾ����������ǰ���gconAddLenStart+ 1 + R ���ַ� �� �������������ַ�
    strVar = Mid(strIn, gconAddLenStart + 1 + R + 1, C)
    If Len(strVar) <> C Then GoTo LineBreak
    
    '��������ʣ�µ�strVar�ַ�
    For K = 1 To C Step 2
        strPt = strPt & gfAsciiSub(Mid(strVar, K, 1))
    Next
    If Len(strPt) <> C / 2 Then GoTo LineBreak
    
    DecryptStringSimple = strPt  '�����ܺõ����ķ��ظ������ĵ�����
    
    Exit Function
    
LineBreak:
'    Err.Clear
'    Err.Number = vbObjectError + 100001
'    Err.Description = "����[" & strIn & "]���ƻ����޷����ܣ�"
'    Call gsAlarmAndLog("���ľ���", False)
    Call gsAlarmAndLogEx("����[" & strIn & "]���ƻ����޷����ܣ�", "���ľ���", False)
End Function

Public Function EncryptStringSimple(ByVal strIn As String) As String
    '��������ַ���(����)���м򵥼��ܣ��������Ĳ����ظ�������
    '���ĳ���<=20���ַ�����ֻ���Ǵ�д��Сд��ĸ�����֣�����ת��ʱ�ᱨ��
    
    Dim strEt As String     '����
    Dim strMid As String    '��ȡ�����ַ����е�ÿһ���ַ�
    Dim strTen As String    '���ĵ�ǰ10���ַ�
    Dim K As Integer, J As Integer, R As Integer  '����
    Dim C As Integer        '���ĵ��ַ�����
    Dim intFill As Integer  '����ַ���
    Dim intRightNum As Integer      'strFill ��λ�ϵ�����
    Dim intAddLenEnd As Integer     '���������ַ�����

    C = Len(Trim(strIn))
    If C = 0 Then
        MsgBox "�����ַ�����Ϊ���ַ����Ҳ����пո�", vbCritical, "���ַ�����"
        Exit Function
    End If
    strIn = Left(strIn, gconMaxPWD) '��ȡǰgconMaxPWD(20)�ַ�
    C = Len(strIn)  '���»�ȡ�ַ���������Ҫ��
    
    'һ�����ַ����е�ÿ���ַ���ASCIIֵǰ��Nλ������һ������ַ��õ�һ���ַ���
    For K = 1 To C
        strEt = strEt & gfAsciiAdd(Mid(strIn, K, 1)) & gfBackOneChar(udUpperLowerNum)
    Next
    If Len(strEt) <> (C * 2) Then
        MsgBox "�����ַ����淶��ֻ�������ֻ���ĸ��", vbCritical, "�ַ�����"
        Exit Function
    End If
    
    '������ת������ַ���strEtǰ�����Ǽ���gconAddLenStart���ַ�
    '   ����gconAddLenStart���ַ��а������ĵĳ�����ϢgconSumLen-C
    '   Ȼ��gconSumLen-C��ֵ�� ��λ��ʮλ����λ��
    '   Ȼ����strTen�ĵڶ�λ����ԭstrTen��Ӧ����������ָ���
    intFill = gconSumLen - C        '����ȥ�����ĸ�����Ҫ�������ַ�����
    intRightNum = intFill Mod 10    '��ȡ��λ�ϵ�����
    strTen = CStr(intRightNum)      '����λ�ϵ����ַ���strTen�ĵ�һλ,Ҳ�����ĵĵ�һλ
    
    '����strTen�ĵ�һλ��ֵ�������������������ֵĸ���
    J = IIf(intRightNum < (gconAddLenStart - 2), intRightNum, gconAddLenStart - 2)
    For K = 1 To J
        strTen = strTen & gfBackOneChar(udNumber)
    Next
    strTen = strTen & CStr(Int(intFill / 10))   '����intFill��ʮλ�ϵ�����
    
    Do
        R = gfBackOneChar(udNumber)     '��ȡһ��1~9�е��������
        If R > 0 Then Exit Do
    Loop
    strTen = Left(strTen, 1) & CStr(R) & Right(strTen, Len(strTen) - 1)
    
    '��strTen�ĳ��Ȳ���gconAddLenStartλ��������������,����strTen���沢�����R������
    J = (gconAddLenStart - 2 - J) + R
    For K = 1 To J
        strTen = strTen & gfBackOneChar(udNumber)
    Next
    strEt = strTen & strEt
    
    '������strEt��׷��intAddLenEnd������ַ��ճ�gconSumLen���ַ�����������
    intAddLenEnd = gconSumLen - (C * 2) - gconAddLenStart - R - 1
    If intAddLenEnd > 0 Then
        For K = 1 To intAddLenEnd
            strEt = strEt & gfBackOneChar(udUpperLowerNum)
        Next
    End If
    
    EncryptStringSimple = strEt  '���strEt���������ķ���ֵ
    
End Function

Public Function gfFileCopy(ByVal strOld As String, ByVal strNew As String, Optional ByVal blnDelOld As Boolean = False) As Boolean
    '�����ļ�
    
    On Error GoTo LineErr
    
    FileCopy strOld, strNew
    gfFileCopy = True
    If blnDelOld Then
        Kill strOld
    End If
    Exit Function
LineErr:
    Call gsAlarmAndLog("�ļ������쳣")
End Function


Public Function gfFileExist(ByVal strPath As String) As Boolean
    '�ж��ļ����ļ�Ŀ¼ �Ƿ����

    Dim strBack As String
        
    On Error GoTo LineErr
    
    If Len(strPath) > 0 Then    '���ַ�������
        strBack = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem)
        If Len(strBack) > 0 Then gfFileExist = True
    End If
  
    Exit Function
    
LineErr:
    Call gsAlarmAndLog("�ж��ļ��쳣")
    
End Function


Public Function gfFileExistEx(ByVal strPath As String) As gtypeValueAndErr
    '��һ�ַ���ֵ��ʽ�����ж��ļ����ļ�Ŀ¼ �Ƿ����
    'ר������Ĺ���gfFileRepair����
    
    Dim strBack As String
    
    On Error GoTo LineErr
    
    If Len(strPath) > 0 Then    '���ַ�������
        strBack = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem)
        If Len(strBack) > 0 Then
            gfFileExistEx.Result = True
        Else
            gfFileExistEx.ErrNum = -1   '�����ڣ�Ҳû�쳣
        End If
    End If
    
    Exit Function
    
LineErr:
    gfFileExistEx.ErrNum = Err.Number   '�쳣�ˣ�Ҳ������������
    Call gsAlarmAndLog("�ļ��жϷ����쳣")
    
End Function

Public Function gfFileIsRun(ByVal pFile As String) As Boolean
    '�ж��ļ��Ƿ񱻴�(������)
    Dim Ret As Long
    
    Ret = CreateFile(pFile, GENERIC_READ Or GENERIC_WRITE, 0&, vbNullString, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    gfFileIsRun = (Ret = INVALID_HANDLE_VALUE)
    CloseHandle Ret
    '��С���ֲ��ԣ��ƺ�û�ã�ֻ���жϿ�ִ���ļ���
End Function


Public Function gfFileOpen(ByVal strFilePath As String) As gtypeValueAndErr
    '��ָ��ȫ·�����ļ�
    
    Dim lngRet As Long
    Dim strDir As String
    
    On Error GoTo LineErr
    
    If gfFileExist(strFilePath) Then
        
        lngRet = ShellExecute(GetDesktopWindow, "open", strFilePath, vbNullString, vbNullString, vbNormalFocus)
        If lngRet = SE_ERR_NOASSOC Then     'û�й����ĳ���
             strDir = Space(260)
             lngRet = GetSystemDirectory(strDir, Len(strDir))
             strDir = Left(strDir, lngRet)
             
            '��ʾ�򿪷�ʽ����
            Call ShellExecute(GetDesktopWindow, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & strFilePath, strDir, vbNormalFocus)
            gfFileOpen.ErrNum = -1   '���ɹ���Ҳû�쳣
        Else
            gfFileOpen.Result = True
        End If
        
    End If
    
    Exit Function
    
LineErr:
    gfFileOpen.ErrNum = Err.Number
    Call gsAlarmAndLog("�ļ����쳣")
    
End Function

Public Function gfFileRename(ByVal strOld As String, ByVal strNew As String) As Boolean
    '�������ļ����ļ���
    
    On Error GoTo LineErr
    
    Close
    Name strOld As strNew
    Close
    gfFileRename = True
    Exit Function
LineErr:
    Close
    Call gsAlarmAndLog("�ļ�/�ļ����������쳣", False)
End Function


Public Function gfFileReNameEx(ByVal strOld As String, ByVal strNew As String) As Boolean
    '�������ļ����ļ�������ɾ�����ڵ����ļ������ļ�
    
    On Error GoTo LineErr
    
    If gfFileExist(strNew) Then
        Kill strNew '���ļ���������ɾ��
    End If
    
    Name strOld As strNew
    gfFileReNameEx = True
    
    Exit Function
LineErr:
    Call gsAlarmAndLog("�ļ�/�ļ����������쳣", False)
End Function


Public Function gfFileRepair(ByVal strFile As String, Optional ByVal blnFolder As Boolean) As Boolean
    '��� �ļ�/�ļ��� ������ �򴴽�
    'ǰ����·�����ϲ�Ŀ¼�ɷ���
    '����blnFolderָ�������·��strFile���ļ�����ΪTrue��Ĭ�����ļ�False
    
    Dim strTemp As String
    Dim typBack As gtypeValueAndErr
    Dim lngLoc As Long
    
    If Right(strFile, 1) = "\" Then
        strFile = Left(strFile, Len(strFile) - 1)   'ȥ����ĩ��"\"
    End If
    strTemp = strFile
    If Len(strTemp) = 0 Then Exit Function          '��ֹ������ַ���
    
    On Error GoTo LineErr

    typBack = gfFileExistEx(strTemp)    '�ж��Ƿ����
    If Not typBack.Result Then          '�ļ�������
        If typBack.ErrNum = -1 Then     '�����쳣
            
            lngLoc = InStrRev(strTemp, "\") '�ж��Ƿ����ϲ�Ŀ¼
            If lngLoc > 0 Then              '���ϲ�Ŀ¼��ݹ�
                strTemp = Left(strTemp, lngLoc - 1) '�ó��ϲ�Ŀ¼�ľ���·��
                Call gfFileRepair(strTemp, True)    '�ݹ���������Ա�֤�ϲ�Ŀ¼����
            End If

            If blnFolder Then                   '����������ļ���
                MkDir strFile                   '�򴴽��ļ���
            Else                                '����������ļ�
                Close                           '�򴴽��ļ�
                Open strFile For Random As #1
                Close
            End If
            
            gfFileRepair = True '�����ɹ�����True
            
        End If
        
    Else
        gfFileRepair = True '·������ֱ�ӷ���True
    End If

LineErr:
    Close
End Function

Public Function gfFolderRepair(ByVal strFile As String) As Boolean
    '��� �ļ��� ������ �򴴽�
    'ǰ����·�����ϲ�Ŀ¼�ɷ���
    
    Dim strTemp As String, strDir As String
    Dim fsObject As Scripting.FileSystemObject
    Dim lngLoc As Long
    
    On Error GoTo LineErr
    
    strTemp = Trim(strFile)
    If Len(strTemp) = 0 Then GoTo LineErr   '��ֹ������ַ���
    
    Set fsObject = New Scripting.FileSystemObject   'ʵ�����ļ�����
    If fsObject.FolderExists(strTemp) Then    '�ж��ļ����Ƿ����
        gfFolderRepair = True '����ֱ�ӷ���True
    Else    '�ļ��в�����
        lngLoc = InStrRev(strTemp, "\") '�ж��Ƿ����ϲ�Ŀ¼��Ŀǰ������\\192.168.2.2����·��
        If lngLoc > 0 Then              '���ϲ�Ŀ¼��ݹ�
            strDir = Left(strTemp, lngLoc - 1) '�ó��ϲ�Ŀ¼�ľ���·��
            Call gfFolderRepair(strDir)        '�ݹ���������Ա�֤�ϲ�Ŀ¼����
        End If
        fsObject.CreateFolder (strTemp) '�ϲ�Ŀ¼ȷ�����ں��򴴽����ļ���
        gfFolderRepair = True           '�����ɹ�ͬʱ����True
    End If
LineErr:
    Set fsObject = Nothing
    If Err.Number > 0 Then
        Call gsAlarmAndLog("�ļ���·��[" & strTemp & "]�쳣��", False)
        Err.Clear
    End If
End Function


Public Function gfFormLoad(ByVal strFormName As String) As Boolean
    '�ж�ָ�������Ƿ񱻼�����
    
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
    'ʹGetSetting�������ص��ַ���ֵ��Ϊ��
    Dim strGet As String
    
    strGet = GetSetting(AppName, Section, Key, Default)
    If BackDefault Then
        If Len(Trim(strGet)) = 0 Then strGet = Default    '����ȡֵΪ���ַ�ʱҲ����Ĭ��ֵ
    End If
    gfGetRegStringValue = strGet
    
End Function

Public Function gfGetRegNumericValue(ByVal AppName As String, ByVal Section As String, _
        ByVal Key As String, Optional ByVal inMinMax As Boolean = True, Optional ByVal Default As Long = 1, _
        Optional ByVal nMin As Long = 1, Optional ByVal nMax As Long = 10) As Long
    'ʹGetSetting��������������ֵ,�������ֵ���ܳ�����С�����ֵ����������Сֵ����
    Dim lngGet As Long
    
    lngGet = Val(GetSetting(AppName, Section, Key, Default))
    If inMinMax Then
        If lngGet < nMin Or lngGet > nMax Then lngGet = Default
    End If
    gfGetRegNumericValue = lngGet
    
End Function

Public Function gfGetSetting(ByVal AppName As String, ByVal Section As String, ByVal Key As String, Optional ByVal strNO As String = "*&^%$#@!") As Boolean
    '�ж�ע�����Ƿ����
    
    Dim strGet As String
    
    strGet = GetSetting(AppName, Section, Key, strNO)
    If strGet <> strNO Then gfGetSetting = True
End Function

Public Function gfLoadAuthority(ByRef frmCur As Form, ByRef ctlCur As Control) As Boolean
    '���ش����еĿ���Ȩ��
    
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
    '�жϴ���Keyֵ�ǲ����Լ����ӽ��
    
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
    '''�����ַ����
    
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

'--------------------------------------------------------------------------

'ö�����ж�������
Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim WindowCaption As String, CaptionLength As Long, WindowClassName As String * 256
    
    CaptionLength = GetWindowTextLength(hwnd)
    WindowCaption = Space(CaptionLength)
    Call GetWindowText(hwnd, WindowCaption, CaptionLength + 1)
    If InStr(1, WindowCaption, MsgBoxTitleText) > 0 Then
        Dlghwnd = hwnd
    End If
    EnumWindowsProc = 1
End Function

'ö�������Ӵ���
Private Function EnumChildWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim WindowCaption As String, CaptionLength As Long, WindowClassName As String * 256
    
    CaptionLength = GetWindowTextLength(hwnd)
    WindowCaption = Space(CaptionLength)
    Call GetWindowText(hwnd, WindowCaption, CaptionLength + 1)
    Call GetClassName(hwnd, WindowClassName, 256)
    If InStr(1, WindowClassName, "Static") > 0 Then
        Dlgtexthwnd = hwnd
    End If
    EnumChildWindowsProc = 1
End Function

Private Function TimeOutString(ByVal strTimeOut As String) As String
    '���ص���ʱ�ִ�
    strTimeOut = CStr(Val(strTimeOut))
    TimeOutString = "(���ڽ���" & strTimeOut & "���ر�)"
End Function

'API����timeSetEventʹ�õĻص�����
Private Function TimeSetProc(ByVal uID As Long, ByVal uMsg As Long, ByVal dwUser As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
    Dim cText As String, nowTime As Long
    
    MediaCount = MediaCount + DelayTime / 1000
    If Dlgtexthwnd > 0 Then
        nowTime = MsgBoxCloseTime - Fix(MediaCount)
        If nowTime <= 0 Then
            Call SendMessage(Dlghwnd, WM_CLOSE, 0, 0) 'ʱ�䵽���رնԻ���
            Call timeKillEvent(TimeID)  'ɾ����ý���ʱ����ʶ
        End If
        cText = MsgBoxPromptText & vbCrLf & TimeOutString(nowTime)
        Call SendMessage(Dlgtexthwnd, WM_SETTEXT, Len(cText), ByVal cText)
    Else
        Call EnumWindows(AddressOf EnumWindowsProc, 0)
        If Dlghwnd > 0 Then
            Call EnumChildWindows(Dlghwnd, AddressOf EnumChildWindowsProc, 0)
        End If
    End If
    TimeSetProc = 1
End Function

'��ʱ�رնԻ���SecondsToClose�������öԻ���ر�ʱ�䣻MsgPrompt�������öԻ�����ʾ�ı���vbButtons���������öԻ���ť��ͼ�ꡣ
Public Function MsgBoxAutoClose(Optional ByVal MsgPrompt As String = "��ʾ��Ϣ", _
        Optional ByVal vbButtons As VbMsgBoxStyle = vbOKOnly + vbInformation, _
        Optional ByVal MsgTitle As String = "�Ի���", _
        Optional ByVal SecondsToClose As Long = 10) As VBA.VbMsgBoxResult
    Dim RetButton As Long '����ֵ��vbAbortRetryIgnore��vbYesNoʱ�޷��Զ��رնԻ���
    
    Dlghwnd = 0
    Dlgtexthwnd = 0
    MsgBoxCloseTime = SecondsToClose
    MsgBoxPromptText = MsgPrompt
    MsgBoxTitleText = MsgTitle
    TimeID = timeSetEvent(DelayTime, 0, AddressOf TimeSetProc, 1, TIME_PERIODIC)  'ʱ����Ϊ500����
    RetButton = MsgBox(MsgBoxPromptText & vbCrLf & TimeOutString(MsgBoxCloseTime), vbButtons, MsgBoxTitleText)      '����msgbox�Ի���
    Call timeKillEvent(TimeID)  'ɾ����ý���ʱ����ʶ
    MediaCount = 0  '����ۼ�ʱ��
    MsgBoxAutoClose = RetButton  '���ذ���ֵ
End Function
'--------------------------------------------------------------------------

Public Function ShellWait(ByVal strShellCommand As String) As Boolean
    '�ȴ�Shell����ִ����ɺ���ִ�к���Ĵ��룬�����ֹShell���첽ִ��.
    Dim osInfo As OSVERSIONINFO
    Dim Ret As Long, nSysVer As Long, pidNotePad As Long, hProcess As Long, lExitCode As Long
    Dim strSave As String, Path As String, sCommPath As String, sExecString As String
    
    On Error Resume Next
    
    osInfo.dwOSVersionInfoSize = Len(osInfo) 'Set the structure size
    Ret& = GetVersionEx(osInfo) 'Get the Windows version
    If Ret& = 0 Then MsgBox "Error Getting Version Information" 'Chack for errors
    nSysVer = osInfo.dwPlatformId

    strSave = String(200, Chr$(0)) 'Create a buffer string
    Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) 'Get the windows directory
    If Mid(Path, Len(Path), 1) <> "\" Then Path = Path & "\"
    sCommPath = Path
    If nSysVer = 1 Then 'windows98
        sExecString = sCommPath + "command.com  /c " + strShellCommand
    Else
        sExecString = strShellCommand
    End If
    
    pidNotePad = Shell(sExecString, vbHide) '����ִ�г��������ID�����ɹ�����0
    hProcess = OpenProcess(Process_query_infomation, True, pidNotePad)  '�򿪽���
    Do
        GetExitCodeProcess hProcess, lExitCode  '��ȡ�����ж��˳�����
        DoEvents
    Loop While lExitCode = Still_Active
    CloseHandle (pidNotePad)    '�رս���
    
    If Err.Number Then
        Call gsAlarmAndLog(Err.Number & "--" & Err.Description, False)
    Else
        ShellWait = True
    End If
End Function

Public Function ShowBackupTimeInfo(ByVal BKInterval As Long, ByVal BKDate As Date) As String
    'ת������Ƶ���뱸��ʱ��
    Dim strShow As String, strNext As String, strTime As String
    
    strTime = Format(BKDate, "HH:mm:ss")
    Select Case BKInterval
        Case 0
            strShow = "��"
        Case 1
            strShow = Format(BKDate, "ÿ��") & strTime
        Case 2
            strShow = "ÿ��" & WeekdayName(Weekday(BKDate)) & strTime
        Case 3
            strShow = Format(BKDate, "ÿ��d��") & strTime
        Case 4
            strShow = Format(BKDate, "ÿ��M��d��") & strTime
        Case 5
            strShow = "ÿ" & gVar.ParaBackupIntervalDays & "��" & strTime
        Case Else
            strShow = "δ����"
    End Select
    ShowBackupTimeInfo = strShow
End Function

Public Function ShowBackupNextTime(ByVal BKInterval As Long, ByVal BKDate As Date) As String
    'ת������Ƶ���뱸��ʱ��
    Dim strShow As String
    Dim nowTime As Date, BackTime As Date, NextDay As Date, ThisYear As Date
    Dim NowWeek As Long, BKWeek As Long, modDay As Long
    Dim NowDay As Long, BKDay As Long, NowMonth As Long, BKMonth As Long
    
    nowTime = Time
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
                Case 1  'ÿ��
                    If nowTime <= BackTime Then
                        NextDay = Date
                    Else
                        NextDay = Date + 1
                    End If
                Case 2  'ÿ��
                    If (NowWeek < BKWeek) Or (NowWeek = BKWeek And nowTime <= BackTime) Then
                        NextDay = Date + (BKWeek - NowWeek)
                    Else
                        NextDay = Date + (7 - NowWeek + BKWeek)
                    End If
                Case 3  'ÿ��
                    If (NowDay < BKDay) Or (NowDay = BKDay And nowTime <= BackTime) Then
                        NextDay = Date + (BKDay - NowDay)
                    Else
                        NextDay = DateAdd("m", DateDiff("m", BKDate, DateAdd("m", 1, Date)), BKDate)
                    End If
                Case 4  'ÿ��
                    If (Date < ThisYear) Or (Date = ThisYear And nowTime <= BackTime) Then
                        NextDay = ThisYear
                    Else
                        NextDay = DateAdd("yyyy", 1, ThisYear)
                    End If
                Case 5  'ÿN��
                    If Now < BKDate Then
                        NextDay = BKDate
                    Else
                        modDay = (Now - BKDate) Mod gVar.ParaBackupIntervalDays
                        If modDay = 0 And nowTime > BackTime Then
                            modDay = modDay + gVar.ParaBackupIntervalDays
                        End If
                        NextDay = Date + modDay
                    End If
            End Select
            strShow = Format(NextDay, "yyyy-MM-dd ") & Format(BKDate, "HH:mm:ss")
        Case Else
            strShow = "��"
    End Select
    ShowBackupNextTime = strShow
End Function

