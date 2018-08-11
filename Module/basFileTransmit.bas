Attribute VB_Name = "basFileTransmit"
'Ҫ��Winsock�ؼ��ڿͻ��������˶����뽨�����飬����Indexֵ���Ӧ������������±�Ҫ��ͬ

Option Explicit



Public Function gfBackVersion(ByVal strFile As String) As String
    '�����ļ��İ汾��
    Dim objFile As Scripting.FileSystemObject
    
    If Not gfDirFile(strFile) Then Exit Function
    Set objFile = New FileSystemObject
    gfBackVersion = objFile.GetFileVersion(strFile)

    Set objFile = Nothing
End Function


Public Function gfCheckIP(ByVal strIP As String) As String
    Dim K As Long
    Dim arrIP() As String
    
    arrIP = Split(strIP, ".")
    If UBound(arrIP) <> 3 Then GoTo LineOver
    For K = 0 To 3
        If Not IsNumeric(arrIP(K)) Then GoTo LineOver
        If Val(arrIP(K)) < 0 Or Val(arrIP(K)) > 255 Then GoTo LineOver
        arrIP(K) = CStr(Val(arrIP(K)))
    Next
    gfCheckIP = arrIP(0) & "." & arrIP(1) & "." & arrIP(2) & "." & arrIP(3)
    Exit Function
    
LineOver:
    gfCheckIP = "127.0.0.1"
End Function

Public Function gfCloseApp(ByVal strName As String) As Boolean
    '�ر�ָ��Ӧ�ó������
    
    Dim winHwnd As Long
    Dim RetVal As Long
    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim objProcess As Object
    
    On Error GoTo LineErr
    
''    winHwnd = FindWindow(vbNullString, strName) '���Ҵ��ڣ�strName���ݼ��������Ͽ����Ĵ��ڱ���
''    If winHwnd <> 0 Then    '��Ϊ0��ʾ�ҵ�����
''        RetVal = PostMessage(winHwnd, WM_CLOSE, 0&, 0&) '���͹رմ�����Ϣ,����ֵΪ0��ʾ�ر�ʧ��
''    End If
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("select * from Win32_Process where Name='" & strName & "' ")
    For Each objProcess In colProcessList
        RetVal = objProcess.Terminate
        If RetVal <> 0 Then Exit Function   '���۲�=0ʱ�رս��̳ɹ������ɹ�ʱ����ֵ��Ϊ��
    Next
    
    gfCloseApp = True   'ȫ���رճɹ��򲻴��ڸý�����ʱ
    
LineErr:
    Set objWMIService = Nothing
    Set colProcessList = Nothing
    Set objProcess = Nothing
End Function


Public Function gfDirFile(ByVal strFile As String) As Boolean
    Dim strDir As String
    
    strFile = Trim(strFile)
    If Len(strFile) = 0 Then Exit Function
    
    On Error GoTo LineErr
    
    strDir = Dir(strFile, vbHidden + vbReadOnly + vbSystem)
    If Len(strDir) > 0 Then
        SetAttr strFile, vbNormal
        gfDirFile = True
    End If
    
    Exit Function
LineErr:
    Debug.Print "Error:gfDirFile--" & Err.Number & "  " & Err.Description
End Function

Public Function gfDirFolder(ByVal strFolder As String) As Boolean
    Dim strDir As String
    
    strFolder = Trim(strFolder)
    If Len(strFolder) = 0 Then Exit Function
    
    On Error GoTo LineErr
    
    strDir = Dir(strFolder, vbHidden + vbReadOnly + vbSystem + vbDirectory)
    If Len(strDir) = 0 Then
        MkDir strFolder
    Else
        SetAttr strFolder, vbNormal
    End If
    gfDirFolder = True
    
    Exit Function
LineErr:
    Debug.Print "Error:gfDirFolder--" & Err.Number & "  " & Err.Description
End Function

Public Function gfFileInfoJoin(ByVal intIndex As Integer, Optional ByVal enmType As genumFileTransimitType = ftSend) As String
    '�ļ���Ϣƴ��
    Dim strType As String
    
    strType = IIf(enmType = ftReceive, gVar.PTFileReceive, gVar.PTFileSend) 'ȷ���ļ��������͡�վ�ڿͻ��˽Ƕ�ȷ����
    With gArr(intIndex)
        gfFileInfoJoin = gVar.PTFileFolder & .FileFolder & gVar.PTFileName & .FileName & gVar.PTFileSize & .FileSizeTotal & strType
    End With
    
End Function

Public Function gfNotifyIconAdd(ByRef frmCur As Form) As Boolean
    '��������ͼ��
    With gNotifyIconData
        .hwnd = frmCur.hwnd
        .uID = frmCur.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP Or NIF_INFO
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frmCur.Icon.Handle
        .szTip = App.Title & " " & App.Major & "." & App.Minor & _
            "." & App.Revision & vbNullChar   '����ƶ�����ͼ��ʱ��ʾ��Tip��Ϣ
        .cbSize = Len(gNotifyIconData)
    End With
    Call Shell_NotifyIcon(NIM_ADD, gNotifyIconData)
End Function

Public Function gfNotifyIconBalloon(ByRef frmCur As Form, ByVal BalloonInfo As String, _
    ByVal BalloonTitle As String, Optional IconFlag As genumNotifyIconFlag = NIIF_INFO) As Boolean
    '����ͼ�굯��������Ϣ
    With gNotifyIconData
        .dwInfoFlags = IconFlag
        .szInfoTitle = BalloonTitle & vbNullChar
        .szInfo = BalloonInfo & vbNullChar
        .cbSize = Len(gNotifyIconData)
    End With
    Call gfNotifyIconModify(gNotifyIconData)
End Function

Public Function gfNotifyIconDelete(ByRef frmCur As Form) As Boolean
    'ɾ������ͼ��
    Call Shell_NotifyIcon(NIM_DELETE, gNotifyIconData)
End Function

Public Function gfNotifyIconModify(nfIconData As gtypeNOTIFYICONDATA) As Boolean
    '�޸�����ͼ����Ϣ
    gNotifyIconData = nfIconData
    Call Shell_NotifyIcon(NIM_MODIFY, gNotifyIconData)
End Function

Public Function gfRegOperate(ByVal RegHKEY As genumRegRootDirectory, ByVal lpSubKey As String, _
    ByVal lpValueName As String, Optional ByVal lpType As genumRegDataType = REG_SZ, _
    Optional ByRef lpValue As String, Optional ByVal lpOp As genumRegOperateType = RegRead) As Boolean
    '
    Dim Ret As Long, hKey As Long, lngLength As Long
    Dim Buff() As Byte
    
    
    Ret = RegOpenKey(RegHKEY, lpSubKey, hKey)
    If Ret = 0 Then
        Select Case lpOp
            Case RegDelete
                Ret = RegDeleteValue(hKey, lpValueName)
                If Ret = 0 Then
                    gfRegOperate = True
                End If
                
            Case RegWrite
                lngLength = LenB(StrConv(lpValue, vbFromUnicode))   '����LenB��StrConv�Ļ�lpValue�ַ������ȶԲ���
                Ret = RegSetValueEx(hKey, lpValueName, 0, lpType, ByVal lpValue, lngLength)
                If Ret = 0 Then
                    gfRegOperate = True
'Debug.Print "W", lpValue, lngLength
                End If
                
            Case Else
                Ret = RegQueryValueEx(hKey, lpValueName, 0, lpType, ByVal 0, lngLength) '��ȡֵ�ĳ���
                If Ret = 0 And lngLength > 0 Then
                    ReDim Buff(lngLength - 1)   '�ض��建���С
                    Ret = RegQueryValueEx(hKey, lpValueName, 0, lpType, Buff(0), lngLength) 'ȡֵ
                    If Ret = 0 And lngLength > 1 Then
                        ReDim Preserve Buff(lngLength - 2)
                        lpValue = StrConv(Buff, vbUnicode)
                        gfRegOperate = True
'Debug.Print "R", lpValue, lngLength - 1
                    End If
                End If
                
        End Select
    End If
    
    Call RegCloseKey(hKey)
    
End Function


Public Function gfRestoreInfo(ByVal strInfo As String, sckGet As MSWinsockLib.Winsock) As Boolean
    '��ԭ���յ����ļ���Ϣ
    
    With gArr(sckGet.Index)
        If InStr(strInfo, gVar.PTFileFolder) > 0 Then
            '���ж��ƺ�����Ӧ�ڿͻ����������ϴ��ļ�ʱ�����������д�ȷ��
            
            Dim lngFod As Long, lngFile As Long, lngSize As Long
            Dim lngSend As Long, lngReceive As Long, lngType As Long
            Dim strFod As String, strSize As String, strType As String
            
            lngFod = InStr(strInfo, gVar.PTFileFolder)
            lngFile = InStr(strInfo, gVar.PTFileName)
            lngSize = InStr(strInfo, gVar.PTFileSize)
            lngSend = InStr(strInfo, gVar.PTFileSend)
            lngReceive = InStr(strInfo, gVar.PTFileReceive)
            
            If lngFile > 0 Then
                gArr(sckGet.Index) = gArr(0)    '�ȳ�ʼ���ļ��������Ϊ����Ϣ
                
                If (lngSend > 0 And lngReceive > 0) Or (lngSend = 0 And lngReceive = 0) Then Exit Function
                strType = IIf(lngSend > 0, gVar.PTFileSend, gVar.PTFileReceive)
                lngType = IIf(lngSend > 0, lngSend, lngReceive)
                
                .FileFolder = Mid(strInfo, lngFod + Len(gVar.PTFileFolder), lngFile - (lngFod + Len(gVar.PTFileFolder)))
                strFod = gVar.AppPath & .FileFolder
                If Not gfDirFolder(strFod) Then Exit Function
                
                .FileName = Mid(strInfo, lngFile + Len(gVar.PTFileName), lngSize - (lngFile + Len(gVar.PTFileName)))
                
                strSize = Mid(strInfo, lngSize + Len(gVar.PTFileSize), lngType - (lngSize + Len(gVar.PTFileSize)))
                If Not IsNumeric(strSize) Then Exit Function
                
                If strType <> Mid(strInfo, lngType) Then Exit Function
                
                If strType = gVar.PTFileSend Then   '��״̬������ڿͻ��˵ġ��ͻ���������������ļ���
                    .FileSizeTotal = CLng(strSize)
                    .FilePath = strFod & "\" & .FileName
                    Call gfSendInfo(gVar.PTFileStart, sckGet)
                    .FileTransmitState = True
                    
                ElseIf strType = gVar.PTFileReceive Then    '�ͻ���Ҫ�����˴���ָ���ļ����ͻ��ˡ�
                    .FilePath = strFod & "\" & .FileName
                    If gfDirFile(.FilePath) Then
                        .FileSizeTotal = FileLen(.FilePath)
                        Call gfSendInfo(gVar.PTFileExist & gVar.PTFileSize & .FileSizeTotal, sckGet)
                    Else
                        gArr(sckGet.Index) = gArr(0)
                        Call gfSendInfo(gVar.PTFileNoExist, sckGet)
                    End If
                End If
                gfRestoreInfo = True
            End If
        ElseIf InStr(strInfo, gVar.PTVersionNotUpdate) > 0 Then '������
            
        End If
    End With

End Function

Public Function gfSendFile(ByVal strFile As String, sckSend As MSWinsockLib.Winsock) As Boolean
    Dim lngSendSize As Long, lngRemain As Long
    Dim byteSend() As Byte
    
    With gArr(sckSend.Index)
        If .FileNumber = 0 Then
            .FileNumber = FreeFile
            Open strFile For Binary As #.FileNumber
            .FileTransmitState = True
        End If
        
        lngSendSize = gVar.FTChunkSize
        lngRemain = .FileSizeTotal - Loc(.FileNumber)
        If lngSendSize > lngRemain Then lngSendSize = lngRemain
        
        ReDim byteSend(lngSendSize - 1)
        Get #.FileNumber, , byteSend
        sckSend.SendData byteSend
        
        .FileSizeCompleted = .FileSizeCompleted + lngSendSize
        If .FileSizeCompleted = .FileSizeTotal Then Close #.FileNumber
        
    End With
    
End Function

Public Function gfSendInfo(ByVal strInfo As String, sckSend As MSWinsockLib.Winsock) As Boolean
    If sckSend.State = 7 Then
        sckSend.SendData strInfo
        DoEvents
'''        Call Sleep(200)
        gfSendInfo = True
    End If
End Function

Public Function gfShell(ByVal strFile As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus) As Boolean
    '����Shell�����쳣
    
    Dim Ret
    
    On Error Resume Next
    
    Ret = Shell(strFile, WindowStyle)

    If Ret > 0 Then gfShell = True
    
End Function

Public Function gfShellExecute(ByVal strFile As String) As Boolean
    'ִ�г������ļ����ļ���
    '''Call ShellExecute(Me.hwnd, "open", strFile, vbNullString, vbNullString, 1)

    Dim lngRet As Long
    Dim strDir As String
    
    lngRet = ShellExecute(GetDesktopWindow, "open", strFile, vbNullString, vbNullString, vbNormalFocus)

    ' û�й����ĳ���
    If lngRet = SE_ERR_NOASSOC Then
         strDir = Space$(260)
         lngRet = GetSystemDirectory(strDir, Len(strDir))
         strDir = Left$(strDir, lngRet)
       ' ��ʾ�򿪷�ʽ����
         lngRet = ShellExecute(GetDesktopWindow, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & strFile, strDir, vbNormalFocus)
    End If
    
    If lngRet > 32 Then gfShellExecute = True
    
End Function


Public Function gfVersionCompare(ByVal strVerCL As String, ByVal strVerSV As String) As String
    '�¾ɰ汾�űȽ�
    Dim ArrCL() As String, ArrSV() As String
    Dim K As Long, C As Long
    
    ArrCL = Split(strVerCL, ".")
    ArrSV = Split(strVerSV, ".")
    K = UBound(ArrCL)
    C = UBound(ArrSV)
    If K = C And K = 3 Then
        For K = 0 To C
            If Not IsNumeric(ArrCL(K)) Then
                gfVersionCompare = "�ͻ��˰汾�쳣"
                Exit For
            End If
            If Not IsNumeric(ArrSV(K)) Then
                gfVersionCompare = "����˰汾�쳣"
                Exit For
            End If
            
            If Val(ArrSV(K)) > Val(ArrCL(K)) Then
                gfVersionCompare = "1" '˵�����°汾
                Exit For
            End If
        Next
        If K = C + 1 Then gfVersionCompare = "0" '˵��û���°棬���ø���
    Else
        If K = 3 And C <> 3 Then
            gfVersionCompare = "����˰汾��ȡ�쳣"
        ElseIf C = 3 And K <> 3 Then
            gfVersionCompare = "�ͻ��˰汾��ȡ�쳣"
        Else
            gfVersionCompare = "�汾��ȡ�쳣"
        End If
    End If
    
End Function

Public Sub gsFormEnable(frmCur As Form, Optional ByVal blnState As Boolean)
    With frmCur
        If blnState Then
            .Enabled = True
            .MousePointer = 0
        Else
            .Enabled = False
            .MousePointer = 13
        End If
    End With
End Sub


