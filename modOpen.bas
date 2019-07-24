Attribute VB_Name = "modOpen"
Option Explicit


'获取一个已中断进程的退出代码.非零表示成功，零表示失败
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
'打开一个已存在的进程对象，并返回进程的句柄
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'关闭一个内核对象。其中包括文件、文件映射、进程、线程、安全和同步对象等
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const CREATE_NEW_CONSOLE = &H10
Private Const Process_query_infomation = &H400  '获取进程的令牌、退出码和优先级等信息
Private Const Still_Active = &H103

'得到当前平台和操作系统有关的版本信息.非零表示成功，零表示失败
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long '初始化为结构的大小
    dwMajorVersion As Long      '系统主版本号
    dwMinorVersion As Long      '系统次版本号
    dwBuildNumber As Long       '系统构建号
    dwPlatformId As Long        '系统支持的平台
    szCSDVersion As String * 128    '
End Type

'得到当前window系统目的完整路径名
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'---------------------------------------------------------------------------------------

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
'-----------------------------------------------------------------------------------

'''网上抄的MD5加密解密
'''链接地址：https://www.cnblogs.com/youyouran/p/5381050.html

Public Declare Function CryptAcquireContext Lib "advapi32.dll" _
    Alias "CryptAcquireContextA" ( _
    ByRef phProv As Long, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long

Public Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long

Public Declare Function CryptCreateHash Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hKey As Long, _
    ByVal dwFlags As Long, _
    ByRef phHash As Long) As Long

Public Declare Function CryptDestroyHash Lib "advapi32.dll" ( _
    ByVal hHash As Long) As Long

Public Declare Function CryptHashData Lib "advapi32.dll" ( _
    ByVal hHash As Long, _
    pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Public Declare Function CryptDeriveKey Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hBaseData As Long, _
    ByVal dwFlags As Long, _
    ByRef phKey As Long) As Long

Public Declare Function CryptDestroyKey Lib "advapi32.dll" ( _
    ByVal hKey As Long) As Long

Public Declare Function CryptEncrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal hHash As Long, _
    ByVal Final As Long, _
    ByVal dwFlags As Long, _
    pbData As Any, _
    ByRef pdwDataLen As Long, _
    ByVal dwBufLen As Long) As Long

Public Declare Function CryptDecrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal hHash As Long, _
    ByVal Final As Long, _
    ByVal dwFlags As Long, _
    pbData As Any, _
    ByRef pdwDataLen As Long) As Long

Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Dest As Any, _
    Src As Any, _
    ByVal Ln As Long)

Private Const PROV_RSA_FULL = 1

Private Const CRYPT_NEWKEYSET = &H8

Private Const ALG_CLASS_HASH = 32768
Private Const ALG_CLASS_DATA_ENCRYPT = 24576&

Private Const ALG_TYPE_ANY = 0
Private Const ALG_TYPE_BLOCK = 1536&
Private Const ALG_TYPE_STREAM = 2048&

Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MD4 = 2
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA1 = 4

Private Const ALG_SID_DES = 1
Private Const ALG_SID_3DES = 3
Private Const ALG_SID_RC2 = 2
Private Const ALG_SID_RC4 = 1

Public Enum HASHALGORITHM
   MD2 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2
   MD4 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4
   MD5 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5
   SHA1 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1
End Enum

Public Enum ENCALGORITHM
   DES = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_DES
   [3DES] = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES
   RC2 = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2
   RC4 = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4
End Enum

Dim HexMatrix(15, 15) As Byte
'---------------------------------------------------------------------------

Private Const mconStrKey As String = "ftkey" '公共密钥
Private Const mconStrBKbak As String = ".bak"    '备份文件的
Private Const mconStrBKrst As String = ".rst"    '备份配置文件扩展名

'---------------------------------------------------------------------------

'================================================
'加密
'================================================
Public Function EncryptString(ByVal str As String, password As String) As String
    Dim byt() As Byte
    Dim HASHALGORITHM As HASHALGORITHM
    Dim ENCALGORITHM As ENCALGORITHM
On Error GoTo LineERR
    byt = str
    HASHALGORITHM = MD5
    ENCALGORITHM = RC4
    EncryptString = BytesToHex(Encrypt(byt, password, HASHALGORITHM, ENCALGORITHM))
LineERR:
    If Err.Number Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "加密异常"
    End If
End Function

Public Function EncryptByte(byt() As Byte, password As String) As Byte()
    Dim HASHALGORITHM As HASHALGORITHM
    Dim ENCALGORITHM As ENCALGORITHM
    HASHALGORITHM = MD5
    ENCALGORITHM = RC4
    EncryptByte = Encrypt(byt, password, HASHALGORITHM, ENCALGORITHM)
End Function

Public Function Encrypt(data() As Byte, ByVal password As String, Optional ByVal HASHALGORITHM As HASHALGORITHM = MD5, Optional ByVal ENCALGORITHM As ENCALGORITHM = RC4) As Byte()
    Dim lRes As Long
    Dim hProv As Long
    Dim hHash As Long
    Dim hKey As Long
    Dim lBufLen As Long
    Dim lDataLen As Long
    Dim abData() As Byte
    lRes = CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, 0)
    If lRes = 0 And Err.LastDllError = &H80090016 Then lRes = CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET)
    If lRes <> 0 Then
        lRes = CryptCreateHash(hProv, HASHALGORITHM, 0, 0, hHash)
        If lRes <> 0 Then
            lRes = CryptHashData(hHash, ByVal password, Len(password), 0)
            If lRes <> 0 Then
                lRes = CryptDeriveKey(hProv, ENCALGORITHM, hHash, 0, hKey)
                If lRes <> 0 Then
                    lBufLen = UBound(data) - LBound(data) + 1
                    lDataLen = lBufLen
                    lRes = CryptEncrypt(hKey, 0&, 1, 0, ByVal 0&, lBufLen, 0)
                    If lRes <> 0 Then
                        If lBufLen < lDataLen Then lBufLen = lDataLen
                        ReDim abData(0 To lBufLen - 1)
                        MoveMemory abData(0), data(LBound(data)), lDataLen
                        lRes = CryptEncrypt(hKey, 0&, 1, 0, abData(0), lBufLen, lDataLen)
                        If lRes <> 0 Then
                            If lDataLen <> lBufLen Then ReDim Preserve abData(0 To lBufLen - 1)
                            Encrypt = abData
                        End If
                    End If
                End If
                CryptDestroyKey hKey
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0
    End If
    If lRes = 0 Then Err.Raise Err.LastDllError
 End Function
 
 
'================================================
'解密
'================================================
Public Function DecryptString(ByVal str As String, password As String) As String
    Dim byt() As Byte
    Dim HASHALGORITHM As HASHALGORITHM
    Dim ENCALGORITHM As ENCALGORITHM
On Error GoTo LineERR
    byt = HexToBytes(str)
    HASHALGORITHM = MD5
    ENCALGORITHM = RC4
    DecryptString = Decrypt(byt, password, HASHALGORITHM, ENCALGORITHM)
LineERR:
    If Err.Number Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "解密异常"
    End If
End Function

Public Function DecryptByte(byt() As Byte, password As String) As Byte()
    Dim HASHALGORITHM As HASHALGORITHM
    Dim ENCALGORITHM As ENCALGORITHM
    HASHALGORITHM = MD5
    ENCALGORITHM = RC4
    DecryptByte = Decrypt(byt, password, HASHALGORITHM, ENCALGORITHM)
End Function

Public Function Decrypt(data() As Byte, ByVal password As String, Optional ByVal HASHALGORITHM As HASHALGORITHM = MD5, Optional ByVal ENCALGORITHM As ENCALGORITHM = RC4) As Byte()
    Dim lRes As Long
    Dim hProv As Long
    Dim hHash As Long
    Dim hKey As Long
    Dim lBufLen As Long
    Dim abData() As Byte
    lRes = CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, 0)
    If lRes = 0 And Err.LastDllError = &H80090016 Then lRes = CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET)
    If lRes <> 0 Then
        lRes = CryptCreateHash(hProv, HASHALGORITHM, 0, 0, hHash)
        If lRes <> 0 Then
            lRes = CryptHashData(hHash, ByVal password, Len(password), 0)
            If lRes <> 0 Then
                lRes = CryptDeriveKey(hProv, ENCALGORITHM, hHash, 0, hKey)
                If lRes <> 0 Then
                    lBufLen = UBound(data) - LBound(data) + 1
                    ReDim abData(0 To lBufLen - 1)
                    MoveMemory abData(0), data(LBound(data)), lBufLen
                    lRes = CryptDecrypt(hKey, 0&, 1, 0, abData(0), lBufLen)
                    If lRes <> 0 Then
                        ReDim Preserve abData(0 To lBufLen - 1)
                        Decrypt = abData
                    End If
                End If
                CryptDestroyKey hKey
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0
    End If
    If lRes = 0 Then Err.Raise Err.LastDllError
End Function

'================================================
'字节与十六进制字符串的转换
'================================================
Public Function BytesToHex(bits() As Byte) As String
    Dim I As Long
    Dim b
    Dim s As String
    For Each b In bits
        If b < 16 Then
            s = s & "0" & Hex(b)
        Else
            s = s & Hex(b)
        End If
    Next
    BytesToHex = s
End Function

Public Function HexToBytes(sHex As String) As Byte()
    Dim b() As Byte
    Dim rst() As Byte
    Dim I As Long
    Dim n As Long
    Dim m1 As Byte
    Dim m2 As Byte
    If HexMatrix(15, 15) = 0 Then Call MatrixInitialize
    b = StrConv(sHex, vbFromUnicode)
    I = (UBound(b) + 1) / 2 - 1
    ReDim rst(I)
    For I = 0 To UBound(b) Step 2
        If b(I) > 96 Then
            m1 = b(I) - 87
        ElseIf b(I) > 64 Then
            m1 = b(I) - 55
        ElseIf b(I) > 47 Then
            m1 = b(I) - 48
        End If
        If b(I + 1) > 96 Then
            m2 = b(I + 1) - 87
        ElseIf b(I + 1) > 64 Then
            m2 = b(I + 1) - 55
        ElseIf b(I + 1) > 47 Then
            m2 = b(I + 1) - 48
        End If
        rst(n) = HexMatrix(m1, m2)
        n = n + 1
    Next I
    HexToBytes = rst
End Function

Public Sub MatrixInitialize()
    HexMatrix(0, 0) = &H0:    HexMatrix(0, 1) = &H1:    HexMatrix(0, 2) = &H2:    HexMatrix(0, 3) = &H3:    HexMatrix(0, 4) = &H4:    HexMatrix(0, 5) = &H5:    HexMatrix(0, 6) = &H6:    HexMatrix(0, 7) = &H7
    HexMatrix(0, 8) = &H8:    HexMatrix(0, 9) = &H9:    HexMatrix(0, 10) = &HA:   HexMatrix(0, 11) = &HB:   HexMatrix(0, 12) = &HC:   HexMatrix(0, 13) = &HD:   HexMatrix(0, 14) = &HE:   HexMatrix(0, 15) = &HF
    HexMatrix(1, 0) = &H10:   HexMatrix(1, 1) = &H11:   HexMatrix(1, 2) = &H12:   HexMatrix(1, 3) = &H13:   HexMatrix(1, 4) = &H14:   HexMatrix(1, 5) = &H15:   HexMatrix(1, 6) = &H16:   HexMatrix(1, 7) = &H17
    HexMatrix(1, 8) = &H18:   HexMatrix(1, 9) = &H19:   HexMatrix(1, 10) = &H1A:  HexMatrix(1, 11) = &H1B:  HexMatrix(1, 12) = &H1C:  HexMatrix(1, 13) = &H1D:  HexMatrix(1, 14) = &H1E:  HexMatrix(1, 15) = &H1F
    HexMatrix(2, 0) = &H20:   HexMatrix(2, 1) = &H21:   HexMatrix(2, 2) = &H22:   HexMatrix(2, 3) = &H23:   HexMatrix(2, 4) = &H24:   HexMatrix(2, 5) = &H25:   HexMatrix(2, 6) = &H26:   HexMatrix(2, 7) = &H27
    HexMatrix(2, 8) = &H28:   HexMatrix(2, 9) = &H29:   HexMatrix(2, 10) = &H2A:  HexMatrix(2, 11) = &H2B:  HexMatrix(2, 12) = &H2C:  HexMatrix(2, 13) = &H2D:  HexMatrix(2, 14) = &H2E:  HexMatrix(2, 15) = &H2F
    HexMatrix(3, 0) = &H30:   HexMatrix(3, 1) = &H31:   HexMatrix(3, 2) = &H32:   HexMatrix(3, 3) = &H33:   HexMatrix(3, 4) = &H34:   HexMatrix(3, 5) = &H35:   HexMatrix(3, 6) = &H36:   HexMatrix(3, 7) = &H37
    HexMatrix(3, 8) = &H38:   HexMatrix(3, 9) = &H39:   HexMatrix(3, 10) = &H3A:  HexMatrix(3, 11) = &H3B:  HexMatrix(3, 12) = &H3C:  HexMatrix(3, 13) = &H3D:  HexMatrix(3, 14) = &H3E:  HexMatrix(3, 15) = &H3F
    HexMatrix(4, 0) = &H40:   HexMatrix(4, 1) = &H41:   HexMatrix(4, 2) = &H42:   HexMatrix(4, 3) = &H43:   HexMatrix(4, 4) = &H44:   HexMatrix(4, 5) = &H45:   HexMatrix(4, 6) = &H46:   HexMatrix(4, 7) = &H47
    HexMatrix(4, 8) = &H48:   HexMatrix(4, 9) = &H49:   HexMatrix(4, 10) = &H4A:  HexMatrix(4, 11) = &H4B:  HexMatrix(4, 12) = &H4C:  HexMatrix(4, 13) = &H4D:  HexMatrix(4, 14) = &H4E:  HexMatrix(4, 15) = &H4F
    HexMatrix(5, 0) = &H50:   HexMatrix(5, 1) = &H51:   HexMatrix(5, 2) = &H52:   HexMatrix(5, 3) = &H53:   HexMatrix(5, 4) = &H54:   HexMatrix(5, 5) = &H55:   HexMatrix(5, 6) = &H56:   HexMatrix(5, 7) = &H57
    HexMatrix(5, 8) = &H58:   HexMatrix(5, 9) = &H59:   HexMatrix(5, 10) = &H5A:  HexMatrix(5, 11) = &H5B:  HexMatrix(5, 12) = &H5C:  HexMatrix(5, 13) = &H5D:  HexMatrix(5, 14) = &H5E:  HexMatrix(5, 15) = &H5F
    HexMatrix(6, 0) = &H60:   HexMatrix(6, 1) = &H61:   HexMatrix(6, 2) = &H62:   HexMatrix(6, 3) = &H63:   HexMatrix(6, 4) = &H64:   HexMatrix(6, 5) = &H65:   HexMatrix(6, 6) = &H66:   HexMatrix(6, 7) = &H67
    HexMatrix(6, 8) = &H68:   HexMatrix(6, 9) = &H69:   HexMatrix(6, 10) = &H6A:  HexMatrix(6, 11) = &H6B:  HexMatrix(6, 12) = &H6C:  HexMatrix(6, 13) = &H6D:  HexMatrix(6, 14) = &H6E:  HexMatrix(6, 15) = &H6F
    HexMatrix(7, 0) = &H70:   HexMatrix(7, 1) = &H71:   HexMatrix(7, 2) = &H72:   HexMatrix(7, 3) = &H73:   HexMatrix(7, 4) = &H74:   HexMatrix(7, 5) = &H75:   HexMatrix(7, 6) = &H76:   HexMatrix(7, 7) = &H77
    HexMatrix(7, 8) = &H78:   HexMatrix(7, 9) = &H79:   HexMatrix(7, 10) = &H7A:  HexMatrix(7, 11) = &H7B:  HexMatrix(7, 12) = &H7C:  HexMatrix(7, 13) = &H7D:  HexMatrix(7, 14) = &H7E:  HexMatrix(7, 15) = &H7F
    HexMatrix(8, 0) = &H80:   HexMatrix(8, 1) = &H81:   HexMatrix(8, 2) = &H82:   HexMatrix(8, 3) = &H83:   HexMatrix(8, 4) = &H84:   HexMatrix(8, 5) = &H85:   HexMatrix(8, 6) = &H86:   HexMatrix(8, 7) = &H87
    HexMatrix(8, 8) = &H88:   HexMatrix(8, 9) = &H89:   HexMatrix(8, 10) = &H8A:  HexMatrix(8, 11) = &H8B:  HexMatrix(8, 12) = &H8C:  HexMatrix(8, 13) = &H8D:  HexMatrix(8, 14) = &H8E:  HexMatrix(8, 15) = &H8F
    HexMatrix(9, 0) = &H90:   HexMatrix(9, 1) = &H91:   HexMatrix(9, 2) = &H92:   HexMatrix(9, 3) = &H93:   HexMatrix(9, 4) = &H94:   HexMatrix(9, 5) = &H95:   HexMatrix(9, 6) = &H96:   HexMatrix(9, 7) = &H97
    HexMatrix(9, 8) = &H98:   HexMatrix(9, 9) = &H99:   HexMatrix(9, 10) = &H9A:  HexMatrix(9, 11) = &H9B:  HexMatrix(9, 12) = &H9C:  HexMatrix(9, 13) = &H9D:  HexMatrix(9, 14) = &H9E:  HexMatrix(9, 15) = &H9F
    HexMatrix(10, 0) = &HA0:  HexMatrix(10, 1) = &HA1:  HexMatrix(10, 2) = &HA2:  HexMatrix(10, 3) = &HA3:  HexMatrix(10, 4) = &HA4:  HexMatrix(10, 5) = &HA5:  HexMatrix(10, 6) = &HA6:  HexMatrix(10, 7) = &HA7
    HexMatrix(10, 8) = &HA8:  HexMatrix(10, 9) = &HA9:  HexMatrix(10, 10) = &HAA: HexMatrix(10, 11) = &HAB: HexMatrix(10, 12) = &HAC: HexMatrix(10, 13) = &HAD: HexMatrix(10, 14) = &HAE: HexMatrix(10, 15) = &HAF
    HexMatrix(11, 0) = &HB0:  HexMatrix(11, 1) = &HB1:  HexMatrix(11, 2) = &HB2:  HexMatrix(11, 3) = &HB3:  HexMatrix(11, 4) = &HB4:  HexMatrix(11, 5) = &HB5:  HexMatrix(11, 6) = &HB6:  HexMatrix(11, 7) = &HB7
    HexMatrix(11, 8) = &HB8:  HexMatrix(11, 9) = &HB9:  HexMatrix(11, 10) = &HBA: HexMatrix(11, 11) = &HBB: HexMatrix(11, 12) = &HBC: HexMatrix(11, 13) = &HBD: HexMatrix(11, 14) = &HBE: HexMatrix(11, 15) = &HBF
    HexMatrix(12, 0) = &HC0:  HexMatrix(12, 1) = &HC1:  HexMatrix(12, 2) = &HC2:  HexMatrix(12, 3) = &HC3:  HexMatrix(12, 4) = &HC4:  HexMatrix(12, 5) = &HC5:  HexMatrix(12, 6) = &HC6:  HexMatrix(12, 7) = &HC7
    HexMatrix(12, 8) = &HC8:  HexMatrix(12, 9) = &HC9:  HexMatrix(12, 10) = &HCA: HexMatrix(12, 11) = &HCB: HexMatrix(12, 12) = &HCC: HexMatrix(12, 13) = &HCD: HexMatrix(12, 14) = &HCE: HexMatrix(12, 15) = &HCF
    HexMatrix(13, 0) = &HD0:  HexMatrix(13, 1) = &HD1:  HexMatrix(13, 2) = &HD2:  HexMatrix(13, 3) = &HD3:  HexMatrix(13, 4) = &HD4:  HexMatrix(13, 5) = &HD5:  HexMatrix(13, 6) = &HD6:  HexMatrix(13, 7) = &HD7
    HexMatrix(13, 8) = &HD8:  HexMatrix(13, 9) = &HD9:  HexMatrix(13, 10) = &HDA: HexMatrix(13, 11) = &HDB: HexMatrix(13, 12) = &HDC: HexMatrix(13, 13) = &HDD: HexMatrix(13, 14) = &HDE: HexMatrix(13, 15) = &HDF
    HexMatrix(14, 0) = &HE0:  HexMatrix(14, 1) = &HE1:  HexMatrix(14, 2) = &HE2:  HexMatrix(14, 3) = &HE3:  HexMatrix(14, 4) = &HE4:  HexMatrix(14, 5) = &HE5:  HexMatrix(14, 6) = &HE6:  HexMatrix(14, 7) = &HE7
    HexMatrix(14, 8) = &HE8:  HexMatrix(14, 9) = &HE9:  HexMatrix(14, 10) = &HEA: HexMatrix(14, 11) = &HEB: HexMatrix(14, 12) = &HEC: HexMatrix(14, 13) = &HED: HexMatrix(14, 14) = &HEE: HexMatrix(14, 15) = &HEF
    HexMatrix(15, 0) = &HF0:  HexMatrix(15, 1) = &HF1:  HexMatrix(15, 2) = &HF2:  HexMatrix(15, 3) = &HF3:  HexMatrix(15, 4) = &HF4:  HexMatrix(15, 5) = &HF5:  HexMatrix(15, 6) = &HF6:  HexMatrix(15, 7) = &HF7
    HexMatrix(15, 8) = &HF8:  HexMatrix(15, 9) = &HF9:  HexMatrix(15, 10) = &HFA: HexMatrix(15, 11) = &HFB: HexMatrix(15, 12) = &HFC: HexMatrix(15, 13) = &HFD: HexMatrix(15, 14) = &HFE: HexMatrix(15, 15) = &HFF
End Sub

''''
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
'--------------------------------------------------------------------------

Public Sub EnabledControl(ByRef frmEN As Form, Optional ByVal blnEN As Boolean = True)
    Dim ctlEn As VB.Control
    On Error Resume Next
    For Each ctlEn In frmEN.Controls
        ctlEn.Enabled = blnEN
    Next
    Screen.MousePointer = IIf(blnEN, 0, 13)
End Sub

Public Function FileBackup(ByVal strFolderSrc As String, ByVal strFolderDes As String, _
    Optional ByVal blnEncrypt As Boolean = True, _
    Optional ByVal strKey As String = mconStrKey) As Boolean
    '备份
    Dim strFS As String, strFD As String
    Dim strFBK As String, strFind As String, strDir As String, strGet As String, strPre As String
    Dim bytFile() As Byte, intSrc As Integer, intDes As Integer, lngSize As Long, bytEncrypt() As Byte
    Dim strFR As String, intFR As Integer, strSize As String
    
    If Not (FolderExist(strFolderSrc) And FolderExist(strFolderDes)) Then
        MsgBox "备份的源路径或目的路径不存在", vbCritical, "警告"
        Exit Function
    End If
    If LCase(strFolderSrc) = LCase(strFolderDes) Then
        MsgBox "备份的源路径或目的路径不能相同", vbCritical, "警告"
        Exit Function
    End If
    
    On Error GoTo LineERR
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
        Get #intSrc, , bytFile '不加密时单个文件大约大于380MB时内存溢出报错
        If blnEncrypt Then
            bytEncrypt = EncryptByte(bytFile, strKey)   '加密时单个文件大约大于185MB时内存溢出报错
            bytFile = bytEncrypt    '加密前后Byte数组的维数与大小相等
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
    FileBackup = True
    
LineERR:
    Close   '关闭所有
    If Err.Number Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
        If FileExist(strFBK) Then Kill strFBK '删除备份文件
        If FileExist(strFR) Then Kill strFR
    End If
End Function

Public Function FileCompress(ByVal strSrcFolder As String, ByVal strDesFolder As String, _
            Optional ByVal MSize As Long = 50) As Boolean
    '压缩文件
    Dim strWinRAR As String, strSrc As String, strDes As String, strSize As String, strCommand As String
    
    strWinRAR = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "WinRAR.exe"
    If Not FileExist(strWinRAR) Then  '压缩程序是否存在
        MsgBox "WinRAR压缩应用程序不存在", vbExclamation, "警告"
        Exit Function
    End If
    If Not FolderExist(strSrcFolder) Then '源文件与目的文件是否存在
        MsgBox "被压缩的文件目录不存在", vbExclamation, "警告"
        Exit Function
    End If
    If Not FolderExist(strDesFolder) Then
        MsgBox "保存压缩文件的目录不存在", vbExclamation, "警告"
        Exit Function
    End If
    If FolderNotNull(strSrcFolder) = 0 Then '源目录是否为空
        MsgBox "被压缩的文件目录无可压缩文件", vbExclamation, "提醒"
        Exit Function
    End If
    
    strSrc = IIf(Right(strSrcFolder, 1) = "\", strSrcFolder, strSrcFolder & "\")    '样式: D:\temp\，'\'有重要意义
    strDes = IIf(Right(strDesFolder, 1) = "\", strDesFolder, strDesFolder & "\") & "FC_" & Format(Now, "yyyy_MM_DD_HH_mm_ss") & ".rar"
    If MSize <= 0 Then MSize = 50
    strSize = "-v" & MSize & "M"
    '生成压缩shell命令。'-k锁定文件，-v50M 以50M分卷，-r 连同子文件夹，-ep1 路径中不包含顶层文件夹
    strCommand = strWinRAR & " a " & strSize & " -s -k -r -ep1 " & strDes & " " & strSrc
    If ShellWait(strCommand) Then
        FileCompress = True '但是如果压缩过程被中断取消也是返回True的
    End If
    
End Function

Public Function FileExist(ByVal strPath As String, Optional ByVal blnSetNormal As Boolean = False) As Boolean
    '判断 [文件] 是否存在
    Dim strDir As String, strMid As String
    
    On Error Resume Next
    strDir = Dir(strPath, vbHidden) '若存在则返回不带路径的文件名
    If Len(strDir) > 0 Then
        strMid = Mid(strPath, InStrRev(strPath, "\") + 1)   '获得不包含文件夹路径的文件名
        If LCase(strMid) = LCase(strDir) Then   '相等则表示文件存在
            If blnSetNormal Then    '若要强制改变文件属性，如删除只读取或隐藏属性
                If GetAttr(strPath) <> vbNormal Then
                    SetAttr strPath, vbNormal   '去除多余的属性，强制改成常规属性文件
                End If
            End If
            FileExist = True
        End If
    End If
    If Err.Number Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
    End If
End Function

Public Function FileExtract(ByVal strSrcFile As String, ByVal strDesFolder As String) As Boolean
    '解压文件
    
    Dim strWinRAR As String, strSrc As String, strDes As String, strCommand As String
    
    strWinRAR = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "WinRAR.exe"
    If Not FileExist(strWinRAR) Then  '压缩程序是否存在
        MsgBox "WinRAR压缩应用程序不存在", vbExclamation, "警告"
        Exit Function
    End If
    If Not FileExist(strSrcFile) Then '源文件与目的文件是否存在
        MsgBox "被解压的文件不存在", vbExclamation, "警告"
        Exit Function
    End If
    If Not FolderExist(strDesFolder) Then
        MsgBox "解压后的文件存放目录不存在", vbExclamation, "警告"
        Exit Function
    End If
        
    strSrc = strSrcFile
    strDes = IIf(Right(strDesFolder, 1) = "\", strDesFolder, strDesFolder & "\")
    
    '生成压缩shell命令
    strCommand = strWinRAR & " x " & strSrc & " " & strDes
    If ShellWait(strCommand) Then
        FileExtract = True
    End If
    
End Function

Public Function FileRestore(ByVal strFileSrc As String, ByVal strFolderDes As String, _
    Optional ByVal blnDecrypt As Boolean = True, _
    Optional ByVal strKey As String = mconStrKey) As Boolean
    '还原
    Dim strFS As String, strFD As String
    Dim strFI As String, strLine As String, strArr() As String, strFBK As String, bytDecrypt() As Byte
    Dim bytFile() As Byte, intFI As Integer, intSrc As Integer, intDes As Integer, lngSize As Long

    strFI = Left(strFileSrc, InStrRev(strFileSrc, ".") - 1) & mconStrBKrst '恢复文件的配置文件
    If Not (FolderExist(strFolderDes) And FileExist(strFileSrc) And FileExist(strFI)) Then
        MsgBox "还原的源文件或还原位置不存在", vbCritical, "警告"
        Exit Function
    End If
    If LCase(Mid(strFileSrc, InStrRev(strFileSrc, "."))) <> LCase(mconStrBKbak) Then
        MsgBox "还原的源文件格式不对", vbCritical, "警告"
        Exit Function
    End If
    
    On Error GoTo LineERR
    strFS = Left(strFileSrc, InStrRev(strFileSrc, "\"))
    strFD = IIf(Right(strFolderDes, 1) = "\", strFolderDes, strFolderDes & "\")
    
    intFI = FreeFile
    Open strFI For Input As #intFI
    intSrc = FreeFile
    Open strFileSrc For Binary As #intSrc
    While Not EOF(intFI)
        Line Input #intFI, strLine
        strArr = Split(strLine, vbTab)
        If UBound(strArr) <> 1 Then GoTo LineERR
        If blnDecrypt Then
            strArr(0) = DecryptString(strArr(0), strKey)
            strArr(1) = DecryptString(strArr(1), strKey)
        End If
        If Not IsNumeric(strArr(1)) Then GoTo LineERR
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
    FileRestore = True
    
LineERR:
    Close   '关闭所有
    If Err.Number Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
    End If
End Function

Public Function FolderExist(ByVal strPath As String, Optional ByVal blnSetNormal As Boolean = False) As Boolean
    '判断 [文件夹] 是否存在
    Dim strFod As String, strGet As String
    
    On Error Resume Next
    If Len(Trim(strPath)) > 0 Then  '以防传入空字符串路径
        If Right(strPath, 1) = "\" Then
            If InStr(strPath, "\") <> InStrRev(strPath, "\") Then   '以防传入的是根目录
                strPath = Left(strPath, Len(strPath) - 1) '非根目录则踢除末尾多余的"\"
            End If
        End If
    End If
    strFod = Dir(strPath, vbDirectory + vbHidden)
    If Len(strFod) > 0 Then '说明有返回值
        If strFod <> "." And strFod <> ".." Then    '若是空路径则返回"."
            If InStr(strPath, "\") = InStrRev(strPath, "\") Then    '以防传入的是根目录如"D:\"
                strGet = strPath
            Else
                strGet = Left(strPath, Len(strPath) - Len(strFod)) & strFod '正常情况下strFod值+上层目录=strPath
            End If
            If GetAttr(strGet) And vbDirectory = vbDirectory Then   '如果是文件夹或者存在的根目录
                If blnSetNormal Then    '若要强制改变文件属性，如删除只读取或隐藏属性
                    If GetAttr(strPath) <> vbNormal Then
                        SetAttr strPath, vbNormal   '去除多余的属性，强制改成常规属性文件
                    End If
                End If
                FolderExist = True
            End If
        End If
    End If
    If Err.Number Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
    End If
End Function

Public Function FolderNotNull(ByVal strFolderPath As String) As Boolean
    '检查文件夹是否为空目录
    Dim objFSO As Object, objFolder As Object, objFiles As Object
    
    If Not FolderExist(strFolderPath) Then Exit Function
    
    On Error Resume Next
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(strFolderPath)
    Set objFiles = objFolder.Files
    If objFiles.Count > 0 Then  '文件个数
        FolderNotNull = True
    ElseIf objFolder.SubFolders.Count > 0 Then  '文件夹个数
        FolderNotNull = True
    End If
    
    Set objFiles = Nothing
    Set objFolder = Nothing
    Set objFSO = Nothing

    If Err.Number Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
    End If
End Function

Public Function ShellWait(ByVal strShellCommand As String) As Boolean
    '等待Shell命令执行完成后再执行后面的代码，间接阻止Shell的异步执行.
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
    
    pidNotePad = Shell(sExecString, vbHide) '返回执行程序的任务ID，不成功返回0
    hProcess = OpenProcess(Process_query_infomation, True, pidNotePad)  '打开进程
    Do
        GetExitCodeProcess hProcess, lExitCode  '获取进程中断退出代码
        DoEvents
    Loop While lExitCode = Still_Active
    CloseHandle (pidNotePad)    '关闭进程
    
    If Err.Number Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "警告"
    Else
        ShellWait = True
    End If
End Function
