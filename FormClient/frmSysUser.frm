VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmSysUser 
   Caption         =   "用户管理"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16185
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   16185
   Begin VB.HScrollBar Hsb 
      Height          =   255
      Left            =   14520
      TabIndex        =   23
      Top             =   6720
      Width           =   1455
   End
   Begin VB.VScrollBar Vsb 
      Height          =   1935
      Left            =   15600
      TabIndex        =   22
      Top             =   4680
      Width           =   255
   End
   Begin VB.Frame ctlMove 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6855
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   15615
      Begin VB.Frame Frame1 
         Caption         =   "用户角色指定"
         ForeColor       =   &H00FF0000&
         Height          =   6375
         Index           =   1
         Left            =   7680
         TabIndex        =   21
         Top             =   0
         Width           =   7815
         Begin VB.Frame Frame1 
            Caption         =   "用户头像"
            ForeColor       =   &H00FF00FF&
            Height          =   2895
            Index           =   5
            Left            =   4000
            TabIndex        =   32
            Top             =   2880
            Width           =   3720
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   120
               Top             =   2160
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton Command4 
               Caption         =   "选择照片"
               Height          =   495
               Left            =   720
               TabIndex        =   36
               Top             =   2160
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "*格式为jpg|png|bmp *文件小于5MB"
               ForeColor       =   &H000000FF&
               Height          =   420
               Index           =   32
               Left            =   1845
               TabIndex        =   37
               Top             =   2200
               Width           =   1710
            End
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1815
               Left            =   600
               Top             =   240
               Width           =   1305
            End
         End
         Begin VB.Frame Frame1 
            Height          =   2415
            Index           =   4
            Left            =   4000
            TabIndex        =   31
            Top             =   150
            Width           =   3720
            Begin VB.TextBox Text1 
               BackColor       =   &H80000003&
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   330
               Index           =   5
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   34
               Text            =   "Text1"
               Top             =   360
               Width           =   2500
            End
            Begin VB.CommandButton Command3 
               Caption         =   "用户角色指定结果保存"
               Height          =   495
               Left            =   720
               TabIndex        =   33
               Top             =   1560
               Width           =   2415
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "已选用户"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   31
               Left            =   120
               TabIndex        =   35
               Top             =   390
               Width           =   900
            End
         End
         Begin MSComctlLib.TreeView TreeView2 
            Height          =   4095
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   7223
            _Version        =   393217
            Indentation     =   441
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "用户"
         ForeColor       =   &H00FF0000&
         Height          =   6375
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   7455
         Begin VB.Frame Frame1 
            Height          =   405
            Index           =   3
            Left            =   675
            TabIndex        =   27
            Top             =   4080
            Width           =   2500
            Begin VB.OptionButton Option1 
               Caption         =   "启用"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   210
               Index           =   2
               Left            =   120
               TabIndex        =   29
               Top             =   150
               Width           =   855
            End
            Begin VB.OptionButton Option1 
               Caption         =   "停用"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   3
               Left            =   1200
               TabIndex        =   28
               Top             =   150
               Width           =   855
            End
         End
         Begin VB.Frame Frame1 
            Height          =   405
            Index           =   2
            Left            =   720
            TabIndex        =   24
            Top             =   2100
            Width           =   2500
            Begin VB.OptionButton Option1 
               Caption         =   "男"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   1
               Left            =   1200
               TabIndex        =   26
               Top             =   150
               Width           =   855
            End
            Begin VB.OptionButton Option1 
               Caption         =   "女"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   210
               Index           =   0
               Left            =   120
               TabIndex        =   25
               Top             =   150
               Width           =   855
            End
         End
         Begin VB.CommandButton Command2 
            Caption         =   "修改用户信息"
            Height          =   495
            Left            =   1800
            TabIndex        =   7
            Top             =   4800
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "添加用户"
            Height          =   495
            Left            =   240
            TabIndex        =   6
            Top             =   4800
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   720
            TabIndex        =   1
            Text            =   "Text2"
            Top             =   720
            Width           =   2500
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   1
            Left            =   1320
            TabIndex        =   12
            Text            =   "Combo2"
            Top             =   3240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000003&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   240
            Width           =   2500
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2640
            Width           =   2500
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   720
            PasswordChar    =   "*"
            TabIndex        =   2
            Text            =   "Text2"
            Top             =   1200
            Width           =   2500
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   720
            TabIndex        =   3
            Text            =   "Text2"
            Top             =   1680
            Width           =   2500
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   4
            Left            =   720
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   3120
            Width           =   2500
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   4095
            Left            =   3480
            TabIndex        =   8
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   7223
            _Version        =   393217
            Indentation     =   441
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "状态"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   120
            TabIndex        =   30
            Top             =   4200
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "密码"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   165
            TabIndex        =   20
            Top             =   1260
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "账号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   165
            TabIndex        =   19
            Top             =   780
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "标识"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   165
            TabIndex        =   18
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "姓名"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   165
            TabIndex        =   17
            Top             =   1740
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "性别"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   165
            TabIndex        =   16
            Top             =   2220
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "部门"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   165
            TabIndex        =   15
            Top             =   2700
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "备注"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   165
            TabIndex        =   14
            Top             =   3180
            Width           =   450
         End
         Begin VB.Label Label1 
            Caption         =   "*** 密码只能包含数字或大小字母，且长度在20位以内"
            ForeColor       =   &H000000FF&
            Height          =   420
            Index           =   21
            Left            =   240
            TabIndex        =   13
            Top             =   5640
            Width           =   3060
         End
      End
   End
End
Attribute VB_Name = "frmSysUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlngID As Long
Dim mstrNewPhotoName As String
Private Const mKeyDept As String = "k"
Private Const mKeyUser As String = "u"
Private Const mOtherKey As String = "kOther"
Private Const mOtherText As String = "其他人员"
Private Const mKeyRole As String = "r"
Private Const mOtherKeyRole As String = "kOtherRole"
Private Const mOtherTextRole As String = "其他角色"
Private Const mTwoBar As String = "--"


Private Function mfSavePhoto(Optional ByVal blnNew As Boolean = True) As Boolean
    '上传个人头像照片至服务器。成功返回新文件名
    Dim strPath As String, strOldName As String, strNewName As String, strExtension As String
    Dim strSaveFolder As String, strClassify As String, strSQL As String
    Dim lngFileLen As Long, K As Long
    
    strPath = Trim(CommonDialog1.Tag) '获取App目录下的文件路径
    If Not gfFileExist(strPath) Then    '不存在则使用原文件路径
        strPath = Trim(CommonDialog1.FileName) '原路径文件
        If Not gfFileExist(strPath) Then    '原文件也不存在了
            MsgBox "头像照片的原文件已丢失，不进行保存！", vbCritical, "照片警报"
            Exit Function
        End If
    End If
    
    strOldName = Mid(strPath, InStrRev(strPath, "\") + 1)   '获取旧文件名
    If Len(strOldName) > 50 Then
        MsgBox "照片文件名不能超过50个字符 ，请重命名！", vbExclamation, "旧文件名警告"
        Exit Function
    End If
    strExtension = Mid(strOldName, InStrRev(strOldName, ".") + 1)   '获取旧文件的扩展名
    If Len(strExtension) > 20 Then
        MsgBox "照片文件的扩展名不能超过20个字符 ，请重命名！", vbExclamation, "扩展名警告"
        Exit Function
    End If
    lngFileLen = FileLen(strOldName)    '获取文件大小
    If lngFileLen > 5242880 Then
        MsgBox "照片大小不能超过5MB！", vbExclamation, "文件大小警告"
        Exit Function
    End If
    For K = 1 To 30
        strNewName = strNewName & gfBackOneChar(udUpperCase)    '获取新文件名（30个随机大写字母）
    Next
    strNewName = strNewName & "." & strExtension    '新文件名加上扩展名
    strSaveFolder = gVar.FolderStore    '服务端存储位置。注意不含路径，仅文件夹名称
    strClassify = "头像照片"            '文件存储类别
    
    With gArr(gWind.Winsock1.Item(1).Index)
        .FilePath = strPath         '本机欲发送的文件路径
        .FileSizeTotal = lngFileLen '本机欲发送的文件大小或服务端接收的文件大小
        .FileFolder = strSaveFolder '本机发送过去的文件在服务端保存的文件夹名称
        .FileName = strNewName      '本机发送过去的文件在服务端保存的文件名
        gVar.FTIsOver = False
    End With
    If gWind.Winsock1.Item(1).State = 7 Then    '使用MDI窗体上的Winsock控件发送文件信息
        If gfSendInfo(gfFileInfoJoin(gWind.Winsock1.Item(1).Index, ftSend), gWind.Winsock1.Item(1)) Then
            Call gsFileProgress(gWind.CommandBars1.StatusBar.FindPane(gID.StatusBarPaneProgress), _
                                gWind.CommandBars1.StatusBar.FindPane(gID.StatusBarPaneProgressText), _
                                ftZero, 0, lngFileLen, 0) '初始化进度条
            Debug.Print "Client：已发送[头像照片]的文件信息给服务端," & Now
            mfSavePhoto = True
            mstrNewPhotoName = strNewName
        End If
    End If
    
End Function

Private Sub msLoadDept(ByRef tvwDept As MSComctlLib.TreeView)
    '加载部门至TreeView控件中
    '要求：1、数据库中部门信息表Dept包含DeptID(Not Null)、DeptName(Not Null)、ParentID(Null)三个字段。
    '要求：2、部门表中只允许顶级部门（锁定为公司名称）的ParentID为Null，其它部门的ParentID都不能为Null。
    
    Dim rsDept As ADODB.Recordset
    Dim strSQL As String
    Dim arrDept() As String '注意下标要从0开始
    Dim I As Long, lngCount As Long, lngOneCompany As Long
    Dim blnLoop As Boolean
        
    strSQL = "SELECT t1.DeptID ,t1.DeptName ,t1.ParentID ,t2.DeptName AS [ParentName] " & _
             "FROM tb_FT_Sys_Department AS [t1] " & _
             "LEFT JOIN tb_FT_Sys_Department AS [t2] " & _
             "ON t1.ParentID = t2.DeptID " & _
             "ORDER BY t1.ParentID ,t1.DeptName"    '注意字段顺序不可变
    Set rsDept = gfBackRecordset(strSQL)
    If rsDept.State = adStateClosed Then Exit Sub
    If rsDept.RecordCount > 0 Then
        
        tvwDept.Nodes.Clear
        Combo1.Item(0).Clear
        Combo1.Item(1).Clear
        
        While Not rsDept.EOF
            If IsNull(rsDept.Fields(3).Value) Then
                lngOneCompany = lngOneCompany + 1
                tvwDept.Nodes.Add , , mKeyDept & rsDept.Fields(0).Value, rsDept.Fields(1).Value, "SysCompany"
                tvwDept.Nodes.Item(mKeyDept & rsDept.Fields(0).Value).Expanded = True
            Else
                ReDim Preserve arrDept(3, lngCount)
                For I = 0 To 3
                    arrDept(I, lngCount) = rsDept.Fields(I).Value
                Next
                lngCount = lngCount + 1
                blnLoop = True
            End If
            
            Combo1.Item(0).AddItem rsDept.Fields(1).Value
            Combo1.Item(1).AddItem rsDept.Fields(0).Value
            
            rsDept.MoveNext
        Wend
        
    End If
    rsDept.Close
    Set rsDept = Nothing
    
    If blnLoop Then Call msLoadDeptTree(tvwDept, arrDept)

End Sub

Private Sub msLoadDeptTree(ByRef tvwTree As MSComctlLib.TreeView, ByRef arrLoad() As String)
    '必须与msLoadDept过程配合使用来加载部门列表
    
    Dim arrOther() As String    '保存剩余的
    Dim blnOther As Boolean     '剩余标识
    Dim I As Long, J As Long, K As Long, lngCount As Long
    Static C As Long
    
    With tvwTree
        For J = LBound(arrLoad, 2) To UBound(arrLoad, 2)
            For I = 1 To .Nodes.Count   '注意此处下标从1开始
                If .Nodes.Item(I).Key = mKeyDept & arrLoad(2, J) Then
                    .Nodes.Add .Nodes.Item(I).Key, tvwChild, mKeyDept & arrLoad(0, J), arrLoad(1, J), "threemen"
                    .Nodes.Item(mKeyDept & arrLoad(0, J)).Expanded = True
                    Exit For
                End If
            Next
            
            If I = .Nodes.Count + 1 Then
                blnOther = True
                ReDim Preserve arrOther(3, lngCount)
                For K = 0 To 3
                    arrOther(K, lngCount) = arrLoad(K, J)
                Next
                lngCount = lngCount + 1
            End If
            
        Next
    End With
    
    C = C + 1
    If C > 64 Then Exit Sub '防止递归层数太深导致堆栈溢出而程序崩溃
    
    If blnOther Then
        Call msLoadDeptTree(tvwTree, arrOther)
    End If

End Sub

Private Sub msLoadRole(ByRef tvwUser As MSComctlLib.TreeView)
    '加载角色，前提是已加载好部门
    
    Dim strSQL As String
    Dim rsRole As ADODB.Recordset
    Dim arrOther() As String    '保存剩余的
    Dim blnOther As Boolean     '剩余标识
    Dim I As Long, J As Long, K As Long, lngCount As Long
    
    If tvwUser.Nodes.Count = 0 Then Exit Sub
    
    strSQL = "SELECT RoleAutoID ,RoleName ,DeptID FROM tb_FT_Sys_Role "
    Set rsRole = gfBackRecordset(strSQL)
    If rsRole.State = adStateClosed Then GoTo LineEnd
    If rsRole.RecordCount = 0 Then GoTo LineEnd
    
    With tvwUser
        While Not rsRole.EOF
            For I = 1 To .Nodes.Count
                If .Nodes(I).Key = mKeyDept & rsRole.Fields("DeptID").Value Then
                    .Nodes.Add .Nodes.Item(I).Key, tvwChild, _
                        mKeyRole & rsRole.Fields("RoleAutoID").Value, rsRole.Fields("RoleName").Value, "SysRole"
                    Exit For
                End If
            Next
            
            If I = .Nodes.Count + 1 Then
                blnOther = True
                ReDim Preserve arrOther(2, lngCount)
                For K = 0 To 2
                    arrOther(K, lngCount) = rsRole.Fields(K).Value & ""
                Next
                lngCount = lngCount + 1
            End If
            
            rsRole.MoveNext
        Wend
        
        If blnOther Then
            .Nodes.Add 1, tvwChild, mOtherKey, mOtherText, "unknown"
            .Nodes(mOtherKey).Expanded = True
            For I = LBound(arrOther, 2) To UBound(arrOther, 2)
                .Nodes.Add mOtherKey, tvwChild, mKeyRole & arrOther(0, I), _
                    arrOther(1, I), "SysRole"
            Next
        End If

    End With
    
LineEnd:
    If rsRole.State = adStateOpen Then rsRole.Close
    Set rsRole = Nothing
    
End Sub

Private Sub msLoadUser(ByRef tvwUser As MSComctlLib.TreeView)
    '加载用户，前提是已加载好部门
    
    Dim strSQL As String
    Dim rsUser As ADODB.Recordset
    Dim arrOther() As String    '保存剩余的
    Dim blnOther As Boolean     '剩余标识
    Dim I As Long, J As Long, K As Long, lngCount As Long
    
    If tvwUser.Nodes.Count = 0 Then Exit Sub
    
    strSQL = "SELECT UserAutoID ,UserFullName ,UserSex ,DeptID FROM tb_FT_Sys_User " & _
             "WHERE UserLoginName <>'" & gVar.AccountAdmin & "' AND UserLoginName <>'" & gVar.AccountSystem & "' "
    Set rsUser = gfBackRecordset(strSQL)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount = 0 Then GoTo LineEnd

    With tvwUser
        While Not rsUser.EOF
            For I = 1 To .Nodes.Count
                If .Nodes(I).Key = mKeyDept & rsUser.Fields("DeptID").Value Then
                    .Nodes.Add .Nodes.Item(I).Key, tvwChild, _
                        mKeyUser & rsUser.Fields("UserAutoID").Value, rsUser.Fields("UserFullName").Value, _
                        IIf(rsUser.Fields("UserSex") = "男", "man", "woman")
                    Exit For
                End If
            Next
            
            If I = .Nodes.Count + 1 Then
                blnOther = True
                ReDim Preserve arrOther(3, lngCount)
                For K = 0 To 3
                    arrOther(K, lngCount) = rsUser.Fields(K).Value & ""
                Next
                lngCount = lngCount + 1
            End If
            
            rsUser.MoveNext
        Wend
        
        If blnOther Then
            .Nodes.Add 1, tvwChild, mOtherKey, mOtherText, "unknown"
            .Nodes(mOtherKey).Expanded = True
            For I = LBound(arrOther, 2) To UBound(arrOther, 2)
                .Nodes.Add mOtherKey, tvwChild, mKeyUser & arrOther(0, I), _
                    arrOther(1, I), IIf(arrOther(2, I) = "男", "man", "woman")
            Next
        End If

    End With
    
LineEnd:
    If rsUser.State = adStateOpen Then rsUser.Close
    Set rsUser = Nothing
End Sub

Private Sub msLoadUserRole(ByVal strUID As String)
    
    Dim strSQL As String
    Dim rsUser As ADODB.Recordset
    Dim I As Long
    
    strSQL = "SELECT UserAutoID ,RoleAutoID FROM tb_FT_Sys_UserRole WHERE UserAutoID =" & strUID
    Set rsUser = gfBackRecordset(strSQL)
    If rsUser.State = adStateOpen Then
        With TreeView2.Nodes
            For I = 2 To .Count
                If Left(.Item(I).Key, Len(mKeyRole)) = mKeyRole Then
                    If rsUser.RecordCount > 0 Then rsUser.MoveFirst
                    Do While Not rsUser.EOF
                        If .Item(I).Key = mKeyRole & rsUser.Fields("RoleAutoID") Then
                            .Item(I).Checked = True
                            Exit Do
                        End If
                        rsUser.MoveNext
                    Loop
                    If rsUser.EOF Then .Item(I).Checked = False
                Else
                     .Item(I).Checked = False
                End If
            Next
        End With
        rsUser.Close
    End If
    
    Set rsUser = Nothing
    
End Sub


Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
            Combo1.Item(Index).ListIndex = -1
        End If
    End If
End Sub

Private Sub Command1_Click()
    '添加
    
    Dim strLoginName As String, strPWD As String, strFullName As String
    Dim strSex As String, strMemo As String, strState As String
    Dim strDept As Variant, strPhotoID As String
    Dim strSQL As String, strMsg As String
    Dim rsUser As ADODB.Recordset
    
    strLoginName = Trim(Text1.Item(1).Text)
    strPWD = Trim(Text1.Item(2).Text)
    strFullName = Trim(Text1.Item(3).Text)
    strMemo = Trim(Text1.Item(4).Text)
    
    strLoginName = Left(strLoginName, 50)
    strPWD = Left(strPWD, 20)
    strFullName = Left(strFullName, 50)
    strMemo = Left(strMemo, 500)
    
    Text1.Item(1).Text = strLoginName
    Text1.Item(2).Text = strPWD
    Text1.Item(3).Text = strFullName
    Text1.Item(4).Text = strMemo
    
    If Option1.Item(0).Value Then strSex = Option1.Item(0).Caption
    If Option1.Item(1).Value Then strSex = Option1.Item(1).Caption
    If Option1.Item(2).Value Then strState = Option1.Item(2).Caption
    If Option1.Item(3).Value Then strState = Option1.Item(3).Caption
    strDept = Combo1.Item(1).List(Combo1.Item(0).ListIndex)
    
    If Len(strLoginName) = 0 Then
        MsgBox Label1.Item(1).Caption & " 不能为空！", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(strLoginName)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strLoginName)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(1).Caption & " 不能含有特殊字符【" & strMsg & "】！", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(strLoginName)
        Exit Sub
    End If
    
    If Len(strPWD) = 0 Then
        MsgBox Label1.Item(2).Caption & " 不能为空！", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(strPWD)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strPWD)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(2).Caption & " 不能含有特殊字符【" & strMsg & "】！", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(strPWD)
        Exit Sub
    End If
    
    If Len(strFullName) = 0 Then
        MsgBox Label1.Item(3).Caption & " 不能为空！", vbExclamation
        Text1.Item(3).SetFocus
        Text1.Item(3).SelStart = 0
        Text1.Item(3).SelLength = Len(strFullName)
        Exit Sub
    End If
    
    If Len(strSex) = 0 Then
        Option1.Item(0).Value = True
        strSex = Option1.Item(0).Caption
    End If
    If Len(strState) = 0 Then
        Option1.Item(3).Value = True
        strState = Option1.Item(3).Caption
    End If
    
    If Len(strDept) = 0 Then strDept = Null
    
    If MsgBox("是否添加用户【" & strLoginName & "】【" & strFullName & "】？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    strSQL = "SELECT UserAutoID ,UserLoginName ,UserPassword ," & _
             "UserFullName ,UserSex ,UserState ,DeptID ,UserMemo ,FileID " & _
             "From tb_FT_Sys_User " & _
             "WHERE UserLoginName = '" & strLoginName & "'"
    Set rsUser = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount > 0 Then
        strMsg = "账号已存在，请更换！"
        GoTo LineBrk
    Else
        On Error GoTo LineErr
        
        rsUser.AddNew
        rsUser.Fields("UserLoginName") = strLoginName
        rsUser.Fields("UserPassword") = EncryptString(strPWD, gVar.EncryptKey)
        rsUser.Fields("UserFullName") = strFullName
        rsUser.Fields("UserSex") = strSex
        rsUser.Fields("UserState") = strState
        rsUser.Fields("DeptID") = strDept
        rsUser.Fields("UserMemo") = strMemo
        rsUser.Fields("FileID") = strPhotoID
        rsUser.Update
        strMsg = rsUser.Fields("UserAutoID").Value
        Text1.Item(0).Text = strMsg
        rsUser.Close
        strMsg = "添加用户【" & strMsg & "】【" & strLoginName & "】【" & strFullName & "】"
        Call gsLogAdd(Me, udInsert, "tb_FT_Sys_User", strMsg)
        MsgBox "用户【" & strLoginName & "】【" & strFullName & "】添加成功！", vbInformation
        Call msLoadDept(TreeView1)
        Call msLoadUser(TreeView1)
    End If
    
    GoTo LineEnd
    
LineBrk:
    rsUser.Close
    MsgBox strMsg, vbExclamation
    GoTo LineEnd
LineErr:
    Call gsAlarmAndLog("添加用户异常")
LineEnd:
    If rsUser.State = adStateOpen Then rsUser.Close
    Set rsUser = Nothing
End Sub

Private Sub Command2_Click()
    '修改
    
    Dim strUID As String, strLoginName As String, strPWD As String, strState As String
    Dim strFullName As String, strSex As String, strDept As String, strMemo As String
    Dim blnLoginName As Boolean, blnPwd As Boolean, blnFullName As Boolean
    Dim blnSex As Boolean, blnDept As Boolean, blnMemo As Boolean, blnState As Boolean
    Dim strSQL As String, strMsg As String
    Dim rsUser As ADODB.Recordset
    
    strUID = Trim(Text1.Item(0).Text)
    strLoginName = Trim(Text1.Item(1).Text)
    strPWD = Trim(Text1.Item(2).Text)
    strFullName = Trim(Text1.Item(3).Text)
    strMemo = Trim(Text1.Item(4).Text)
    
    strLoginName = Left(strLoginName, 50)
    strPWD = Left(strPWD, 20)
    strFullName = Left(strFullName, 50)
    strMemo = Left(strMemo, 500)
    
    Text1.Item(1).Text = strLoginName
    Text1.Item(2).Text = strPWD
    Text1.Item(3).Text = strFullName
    Text1.Item(4).Text = strMemo
    
    If Option1.Item(0).Value Then strSex = Option1.Item(0).Caption
    If Option1.Item(1).Value Then strSex = Option1.Item(1).Caption
    If Option1.Item(2).Value Then strState = Option1.Item(2).Caption
    If Option1.Item(3).Value Then strState = Option1.Item(3).Caption
    strDept = Combo1.Item(1).List(Combo1.Item(0).ListIndex)
    
    If Len(strLoginName) = 0 Then
        MsgBox Label1.Item(1).Caption & " 不能为空！", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(Text1.Item(1).Text)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strLoginName)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(1).Caption & " 不能含有特殊字符【" & strMsg & "】！", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(Text1.Item(1).Text)
        Exit Sub
    End If
    
    If Len(strPWD) = 0 Then
        MsgBox Label1.Item(2).Caption & " 不能为空！", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(Text1.Item(2).Text)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strPWD)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(2).Caption & " 不能含有特殊字符【" & strMsg & "】！", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(strPWD)
        Exit Sub
    End If
    
    If Len(strFullName) = 0 Then
        MsgBox Label1.Item(3).Caption & " 不能为空！", vbExclamation
        Text1.Item(3).SetFocus
        Text1.Item(3).SelStart = 0
        Text1.Item(3).SelLength = Len(Text1.Item(3).Text)
        Exit Sub
    End If
    
    If Len(strDept) = 0 Then
        MsgBox Label1.Item(5).Caption & " 不能为空！", vbExclamation
        Combo1.Item(0).SetFocus
        Exit Sub
    End If
    
    If Len(strSex) = 0 Then
        Option1.Item(0).Value = True
        strSex = Option1.Item(0).Caption
    End If
    If Len(strState) = 0 Then
        Option1.Item(3).Value = True
        strState = Option1.Item(3).Caption
    End If
    
    strSQL = "SELECT UserAutoID ,UserLoginName ,UserPassword ," & _
             "UserFullName ,UserSex ,UserState ,DeptID ,UserMemo " & _
             "From tb_FT_Sys_User " & _
             "WHERE UserAutoID = '" & strUID & "'"
    Set rsUser = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount = 0 Then
        strMsg = "该账号相关信息已丢失，请联系管理员！"
        GoTo LineBrk
    ElseIf rsUser.RecordCount > 1 Then
        strMsg = "该账号相关信息异常，请联系管理员！"
        GoTo LineBrk
    Else
        If strLoginName <> rsUser.Fields("UserLoginName") Then blnLoginName = True
        If IsNull(rsUser.Fields("UserPassword")) Or strPWD <> DecryptString(rsUser.Fields("UserPassword"), gVar.EncryptKey) Then blnPwd = True
        If IsNull(rsUser.Fields("UserFullName")) Or strFullName <> rsUser.Fields("UserFullName") Then blnFullName = True
        If IsNull(rsUser.Fields("UserSex")) Or strSex <> rsUser.Fields("UserSex") Then blnSex = True
        If IsNull(rsUser.Fields("DeptID")) Or strDept <> rsUser.Fields("DeptID") Then blnDept = True
        If IsNull(rsUser.Fields("UserMemo")) Or strMemo <> rsUser.Fields("UserMemo") Then blnMemo = True
        If IsNull(rsUser.Fields("UserState")) Or strState <> rsUser.Fields("UserState") Then blnState = True
        
        If Not (blnLoginName Or blnPwd Or blnFullName Or blnSex Or blnState Or blnDept Or blnMemo) Then
            strMsg = "没有实质性的改动，不进行修改！"
            GoTo LineBrk
        End If
        
        strMsg = "确定要修改" & Label1.Item(0).Caption & "【" & strUID & "】的用户信息吗？"
        If MsgBox(strMsg, vbQuestion + vbYesNo) = vbNo Then GoTo LineEnd
        
        On Error GoTo LineErr
        
        If blnLoginName Then rsUser.Fields("UserLoginName") = strLoginName
        If blnPwd Then rsUser.Fields("UserPassword") = EncryptString(strPWD, gVar.EncryptKey)
        If blnFullName Then rsUser.Fields("UserFullName") = strFullName
        If blnSex Then rsUser.Fields("UserSex") = strSex
        If blnDept Then rsUser.Fields("DeptID") = strDept
        If blnMemo Then rsUser.Fields("UserMemo") = strMemo
        If blnState Then rsUser.Fields("UserState") = strState
        
        rsUser.Update
        rsUser.Close
        
        strMsg = "修改ID【" & strUID & "】的"
        If blnLoginName Then strMsg = strMsg & "【" & Label1.Item(1).Caption & "】"
        If blnPwd Then strMsg = strMsg & "【" & Label1.Item(2).Caption & "】"
        If blnFullName Then strMsg = strMsg & "[" & Label1.Item(3).Caption & "】"
        If blnSex Then strMsg = strMsg & "【" & Label1.Item(4).Caption & "】"
        If blnDept Then strMsg = strMsg & "【" & Label1.Item(5).Caption & "】"
        If blnMemo Then strMsg = strMsg & "【" & Label1.Item(6).Caption & "】"
        If blnState Then strMsg = strMsg & "【" & Label1.Item(7).Caption & "】"
        Call gsLogAdd(Me, udUpdate, "tb_FT_Sys_User", strMsg)
        
        MsgBox "已成功" & strMsg & "。", vbInformation
        
        If blnFullName Or blnSex Or blnDept Then
            Call msLoadDept(TreeView1)
            Call msLoadUser(TreeView1)
        End If
        
    End If
    
    GoTo LineEnd
    
LineBrk:
    rsUser.Close
    MsgBox strMsg, vbExclamation
    GoTo LineEnd
LineErr:
    Call gsAlarmAndLog("用户信息修改异常")
LineEnd:
    If rsUser.State = adStateOpen Then rsUser.Close
    Set rsUser = Nothing
    
End Sub

Private Sub Command3_Click()
    '保存
    
    Dim strUID As String, strTemp As String, strMsg As String, strSQL As String
    Dim cnUser As ADODB.Connection
    Dim rsUser As ADODB.Recordset
    Dim blnTran As Boolean
    Dim I As Long
    
    If (TreeView1.Nodes.Count = 0) Or (TreeView2.Nodes.Count = 0) Then
        MsgBox "请先保证部门、用户、角色都已经设置好，再为用户指定角色！", vbExclamation
        Exit Sub
    End If
    strTemp = Trim(Text1.Item(5).Text)
    If (TreeView1.SelectedItem Is Nothing) Or (Len(strTemp) = 0) Then
        MsgBox "请先选择一个用户！", vbExclamation
        Exit Sub
    End If
    strUID = Left(strTemp, InStr(strTemp, mTwoBar) - 1)
    If Trim(strUID) <> Trim(Text1.Item(0).Text) Then
        MsgBox "用户检测异常，请重新选择一次用户！", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("确定对【" & strTemp & "】进行" & Command3.Caption & "吗？", vbQuestion + vbOKCancel, Command3.Caption & "询问") = vbCancel Then Exit Sub
    
    
    Set cnUser = New ADODB.Connection
    Set rsUser = New ADODB.Recordset
    cnUser.CursorLocation = adUseClient
    
    On Error GoTo LineErr
    
    cnUser.Open gVar.ConString
    strSQL = "DELETE FROM tb_FT_Sys_UserRole WHERE UserAutoID =" & strUID
    cnUser.BeginTrans
    blnTran = True
    cnUser.Execute strSQL
    
    strSQL = "SELECT UserAutoID ,RoleAutoID FROM tb_FT_Sys_UserRole WHERE UserAutoID =" & strUID
    rsUser.Open strSQL, cnUser, adOpenStatic, adLockBatchOptimistic
    If rsUser.RecordCount > 0 Then
        strMsg = strTemp & " 的后台数据检测异常，请重试或联系管理员！"
        GoTo LineErr
    Else
        With TreeView2.Nodes
            For I = 2 To .Count
                If Left(.Item(I).Key, Len(mKeyRole)) = mKeyRole Then
                    If .Item(I).Checked Then
                        rsUser.AddNew
                        rsUser.Fields("UserAutoID") = strUID
                        rsUser.Fields("RoleAutoID") = Right(.Item(I).Key, Len(.Item(I).Key) - Len(mKeyRole))
                    End If
                End If
            Next
        End With
        rsUser.UpdateBatch
        cnUser.CommitTrans
       
    End If
    
    rsUser.Close
    cnUser.Close
    Set rsUser = Nothing
    Set cnUser = Nothing
    
    Call gsLogAdd(Me, udInsertBatch, "tb_FT_Sys_UserRole", "为【" & strTemp & "】指定角色")
    MsgBox "【" & strTemp & "】" & Command3.Caption & " 成功！", vbInformation
    
    Exit Sub
    
LineErr:
    If blnTran Then cnUser.RollbackTrans
    If rsUser.State = adStateOpen Then rsUser.Close
    If cnUser.State = adStateOpen Then cnUser.Close
    Set rsUser = Nothing
    Set cnUser = Nothing
    If Len(strMsg) = 0 Then
        Call gsAlarmAndLog(Command3.Caption & "异常")
    Else
        MsgBox strMsg, vbExclamation
    End If
    
End Sub

Private Sub Command4_Click()
    '选择照片
    Dim strPhoto As String, strCopy As String
    Dim lngFiveMB As Long
    
    On Error GoTo LineErr
    
    Image1.Stretch = True
    With CommonDialog1
        .DialogTitle = "照片选择"
        .Filter = "图片(*.jpg;*.png;*.bmp)|*.jpg;*.png;*.bmp;*.gif"
        .Flags = cdlOFNCreatePrompt '当文件不存在时对话框要提示创建文件
        .ShowOpen   '弹出打开窗口
        strPhoto = .FileName    '将所选图片路径放到变量中
    End With
    
    If gfFileRepair(strPhoto) Then  '判断路径是否有效
        lngFiveMB = 5242880 ' 5 * 1024 * 1024
        If FileLen(strPhoto) > lngFiveMB Then   '文件不能大于5MB
            MsgBox "您选择的图片大于5MB了！", vbExclamation, "文件警告"
            CommonDialog1.FileName = ""
            Exit Sub
        End If
        Image1.Picture = LoadPicture(strPhoto)  '加载图片
        strCopy = gVar.FolderNameData & Mid(strPhoto, InStrRev(strPhoto, "\") + 1)  '图片复制一份的路径
        If gfFileRepair(strCopy) Then   '判断路径是否有效
            FileCopy strPhoto, strCopy  '将图片复制一份
            CommonDialog1.Tag = strCopy '路径保存备用
        End If
    End If
    Call mfSavePhoto(True)
    
LineErr:
    If Err.Number > 0 Then  '有异常发生
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "图片加载异常"
        CommonDialog1.FileName = "" '清空无效图片的路径
        CommonDialog1.Tag = ""
    End If
End Sub

Private Sub Form_Load()

    Set Me.Icon = gWind.ImageList1.ListImages("SysUser").Picture
    Me.Caption = gWind.CommandBars1.Actions(gID.SysAuthUser).Caption
    Frame1.Item(0).Caption = Me.Caption
    Command4.ToolTipText = "照片请通过【" & Command1.Caption & "】【" & Command2.Caption & "】按钮进行保存"
    
    For mlngID = Text1.LBound To Text1.UBound
        Text1.Item(mlngID).Text = ""
    Next
    
    Text1.Item(2).ToolTipText = "密码只能包含数字或大小字母，且长度在20位以内"
    
    TreeView1.Nodes.Clear
    TreeView1.ImageList = gWind.ImageList1
    TreeView2.Nodes.Clear
    TreeView2.ImageList = gWind.ImageList1
    
    Call msLoadDept(TreeView1)  '部门先加载
    Call msLoadUser(TreeView1)  '人员后加载
    
    Call msLoadDept(TreeView2)
    Call msLoadRole(TreeView2)
    
    Call gfLoadAuthority(Me, Command1)
    Call gfLoadAuthority(Me, Command2)
    Call gfLoadAuthority(Me, Command3)
    Call gfLoadAuthority(Me, TreeView1)
    
End Sub

Private Sub Form_Resize()
    
    Const conHeight As Long = 6500
    Const conEdge As Long = 120
    Const conTB As Long = 400
    
    If Me.WindowState <> vbMinimized Then
        If Me.Height > conHeight Then
            If Me.ScaleHeight > conEdge * 2 Then
                Frame1.Item(0).Height = Me.ScaleHeight - conEdge * 2
                Frame1.Item(1).Height = Frame1.Item(0).Height
                TreeView1.Height = Frame1.Item(0).Height - conTB
                TreeView2.Height = TreeView1.Height
                ctlMove.Height = Frame1.Item(0).Height
            End If
        End If
    End If
    
    Call gsFormScrollBar(Me, Me.ctlMove, Me.Hsb, Me.Vsb, 16000, 9000)  '注意长、宽的修改
    
End Sub

Private Sub Hsb_Change()
    ctlMove.Left = -Hsb.Value
End Sub

Private Sub Hsb_Scroll()
    Call Hsb_Change    '当滑动滚动条中的滑块时会同时更新对应内容，以下同。
End Sub

Private Sub Vsb_Change()
    ctlMove.Top = -Vsb.Value
End Sub

Private Sub Vsb_Scroll()
    Call Vsb_Change
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim lngLen As Long, I As Long
    Dim strKey As String, strUID As String, strSQL As String, strMsg As String
    Dim rsUser As ADODB.Recordset
    
    strKey = Node.Key
    lngLen = Len(strKey)
    If lngLen < Len(mKeyUser) Then Exit Sub
    If Left(strKey, Len(mKeyDept)) = mKeyDept Then
        For mlngID = Text1.LBound To Text1.UBound
            Text1.Item(mlngID).Text = ""
        Next
        Option1.Item(1).Value = True
        Combo1.Item(0).ListIndex = -1
        Exit Sub
    End If
    If Left(strKey, Len(mKeyUser)) <> mKeyUser Then Exit Sub
    
    strUID = Right(Node.Key, lngLen - Len(mKeyUser))
    strSQL = "EXEC sp_FT_Sys_UserInfo '" & strUID & "'"
    Set rsUser = gfBackRecordset(strSQL)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount = 0 Then
        strMsg = "用户信息丢失了，请联系管理员！"
        GoTo LineBreak
    ElseIf rsUser.RecordCount > 1 Then
        strMsg = "用户信息异常，请联系管理员！"
        GoTo LineBreak
    Else
        Text1.Item(0).Text = strUID
        Text1.Item(1).Text = rsUser.Fields("UserLoginName").Value & ""
        Text1.Item(2).Text = DecryptString(rsUser.Fields("UserPassword").Value & "", gVar.EncryptKey)
        Text1.Item(3).Text = rsUser.Fields("UserFullName").Value & ""
        Text1.Item(4).Text = rsUser.Fields("UserMemo").Value & ""
        Text1.Item(5).Text = strUID & mTwoBar & rsUser.Fields("UserFullName")
        
        Option1.Item(0).Value = IIf(rsUser.Fields("UserSex") = Option1.Item(0).Caption, True, False)
        Option1.Item(1).Value = IIf(rsUser.Fields("UserSex") = Option1.Item(1).Caption, True, False)
        Option1.Item(2).Value = IIf(rsUser.Fields("UserState") = Option1.Item(2).Caption, True, False)
        Option1.Item(3).Value = IIf(rsUser.Fields("UserState") = Option1.Item(3).Caption, True, False)
        
        If IsNull(rsUser.Fields("DeptID").Value) Then
            Combo1.Item(0).ListIndex = -1
        Else
            For I = 0 To Combo1.Item(1).ListCount - 1
                If rsUser.Fields("DeptID").Value = Combo1.Item(1).List(I) Then
                    Combo1.Item(0).ListIndex = I
                    Exit For
                End If
            Next
            If I = Combo1.Item(1).ListCount Then Combo1.Item(0).ListIndex = -1
        End If

        Node.SelectedImage = "SelectedMen"
        
        rsUser.Close
        
        Call msLoadUserRole(strUID)
        
    End If
    
    GoTo LineEnd
    
LineBreak:
    rsUser.Close
    MsgBox strMsg, vbExclamation
LineEnd:
    If rsUser.State = adStateOpen Then rsUser.Close
    Set rsUser = Nothing
    
End Sub

Private Sub TreeView2_NodeCheck(ByVal Node As MSComctlLib.Node)
    '
    Call gsNodeCheckCascade(Node, Node.Checked)
    
End Sub
