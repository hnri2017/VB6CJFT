VERSION 5.00
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ѡ��"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin FlexCell.Grid Grid1 
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3201
      Cols            =   5
      GridColor       =   12632256
      Rows            =   30
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub msLoadParameter(Optional ByVal blnLoad As Boolean = True)
    
    If Not blnLoad Then Exit Sub
    
    '�ӹ���������ע����м���������Ϣ
    With Me.Grid1
        .Cell(2, 1).Text = gVar.ParaBlnWindowCloseMin   '�ر�ʱ��С��
        .Cell(2, 5).Text = gVar.ParaBlnWindowMinHide    '��С��ʱ����
        
    
    End With
    
End Sub

Private Sub msSaveParameter(Optional ByVal blnSave As Boolean = True)
    
    If Not blnSave Then Exit Sub
    
    '����ֵ��������������
    With Grid1
        gVar.ParaBlnWindowCloseMin = .Cell(2, 1).Text
        gVar.ParaBlnWindowMinHide = .Cell(2, 5).Text

    End With
    
    '����ֵͨ�����ñ��������ע�����
    With gVar
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowMinHide, IIf(.ParaBlnWindowMinHide, 1, 0))
        Call SaveSetting(.RegAppName, .RegSectionSettings, .RegKeyParaWindowCloseMin, IIf(.ParaBlnWindowCloseMin, 1, 0))
        
    End With
    
    If MsgBox("����������ɣ��Ƿ������˳����ڣ�", vbInformation + vbYesNo, "��ʾ") = vbYes Then Unload Me
    
End Sub


Private Sub Form_Load()
    Dim strFile As String
    
    strFile = gVar.FolderNameBin & "OptionWindow.cel"
    If Not gfFileExist(strFile) Then
        MsgBox "���������ļ�����ʧ�ܣ������������´򿪴��ڡ�" & vbCrLf & strFile, vbCritical, "�쳣��ʾ"
        Exit Sub
    End If
    With Grid1
        .AutoRedraw = False
        .OpenFile (strFile)
        .Appearance = Flat
        .Column(0).Width = 0
        .RowHeight(0) = 0
        .ExtendLastCol = True
        .GridColor = vbWhite
        .BorderColor = Me.BackColor
        .BackColorBkg = Me.BackColor
        
        Call msLoadParameter(True)
        
        .AutoRedraw = True
        .Refresh
    End With
End Sub

Private Sub Form_Resize()
    Grid1.Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - 240
End Sub

Private Sub Grid1_HyperLinkClick(ByVal Row As Long, ByVal Col As Long, URL As String, Changed As Boolean)
    '��������ֵ
    
    URL = ""
    Changed = True
    If Row <> (Grid1.Rows - 1) Then Exit Sub
    
    If Col = 1 Then '����
        If MsgBox("ȷ���������в���ֵ��", vbQuestion + vbOKCancel, "����ѯ��") = vbOK Then Call msSaveParameter(True)
    ElseIf Col = 5 Then '�˳�
        Unload Me
    End If
End Sub
