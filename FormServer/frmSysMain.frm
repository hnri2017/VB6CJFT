VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.CommandBars.v15.3.1.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form frmSysMain 
   Caption         =   "Main�����"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   12675
   StartUpPosition =   2  '��Ļ����
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   720
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   3840
   End
   Begin FlexCell.Grid Grid1 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5106
      Cols            =   5
      GridColor       =   12632256
      Rows            =   30
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   2040
      Top             =   3840
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   2640
      Top             =   3840
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


Private Sub msGridSet()
    With Grid1
        .AutoRedraw = False
        .Appearance = Flat
        .BackColorBkg = Me.BackColor
        .DisplayRowIndex = True
        .ExtendLastCol = True
        
        .Cols = 7
        .Cell(0, 0).Text = "���"
        .Cell(0, 1).Text = "�û�IP��ַ"
        .Cell(0, 2).Text = "���ӱ�ʶ"
        .Cell(0, 3).Text = "���Ӻ���"
        .Cell(0, 4).Text = "��½�˺�"
        .Cell(0, 5).Text = "�û�����"
        .Cell(0, 6).Text = "����ʱ��"
        .Column(1).Width = 120
        .RowHeight(0) = 40
        .Range(0, 0, 0, .Cols - 1).WrapText = True
        .ReadOnly = True
        
        .AutoRedraw = True
        .Refresh
    End With
End Sub

Private Sub CommandBars1_Resize()
    
    Dim L As Long, T As Long, R As Long, B As Long
    
    On Error Resume Next
    
    CommandBars1.GetClientRect L, T, R, B
    Grid1.Move L, T, R - L, B - T
    
End Sub

Private Sub Form_Load()
    '�������
    
    
    Call Main   '��ʼ������
    Call gfLoadSkin(Me, SkinFramework1, sMSO7)  '���ش�������
    
    Call CommandBars1.LoadCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegSectionSettings)
    Call gsFormSizeLoad(Me)
    
    Call msGridSet  '�������
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '����ע�����Ϣ-CommandBars����
    Call CommandBars1.SaveCommandBars(gVar.RegKeyCommandBars, gVar.RegAppName, gVar.RegSectionSettings)
    
    '����ע�����Ϣ-���ڴ�С
    Call gsFormSizeSave(Me)
    
End Sub
