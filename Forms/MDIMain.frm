VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "salary"
   ClientHeight    =   4605
   ClientLeft      =   2715
   ClientTop       =   2640
   ClientWidth     =   8205
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4335
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6192
            MinWidth        =   5292
            Text            =   "״̬"
            TextSave        =   "״̬"
            Key             =   "display"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "user"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2010-11-8"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "15:28"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_File 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnu_File_Close 
         Caption         =   "�ر�(&C)"
      End
      Begin VB.Menu mnu_File_CloseAll 
         Caption         =   "�ر�����"
      End
      Begin VB.Menu mnu_File_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnu_Oper 
      Caption         =   "����(&O)"
      Begin VB.Menu mnu_Oper_Emp 
         Caption         =   "Ա����Ϣ(&E)"
      End
      Begin VB.Menu mnu_Oper_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Oper_Service 
         Caption         =   "��������(&C)"
         Index           =   0
      End
      Begin VB.Menu mnu_Oper_Service 
         Caption         =   "����������(&W)"
         Index           =   1
      End
      Begin VB.Menu mnu_Oper_Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Oper_Salary 
         Caption         =   "���ݵ���(&L)"
         Index           =   0
      End
      Begin VB.Menu mnu_Oper_Salary 
         Caption         =   "�ʼ�����(&S)"
         Index           =   1
      End
   End
   Begin VB.Menu mnu_Report 
      Caption         =   "����(&R)"
      Begin VB.Menu mnu_Report_Query 
         Caption         =   "�ۺϲ�ѯ(&Q)"
      End
      Begin VB.Menu mnu_Report_Bar 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Report_Oper 
         Caption         =   "��������(&C)"
         Index           =   0
      End
      Begin VB.Menu mnu_Report_Oper 
         Caption         =   "��������(&G)"
         Index           =   1
      End
   End
   Begin VB.Menu mnu_Tool 
      Caption         =   "����(&T)"
      Begin VB.Menu mnu_Tool_Options 
         Caption         =   "ѡ��(&O)"
      End
      Begin VB.Menu mnu_Tool_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Tool_Sys 
         Caption         =   "ϵͳ����(&S)"
         Begin VB.Menu mnu_Tool_Sys_Maint 
            Caption         =   "�Զ��޸�"
            Index           =   0
         End
         Begin VB.Menu mnu_Tool_Sys_Maint 
            Caption         =   "������ָ�"
            Index           =   1
         End
         Begin VB.Menu mnu_Tool_Sys_Maint 
            Caption         =   "ϵͳ��ʼ��"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnu_Win 
      Caption         =   "����(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnu_Win_Cascade 
         Caption         =   "�ص�(&C)"
      End
      Begin VB.Menu mnu_Win_Horizontal 
         Caption         =   "ˮƽ����(&H)"
      End
      Begin VB.Menu mnu_Win_Vertical 
         Caption         =   "��ֱ����(&V)"
      End
      Begin VB.Menu mnu_Win_ArrangeIcons 
         Caption         =   "����ͼ��(&A)"
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "����(&H)"
      Begin VB.Menu mnu_Help_Info 
         Caption         =   "ʹ���ֲ�(&M)"
         Index           =   0
      End
      Begin VB.Menu mnu_Help_Info 
         Caption         =   "��������(&F)"
         Index           =   1
      End
      Begin VB.Menu mnu_Help_Info 
         Caption         =   "����˵��(&U)"
         Index           =   2
      End
      Begin VB.Menu mnu_Help_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Help_About 
         Caption         =   "����(&A)"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    Me.Caption = App.Title & " -- Ver " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If DisplayMsg("���Ҫ�˳�" & App.Title & "��", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    CloseAllSubForm
    CloseDB
End Sub

Private Sub mnu_File_Close_Click()
    If Not Me.ActiveForm Is Nothing Then
        Unload Me.ActiveForm
    End If
End Sub

Private Sub mnu_File_CloseAll_Click()
    CloseAllSubForm
End Sub

Private Sub mnu_File_Exit_Click()
    Unload Me
End Sub

Private Sub mnu_Help_About_Click()
    frmAbout.Show vbModal
    'OpenForm "frmAbout", , vbModal
End Sub

Private Sub mnu_Help_Info_Click(Index As Integer)
    Dim strCaption As String
    Select Case Index
    Case 0  'ʹ���ֲ�Manual
        strCaption = "ʹ���ֲ�"
    Case 1  '��������FAQ
        strCaption = "��������"
    Case 2  '����˵��Change Log
        strCaption = "����˵��"
    End Select

    Load frmMsgLog
    frmMsgLog.Caption = strCaption
    If Dir(App.Path & "\" & strCaption & ".txt") = "" Then
        frmMsgLog.RTDesc.Text = strCaption & "�ļ����ƻ�"
    Else
        frmMsgLog.RTDesc.LoadFile App.Path & "\" & strCaption & ".txt"
    End If
    frmMsgLog.Show vbModal
End Sub

Private Sub mnu_Oper_Emp_Click()
    frmMaint.SetFocus
    'OpenForm "frmMaint"
End Sub

Private Sub mnu_Oper_Salary_Click(Index As Integer)
    Dim strCaption As String

    Select Case Index
    Case 0  '���ݵ���
        strCaption = "���ݵ���"
    Case 1  '�ʼ�����
        strCaption = "�ʼ�����"
    End Select

    Dim iPos As Integer
    iPos = FindForm("frmSalary", strCaption)
    If iPos > -1 Then
        Forms(iPos).SetFocus
        Exit Sub
    End If

    Dim oForm As New frmSalary
    Load oForm
    If Index = 1 Then   '�ʼ�����
        With oForm
            .Caption = strCaption
            .dtServPeriod.Visible = False
            .cbServPeriod.Visible = True
            .cmdLoad.Visible = False
            .cmdDel.Visible = True
        End With
    End If
    oForm.Show
End Sub

Private Sub mnu_Oper_Service_Click(Index As Integer)
    Select Case Index
    Case 0  '��������
        frmService.SetFocus
    Case 1  '����������
        frmServWizard.Show vbModal
    End Select
End Sub

Private Sub mnu_Report_Oper_Click(Index As Integer)
    Select Case Index
    Case 0  '��������
        frmRptMain.SetFocus
    Case 1  '��������
        frmReport.SetFocus
    End Select
End Sub

Private Sub mnu_Report_Query_Click()
    frmQuery.SetFocus
End Sub

Private Sub mnu_Tool_Options_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnu_Tool_Sys_Maint_Click(Index As Integer)
    If Forms.Count > 1 Then
        If DisplayMsg("����ϵͳ������ر��������ڡ�" & vbCrLf & vbCrLf & "���ڹر������Ĵ�����", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
        Call CloseAllSubForm
    End If

    Select Case Index
    Case 0  '�Զ��޸�
        frmCheck.Show vbModal
    Case 1  '������ָ�
        frmBackup.Show vbModal
    Case 2  'ϵͳ��ʼ��
        frmInit.Show vbModal
    Case 3  '������
    End Select
End Sub

Private Sub mnu_Win_ArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnu_Win_Cascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnu_Win_Horizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnu_Win_Vertical_Click()
    Me.Arrange vbTileVertical
End Sub
