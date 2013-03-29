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
            Text            =   "状态"
            TextSave        =   "状态"
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
      Caption         =   "文件(&F)"
      Begin VB.Menu mnu_File_Close 
         Caption         =   "关闭(&C)"
      End
      Begin VB.Menu mnu_File_CloseAll 
         Caption         =   "关闭所有"
      End
      Begin VB.Menu mnu_File_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnu_Oper 
      Caption         =   "操作(&O)"
      Begin VB.Menu mnu_Oper_Emp 
         Caption         =   "员工信息(&E)"
      End
      Begin VB.Menu mnu_Oper_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Oper_Service 
         Caption         =   "服务配置(&C)"
         Index           =   0
      End
      Begin VB.Menu mnu_Oper_Service 
         Caption         =   "服务配置向导(&W)"
         Index           =   1
      End
      Begin VB.Menu mnu_Oper_Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Oper_Salary 
         Caption         =   "数据导入(&L)"
         Index           =   0
      End
      Begin VB.Menu mnu_Oper_Salary 
         Caption         =   "邮件发送(&S)"
         Index           =   1
      End
   End
   Begin VB.Menu mnu_Report 
      Caption         =   "报表(&R)"
      Begin VB.Menu mnu_Report_Query 
         Caption         =   "综合查询(&Q)"
      End
      Begin VB.Menu mnu_Report_Bar 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Report_Oper 
         Caption         =   "报表配置(&C)"
         Index           =   0
      End
      Begin VB.Menu mnu_Report_Oper 
         Caption         =   "报表生成(&G)"
         Index           =   1
      End
   End
   Begin VB.Menu mnu_Tool 
      Caption         =   "工具(&T)"
      Begin VB.Menu mnu_Tool_Options 
         Caption         =   "选项(&O)"
      End
      Begin VB.Menu mnu_Tool_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Tool_Sys 
         Caption         =   "系统工具(&S)"
         Begin VB.Menu mnu_Tool_Sys_Maint 
            Caption         =   "自动修复"
            Index           =   0
         End
         Begin VB.Menu mnu_Tool_Sys_Maint 
            Caption         =   "备份与恢复"
            Index           =   1
         End
         Begin VB.Menu mnu_Tool_Sys_Maint 
            Caption         =   "系统初始化"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnu_Win 
      Caption         =   "窗口(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnu_Win_Cascade 
         Caption         =   "重叠(&C)"
      End
      Begin VB.Menu mnu_Win_Horizontal 
         Caption         =   "水平排列(&H)"
      End
      Begin VB.Menu mnu_Win_Vertical 
         Caption         =   "垂直排列(&V)"
      End
      Begin VB.Menu mnu_Win_ArrangeIcons 
         Caption         =   "排列图标(&A)"
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnu_Help_Info 
         Caption         =   "使用手册(&M)"
         Index           =   0
      End
      Begin VB.Menu mnu_Help_Info 
         Caption         =   "常见问题(&F)"
         Index           =   1
      End
      Begin VB.Menu mnu_Help_Info 
         Caption         =   "更新说明(&U)"
         Index           =   2
      End
      Begin VB.Menu mnu_Help_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Help_About 
         Caption         =   "关于(&A)"
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
    If DisplayMsg("真的要退出" & App.Title & "吗？", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
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
    Case 0  '使用手册Manual
        strCaption = "使用手册"
    Case 1  '常见问题FAQ
        strCaption = "常见问题"
    Case 2  '更新说明Change Log
        strCaption = "更新说明"
    End Select

    Load frmMsgLog
    frmMsgLog.Caption = strCaption
    If Dir(App.Path & "\" & strCaption & ".txt") = "" Then
        frmMsgLog.RTDesc.Text = strCaption & "文件被破坏"
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
    Case 0  '数据导入
        strCaption = "数据导入"
    Case 1  '邮件发送
        strCaption = "邮件发送"
    End Select

    Dim iPos As Integer
    iPos = FindForm("frmSalary", strCaption)
    If iPos > -1 Then
        Forms(iPos).SetFocus
        Exit Sub
    End If

    Dim oForm As New frmSalary
    Load oForm
    If Index = 1 Then   '邮件发送
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
    Case 0  '服务配置
        frmService.SetFocus
    Case 1  '服务配置向导
        frmServWizard.Show vbModal
    End Select
End Sub

Private Sub mnu_Report_Oper_Click(Index As Integer)
    Select Case Index
    Case 0  '报表配置
        frmRptMain.SetFocus
    Case 1  '报表生成
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
        If DisplayMsg("运行系统工具需关闭其他窗口。" & vbCrLf & vbCrLf & "现在关闭其他的窗口吗？", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
        Call CloseAllSubForm
    End If

    Select Case Index
    Case 0  '自动修复
        frmCheck.Show vbModal
    Case 1  '备份与恢复
        frmBackup.Show vbModal
    Case 2  '系统初始化
        frmInit.Show vbModal
    Case 3  '检查更新
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
