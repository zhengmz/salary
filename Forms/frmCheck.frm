VERSION 5.00
Begin VB.Form frmCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "自动修复"
   ClientHeight    =   6015
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   8070
   Icon            =   "frmCheck.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4980
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   4980
      ScaleWidth      =   7605
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Width           =   7605
      Begin VB.Frame fraFixLog 
         Caption         =   "自动修复日志"
         Height          =   4785
         Left            =   240
         TabIndex        =   8
         Top             =   0
         Width           =   7095
         Begin VB.TextBox txtFixLog 
            Height          =   4215
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   360
            Width           =   6615
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4980
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   4980
      ScaleWidth      =   7605
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   240
      Width           =   7605
      Begin VB.Frame fraStep3 
         Caption         =   "自动修复"
         Height          =   4785
         Left            =   240
         TabIndex        =   7
         Top             =   0
         Width           =   7095
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "按 ""下一步"" 进行自动修复。"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   29
            Top             =   4200
            Width           =   2445
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "3、主要是对服务配置与各数据之间的协同进行理顺。"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   840
            TabIndex        =   27
            Top             =   1920
            Width           =   4590
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "2、只会对影响系统运行的内部关联进行调整；"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   840
            TabIndex        =   26
            Top             =   1440
            Width           =   4005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "1、不会对业务数据进行修改；"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   840
            TabIndex        =   25
            Top             =   960
            Width           =   2640
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "修复功能介绍："
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   20
            Top             =   480
            Width           =   1365
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4980
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   4980
      ScaleWidth      =   7605
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   7605
      Begin VB.Frame fraStep2 
         Caption         =   "修复建议"
         Height          =   1215
         Left            =   240
         TabIndex        =   21
         Top             =   3600
         Width           =   7095
         Begin VB.CheckBox chkBackup 
            Caption         =   "建议备份"
            Height          =   195
            Left            =   4080
            TabIndex        =   24
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lbResult 
            AutoSize        =   -1  'True
            Caption         =   "系统完好，无需修复。"
            Height          =   195
            Left            =   360
            TabIndex        =   23
            Top             =   480
            Width           =   1800
         End
         Begin VB.Label lbBackup 
            AutoSize        =   -1  'True
            Caption         =   "建议在修改系统之前，对数据进行备份，而且是完整的备份。"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   22
            Top             =   840
            Visible         =   0   'False
            Width           =   5265
         End
      End
      Begin VB.Frame fraCheckLog 
         Caption         =   "系统自查报告"
         Height          =   3585
         Left            =   240
         TabIndex        =   6
         Top             =   0
         Width           =   7095
         Begin VB.TextBox txtCheckLog 
            Height          =   3015
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   360
            Width           =   6615
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4980
      Index           =   0
      Left            =   210
      ScaleHeight     =   4980
      ScaleWidth      =   7605
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   7605
      Begin VB.Frame fraDesc 
         Caption         =   "向导说明"
         Height          =   4815
         Left            =   240
         TabIndex        =   9
         Top             =   0
         Width           =   7215
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "第一步"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   18
            Top             =   480
            Width           =   585
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "先对系统进行自查，并产生检查报告；"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   17
            Top             =   480
            Width           =   3060
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "第二步"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   16
            Top             =   990
            Width           =   585
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "可根据第一步产生的检查报告，决定是否进行修复；"
            Height          =   195
            Index           =   3
            Left            =   1440
            TabIndex        =   15
            Top             =   990
            Width           =   4140
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "如果需要自动修复，建议进行备份工作；"
            Height          =   195
            Index           =   4
            Left            =   1440
            TabIndex        =   14
            Top             =   1500
            Width           =   3240
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "第三步"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   480
            TabIndex        =   13
            Top             =   2010
            Width           =   585
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "调用备份与恢复功能进行备份工作（可选）；"
            Height          =   195
            Index           =   6
            Left            =   1440
            TabIndex        =   12
            Top             =   2010
            Width           =   3600
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "第四步"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   480
            TabIndex        =   11
            Top             =   2520
            Width           =   585
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "对系统进行自动化的修复，并产生相应的结果报告。"
            Height          =   195
            Index           =   8
            Left            =   1440
            TabIndex        =   10
            Top             =   2520
            Width           =   4140
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   5415
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一步"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   5415
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   5175
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iStep As Integer
Private m_blNeedFix As Boolean
Private Const m_iMaxStep As Integer = 4
Private m_strDelService As String
Private m_strModService As String
Private m_blNeedCompareDB As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    m_iStep = m_iStep + 1
    Call StepProc
End Sub

Private Sub Form_Load()
    '初始化变量
    m_iStep = 1
    m_blNeedFix = False     '需要修复标识
    m_strDelService = ""    '无效或无用的服务配置列表，可删除
    m_strModService = ""    '需要修改字段映射关系的服务配置列表
    m_blNeedCompareDB = False '是否需要压缩数据库

    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub StepProc()
    Dim i As Integer
    For i = 0 To m_iMaxStep - 1
        If i = m_iStep - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next

    Select Case m_iStep
    Case 1  '向导首页
    Case 2  '自查，并产生报告
        Call AutoCheck
        If m_blNeedFix = False Then
            cmdNext.Enabled = False
        Else
            lbBackup.Visible = True
            chkBackup.Visible = True
            chkBackup.value = 1
            lbResult.Caption = "系统存在问题，建议修复。"
            cmdNext.Enabled = True
        End If
    Case 3  '备份
        If chkBackup.value = 1 Then
            Me.Hide
            Me.Show vbModeless
            frmBackup.Show vbModal
            Me.Hide
            Me.Show vbModal
        End If
    Case 4  '自动修复，并产生日志
        Call AutoFix
        cmdNext.Enabled = False
    End Select
End Sub

Private Sub AutoFix()
    Dim strMsgLog As String
    
    txtFixLog.Text = "修复日志如下：" & vbCrLf & vbCrLf
    
    '先修复服务配置表
    strMsgLog = FixServices()
    If strMsgLog <> "" Then
        txtFixLog.Text = txtFixLog.Text & strMsgLog & vbCrLf & vbCrLf
    End If
    
    '修复服务配置序号
    strMsgLog = FixServiceID()
    If strMsgLog <> "" Then
        txtFixLog.Text = txtFixLog.Text & strMsgLog & vbCrLf & vbCrLf
    End If
    
    '修复报表ID序号
    strMsgLog = FixReportID()
    If strMsgLog <> "" Then
        txtFixLog.Text = txtFixLog.Text & strMsgLog & vbCrLf & vbCrLf
    End If
    
    If m_strModService <> "" Then
        txtFixLog.Text = txtFixLog.Text & "无法修复，需人工增加字段映射的服务配置列表如下：" & vbCrLf & m_strModService & vbCrLf & vbCrLf
    End If
    
    '压缩数据文件
    strMsgLog = CompareDB()
    If strMsgLog <> "" Then
        txtFixLog.Text = txtFixLog.Text & strMsgLog & vbCrLf & vbCrLf
    End If
    
    '最后显示基本表数据
    strMsgLog = CheckTabs()
    If strMsgLog <> "" Then
        txtFixLog.Text = txtFixLog.Text & "修复调整后的" & strMsgLog
    End If
End Sub

Private Function FixServices() As String
    Dim strSQL As String
    Dim iLastPos As Integer
    Dim iCurrPos As Integer
    Dim strServId As String
    Dim strSuccMsgLog As String
    Dim strErrMsgLog As String
    
    strSuccMsgLog = ""
    strErrMsgLog = ""
    If m_strDelService <> "" Then
        iLastPos = 1
        iCurrPos = InStr(iLastPos, m_strDelService, ",")
        Do While iCurrPos > 0
            If iCurrPos = 0 Then
                strServId = m_strDelService
            Else
                strServId = Mid(m_strDelService, iLastPos, iCurrPos - iLastPos)
            End If
            strSQL = "delete from services where serv_id='" & strServId & "'"
            If gExecSql(strSQL) = True Then
                strSuccMsgLog = strSuccMsgLog & vbCrLf & strServId
            Else
                strErrMsgLog = strErrMsgLog & vbCrLf & strServId
            End If
            iLastPos = iCurrPos + 1
            iCurrPos = InStr(iLastPos, m_strDelService, ",")
        Loop
        If iCurrPos = 0 And iLastPos < Len(m_strDelService) Then
            strServId = Mid(m_strDelService, iLastPos, Len(m_strDelService) - iLastPos + 1)
            strSQL = "delete from services where serv_id='" & strServId & "'"
            If gExecSql(strSQL) = True Then
                strSuccMsgLog = strSuccMsgLog & vbCrLf & strServId
            Else
                strErrMsgLog = strErrMsgLog & vbCrLf & strServId
            End If
        End If
    End If
    FixServices = ""
    If strSuccMsgLog <> "" Then
        FixServices = FixServices & "成功删除的无用服务配置列表：" & strSuccMsgLog & vbCrLf
    End If
    If strErrMsgLog <> "" Then
        FixServices = FixServices & "无法删除的无用服务配置列表：" & strErrMsgLog & vbCrLf
    End If
End Function

Private Function FixServiceID() As String
    Dim rsDicts As New ADODB.Recordset
    Dim rsService As New ADODB.Recordset
    Dim iOldMaxSeq As Integer
    Dim iSeq As Integer
    Dim iNewMaxSeq As Integer
    Dim strSQL As String
    Dim strSuccMsgLog As String
    Dim strErrMsgLog As String
    
    strSuccMsgLog = ""
    strErrMsgLog = ""
    strSQL = "select dict_key, dict_type, dict_value from dicts where dict_sect='OPT_SERV_TYPE'"
    rsDicts.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    Do While rsDicts.EOF = False
        iOldMaxSeq = Val(rsDicts("dict_value"))
        strSQL = "select serv_id from services where serv_id like '" & rsDicts("dict_type") & "%' order by serv_id desc"
        rsService.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
        iNewMaxSeq = 0
        Do While rsService.EOF = False
            iSeq = Val(Mid(rsService(0), Len(rsDicts("dict_type")) + 1, Len(rsService(0)) - Len(rsDicts("dict_type"))))
            If iSeq > iNewMaxSeq Then
                iNewMaxSeq = iSeq
            End If
            rsService.MoveNext
        Loop
        rsService.Close
        If iOldMaxSeq <> iNewMaxSeq Then
            strSQL = "update dicts set dict_value='" & iNewMaxSeq & "' where dict_sect='OPT_SERV_TYPE' and dict_key='" & rsDicts("dict_key") & "'"
            If gExecSql(strSQL) = True Then
                strSuccMsgLog = strSuccMsgLog & vbCrLf & "将类型 '" & rsDicts("dict_key") & "' 的当前序号从 " & iOldMaxSeq & " 更改为 " & iNewMaxSeq
            Else
                strErrMsgLog = strErrMsgLog & vbCrLf & "类型 '" & rsDicts("dict_key") & "' 的当前序号是： " & iOldMaxSeq & " ，而数据字典中的当前序号为： " & iNewMaxSeq
            End If
        End If
        rsDicts.MoveNext
    Loop
    rsDicts.Close
    Set rsDicts = Nothing
    Set rsService = Nothing
    FixServiceID = ""
    If strSuccMsgLog <> "" Then
        FixServiceID = FixServiceID & "成功修复服务配置序号，具体如下：" & strSuccMsgLog & vbCrLf
    End If
    If strErrMsgLog <> "" Then
        FixServiceID = FixServiceID & "无法修复服务配置序号，具体如下：" & strErrMsgLog & vbCrLf
    End If
End Function

Private Function FixReportID() As String
    Dim rsDicts As New ADODB.Recordset
    Dim rsReport As New ADODB.Recordset
    Dim iOldMaxSeq As Integer
    Dim iSeq As Integer
    Dim iNewMaxSeq As Integer
    Dim strSQL As String
    Dim strSuccMsgLog As String
    Dim strErrMsgLog As String
    
    strSuccMsgLog = ""
    strErrMsgLog = ""
    strSQL = "select dict_key, dict_type,dict_value from dicts where dict_sect='OPT_REPORT' and dict_key='REPORT_ID'"
    rsDicts.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    Do While rsDicts.EOF = False
        iOldMaxSeq = Val(rsDicts("dict_value"))
        strSQL = "select rpt_id from reports where rpt_id like '" & rsDicts("dict_type") & "%' order by rpt_id desc"
        rsReport.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
        iNewMaxSeq = 0
        Do While rsReport.EOF = False
            iSeq = Val(Mid(rsReport(0), Len(rsDicts("dict_type")) + 1, Len(rsReport(0)) - Len(rsDicts("dict_type"))))
            If iSeq > iNewMaxSeq Then
                iNewMaxSeq = iSeq
            End If
            rsReport.MoveNext
        Loop
        rsReport.Close
        If iOldMaxSeq <> iNewMaxSeq Then
            strSQL = "update dicts set dict_value='" & iNewMaxSeq & "' where dict_sect='OPT_REPORT' and dict_key='" & rsDicts("dict_key") & "'"
            If gExecSql(strSQL) = True Then
                strSuccMsgLog = strSuccMsgLog & vbCrLf & "将报表序号的当前序号从 " & iOldMaxSeq & " 更改为 " & iNewMaxSeq
            Else
                strErrMsgLog = strErrMsgLog & vbCrLf & "报表使用的当前序号是： " & iOldMaxSeq & " ，而数据字典中的当前序号为： " & iNewMaxSeq
            End If
        End If
        rsDicts.MoveNext
    Loop
    rsDicts.Close
    Set rsDicts = Nothing
    Set rsReport = Nothing

    FixReportID = ""
    If strSuccMsgLog <> "" Then
        FixReportID = FixReportID & "成功修复报表序号，具体如下：" & strSuccMsgLog & vbCrLf
    End If
    If strErrMsgLog <> "" Then
        FixReportID = FixReportID & "无法修复报表序号，具体如下：" & strErrMsgLog & vbCrLf
    End If
End Function

'与CheckUsedCount对应
Private Function CompareDB() As String
    If m_blNeedCompareDB = False Then
        CompareDB = ""
        Exit Function
    End If

    Dim strRetMsg As String

    If gCompareDB(strRetMsg) = False Then
        DisplayMsg "压缩数据库时出错", vbCritical
        CompareDB = "在压缩数据库时出错，具体错误信息如下：" & vbCrLf & strRetMsg
        Exit Function
    End If

    '修改数据字典
    Dim rsDicts As New ADODB.Recordset
    Dim strSQL As String
    Dim strArrCnt() As String

    strSQL = "select dict_type, dict_value from dicts where dict_sect='SYS' and dict_key='USED_COUNT'"
    rsDicts.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText

    If rsDicts.RecordCount > 0 Then
        strArrCnt = Split(rsDicts("dict_value"), "-", 2, vbTextCompare)
        If UBound(strArrCnt) = 0 Then
            strSQL = "update dicts set dict_value='" & strArrCnt(0) & "-" & strArrCnt(0) & "' where dict_sect='SYS' and dict_key='USED_COUNT'"
        Else
            strSQL = "update dicts set dict_value='" & strArrCnt(1) & "-" & strArrCnt(1) & "' where dict_sect='SYS' and dict_key='USED_COUNT'"
        End If
        gExecSql (strSQL)
    End If
    rsDicts.Close
    Set rsDicts = Nothing
    CompareDB = strRetMsg
End Function

Private Sub AutoCheck()
    Dim strMsgLog As String
    
    txtCheckLog.Text = "检查报告如下：" & vbCrLf & vbCrLf
    
    '检查基本表数据
    strMsgLog = CheckTabs()
    If strMsgLog <> "" Then
        txtCheckLog.Text = txtCheckLog.Text & strMsgLog
    End If
    '检查服务配置
    strMsgLog = CheckServices()
    If strMsgLog <> "" Then
        txtCheckLog.Text = txtCheckLog.Text & vbCrLf & vbCrLf & strMsgLog
    End If
    
    '检查ServiceID序号的唯一性
    strMsgLog = CheckServiceID()
    If strMsgLog <> "" Then
        txtCheckLog.Text = txtCheckLog.Text & vbCrLf & vbCrLf & strMsgLog
    End If
    
    '检查ReportID序号的唯一性
    strMsgLog = CheckReportID()
    If strMsgLog <> "" Then
        txtCheckLog.Text = txtCheckLog.Text & vbCrLf & vbCrLf & strMsgLog
    End If
    
    '检查是否需要压缩
    strMsgLog = CheckUsedCount()
    If strMsgLog <> "" Then
        txtCheckLog.Text = txtCheckLog.Text & vbCrLf & vbCrLf & strMsgLog
    End If
    
    If m_blNeedFix = True Then
        txtCheckLog.Text = txtCheckLog.Text & vbCrLf & vbCrLf & "建议对系统进行自动修复。"
    Else
        txtCheckLog.Text = txtCheckLog.Text & vbCrLf & vbCrLf & "系统完好，无需修复。"
    End If
    
End Sub

Private Function CheckTabs() As String
    On Error GoTo ErrHandle
    Dim rsTabs As New ADODB.Recordset
    Dim rsTabCount As New ADODB.Recordset
    Dim strSQL As String
    Dim iCount As Integer
    Dim strMsgLog As String
    
    strSQL = "select dict_key as table_name, dict_value as table_desc from dicts where dict_sect='TABLE_NAME'"
    rsTabs.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    strMsgLog = ""
    Do While rsTabs.EOF = False
        strSQL = "select count(*) from " & rsTabs("table_name")
        iCount = -1
        rsTabCount.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
        iCount = rsTabCount(0)
        rsTabCount.Close
        If iCount = -1 Then
            strMsgLog = strMsgLog & "不存在" & rsTabs("table_desc") & UCase(rsTabs("table_name")) & "，请找系统管理员检查。" & vbCrLf
        Else
            strMsgLog = strMsgLog & rsTabs("table_desc") & UCase(rsTabs("table_name")) & "，共有 " & iCount & " 条记录。" & vbCrLf
        End If
        rsTabs.MoveNext
    Loop
    rsTabs.Close
    Set rsTabs = Nothing
    Set rsTabCount = Nothing

    CheckTabs = ""
    If strMsgLog = "" Then
        CheckTabs = "没有基本数据的配置表，请找系统管理员检查数据字典。" & vbCrLf
    Else
        CheckTabs = "基本数据：" & vbCrLf & strMsgLog
    End If
    Exit Function
    
ErrHandle:
    iCount = -1
    Resume Next
End Function

Private Function CheckServices() As String
    Dim rsService As New ADODB.Recordset
    Dim rsServConf As New ADODB.Recordset
    Dim rsSalary As New ADODB.Recordset
    Dim strSQL As String
    Dim strMsgLog As String

    m_strDelService = ""
    m_strModService = ""
    strSQL = "select * from services"
    rsService.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    Do While rsService.EOF = False
        strSQL = "select count(*) from serv_field where serv_id='" & rsService("serv_id") & "'"
        rsServConf.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
        If rsServConf(0) = 0 Then
            strSQL = "select count(*) from salary where serv_id='" & rsService("serv_id") & "'"
            rsSalary.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
            If rsSalary(0) = 0 Then
                m_strDelService = m_strDelService & "," & rsService("serv_id")
            Else
                m_strModService = m_strModService & "," & rsService("serv_id")
            End If
            rsSalary.Close
        End If
        rsServConf.Close
        rsService.MoveNext
    Loop
    rsService.Close
    Set rsService = Nothing
    Set rsServConf = Nothing
    Set rsSalary = Nothing

    strMsgLog = ""
    If m_strDelService <> "" Then
        m_strDelService = Mid(m_strDelService, 2, Len(m_strDelService) - 1)
        strMsgLog = strMsgLog & "无效的服务配置列表如下：" & vbCrLf & m_strDelService
        m_blNeedFix = True
    End If
    If m_strModService <> "" Then
        m_strModService = Mid(m_strModService, 2, Len(m_strModService) - 1)
        strMsgLog = strMsgLog & "需要增加字段映射关系的服务配置列表如下：" & vbCrLf & m_strModService
        m_blNeedFix = True
    End If
    CheckServices = strMsgLog
End Function

Private Function CheckServiceID() As String
    Dim rsDicts As New ADODB.Recordset
    Dim rsService As New ADODB.Recordset
    Dim iMaxSeq As Integer
    Dim iSeq As Integer
    Dim strSQL As String
    Dim strMsgLog As String
    
    strMsgLog = ""
    strSQL = "select dict_key, dict_type, dict_value from dicts where dict_sect='OPT_SERV_TYPE'"
    rsDicts.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    Do While rsDicts.EOF = False
        iMaxSeq = Val(rsDicts("dict_value"))
        strSQL = "select serv_id from services where serv_id like '" & rsDicts("dict_type") & "%' order by serv_id desc"
        rsService.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
        Do While rsService.EOF = False
            iSeq = Val(Mid(rsService(0), Len(rsDicts("dict_type")) + 1, Len(rsService(0)) - Len(rsDicts("dict_type"))))
            If iSeq > iMaxSeq Then
                strMsgLog = "服务配置序号ID存在唯一性重复问题。"
                Exit Do
            End If
            rsService.MoveNext
        Loop
        rsService.Close
        If strMsgLog <> "" Then
            Exit Do
        End If
        rsDicts.MoveNext
    Loop
    rsDicts.Close
    Set rsDicts = Nothing
    Set rsService = Nothing
    If strMsgLog <> "" Then
        strMsgLog = "严重影响系统运行的错误如下：" & vbCrLf & strMsgLog
        m_blNeedFix = True
    End If
    CheckServiceID = strMsgLog
End Function

Private Function CheckReportID() As String
    Dim rsDicts As New ADODB.Recordset
    Dim rsReport As New ADODB.Recordset
    Dim iMaxSeq As Integer
    Dim iSeq As Integer
    Dim strSQL As String
    Dim strMsgLog As String
    
    strMsgLog = ""
    strSQL = "select dict_type, dict_value from dicts where dict_sect='OPT_REPORT' and dict_key='REPORT_ID'"
    rsDicts.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    Do While rsDicts.EOF = False
        iMaxSeq = Val(rsDicts("dict_value"))
        strSQL = "select rpt_id from reports where rpt_id like '" & rsDicts("dict_type") & "%' order by rpt_id desc"
        rsReport.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
        Do While rsReport.EOF = False
            iSeq = Val(Mid(rsReport(0), Len(rsDicts("dict_type")) + 1, Len(rsReport(0)) - Len(rsDicts("dict_type"))))
            If iSeq > iMaxSeq Then
                strMsgLog = "报表序号ID存在唯一性重复问题。"
                Exit Do
            End If
            rsReport.MoveNext
        Loop
        rsReport.Close
        If strMsgLog <> "" Then
            Exit Do
        End If
        rsDicts.MoveNext
    Loop
    rsDicts.Close
    Set rsDicts = Nothing
    Set rsReport = Nothing
    If strMsgLog <> "" Then
        strMsgLog = "严重影响系统运行的错误如下：" & vbCrLf & strMsgLog
        m_blNeedFix = True
    End If
    CheckReportID = strMsgLog
End Function

Private Function CheckUsedCount() As String
    Dim rsDicts As New ADODB.Recordset
    Dim lngGene As Long
    Dim lngUsedCount As Long
    Dim strSQL As String
    Dim strMsgLog As String
    Dim strArrCnt() As String
    
    strMsgLog = ""
    strSQL = "select dict_type as gene, dict_value as used_count from dicts where dict_sect='SYS' and dict_key='USED_COUNT'"
    rsDicts.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    lngGene = -1
    Do While rsDicts.EOF = False
        lngGene = Val(rsDicts("gene"))
        strArrCnt = Split(rsDicts("used_count"), "-", 2, vbTextCompare)
        If UBound(strArrCnt) > 0 Then
            lngUsedCount = CLng(strArrCnt(1)) - CLng(strArrCnt(0))
            If lngUsedCount > lngGene Then
                strMsgLog = "自上次使用次数 " & strArrCnt(0) & " 至现在使用次数 " & strArrCnt(1) & " ，已累计超过约定因子 " & lngGene & " 。"
                Exit Do
            End If
        Else
            lngGene = 0
        End If
        rsDicts.MoveNext
    Loop
    rsDicts.Close
    Set rsDicts = Nothing
    If lngGene = 0 Then
        strMsgLog = "建议对数据库进行压缩，以减少数据库出错的风险。"
        m_blNeedFix = True
        m_blNeedCompareDB = True
    ElseIf strMsgLog <> "" Then
        strMsgLog = "建议对数据库进行压缩，以减少数据库出错的风险：" & vbCrLf & strMsgLog
        m_blNeedFix = True
        m_blNeedCompareDB = True
    End If
    CheckUsedCount = strMsgLog
End Function

