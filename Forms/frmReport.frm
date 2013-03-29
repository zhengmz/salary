VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmReport 
   Caption         =   "报表生成"
   ClientHeight    =   4635
   ClientLeft      =   315
   ClientTop       =   1725
   ClientWidth     =   9075
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4635
   ScaleWidth      =   9075
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   9075
      TabIndex        =   2
      Top             =   0
      Width           =   9075
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "关 闭(&C)"
         Height          =   360
         Left            =   7800
         TabIndex        =   6
         Top             =   217
         Width           =   1095
      End
      Begin VB.CommandButton cmdReBuild 
         Caption         =   "重新生成"
         Height          =   360
         Left            =   6480
         TabIndex        =   5
         Top             =   217
         Width           =   1095
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "导出(&E)"
         Height          =   360
         Left            =   5160
         TabIndex        =   4
         Top             =   217
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dcbReport 
         Bindings        =   "frmReport.frx":212A
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "报表名称："
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   900
      End
   End
   Begin MSDataGridLib.DataGrid grdRptData 
      Bindings        =   "frmReport.frx":2142
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2143
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   8454016
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoRptData 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   4305
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "无数据"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoReport 
      Height          =   330
      Left            =   4800
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dgFile 
      Left            =   5400
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_strReportID As String
Private m_strReportName As String
Private m_strRptPattern As String   '报表ID前缀
Private m_strCntPattern As String   '汇总字段名称

Private Sub adoRptData_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If adoRptData.Recordset.RecordCount = 0 Then
        adoRptData.Caption = "无记录"
    Else
        If adoRptData.Recordset.EOF = False Then
            adoRptData.Caption = "当前记录位置: " & adoRptData.Recordset.AbsolutePosition & _
                         "/" & adoRptData.Recordset.RecordCount
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    On Error GoTo Err_Handle

    Dim strFileName As String
    With dgFile
        .CancelError = True
        .Filter = "导出Excel文件 (*.xls)|*.xls"
        .DefaultExt = "xls"
        .DialogTitle = "导出报表数据"
        .Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
        .ShowSave
        strFileName = .FileName
    End With

    If Dir(strFileName) <> "" Then
        Kill (strFileName)
    End If
    
    Dim rsRptField As New ADODB.Recordset
    Dim strSql As String

    strSql = "select field_name, display_name from rpt_field where rpt_id='" & m_strReportID & "' order by field_name"
    rsRptField.Open strSql, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    strSql = "select emp_id as 员工编码, emp_name as 姓名"
    Do Until rsRptField.EOF
        strSql = strSql & "," & rsRptField("field_name") & " as " & rsRptField("display_name")
        rsRptField.MoveNext
    Loop
    rsRptField.Close
    Set rsRptField = Nothing

    strSql = strSql & " into [Excel 8.0;database=" & strFileName & "].[" & m_strReportName & "] from rpt_data where rpt_id='" & m_strReportID & "'"
    If gExecSql(strSql) = False Then
        DisplayMsg "导出失败。", vbExclamation
    Else
        DisplayMsg "成功导出到文件 " & strFileName & "！", vbInformation
    End If
    Exit Sub

Err_Handle:
    If Err.Number = 32755 Then
    '按了取消
        Exit Sub
    End If
    DisplayMsg "导出失败！", vbCritical
End Sub

Private Sub cmdReBuild_Click()
    If adoRptData.Recordset.RecordCount > 0 Then
        If DisplayMsg("将覆盖现有数据，请确认是否重新生成？", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If

    On Error GoTo ErrHandle

    Dim rsRptField As New ADODB.Recordset
    Dim strSql As String
    Dim strSumSQL As String
    Dim strArrSQL() As String
    Dim iCount As Integer

    strSql = "select * from rpt_field where rpt_id='" & m_strReportID & "' order by field_name"
    rsRptField.Open strSql, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    iCount = rsRptField.RecordCount
    ReDim strArrSQL(iCount)
    If iCount > 0 Then
        ReDim strArrSQL(iCount - 1)
    End If

    Dim iPos As Integer
    Dim strErrMsg As String
    Dim strFieldName As String
    Dim strServId As String
    Dim strServPeriod As String
    Dim strServField As String

    iPos = 0
    strSumSQL = "0"
    strErrMsg = ""
    Do Until rsRptField.EOF
        strFieldName = rsRptField("field_name")
        strServId = IIf(IsNull(rsRptField("serv_id")), "", rsRptField("serv_id"))
        strServPeriod = IIf(IsNull(rsRptField("serv_period")), "", rsRptField("serv_period"))
        strServField = IIf(IsNull(rsRptField("serv_field")), "", rsRptField("serv_field"))
        If strServId = "" Then
            strErrMsg = strErrMsg & vbCrLf & iPos + 1 & ": 字段 '" & strFieldName & "' 没有配置"
        '汇总字段
        ElseIf strServId = m_strCntPattern Then
            If iPos = iCount - 1 Then   '应在最后一列
                strArrSQL(iPos) = "UPDATE rpt_data SET field04 = " & strSumSQL & _
                                  " WHERE rpt_id = '" & m_strReportID & "'"
                strArrSQL(iPos) = "update rpt_data set " & strFieldName & "=" & strSumSQL & " where rpt_id='" & m_strReportID & "'"
            Else
                strErrMsg = strErrMsg & vbCrLf & iPos + 1 & ": 汇总字段 '" & strFieldName & "' 没有设在最后一列"
            End If
        '报表字段，读rpt_data
        ElseIf m_strRptPattern <> "" And Mid(strServId, 1, Len(m_strRptPattern)) = m_strRptPattern Then
            If strServField <> "" Then
                strArrSQL(iPos) = "UPDATE rpt_data INNER JOIN rpt_data AS rpt_data_1 ON rpt_data.emp_id=rpt_data_1.emp_id " & _
                                  "SET rpt_data." & strFieldName & " = rpt_data_1." & strServField & _
                                  " WHERE rpt_data.rpt_id='" & m_strReportID & "' AND rpt_data_1.rpt_id='" & strServId & "'"
            Else
                strErrMsg = strErrMsg & vbCrLf & iPos + 1 & ": 报表字段 '" & strFieldName & "' 没有对应的字段配置"
            End If
        '服务字段，读salary
        Else
            If strServField <> "" And strServPeriod <> "" Then
                strArrSQL(iPos) = "UPDATE salary INNER JOIN rpt_data ON salary.emp_id = rpt_data.emp_id " & _
                                  "SET rpt_data." & strFieldName & " = salary." & strServField & _
                                  " WHERE rpt_data.rpt_id='" & m_strReportID & "' AND salary.serv_period='" & strServPeriod & "' AND salary.serv_id='" & strServId & "'"
            Else
                strErrMsg = strErrMsg & vbCrLf & iPos + 1 & ": 服务字段 '" & strFieldName & "' 没有对应的条件或字段配置"
            End If
        End If
        strSumSQL = strSumSQL & "+IIF(ISNULL(" & strFieldName & "),0,VAL(" & strFieldName & "))"
        iPos = iPos + 1
        rsRptField.MoveNext
    Loop
    rsRptField.Close
    Set rsRptField = Nothing
    If iPos = 0 Then
        strErrMsg = strErrMsg & vbCrLf & "报表配置为空，无法生成"
    End If

    If strErrMsg <> "" Then
        DisplayMsg "报表配置验证失败，具体信息如下：" & vbCrLf & strErrMsg, vbExclamation
        Exit Sub
    End If
    
    On Error GoTo TransErrHandle
    gAdoConnDB.BeginTrans
    strSql = "delete from rpt_data where rpt_id='" & m_strReportID & "'"
    gAdoConnDB.Execute strSql
    strSql = "insert into rpt_data(rpt_id,emp_id,emp_name) " & _
             "select '" & m_strReportID & "',emp_id,emp_name from emp"
    gAdoConnDB.Execute strSql
    For iPos = 0 To iCount - 1
        gAdoConnDB.Execute strArrSQL(iPos)
    Next
    gAdoConnDB.CommitTrans
    Call RefreshGrid
    
    Exit Sub

ErrHandle:
    DisplayMsg "重新生成失败！", vbCritical
    Exit Sub
    
TransErrHandle:
    DisplayMsg "重新生成失败！出错参考语句：" & vbCrLf & strArrSQL(iPos), vbCritical
    gAdoConnDB.RollbackTrans
End Sub

Private Sub dcbReport_Change()
    If dcbReport.BoundText = m_strReportID Then
        Exit Sub
    End If
    cmdExport.Enabled = True
    cmdReBuild.Enabled = True
    m_strReportID = dcbReport.BoundText
    m_strReportName = dcbReport.Text
    Call RefreshGrid
End Sub

Private Sub Form_Load()
    '初始化变量
    m_strReportID = ""
    m_strReportName = ""
    
    Call ReadOption

    '初始化控件
    cmdExport.Enabled = False
    cmdReBuild.Enabled = False
    grdRptData.MarqueeStyle = dbgHighlightRow

    adoRptData.ConnectionString = gStrConnDB
    adoRptData.CommandType = adCmdText

    adoReport.ConnectionString = gStrConnDB
    adoReport.CommandType = adCmdText
    adoReport.RecordSource = "select rpt_id,rpt_name from reports where valid_flag=1 order by modify_dt desc"
    adoReport.Refresh
    dcbReport.ListField = "rpt_name"
    dcbReport.BoundColumn = "rpt_id"
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  '当窗体调整时会调整网格
  grdRptData.Left = 0
  grdRptData.Width = Me.ScaleWidth
  grdRptData.Height = Me.ScaleHeight - adoRptData.Height - picTop.Height
  grdRptData.Top = picTop.Height
End Sub

Private Sub RefreshGrid()
    If m_strReportID = "" Then
        adoRptData.Recordset.Close
        Exit Sub
    End If
    adoRptData.RecordSource = "select * from rpt_data where rpt_id='" & m_strReportID & "' order by emp_id"
    adoRptData.Refresh
    DisplayGrid grdRptData, "rpt_data", m_strReportID
End Sub

Private Sub ReadOption()
    Dim rsDicts As New ADODB.Recordset
    Dim strSql As String

    strSql = "select dict_key, dict_type as pattern from dicts where dict_sect='OPT_REPORT' and dict_key in ('REPORT_ID','SERV_ID')"
    rsDicts.Open strSql, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText

    m_strRptPattern = ""
    m_strCntPattern = ""
    Do Until rsDicts.EOF
        If rsDicts("dict_key") = "REPORT_ID" Then
            m_strRptPattern = rsDicts("pattern")
        ElseIf rsDicts("dict_key") = "SERV_ID" Then
            m_strCntPattern = rsDicts("pattern")
        End If
        rsDicts.MoveNext
    Loop
    rsDicts.Close
    Set rsDicts = Nothing
End Sub
