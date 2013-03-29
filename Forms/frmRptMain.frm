VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRptMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "报表配置"
   ClientHeight    =   5160
   ClientLeft      =   3675
   ClientTop       =   3240
   ClientWidth     =   7995
   Icon            =   "frmRptMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraService 
      Caption         =   "报表配置表"
      Height          =   4815
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   5895
      Begin MSDataGridLib.DataGrid grdReport 
         Bindings        =   "frmRptMain.frx":000C
         Height          =   4455
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
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
   End
   Begin VB.PictureBox picCommand 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5160
      Left            =   6165
      ScaleHeight     =   5160
      ScaleWidth      =   1830
      TabIndex        =   0
      Top             =   0
      Width           =   1830
      Begin VB.CommandButton cmdOper 
         Caption         =   "复制(&C)"
         Height          =   420
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "关 闭(&C)"
         Height          =   420
         Left            =   240
         TabIndex        =   6
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "删 除(&D)"
         Height          =   420
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "修 改(&M)"
         Height          =   420
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "增 加(&A)"
         Height          =   420
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "查 看(&Q)"
         Height          =   420
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc adoReport 
      Height          =   330
      Left            =   -240
      Top             =   2160
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
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
      Caption         =   "数据源"
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
End
Attribute VB_Name = "frmRptMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub adoReport_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    DisplayMsg "数据库错误", vbCritical
End Sub

Private Sub adoReport_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    '按纽显示控制
    Dim i As Integer
    If adoReport.Recordset.RecordCount = 0 Then
        For i = 0 To cmdOper.Count - 1
            cmdOper.Item(i).Enabled = False
        Next
        cmdOper.Item(1).Enabled = True
    Else
        For i = 0 To cmdOper.Count - 1
            cmdOper.Item(i).Enabled = True
        Next
        If adoReport.Recordset.EOF = True Then
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOper_Click(Index As Integer)
    If Index = OPER_ADD Then           '增加
        Call ShowMaintWin(Index)
        Exit Sub
    End If
    If adoReport.Recordset.RecordCount = 0 Then
        DisplayMsg "请选择所要操作的报表", vbExclamation
        cmdOper(Index).Enabled = False
        Exit Sub
    End If
    Select Case Index
    Case OPER_QUERY, OPER_MODIFY '查看和修改
        Call ShowMaintWin(Index)
    Case OPER_DEL  '删除
        Call DelReport
    Case OPER_COPY  '复制
        Call CopyReport
    End Select
End Sub

Private Sub Form_Load()
    grdReport.MarqueeStyle = dbgHighlightRow
    grdReport.AllowRowSizing = False
    
    adoReport.ConnectionString = gStrConnDB
    adoReport.CommandType = adCmdText
    adoReport.RecordSource = "select * from reports where valid_flag=1 order by modify_dt desc"

    Call RefreshData
End Sub

Private Sub grdReport_DblClick()
    Call cmdOper_Click(0)
End Sub

Private Sub RefreshData()
    adoReport.Refresh
    Call DisplayGrid(grdReport, "reports")
    On Error Resume Next
    grdReport.SetFocus
End Sub

Private Sub ShowMaintWin(ByVal Index As Integer)
    Dim i As Integer
    Dim strRptId As String
    Dim strRetMsg As String
    i = adoReport.Recordset.AbsolutePosition

    Load frmRptDetail
    strRptId = ""
    If Index <> OPER_ADD Then
        strRptId = adoReport.Recordset("rpt_id")
    End If
    If frmRptDetail.Init(Index, strRptId, strRetMsg) = False Then
        DisplayMsg "打开维护界面时出错，错误信息如下：" & vbCrLf & strRetMsg, vbExclamation
        Unload frmRptDetail
        Exit Sub
    End If
    frmRptDetail.Show vbModal
    Call RefreshData
    
    If adoReport.Recordset.RecordCount = 0 Then
        Exit Sub
    End If
    If i > adoReport.Recordset.RecordCount Then
        adoReport.Recordset.MoveLast
    Else
        adoReport.Recordset.Move i - 1, 1
    End If
End Sub

Private Sub DelReport()
    Dim rsRptData As New ADODB.Recordset
    Dim strSQL As String
    Dim iCount As Integer
    
    strSQL = "select count(*) from rpt_data where rpt_id='" & adoReport.Recordset("rpt_id") & "'"
    rsRptData.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    
    iCount = rsRptData(0)
    rsRptData.Close
    Set rsRptData = Nothing
    If iCount > 0 Then
        If DisplayMsg("此报表有历史数据，请确认是否删除？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    Else
        If DisplayMsg("请确认删除当前记录？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    End If

    Dim i As Integer
    i = adoReport.Recordset.AbsolutePosition
    gExecSql ("update reports set valid_flag=0, modify_dt=#" & Format(Date, "yyyy-MM-dd") & "# where rpt_id='" & adoReport.Recordset("rpt_id") & "'")
    Call RefreshData
    If adoReport.Recordset.RecordCount = 0 Then
        Exit Sub
    End If
    If i > adoReport.Recordset.RecordCount Then
        adoReport.Recordset.MoveLast
    Else
        adoReport.Recordset.Move i - 1, 1
    End If
End Sub

Private Sub CopyReport()
    Dim rsDicts As New ADODB.Recordset
    Dim strSQL As String
    Dim iSeq As Integer
    Dim strReportID As String
    Dim strReportName As String
    
    strSQL = "select dict_type, dict_value from dicts where dict_sect='OPT_REPORT' and dict_key='REPORT_ID'"
    rsDicts.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    iSeq = Val(rsDicts("dict_value")) + 1
    strReportID = rsDicts("dict_type") & Format(iSeq, RPT_ID_PATTERN)
    rsDicts.Close
    Set rsDicts = Nothing
    strReportName = InputBox("请输入报表名称：", "复制报表", adoReport.Recordset("rpt_name") & "-" & Format(Date, "yyyyMMdd"))
    If strReportName = "" Then
        Exit Sub
    End If
    
    On Error GoTo ErrHandle
    gAdoConnDB.BeginTrans
    strSQL = "insert into reports(rpt_id, rpt_name) values('" & strReportID & "','" & strReportName & "')"
    gAdoConnDB.Execute strSQL
    strSQL = "insert into rpt_field(rpt_id,field_name,display_name,serv_id,serv_period,serv_field) " & _
            "select '" & strReportID & "', field_name,display_name,serv_id,serv_period,serv_field from rpt_field where rpt_id='" & adoReport.Recordset("rpt_id") & "'"
    gAdoConnDB.Execute strSQL
    strSQL = "update dicts set dict_value='" & iSeq & "' where dict_sect='OPT_REPORT' and dict_key='REPORT_ID'"
    gAdoConnDB.Execute strSQL
    gAdoConnDB.CommitTrans
    Call RefreshData
    adoReport.Recordset.MoveFirst
    Exit Sub
    
ErrHandle:
    DisplayMsg "复制报表时出错，信息如下：", vbCritical
    On Error Resume Next
    gAdoConnDB.RollbackTrans
End Sub
