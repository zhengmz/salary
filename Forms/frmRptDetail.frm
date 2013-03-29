VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRptDetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "报表配置维护界面"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   Icon            =   "frmRptDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   4590
      TabIndex        =   10
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用"
      Height          =   375
      Left            =   5910
      TabIndex        =   9
      Top             =   4680
      Width           =   1095
   End
   Begin TabDlg.SSTab ssTabReport 
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "报表信息"
      TabPicture(0)   =   "frmRptDetail.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraReport"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "报表字段映射表"
      TabPicture(1)   =   "frmRptDetail.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdOper(4)"
      Tab(1).Control(1)=   "cmdOper(3)"
      Tab(1).Control(2)=   "cmdOper(2)"
      Tab(1).Control(3)=   "cmdOper(1)"
      Tab(1).Control(4)=   "cmdOper(0)"
      Tab(1).Control(5)=   "grdRptField"
      Tab(1).Control(6)=   "adoRptField"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "字段映射维护"
      TabPicture(2)   =   "frmRptDetail.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraRptField"
      Tab(2).Control(1)=   "adoServExt"
      Tab(2).Control(2)=   "adoServPeriod"
      Tab(2).Control(3)=   "adoServField"
      Tab(2).ControlCount=   4
      Begin VB.CommandButton cmdOper 
         Caption         =   "复 制"
         Height          =   420
         Index           =   4
         Left            =   -69480
         TabIndex        =   27
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "删 除"
         Height          =   420
         Index           =   3
         Left            =   -70770
         TabIndex        =   26
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "修 改"
         Height          =   420
         Index           =   2
         Left            =   -72060
         TabIndex        =   25
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "增 加"
         Height          =   420
         Index           =   1
         Left            =   -73350
         TabIndex        =   24
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "查 看"
         Height          =   420
         Index           =   0
         Left            =   -74640
         TabIndex        =   23
         Top             =   3720
         Width           =   855
      End
      Begin VB.Frame fraRptField 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   12
         Top             =   480
         Width           =   6255
         Begin MSDataListLib.DataCombo dcbServID 
            Bindings        =   "frmRptDetail.frx":0060
            DataField       =   "serv_id"
            DataSource      =   "adoRptField"
            Height          =   315
            Left            =   1680
            TabIndex        =   18
            Top             =   1560
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "serv_name"
            BoundColumn     =   "serv_id"
            Text            =   ""
         End
         Begin VB.TextBox txtDisplayName 
            Height          =   285
            Left            =   1680
            TabIndex        =   15
            Top             =   941
            Width           =   3495
         End
         Begin VB.TextBox txtFieldName 
            Height          =   285
            Left            =   1680
            TabIndex        =   13
            Top             =   315
            Width           =   2415
         End
         Begin MSDataListLib.DataCombo dcbServPeriod 
            Bindings        =   "frmRptDetail.frx":0079
            DataField       =   "serv_period"
            DataSource      =   "adoRptField"
            Height          =   315
            Left            =   1680
            TabIndex        =   19
            Top             =   2220
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "serv_period"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcbServField 
            Bindings        =   "frmRptDetail.frx":0095
            DataField       =   "serv_field"
            DataSource      =   "adoRptField"
            Height          =   315
            Left            =   1680
            TabIndex        =   20
            Top             =   2880
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "display_name"
            BoundColumn     =   "field_name"
            Text            =   ""
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "对应字段："
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   22
            Top             =   2940
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "对应条件："
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   21
            Top             =   2280
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "对应服务："
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   17
            Top             =   1620
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "显示名称："
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   16
            Top             =   986
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "字段名称："
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   14
            Top             =   360
            Width           =   900
         End
      End
      Begin MSDataGridLib.DataGrid grdRptField 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
      Begin VB.Frame fraReport 
         Height          =   3495
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   6015
         Begin MSAdodcLib.Adodc adoReport 
            Height          =   330
            Left            =   240
            Top             =   2280
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
         Begin VB.TextBox txtRptDesc 
            Height          =   1725
            Left            =   1680
            TabIndex        =   8
            Top             =   1440
            Width           =   3375
         End
         Begin VB.TextBox txtRptName 
            Height          =   285
            Left            =   1680
            TabIndex        =   6
            Top             =   855
            Width           =   3375
         End
         Begin VB.TextBox txtRptID 
            BackColor       =   &H80000011&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   4
            Top             =   315
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "报表描述："
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   7
            Top             =   1440
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "报表名称："
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   5
            Top             =   900
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "报表ID："
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   360
            Width           =   705
         End
      End
      Begin MSAdodcLib.Adodc adoRptField 
         Height          =   330
         Left            =   -74880
         Top             =   2400
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
      Begin MSAdodcLib.Adodc adoServExt 
         Height          =   330
         Left            =   -74880
         Top             =   2040
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
      Begin MSAdodcLib.Adodc adoServPeriod 
         Height          =   330
         Left            =   -74880
         Top             =   2400
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
      Begin MSAdodcLib.Adodc adoServField 
         Height          =   330
         Left            =   -74880
         Top             =   2760
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
   End
End
Attribute VB_Name = "frmRptDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_strReportID As String
Private m_intCmdType As Integer

Public Function Init(ByVal pmIntCmdType As Integer, ByVal pmStrRptID As String, pmStrRetMsg As String) As Boolean
    On Error GoTo ErrHandle
    
    Init = False
    m_intCmdType = pmIntCmdType
    m_strReportID = pmStrRptID
    pmStrRetMsg = ""

    adoReport.RecordSource = "select * from reports where rpt_id='" & m_strReportID & "'"
    adoReport.Refresh
    Select Case m_intCmdType
    Case OPER_QUERY  '查看
        txtRptID.Enabled = False
        txtRptID.BackColor = &H80000011
        txtRptName.Enabled = False
        txtRptName.BackColor = &H80000011
        txtRptDesc.Enabled = False
        txtRptDesc.BackColor = &H80000011
        txtFieldName.Enabled = False
        txtFieldName.BackColor = &H80000011
        txtDisplayName.Enabled = False
        txtDisplayName.BackColor = &H80000011
        dcbServID.Enabled = False
        dcbServID.BackColor = &H80000011
        dcbServPeriod.Enabled = False
        dcbServPeriod.BackColor = &H80000011
        dcbServField.Enabled = False
        dcbServField.BackColor = &H80000011
    Case OPER_ADD '增加
        adoReport.Recordset.AddNew
        Dim rsDicts As New ADODB.Recordset
        Dim strSQL As String
        
        strSQL = "select dict_type, dict_value from dicts where dict_sect='OPT_REPORT' and dict_key='REPORT_ID'"
        rsDicts.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
        m_strReportID = rsDicts("dict_type") & Format(Val(rsDicts("dict_value")) + 1, RPT_ID_PATTERN)
        rsDicts.Close
        Set rsDicts = Nothing
        txtRptID.Text = m_strReportID
    Case OPER_MODIFY  '修改
    Case Else
        Init = False
        pmStrRetMsg = "非法进入本页面，传入参数为 " & pmIntCmdType & " , " & pmStrRptID
        Exit Function
    End Select
    Call RefreshGrid
    
    adoServExt.RecordSource = "select * from v_serv_ext where serv_id<>'" & m_strReportID & "'"
    adoServExt.Refresh
    If adoRptField.Recordset.RecordCount > 0 Then
        Call RefreshServ(IIf(IsNull(adoRptField.Recordset("serv_id")), "", adoRptField.Recordset("serv_id")))
    End If

    Init = True
    Exit Function
    
ErrHandle:
    pmStrRetMsg = "错误号：" & Err.Number & vbCrLf & _
                  "错误内容：" & Err.Description
End Function

Private Sub adoReport_WillChangeField(ByVal cFields As Long, Fields As Variant, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If cFields > 0 Then
        cmdApply.Enabled = True
    End If
End Sub

Private Sub adoRptField_WillChangeField(ByVal cFields As Long, Fields As Variant, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If cFields > 0 Then
        cmdApply.Enabled = True
    End If
End Sub

Private Sub cmdApply_Click()
    On Error GoTo ErrHandle

    Select Case ssTabReport.Tab
    Case 0  'update reports
        gAdoConnDB.BeginTrans
        adoReport.Recordset("modify_dt") = Date
        adoReport.Recordset.Update
        If m_intCmdType = OPER_ADD Then
            Dim strSQL As String
            Dim iSeq As Integer
            iSeq = Val(Right(m_strReportID, Len(RPT_ID_PATTERN)))
            strSQL = "update dicts set dict_value='" & iSeq & "' where dict_sect='OPT_REPORT' and dict_key='REPORT_ID'"
            gAdoConnDB.Execute strSQL
        End If
        gAdoConnDB.CommitTrans
    Case 2  'update rpt_field
        Dim i As Integer
        i = adoRptField.Recordset.AbsolutePosition
        adoRptField.Recordset.Update
        adoRptField.Refresh
        Call RefreshGrid
        If adoRptField.Recordset.RecordCount = 0 Then
            Exit Sub
        End If
        If i > adoRptField.Recordset.RecordCount Then
            adoRptField.Recordset.MoveLast
        Else
            adoRptField.Recordset.Move i - 1, 1
        End If
    End Select
    
NormalExit:
    cmdApply.Enabled = False
    Exit Sub
    
ErrHandle:
    DisplayMsg "保存更新时出错，将取消更新，错误信息如下，：", vbCritical
    On Error Resume Next
    If ssTabReport.Tab = 0 Then
        gAdoConnDB.RollbackTrans
    Else
        adoRptField.Recordset.Cancel
    End If
    GoTo NormalExit
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    If cmdApply.Enabled = True Then
        adoReport.Recordset.Cancel
        adoRptField.Recordset.Cancel
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If cmdApply.Enabled = True Then
        Call cmdApply_Click
    End If
    If cmdApply.Enabled = False Then    '更新成功后便可退出
        Unload Me
    End If
End Sub

Private Sub cmdOper_Click(Index As Integer)
    Dim iSeq As Integer

    Select Case Index
    Case OPER_QUERY, OPER_MODIFY
        ssTabReport.Tab = 2
    Case OPER_ADD
        ssTabReport.Tab = 2
        iSeq = 1
        If adoRptField.Recordset.RecordCount > 0 Then
            adoRptField.Recordset.MoveLast
            iSeq = Val(Right(adoRptField.Recordset("field_name"), Len(RPT_FIELD_PATTERN))) + 1
        End If
        adoRptField.Recordset.AddNew
        adoRptField.Recordset("rpt_id") = m_strReportID
        If iSeq <= RPT_FIELD_CNT_MAX Then
            adoRptField.Recordset("field_name") = "field" & Format(iSeq, RPT_FIELD_PATTERN)
        End If
    Case OPER_DEL
        If DisplayMsg("请确认是否删除当前记录？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
        adoRptField.Recordset.Delete
    Case OPER_COPY
        Dim strFieldName As String
        Dim strDisplayName As String
        Dim strServId As String
        Dim strServPeriod As String
        Dim strServField As String
        strDisplayName = adoRptField.Recordset("display_name")
        strServId = IIf(IsNull(adoRptField.Recordset("serv_id")), "", adoRptField.Recordset("serv_id"))
        strServPeriod = adoRptField.Recordset("serv_period")
        strServField = adoRptField.Recordset("serv_field")
        adoRptField.Recordset.MoveLast
        iSeq = Val(Right(adoRptField.Recordset("field_name"), Len(RPT_FIELD_PATTERN))) + 1
        ssTabReport.Tab = 2
        adoRptField.Recordset.AddNew
        adoRptField.Recordset("rpt_id") = m_strReportID
        If iSeq <= RPT_FIELD_CNT_MAX Then
            adoRptField.Recordset("field_name") = "field" & Format(iSeq, RPT_FIELD_PATTERN)
        End If
        adoRptField.Recordset("display_name") = strDisplayName
        adoRptField.Recordset("serv_id") = strServId
        adoRptField.Recordset("serv_period") = strServPeriod
        adoRptField.Recordset("serv_field") = strServField
    End Select
End Sub

Private Sub dcbServID_Validate(Cancel As Boolean)
    If dcbServID.BoundText = adoRptField.Recordset("serv_id") Then
        Exit Sub
    End If

    adoRptField.Recordset("serv_period") = ""
    adoRptField.Recordset("serv_field") = ""
    Call RefreshServ(dcbServID.BoundText)
End Sub

Private Sub Form_Load()
    '初始化变量
    m_strReportID = ""
    m_intCmdType = -1
    
    '初始化控件
    ssTabReport.Tab = 0
    cmdApply.Enabled = False
    adoReport.ConnectionString = gStrConnDB
    adoReport.CommandType = adCmdText
    adoReport.RecordSource = "select * from reports where valid_flag=1"

    adoRptField.ConnectionString = gStrConnDB
    adoRptField.CommandType = adCmdText
    adoRptField.RecordSource = "select * from rpt_field"
    
    adoServExt.ConnectionString = gStrConnDB
    adoServExt.CommandType = adCmdText
    adoServExt.RecordSource = "select * from v_serv_ext"
    adoServExt.Refresh
    
    adoServPeriod.ConnectionString = gStrConnDB
    adoServPeriod.CommandType = adCmdText
    
    adoServField.ConnectionString = gStrConnDB
    adoServField.CommandType = adCmdText
    
    Set txtRptID.DataSource = adoReport
    txtRptID.DataField = "rpt_id"
    Set txtRptName.DataSource = adoReport
    txtRptName.DataField = "rpt_name"
    Set txtRptDesc.DataSource = adoReport
    txtRptDesc.DataField = "rpt_desc"
    grdRptField.MarqueeStyle = dbgHighlightRow
    Set grdRptField.DataSource = adoRptField
    Set txtFieldName.DataSource = adoRptField
    txtFieldName.DataField = "field_name"
    Set txtDisplayName.DataSource = adoRptField
    txtDisplayName.DataField = "display_name"

    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub ssTabReport_Click(PreviousTab As Integer)
    If cmdApply.Enabled = True Then     '保存更新
        Call cmdApply_Click
    End If
End Sub

Private Sub RefreshGrid()
    adoRptField.RecordSource = "select * from rpt_field where rpt_id='" & m_strReportID & "' order by field_name"
    adoRptField.Refresh
    Call DisplayGrid(grdRptField, "rpt_field")
    
    Dim i As Integer

    If adoRptField.Recordset.RecordCount = 0 Then
        For i = 0 To cmdOper.Count - 1
            cmdOper.Item(i).Enabled = False
        Next
        If m_intCmdType <> OPER_QUERY Then
            cmdOper.Item(1).Enabled = True
        End If
    Else
        If m_intCmdType = OPER_QUERY Then
            For i = 0 To cmdOper.Count - 1
                cmdOper.Item(i).Enabled = False
            Next
            cmdOper.Item(0).Enabled = True
        Else
            For i = 0 To cmdOper.Count - 1
                cmdOper.Item(i).Enabled = True
            Next
        End If
    End If
End Sub

Private Sub RefreshServ(ByVal pmStrServId As String)
    adoServPeriod.RecordSource = "select distinct serv_period from salary where serv_id='" & pmStrServId & "' order by serv_period desc"
    adoServPeriod.Refresh
    adoServField.RecordSource = "select * from v_serv_field_ext where serv_id='" & pmStrServId & "' order by field_name"
    adoServField.Refresh
    If Not IsNull(adoRptField.Recordset("serv_field")) Then
        dcbServField.BoundText = adoRptField.Recordset("serv_field")
    End If
End Sub
