VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "综合查询"
   ClientHeight    =   5415
   ClientLeft      =   2955
   ClientTop       =   2490
   ClientWidth     =   7500
   Icon            =   "frmQuery.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab ssTabQuery 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "查询条件"
      TabPicture(0)   =   "frmQuery.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "adoEmp"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "adoServType"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "grdResult"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "adoResult"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraQuery"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "结果明细"
      TabPicture(1)   =   "frmQuery.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "adoSalary"
      Tab(1).Control(1)=   "grdSalary"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraQuery 
         Height          =   1815
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   6495
         Begin VB.CommandButton cmdQuery 
            Caption         =   "查 询(&Q)"
            Default         =   -1  'True
            Height          =   375
            Left            =   3120
            TabIndex        =   5
            Top             =   1260
            Width           =   1215
         End
         Begin VB.CommandButton cmdClose 
            Cancel          =   -1  'True
            Caption         =   "关 闭(&C)"
            Height          =   375
            Left            =   4920
            TabIndex        =   4
            Top             =   1260
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dcbServType 
            Bindings        =   "frmQuery.frx":0044
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "serv_type"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcbEmpId 
            Bindings        =   "frmQuery.frx":005E
            Height          =   315
            Left            =   3720
            TabIndex        =   7
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "emp_name"
            BoundColumn     =   "emp_id"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker dtBegin 
            Height          =   315
            Left            =   1320
            TabIndex        =   8
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   25493507
            CurrentDate     =   39069
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker dtEnd 
            Height          =   315
            Left            =   3720
            TabIndex        =   9
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   25493507
            CurrentDate     =   39069
            MinDate         =   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "单据分类："
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   13
            Top             =   300
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "员工："
            Height          =   195
            Index           =   1
            Left            =   3120
            TabIndex        =   12
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "查询期间："
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   11
            Top             =   780
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   195
            Index           =   3
            Left            =   3120
            TabIndex        =   10
            Top             =   780
            Width           =   180
         End
      End
      Begin MSAdodcLib.Adodc adoResult 
         Height          =   330
         Left            =   360
         Top             =   4560
         Width           =   6495
         _ExtentX        =   11456
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
         Caption         =   "无记录"
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
      Begin MSDataGridLib.DataGrid grdResult 
         Bindings        =   "frmQuery.frx":0073
         Height          =   2295
         Left            =   360
         TabIndex        =   2
         Top             =   2280
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4048
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
      Begin MSAdodcLib.Adodc adoSalary 
         Height          =   375
         Left            =   -74640
         Top             =   4440
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
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
         Caption         =   "无记录"
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
      Begin MSDataGridLib.DataGrid grdSalary 
         Bindings        =   "frmQuery.frx":008B
         Height          =   3855
         Left            =   -74640
         TabIndex        =   1
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6800
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
      Begin MSAdodcLib.Adodc adoServType 
         Height          =   330
         Left            =   0
         Top             =   2880
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
      Begin MSAdodcLib.Adodc adoEmp 
         Height          =   330
         Left            =   0
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_strServType As String
Private m_strEmpId As String
Private m_dtBegin As Date
Private m_dtEnd As Date

Private Sub adoResult_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If adoResult.Recordset.RecordCount = 0 Then
        adoResult.Caption = "无记录"
    Else
        If adoResult.Recordset.EOF = False Then
            adoResult.Caption = "当前记录位置: " & adoResult.Recordset.AbsolutePosition & _
                         "/" & adoResult.Recordset.RecordCount
        End If
    End If
End Sub

Private Sub adoSalary_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If adoSalary.Recordset.RecordCount = 0 Then
        adoSalary.Caption = "无记录"
    Else
        If adoSalary.Recordset.EOF = False Then
            adoSalary.Caption = "当前记录位置: " & adoSalary.Recordset.AbsolutePosition & _
                         "/" & adoSalary.Recordset.RecordCount
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    m_strServType = dcbServType.BoundText
    m_strEmpId = dcbEmpId.BoundText
    m_dtBegin = dtBegin.value
    m_dtEnd = dtEnd.value
    
    If m_dtBegin > m_dtEnd Then
        DisplayMsg "起始时间大于终止时间", vbExclamation
        Exit Sub
    End If
    
    Dim strSQLYear As String
    Dim strSQLMonth As String
    Dim strSQLDay As String
    
    strSQLDay = "SELECT DISTINCT services.serv_type, salary.serv_id, services.serv_name, salary.serv_period " & _
                 "FROM salary INNER JOIN services ON salary.serv_id = services.serv_id "
    
    strSQLYear = strSQLDay & "WHERE services.serv_period='YEAR' " & _
                "AND salary.serv_period>='" & Format(m_dtBegin, DATE_FORMAT_YEAR) & "' " & _
                "AND salary.serv_period<='" & Format(m_dtEnd, DATE_FORMAT_YEAR) & "'"
    
    strSQLMonth = strSQLDay & "WHERE services.serv_period='MONTH' " & _
                "AND salary.serv_period>='" & Format(m_dtBegin, DATE_FORMAT_MONTH) & "' " & _
                "AND salary.serv_period<='" & Format(m_dtEnd, DATE_FORMAT_MONTH) & "'"
    
    strSQLDay = strSQLDay & "WHERE services.serv_period='DAY' " & _
                "AND salary.serv_period>='" & Format(m_dtBegin, DATE_FORMAT_DAY) & "' " & _
                "AND salary.serv_period<='" & Format(m_dtEnd, DATE_FORMAT_DAY) & "'"
    
    If m_strServType <> "" Then
        strSQLYear = strSQLYear & " AND services.serv_type='" & m_strServType & "'"
        strSQLMonth = strSQLMonth & " AND services.serv_type='" & m_strServType & "'"
        strSQLDay = strSQLDay & " AND services.serv_type='" & m_strServType & "'"
    End If
    If m_strEmpId <> "" Then
        strSQLYear = strSQLYear & "AND salary.emp_id='" & m_strEmpId & "'"
        strSQLMonth = strSQLMonth & "AND salary.emp_id='" & m_strEmpId & "'"
        strSQLDay = strSQLDay & "AND salary.emp_id='" & m_strEmpId & "'"
    End If
    
    adoResult.RecordSource = strSQLYear & " union " & strSQLMonth & " union " & strSQLDay
    adoResult.Refresh
    DisplayGrid grdResult, "SERV_QUERY"
End Sub

Private Sub Form_Load()
    '初始化控件
    dtBegin.value = Date
    dtBegin.Day = 1
    dtEnd.value = Date
    grdResult.MarqueeStyle = dbgHighlightRow
    grdSalary.MarqueeStyle = dbgHighlightRow
    
    adoResult.ConnectionString = gStrConnDB
    adoResult.CommandType = adCmdText

    adoSalary.ConnectionString = gStrConnDB
    adoSalary.CommandType = adCmdText
    
    adoServType.ConnectionString = gStrConnDB
    adoServType.CommandType = adCmdText
    adoServType.RecordSource = "select dict_key as serv_type from dicts where dict_sect='OPT_SERV_TYPE' order by dict_flag"
    adoServType.Refresh
    
    adoEmp.ConnectionString = gStrConnDB
    adoEmp.CommandType = adCmdText
    adoEmp.RecordSource = "select emp_id, emp_id + '  ' + emp_name as emp_name from emp order by emp_id"
    adoEmp.Refresh
End Sub

Private Sub grdResult_DblClick()
    ssTabQuery.Tab = 1
End Sub

Private Sub ssTabQuery_Click(PreviousTab As Integer)
    If ssTabQuery.Tab <> 1 Then
        Exit Sub
    End If
    If adoResult.RecordSource = "" Then
        Exit Sub
    ElseIf adoResult.Recordset.RecordCount = 0 Then
        If adoSalary.RecordSource = "" Then
            Exit Sub
        End If
        If adoSalary.Recordset.State = adStateOpen Then
            adoSalary.Caption = "无记录"
            adoSalary.Recordset.Close
        End If
        Exit Sub
    End If

    Dim strSQL As String
    
    strSQL = "select * from salary where serv_id='" & adoResult.Recordset(1) & "' " & _
             "and serv_period='" & adoResult.Recordset(3) & "'"
             
    If m_strEmpId <> "" Then
        strSQL = strSQL & " and emp_id='" & m_strEmpId & "'"
    End If
    
    If strSQL = adoSalary.RecordSource Then
        Exit Sub
    End If
    
    adoSalary.RecordSource = strSQL
    adoSalary.Refresh
    DisplayGrid grdSalary, "salary", adoResult.Recordset(1)
End Sub
