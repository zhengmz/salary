VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalary 
   Caption         =   "数据导入"
   ClientHeight    =   4635
   ClientLeft      =   315
   ClientTop       =   1725
   ClientWidth     =   8880
   Icon            =   "frmSalary.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4635
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   8880
      TabIndex        =   6
      Top             =   0
      Width           =   8880
      Begin MSComCtl2.DTPicker dtServPeriod 
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyyMM"
         Format          =   21233667
         CurrentDate     =   39069
         MinDate         =   2
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "清 除(&D)"
         Height          =   360
         Left            =   4320
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "导 入(&L)"
         Height          =   360
         Left            =   4320
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.Frame fraSelect 
         Caption         =   "请选择模板"
         Height          =   735
         Left            =   360
         TabIndex        =   9
         Top             =   120
         Width           =   8175
         Begin VB.ComboBox cbServType 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo dcbService 
            Bindings        =   "frmSalary.frx":212A
            Height          =   315
            Left            =   3480
            TabIndex        =   2
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lbType 
            AutoSize        =   -1  'True
            Caption         =   "类型："
            Height          =   315
            Left            =   240
            TabIndex        =   11
            Top             =   270
            Width           =   540
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "名称："
            Height          =   315
            Left            =   2760
            TabIndex        =   10
            Top             =   270
            Width           =   540
         End
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "发 送(&S)"
         Height          =   360
         Left            =   5760
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "关 闭(&C)"
         Height          =   360
         Left            =   7200
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cbServPeriod 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lbPeriod 
         AutoSize        =   -1  'True
         Caption         =   "月份："
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   1020
         Width           =   540
      End
      Begin VB.Label lbPeriodFormat 
         AutoSize        =   -1  'True
         Caption         =   "（YYYY-MM）"
         Height          =   195
         Left            =   2880
         TabIndex        =   7
         Top             =   1020
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid tbSalary 
      Bindings        =   "frmSalary.frx":2143
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   1800
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
   Begin MSAdodcLib.Adodc adoSalary 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   4305
      Width           =   8880
      _ExtentX        =   15663
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
   Begin MSAdodcLib.Adodc adoService 
      Height          =   330
      Left            =   4440
      Top             =   1800
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Caption         =   "Service"
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
Attribute VB_Name = "frmSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_strTableName As String
Private m_strServId As String
Private m_strServType As String
Private m_strServPeriod As String
Private m_strServSheet As String
Private m_strServPeriodMethod As String

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

Private Sub cbServPeriod_Validate(Cancel As Boolean)
    If cbServPeriod.Text = m_strServPeriod Then
        Exit Sub
    End If

    m_strServPeriod = cbServPeriod.Text

    Call RefreshGrid
End Sub

Private Sub cbServType_LostFocus()
    If cbServType.Text = m_strServType Then
        Exit Sub
    End If
    m_strServType = cbServType.Text

    If m_strServType = "" Or m_strServType = "全部" Then
        adoService.RecordSource = "select * from services where valid_flag=1 order by modify_dt desc"
    Else
        adoService.RecordSource = "select * from services where valid_flag=1 and serv_type='" & m_strServType & "' order by modify_dt desc"
    End If
    adoService.Refresh
    dcbService.BoundText = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    If m_strTableName = "" Then
        DisplayMsg "请选择要被导入的表！", vbExclamation
        Exit Sub
    End If
    If m_strServId = "" Then
        DisplayMsg "请选择相应的模板!", vbExclamation
        dcbService.SetFocus
        Exit Sub
    End If
    If m_strServPeriod = "" Then
        DisplayMsg "输入有效的时间格式!", vbExclamation
        cbServPeriod.SetFocus
        Exit Sub
    End If

    Dim strSQL As String
    If DisplayMsg("请确认是否删除这些数据？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        strSQL = "delete from salary where serv_period='" & m_strServPeriod & _
                 "' and serv_id='" & m_strServId & "'"
        Call gExecSql(strSQL)
        Call RefreshGrid
    End If
End Sub

Private Sub cmdLoad_Click()
    If m_strTableName = "" Then
        DisplayMsg "请选择要被导入的表！", vbExclamation
        Exit Sub
    End If
    If m_strServId = "" Then
        DisplayMsg "请选择相应的模板!", vbExclamation
        dcbService.SetFocus
        Exit Sub
    End If
    If m_strServPeriod = "" Then
        DisplayMsg "输入有效的时间格式!", vbExclamation
        dtServPeriod.SetFocus
        Exit Sub
    End If
    
    Dim strFileName As String
    Dim intFileFormat As Integer  '导入的文件格式

    On Error GoTo ErrHandle
    With dgFile
        .FileName = ""
        .Filter = "数据文件 (*.xls)|*.xls|数据文件 (*.csv)|*.csv"
        .DialogTitle = "打开数据文件"
        .CancelError = True
        .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNFileMustExist
        .ShowOpen
        strFileName = .FileName
        intFileFormat = .FilterIndex
    End With
    
    Dim strSQL As String
    If adoSalary.Recordset.RecordCount > 0 Then
        If DisplayMsg("存在旧数据，是否追加？" & vbCrLf & vbCrLf & "选择[是]将追加新数据" & vbCrLf & "选择[否]将先删除旧数据，后再导入新数据", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            strSQL = "delete from salary where serv_period='" & m_strServPeriod & _
                     "' and serv_id='" & m_strServId & "'"
            Call gExecSql(strSQL)
            Call RefreshGrid
        End If
    End If

    Dim strRetMsg As String
    Dim iFieldCount As Integer
    Dim rsServConf As New ADODB.Recordset
    Dim strSourceRS As String
    Dim strTargetRS As String

    strSQL = "Select field_name,display_name From serv_field" & _
            " Where serv_id='" & m_strServId & "' AND valid_flag = 1"

    rsServConf.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    If rsServConf.RecordCount = 0 Then
        DisplayMsg "没有找到对应的字段映射关系，请检查！服务配置为 '" & dcbService.Text & "'"
        rsServConf.Close
        Set rsServConf = Nothing
        Exit Sub
    End If
    iFieldCount = rsServConf.RecordCount + 2

    strTargetRS = "select serv_id, serv_period"
    strSourceRS = "select '" & m_strServId & "', '" & m_strServPeriod & "'"
    Do Until rsServConf.EOF
        strTargetRS = strTargetRS & ", " & rsServConf("field_name")
        strSourceRS = strSourceRS & ", [" & rsServConf("display_name") & "]"
        rsServConf.MoveNext
    Loop
    rsServConf.Close
    Set rsServConf = Nothing
    strTargetRS = strTargetRS & " from " & m_strTableName
    If intFileFormat = 1 Then   'Excel 文件
        '说明：HDR=表示数据有无标题行(yes/no)，
        'IMEX=1通知驱动程序始终将“互混”数据列作为文本读取，
        '    在这里需要清楚的是，系统在判断该字段(列)到底是数值还是文本时，是通过该列的前8条记录是否有文本数据，有则该列作为文本读取，
        '    否则，即使后面的记录有文本，也还是按数值读取（导致文本为空）。
        '修改“HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Excel”下的TypeGuessRows注册表值，
        '    可以将其改大些，如1000,这样在1000条记录前只要有文本，就不会出现文本为空的现象，具体情况要视你的数据而定。
        strSourceRS = strSourceRS & " from [Excel 8.0;HDR=yes;IMEX=1;database=" & strFileName & "].[" & m_strServSheet & "]"
    Else  'CSV文件
        Dim fsObj As FileSystemObject
        Set fsObj = CreateObject("Scripting.FileSystemObject")
        Dim strFileNameDir As String
        Dim strFileNameBase As String
        
        strFileNameDir = fsObj.GetParentFolderName(strFileName)
        strFileNameBase = fsObj.GetFileName(strFileName)
        Set fsObj = Nothing

        strSourceRS = strSourceRS & " from [Text;database=" & strFileNameDir & "].[" & strFileNameBase & "]"
    End If

    Me.MousePointer = vbHourglass
    Call gImportData(strSourceRS, strTargetRS, iFieldCount, strRetMsg)
    Me.MousePointer = vbDefault
    Call RefreshGrid

    If strRetMsg <> "" Then
        Load frmMsgLog
        frmMsgLog.Caption = "导入信息说明"
        frmMsgLog.RTDesc.Text = "导入信息说明：" & vbCrLf & "导入文件名：" & strFileName & vbCrLf & strRetMsg
        frmMsgLog.Show vbModal
    End If
    Exit Sub
    
ErrHandle:
    If Err.Number = 32755 Then
    '按了取消
        dgFile.FileName = ""
        Exit Sub
    End If
    DisplayMsg "导入时出错!", vbCritical
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdSend_Click()
    If adoSalary.RecordSource = "" Then
        DisplayMsg "数据源为空，请选择要导入的数据", vbExclamation
        cbServType.SetFocus
        Exit Sub
    ElseIf adoSalary.Recordset.State = adStateClosed Then
        DisplayMsg "数据源为空，请选择要导入的数据", vbExclamation
        cbServType.SetFocus
        Exit Sub
    ElseIf adoSalary.Recordset.RecordCount = 0 Then
        DisplayMsg "记录数为０，请选择要导入的数据", vbExclamation
        dcbService.SetFocus
        Exit Sub
    End If
    If m_strServId = "" Then
        DisplayMsg "请选择模板", vbExclamation
        dcbService.SetFocus
        Exit Sub
    End If

    Load frmSendMail
    With frmSendMail
        .txtEmailSubject.Text = adoService.Recordset("serv_subject") & "（" & m_strServPeriod & "）"
        .txtServId.Text = m_strServId
        .adoSalary.RecordSource = Me!adoSalary.RecordSource
        .adoSalary.Refresh
        .adoSalary.Recordset.Move adoSalary.Recordset.AbsolutePosition - 1, 1
    End With
    frmSendMail.Show vbModal
End Sub

Private Sub dcbService_Change()
    If dcbService.BoundText = m_strServId Then
        Exit Sub
    End If
    m_strServId = dcbService.BoundText

    If m_strServId <> "" And adoService.Recordset.RecordCount > 0 Then
        adoService.Recordset.Move dcbService.SelectedItem - 1, 1
        m_strServPeriodMethod = UCase(adoService.Recordset("serv_period"))
        m_strServSheet = adoService.Recordset("serv_sheet")
    Else
        m_strServPeriodMethod = ""
        m_strServSheet = ""
    End If
    Call RefreshServPeriod
    If dtServPeriod.Visible = True Then
        Call dtServPeriod_Validate(False)
    End If

    Call RefreshGrid
End Sub

Private Sub dtServPeriod_Validate(Cancel As Boolean)
    If m_strServPeriod = dtServPeriod.value Then
        Exit Sub
    End If
    
    Select Case m_strServPeriodMethod
    Case "YEAR"
        m_strServPeriod = Format(dtServPeriod.value, DATE_FORMAT_YEAR)
    Case "DAY"
        m_strServPeriod = Format(dtServPeriod.value, DATE_FORMAT_DAY)
    Case Else
        m_strServPeriod = Format(dtServPeriod.value, DATE_FORMAT_MONTH)
    End Select

    Call RefreshGrid
End Sub

Private Sub Form_Load()
    '默认为数据导入
    '初始化变量
    m_strTableName = "salary"
    m_strServId = ""
    m_strServPeriodMethod = ""
    m_strServPeriod = ""
    m_strServType = ""
    cmdDel.Visible = False
    lbPeriodFormat.Caption = "（" & UCase(DATE_FORMAT_MONTH) & "）"
    dtServPeriod.CustomFormat = DATE_FORMAT_MONTH
    dtServPeriod.value = Date
    cbServPeriod.Visible = False
    
    tbSalary.MarqueeStyle = dbgHighlightRow

    adoSalary.ConnectionString = gStrConnDB
    adoSalary.CommandType = adCmdText

    adoService.ConnectionString = gStrConnDB
    adoService.CommandType = adCmdText
    adoService.RecordSource = "select * from services where valid_flag=1 order by modify_dt desc"
    adoService.Refresh
    dcbService.ListField = "serv_name"
    dcbService.BoundColumn = "serv_id"

    Dim rsServType As New ADODB.Recordset
    rsServType.Open "select dict_key from dicts where dict_sect='OPT_SERV_TYPE' order by dict_flag", _
                gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText

    cbServType.AddItem "全部"
    While rsServType.EOF <> True
        cbServType.AddItem rsServType(0)
        rsServType.MoveNext
    Wend
    rsServType.Close
    Set rsServType = Nothing
    cbServType.ListIndex = 0
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  '当窗体调整时会调整网格
  tbSalary.Left = 0
  tbSalary.Width = Me.ScaleWidth
  tbSalary.Height = Me.ScaleHeight - adoSalary.Height - picTop.Height
  tbSalary.Top = picTop.Height
End Sub

Private Sub RefreshGrid()
    Dim blDisplay As Boolean

    blDisplay = True
    If m_strServPeriod <> "" And m_strServId <> "" Then
        adoSalary.RecordSource = "SELECT emp.emp_email,salary.* FROM emp RIGHT JOIN salary ON emp.emp_id = salary.emp_id" & _
                                " WHERE salary.serv_period='" & m_strServPeriod & "'" & _
                                " AND salary.serv_id='" & m_strServId & "'" & _
                                " ORDER BY salary.emp_id"
        adoSalary.Refresh
    Else
        If adoSalary.RecordSource = "" Then
        ElseIf adoSalary.Recordset.State = adStateOpen Then
            adoSalary.Recordset.Close
        End If
        adoSalary.Caption = "无数据"
        blDisplay = False
    End If

    If blDisplay = True Then
        DisplayGrid tbSalary, m_strTableName, m_strServId
    End If
End Sub

Private Sub RefreshComboServPeriod()
    If cbServPeriod.Visible = False Then
        Exit Sub
    End If

    If m_strServId = "" Then
        cbServPeriod.Clear
        Exit Sub
    End If
    
    '读取已有信息
    Dim rsServDate As New ADODB.Recordset
    Dim strSQL As String
    strSQL = "select distinct serv_period from salary where serv_id='" & m_strServId & "' order by serv_period desc"

    rsServDate.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    cbServPeriod.Clear
    While rsServDate.EOF <> True
        cbServPeriod.AddItem rsServDate(0)
        rsServDate.MoveNext
    Wend
    rsServDate.Close
    Set rsServDate = Nothing
End Sub

Private Sub RefreshServPeriod()
    m_strServPeriod = ""

    Select Case m_strServPeriodMethod
    Case "YEAR"
        lbPeriod.Caption = "年份："
        lbPeriodFormat.Caption = "（" & UCase(DATE_FORMAT_YEAR) & "）"
        dtServPeriod.CustomFormat = DATE_FORMAT_YEAR
    Case "DAY"
        lbPeriod.Caption = "日期："
        lbPeriodFormat.Caption = "（" & UCase(DATE_FORMAT_DAY) & "）"
        dtServPeriod.CustomFormat = DATE_FORMAT_DAY
    Case Else
        lbPeriod.Caption = "月份："
        lbPeriodFormat.Caption = "（" & UCase(DATE_FORMAT_MONTH) & "）"
        dtServPeriod.CustomFormat = DATE_FORMAT_MONTH
    End Select

    Call RefreshComboServPeriod
End Sub

