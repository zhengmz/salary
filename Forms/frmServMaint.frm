VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmServMaint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "服务配置维护界面"
   ClientHeight    =   5745
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   7080
   Icon            =   "frmServMaint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adoService 
      Height          =   330
      Left            =   840
      Top             =   5040
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
      Caption         =   "Adodc1"
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
   Begin VB.TextBox txtCmdType 
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   9
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   8
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4380
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   4380
      ScaleWidth      =   6525
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   6525
      Begin MSDataGridLib.DataGrid grdServConf 
         Bindings        =   "frmServMaint.frx":000C
         Height          =   4215
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   7435
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4380
      Index           =   0
      Left            =   210
      ScaleHeight     =   4380
      ScaleWidth      =   6525
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   6525
      Begin VB.Frame fraBase2 
         Height          =   4095
         Left            =   240
         TabIndex        =   11
         Top             =   120
         Width           =   6015
         Begin VB.TextBox txtServTemplFn 
            DataField       =   "serv_templ_fn"
            DataSource      =   "adoService"
            Height          =   285
            Left            =   1320
            MaxLength       =   200
            TabIndex        =   29
            Top             =   960
            Width           =   4335
         End
         Begin VB.TextBox txtServId 
            BackColor       =   &H80000011&
            DataField       =   "serv_id"
            DataSource      =   "adoService"
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   25
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtServType 
            BackColor       =   &H80000011&
            DataField       =   "serv_type"
            DataSource      =   "adoService"
            Enabled         =   0   'False
            Height          =   285
            Left            =   3720
            TabIndex        =   24
            Top             =   240
            Width           =   1935
         End
         Begin VB.ComboBox cbServName 
            Height          =   315
            Left            =   1320
            TabIndex        =   23
            Top             =   1365
            Visible         =   0   'False
            Width           =   4335
         End
         Begin VB.TextBox txtServSubject 
            DataField       =   "serv_subject"
            DataSource      =   "adoService"
            Height          =   285
            Left            =   1320
            MaxLength       =   200
            TabIndex        =   18
            Top             =   2640
            Width           =   4335
         End
         Begin VB.TextBox txtServDesc 
            DataField       =   "serv_desc"
            DataSource      =   "adoService"
            Height          =   855
            Left            =   1320
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   3120
            Width           =   4335
         End
         Begin VB.OptionButton optServPeriod 
            Caption         =   "每年"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   16
            Top             =   2250
            Width           =   855
         End
         Begin VB.OptionButton optServPeriod 
            Caption         =   "每月"
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   15
            Top             =   2250
            Width           =   735
         End
         Begin VB.OptionButton optServPeriod 
            Caption         =   "不定期"
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   14
            Top             =   2250
            Width           =   855
         End
         Begin VB.TextBox txtServSheet 
            DataField       =   "serv_sheet"
            DataSource      =   "adoService"
            Height          =   285
            Left            =   1320
            TabIndex        =   13
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txtServName 
            DataField       =   "serv_name"
            DataSource      =   "adoService"
            Height          =   285
            Left            =   1320
            TabIndex        =   12
            Top             =   1380
            Width           =   4335
         End
         Begin MSDataListLib.DataCombo dcbServType 
            Bindings        =   "frmServMaint.frx":0026
            Height          =   315
            Left            =   3720
            TabIndex        =   28
            Top             =   225
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "名称："
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   32
            Top             =   1425
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Sheet名称："
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   31
            Top             =   1845
            Width           =   960
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "模板文件："
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   1005
            Width           =   900
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   5760
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "标识ID："
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   285
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "类型："
            Height          =   195
            Left            =   3000
            TabIndex        =   26
            Top             =   285
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "描述："
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   3150
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "周期："
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "邮件主题："
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   19
            Top             =   2685
            Width           =   900
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   5175
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   5175
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   3330
      TabIndex        =   1
      Top             =   5175
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4845
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8546
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "基本信息"
            Key             =   "BaseGrp"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "字段映射关系"
            Key             =   "ExtendGrp"
            ImageVarType    =   2
         EndProperty
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
   End
   Begin MSAdodcLib.Adodc adoServType 
      Height          =   330
      Left            =   2040
      Top             =   5040
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
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc adoServConf 
      Height          =   330
      Left            =   960
      Top             =   5400
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
      Caption         =   "无数据"
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
Attribute VB_Name = "frmServMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_strServId As String
Private m_blUpdateSeq As Boolean
Private m_blUpdateSucc As Boolean

Private Sub adoService_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    DisplayMsg "数据库错误", vbCritical
End Sub

Private Sub adoService_WillChangeField(ByVal cFields As Long, Fields As Variant, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If cFields > 0 Then
        cmdApply.Enabled = True
    End If
End Sub

Private Sub cbServName_Validate(Cancel As Boolean)
    If cbServName.ListIndex = -1 And cbServName.Text = "" And txtServName.Text = "" Then
        Exit Sub
    End If
    If cbServName.Text = txtServName.Text Then
        Exit Sub
    End If

    '查重
    Dim i As Integer
    For i = 0 To cbServName.ListCount - 1
        If cbServName.Text = cbServName.List(i) Then
            DisplayMsg "名字重复，请重新输入。", vbExclamation
            Cancel = True
            Exit Sub
        End If
    Next i
    txtServName.Text = cbServName.Text
End Sub

Private Sub cmdApply_Click()
    On Error GoTo ErrHandle

    m_blUpdateSucc = False
    If txtServId.Text = "" Then
        DisplayMsg "无法保存，标识ID为空。", vbExclamation
        Exit Sub
    End If
    If txtServType.Text = "" Then
        DisplayMsg "无法保存，类型为空。", vbExclamation
        Exit Sub
    End If
    If txtServName.Text = "" Then
        DisplayMsg "无法保存，名称为空。", vbExclamation
        Exit Sub
    End If
    
    If cmdApply.Visible = True Then
        adoService.Recordset("modify_dt") = Date
        adoService.Recordset.Update
        If m_blUpdateSeq = True Then
            adoServType.Recordset("seq") = Val(adoServType.Recordset("seq")) + 1
            adoServType.Recordset.Update
            m_blUpdateSeq = False
        End If
        m_blUpdateSucc = True
    End If
    cmdApply.Enabled = False
    Exit Sub
    
ErrHandle:
    DisplayMsg "更新错误", vbCritical
    adoService.Recordset.Cancel
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    adoServConf.Recordset.Cancel
    If cmdApply.Enabled = True Then
        adoService.Recordset.Cancel
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If cmdApply.Enabled = True Then
        Call cmdApply_Click
    End If
    If m_blUpdateSucc = True Or cmdApply.Visible = False Then
        Unload Me
    End If
End Sub

Private Sub dcbServType_Change()
    adoServType.Recordset.Move dcbServType.SelectedItem - 1, 1
    txtServId.Text = adoServType.Recordset("code") & Format(Val(adoServType.Recordset("seq")) + 1, SERV_ID_PATTERN)
    txtServType.Text = dcbServType.Text
    cmdApply.Enabled = True
    m_blUpdateSeq = True

    '读取已有信息
    Dim rsService As New ADODB.Recordset
    Dim strSQL As String
    strSQL = "select serv_name from services where valid_flag=1 and serv_type='" & dcbServType.Text & "' order by modify_dt desc"

    rsService.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    cbServName.Clear
    While rsService.EOF <> True
        cbServName.AddItem rsService(0)
        rsService.MoveNext
    Wend
    rsService.Close
    Set rsService = Nothing
    txtServName.Text = ""
End Sub

Private Sub Form_Activate()
    Select Case txtCmdType.Text
    Case "0"    '查看
        grdServConf.AllowAddNew = False
        grdServConf.AllowDelete = False
        grdServConf.AllowUpdate = False
        grdServConf.MarqueeStyle = dbgHighlightRow
        cmdApply.Enabled = False
        cmdOK.Enabled = False
        cmdCancel.Caption = "关闭"
        txtServName.Enabled = False
        txtServName.BackColor = &H80000011
        txtServSheet.Enabled = False
        txtServSheet.BackColor = &H80000011
        txtServTemplFn.Enabled = False
        txtServTemplFn.BackColor = &H80000011
        txtServSubject.Enabled = False
        txtServSubject.BackColor = &H80000011
        txtServDesc.Enabled = False
        txtServDesc.BackColor = &H80000011
        optServPeriod.Item(0).Enabled = False
        optServPeriod.Item(1).Enabled = False
        optServPeriod.Item(2).Enabled = False
    Case "1"    '新增
        adoService.Recordset.AddNew
        dcbServType.Visible = True
        txtServType.Visible = False
        cbServName.Visible = True
        txtServName.Visible = False
    Case "2"    '修改
    Case Else
        DisplayMsg "无效指令，或是非法进入本页面", vbExclamation
        Unload Me
        Exit Sub
    End Select

    Select Case UCase(adoService.Recordset("serv_period"))
    Case "YEAR"
        optServPeriod.Item(0).value = True
    Case "MONTH"
        optServPeriod.Item(1).value = True
    Case "DAY"
        optServPeriod.Item(2).value = True
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
    '初始化变量
    m_strServId = ""
    m_blUpdateSeq = False
    m_blUpdateSucc = True   '默认未做修改，可直接退出

    adoServConf.ConnectionString = gStrConnDB
    adoServConf.CommandType = adCmdText
    adoService.ConnectionString = gStrConnDB
    adoService.CommandType = adCmdText
    grdServConf.AllowRowSizing = False
    
    adoServType.ConnectionString = gStrConnDB
    adoServType.CommandType = adCmdText
    adoServType.RecordSource = "select dict_key as serv_type, dict_type as code, dict_value as seq" & _
                            " from dicts where dict_sect='OPT_SERV_TYPE' order by dict_flag"
    adoServType.Refresh
    dcbServType.ListField = "serv_type"

    cmdApply.Enabled = False
    dcbServType.Visible = False
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub grdServConf_Click()
    If adoServConf.Recordset.RecordCount = 0 Then
        adoServConf.Recordset.AddNew
        adoServConf.Recordset("serv_id") = m_strServId
        adoServConf.Recordset("valid_flag") = 1
    End If
End Sub

Private Sub grdServConf_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
    adoServConf.Recordset.CancelBatch adAffectCurrent
    adoServConf.Recordset.Cancel
End Sub

Private Sub grdServConf_OnAddNew()
    adoServConf.Recordset.AddNew
    adoServConf.Recordset("serv_id") = m_strServId
    adoServConf.Recordset("valid_flag") = 1
End Sub

Private Sub optServPeriod_Click(Index As Integer)
    Dim strServPeriod As String
    
    strServPeriod = Choose(Index + 1, "YEAR", "MONTH", "DAY")
    If strServPeriod = UCase(adoService.Recordset("serv_period")) Then
        Exit Sub
    End If
    
    adoService.Recordset("serv_period") = strServPeriod
    'cmdApply.Enabled = True
End Sub

Private Sub tbsOptions_Click()
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    
    Dim iCurrentTab As Integer
    iCurrentTab = tbsOptions.SelectedItem.Index
    '业务处理或判断
    Select Case iCurrentTab
    Case 2
        If cmdApply.Enabled = True Then
            Call cmdApply_Click
            If m_blUpdateSucc = False Then
                Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
            End If
        End If
        
        Call RefreshGrid
    End Select
End Sub

Private Sub RefreshGrid()
    If m_strServId = txtServId.Text Then
        Exit Sub
    End If
    m_strServId = txtServId.Text
    
    If m_strServId = "" Then
        If adoServConf.RecordSource = "" Then
            Exit Sub
        ElseIf adoServConf.Recordset.State = adStateOpen Then
            adoServConf.Recordset.Close
        End If
    Else
        adoServConf.RecordSource = "Select * from serv_field where serv_id='" & m_strServId & "' order by field_name"
        adoServConf.Refresh
        Call DisplayGrid(grdServConf, "serv_field")
        grdServConf.Columns(0).Locked = True
        grdServConf.Columns(3).Locked = True
    End If
End Sub
