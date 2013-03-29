VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmService 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "服务配置"
   ClientHeight    =   5160
   ClientLeft      =   3675
   ClientTop       =   3240
   ClientWidth     =   7995
   Icon            =   "frmService.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraService 
      Caption         =   "服务配置表"
      Height          =   4815
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   5895
      Begin MSDataGridLib.DataGrid grdService 
         Bindings        =   "frmService.frx":000C
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
         Caption         =   "设为默认"
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
   Begin MSAdodcLib.Adodc adoService 
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
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub adoService_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    DisplayMsg "数据库错误", vbCritical
End Sub

Private Sub adoService_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    '按纽显示控制
    Dim i As Integer
    If adoService.Recordset.RecordCount = 0 Then
        For i = 0 To cmdOper.Count - 1
            cmdOper.Item(i).Enabled = False
        Next
        cmdOper.Item(1).Enabled = True
    Else
        For i = 0 To cmdOper.Count - 1
            cmdOper.Item(i).Enabled = True
        Next
        If adoService.Recordset.EOF = True Then
            Exit Sub
        End If
        If adoService.Recordset("default_flag") = 1 Then
            cmdOper.Item(4).Enabled = False
        Else
            cmdOper.Item(4).Enabled = True
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOper_Click(Index As Integer)
    If Index = OPER_ADD Then           '增加
        Dim strMsg As String
        strMsg = "请选择增加方式。" & vbCrLf & vbCrLf & _
                 "[是]向导方式（推荐方式）" & vbCrLf & _
                 "[否]手工方式"
        If DisplayMsg(strMsg, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            frmServWizard.Show vbModal
        Else
            Call ShowMaintWin(Index)
        End If
        Call RefreshData
        Exit Sub
    End If
    If adoService.Recordset.RecordCount = 0 Then
        DisplayMsg "没有选择服务，请确认是否有数据", vbExclamation
        cmdOper(Index).Enabled = False
        Exit Sub
    End If
    Select Case Index
    Case OPER_QUERY  '查看
        Call ShowMaintWin(Index)
    Case OPER_ADD  '增加
    Case OPER_MODIFY  '修改
        DisplayMsg "推荐使用配置向导对模板修改。", vbInformation
        Call ShowMaintWin(Index)
        Call RefreshData
    Case OPER_DEL  '删除
        Call DelService
    Case 4  '设为默认
        Dim strServId As String
        Dim i As Integer
        strServId = adoService.Recordset("serv_id")
        i = adoService.Recordset.AbsolutePosition
        gAdoConnDB.BeginTrans
        gAdoConnDB.Execute "update services set default_flag=0, modify_dt=#" & Format(Date, "yyyy-MM-dd") & "# where default_flag=1"
        gAdoConnDB.Execute "update services set default_flag=1, modify_dt=#" & Format(Date, "yyyy-MM-dd") & "# where serv_id='" & strServId & "'"
        gAdoConnDB.CommitTrans
        Call RefreshData
        adoService.Recordset.Move i - 1, 1
    End Select
End Sub

Private Sub Form_Load()
    grdService.MarqueeStyle = dbgHighlightRow
    grdService.AllowRowSizing = False
    
    adoService.ConnectionString = gStrConnDB
    adoService.CommandType = adCmdText
    adoService.RecordSource = "select * from services where valid_flag=1 order by serv_type, modify_dt desc"

    Call RefreshData
End Sub

Private Sub grdService_DblClick()
    Call cmdOper_Click(0)
End Sub

Private Sub RefreshData()
    adoService.Refresh
    Call DisplayGrid(grdService, "services")
    On Error Resume Next
    grdService.SetFocus
End Sub

Private Sub ShowMaintWin(ByVal Index As Integer)
    Load frmServMaint
    With frmServMaint
        .txtCmdType.Text = Index
        If adoService.Recordset.RecordCount = 0 Then
            .adoService.RecordSource = "select * from services where valid_flag=1"
        Else
            .adoService.RecordSource = "select * from services where serv_id='" & adoService.Recordset("serv_id") & "'"
        End If
        .adoService.Refresh
    End With
    frmServMaint.Show vbModal
End Sub

Private Sub DelService()
    Dim rsSalary As New ADODB.Recordset
    Dim strSql As String
    Dim iSalaryCount As Integer
    
    strSql = "select count(*) from salary where serv_id='" & adoService.Recordset("serv_id") & "'"
    rsSalary.Open strSql, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    
    iSalaryCount = rsSalary(0)
    rsSalary.Close
    Set rsSalary = Nothing
    If iSalaryCount > 0 Then
        If DisplayMsg("还有与此配置相关的历史数据，请确认是否删除？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    Else
        If DisplayMsg("请确认删除当前记录？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    End If

    Dim i As Integer
    i = adoService.Recordset.AbsolutePosition
    gExecSql ("update services set valid_flag=0, modify_dt=#" & Format(Date, "yyyy-MM-dd") & "# where serv_id='" & adoService.Recordset("serv_id") & "'")
    Call RefreshData
    If adoService.Recordset.RecordCount = 0 Then
        Exit Sub
    End If
    If i > adoService.Recordset.RecordCount Then
        adoService.Recordset.MoveLast
    Else
        adoService.Recordset.Move i - 1, 1
    End If
End Sub
