VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSendMail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "邮件发送"
   ClientHeight    =   4965
   ClientLeft      =   300
   ClientTop       =   1710
   ClientWidth     =   7230
   Icon            =   "frmSendMail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFrequency 
      Height          =   285
      Left            =   1680
      TabIndex        =   20
      Text            =   "0"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "预 览(&P)"
      Height          =   420
      Left            =   2040
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtServId 
      Height          =   285
      Left            =   360
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc adoSalary 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   4635
      Width           =   7230
      _ExtentX        =   12753
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
   Begin VB.Frame fraBase 
      Caption         =   "基本信息"
      Height          =   1215
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtEmailServer 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   690
         Width           =   2655
      End
      Begin VB.TextBox txtEmailServerPort 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   14
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox txtEmailFromName 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   11
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox txtEmailFrom 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   330
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "发送服务器："
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   735
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "发送端口："
         Height          =   195
         Left            =   4320
         TabIndex        =   16
         Top             =   735
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "发送人："
         Height          =   195
         Left            =   4320
         TabIndex        =   13
         Top             =   375
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "发送人地址："
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   375
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "发 送(&S)"
      Height          =   420
      Left            =   3720
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关 闭(&C)"
      Height          =   420
      Left            =   5400
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtEmailRemark 
      Height          =   1215
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1800
      Width           =   5055
   End
   Begin VB.TextBox txtEmailSubject 
      Height          =   270
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   5055
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   480
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "（每次发送的记录数，默认全部，大于零有效）"
      Height          =   195
      Left            =   3000
      TabIndex        =   21
      Top             =   3165
      Width           =   3780
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "发送频率："
      Height          =   195
      Left            =   480
      TabIndex        =   19
      Top             =   3165
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "状态："
      Height          =   180
      Left            =   480
      TabIndex        =   8
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label lbStatus 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "准备发送......"
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   3480
      Width           =   990
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "邮件备注："
      Height          =   180
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "邮件标题："
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   900
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub adoSalary_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If adoSalary.Recordset.RecordCount = 0 Then
        adoSalary.Caption = "无记录"
    Else
        If adoSalary.Recordset.EOF = False Then
            adoSalary.Caption = "当前记录：[" & adoSalary.Recordset("emp_name") & _
                            "]，当前记录位置: " & adoSalary.Recordset.AbsolutePosition & _
                            "/" & adoSalary.Recordset.RecordCount
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    Call SendMail(True)
End Sub

Private Sub cmdSend_Click()
    If DisplayMsg("确认数据已经检查过，需要正式发送邮件？", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
        Exit Sub
    End If

    If SendMail(False) = True Then
        DisplayMsg "发送完毕", vbInformation
    End If
End Sub

Public Sub Preview()
    Call SendMail(True)
End Sub

Private Sub Form_Load()
    Dim m_clsRegConfig As New CRegConfig
    
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    txtEmailFromName.Text = m_clsRegConfig.GetConfig("Options", "EmailFromName")
    txtEmailFrom.Text = m_clsRegConfig.GetConfig("Options", "EmailFromAddr")
    txtEmailServer.Text = m_clsRegConfig.GetConfig("Options", "EmailServerIP")
    txtEmailServerPort.Text = m_clsRegConfig.GetConfig("Options", "EmailServerPort")
    
    adoSalary.ConnectionString = gStrConnDB
    adoSalary.CommandType = adCmdText
End Sub

Private Function SendMail(ByVal pmBlPrevFlag As Boolean) As Boolean
    SendMail = False

    If txtEmailFromName.Text = "" Or txtEmailFrom.Text = "" Or txtEmailServer.Text = "" Or txtEmailServerPort.Text = "" Then
        DisplayMsg "请配置邮件服务器地址或发送人邮件地址，无法发送！", vbExclamation
        cmdSend.Enabled = False
        Exit Function
    End If
    If txtServId.Text = "" Then
        DisplayMsg "没有选择合适的服务配置，无法发送！", vbExclamation
        cmdSend.Enabled = False
        Exit Function
    End If
    If adoSalary.Recordset.RecordCount = 0 Then
        DisplayMsg "所要发送的记录为0，请检查相关数据，无法发送！", vbExclamation
        cmdSend.Enabled = False
        Exit Function
    End If

    On Error GoTo Err_Handle
    
    '生成模版
    Dim strSQL As String
    Dim iArrLen As Integer
    Dim rsExpFormat As New ADODB.Recordset

    strSQL = "Select field_name,display_name From serv_field" & _
                " Where serv_id='" & txtServId.Text & "' AND valid_flag = 1 order by field_name"

    rsExpFormat.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    iArrLen = rsExpFormat.RecordCount
    If iArrLen = 0 Then
        DisplayMsg "没有找到相应的字段映射关系，无法生成发送邮件内容！", vbCritical
        rsExpFormat.Close
        Set rsExpFormat = Nothing
        cmdSend.Enabled = False
        Exit Function
    End If
    
    '创建显示的HTML模板
    Dim strBodyTemplate As String
    Dim i As Integer
    
    If Not CreateEmailTemplate(rsExpFormat, txtServId.Text, strBodyTemplate) Then
        Exit Function
    End If

    '创建字段映射关系表
    Dim strArrField() As String
    ReDim strArrField(iArrLen)

    rsExpFormat.MoveFirst
    For i = 0 To iArrLen - 1
        strArrField(i) = rsExpFormat("field_name")
        rsExpFormat.MoveNext
    Next
    rsExpFormat.Close
    Set rsExpFormat = Nothing

    '发送邮件
    Dim strToEmail As String
    Dim strBody As String
    Dim strFieldName As String
    Dim strErrMsg As String
    Dim strRetMsg As String
    
    Dim clsMail As CMail
    Set clsMail = New CMail
    clsMail.BindObj Winsock1
    
    '连接mail server
    lbStatus.Caption = "正在连接Email服务器中......"
    If pmBlPrevFlag = False Then
        If Not clsMail.Connect(txtEmailServer.Text, txtEmailServerPort.Text) Then
            DisplayMsg "连接服务器错误", vbCritical
            Exit Function
        End If
        If Not clsMail.Init(txtEmailFrom.Text, txtEmailFromName.Text, txtEmailSubject.Text) Then
            DisplayMsg "初始化服务器错误", vbCritical
            Exit Function
        End If
        'clsMail.Init txtEmailFrom.Text, txtEmailFromName.Text, txtEmailSubject.Text
        lbStatus.Caption = "已连上Email服务器"
        
        WriteLog "frmSendMail:SendMail", "服务器连接后的状态是" & clsMail.GetState, LOG_LEVEL_DEBUG
        adoSalary.Recordset.MoveFirst
    Else
        lbStatus.Caption = "预览测试中"
    End If

    '获取发送频率，即每次发送记录数
    Dim intFrequency As Integer
    Dim intRecordInd As Integer

    intFrequency = Val(txtFrequency.Text)
    intRecordInd = 0

    lbStatus.Caption = "发送中......"
    strErrMsg = ""
    Do Until adoSalary.Recordset.EOF
        strBody = strBodyTemplate
        For i = 0 To iArrLen - 1
            strFieldName = strArrField(i)
            If IsNull(adoSalary.Recordset(strFieldName)) Or adoSalary.Recordset(strFieldName) = "" Then
                strBody = Replace(strBody, "{" & strFieldName & "}", "&nbsp; ")
            Else
                strBody = Replace(strBody, "{" & strFieldName & "}", adoSalary.Recordset(strFieldName))
            End If
        Next

        If pmBlPrevFlag = True Then
            Load frmBrowser
            With frmBrowser
                .brwWebBrowser.Document.Open
                .brwWebBrowser.Document.writeln strBody
            End With
            Me.Hide
            Me.Show vbModeless
            frmBrowser.Show vbModal
            Me.Hide
            Me.Show vbModal
            Exit Do
        End If

        If IsNull(adoSalary.Recordset("emp_email")) Or adoSalary.Recordset("emp_email") = "" Then
            strErrMsg = strErrMsg & vbCrLf & _
                        "姓名：" & adoSalary.Recordset("emp_name") & "，邮箱：空" & vbCrLf & _
                        "错误原因：没有相应的邮箱" & vbCrLf
        Else
            strToEmail = adoSalary.Recordset("emp_email")
            '过滤从Domino服务器导出的特殊字符
            strToEmail = Replace(strToEmail, Chr(9), "")
            'printChr (strToEmail)
            If Not clsMail.SendMail(strToEmail, strBody, strRetMsg) Then
                strErrMsg = strErrMsg & vbCrLf & _
                        "姓名：" & adoSalary.Recordset("emp_name") & "，邮箱：" & "[" & strToEmail & "]" & vbCrLf & _
                        "错误原因：" & strRetMsg
                If clsMail.GetState = sckError Then
'                    strErrMsg = strErrMsg & vbCrLf & _
'                        "socket连接出错，无法正常发送。"
'                    Exit Do
                    DisplayMsg "socket连接出错，无法正常发送!", vbCritical
                    Exit Function
                End If
            Else
                'WriteLog "frmSendMail:SendMail", "发送[" & strToEmail & "]进入队列成功。" & strRetMsg, LOG_LEVEL_DEBUG
                adoSalary.Recordset.Update "send_dt", Date
            End If
        End If
        adoSalary.Recordset.MoveNext
        
        If intFrequency > 0 Then
            intRecordInd = intRecordInd + 1
            WriteLog "frmSendMail:SendMail", "目前处理的记录数: " & intRecordInd, LOG_LEVEL_DEBUG
            
            If (intRecordInd Mod intFrequency) = 0 Then
                '断开并重新连接邮箱服务器，以便其及时发出邮件
                'WriteLog "frmSendMail:SendMail", "重新连接服务器前的连接状态是" & clsMail.GetState, LOG_LEVEL_DEBUG
                If Not clsMail.Disconnect() Then
                    'WriteLog "frmSendMail:SendMail", "重新连接服务器时，正在断开服务器后的连接状态是" & clsMail.GetState, LOG_LEVEL_DEBUG
                    DisplayMsg "断开服务器失败!" & vbCrLf & vbCrLf & "已发送批次为" & (intRecordInd \ intFrequency), vbCritical
                    Exit Function
                End If
                'WriteLog "frmSendMail:SendMail", "重新连接服务器时，断开服务器后的连接状态是" & clsMail.GetState, LOG_LEVEL_DEBUG
                If Not clsMail.Connect(txtEmailServer.Text, txtEmailServerPort.Text) Then
                    DisplayMsg "连接服务器错误" & vbCrLf & vbCrLf & "已发送批次为" & (intRecordInd \ intFrequency), vbCritical
                    Exit Function
                End If
                If Not clsMail.Init(txtEmailFrom.Text, txtEmailFromName.Text, txtEmailSubject.Text) Then
                    DisplayMsg "初始化服务器错误", vbCritical
                    Exit Function
                End If
                'clsMail.Init txtEmailFrom.Text, txtEmailFromName.Text, txtEmailSubject.Text
                'WriteLog "frmSendMail:SendMail", "重新连接服务器时，连接服务器后的连接状态是" & clsMail.GetState, LOG_LEVEL_DEBUG
                'If DisplayMsg("是否继续？", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                '    Exit Do
                'End If
            End If
        End If
    Loop
    '断开与mail server的连接
    If pmBlPrevFlag = False Then
        'WriteLog "frmSendMail:SendMail", "发送完成后，断开服务器前的连接状态是" & clsMail.GetState, LOG_LEVEL_DEBUG
        If clsMail.GetState <> sckClosed And clsMail.GetState <> sckError Then
            If Not clsMail.Disconnect() Then
                DisplayMsg "断开服务器失败，无法正常发送!", vbCritical
                Exit Function
            End If
        End If
        'WriteLog "frmSendMail:SendMail", "发送完成后，断开服务器后的连接状态是" & clsMail.GetState, LOG_LEVEL_DEBUG
        clsMail.ReleaseObj
        lbStatus.Caption = "发送完毕"

        If strErrMsg <> "" Then
            Me.Hide
            Load frmMsgLog
            frmMsgLog.Caption = "邮件发送日志"
            frmMsgLog.RTDesc.Text = "以下的邮件无法送达：" & vbCrLf & strErrMsg
            frmMsgLog.Show vbModal
            Me.Show vbModal
        End If
    Else
        lbStatus.Caption = "预览完成"
    End If
    SendMail = True
    Exit Function

Err_Handle:
    DisplayMsg "发送错误！", vbCritical
End Function

Private Sub printChr(ByVal pmStr As String)
    Dim i As Integer
    For i = 1 To Len(pmStr)
        Debug.Print Mid(pmStr, i, 1) & "--" & Asc(Mid(pmStr, i, 1))
    Next
End Sub

Private Function CreateEmailTemplate(ByVal pmRsServField As ADODB.Recordset, ByVal pmStrServId As String, pmStrTemplate As String) As Boolean
    On Error GoTo ErrHandle
    CreateEmailTemplate = False
    pmStrTemplate = ""

    Dim fsObj As FileSystemObject
    Dim strTemplFileName As String
    Dim strBodyTemplate As String
    Dim i As Integer
    Dim iArrLen As Integer

    strTemplFileName = App.Path & "\Template\" & pmStrServId & ".html"
    strBodyTemplate = ""
    iArrLen = pmRsServField.RecordCount
    pmRsServField.MoveFirst

    Set fsObj = CreateObject("Scripting.FileSystemObject")
    If fsObj.FileExists(strTemplFileName) Then
        Dim ts As TextStream
        
        Set ts = fsObj.OpenTextFile(strTemplFileName, ForReading)
        strBodyTemplate = ts.ReadAll()
        'Call WriteLog("CreateEmailTemplate", strBodyTemplate, LOG_LEVEL_INFO)
        ts.Close
        Set ts = Nothing
        
        For i = 0 To iArrLen - 1
            strBodyTemplate = Replace(strBodyTemplate, "{" & pmRsServField("display_name") & "}", "{" & pmRsServField("field_name") & "}")
            pmRsServField.MoveNext
        Next
    Else
        '创建显示的HTML模板
        Dim strFirstLine As String
        Dim strSecondLine As String
        Dim iSplitField As Integer

        'HTML标题
        strBodyTemplate = "<html><body>" + vbCrLf
    
        '主要内容（TABLE格式）
        strBodyTemplate = strBodyTemplate + "<table border='1' cellpadding='2' cellspacing='0' bordercolor='#000000'>" + vbCrLf
        
        strFirstLine = "<tr bgcolor='#FFCC00'>" + vbCrLf
        strSecondLine = "<tr>" + vbCrLf
        iSplitField = 0
        For i = 0 To iArrLen - 1
            strFirstLine = strFirstLine + "<td align='center'><strong>" & pmRsServField("display_name") & "</strong></td>" + vbCrLf
            strSecondLine = strSecondLine + "<td align='center'>{" & pmRsServField("field_name") & "}</td>" + vbCrLf
            iSplitField = iSplitField + 1
            If gSysSplitFields > 0 Then
                If iSplitField Mod gSysSplitFields = 0 Then
                    strFirstLine = strFirstLine + "</tr>" + vbCrLf
                    strSecondLine = strSecondLine + "</tr>" + vbCrLf
                    strBodyTemplate = strBodyTemplate + strFirstLine + strSecondLine + "</table>" + vbCrLf
                    strBodyTemplate = strBodyTemplate + "<br>" + vbCrLf
                    'strBodyTemplate = strBodyTemplate + "<p>&nbsp; </p>"
                    strBodyTemplate = strBodyTemplate + "<table border='1' cellpadding='2' cellspacing='0' bordercolor='#000000'>" + vbCrLf
                    strFirstLine = "<tr bgcolor='#FFCC00'>" + vbCrLf
                    strSecondLine = "<tr>" + vbCrLf
                End If
            End If
            pmRsServField.MoveNext
        Next
    
        strFirstLine = strFirstLine + "</tr>" + vbCrLf
        strSecondLine = strSecondLine + "</tr>" + vbCrLf
        strBodyTemplate = strBodyTemplate + strFirstLine + strSecondLine + "</table>" + vbCrLf
        
'        '备注
'        Dim strArrRemark() As String
'        Dim iLine As Integer
'
'        strArrRemark = Split(txtEmailRemark.Text, vbCrLf)
'
'        For i = 0 To UBound(strArrRemark)
'            strBodyTemplate = strBodyTemplate + "<br>" + strArrRemark(i) + vbCrLf
'        Next
        'HTML结尾
        strBodyTemplate = strBodyTemplate + "</body></html>"
    End If
    
    '备注
    Dim strArrRemark() As String
    Dim iLine As Integer

    strArrRemark = Split(txtEmailRemark.Text, vbCrLf)

    For i = 0 To UBound(strArrRemark)
        strBodyTemplate = strBodyTemplate + "<br>" + strArrRemark(i) + vbCrLf
    Next
    
    Set fsObj = Nothing
    CreateEmailTemplate = True
    pmStrTemplate = strBodyTemplate
    Exit Function

ErrHandle:
    DisplayMsg "创建邮件模板出错!", vbCritical
End Function
