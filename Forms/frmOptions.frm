VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选项"
   ClientHeight    =   5535
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   7575
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adoDicts 
      Height          =   330
      Left            =   360
      Top             =   5040
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
      Caption         =   "adoDicts"
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
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   10
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   9
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
      ScaleWidth      =   7125
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   7125
      Begin MSDataGridLib.DataGrid grdDicts 
         Bindings        =   "frmOptions.frx":000C
         Height          =   4095
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7223
         _Version        =   393216
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
      ScaleWidth      =   7125
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   7125
      Begin VB.Frame fraEmail2 
         Caption         =   "发送服务器信息"
         Height          =   1185
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   5775
         Begin VB.TextBox txtEmailServerPort 
            Height          =   285
            Left            =   1920
            TabIndex        =   19
            Top             =   705
            Width           =   1095
         End
         Begin VB.TextBox txtEmailServer 
            Height          =   285
            Left            =   1920
            TabIndex        =   17
            Top             =   345
            Width           =   2295
         End
         Begin VB.Label Label4 
            Caption         =   "发送端口："
            Height          =   255
            Left            =   480
            TabIndex        =   18
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "发送服务器："
            Height          =   255
            Left            =   480
            TabIndex        =   16
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame fraEmail1 
         Caption         =   "发送人信息"
         Height          =   1185
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   5775
         Begin VB.TextBox txtEmailFrom 
            Height          =   285
            Left            =   1920
            TabIndex        =   14
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtEmailFromName 
            Height          =   285
            Left            =   1920
            TabIndex        =   11
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "发送人地址："
            Height          =   255
            Left            =   480
            TabIndex        =   13
            Top             =   735
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "发送人姓名："
            Height          =   255
            Left            =   480
            TabIndex        =   12
            Top             =   375
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   5055
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   5055
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   3690
      TabIndex        =   1
      Top             =   5055
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4845
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   8546
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "邮件发送"
            Key             =   "EmailGrp"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "数据字典"
            Key             =   "DictsGrp"
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
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsRegConfig As CRegConfig
Private m_strEmailFromName As String
Private m_strEmailFrom As String
Private m_strEmailServer As String
Private m_strEmailServerPort As String

Private Sub cmdApply_Click()
    m_clsRegConfig.SetConfig "Options", "EmailFromName", m_strEmailFromName
    m_clsRegConfig.SetConfig "Options", "EmailFromAddr", m_strEmailFrom
    m_clsRegConfig.SetConfig "Options", "EmailServerIP", m_strEmailServer
    m_clsRegConfig.SetConfig "Options", "EmailServerPort", m_strEmailServerPort
    cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If cmdApply.Enabled Then
        Call cmdApply_Click
    End If
    Unload Me
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
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    Set m_clsRegConfig = New CRegConfig
    m_strEmailFromName = m_clsRegConfig.GetConfig("Options", "EmailFromName")
    m_strEmailFrom = m_clsRegConfig.GetConfig("Options", "EmailFromAddr")
    m_strEmailServer = m_clsRegConfig.GetConfig("Options", "EmailServerIP")
    m_strEmailServerPort = m_clsRegConfig.GetConfig("Options", "EmailServerPort")
    
    txtEmailFromName.Text = m_strEmailFromName
    txtEmailFrom.Text = m_strEmailFrom
    txtEmailServer.Text = m_strEmailServer
    txtEmailServerPort.Text = m_strEmailServerPort
    cmdApply.Enabled = False
    
    adoDicts.ConnectionString = gStrConnDB
    adoDicts.CommandType = adCmdText
    adoDicts.RecordSource = "select * from dicts order by dict_sect, dict_flag"
    adoDicts.Refresh
    grdDicts.AllowRowSizing = False
    Call DisplayGrid(grdDicts, "dicts")
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
    
    Select Case iCurrentTab
    Case 2
        If cmdApply.Enabled = True Then
            Call cmdApply_Click
        End If
        DisplayMsg "如果你不明白这些配置，请不要修改，否则可能造成系统无法正常运行。" & _
            vbCrLf & "其中，系统设置[SYS]的修改，需重新启动应用才会有效。" & _
            vbCrLf & "如有问题，请及时联系管理员", vbExclamation
    End Select
End Sub

Private Sub txtEmailFrom_LostFocus()
    If m_strEmailFrom <> txtEmailFrom.Text Then
        m_strEmailFrom = txtEmailFrom.Text
        cmdApply.Enabled = True
    End If
End Sub

Private Sub txtEmailFromName_LostFocus()
    If m_strEmailFromName <> txtEmailFromName.Text Then
        m_strEmailFromName = txtEmailFromName.Text
        cmdApply.Enabled = True
    End If
End Sub

Private Sub txtEmailServer_LostFocus()
    If m_strEmailServer <> txtEmailServer.Text Then
        m_strEmailServer = txtEmailServer.Text
        cmdApply.Enabled = True
    End If
End Sub

Private Sub txtEmailServerPort_LostFocus()
    If m_strEmailServerPort <> txtEmailServerPort.Text Then
        m_strEmailServerPort = txtEmailServerPort.Text
        cmdApply.Enabled = True
    End If
End Sub
