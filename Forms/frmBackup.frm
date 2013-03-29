VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "备份与恢复"
   ClientHeight    =   6015
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   8070
   Icon            =   "frmBackup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一步"
      Height          =   375
      Left            =   3360
      TabIndex        =   30
      Top             =   5415
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cbTab 
      Height          =   315
      ItemData        =   "frmBackup.frx":000C
      Left            =   240
      List            =   "frmBackup.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dgFile 
      Left            =   2040
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   5100
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   5100
      ScaleWidth      =   7605
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Width           =   7605
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   5100
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   5100
      ScaleWidth      =   7605
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   240
      Width           =   7605
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
      Begin VB.Frame fraLog 
         Caption         =   "日志信息"
         Height          =   3735
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   7095
         Begin VB.TextBox txtLog 
            Height          =   3135
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   360
            Width           =   6375
         End
      End
      Begin VB.Frame fraStep2 
         Caption         =   "请选择策略"
         Height          =   1035
         Left            =   240
         TabIndex        =   6
         Top             =   0
         Width           =   7095
         Begin VB.ComboBox cbTable 
            Height          =   315
            ItemData        =   "frmBackup.frx":0010
            Left            =   2160
            List            =   "frmBackup.frx":0012
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   570
            Width           =   3375
         End
         Begin VB.OptionButton OptMethod 
            Caption         =   "表级"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   26
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton OptMethod 
            Caption         =   "完整"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   25
            Top             =   240
            Width           =   1335
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
      Begin VB.Frame fraChoice 
         Height          =   855
         Left            =   240
         TabIndex        =   20
         Top             =   3960
         Width           =   7095
         Begin VB.OptionButton optChoice 
            Caption         =   "恢复"
            Height          =   195
            Index           =   1
            Left            =   3600
            TabIndex        =   23
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "备份"
            Height          =   195
            Index           =   0
            Left            =   1920
            TabIndex        =   22
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "请选择："
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
            Left            =   480
            TabIndex        =   21
            Top             =   360
            Width           =   780
         End
      End
      Begin VB.Frame fraDesc 
         Caption         =   "工具介绍"
         Height          =   3975
         Left            =   240
         TabIndex        =   7
         Top             =   0
         Width           =   7095
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
            Left            =   480
            TabIndex        =   24
            Top             =   3600
            Visible         =   0   'False
            Width           =   5265
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "文件进行验证，通过验证才进行恢复。"
            Height          =   195
            Index           =   10
            Left            =   1440
            TabIndex        =   19
            Top             =   3120
            Width           =   3060
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "此两个文件都必须保存好，并存放在同一个目录下。"
            Height          =   195
            Index           =   9
            Left            =   1440
            TabIndex        =   18
            Top             =   2040
            Width           =   4140
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "表级备份，只对一张表的数据进行备份；"
            Height          =   195
            Index           =   5
            Left            =   1440
            TabIndex        =   17
            Top             =   1425
            Width           =   3240
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "恢复："
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
            Top             =   2520
            Width           =   585
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "备份："
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
            TabIndex        =   15
            Top             =   720
            Width           =   585
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "本工具提供对系统数据的备份和恢复功能。"
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
            TabIndex        =   14
            Top             =   360
            Width           =   3705
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "主要原理就是将备份后的信息恢复到系统中，在恢复前，会对数据"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   12
            Top             =   2880
            Width           =   5220
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "分为完整备份与表级备份。"
            Height          =   195
            Index           =   3
            Left            =   1440
            TabIndex        =   11
            Top             =   720
            Width           =   2160
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "完整备份，请对整个数据文件进行备份；"
            Height          =   195
            Index           =   4
            Left            =   1440
            TabIndex        =   10
            Top             =   1065
            Width           =   3240
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "备份将产生两个文件，一个是数据文件，另一个是验证文件；"
            Height          =   195
            Index           =   6
            Left            =   1440
            TabIndex        =   9
            Top             =   1770
            Width           =   4860
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "分为完整恢复和表级恢复两类。"
            Height          =   195
            Index           =   8
            Left            =   1440
            TabIndex        =   8
            Top             =   2520
            Width           =   2520
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
      Left            =   4860
      TabIndex        =   0
      Top             =   5415
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   5175
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iBakFlag As Integer
Private m_iMethod As Integer

Private Sub cbTable_LostFocus()
    cbTab.ListIndex = cbTable.ListIndex
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    Dim strCaption As String

    If m_iBakFlag = 0 Then
        strCaption = "备份"
    Else
        strCaption = "恢复"
    End If
    If cmdNext.Caption = "下一步" Then
        Me.Caption = "备份与恢复 - " & strCaption
        cmdNext.Caption = strCaption
        fraStep2.Caption = "请选择" & strCaption & "策略"
        OptMethod(0).Caption = "完整" & strCaption
        OptMethod(1).Caption = "表级" & strCaption
        picOptions(0).Left = -20000
        picOptions(1).Left = 210
        cmdPrevious.Visible = True
        Exit Sub
    End If

    If m_iMethod = 1 And cbTab.Text = "" Then
        DisplayMsg "请选择要处理的表", vbExclamation
        cbTable.SetFocus
        Exit Sub
    End If
    
    Dim strFileName As String
    Dim strRetMsg As String
    Dim strTableName As String

    strRetMsg = ""
    txtLog.Text = ""
    strFileName = ""
    If m_iMethod = 0 Then   '完整
        strTableName = "ALL"
    Else                    '表级
        strTableName = cbTab.Text
    End If

    On Error GoTo ErrHandle
    With dgFile
        .FileName = ""
        If Dir(App.Path & "\Backup", vbDirectory) <> "" Then
            .InitDir = App.Path & "\Backup"
        Else
            .InitDir = App.Path
        End If
        .Filter = "数据文件 (*.dat)|*.dat"
    End With

    Select Case m_iBakFlag
    Case 0  '备份
        With dgFile
            .FileName = .InitDir & "\" & strTableName & "_" & Format(Date, "yyyyMMdd") & ".dat"
            .DialogTitle = "备份数据文件"
            .CancelError = True
            .Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
            .ShowSave
            strFileName = .FileName
        End With

        If gBackup(strFileName, strRetMsg, strTableName) = True Then
            DisplayMsg "备份成功"
            txtLog.Text = "成功备份。" & vbCrLf & "备份数据文件为 '" & strFileName & "'。"
        Else
            DisplayMsg "备份失败", vbExclamation
            txtLog.Text = "备份文件时出错，具体信息如下：" & vbCrLf & strRetMsg
        End If
    Case 1  '恢复
        With dgFile
            .DialogTitle = "打开数据文件"
            .CancelError = True
            .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNFileMustExist
            .ShowOpen
            strFileName = .FileName
        End With

        If gRecover(strFileName, strRetMsg, strTableName) = True Then
            DisplayMsg "恢复成功"
            txtLog.Text = "恢复成功。" & vbCrLf & "成功从文件 '" & strFileName & "' 恢复数据。"
            
            '刷新数据
            Call GetTabList
        Else
            DisplayMsg "恢复失败", vbExclamation
            txtLog.Text = "恢复文件时出错，具体信息如下：" & vbCrLf & strRetMsg
        End If
    End Select
    Exit Sub
    
ErrHandle:
    If Err.Number = 32755 Then
    '按了取消
        dgFile.FileName = ""
        Exit Sub
    End If
    DisplayMsg strCaption & "时出错!", vbCritical
End Sub

Private Sub cmdPrevious_Click()
    picOptions(0).Left = 210
    picOptions(1).Left = -20000
    cmdNext.Caption = "下一步"
    cmdPrevious.Visible = False
    Me.Caption = "备份与恢复"
    txtLog.Text = ""
End Sub

Private Sub Form_Load()
    '默认为备份
    optChoice(0).value = True
    optChoice(1).value = False
    m_iBakFlag = 0          '标识，0为备份，1为恢复
    
    '默认为完整
    OptMethod(0).value = True
    OptMethod(1).value = False
    m_iMethod = 0           '方式，0为完整，1为表级
    cbTable.Enabled = False
    
    Call GetTabList

    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub optChoice_Click(Index As Integer)
    m_iBakFlag = Index
End Sub

Private Sub OptMethod_Click(Index As Integer)
    m_iMethod = Index
    If m_iMethod = 0 Then
        cbTable.Enabled = False
    Else
        cbTable.Enabled = True
    End If
End Sub

Private Sub GetTabList()
    Dim strSQL As String
    Dim rsTab As New ADODB.Recordset
    strSQL = "select dict_key as table_name, dict_value as table_desc from dicts where dict_sect='TABLE_NAME' and dict_flag>0 order by dict_flag"
    rsTab.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    cbTable.Clear
    cbTab.Clear
    While rsTab.EOF <> True
        cbTable.AddItem rsTab("table_desc")
        cbTab.AddItem rsTab("table_name")
        rsTab.MoveNext
    Wend
    rsTab.Close
    Set rsTab = Nothing
End Sub
