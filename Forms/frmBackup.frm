VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ָ�"
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
      Caption         =   "��һ��"
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
         Caption         =   "��־��Ϣ"
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
         Caption         =   "��ѡ�����"
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
            Caption         =   "��"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   26
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton OptMethod 
            Caption         =   "����"
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
            Caption         =   "�ָ�"
            Height          =   195
            Index           =   1
            Left            =   3600
            TabIndex        =   23
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "����"
            Height          =   195
            Index           =   0
            Left            =   1920
            TabIndex        =   22
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "��ѡ��"
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
         Caption         =   "���߽���"
         Height          =   3975
         Left            =   240
         TabIndex        =   7
         Top             =   0
         Width           =   7095
         Begin VB.Label lbBackup 
            AutoSize        =   -1  'True
            Caption         =   "�������޸�ϵͳ֮ǰ�������ݽ��б��ݣ������������ı��ݡ�"
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
            Caption         =   "�ļ�������֤��ͨ����֤�Ž��лָ���"
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
            Caption         =   "�������ļ������뱣��ã��������ͬһ��Ŀ¼�¡�"
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
            Caption         =   "�����ݣ�ֻ��һ�ű�����ݽ��б��ݣ�"
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
            Caption         =   "�ָ���"
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
            Caption         =   "���ݣ�"
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
            Caption         =   "�������ṩ��ϵͳ���ݵı��ݺͻָ����ܡ�"
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
            Caption         =   "��Ҫԭ����ǽ����ݺ����Ϣ�ָ���ϵͳ�У��ڻָ�ǰ���������"
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
            Caption         =   "��Ϊ��������������ݡ�"
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
            Caption         =   "�������ݣ�������������ļ����б��ݣ�"
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
            Caption         =   "���ݽ����������ļ���һ���������ļ�����һ������֤�ļ���"
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
            Caption         =   "��Ϊ�����ָ��ͱ��ָ����ࡣ"
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
      Caption         =   "�ر�(&C)"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   5415
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ��"
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
        strCaption = "����"
    Else
        strCaption = "�ָ�"
    End If
    If cmdNext.Caption = "��һ��" Then
        Me.Caption = "������ָ� - " & strCaption
        cmdNext.Caption = strCaption
        fraStep2.Caption = "��ѡ��" & strCaption & "����"
        OptMethod(0).Caption = "����" & strCaption
        OptMethod(1).Caption = "��" & strCaption
        picOptions(0).Left = -20000
        picOptions(1).Left = 210
        cmdPrevious.Visible = True
        Exit Sub
    End If

    If m_iMethod = 1 And cbTab.Text = "" Then
        DisplayMsg "��ѡ��Ҫ����ı�", vbExclamation
        cbTable.SetFocus
        Exit Sub
    End If
    
    Dim strFileName As String
    Dim strRetMsg As String
    Dim strTableName As String

    strRetMsg = ""
    txtLog.Text = ""
    strFileName = ""
    If m_iMethod = 0 Then   '����
        strTableName = "ALL"
    Else                    '��
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
        .Filter = "�����ļ� (*.dat)|*.dat"
    End With

    Select Case m_iBakFlag
    Case 0  '����
        With dgFile
            .FileName = .InitDir & "\" & strTableName & "_" & Format(Date, "yyyyMMdd") & ".dat"
            .DialogTitle = "���������ļ�"
            .CancelError = True
            .Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
            .ShowSave
            strFileName = .FileName
        End With

        If gBackup(strFileName, strRetMsg, strTableName) = True Then
            DisplayMsg "���ݳɹ�"
            txtLog.Text = "�ɹ����ݡ�" & vbCrLf & "���������ļ�Ϊ '" & strFileName & "'��"
        Else
            DisplayMsg "����ʧ��", vbExclamation
            txtLog.Text = "�����ļ�ʱ����������Ϣ���£�" & vbCrLf & strRetMsg
        End If
    Case 1  '�ָ�
        With dgFile
            .DialogTitle = "�������ļ�"
            .CancelError = True
            .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNFileMustExist
            .ShowOpen
            strFileName = .FileName
        End With

        If gRecover(strFileName, strRetMsg, strTableName) = True Then
            DisplayMsg "�ָ��ɹ�"
            txtLog.Text = "�ָ��ɹ���" & vbCrLf & "�ɹ����ļ� '" & strFileName & "' �ָ����ݡ�"
            
            'ˢ������
            Call GetTabList
        Else
            DisplayMsg "�ָ�ʧ��", vbExclamation
            txtLog.Text = "�ָ��ļ�ʱ����������Ϣ���£�" & vbCrLf & strRetMsg
        End If
    End Select
    Exit Sub
    
ErrHandle:
    If Err.Number = 32755 Then
    '����ȡ��
        dgFile.FileName = ""
        Exit Sub
    End If
    DisplayMsg strCaption & "ʱ����!", vbCritical
End Sub

Private Sub cmdPrevious_Click()
    picOptions(0).Left = 210
    picOptions(1).Left = -20000
    cmdNext.Caption = "��һ��"
    cmdPrevious.Visible = False
    Me.Caption = "������ָ�"
    txtLog.Text = ""
End Sub

Private Sub Form_Load()
    'Ĭ��Ϊ����
    optChoice(0).value = True
    optChoice(1).value = False
    m_iBakFlag = 0          '��ʶ��0Ϊ���ݣ�1Ϊ�ָ�
    
    'Ĭ��Ϊ����
    OptMethod(0).value = True
    OptMethod(1).value = False
    m_iMethod = 0           '��ʽ��0Ϊ������1Ϊ��
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
