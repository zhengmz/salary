VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMaint 
   Caption         =   "Ա��������Ϣ"
   ClientHeight    =   4905
   ClientLeft      =   3180
   ClientTop       =   2955
   ClientWidth     =   9645
   Icon            =   "frmMaint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   9645
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   9645
      TabIndex        =   7
      Top             =   0
      Width           =   9645
      Begin VB.Frame Frame1 
         Caption         =   "��ѡ�����ļ�"
         Height          =   975
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   9135
         Begin VB.CommandButton cmdLoad 
            Caption         =   "���µ���"
            Height          =   420
            Left            =   5940
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "�� ��(&O)"
            Default         =   -1  'True
            Height          =   420
            Left            =   4320
            TabIndex        =   12
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtFileName 
            Height          =   375
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   390
            Width           =   3975
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "�������"
            Height          =   420
            Left            =   7560
            TabIndex        =   10
            Top             =   360
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   9645
      TabIndex        =   8
      Top             =   3675
      Width           =   9645
      Begin VB.CommandButton cmdExport 
         Caption         =   "�� ��(&E)"
         Height          =   420
         Left            =   6525
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "�� ��(&A)"
         Height          =   420
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "�� ��(&U)"
         Height          =   420
         Left            =   1815
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ ��(&D)"
         Height          =   420
         Left            =   3390
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ ��(&R)"
         Height          =   420
         Left            =   4950
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "�� ��(&C)"
         Height          =   420
         Left            =   8100
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Bindings        =   "frmMaint.frx":030A
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   8454016
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   4575
      Width           =   9645
      _ExtentX        =   17013
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
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dgFile 
      Left            =   5160
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_strTableName As String

Private Sub cmdClear_Click()
    If m_strTableName = "" Then
        DisplayMsg "��ѡ��Ҫ������ı�", vbExclamation
        Exit Sub
    End If
    If DisplayMsg("�Ƿ�ȷ����� [ " & m_strTableName & " ] �����ݣ�", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    Call gExecSql("Delete From " & m_strTableName)

    datPrimaryRS.Refresh
    Call DisplayGrid(grdDataGrid, m_strTableName)
End Sub

Private Sub cmdExport_Click()
    On Error GoTo Err_Handle
    If m_strTableName = "" Then
        DisplayMsg "��ѡ��Ҫ������ı�", vbExclamation
        Exit Sub
    End If
    With dgFile
        .CancelError = True
        .Filter = "����Excel�ļ� (*.xls)|*.xls"
        .DefaultExt = "xls"
        .DialogTitle = "��������"
        .Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
        .ShowSave
    End With
    
    Dim strFileName As String
    Dim strExportSql As String
    
    strFileName = Trim(dgFile.FileName)
    If strFileName = "" Then
        DisplayMsg "�����ļ�����Ϊ�գ�", vbExclamation
        Exit Sub
    End If
    
'    If Dir(strFileName) <> "" Then
'        'Kill (strFileName)
'        strExportSql = "insert into [Excel 8.0;database=" & strFileName & "].[sheet1] select * from " & m_strTableName
'    Else
'        strExportSql = "select * into [Excel 8.0;database=" & strFileName & "].[sheet1] from " & m_strTableName
'    End If
    
    If Dir(strFileName) <> "" Then
        Kill (strFileName)
    End If
    strExportSql = "select * into [Excel 8.0;database=" & strFileName & "].[sheet1] from " & m_strTableName
    gAdoConnDB.BeginTrans
    gAdoConnDB.Execute strExportSql
    gAdoConnDB.CommitTrans

    DisplayMsg "�ɹ��������ļ� " & strFileName & "��", vbInformation
    Exit Sub

Err_Handle:
    If Err.Number = 32755 Then
    '����ȡ��
        Exit Sub
    End If
    DisplayMsg "��������ʱ����", vbCritical
End Sub

Private Sub cmdLoad_Click()
    If txtFileName.Text = Null Or txtFileName.Text = "" Then
        DisplayMsg "��ѡ�������ļ���", vbExclamation
        cmdOpen.SetFocus
        Exit Sub
    End If
    If Dir(txtFileName.Text) = "" Then
        DisplayMsg "�����ļ�[ " & txtFileName.Text & " ]�����ڣ�", vbExclamation
        cmdOpen.SetFocus
        Exit Sub
    End If

    If m_strTableName = "" Then
        DisplayMsg "��ѡ��Ҫ������ı�", vbExclamation
        Exit Sub
    End If
    
    cmdLoad.Enabled = True
    
    Dim strSQL As String
'    strSQL = "insert into " & m_strTableName & " select * From [Excel 8.0;database=" & txtFileName.Text & "].[sheet1$]"
'    strSQL = "insert into " & m_strTableName & "(emp_id,emp_name,email) select Ա������,����,�������� From [Excel 8.0;database=" & txtFileName.Text & "].[sheet1$]"
'    strSQL = "insert into " & m_strTableName & "(emp_id,emp_name,email) select Ա������,����,'zhengmz' From [Excel 8.0;database=" & txtFileName.Text & "].[sheet1$]"
'    Call gExecSql(strSQL)
    Dim strRetMsg As String
    Me.MousePointer = vbHourglass
    'Call gImportData("select * From [Excel 8.0;HDR=yes;IMEX=1;database=" & txtFileName.Text & "].[sheet1$]", "select * from emp", 3, strRetMsg)
    Call gImportData("select * From [Excel 8.0;HDR=yes;IMEX=1;database=" & txtFileName.Text & "].[sheet1$]", "select emp_id, emp_name, emp_email from emp", 3, strRetMsg)
    Me.MousePointer = vbDefault
    datPrimaryRS.Refresh
    Call DisplayGrid(grdDataGrid, m_strTableName)

    If strRetMsg <> "" Then
        Load frmMsgLog
        frmMsgLog.Caption = "������Ϣ˵��"
        frmMsgLog.RTDesc.Text = "������Ϣ˵����" & vbCrLf & "�����ļ�Ϊ��" & txtFileName.Text & vbCrLf & strRetMsg
        frmMsgLog.Show vbModal
    End If
End Sub

Private Sub cmdOpen_Click()
    On Error GoTo ErrHandle
    With dgFile
        .FileName = ""
        .Filter = "�����ļ� (*.xls)|*.xls"
        .DialogTitle = "�������ļ�"
        .CancelError = True
        .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNFileMustExist
        .ShowOpen
        txtFileName.Text = .FileName
    End With

    Call cmdLoad_Click
    Exit Sub

ErrHandle:
    If Err.Number = 32755 Then
    '����ȡ��
        dgFile.FileName = ""
        Exit Sub
    End If
    DisplayMsg "����ʱ����!", vbCritical
End Sub

Private Sub Form_Load()
    m_strTableName = "emp"
    cmdLoad.Enabled = False

    datPrimaryRS.ConnectionString = gStrConnDB
    datPrimaryRS.CommandType = adCmdText
    datPrimaryRS.RecordSource = "select * from emp order by emp_id"
    datPrimaryRS.Refresh
    Call DisplayGrid(grdDataGrid, m_strTableName)
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  '���������ʱ���������
  grdDataGrid.Top = picTop.Height
  grdDataGrid.Left = 0
  grdDataGrid.Width = Me.ScaleWidth
  grdDataGrid.Height = Me.ScaleHeight - datPrimaryRS.Height - picButtons.Height - picTop.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  '���������������ڴ˴�
  '��Ҫ���Դ���ע�͵���һ��
  '��Ҫ�������ǣ��ڴ���Ӵ����Դ�������
  DisplayMsg "���ݿ����", vbCritical
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Ϊ��� recordset ��ʾ��ǰ��¼λ��
  If datPrimaryRS.Recordset.RecordCount = 0 Then
    datPrimaryRS.Caption = "�޼�¼"
  Else
    datPrimaryRS.Caption = "��ǰ��¼λ��: " & CStr(datPrimaryRS.Recordset.AbsolutePosition) & _
                         "/" & CStr(datPrimaryRS.Recordset.RecordCount)
  End If
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  '��֤�������ڴ˴�
  '���ж�������ʱ���¼�������
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.MoveLast
  grdDataGrid.SetFocus
  SendKeys "{down}"

  Exit Sub
AddErr:
  DisplayMsg "���Ӽ�¼����", vbCritical
End Sub

Private Sub cmdDelete_Click()
  If DisplayMsg("Ҫɾ����ǰ��¼��", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
    Exit Sub
  End If

  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  DisplayMsg "ɾ����¼����", vbCritical
End Sub

Private Sub cmdRefresh_Click()
  'ֻ�ж��û�Ӧ�ó�����Ҫ
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Call DisplayGrid(grdDataGrid, m_strTableName)
  Exit Sub
RefreshErr:
  DisplayMsg "ˢ�¼�¼����", vbCritical
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  DisplayMsg "���¼�¼����", vbCritical
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

