VERSION 5.00
Begin VB.Form frmInit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ϵͳ��ʼ��"
   ClientHeight    =   6015
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   8070
   Icon            =   "frmInit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
         Height          =   4335
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   7095
         Begin VB.TextBox txtLog 
            Height          =   3735
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   360
            Width           =   6375
         End
      End
      Begin VB.Label lbDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϊ��һ��ʹ��ϵͳ�ṩ��ʼ�����ܣ������ʹ�á�"
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
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   4875
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
      Begin VB.Frame fraDesc 
         Caption         =   "���߽���"
         Height          =   4815
         Left            =   240
         TabIndex        =   4
         Top             =   0
         Width           =   7095
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ϵͳ�Ѿ�ʹ�ã�������ʵ�����ݣ������á�"
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
            Index           =   3
            Left            =   1440
            TabIndex        =   13
            Top             =   1140
            Width           =   4095
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ȼ���Ĭ�ϳ�ʼ�����лָ���ԭʼ���ݣ���Щ����ֻ����һЩ"
            Height          =   195
            Index           =   5
            Left            =   1440
            TabIndex        =   11
            Top             =   2220
            Width           =   4860
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "˵����"
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
            TabIndex        =   10
            Top             =   1680
            Width           =   585
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ע�⣺"
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
            TabIndex        =   9
            Top             =   1140
            Width           =   585
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������Ϊ��һ��ʹ��ϵͳ�ṩ��ʼ�����ܡ�"
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
            TabIndex        =   8
            Top             =   600
            Width           =   3705
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ʹ�ñ����ߣ���ɾ��ϵͳ���е�ȫ�����ݡ�"
            Height          =   195
            Index           =   4
            Left            =   1440
            TabIndex        =   6
            Top             =   1680
            Width           =   3420
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ϵͳ����Ļ����������ݡ�"
            Height          =   195
            Index           =   6
            Left            =   1440
            TabIndex        =   5
            Top             =   2760
            Width           =   2160
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
      TabIndex        =   14
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_strMd5 As String = "d9b847d79d9b8a0da23bec1570bb17ff"

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If cmdNext.Caption = "��һ��" Then
        cmdNext.Caption = "��ʼ��"
        picOptions(0).Left = -20000
        picOptions(1).Left = 210
        Exit Sub
    End If
    
    Dim strFileName As String
    Dim strRetMsg As String
    Dim strMd5FileName As String
    Dim strTableName As String

    strRetMsg = ""
    txtLog.Text = ""
    strFileName = App.Path & "\Init.dat"
    strMd5FileName = App.Path & "\Init.md5"
    strTableName = "ALL"

    On Error GoTo ErrHandle

    If Dir(strFileName) = "" Then
        strRetMsg = "ԭʼ�ļ� '" & strFileName & "' �����ڣ�ϵͳ�ѱ��ƻ��������°�װ��"
        txtLog.Text = strRetMsg
        DisplayMsg strRetMsg
        Exit Sub
    End If
    '������֤�ļ�
    Dim clsMD5 As CMD5
    Dim fsObj As FileSystemObject
    Dim ts As TextStream

    Set fsObj = CreateObject("Scripting.FileSystemObject")
    Set clsMD5 = New CMD5

    Set ts = fsObj.OpenTextFile(strMd5FileName, ForWriting, True)
    ts.WriteLine (strTableName)
    ts.WriteLine (clsMD5.DigestStrToHexStr("TABLE:" & strTableName))
    ts.WriteLine (m_strMd5)
    ts.Close
    
    '���ûָ����ܽ���ϵͳ��ʼ��
    If gRecover(strFileName, strRetMsg) = True Then
        txtLog.Text = "ϵͳ��ʼ���ɹ�����ӭ��ʼʹ�á�"
        DisplayMsg txtLog.Text
        
        '���ԭ������
        Dim clsRegConfig As CRegConfig
        Set clsRegConfig = New CRegConfig

        clsRegConfig.DelConfig
    Else
        DisplayMsg "ϵͳ��ʼ��ʧ��", vbExclamation
        txtLog.Text = "ϵͳ��ʼ��ʧ�ܣ����������Ϣ���£�" & strRetMsg
    End If
    'ɾ����֤�ļ�
    fsObj.DeleteFile strMd5FileName
    Exit Sub
    
ErrHandle:
    DisplayMsg "ϵͳ��ʼ��ʱ����!", vbCritical
End Sub

Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

