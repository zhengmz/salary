VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   $"frmAbout.frx":0000
   ClientHeight    =   3555
   ClientLeft      =   3780
   ClientTop       =   3315
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":034D
   LinkTopic       =   "frmAbout"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0359
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确  定"
      Default         =   -1  'True
      Height          =   345
      Left            =   4125
      TabIndex        =   0
      Top             =   2625
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "电  邮： "
      DragIcon        =   "frmAbout.frx":0C23
      Height          =   195
      Left            =   255
      TabIndex        =   9
      Top             =   3240
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "博  客： "
      DragIcon        =   "frmAbout.frx":0F2D
      Height          =   195
      Left            =   255
      TabIndex        =   8
      Top             =   3030
      Width           =   675
   End
   Begin VB.Label lblUrl 
      AutoSize        =   -1  'True
      Caption         =   "http://blog.sina.com.cn/zhengmz"
      DragIcon        =   "frmAbout.frx":1237
      Height          =   195
      Left            =   937
      TabIndex        =   7
      Top             =   3030
      Width           =   2370
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "zhengmz@hotmail.com"
      DragIcon        =   "frmAbout.frx":1541
      Height          =   195
      Left            =   937
      TabIndex        =   6
      Top             =   3240
      Width           =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   "应用程序描述"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   2
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "应用程序标题"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "版本"
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "警告: ..."
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   255
      TabIndex        =   3
      Top             =   2625
      Width           =   3630
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "关于 " & App.Title
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription.Caption = _
        "应用程序的功能是" & App.Comments & vbCrLf & vbCrLf & _
        "本软件已授权使用。" & vbCrLf & vbCrLf & _
        "                         制作者 " & Format(Now(), "yyyy.MM")
    lblDisclaimer.Caption = _
        "警  告： 版权所有(C) 2004.5-" & Format(Now(), "yyyy.MM") & vbCrLf & _
        "作  者： 郑明忠"
End Sub

Private Sub lblEmail_DragDrop(Source As Control, x As Single, Y As Single)
    If Source Is lblEmail Then
        With lblEmail
            .Font.Underline = False
            .ForeColor = vbBlack
            Call ShellExecute(0&, vbNullString, "mailto:" & .Caption, vbNullString, vbNullString, vbNormalFocus)
        End With
    End If
End Sub

Private Sub lblEmail_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
    If State = vbLeave Then
        With lblEmail
            .Drag vbEndDrag
            .Font.Underline = False
            .ForeColor = vbBlack
        End With
    End If
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With lblEmail
        .ForeColor = vbBlue
        .Font.Underline = True
        .Drag vbBeginDrag
    End With
End Sub

Private Sub lblUrl_DragDrop(Source As Control, x As Single, Y As Single)
    If Source Is lblUrl Then
        With lblUrl
            .Font.Underline = False
            .ForeColor = vbBlack
            Call ShellExecute(0&, vbNullString, .Caption, vbNullString, vbNullString, vbNormalFocus)
        End With
    End If
End Sub

Private Sub lblUrl_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
    If State = vbLeave Then
        With lblUrl
            .Drag vbEndDrag
            .Font.Underline = False
            .ForeColor = vbBlack
        End With
    End If
End Sub

Private Sub lblUrl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With lblUrl
        .ForeColor = vbBlue
        .Font.Underline = True
        .Drag vbBeginDrag
    End With
End Sub
