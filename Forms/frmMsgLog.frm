VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMsgLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "更新说明"
   ClientHeight    =   7455
   ClientLeft      =   2955
   ClientTop       =   2085
   ClientWidth     =   8370
   Icon            =   "frmMsgLog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "确 定(&C)"
      Default         =   -1  'True
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   6720
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RTDesc 
      Height          =   6375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   11245
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMsgLog.frx":000C
   End
End
Attribute VB_Name = "frmMsgLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

