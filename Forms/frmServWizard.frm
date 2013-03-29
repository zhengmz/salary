VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmServWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "服务配置向导"
   ClientHeight    =   6180
   ClientLeft      =   4230
   ClientTop       =   3240
   ClientWidth     =   8400
   Icon            =   "frmServWizard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc adoServType 
      Height          =   330
      Left            =   480
      Top             =   5880
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
   Begin MSAdodcLib.Adodc adoSalary 
      Height          =   330
      Left            =   1320
      Top             =   5520
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
      Left            =   120
      Top             =   5520
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
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4860
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   4860
      ScaleWidth      =   7965
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   7965
      Begin VB.Frame fraStep4 
         Caption         =   "保存信息"
         Height          =   4545
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   7455
         Begin VB.TextBox txtServTemplFn 
            BackColor       =   &H80000011&
            Height          =   285
            Left            =   1680
            MaxLength       =   200
            TabIndex        =   52
            Top             =   1560
            Width           =   4575
         End
         Begin VB.OptionButton optServPeriod 
            Caption         =   "不定期"
            Height          =   255
            Index           =   2
            Left            =   4200
            TabIndex        =   51
            Top             =   2040
            Width           =   1215
         End
         Begin VB.OptionButton optServPeriod 
            Caption         =   "每月"
            Height          =   255
            Index           =   1
            Left            =   2700
            TabIndex        =   50
            Top             =   2040
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optServPeriod 
            Caption         =   "每年"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   48
            Top             =   2040
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dcbServType 
            Bindings        =   "frmServWizard.frx":000C
            Height          =   315
            Left            =   960
            TabIndex        =   27
            Top             =   330
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.CheckBox chkServFlag 
            Caption         =   "是否设为默认"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   4080
            Width           =   2775
         End
         Begin VB.TextBox txtServDesc 
            Height          =   855
            Left            =   1680
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   3000
            Width           =   5415
         End
         Begin VB.TextBox txtServSubject 
            Height          =   285
            Left            =   1680
            MaxLength       =   200
            TabIndex        =   34
            Top             =   2505
            Width           =   5415
         End
         Begin VB.TextBox txtServSheet 
            BackColor       =   &H80000011&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4440
            TabIndex        =   32
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtServID 
            BackColor       =   &H80000011&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   30
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cbServName 
            Height          =   315
            Left            =   3480
            TabIndex        =   28
            Top             =   330
            Width           =   3735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "模板文件名："
            Height          =   195
            Left            =   240
            TabIndex        =   53
            Top             =   1605
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "周期："
            Height          =   195
            Left            =   240
            TabIndex        =   49
            Top             =   2070
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "描述："
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   3015
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "默认邮件主题："
            Height          =   195
            Left            =   240
            TabIndex        =   33
            Top             =   2520
            Width           =   1260
         End
         Begin VB.Label Label2 
            Caption         =   "Sheet名称："
            Height          =   255
            Left            =   3120
            TabIndex        =   31
            Top             =   1095
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "标识ID："
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   1095
            Width           =   705
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   7320
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "名称："
            Height          =   195
            Left            =   2880
            TabIndex        =   26
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lbType 
            AutoSize        =   -1  'True
            Caption         =   "类型："
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   540
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4860
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   4860
      ScaleWidth      =   7965
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   7965
      Begin VB.Frame fraStep3_2 
         Caption         =   "预览"
         Height          =   1455
         Left            =   240
         TabIndex        =   23
         Top             =   3240
         Width           =   7455
         Begin MSDataGridLib.DataGrid grdBrower 
            Bindings        =   "frmServWizard.frx":0026
            Height          =   975
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   1720
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   0   'False
            HeadLines       =   1
            RowHeight       =   15
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
      Begin VB.Frame fraStep3_1 
         Caption         =   "映射关系"
         Height          =   2985
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   7455
         Begin MSDataGridLib.DataGrid grdRelation 
            Bindings        =   "frmServWizard.frx":003E
            Height          =   2415
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   4260
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
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
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4860
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   4860
      ScaleWidth      =   7965
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   7965
      Begin VB.Frame fraStep2_fld 
         Caption         =   "请选择要导入的字段"
         Height          =   3735
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   7455
         Begin VB.CommandButton cmdChoice 
            Caption         =   "<<"
            Height          =   375
            Index           =   3
            Left            =   3480
            TabIndex        =   20
            Top             =   2640
            Width           =   495
         End
         Begin VB.CommandButton cmdChoice 
            Caption         =   "<"
            Height          =   375
            Index           =   2
            Left            =   3480
            TabIndex        =   19
            Top             =   2040
            Width           =   495
         End
         Begin VB.CommandButton cmdChoice 
            Caption         =   ">"
            Height          =   375
            Index           =   1
            Left            =   3480
            TabIndex        =   18
            Top             =   1440
            Width           =   495
         End
         Begin VB.CommandButton cmdChoice 
            Caption         =   ">>"
            Height          =   375
            Index           =   0
            Left            =   3480
            TabIndex        =   17
            Top             =   840
            Width           =   495
         End
         Begin VB.ListBox lstChoice 
            Height          =   3180
            Left            =   4320
            MultiSelect     =   2  'Extended
            TabIndex        =   16
            Top             =   360
            Width           =   2895
         End
         Begin VB.ListBox lstField 
            Height          =   3180
            Left            =   240
            MultiSelect     =   2  'Extended
            TabIndex        =   15
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame fraStep2_tbl 
         Caption         =   "请选择表"
         Height          =   705
         Left            =   240
         TabIndex        =   8
         Top             =   120
         Width           =   5055
         Begin VB.ComboBox cbTable 
            Height          =   315
            ItemData        =   "frmServWizard.frx":0058
            Left            =   360
            List            =   "frmServWizard.frx":005A
            TabIndex        =   21
            Text            =   "cbTable"
            Top             =   240
            Width           =   4335
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4860
      Index           =   0
      Left            =   210
      ScaleHeight     =   4860
      ScaleWidth      =   7965
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   7965
      Begin VB.Frame fraDesc 
         Caption         =   "向导说明"
         Height          =   2895
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   7455
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "保存为导入模板，以便之后的数据导入和显示。"
            Height          =   195
            Index           =   8
            Left            =   1440
            TabIndex        =   47
            Top             =   2400
            Width           =   3780
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "第四步"
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
            TabIndex        =   46
            Top             =   2400
            Width           =   585
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "可以查看字段的映射关系，并预览显示的结果；"
            Height          =   195
            Index           =   6
            Left            =   1440
            TabIndex        =   45
            Top             =   1890
            Width           =   3780
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "第三步"
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
            Index           =   5
            Left            =   480
            TabIndex        =   44
            Top             =   1890
            Width           =   585
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "如第一个字段一般为员工编码，而第二个字段一般为姓名等；"
            Height          =   195
            Index           =   4
            Left            =   1440
            TabIndex        =   43
            Top             =   1380
            Width           =   4860
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择所要导入的字段，结果列表中，是有顺序的要求，"
            Height          =   195
            Index           =   3
            Left            =   1440
            TabIndex        =   42
            Top             =   870
            Width           =   4320
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "第二步"
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
            TabIndex        =   41
            Top             =   870
            Width           =   585
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选择所要导入的模板文件，格式为Excel，第一行为标题；"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   40
            Top             =   360
            Width           =   4530
         End
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "第一步"
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
            TabIndex        =   39
            Top             =   360
            Width           =   585
         End
      End
      Begin VB.Frame fraStep1 
         Caption         =   "请选择模板文件"
         Height          =   975
         Left            =   240
         TabIndex        =   11
         Top             =   3480
         Width           =   7455
         Begin VB.CommandButton cmdOpen 
            Caption         =   "浏览"
            Height          =   405
            Left            =   6240
            TabIndex        =   13
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox txtFileName 
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   360
            Width           =   5775
         End
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存退出"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   5655
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一步"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   5655
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一步"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   5655
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   5325
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9393
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第一步"
            Key             =   "Step1"
            Object.ToolTipText     =   "选模板文件"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第二步"
            Key             =   "Step2"
            Object.ToolTipText     =   "选字段"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第三步"
            Key             =   "Step3"
            Object.ToolTipText     =   "选对应关系"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "第四步"
            Key             =   "Step4"
            Object.ToolTipText     =   "保存配置"
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
   Begin MSComDlg.CommonDialog dgFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmServWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_strFileName As String
Private m_strTableName As String
Private m_strServId As String
Private m_strServPeriod As String
Private m_blChoiceUpdate As Boolean

Private Sub cbServName_Validate(Cancel As Boolean)
    If cbServName.ListIndex = -1 And cbServName.Text = "" Then
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
End Sub

Private Sub cbTable_Validate(Cancel As Boolean)
    If cbTable.Text <> cbTable.List(cbTable.ListIndex) Then
        DisplayMsg "请选择列表内的内容", vbExclamation
        Cancel = True
        Exit Sub
    End If
    If cbTable.Text = m_strTableName Then
        Exit Sub
    End If
    m_strTableName = cbTable.Text
    lstField.Clear
    lstChoice.Clear
    If cbTable.ListIndex = -1 Then
        Exit Sub
    End If

    '获取列
    Dim cnExcel As ADODB.Connection
    Dim rsColumns As ADODB.Recordset
    Dim strColumnName As String
    Dim iColumnPos As Integer
    Dim iColumnCount As Integer
    Dim strColName() As String
    Set cnExcel = New ADODB.Connection
    With cnExcel
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & m_strFileName & _
            ";Extended Properties=Excel 8.0;"
        .CursorLocation = adUseClient
        .Open
    End With

    Set rsColumns = cnExcel.OpenSchema(adSchemaColumns, Array(Empty, Empty, m_strTableName, Empty))
    iColumnCount = rsColumns.RecordCount
    If iColumnCount > 0 Then
        ReDim strColName(iColumnCount - 1) As String
    End If
    While rsColumns.EOF <> True
        strColumnName = rsColumns!COLUMN_NAME
        iColumnPos = rsColumns!ORDINAL_POSITION
        strColName(iColumnPos - 1) = strColumnName
        rsColumns.MoveNext
    Wend
    rsColumns.Close
    cnExcel.Close
    Set rsColumns = Nothing
    Set cnExcel = Nothing
    
    '按列的位置排序列出
    Dim i As Integer
    For i = 0 To iColumnCount - 1
        strColumnName = strColName(i)
        If strColumnName <> "F" & (i + 1) Then
            lstField.AddItem strColumnName
        End If
    Next
End Sub

Private Sub cmdChoice_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
    '>>
    Case 0
        For i = 0 To lstField.ListCount - 1
            lstChoice.AddItem lstField.List(i)
        Next
        lstField.Clear
    '>
    Case 1
        For i = 0 To lstField.ListCount - 1
            If lstField.Selected(i) = True Then
                lstChoice.AddItem lstField.List(i)
            End If
        Next
        For i = lstField.ListCount - 1 To 0 Step -1
            If lstField.Selected(i) = True Then
                lstField.RemoveItem i
            End If
        Next
    '<
    Case 2
        For i = 0 To lstChoice.ListCount - 1
            If lstChoice.Selected(i) = True Then
                lstField.AddItem lstChoice.List(i)
            End If
        Next
        For i = lstChoice.ListCount - 1 To 0 Step -1
            If lstChoice.Selected(i) = True Then
                lstChoice.RemoveItem i
            End If
        Next
    '<<
    Case 3
        For i = 0 To lstChoice.ListCount - 1
            lstField.AddItem lstChoice.List(i)
        Next
        lstChoice.Clear
    End Select
    m_blChoiceUpdate = True
End Sub

Private Sub cmdNext_Click()
    Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.SelectedItem.Index + 1)
End Sub

Private Sub cmdOpen_Click()
    On Error GoTo ErrHandle

    With dgFile
        .FileName = ""
        If Dir(App.Path & "\Template", vbDirectory) <> "" Then
            .InitDir = App.Path & "\Template"
        End If
        .Filter = "模板文件 (*.xls)|*.xls"
        .DialogTitle = "打开模板文件"
        .CancelError = True
        .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNFileMustExist
        .ShowOpen
        txtFileName.Text = .FileName
    End With
    txtFileName.SetFocus
    Exit Sub
    
ErrHandle:
    If Err.Number = 32755 Then
    '按了取消
        dgFile.FileName = ""
        Exit Sub
    End If
    DisplayMsg "打开文件时出错!", vbCritical
End Sub

Private Sub cmdPrevious_Click()
    Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.SelectedItem.Index - 1)
End Sub

Private Sub cmdSave_Click()
    If dcbServType.Text = "" Then
        DisplayMsg "请选择类型", vbExclamation
        dcbServType.SetFocus
        Exit Sub
    End If
    If cbServName.Text = "" Then
        DisplayMsg "请填写名称", vbExclamation
        cbServName.SetFocus
        Exit Sub
    End If

    On Error GoTo ErrHandle

    Dim strSQL As String
    gAdoConnDB.BeginTrans
    '修改序号
    strSQL = "Update dicts Set dict_value='" & Val(adoServType.Recordset("seq")) + 1 & "'" & _
             " Where dict_sect='OPT_SERV_TYPE' And dict_key='" & dcbServType.Text & "'"
    gAdoConnDB.Execute strSQL
    '修改默认标识
    If chkServFlag.value = 1 Then
        strSQL = "Update services Set default_flag=0 where default_flag=1"
        gAdoConnDB.Execute strSQL
    End If
    '增加记录
    strSQL = "Insert into services(serv_id,serv_type,serv_name,serv_sheet,serv_templ_fn,serv_subject,serv_desc,default_flag,serv_period)" & _
             " values('" & txtServId.Text & "','" & dcbServType.Text & "','" & cbServName.Text & _
             "','" & txtServSheet.Text & "','" & txtServTemplFn.Text & "','" & txtServSubject.Text & "','" & txtServDesc.Text & _
             "'," & chkServFlag.value & ",'" & m_strServPeriod & "')"
    gAdoConnDB.Execute strSQL
    '修改配置
    strSQL = "Update serv_field Set serv_id='" & txtServId.Text & "' Where serv_id='-1'"
    gAdoConnDB.Execute strSQL
    gAdoConnDB.CommitTrans

    On Error GoTo 0
    Unload Me
    Exit Sub
    
ErrHandle:
    If Err.Number = -2147467259 Then
        If gAdoConnDB.Errors(0).NativeError = -105121349 Then
            DisplayMsg "编号ID重复，请使用工具菜单的自动修复功能。具体错误信息如下：", vbCritical
        End If
    Else
        DisplayMsg "信息保存出错", vbCritical
    End If
    gAdoConnDB.RollbackTrans
End Sub

Private Sub dcbServType_Change()
    adoServType.Recordset.Move dcbServType.SelectedItem - 1, 1
    txtServId.Text = adoServType.Recordset("code") & Format(Val(adoServType.Recordset("seq")) + 1, SERV_ID_PATTERN)
    
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
    m_strFileName = ""
    m_strTableName = ""
    m_strServId = "-1"
    m_strServPeriod = "MONTH"
    m_blChoiceUpdate = False
    grdRelation.MarqueeStyle = dbgHighlightRow

    adoServConf.ConnectionString = gStrConnDB
    adoServConf.CommandType = adCmdText
    adoServConf.RecordSource = "select * from serv_field where serv_id='-1' order by field_name"
    adoSalary.ConnectionString = gStrConnDB
    adoSalary.CommandType = adCmdText
    adoSalary.RecordSource = "select * from salary where serv_id='-1'"
    adoServType.ConnectionString = gStrConnDB
    adoServType.CommandType = adCmdText
    adoServType.RecordSource = "select dict_key as type, dict_type as code, dict_value as seq" & _
                            " from dicts where dict_sect='OPT_SERV_TYPE' order by dict_flag"
    adoServConf.Refresh
    adoSalary.Refresh
    adoServType.Refresh
    dcbServType.ListField = "type"

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Call tbsOptions_Click
End Sub

Private Sub lstChoice_DblClick()
    Call cmdChoice_Click(2)
End Sub

Private Sub lstField_DblClick()
    Call cmdChoice_Click(1)
End Sub

Private Sub optServPeriod_Click(Index As Integer)
    Select Case Index
    Case 0
        m_strServPeriod = "YEAR"
    Case 1
        m_strServPeriod = "MONTH"
    Case 2
        m_strServPeriod = "DAY"
    End Select
End Sub

Private Sub tbsOptions_Click()
    Dim i As Integer
    '显示和生效相应的控制面板，隐藏和失效其他的面板。
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
    '业务判断
    If iCurrentTab > 1 And m_strFileName = "" Then
        DisplayMsg "请选择模板文件", vbExclamation
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        txtFileName.SetFocus
    ElseIf iCurrentTab > 2 Then
        If cbTable.ListIndex = -1 Then
            DisplayMsg "请选择所要导入的表名", vbExclamation
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(2)
            cbTable.SetFocus
        ElseIf lstChoice.ListCount = 0 Then
            DisplayMsg "请选择所要导入的字段", vbExclamation
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(2)
            lstChoice.SetFocus
        End If
    ElseIf iCurrentTab > 3 And adoServConf.Recordset.RecordCount = 0 Then
        DisplayMsg "没有建立相应的映射关系", vbExclamation
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(2)
        lstChoice.SetFocus
    End If
    iCurrentTab = tbsOptions.SelectedItem.Index

    '按纽控制
    If iCurrentTab = 1 Then
        cmdPrevious.Enabled = False
        cmdNext.Enabled = True
        cmdSave.Enabled = False
    ElseIf iCurrentTab = tbsOptions.Tabs.Count Then
        cmdPrevious.Enabled = True
        cmdNext.Enabled = False
        cmdSave.Enabled = True
    Else
        cmdPrevious.Enabled = True
        cmdNext.Enabled = True
        cmdSave.Enabled = False
    End If
    
    '业务处理
    Select Case iCurrentTab
    Case 2
        Call cbTable_Validate(False)
    Case 3
        Call Step3
    Case 4
        Call Step3
        Call Step4
    End Select
End Sub

Private Sub txtFileName_Validate(Cancel As Boolean)
    If m_strFileName = txtFileName.Text Then
        Exit Sub
    End If

    If Dir(txtFileName.Text) = "" Then
        DisplayMsg "文件[" & txtFileName.Text & "]不存在，请重新选择。", vbExclamation
        Cancel = True
        Exit Sub
    End If
    m_strFileName = txtFileName.Text
    
    '读取表名
    Dim cnExcel As ADODB.Connection
    Dim rsTable As ADODB.Recordset
    Dim strTableName As String
    Set cnExcel = New ADODB.Connection
    With cnExcel
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & m_strFileName & _
            ";Extended Properties=Excel 8.0;"
        .CursorLocation = adUseClient
        .Open
    End With
    Set rsTable = cnExcel.OpenSchema(adSchemaTables)
    cbTable.Clear
    While rsTable.EOF <> True
        strTableName = rsTable.Fields("TABLE_NAME").value
        If Right(strTableName, 1) = "$" Or Right(strTableName, 2) = "$'" Then
            cbTable.AddItem strTableName
        End If
        rsTable.MoveNext
    Wend
    rsTable.Close
    cnExcel.Close
    Set rsTable = Nothing
    Set cnExcel = Nothing
End Sub

Private Sub Step3()
    If m_blChoiceUpdate = False Then
        Exit Sub
    End If

    '清理数据库
    Dim strSQL As String

    gAdoConnDB.BeginTrans
    strSQL = "DELETE FROM serv_field WHERE serv_id='-1'"
    gAdoConnDB.Execute strSQL
    gAdoConnDB.CommitTrans

    Dim i As Integer
    gAdoConnDB.BeginTrans
    For i = 0 To lstChoice.ListCount - 1
        strSQL = "insert into serv_field(serv_id,field_name,display_name,valid_flag)"
        If i = 0 Then
            strSQL = strSQL & " values('" & m_strServId & "','emp_id','" & lstChoice.List(i) & "',1)"
        ElseIf i = 1 Then
            strSQL = strSQL & " values('" & m_strServId & "','emp_name','" & lstChoice.List(i) & "',1)"
        ElseIf i > SERV_FIELD_CNT_MAX Then
            DisplayMsg "字段过多，请扩展系统，谢谢。", vbExclamation
            Exit For
        Else
            strSQL = strSQL & " values('" & m_strServId & "','field" & Format(i - 1, SERV_FIELD_PATTERN) & "','" & lstChoice.List(i) & "',1)"
        End If
        gAdoConnDB.Execute strSQL
    Next
    gAdoConnDB.CommitTrans
    adoServConf.Refresh
    adoSalary.Refresh
    DisplayGrid grdRelation, "serv_field"
    DisplayGrid grdBrower, "salary"
    m_blChoiceUpdate = False
End Sub

Private Sub Step4()
    txtServSheet.Text = Replace(m_strTableName, "'", "")
    txtServTemplFn.Text = m_strFileName
    adoServType.Refresh
End Sub
