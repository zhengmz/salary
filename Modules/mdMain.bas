Attribute VB_Name = "mdMain"
Option Explicit

'数据库连接串
Public gStrConnDB As String
'数据库文件名，含绝对路径
Public gStrDBFileName As String
'数据库连接
Public gAdoConnDB As New ADODB.Connection
'应用名称
Public gStrAppName As String

'系统变量
'系统日志级别，默认为0,即不记录
Public gSysLogLevel As Integer
'系统日志目录，默认为App.path\Log
Public gSysLogDir As String
'分表字段数，默认为0，表示不分表
Public gSysSplitFields As Integer


Sub Main()
    '初始化
    gStrDBFileName = App.Path & "\Data\salary.dat"
    gStrConnDB = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gStrDBFileName & ";Persist Security Info=False;Jet OLEDB:Database password=mdb@salary"
    gStrAppName = App.EXEName
    gSysLogLevel = 0
    gSysLogDir = App.Path & "\Log"
    gSysSplitFields = 0
    
    '连接数据库
    If Not ConnectDB() Then
        DisplayMsg "连接数据库错误，请检查数据库文件！", vbCritical
        End
    End If
    
    '检查并创建目录结构
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(App.Path & "\Backup") = False Then
        fso.CreateFolder App.Path & "\Backup"
    End If
    If fso.FolderExists(App.Path & "\Template") = False Then
        fso.CreateFolder App.Path & "\Template"
    End If
    If fso.FolderExists(gSysLogDir) = False Then
        fso.CreateFolder gSysLogDir
    End If
    Set fso = Nothing
    
    '获取系统参数
    Call gGetParameter

    '修改系统使用次数
    Call gUpdateUsedCount

    MDIMain.Show
End Sub
