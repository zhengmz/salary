VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetConfig(ByVal pmStrSection As String, ByVal pmStrKeyName As String, _
                          Optional ByVal pmStrDefault As String = "") As String
    GetConfig = GetSetting(gStrAppName, pmStrSection, pmStrKeyName, pmStrDefault)
End Function

Public Sub SetConfig(ByVal pmStrSection As String, ByVal pmStrKeyName As String, _
                          Optional ByVal pmStrDefault As String = "")
    SaveSetting gStrAppName, pmStrSection, pmStrKeyName, pmStrDefault
End Sub

Public Sub DelConfig(Optional ByVal pmStrSection As String = "", Optional ByVal pmStrKeyName As String = "")
    If pmStrSection <> "" And pmStrKeyName <> "" Then
        DeleteSetting gStrAppName, pmStrSection, pmStrKeyName
    End If
    If pmStrSection <> "" And pmStrKeyName = "" Then
        DeleteSetting gStrAppName, pmStrSection
    End If
    If pmStrSection = "" Then
        DeleteSetting gStrAppName
    End If
End Sub

Private Sub Class_Initialize()
    '判断是否第一次运行
    If GetSetting(gStrAppName, "Common", "Startup", 0) = 1 Then
        Exit Sub
    End If
    '判断数据库是否连接
    If gAdoConnDB.State <> adStateOpen Then
        Exit Sub
    End If

    '设置默认参数到注册表中
    Dim rsConfig As New ADODB.Recordset
    Dim strSql As String
    strSql = "select dict_key,dict_value from dicts where dict_sect='OPT_EMAIL'"

    rsConfig.Open strSql, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    While rsConfig.EOF <> True
        SaveSetting gStrAppName, "Options", rsConfig("dict_key"), rsConfig("dict_value")
        rsConfig.MoveNext
    Wend
    rsConfig.Close
    Set rsConfig = Nothing

    '设置Startup
    SaveSetting gStrAppName, "Common", "Startup", 1
End Sub
