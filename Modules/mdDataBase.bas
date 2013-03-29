Attribute VB_Name = "mdDataBase"
Option Explicit

'连接到数据库
Function ConnectDB() As Boolean
    On Error GoTo Err_Process

    If gAdoConnDB.State = adStateClosed Then
        gAdoConnDB.Open gStrConnDB
    End If
    ConnectDB = True
    Exit Function
      
Err_Process:
    DisplayMsg "连接数据库出错！", vbCritical
    ConnectDB = False
End Function

'关闭数据库
Function CloseDB() As Boolean
    On Error Resume Next
    gAdoConnDB.Close
    Set gAdoConnDB = Nothing
End Function

Public Function gCompareDB(pmStrRetMsg As String) As Boolean
    On Error GoTo ErrHandle
    gCompareDB = False
    
    pmStrRetMsg = "数据库压缩信息如下：" & vbCrLf & _
                "原始数据库文件的大小为：" & FileLen(gStrDBFileName) & "B"
    
    Dim fsObj As FileSystemObject
    Dim je As JRO.JetEngine
    Dim strTempFileName As String

    '关闭数据库
    Call CloseDB
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    Set je = New JRO.JetEngine
    
    '删除原临时.xls文件
    strTempFileName = gStrDBFileName & ".tmp"
    If Dir(strTempFileName) <> "" Then
        fsObj.DeleteFile strTempFileName
    End If

    '压缩文件
    je.CompactDatabase _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & gStrDBFileName & ";Jet OLEDB:Database password=mdb@salary", _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & strTempFileName

    'MOVE成正式文件数据文件
    fsObj.CopyFile strTempFileName, gStrDBFileName, True
    fsObj.DeleteFile strTempFileName
    
    '重连数据库
    Call ConnectDB
    
    pmStrRetMsg = pmStrRetMsg & vbCrLf & _
                "压缩后数据库文件的大小为：" & FileLen(gStrDBFileName) & "B"
    
    gCompareDB = True
    Exit Function

ErrHandle:
    pmStrRetMsg = "错误代码：" & Err.Number & vbCrLf & _
                "错误内容：" & Err.Description
    
NormalExit:
    On Error Resume Next
    Call ConnectDB
End Function

Public Function gExecSql(ByVal pmStrSQL As String) As Boolean
    On Error GoTo ErrHandle
    gAdoConnDB.BeginTrans
    gAdoConnDB.Execute pmStrSQL
    gAdoConnDB.CommitTrans
    gExecSql = True
    Exit Function

ErrHandle:
    gExecSql = False
    gAdoConnDB.RollbackTrans
    DisplayMsg "执行语句[" & pmStrSQL & "]时发生错误", vbCritical
End Function

'导入程序 gImportData
'描述说明：
'   使用Ado访问技术，从外部的数据源（如Excel）中逐行导入数据，并判断重复信息。
'输入说明
'   pmStrSourceRS: 外部数据源
'   pmStrTargetRS:  目标数据源
'   pmIntFieldCount:  所要导入的字段数
'输出说明
'   函数返回值: 成功为True, 失败为False
'   pmRetMsg:  导入日志信息
Function gImportData(ByVal pmStrSourceRS As String, _
                    ByVal pmStrTargetRS As String, _
                    ByVal pmIntFieldCount As Integer, _
                    pmRetMsg As String) As Boolean
'程序开始
    On Error GoTo Err_Process
    
    WriteLog "mdDataBase:gImportData", "外部数据源：" & pmStrSourceRS, LOG_LEVEL_DEBUG
    WriteLog "mdDataBase:gImportData", "目标数据源：" & pmStrTargetRS, LOG_LEVEL_DEBUG

    Dim iLoadCount As Integer  '导入文件行数
    Dim iSuccInst As Integer    '成功导入记录数
    Dim iRepeat As Integer      '重复记录数
    Dim strRepeat As String     '重复行集
    Dim iRelate As Integer      '需要表关联的记录数
    Dim strRelate As String     '需要表关联的行集

    '数据操作
    Dim rsTarget As New ADODB.Recordset
    Dim rsSource As New ADODB.Recordset
    
    '开始事务
    gImportData = False
    gAdoConnDB.BeginTrans
    rsSource.Open pmStrSourceRS, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    'rsTarget.Open pmStrTargetRS, gAdoConnDB, adOpenKeyset, adLockOptimistic, adCmdText
    rsTarget.Open pmStrTargetRS, gStrConnDB, adOpenKeyset, adLockOptimistic, adCmdText

    iLoadCount = 0
    iSuccInst = 0
    iRepeat = 0
    strRepeat = "因重复无法导入的行号有: "
    iRelate = 0
    strRelate = "需要关联，而无法导入的行号有: "

    Dim i As Integer
    Dim j As Integer

    Do Until rsSource.EOF
        iLoadCount = iLoadCount + 1
        'Call WriteLog("gImportData", "记录信息：" & rsSource.GetString(adClipString, 1, ",", ";"), LOG_LEVEL_DEBUG)
        Call WriteLog("gImportData", "记录行号：" & iLoadCount, LOG_LEVEL_DEBUG)
        rsTarget.AddNew
        For i = 0 To pmIntFieldCount - 1
            Call WriteLog("gImportData", "字段对应：" & rsSource(i).Name & "(" & rsSource(i) & ")-" & rsTarget(i).Name, LOG_LEVEL_DEBUG)
            If IsNull(rsSource(i)) Then
                If rsTarget(i).Type = adDate Then
                    rsTarget(i) = Date
                End If
            ElseIf rsSource(i) <> "0" Then
                rsTarget(i) = rsSource(i)
            End If
        Next
        rsTarget.Update
        rsSource.MoveNext
    Loop
    gAdoConnDB.CommitTrans

    '生成日志信息
    iSuccInst = iLoadCount - iRepeat - iRelate
    pmRetMsg = vbCrLf & _
            "导入文件的总行数为： " & iLoadCount & vbCrLf & _
            "成功导入的行数为： " & iSuccInst
    If iRepeat > 0 Then
        pmRetMsg = pmRetMsg & vbCrLf & _
            "因重复无法导入的行数有： " & iRepeat & vbCrLf & _
            strRepeat
    End If
    If iRelate > 0 Then
        pmRetMsg = pmRetMsg & vbCrLf & _
            "需要关联，而无法导入的行数有： " & iRelate & vbCrLf & _
            strRelate
    End If
    pmRetMsg = pmRetMsg & vbCrLf & "导入成功！"
    DisplayMsg "数据导入成功！", vbInformation
    gImportData = True
    On Error GoTo 0

Normal_Finish:
    On Error Resume Next
    If Not IsNull(rsTarget) And rsTarget.State = adStateOpen Then rsTarget.Close
    If Not IsNull(rsSource) And rsSource.State = adStateOpen Then rsSource.Close

    Set rsTarget = Nothing
    Set rsSource = Nothing
    
    Exit Function

Err_Process:
    If Err.Number = -2147217887 Then
        '违背唯一索引的约束
        If gAdoConnDB.Errors(0).NativeError = -105121349 Then
        'sqlstate=3022
            rsTarget.CancelUpdate
            iRepeat = iRepeat + 1
            strRepeat = strRepeat & iLoadCount & ","
            Resume Next
        End If
        '违背外键关联系约束
        If gAdoConnDB.Errors(0).NativeError = -535037517 Then
        'sqlstate=3201
            rsTarget.CancelUpdate
            iRelate = iRelate + 1
            strRelate = strRelate & iLoadCount & ","
            Resume Next
        End If
    End If

    DisplayMsg "数据导入失败！" & vbCrLf & vbCrLf & "请检查导入数据源 '" & pmStrSourceRS & "' 是否符合格式，" _
        & vbCrLf & vbCrLf & "如果不确定，请询问维护人员！", vbCritical

    pmRetMsg = "导入失败，请检查文件是否符合格式！"
    Select Case Err.Number
    Case 3265
        pmRetMsg = pmRetMsg & vbCrLf & "错误原因：数据文件无此字段“" & rsSource(i).Name & "”。"
    Case -2147217904
        pmRetMsg = pmRetMsg & vbCrLf & "错误原因：数据文件至少有一个以上的字段找不到。"
    Case -2147467259
        pmRetMsg = pmRetMsg & vbCrLf & "错误原因：数据文件找不到相应的表格。"
    Case -2147217887
        pmRetMsg = pmRetMsg & vbCrLf & "错误原因：数据文件中第 " & iLoadCount & " 数据行的关键信息如员工编码和姓名等读取出错。" & vbCrLf & _
                        vbTab & "  读到的数据是" & rsSource.GetString(adClipString, , ",")
    End Select

    pmRetMsg = pmRetMsg & vbCrLf & vbCrLf & "传入的参数信息如下：" & vbCrLf & _
                "外部数据源：" & pmStrSourceRS & vbCrLf & _
                "目标数据源：" & pmStrTargetRS & vbCrLf & _
                "字段数：" & pmIntFieldCount & vbCrLf & _
                "正在处理的行号：" & iLoadCount & vbCrLf & vbCrLf & _
                "报告错误的详细信息如下：" & vbCrLf & _
                "错误号# " & Err.Number & vbCrLf & _
                "错误内容: " & Err.Description
    gAdoConnDB.RollbackTrans
    GoTo Normal_Finish
End Function

Public Sub gGetParameter()
    Dim rsDicts As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select dict_key, dict_value from dicts where dict_sect='SYS'"
    rsDicts.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText

    Do Until rsDicts.EOF
        Select Case rsDicts("dict_key")
        Case "LOG_LEVEL"
            gSysLogLevel = CInt(rsDicts("dict_value"))
        Case "SPLIT_FIELDS"
            gSysSplitFields = CInt(rsDicts("dict_value"))
        Case Else
        End Select
        rsDicts.MoveNext
    Loop
    rsDicts.Close
    Set rsDicts = Nothing
End Sub

Public Sub gUpdateUsedCount()
    '修改系统使用次数
    Dim rsDicts As New ADODB.Recordset
    Dim strSQL As String

    strSQL = "select dict_type, dict_value from dicts where dict_sect='SYS' and dict_key='USED_COUNT'"
    rsDicts.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText

    If rsDicts.RecordCount > 0 Then
        Dim strArrCnt() As String
        strArrCnt = Split(rsDicts("dict_value"), "-", 2, vbTextCompare)
        If UBound(strArrCnt) = 0 Then
            strSQL = "update dicts set dict_value='0-" & CLng(strArrCnt(0)) + 1 & "' where dict_sect='SYS' and dict_key='USED_COUNT'"
        Else
            strSQL = "update dicts set dict_value='" & strArrCnt(0) & "-" & CLng(strArrCnt(1)) + 1 & "' where dict_sect='SYS' and dict_key='USED_COUNT'"
        End If
        gExecSql (strSQL)
    End If
    rsDicts.Close
    Set rsDicts = Nothing
End Sub
