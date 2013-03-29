Attribute VB_Name = "mdBackup"
Option Explicit

Private m_strErrMsg As String

Public Function gBackup(ByVal pmStrFileName As String, pmStrRetMsg As String, Optional ByVal pmStrTableName As String = "ALL") As Boolean
    gBackup = False
    
    pmStrRetMsg = ""
    m_strErrMsg = ""
    
    Screen.MousePointer = vbHourglass
    If pmStrTableName = "ALL" Then      '完整模式
        If gBackupFile(pmStrFileName, pmStrRetMsg) = False Then
            GoTo NormalExit
        End If
    Else                                '表级模式
        If gBackupTable(pmStrFileName, pmStrRetMsg, pmStrTableName) = False Then
            GoTo NormalExit
        End If
    End If
    
    '生成验证文件
    If GenMd5File(pmStrFileName, pmStrTableName) = False Then
        pmStrRetMsg = "生成验证文件时出错：" & m_strErrMsg
        GoTo NormalExit
    End If
    Screen.MousePointer = vbDefault
    gBackup = True
    Exit Function

NormalExit:
    Screen.MousePointer = vbDefault
End Function

Public Function gBackupFile(ByVal pmStrFileName As String, pmStrRetMsg As String) As Boolean
    On Error GoTo ErrHandle
    gBackupFile = False
    
    pmStrRetMsg = ""
    m_strErrMsg = ""
    
    Dim fsObj As FileSystemObject
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    
    '关闭数据库
    Call CloseDB
    '拷贝数据文件
    fsObj.CopyFile gStrDBFileName, pmStrFileName, True
    '重连数据库
    Call ConnectDB
    
    gBackupFile = True
    Exit Function

ErrHandle:
    pmStrRetMsg = "错误代码：" & Err.Number & vbCrLf & _
                "错误内容：" & Err.Description
    
NormalExit:
    On Error Resume Next
    Call ConnectDB
End Function

Public Function gBackupTable(ByVal pmStrFileName As String, pmStrRetMsg As String, ByVal pmStrTableName As String) As Boolean
    On Error GoTo ErrHandle
    gBackupTable = False
    
    pmStrRetMsg = ""
    m_strErrMsg = ""
    
    Dim fsObj As FileSystemObject
    Dim strSql As String
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    
    '生成临时.xls文件
    Dim strTempFileName As String
    strTempFileName = fsObj.GetParentFolderName(pmStrFileName) & "\" & fsObj.GetBaseName(pmStrFileName) & ".xls"
    If Dir(strTempFileName) <> "" Then
        fsObj.DeleteFile strTempFileName
    End If

    strSql = "select * into [Excel 8.0;database=" & strTempFileName & "].[" & pmStrTableName & "] from " & pmStrTableName
    If gExecSql(strSql) = False Then
        pmStrRetMsg = "运行SQL语句 '" & strSql & "' 时出错。"
        Exit Function
    End If
    
    If pmStrFileName <> strTempFileName Then
        '删除原有文件
        If Dir(pmStrFileName) <> "" Then
            fsObj.DeleteFile pmStrFileName
        End If
        'MOVE成正式文件
        fsObj.MoveFile strTempFileName, pmStrFileName
    End If

    gBackupTable = True
    Exit Function

ErrHandle:
    pmStrRetMsg = "错误代码：" & Err.Number & vbCrLf & _
                "错误内容：" & Err.Description
End Function

Public Function gRecover(ByVal pmStrFileName As String, pmStrRetMsg As String, Optional ByVal pmStrTableName As String = "ALL") As Boolean
    gRecover = False

    m_strErrMsg = ""
    pmStrRetMsg = ""
    Screen.MousePointer = vbHourglass
    '生成验证文件
    If ValidMd5File(pmStrFileName, pmStrTableName) = False Then
        pmStrRetMsg = "验证文件不通过：" & m_strErrMsg
        GoTo NormalExit
    End If

    Dim fsObj As FileSystemObject
    Dim strSql As String
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    
    If pmStrTableName = "ALL" Then      '完整模式
        If gRecoverFile(pmStrFileName, pmStrRetMsg) = False Then
            GoTo NormalExit
        End If
    Else
        If gRecoverTable(pmStrFileName, pmStrRetMsg, pmStrTableName) = False Then
            GoTo NormalExit
        End If
    End If

    Screen.MousePointer = vbDefault
    gRecover = True
    Exit Function

NormalExit:
    Screen.MousePointer = vbDefault
End Function

Public Function gRecoverFile(ByVal pmStrFileName As String, pmStrRetMsg As String) As Boolean
    On Error GoTo ErrHandle
    gRecoverFile = False

    m_strErrMsg = ""
    pmStrRetMsg = ""
    
    Dim fsObj As FileSystemObject
    Set fsObj = CreateObject("Scripting.FileSystemObject")

    '关闭数据库
    Call CloseDB
    '拷贝数据文件
    fsObj.CopyFile pmStrFileName, gStrDBFileName, True
    '重连数据库
    Call ConnectDB

    gRecoverFile = True
    Exit Function

ErrHandle:
    pmStrRetMsg = "错误代码：" & Err.Number & vbCrLf & _
                "错误内容：" & Err.Description

NormalExit:
    On Error Resume Next
    Call ConnectDB
End Function

Public Function gRecoverTable(ByVal pmStrFileName As String, pmStrRetMsg As String, ByVal pmStrTableName As String) As Boolean
    On Error GoTo ErrHandle
    gRecoverTable = False

    m_strErrMsg = ""
    pmStrRetMsg = ""

    Dim fsObj As FileSystemObject
    Dim strSql As String
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    
    '生成临时.xls文件
    Dim strTempFileName As String
    strTempFileName = fsObj.GetParentFolderName(pmStrFileName) & "\" & fsObj.GetBaseName(pmStrFileName) & ".xls"
    If pmStrFileName <> strTempFileName Then
        fsObj.CopyFile pmStrFileName, strTempFileName, True
    End If

    '拷贝表数据
    gAdoConnDB.BeginTrans
    strSql = "delete from " & pmStrTableName
    gAdoConnDB.Execute strSql
    strSql = "insert into " & pmStrTableName & " select * From [Excel 8.0;database=" & strTempFileName & "].[" & pmStrTableName & "$]"
    gAdoConnDB.Execute strSql
    gAdoConnDB.CommitTrans
    
    If pmStrFileName <> strTempFileName Then
        '删除原有文件
        fsObj.DeleteFile strTempFileName
    End If
    
    gRecoverTable = True
    Exit Function

ErrHandle:
    pmStrRetMsg = "错误代码：" & Err.Number & vbCrLf & _
                "错误内容：" & Err.Description

NormalExit:
    On Error Resume Next
    gAdoConnDB.RollbackTrans
End Function

Private Function GenMd5File(ByVal pmStrFileName As String, Optional ByVal pmStrTableName As String = "ALL") As Boolean
    On Error GoTo ErrHandle
    GenMd5File = False

    Dim clsMD5 As CMD5
    Dim strMd5 As String
    Dim strMd5FileName As String
    Dim fsObj As FileSystemObject
    Dim ts As TextStream

    Set fsObj = CreateObject("Scripting.FileSystemObject")
    Set clsMD5 = New CMD5
    strMd5FileName = fsObj.GetParentFolderName(pmStrFileName) & "\" & fsObj.GetBaseName(pmStrFileName) & ".md5"

    Set ts = fsObj.OpenTextFile(strMd5FileName, ForWriting, True)
    ts.WriteLine (pmStrTableName)
    strMd5 = clsMD5.DigestStrToHexStr("TABLE:" & pmStrTableName)
    ts.WriteLine (strMd5)
    strMd5 = clsMD5.DigestFileToHexStr(pmStrFileName)
    ts.WriteLine (strMd5)
    ts.Close
    GenMd5File = True
    Exit Function

ErrHandle:
    m_strErrMsg = "错误代码：" & Err.Number & vbCrLf & _
                "错误内容：" & Err.Description
End Function

Private Function ValidMd5File(ByVal pmStrFileName As String, Optional ByVal pmStrTableName As String) As Boolean
    On Error GoTo ErrHandle
    ValidMd5File = False

    Dim clsMD5 As CMD5
    Dim strMd5 As String
    Dim strMd5FileName As String
    Dim strLine(3) As String
    Dim fsObj As FileSystemObject
    Dim ts As TextStream
    Dim i As Integer

    Set fsObj = CreateObject("Scripting.FileSystemObject")
    Set clsMD5 = New CMD5
    strMd5FileName = fsObj.GetParentFolderName(pmStrFileName) & "\" & fsObj.GetBaseName(pmStrFileName) & ".md5"
    If fsObj.FileExists(strMd5FileName) = False Then
        m_strErrMsg = "验证文件 '" & strMd5FileName & "' 不存在。"
        Exit Function
    End If
    Set ts = fsObj.OpenTextFile(strMd5FileName, ForReading, False)
    For i = 0 To 2
        If ts.AtEndOfStream = True Then
            m_strErrMsg = "验证文件 '" & strMd5FileName & "' 信息不全或被破坏。"
            Exit Function
        End If
        strLine(i) = ts.ReadLine
    Next
    ts.Close
    strMd5 = clsMD5.DigestStrToHexStr("TABLE:" & strLine(0))
    If strMd5 <> strLine(1) Then
        m_strErrMsg = "验证 '" & strLine(0) & "' 不通过，验证文件内容可能被修改。"
        Exit Function
    End If

    If strLine(0) <> pmStrTableName Then
        m_strErrMsg = "验证 '" & strLine(0) & "' 不通过，此数据文件不是相应的内容(" & pmStrTableName & ")。"
        Exit Function
    End If

    strMd5 = clsMD5.DigestFileToHexStr(pmStrFileName)
    If strMd5 <> strLine(2) Then
        m_strErrMsg = "验证 '" & pmStrFileName & "' 不通过，请确认此验证文件是否与之相关。"
        Exit Function
    End If

    ValidMd5File = True
    Exit Function

ErrHandle:
    m_strErrMsg = "错误代码：" & Err.Number & vbCrLf & _
                "错误内容：" & Err.Description
End Function
