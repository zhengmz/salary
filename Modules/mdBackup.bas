Attribute VB_Name = "mdBackup"
Option Explicit

Private m_strErrMsg As String

Public Function gBackup(ByVal pmStrFileName As String, pmStrRetMsg As String, Optional ByVal pmStrTableName As String = "ALL") As Boolean
    gBackup = False
    
    pmStrRetMsg = ""
    m_strErrMsg = ""
    
    Screen.MousePointer = vbHourglass
    If pmStrTableName = "ALL" Then      '����ģʽ
        If gBackupFile(pmStrFileName, pmStrRetMsg) = False Then
            GoTo NormalExit
        End If
    Else                                '��ģʽ
        If gBackupTable(pmStrFileName, pmStrRetMsg, pmStrTableName) = False Then
            GoTo NormalExit
        End If
    End If
    
    '������֤�ļ�
    If GenMd5File(pmStrFileName, pmStrTableName) = False Then
        pmStrRetMsg = "������֤�ļ�ʱ����" & m_strErrMsg
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
    
    '�ر����ݿ�
    Call CloseDB
    '���������ļ�
    fsObj.CopyFile gStrDBFileName, pmStrFileName, True
    '�������ݿ�
    Call ConnectDB
    
    gBackupFile = True
    Exit Function

ErrHandle:
    pmStrRetMsg = "������룺" & Err.Number & vbCrLf & _
                "�������ݣ�" & Err.Description
    
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
    
    '������ʱ.xls�ļ�
    Dim strTempFileName As String
    strTempFileName = fsObj.GetParentFolderName(pmStrFileName) & "\" & fsObj.GetBaseName(pmStrFileName) & ".xls"
    If Dir(strTempFileName) <> "" Then
        fsObj.DeleteFile strTempFileName
    End If

    strSql = "select * into [Excel 8.0;database=" & strTempFileName & "].[" & pmStrTableName & "] from " & pmStrTableName
    If gExecSql(strSql) = False Then
        pmStrRetMsg = "����SQL��� '" & strSql & "' ʱ����"
        Exit Function
    End If
    
    If pmStrFileName <> strTempFileName Then
        'ɾ��ԭ���ļ�
        If Dir(pmStrFileName) <> "" Then
            fsObj.DeleteFile pmStrFileName
        End If
        'MOVE����ʽ�ļ�
        fsObj.MoveFile strTempFileName, pmStrFileName
    End If

    gBackupTable = True
    Exit Function

ErrHandle:
    pmStrRetMsg = "������룺" & Err.Number & vbCrLf & _
                "�������ݣ�" & Err.Description
End Function

Public Function gRecover(ByVal pmStrFileName As String, pmStrRetMsg As String, Optional ByVal pmStrTableName As String = "ALL") As Boolean
    gRecover = False

    m_strErrMsg = ""
    pmStrRetMsg = ""
    Screen.MousePointer = vbHourglass
    '������֤�ļ�
    If ValidMd5File(pmStrFileName, pmStrTableName) = False Then
        pmStrRetMsg = "��֤�ļ���ͨ����" & m_strErrMsg
        GoTo NormalExit
    End If

    Dim fsObj As FileSystemObject
    Dim strSql As String
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    
    If pmStrTableName = "ALL" Then      '����ģʽ
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

    '�ر����ݿ�
    Call CloseDB
    '���������ļ�
    fsObj.CopyFile pmStrFileName, gStrDBFileName, True
    '�������ݿ�
    Call ConnectDB

    gRecoverFile = True
    Exit Function

ErrHandle:
    pmStrRetMsg = "������룺" & Err.Number & vbCrLf & _
                "�������ݣ�" & Err.Description

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
    
    '������ʱ.xls�ļ�
    Dim strTempFileName As String
    strTempFileName = fsObj.GetParentFolderName(pmStrFileName) & "\" & fsObj.GetBaseName(pmStrFileName) & ".xls"
    If pmStrFileName <> strTempFileName Then
        fsObj.CopyFile pmStrFileName, strTempFileName, True
    End If

    '����������
    gAdoConnDB.BeginTrans
    strSql = "delete from " & pmStrTableName
    gAdoConnDB.Execute strSql
    strSql = "insert into " & pmStrTableName & " select * From [Excel 8.0;database=" & strTempFileName & "].[" & pmStrTableName & "$]"
    gAdoConnDB.Execute strSql
    gAdoConnDB.CommitTrans
    
    If pmStrFileName <> strTempFileName Then
        'ɾ��ԭ���ļ�
        fsObj.DeleteFile strTempFileName
    End If
    
    gRecoverTable = True
    Exit Function

ErrHandle:
    pmStrRetMsg = "������룺" & Err.Number & vbCrLf & _
                "�������ݣ�" & Err.Description

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
    m_strErrMsg = "������룺" & Err.Number & vbCrLf & _
                "�������ݣ�" & Err.Description
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
        m_strErrMsg = "��֤�ļ� '" & strMd5FileName & "' �����ڡ�"
        Exit Function
    End If
    Set ts = fsObj.OpenTextFile(strMd5FileName, ForReading, False)
    For i = 0 To 2
        If ts.AtEndOfStream = True Then
            m_strErrMsg = "��֤�ļ� '" & strMd5FileName & "' ��Ϣ��ȫ���ƻ���"
            Exit Function
        End If
        strLine(i) = ts.ReadLine
    Next
    ts.Close
    strMd5 = clsMD5.DigestStrToHexStr("TABLE:" & strLine(0))
    If strMd5 <> strLine(1) Then
        m_strErrMsg = "��֤ '" & strLine(0) & "' ��ͨ������֤�ļ����ݿ��ܱ��޸ġ�"
        Exit Function
    End If

    If strLine(0) <> pmStrTableName Then
        m_strErrMsg = "��֤ '" & strLine(0) & "' ��ͨ�����������ļ�������Ӧ������(" & pmStrTableName & ")��"
        Exit Function
    End If

    strMd5 = clsMD5.DigestFileToHexStr(pmStrFileName)
    If strMd5 <> strLine(2) Then
        m_strErrMsg = "��֤ '" & pmStrFileName & "' ��ͨ������ȷ�ϴ���֤�ļ��Ƿ���֮��ء�"
        Exit Function
    End If

    ValidMd5File = True
    Exit Function

ErrHandle:
    m_strErrMsg = "������룺" & Err.Number & vbCrLf & _
                "�������ݣ�" & Err.Description
End Function
