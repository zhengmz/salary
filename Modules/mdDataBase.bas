Attribute VB_Name = "mdDataBase"
Option Explicit

'���ӵ����ݿ�
Function ConnectDB() As Boolean
    On Error GoTo Err_Process

    If gAdoConnDB.State = adStateClosed Then
        gAdoConnDB.Open gStrConnDB
    End If
    ConnectDB = True
    Exit Function
      
Err_Process:
    DisplayMsg "�������ݿ����", vbCritical
    ConnectDB = False
End Function

'�ر����ݿ�
Function CloseDB() As Boolean
    On Error Resume Next
    gAdoConnDB.Close
    Set gAdoConnDB = Nothing
End Function

Public Function gCompareDB(pmStrRetMsg As String) As Boolean
    On Error GoTo ErrHandle
    gCompareDB = False
    
    pmStrRetMsg = "���ݿ�ѹ����Ϣ���£�" & vbCrLf & _
                "ԭʼ���ݿ��ļ��Ĵ�СΪ��" & FileLen(gStrDBFileName) & "B"
    
    Dim fsObj As FileSystemObject
    Dim je As JRO.JetEngine
    Dim strTempFileName As String

    '�ر����ݿ�
    Call CloseDB
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    Set je = New JRO.JetEngine
    
    'ɾ��ԭ��ʱ.xls�ļ�
    strTempFileName = gStrDBFileName & ".tmp"
    If Dir(strTempFileName) <> "" Then
        fsObj.DeleteFile strTempFileName
    End If

    'ѹ���ļ�
    je.CompactDatabase _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & gStrDBFileName & ";Jet OLEDB:Database password=mdb@salary", _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & strTempFileName

    'MOVE����ʽ�ļ������ļ�
    fsObj.CopyFile strTempFileName, gStrDBFileName, True
    fsObj.DeleteFile strTempFileName
    
    '�������ݿ�
    Call ConnectDB
    
    pmStrRetMsg = pmStrRetMsg & vbCrLf & _
                "ѹ�������ݿ��ļ��Ĵ�СΪ��" & FileLen(gStrDBFileName) & "B"
    
    gCompareDB = True
    Exit Function

ErrHandle:
    pmStrRetMsg = "������룺" & Err.Number & vbCrLf & _
                "�������ݣ�" & Err.Description
    
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
    DisplayMsg "ִ�����[" & pmStrSQL & "]ʱ��������", vbCritical
End Function

'������� gImportData
'����˵����
'   ʹ��Ado���ʼ��������ⲿ������Դ����Excel�������е������ݣ����ж��ظ���Ϣ��
'����˵��
'   pmStrSourceRS: �ⲿ����Դ
'   pmStrTargetRS:  Ŀ������Դ
'   pmIntFieldCount:  ��Ҫ������ֶ���
'���˵��
'   ��������ֵ: �ɹ�ΪTrue, ʧ��ΪFalse
'   pmRetMsg:  ������־��Ϣ
Function gImportData(ByVal pmStrSourceRS As String, _
                    ByVal pmStrTargetRS As String, _
                    ByVal pmIntFieldCount As Integer, _
                    pmRetMsg As String) As Boolean
'����ʼ
    On Error GoTo Err_Process
    
    WriteLog "mdDataBase:gImportData", "�ⲿ����Դ��" & pmStrSourceRS, LOG_LEVEL_DEBUG
    WriteLog "mdDataBase:gImportData", "Ŀ������Դ��" & pmStrTargetRS, LOG_LEVEL_DEBUG

    Dim iLoadCount As Integer  '�����ļ�����
    Dim iSuccInst As Integer    '�ɹ������¼��
    Dim iRepeat As Integer      '�ظ���¼��
    Dim strRepeat As String     '�ظ��м�
    Dim iRelate As Integer      '��Ҫ������ļ�¼��
    Dim strRelate As String     '��Ҫ��������м�

    '���ݲ���
    Dim rsTarget As New ADODB.Recordset
    Dim rsSource As New ADODB.Recordset
    
    '��ʼ����
    gImportData = False
    gAdoConnDB.BeginTrans
    rsSource.Open pmStrSourceRS, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    'rsTarget.Open pmStrTargetRS, gAdoConnDB, adOpenKeyset, adLockOptimistic, adCmdText
    rsTarget.Open pmStrTargetRS, gStrConnDB, adOpenKeyset, adLockOptimistic, adCmdText

    iLoadCount = 0
    iSuccInst = 0
    iRepeat = 0
    strRepeat = "���ظ��޷�������к���: "
    iRelate = 0
    strRelate = "��Ҫ���������޷�������к���: "

    Dim i As Integer
    Dim j As Integer

    Do Until rsSource.EOF
        iLoadCount = iLoadCount + 1
        'Call WriteLog("gImportData", "��¼��Ϣ��" & rsSource.GetString(adClipString, 1, ",", ";"), LOG_LEVEL_DEBUG)
        Call WriteLog("gImportData", "��¼�кţ�" & iLoadCount, LOG_LEVEL_DEBUG)
        rsTarget.AddNew
        For i = 0 To pmIntFieldCount - 1
            Call WriteLog("gImportData", "�ֶζ�Ӧ��" & rsSource(i).Name & "(" & rsSource(i) & ")-" & rsTarget(i).Name, LOG_LEVEL_DEBUG)
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

    '������־��Ϣ
    iSuccInst = iLoadCount - iRepeat - iRelate
    pmRetMsg = vbCrLf & _
            "�����ļ���������Ϊ�� " & iLoadCount & vbCrLf & _
            "�ɹ����������Ϊ�� " & iSuccInst
    If iRepeat > 0 Then
        pmRetMsg = pmRetMsg & vbCrLf & _
            "���ظ��޷�����������У� " & iRepeat & vbCrLf & _
            strRepeat
    End If
    If iRelate > 0 Then
        pmRetMsg = pmRetMsg & vbCrLf & _
            "��Ҫ���������޷�����������У� " & iRelate & vbCrLf & _
            strRelate
    End If
    pmRetMsg = pmRetMsg & vbCrLf & "����ɹ���"
    DisplayMsg "���ݵ���ɹ���", vbInformation
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
        'Υ��Ψһ������Լ��
        If gAdoConnDB.Errors(0).NativeError = -105121349 Then
        'sqlstate=3022
            rsTarget.CancelUpdate
            iRepeat = iRepeat + 1
            strRepeat = strRepeat & iLoadCount & ","
            Resume Next
        End If
        'Υ���������ϵԼ��
        If gAdoConnDB.Errors(0).NativeError = -535037517 Then
        'sqlstate=3201
            rsTarget.CancelUpdate
            iRelate = iRelate + 1
            strRelate = strRelate & iLoadCount & ","
            Resume Next
        End If
    End If

    DisplayMsg "���ݵ���ʧ�ܣ�" & vbCrLf & vbCrLf & "���鵼������Դ '" & pmStrSourceRS & "' �Ƿ���ϸ�ʽ��" _
        & vbCrLf & vbCrLf & "�����ȷ������ѯ��ά����Ա��", vbCritical

    pmRetMsg = "����ʧ�ܣ������ļ��Ƿ���ϸ�ʽ��"
    Select Case Err.Number
    Case 3265
        pmRetMsg = pmRetMsg & vbCrLf & "����ԭ�������ļ��޴��ֶΡ�" & rsSource(i).Name & "����"
    Case -2147217904
        pmRetMsg = pmRetMsg & vbCrLf & "����ԭ�������ļ�������һ�����ϵ��ֶ��Ҳ�����"
    Case -2147467259
        pmRetMsg = pmRetMsg & vbCrLf & "����ԭ�������ļ��Ҳ�����Ӧ�ı��"
    Case -2147217887
        pmRetMsg = pmRetMsg & vbCrLf & "����ԭ�������ļ��е� " & iLoadCount & " �����еĹؼ���Ϣ��Ա������������ȶ�ȡ����" & vbCrLf & _
                        vbTab & "  ������������" & rsSource.GetString(adClipString, , ",")
    End Select

    pmRetMsg = pmRetMsg & vbCrLf & vbCrLf & "����Ĳ�����Ϣ���£�" & vbCrLf & _
                "�ⲿ����Դ��" & pmStrSourceRS & vbCrLf & _
                "Ŀ������Դ��" & pmStrTargetRS & vbCrLf & _
                "�ֶ�����" & pmIntFieldCount & vbCrLf & _
                "���ڴ�����кţ�" & iLoadCount & vbCrLf & vbCrLf & _
                "����������ϸ��Ϣ���£�" & vbCrLf & _
                "�����# " & Err.Number & vbCrLf & _
                "��������: " & Err.Description
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
    '�޸�ϵͳʹ�ô���
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
