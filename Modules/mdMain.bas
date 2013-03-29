Attribute VB_Name = "mdMain"
Option Explicit

'���ݿ����Ӵ�
Public gStrConnDB As String
'���ݿ��ļ�����������·��
Public gStrDBFileName As String
'���ݿ�����
Public gAdoConnDB As New ADODB.Connection
'Ӧ������
Public gStrAppName As String

'ϵͳ����
'ϵͳ��־����Ĭ��Ϊ0,������¼
Public gSysLogLevel As Integer
'ϵͳ��־Ŀ¼��Ĭ��ΪApp.path\Log
Public gSysLogDir As String
'�ֱ��ֶ�����Ĭ��Ϊ0����ʾ���ֱ�
Public gSysSplitFields As Integer


Sub Main()
    '��ʼ��
    gStrDBFileName = App.Path & "\Data\salary.dat"
    gStrConnDB = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gStrDBFileName & ";Persist Security Info=False;Jet OLEDB:Database password=mdb@salary"
    gStrAppName = App.EXEName
    gSysLogLevel = 0
    gSysLogDir = App.Path & "\Log"
    gSysSplitFields = 0
    
    '�������ݿ�
    If Not ConnectDB() Then
        DisplayMsg "�������ݿ�����������ݿ��ļ���", vbCritical
        End
    End If
    
    '��鲢����Ŀ¼�ṹ
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
    
    '��ȡϵͳ����
    Call gGetParameter

    '�޸�ϵͳʹ�ô���
    Call gUpdateUsedCount

    MDIMain.Show
End Sub
