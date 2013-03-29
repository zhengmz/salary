VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_wsk As Winsock
Attribute m_wsk.VB_VarHelpID = -1
Private m_strResponse As String
Private m_strFrom As String
Private m_strFromName As String
Private m_strSubject As String
Private m_strErrMsg As String

Private Sub Class_Initialize()
    Set m_wsk = Nothing
End Sub

Public Sub BindObj(pmWsk As Winsock)
    Set m_wsk = pmWsk
End Sub

Public Sub ReleaseObj()
    Set m_wsk = Nothing
End Sub

Private Sub Class_Terminate()
    Set m_wsk = Nothing
End Sub

'���շ���������Ӧ��Ϣ
Private Sub m_wsk_DataArrival(ByVal bytesTotal As Long)
    m_wsk.GetData m_strResponse
    m_strErrMsg = m_strResponse
    WriteLog "CMail:m_wsk_DataArrival", "winsock�յ�����Ϣ��: " & m_strResponse, LOG_LEVEL_DEBUG
End Sub

Private Sub m_wsk_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    m_strErrMsg = "winsock����" & vbCrLf & _
            "����#��" & Number & vbCrLf & _
            "�������ݣ�" & Description
End Sub

'�ú������������ʼ�������
Public Function Connect(ByVal ServerIp As String, Optional ByVal ServerPort As String) As Boolean
    Connect = False
    '����Winsock
    m_wsk.Close
    m_wsk.RemoteHost = ServerIp
    If ServerPort = "" Then
        m_wsk.RemotePort = 25
    Else
        m_wsk.RemotePort = Val(ServerPort)
    End If
    '����SMTP������
    WriteLog "CMail:Connect", "Connecting Server [" & ServerIp & "] ... ", LOG_LEVEL_DEBUG
    m_wsk.Connect
    If WaitForResponse("220", 10) <> 0 Then
        Exit Function
    End If
    Connect = True
End Function

'��ʼ����������ʹ򿪻Ự
'Ҳ��ʵ����֤��¼
Public Function Init(ByVal strFrom As String, ByVal strFromName As String, ByVal strSubject As String) As Boolean
    m_strFrom = strFrom
    m_strFromName = strFromName
    m_strSubject = strSubject
    
    Init = False
    WriteLog "CMail:Init", "Reset and open session ... ", LOG_LEVEL_DEBUG
    m_wsk.SendData "RSET" & vbCrLf
    If WaitForResponse("250", 10) <> 0 Then
        Exit Function
    End If
    '�򿪶Ի�
    m_wsk.SendData "EHLO " & vbCrLf
    If WaitForResponse("250", 10) <> 0 Then
        Exit Function
    End If
    
    '��֤��¼
'    WriteLog "CMail:Init", "Auth login ... ", LOG_LEVEL_DEBUG
'    m_wsk.SendData "AUTH LOGIN" & vbCrLf
'    If WaitForResponse("334", 10) <> 0 Then
'        Exit Function
'    End If
'    m_wsk.SendData "emhlbmdteg==" & vbCrLf  'zhengmz
'    If WaitForResponse("334", 10) <> 0 Then
'        Exit Function
'    End If
'    m_wsk.SendData "Y21zekAxMzMw" & vbCrLf
'    If WaitForResponse("235", 10) <> 0 Then
'        Exit Function
'    End If

    Init = True
End Function

'�Ͽ�����������
Public Function Disconnect() As Boolean
    Disconnect = False
    If m_wsk.State = sckConnected Then
        WriteLog "CMail:Disconnect", "Disconnecting Server  ... ", LOG_LEVEL_DEBUG
        m_wsk.SendData "QUIT" & vbCrLf
        If WaitForResponse("221", 10) <> 0 Then
            Exit Function
        End If
    End If
    m_wsk.Close
    Disconnect = True
End Function

'�ú������ڵȴ���������Ӧ��
'����ֵ��
'  0   ����
'  -1  �쳣
'  -2  ��ʱ
Private Function WaitForResponse(ByVal strResponse As String, ByVal WaitTime As Integer) As Integer
    Dim WaitSt As Date
    Dim iRet As Integer
    WaitSt = Now()

    WaitForResponse = -1
    WriteLog "CMail:WaitForResponse", "�ȴ�����ֵ[" & strResponse & "], Wait...", LOG_LEVEL_DEBUG
    Do Until DateDiff("s", WaitSt, Now) > WaitTime
        If InStr(1, m_strResponse, strResponse, vbTextCompare) > 0 Then
            m_strResponse = ""
            WaitForResponse = 0
            Exit Function
        ElseIf Len(m_strResponse) > 0 Then
            m_strErrMsg = "�ȴ�����ֵ[" & strResponse & "]��ʵ�ʷ������ݣ�" & m_strResponse
            m_strResponse = ""
            WriteLog "CMail:WaitForResponse", m_strErrMsg, LOG_LEVEL_DEBUG
            Exit Function
        End If
        If m_wsk.State = sckError Then
            m_strErrMsg = "winsock���󣬴����:" & Err.Number & ", ��������: " & Err.Description
            m_strResponse = ""
            WriteLog "CMail:WaitForResponse", m_strErrMsg, LOG_LEVEL_DEBUG
            Exit Function
        End If
        DoEvents
        'WriteLog "CMail:WaitForResponse", "�ȴ�����ֵ[" & strResponse & "], Wait...", LOG_LEVEL_DEBUG
    Loop
    m_strResponse = ""
    WaitForResponse = -2
    m_strErrMsg = "�ȴ�����ֵ[" & strResponse & "]��ʱ���ȴ�ʱ��Ϊ" & WaitTime & "��"
    WriteLog "CMail:WaitForResponse", m_strErrMsg, LOG_LEVEL_DEBUG
End Function

Public Function SendMail(ByVal strTo As String, ByVal strBody As String, Optional pmStrErrMsg As String) As Boolean
    On Error GoTo ErrHandle

    pmStrErrMsg = ""
    SendMail = False
'    WriteLog "CMail:SendMail", "winsock current state: " & m_wsk.State, LOG_LEVEL_DEBUG
    WriteLog "CMail:SendMail", "Begin Send to: [" & strTo & "] ...", LOG_LEVEL_DEBUG
    m_wsk.SendData "RSET" & vbCrLf
    If WaitForResponse("250", 10) <> 0 Then
        GoTo ErrHandle
    End If
    '�򿪶Ի�
    m_wsk.SendData "EHLO " & vbCrLf
    If WaitForResponse("250", 10) <> 0 Then
        GoTo ErrHandle
    End If
    
    '��֤��¼
'    m_wsk.SendData "AUTH LOGIN" & vbCrLf
'    If WaitForResponse("334", 10) <> 0 Then
'        GoTo ErrHandle
'    End If
'    m_wsk.SendData "emhlbmdteg==" & vbCrLf  'zhengmz
'    If WaitForResponse("334", 10) <> 0 Then
'        GoTo ErrHandle
'    End If
'    m_wsk.SendData "Y21zekAxMzMw" & vbCrLf
'    If WaitForResponse("235", 10) <> 0 Then
'        GoTo ErrHandle
'    End If
    
    '���ͷ���ַ
    m_wsk.SendData "MAIL FROM:" & m_strFrom & vbCrLf
    If WaitForResponse("250", 10) <> 0 Then
        GoTo ErrHandle
    End If
    '���շ���ַ
    m_wsk.SendData "RCPT TO:" & strTo & vbCrLf
    If WaitForResponse("250", 10) <> 0 Then
        GoTo ErrHandle
    End If
    '������Ϣ��
    m_wsk.SendData "DATA" & vbCrLf
    If WaitForResponse("354", 10) <> 0 Then
        GoTo ErrHandle
    End If
    m_wsk.SendData "From: " & m_strFromName & " <" & m_strFrom & ">" & vbCrLf
    m_wsk.SendData "To: <" & strTo & ">" & vbCrLf
    m_wsk.SendData "Subject: " & m_strSubject & vbCrLf
    m_wsk.SendData "Content-type: text/html; charset=GBK" & vbCrLf & vbCrLf
    m_wsk.SendData strBody & vbCrLf
    m_wsk.SendData vbCrLf & "." & vbCrLf
    If WaitForResponse("250", 20) <> 0 Then
        GoTo ErrHandle
    End If
'    m_wsk.SendData "NOOP" & vbCrLf
'    If WaitForResponse("250", 10) <> 0 Then
'        GoTo ErrHandle
'    End If
    SendMail = True
    WriteLog "CMail:SendMail", "End Send to: [" & strTo & "] OK!", LOG_LEVEL_DEBUG
    Exit Function

ErrHandle:
    If m_wsk.State = sckError Then
        m_strErrMsg = m_strErrMsg & vbCrLf & _
            "winsock���󣬴����:" & Err.Number & ", ��������: " & Err.Description
    End If
    pmStrErrMsg = m_strErrMsg
    WriteLog "CMail:SendMail", "End Send to: [" & strTo & "] ERR!", LOG_LEVEL_DEBUG
End Function

'��ȡ����������״̬
'sckClosed 0 ȱʡ�ġ��ر�
'sckOpen 1 ��
'sckListening 2 ����
'sckConnectionPending  3 ���ӹ���
'sckResolvingHost  4 ʶ������
'sckHostResolved  5 ��ʶ������
'sckConnecting  6 ��������
'sckConnected  7 ������
'sckClosing  8 ͬ����Ա���ڹر�����
'sckError  9 ����
Public Function GetState() As Integer
    GetState = m_wsk.State
End Function