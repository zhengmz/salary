Attribute VB_Name = "mdCommon"
Option Explicit

'最大支持字段数，包括姓名和编号
Public Const SERV_FIELD_CNT_MAX As Integer = 101
Public Const SERV_ID_PATTERN As String = "00000000"
Public Const SERV_FIELD_PATTERN As String = "00"

'最大报表字段数
Public Const RPT_FIELD_CNT_MAX As Integer = 70
Public Const RPT_ID_PATTERN As String = "00000000"
Public Const RPT_FIELD_PATTERN As String = "00"

Public Const OPER_QUERY As Integer = 0
Public Const OPER_ADD As Integer = 1
Public Const OPER_MODIFY As Integer = 2
Public Const OPER_DEL As Integer = 3
Public Const OPER_COPY As Integer = 4

Public Const DATE_FORMAT_YEAR As String = "yyyy"
Public Const DATE_FORMAT_MONTH As String = "yyyyMM"
Public Const DATE_FORMAT_DAY As String = "yyyyMMdd"

Public Enum SYS_LOG_LEVEL
    LOG_LEVEL_ALARM = 1
    LOG_LEVEL_SERIOUS = 2
    LOG_LEVEL_INFO = 3
    LOG_LEVEL_DEBUG = 4
End Enum

Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Sub WriteLog(ByVal pmStrFunc As String, ByVal pmStrMsg As String, Optional ByVal pmIntLevel As SYS_LOG_LEVEL = LOG_LEVEL_INFO)
    If pmIntLevel > gSysLogLevel Then
        Exit Sub
    End If
    
    
    Dim fso As FileSystemObject
    Dim strFileName As String
    Dim ts As TextStream
    Dim strLogMsg As String

    strFileName = gSysLogDir & "\log_" & Format(Date, "yyyymmdd") & ".txt"
    strLogMsg = Format(Date, "yyyymmdd") & Format(Time, "HhNnSs") & ":" & pmIntLevel & ":" & pmStrFunc & ":" & pmStrMsg
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(strFileName, ForAppending, True)
    ts.WriteLine (strLogMsg)
    ts.Close
    Set fso = Nothing
End Sub

'用msgbox中止后台程序
Function DisplayMsg(ByVal pmMsg As String, Optional ByVal pmLevel As VbMsgBoxStyle = vbInformation) As VbMsgBoxResult
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = App.Title
    strMsg = pmMsg

    If pmLevel = vbCritical Then
        If Err.Number <> 0 Then
            strMsg = strMsg & vbCrLf & vbCrLf & _
                "错误号# " & Err.Number & vbCrLf & _
                "错误内容: " & Err.Description
        End If
        strTitle = strTitle & "--错误"
    ElseIf pmLevel = vbExclamation Then
        strTitle = strTitle & "--告警"
    Else
        strTitle = strTitle & "--提示"
    End If
    
    DisplayMsg = MsgBox(strMsg, pmLevel, strTitle)
End Function

'用API不会中止后台程序，但缺点是必须指定要返回的窗口句柄，否则效果就象无模式窗口一样
Function DisplayMsgAPI(ByVal pmMsg As String, Optional ByVal pmLevel As VbMsgBoxStyle = vbInformation, Optional ByVal hwnd As Long = 0) As VbMsgBoxResult
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = App.Title
    strMsg = pmMsg

    If pmLevel = vbCritical Then
        If Err.Number <> 0 Then
            strMsg = strMsg & vbCrLf & vbCrLf & _
                "错误号# " & Err.Number & vbCrLf & _
                "错误内容: " & Err.Description
        End If
        strTitle = strTitle & "--错误"
    ElseIf pmLevel = vbExclamation Then
        strTitle = strTitle & "--告警"
    Else
        strTitle = strTitle & "--提示"
    End If
    
    If hwnd = 0 Then
        DisplayMsgAPI = MessageBox(Forms(0).hwnd, strMsg, strTitle, pmLevel)
    Else
        DisplayMsgAPI = MessageBox(hwnd, strMsg, strTitle, pmLevel)
    End If
End Function

'打开一个窗体，先查找如果存在，就设置焦点，否则打开
Public Function OpenForm(ByVal pmFormName As String, Optional ByVal pmFormCaption As String = "", Optional ByVal lngModal As FormShowConstants = vbModeless) As Form
'    On Error Resume Next
    Dim i As Integer

    i = FindForm(pmFormName, pmFormCaption)
    If i > -1 Then
        Set OpenForm = Forms(i)
        Forms(i).SetFocus
        Exit Function
    End If

    Dim newForm As Form
    Set newForm = Forms.Add(pmFormName)
    Load newForm
    If pmFormCaption <> "" Then
        newForm.Caption = pmFormCaption
    End If
    newForm.Show lngModal
    Set OpenForm = newForm
End Function

'查找在应用程序中打开窗体的名称
Public Function FindForm(ByVal pmFormName As String, Optional ByVal pmFormCaption As String = "") As Integer
'    On Error Resume Next
    Dim i As Integer
    
    FindForm = -1
    For i = 0 To Forms.Count - 1
        If Forms(i).Name = pmFormName Then
            If pmFormCaption = "" Then
                FindForm = i
                Exit Function
            ElseIf Forms(i).Caption = pmFormCaption Then
                FindForm = i
                Exit Function
            End If
        End If
    Next i
End Function

Public Sub CloseAllSubForm()
    '关闭原有子窗口
    'On Error Resume Next
    If Forms.Count = 1 Then
        Exit Sub
    End If
    Dim i As Integer
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> "MDIMain" Then
            Unload Forms(i)
            Exit For
        End If
    Next i
    CloseAllSubForm
End Sub

Public Sub DisplayGrid(pmGrid As DataGrid, ByVal pmTableName As String, _
                    Optional ByVal pmStrServId As String = "-1")

    On Error GoTo ErrHandle
    
    Dim i As Integer
    Dim iColCount As Integer
    Dim sArrCaption() As String

    iColCount = pmGrid.Columns.Count
    ReDim sArrCaption(iColCount) As String
    For i = 0 To iColCount - 1
        sArrCaption(i) = pmGrid.Columns(i).Caption
    Next i

    Dim strTableName As String
    Dim rsDisplay As New ADODB.Recordset
    Dim strSQL As String

    strTableName = UCase(pmTableName)
    Select Case strTableName
    Case "SALARY"
        strSQL = "SELECT field_name,display_name,'1000' as display_width FROM serv_field WHERE serv_id='" & pmStrServId & "' and valid_flag=1"
    Case "RPT_DATA"
        strSQL = "SELECT field_name,display_name,'1000' as display_width FROM rpt_field WHERE rpt_id='" & pmStrServId & "' and valid_flag=1"
    Case Else
        strSQL = "SELECT dict_key as field_name, dict_value as display_name, dict_type as display_width FROM dicts WHERE dict_sect='TBL_" & strTableName & "' and dict_flag=1"
    End Select

    rsDisplay.Open strSQL, gStrConnDB, adOpenStatic, adLockReadOnly, adCmdText
    i = 0
    Do Until rsDisplay.EOF
        pmGrid.Columns(rsDisplay("field_name")).Visible = True
        pmGrid.Columns(rsDisplay("field_name")).Width = Val(rsDisplay("display_width"))
        pmGrid.Columns(rsDisplay("field_name")).Caption = rsDisplay("display_name")
        rsDisplay.MoveNext
        i = i + 1
    Loop
    rsDisplay.Close
    Set rsDisplay = Nothing
    
    If i = 0 Or i = iColCount Then
        Exit Sub
    End If
    On Error Resume Next
    If strTableName = "SALARY" Then
        pmGrid.Columns("emp_email").Visible = True
        pmGrid.Columns("emp_email").Width = 3000
        pmGrid.Columns("emp_email").Caption = "电子邮箱"
    End If
    If strTableName = "RPT_DATA" Then
        pmGrid.Columns("emp_id").Visible = True
        pmGrid.Columns("emp_id").Width = 1000
        pmGrid.Columns("emp_id").Caption = "员工编码"
        pmGrid.Columns("emp_name").Visible = True
        pmGrid.Columns("emp_name").Width = 1000
        pmGrid.Columns("emp_name").Caption = "姓名"
    End If

    For i = 0 To iColCount - 1
        If sArrCaption(i) = pmGrid.Columns(i).Caption Then
            pmGrid.Columns(i).Visible = False
        End If
    Next i
    Exit Sub
    
ErrHandle:
    DisplayMsg "显示格式化错误", vbCritical
End Sub

'Base64算法
'按照RFC2045，Base64被定义为：Base64内容传送编码被设计用来把任意序列的8位字节描述为一种不易被人直接识别的形式。
'（The Base64 Content-Transfer-Encoding is designed to represent arbitrary sequences of octets in a form that need not be humanly readable.）
Public Function Base64Encode(ByVal pmStrSource As String) As String
    Const BASE64_TABLE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim strEncode As String
    Dim i As Integer
    For i = 1 To (Len(pmStrSource) - Len(pmStrSource) Mod 3) Step 3
        strEncode = strEncode + Mid(BASE64_TABLE, (Asc(Mid(pmStrSource, i, 1)) \ 4) + 1, 1)
        strEncode = strEncode + Mid(BASE64_TABLE, ((Asc(Mid(pmStrSource, i, 1)) Mod 4) * 16 _
                      + Asc(Mid(pmStrSource, i + 1, 1)) \ 16) + 1, 1)
        strEncode = strEncode + Mid(BASE64_TABLE, ((Asc(Mid(pmStrSource, i + 1, 1)) Mod 16) * 4 _
                      + Asc(Mid(pmStrSource, i + 2, 1)) \ 64) + 1, 1)
        strEncode = strEncode + Mid(BASE64_TABLE, (Asc(Mid(pmStrSource, i + 2, 1)) Mod 64) + 1, 1)
    Next i
    If Not (Len(pmStrSource) Mod 3) = 0 Then
         If (Len(pmStrSource) Mod 3) = 2 Then
            strEncode = strEncode + Mid(BASE64_TABLE, (Asc(Mid(pmStrSource, i, 1)) \ 4) + 1, 1)
            strEncode = strEncode + Mid(BASE64_TABLE, (Asc(Mid(pmStrSource, i, 1)) Mod 4) * 16 _
                      + Asc(Mid(pmStrSource, i + 1, 1)) \ 16 + 1, 1)
             strEncode = strEncode + Mid(BASE64_TABLE, (Asc(Mid(pmStrSource, i + 1, 1)) Mod 16) * 4 + 1, 1)
            strEncode = strEncode & "="
        ElseIf (Len(pmStrSource) Mod 3) = 1 Then
            strEncode = strEncode + Mid(BASE64_TABLE, Asc(Mid(pmStrSource, i, 1)) \ 4 + 1, 1)
            strEncode = strEncode + Mid(BASE64_TABLE, (Asc(Mid(pmStrSource, i, 1)) Mod 4) * 16 + 1, 1)
             strEncode = strEncode & "=="
        End If
     End If
    Base64Encode = strEncode
End Function
