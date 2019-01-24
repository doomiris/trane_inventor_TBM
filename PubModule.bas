Attribute VB_Name = "PubModule"
Public Const c_ver = "3.54"
Public Const c_date = "2008-2-14"
Public Const DateForTrane = "dd-mmm-yyyy"
Public oApp As Inventor.Application
Public dbpath As String
Public conn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public TbmLan As Integer
Public Function LS(name, default) 'Load Setting
LS = GetSetting("Domisoft", "TBM", name, default)
End Function

Public Sub WS(name, Value)  'Write Setting
SaveSetting "Domisoft", "TBM", name, Value
End Sub

Public Function checknull(str)
str = "" & str
If Format(str, ">") = "NULL" Or CStr(str) = "0" Then
checknull = ""
Else:
checknull = str
End If
End Function
Public Sub CheckUpdate() '检查更新模块

Dim ToCheck As String
ToCheck = LS("DayChecked", "0")

If Not Day(Now()) = ToCheck Then
    Dim ConnStr As String
    UpdateServer = LS("UpdateServer", "\\shahns07\NPI_Program\Chiller_Projects\Global_Scroll_AC\Member_backup\Dominic\UpdateList.xls")
    On Error Resume Next
    ConnStr = "Driver={Microsoft Excel Driver (*.xls)};DBQ=" & UpdateServer
    conn.Open ConnStr
    Set rs = CreateObject("ADODB.Recordset")
    Dim mySQL As String
    mySQL = "select * from [update$];"
    rs.Open mySQL, conn, 1, 3
    
    Call WS("ServerVer", rs("n_ver"))
    
    Set rs = Nothing
    rs.Close
    Set conn = Nothing
    conn.Close
    
    Call WS("DayChecked", Day(Now()))
 '   Call VersionDisabledCheck
End If

Dim ServerVer As String
ServerVer = LS("ServerVer", c_ver)

If ServerVer > c_ver Then
    With frmtbm.sb
        .Enabled = True
        .Caption = "New version:" & ServerVer & " has been released, click here for update!"
    End With
    If LS("AllowAutoUpdate", 1) = 1 Then
    Shell "\\shahns07\NPI_Program\Chiller_Projects\Global_Scroll_AC\Member_backup\Dominic\TBM\Update\AddRunOnce.cmd", vbHide
    End If
End If
End Sub
Public Sub TbmLog(ModeName As String)
'Dim ToLog As String
'ToLog = LS("DayLoged", "0")
'If Day(Now()) = ToLog Then
'    Exit Sub
'End If

If LS("DisableLog", "0") = "1" Then
    Exit Sub
End If

Dim ConnStr As String
Dim c_UserName As String
Dim c_Count As Integer

LogDBpath = LS("LogDBpath", "\\shahns07\NPI_Program\Chiller_Projects\Global_Scroll_AC\Member_backup\Dominic\TBM\tbmlog.tbmlog")
On Error Resume Next
c_UserName = oApp.UserName
c_Count = 1
ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & LogDBpath
conn.Open ConnStr
Set rs = CreateObject("ADODB.Recordset")
Dim mySQL As String
Select Case ModeName
    Case "E"
        mySQL = "select * from [logE] where Username='" & c_UserName & "';"
    Case Else
        mySQL = "select * from [log] where Username='" & c_UserName & "';"
End Select
rs.Open mySQL, conn, 1, 3
Select Case rs.RecordCount
Case 0
    rs.AddNew
    rs("Username") = c_UserName
Case 1
    c_Count = rs("tbmCount") + 1
End Select
rs("tbmCount") = c_Count
rs("LastUse") = Now()
rs("cc_ver") = c_ver
rs.Update
Call dbClose

End Sub
Public Sub dbOpen()
Dim ConnStr As String
On Error GoTo xlsopenerror
ConnStr = "Driver={Microsoft Excel Driver (*.xls)};DBQ=" & dbpath
conn.Open ConnStr
Set rs = CreateObject("ADODB.Recordset")
Exit Sub
xlsopenerror:
MsgBox "Cannot open EXCEL file " & dbpath & ", the file is either in use or not exist, please check up.", vbCritical, "Error"
Call dbInput
End Sub
Public Sub dbClose()
On Error Resume Next
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
End Sub
Private Sub dbInput()
tempdb = InputBox("Input EXCEL database path excatly!", "Input full path", dbpath)
If Not Len(tempdb) = 0 Then
    dbpath = tempdb
    Call dbOpen
Else:
    MsgBox "Please Confirm Database!", vbCritical, "TBM Critical Error"
Unload All
End If
End Sub
'Public Sub VersionDisabledCheck()
'On Error Resume Next
'Dim ConnStr As String
'
'UpdateServer = LS("UpdateServer", "\\shahns07\NPI_Program\Chiller_Projects\Global_Scroll_AC\Member_backup\Dominic\UpdateList.xls")
'On Error Resume Next
'ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & UpdateServer
'conn.Open ConnStr
'Set rs = CreateObject("ADODB.Recordset")
'Dim mySQL As String
'mySQL = "select * from [DisabledVersion$] where Ver='" & c_ver & "';"
'rs.Open mySQL, conn, 1, 3
'
'If rs("Disabled") = "FALSE" Then
'    Call WS("VersionAllowed", c_ver)
'Else:
'    Call WS("VersionAllowed", LS("ServerVer", "unkonwn"))
'End If
'Call dbClose
'
'End Sub
Public Function TrimFileName(FileName As String)
If Not Len(FileName) = 0 Then
    Dim TempSTR
    TempSTR = Split(FileName, ".")
    TrimFileName = TempSTR(0)
Else:
    TrimFileName = FileName
End If
End Function

Public Function LRS(Number As Integer)
LRS = LoadResString(CInt(TbmLan & Number))
End Function
