<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>

<%
Dim fid
If (Request.QueryString("fid") <> "") Then 
fid = cInt(Request.QueryString("fid"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
Dim comeback
If (Request.QueryString("comeback") <> "") Then 
comeback = cStr(Request.QueryString("comeback"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = Connect
  MM_editTable = "tr_feedback"
  MM_editColumn = "ID_feedback"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "" + comeback + ""

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
  
End If
%>
<%
' *** Delete Record: construct a sql delete staatement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete staatement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    if Edit_OK = true then MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    call log_the_page ("Training Execute - DELETE Feedback: " & MM_recordId)
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
set f_del = Server.CreateObject("ADODB.Recordset")
f_del.ActiveConnection = Connect
f_del.Source = "SELECT ID_feedback  FROM tr_feedback  WHERE ID_feedback = " + Replace(fid, "'", "''") + ""
f_del.CursorType = 0
f_del.CursorLocation = 3
f_del.LockType = 3
f_del.Open()
f_del_numRows = 0
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>Quiz admin </TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY onload="javascript:document.forms[0].submit();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
  <input type="hidden" name="MM_delete" value="true">
  <input type="hidden" name="MM_recordId" value="<%= f_del.Fields.Item("ID_feedback").Value %>">
</form>
</body>
</HTML>
<%
f_del.Close()
%>
