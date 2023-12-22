<%@LANGUAGE="VBSCRIPT"%>

<% 
'Response buffer is used to buffer the output page. That means if any database exception occurs the contents can be cleared without processed any script to browser
 Response.Buffer = True
 
' "On Error Resume Next" method allows page to move to the next script even if any error present on page whcich will be caught after processing all asp script on page
 On Error Resume Next
 
'Changed by PR on 25.02.16
%>

<% bbp_training = true%>
<!--#include file="connections/bbg_conn.asp" -->
<!--#include file="connections/include.asp" -->

<%
if NOT pref_quiz_avail then response.redirect("error.asp?" & request.QueryString)



Dim position
'If (Request.QueryString("nextID") <> "") Then
'nextID = cInt(Request.QueryString("nextID"))
'Else
'Response.Redirect("error.asp?" & request.QueryString)
'End If
Dim userid
if (Session("UserID") <> "") Then
userid = cInt(Session("UserID"))
Else
Response.Redirect("error.asp?" & request.QueryString)
End If

Dim SessionID
if (Session("SessionID") <> "") Then
SessionID = cLng(Session("SessionID"))
Else
Response.Redirect("error.asp?" & request.QueryString)
End If
%>

        <input type="hidden" name="discard" value="<%= Request.Form("discard") %>">
        <input type="hidden" name="user" value="<%=userID%>">
        <input type="hidden" name="subject" value="<%=ID_subject_prm%>">
        <input type="hidden" name="date" value="<%=cDateSql(Now())%>">
        <input type="hidden" name="total" value="<%=total%>">
        <input type="hidden" name="quizcorrect" value="0">
        <input type="hidden" name="stop" value="1">
        <input type="hidden" name="done" value="NO">
<%
if Request.Form("discard") = "TRUE" then
  MM_editConnection = Connect
  MM_editTable = "q_session"
  MM_editQuery = "delete from " & MM_editTable & " where ID_session = " & SessionID
if Err.Number = 0 then
 	Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
end if
  MM_editTable = "q_result"
  MM_editQuery = "delete from " & MM_editTable & " where result_session = " & SessionID
if Err.Number = 0 then
 	Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
end if

    MM_editCmd.ActiveConnection.Close
	response.redirect("t_index.asp")
end if
response.redirect("t_question.asp?nextID="&request.querystring("nextID")&"" )
%>
</body>
</HTML>
<!-- #include file = "errorhandler/index.asp"-->


