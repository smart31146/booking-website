<%@LANGUAGE="VBSCRIPT"%>
<% Response.Buffer="true" %>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
Dim user__MMColParam
user__MMColParam = "1"
if (Request.QueryString("user") <> "") then user__MMColParam = Request.QueryString("user")
numbers=1
set user = Server.CreateObject("ADODB.Recordset")
user.ActiveConnection = Connect
user.Source = "SELECT * FROM q_user WHERE ID_user = " + Replace(user__MMColParam, "'", "''") + ""
user.CursorType = 0
user.CursorLocation = 3
user.LockType = 3
user.Open()
user_numRows = 0

Dim sessions__MMColParam
sessions__MMColParam = "1"
if (Request.QueryString("user") <> "") then sessions__MMColParam = Request.QueryString("user")

subject_prm = cint(request("subject"))
passrate = cint(request("passrate"))

if cint(request("subject")) <> 0 then
	t ="and (q_session.Session_subject ="&subject_prm&")"
else
	t = ""
end if	

if cstr(request("fromdate")) <> "" and cstr(request("todate"))= "" then
	t1 = "and session_finish >='"&request("fromdate")&"'"
else if cstr(request("todate")) <> "" and cstr(request("fromdate")) ="" then
	t1 = "and  session_finish <='"&request("todate")&"'"
else if cstr(request("fromdate")) <> "" and cstr(request("todate")) <> "" then
	t1 = "and q_session.session_finish between '"&request("fromdate")&"' and '"&request("todate")&"'"
else
	t1=""
end if
end if
end if

set sessions = Server.CreateObject("ADODB.Recordset")
sessions.ActiveConnection = Connect
If session("mths") = "" then
	sessions.Source = "SELECT q_session.ID_Session, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE Session_users = " + Replace(sessions__MMColParam, "'", "''") + " "&t&" "&t1&" ORDER BY subjects.subject_name, q_session.Session_date desc;"
else
	sessions.Source = "SELECT q_session.ID_Session, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE Session_users = " + Replace(sessions__MMColParam, "'", "''") + "AND (q_session.Session_done = 1) "&t&" "&t1&" ORDER BY subjects.id_subject, q_session.Session_date DESC ;"
end if
'Response.Write sessions.Source
sessions.CursorType = 0
sessions.CursorLocation = 3
sessions.LockType = 3
sessions.Open()
sessions_numRows = 0
%>
<HEAD>
</HEAD>
<BODY>
<%
Response.Clear()
Response.AddHeader "Content-Disposition","attachment; filename=USER_Sessions" & day(now()) & "_" & month(now()) & "_" & year(now()) & ".csv"
Response.ContentType="application/vnd.ms-excel"
%>
"USERS - Sessions"<%=vbcrlf%><%=vbcrlf%>"<%=(user.Fields.Item("user_firstname").Value) & " "%><%=(user.Fields.Item("user_lastname").Value)%>'s session(s):"<%=vbcrlf%><%=vbcrlf%>"Subject","Date","Correct","Total Quest","Up to Page","Finished","Rate","Pass"<%=vbcrlf%>"-------------------------------------------------------------------------------------------------------------------------"<%=vbcrlf%><%
overall_rate = 0
sum_tests = 0
passrate = request("passrate")
While (NOT sessions.EOF)
if session("mths") <> "" then
if cint(subid) <> (sessions.Fields.Item("id_subject").Value)then 
	subid = (sessions.Fields.Item("id_subject").Value)
user_rate = FormatNumber((sessions.Fields.Item("Session_correct").Value)/(sessions.Fields.Item("Session_total").Value)*100,2)
if cInt(user_rate) >= cInt(passrate) then user_pass = 1 else user_pass = 0
overall_rate = overall_rate + user_rate
%>"<%=(sessions.Fields.Item("subject_name").Value)%>","<%=(sessions.Fields.Item("Session_date").Value)%>","<%=(sessions.Fields.Item("Session_correct").Value)%>","<%=(sessions.Fields.Item("Session_total").Value)%>","<%=(sessions.Fields.Item("Session_stop").Value)%>","<%
if abs(sessions.Fields.Item("Session_done").Value) = 1 then 
response.write "YES" 
sum_tests = sum_tests+1
else 
response.write "NO"
end if
%>","<%if user_pass = 1 then response.write (user_rate & "%") else response.write (user_rate & "%")%>","<%if user_pass = 1 then response.write "PASSED" else response.write "FAILED"%>"<%=vbcrlf%><% 
sessions.MoveNext()
else
	sessions.MoveNext()
end if
else
user_rate = FormatNumber((sessions.Fields.Item("Session_correct").Value)/(sessions.Fields.Item("Session_total").Value)*100,2)
if cInt(user_rate) >= cInt(passrate) then user_pass = 1 else user_pass = 0
overall_rate = overall_rate + user_rate
%>"<%=(sessions.Fields.Item("subject_name").Value)%>","<%=(sessions.Fields.Item("Session_date").Value)%>","<%=(sessions.Fields.Item("Session_correct").Value)%>","<%=(sessions.Fields.Item("Session_total").Value)%>","<%=(sessions.Fields.Item("Session_stop").Value)%>","<%
if abs(sessions.Fields.Item("Session_done").Value) = 1 then 
response.write "YES" 
sum_tests = sum_tests+1
else 
response.write "NO"
end if
%>","<%if user_pass = 1 then response.write (user_rate & "%") else response.write (user_rate & "%")%>","<%if user_pass = 1 then response.write "PASSED" else response.write "FAILED"%>"<%=vbcrlf%><% 
sessions.MoveNext()
end if
Wend
if sum_tests = 0 then overall_pass = 0 else overall_pass = overall_rate/sum_tests
%>"-------------------------------------------------------------------------------------------------------------------------"<%=vbcrlf%>"Completed Sessions:","<%=sum_tests%>"<%=vbcrlf%>"Average %:","<%if user_pass = 1 then response.write (FormatNumber(overall_pass,2) & "%") else response.write (FormatNumber(overall_pass,2) & "%") %>"
<%=vbcrlf%>"-------------------------------------------------------------------------------------------------------------------------"<%=vbcrlf%>"Generated on:","<%=Now()%>"<%=vbcrlf%><%=vbcrlf%>"Copyright <% Response.Write Year(now) %> (c) Law of the Jungle Pty Limited" 
<%
Response.Flush()
Response.End()
%>
</BODY>
<%
user.Close()
%>
<%
sessions.Close()
%>