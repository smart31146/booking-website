<%@LANGUAGE="VBSCRIPT"%>
<% Response.Buffer="true" %>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
Dim user__MMColParam
user__MMColParam = "1"
if (Request.QueryString("user") <> "") then user__MMColParam = Request.QueryString("user")
%>
<%
numbers=1
%>
<%
set user = Server.CreateObject("ADODB.Recordset")
user.ActiveConnection = Connect
user.Source = "SELECT * FROM q_user WHERE ID_user = " + Replace(user__MMColParam, "'", "''") + ""
user.CursorType = 0
user.CursorLocation = 3
user.LockType = 3
user.Open()
user_numRows = 0
%>
<%
Dim sessions__MMColParam
sessions__MMColParam = "1"
if (Request.QueryString("user") <> "") then sessions__MMColParam = Request.QueryString("user")
%>
<%
set sessions = Server.CreateObject("ADODB.Recordset")
sessions.ActiveConnection = Connect
sessions.Source = "SELECT q_session.ID_Session, q_session.Session_date, subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE Session_users = " + Replace(sessions__MMColParam, "'", "''") + "  ORDER BY q_session.Session_date DESC , subjects.subject_name DESC;"
sessions.CursorType = 0
sessions.CursorLocation = 3
sessions.LockType = 3
sessions.Open()
sessions_numRows = 0
%>
<%
Dim answers__MMColParam
answers__MMColParam = "1"
if (Request.QueryString("user_session") <> "") then answers__MMColParam = Request.QueryString("user_session")
%>
<%
set answers = Server.CreateObject("ADODB.Recordset")
answers.ActiveConnection = Connect
answers.Source = "SELECT q_result.ID_result, q_question.ID_question, q_question.question_body, q_topics.topic_name, q_choice.choice_label, q_choice.choice_body, q_choice.choice_cor  FROM q_topics INNER JOIN ((q_result INNER JOIN q_question ON q_result.result_question = q_question.ID_question) INNER JOIN q_choice ON (q_question.ID_question = q_choice.choice_question) AND (q_result.result_answer = q_choice.ID_choice)) ON q_topics.ID_topic = q_question.question_topic  WHERE result_session = " + Replace(answers__MMColParam, "'", "''") + " ORDER BY q_topics.topic_ord, q_topics.ID_topic, q_result.ID_result;"
answers.CursorType = 0
answers.CursorLocation = 3
answers.LockType = 3
answers.Open()
answers_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
answers_numRows = answers_numRows + Repeat1__numRows
%>
<HEAD>
</HEAD>
<BODY>
<%
Response.Clear()
Response.AddHeader "Content-Disposition","attachment; filename=USER_Results" & day(now()) & "_" & month(now()) & "_" & year(now()) & ".csv"
Response.ContentType="application/vnd.ms-excel"
'Basic-ultradev DownloadRepeatRegion Server Behavior
%>"USERS - Results"<%=vbcrlf%><%=vbcrlf%>""<%=(user.Fields.Item("user_firstname").Value)%>,<%=(user.Fields.Item("user_lastname").Value)%>'s session in <%=(sessions.Fields.Item("subject_name").Value)%>, <%=(sessions.Fields.Item("Session_date").Value)%>:"<%=vbcrlf%><%=vbcrlf%>"Question","Topic","Answer","Correct"<%=vbcrlf%>"-------------------------------------------------------------------------------------------------------------------------"<%=vbcrlf%><%
user_all = 0
user_correct = 0
While ((Repeat1__numRows <> 0) AND (NOT answers.EOF)) 
%>"<%=ClearHTMLTags(answers.Fields.Item("question_body").Value,2)%>","<%=(answers.Fields.Item("topic_name").Value)%>","<%=(answers.Fields.Item("choice_label").Value)%>","<%
if abs(answers.Fields.Item("choice_cor").Value) = 1 then response.write("Yes") else response.write("No")
%>"<%=vbcrlf%><% 
  if abs(answers.Fields.Item("choice_cor").Value) = 1 then user_correct = user_correct +1
  user_all = user_all +1
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  answers.MoveNext()
Wend
%>"-------------------------------------------------------------------------------------------------------------------------"<%=vbcrlf%>"Overall Pass & Rate:","<%if (user_correct/user_all*100) >= passrate then response.write("PASSED - " & FormatNumber(user_correct/user_all*100, 2) & "%") else response.write("FAILED - " & FormatNumber(user_correct/user_all*100, 2) & "%")%>"<%=vbcrlf%>"Correct:","<%=user_correct%>"<%=vbcrlf%>"Incorrect:","<%=user_all-user_correct%>"<%=vbcrlf%>"Pass rate is currently: <%=request("passrate")%> %"<%=vbcrlf%>"-------------------------------------------------------------------------------------------------------------------------"<%=vbcrlf%>"Generated on:","<%=Now()%>"<%=vbcrlf%><%=vbcrlf%>"Copyright 2002 - 2011 (c) Law of the Jungle Pty Limited" 
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
<%
answers.Close()
%>
