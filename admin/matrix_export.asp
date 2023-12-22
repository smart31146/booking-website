<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>Matrix</TITLE>
</HEAD>
<%
subj = Request.QueryString("subj")
topic = Request.QueryString("topic")
info1 = Request.QueryString("info1")
info2 = Request.QueryString("info2")
info3 = Request.QueryString("info3")
' ADDED 3 JAN 2007 / Johan Bredenholt
fromdate = Request.QueryString("fromdate")
todate = Request.QueryString("todate")
sortby = Request.QueryString("sortby")

if sortby = "" then sortby = "session_finish"

set tmp = Server.CreateObject("ADODB.Recordset")
tmp.ActiveConnection = Connect

business =""
if info1 <> "" then
tmp.Source = "SELECT info1 from q_info1 where id_info1="&info1
tmp.CursorType = 0
tmp.CursorLocation = 3
tmp.LockType = 3
tmp.Open()
while not tmp.EOF
	business = tmp("info1")
	tmp.MoveNext
wend
tmp.Close
else
info1="0"
end if
if business ="" then
	business ="All Businesses"
end if

site = ""
if info2 <> "" then
tmp.Source = "SELECT info2 from q_info2 where id_info2="&info2
tmp.CursorType = 0
tmp.CursorLocation = 3
tmp.LockType = 3
tmp.Open()
while not tmp.EOF
	site = tmp("info2")
	tmp.MoveNext
wend
tmp.Close
else
info2="0"
end if
if site ="" then
	site ="All Sites"
end if

act= ""
if info3 <> "" then
tmp.Source = "SELECT info3 from q_info3 where id_info3="&info3
tmp.CursorType = 0
tmp.CursorLocation = 3
tmp.LockType = 3
tmp.Open()
while not tmp.EOF
	act = tmp("info3")
	tmp.MoveNext
wend
tmp.Close
else
info3="0"
end if
if act ="" then
	act ="All Activities"
end if


tname=""
if topic <> "" then
tmp.Source = "SELECT topic_name from q_topics where id_topic="&topic
tmp.CursorType = 0
tmp.CursorLocation = 3
tmp.LockType = 3
tmp.Open()
while not tmp.EOF
	tname = tmp("topic_name")
	tmp.MoveNext
wend
tmp.Close
else
topic="0"
end if
if tname ="" then
	tname ="All Topics"
end if

set questions = Server.CreateObject("ADODB.Recordset")
questions.ActiveConnection = Connect
set results = Server.CreateObject("ADODB.Recordset")
results.ActiveConnection = Connect
set total_answers = Server.CreateObject("ADODB.Recordset")
total_answers.ActiveConnection = Connect

if cstr(topic) <> "0" then
	questions.Source = "SELECT a.ID_question,a.question_body, b.ID_topic, b.topic_name, b.topic_subject, c.ID_subject, c.subject_name from q_question a, q_topics b, subjects c where c.Id_subject = b.topic_subject and b.Id_topic = a.question_topic and b.Id_topic="&topic & " and c.Id_subject="& subj & " order by a.ID_question"
else
	questions.Source = "SELECT a.ID_question,a.question_body, b.ID_topic, b.topic_name, b.topic_subject, c.ID_subject, c.subject_name from q_question a, q_topics b, subjects c where c.Id_subject = b.topic_subject and b.Id_topic = a.question_topic and c.Id_subject="& subj & " order by a.ID_question"
end if
questions.CursorType = 0
questions.CursorLocation = 3
questions.LockType = 3
questions.Open()

set users = Server.CreateObject("ADODB.Recordset")
users.ActiveConnection = Connect
if cstr(topic)<> "0" then
	users.Source = "select a.id_user, a.user_lastname,a.user_firstname, b.id_session, b.session_date from q_user a, q_session b where a.id_user = b.session_users and b.session_done=1 and b.session_subject="& subj &""
	if cstr(info1) <> "0" then
		users.Source = users.Source & " and a.user_info1="&info1
	end if
	if cstr(info2) <> "0" then
		users.Source = users.Source & " and a.user_info2="&info2
	end if
	if cstr(info3) <> "0" then
		users.Source = users.Source & " and a.user_info3="&info3
	end if
	' ------------------- ADDED DATE FILTER 20 DECEMBER 2006 ---------------------- Johan Bredenholt
	
	if cstr(fromdate) <> "" and cstr(todate)= "" then
			users.Source = users.Source & " and session_finish >= '"&fromdate&"'"
	elseif cstr(todate) <> "" and cstr(fromdate) ="" then
			users.Source = users.Source & " and session_finish <= '"&todate&"'"
	elseif cstr(fromdate) <> "" and cstr(todate) <> "" then
			users.Source = users.Source & " and session_finish <='"&todate&"' and session_finish >='"&fromdate&"'"
	end if
	' ------------------------- END
	
else
	users.Source = "select a.id_question, a.question_topic, b.id_result, c.id_session, e.id_user, e.user_lastname,e.user_firstname, c.id_session,c.session_date from q_question a, q_result b, q_session c, q_choice d, q_user e where 	e.id_user = c.session_users	and c.session_subject="&subj&" and c.session_done=1 and a.id_question = b.result_question and a.question_topic in (select id_topic from q_topics where topic_subject="&subj&") and c.id_session = b.result_session and d.id_choice=b.result_answer and d.choice_cor =1"
	if cstr(info1) <> "0" then
		users.Source = users.Source & " and e.user_info1="&info1
	end if
	if cstr(info2) <> "0" then
		users.Source = users.Source & " and e.user_info2="&info2
	end if
	if cstr(info3) <> "0" then
		users.Source = users.Source & " and e.user_info3="&info3
	end if

' ------------------- ADDED DATE FILTER 20 DECEMBER 2006 ---------------------- Johan Bredenholt
	
	if cstr(fromdate) <> "" and cstr(todate)= "" then
			users.Source = users.Source & " and session_finish >= '"&fromdate&"'"
	elseif cstr(todate) <> "" and cstr(fromdate) ="" then
			users.Source = users.Source & " and session_finish <= '"&todate&"'"
	elseif cstr(fromdate) <> "" and cstr(todate) <> "" then
			users.Source = users.Source & " and session_finish <='"&todate&"' and session_finish >='"&fromdate&"'"
	end if
	
	if sortby = "session_finish" then
	users.Source = users.Source & " order by session_finish,c.id_session"
	else
	users.Source = users.Source & " order by e."&sortby&",c.id_session"
	end if
	' ------------------------- END
end if
'Response.Write users.Source
users.CursorType = 0
users.CursorLocation = 3
users.LockType = 3
users.Open()

%>
<BODY>
<%
Response.Clear()
Response.AddHeader "Content-Disposition","attachment; filename=matrix_users" & day(now()) & "_" & month(now()) & "_" & year(now()) & ".csv"
Response.ContentType="application/vnd.ms-excel"
%>

Subject, <%=questions("subject_name")%>
Topic, <%=tname%>
Business, <%=business%>
<% =BBPinfo3 %>, <%=site%>
<% =BBPinfo3 %>, <%=act%>
Sessions <% if cstr(fromdate) <> "" and cstr(todate)= "" then %>from<%elseif cstr(todate) <> "" and cstr(fromdate) ="" then%>to<%elseif cstr(fromdate) <> "" and cstr(todate) <> "" then%>between<%end if %>, <% if cstr(fromdate) <> "" and cstr(todate)= "" then%><%=fromdate%><%elseif cstr(todate) <> "" and cstr(fromdate) ="" then%><%=todate%><%elseif cstr(fromdate) <> "" and cstr(todate) <> "" then%><%=fromdate%> and <%=todate%><%end if %>
Sort by, <%if sortby="user_firstname" then%>First name<%elseif sortby="user_lastname" then%>Last name<%else%>Date<%end if%>	

Quiz results <%=vbcrlf%>"-------------------------------------------------------------------------------------------------------------------------"  


<%
colspan = questions.RecordCount +1
total_Q_cnt = questions.RecordCount
dim quest_arr()
redim quest_arr(total_Q_cnt) 
t = 0%>Question ID's/User Sessions,,,<%
while not questions.EOF
%>
<%=questions("ID_question")%>
<%
quest_arr(t) = questions("ID_question")%>,<%
t = t + 1
questions.MoveNext
wend
questions.Close%>

<%
Response.Flush
userid=0
if users.EOF then%>
No users found
<%else
sess = 0
while not users.EOF
%>
<%
if sess <>  users("id_session") then%><%if userid <> users("id_user") then%><%fname=users("user_firstname")%><%lname=users("user_lastname")%><%sdate=users("session_date")%><%else%><%fname=""%><%lname=""%><%sdate=users("session_date")%><%end if%>
<%		
'results.Source = "select a.id_question, a.question_topic, b.id_result, c.id_session from q_question a, q_result b, q_session c, q_choice d where a.id_question = b.result_question and a.question_topic=" &topic &" and c.id_session = b.result_session and d.id_choice=b.result_answer and d.choice_cor = 1 "
id_user = users("id_user")
id_session = users("id_session")
if cstr(topic)<> "0" then
	results.Source = "select a.id_question, a.question_topic, b.id_result, c.id_session, e.id_user, e.user_lastname,e.user_firstname, c.id_session,c.session_date, d.choice_cor from q_question a, q_result b, q_session c, q_choice d, q_user e where e.id_user = c.session_users and c.session_subject="&subj&" and a.id_question = b.result_question and a.question_topic="&topic&" and c.id_session = b.result_session and d.id_choice=b.result_answer and e.id_user = "&id_user&" and c.id_session = "&id_session&"  order by c.id_session, a.id_question"
else
	results.Source = "select a.id_question, a.question_topic, b.id_result, c.id_session, e.id_user, e.user_lastname,e.user_firstname, c.id_session,c.session_date, d.choice_cor from q_question a, q_result b, q_session c, q_choice d, q_user e where e.id_user = c.session_users and c.session_subject="&subj&" and a.id_question = b.result_question and c.id_session = b.result_session and d.id_choice=b.result_answer and e.id_user = "&id_user&" and c.id_session = "&id_session&"  order by c.id_session, a.id_question"
end if
		
results.CursorType = 0
results.CursorLocation = 3
results.LockType = 3
results.Open()
i = 0 
green= 0 
red = 0
temp = true%>

"<%=fname%>","<%=lname%>","<%=sdate%>"<%while not results.EOF%><%id_question = results("id_question")%><%while temp%><%if quest_arr(i)<> id_question then%>,-<%else%><%temp = false%><%end if%><%i = i + 1%><%wend%><%if results("choice_cor") then%>,Correct<%green = green + 1%><%else%>,Incorrect<%red = red + 1%><%end if%><%Response.Flush%><%results.MoveNext%><%temp = true%><%wend%><%results.close()%><%while i < total_Q_cnt%>,-<%i = i + 1%><%wend%>,<%=green%>+<%=red%>=<%=green + red%><%userid = users("id_user")%><%green= 0 %><%red = 0%><%end if%><%sess =  users("id_session")%><%Response.Flush%><%users.MoveNext%><%wend%>


Correctly answered : ,,,<%
for i=0 to (ubound(quest_arr)-1)
	quest = cint(quest_arr(i))
	
	if cstr(topic)<> "0" then
		total_answers.Source = "select count(d.choice_cor) as choice from q_question a, q_result b, q_session c, q_choice d, q_user e where e.id_user = c.session_users and c.session_done=1 and c.session_subject="&subj&" and a.id_question = b.result_question and a.question_topic="&topic&" and c.id_session = b.result_session and d.id_choice=b.result_answer and d.choice_cor=1 and a.id_question ="&quest&""		
		if cstr(info1) <> "0" then
			total_answers.Source = total_answers.Source & " and e.user_info1="&info1
		end if
		if cstr(info2) <> "0" then
			total_answers.Source = total_answers.Source & " and e.user_info2="&info2
		end if
		if cstr(info3) <> "0" then
			total_answers.Source = total_answers.Source & " and e.user_info3="&info3
		end if
	else
		total_answers.Source = "select count(d.choice_cor) as choice from q_question a, q_result b, q_session c, q_choice d, q_user e where e.id_user = c.session_users and c.session_done=1 and c.session_subject="&subj&" and a.id_question = b.result_question and c.id_session = b.result_session and d.id_choice=b.result_answer and d.choice_cor=1 and a.id_question ="&quest&""
		if cstr(info1) <> "0" then
			total_answers.Source = total_answers.Source & " and e.user_info1="&info1
		end if
		if cstr(info2) <> "0" then
			total_answers.Source = total_answers.Source & " and e.user_info2="&info2
		end if
		if cstr(info3) <> "0" then
			total_answers.Source = total_answers.Source & " and e.user_info3="&info3
		end if
	end if
	' ------------------- ADDED DATE FILTER 20 DECEMBER 2006 ---------------------- Johan Bredenholt
	
	if cstr(fromdate) <> "" and cstr(todate)= "" then
			total_answers.Source = total_answers.Source & " and session_finish >='"&fromdate&"'"
	elseif cstr(todate) <> "" and cstr(fromdate) ="" then
			total_answers.Source = total_answers.Source & " and session_finish <= '"&todate&"'"
	elseif cstr(fromdate) <> "" and cstr(todate) <> "" then
			total_answers.Source = total_answers.Source & " and session_finish <='"&todate&"' and session_finish >='"&fromdate&"'"
	end if
	' ------------------------- END
'Response.Write total_answers.source
total_answers.CursorType = 0
total_answers.CursorLocation = 3
total_answers.LockType = 3
total_answers.Open()%>
<%while not total_answers.EOF%><%=total_answers("choice")%>,<%total_answers.MoveNext%><%wend%><%total_answers.Close()%><%next%>
Wrongly answered : ,,,<%
for i=0 to (ubound(quest_arr)-1)
	quest = cint(quest_arr(i))
	if cstr(topic)<> "0" then
		total_answers.Source = "select count(d.choice_cor) as choice from q_question a, q_result b, q_session c, q_choice d, q_user e where e.id_user = c.session_users and c.session_subject="&subj&" and a.id_question = b.result_question and a.question_topic="&topic&" and c.session_done=1 and c.id_session = b.result_session and d.id_choice=b.result_answer and d.choice_cor=0 and a.id_question ="&quest&""
		if cstr(info1) <> "0" then
			total_answers.Source = total_answers.Source & " and e.user_info1="&info1
		end if
		if cstr(info2) <> "0" then
			total_answers.Source = total_answers.Source & " and e.user_info2="&info2
		end if
		if cstr(info3) <> "0" then
			total_answers.Source = total_answers.Source & " and e.user_info3="&info3
		end if
	else
		total_answers.Source = "select count(d.choice_cor) as choice from q_question a, q_result b, q_session c, q_choice d, q_user e where e.id_user = c.session_users and c.session_subject="&subj&" and a.id_question = b.result_question and c.session_done=1 and c.id_session = b.result_session and d.id_choice=b.result_answer and d.choice_cor=0 and a.id_question ="&quest&""
		if cstr(info1) <> "0" then
			total_answers.Source = total_answers.Source & " and e.user_info1="&info1
		end if
		if cstr(info2) <> "0" then
			total_answers.Source = total_answers.Source & " and e.user_info2="&info2
		end if
		if cstr(info3) <> "0" then
			total_answers.Source = total_answers.Source & " and e.user_info3="&info3
		end if
	end if
' ------------------- ADDED DATE FILTER 20 DECEMBER 2006 ---------------------- Johan Bredenholt
	
	if cstr(fromdate) <> "" and cstr(todate)= "" then
			total_answers.Source = total_answers.Source & " and session_finish >='"&fromdate&"'"
	elseif cstr(todate) <> "" and cstr(fromdate) ="" then
			total_answers.Source = total_answers.Source & " and session_finish <= '"&todate&"'"
	elseif cstr(fromdate) <> "" and cstr(todate) <> "" then
			total_answers.Source = total_answers.Source & " and session_finish <='"&todate&"' and session_finish >='"&fromdate&"'"
	end if
	' ------------------------- END
total_answers.CursorType = 0
total_answers.CursorLocation = 3
total_answers.LockType = 3
total_answers.Open()
while not total_answers.EOF%><%=total_answers("choice")%>,<%total_answers.MoveNext%><%wend%><%total_answers.Close()%><%next%><%end if%>


<%=vbcrlf%>"-------------------------------------------------------------------------------------------------------------------------"<%=vbcrlf%>"Generated on:","<%=Now()%>"<%=vbcrlf%><%=vbcrlf%>"Copyright 2002 - 2011 (c) Law of the Jungle Pty Ltd"

<%
call log_the_page ("Matrix Export Topics: " & subj)
%>
<%
users.close

%>

