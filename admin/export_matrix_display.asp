<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
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


set tmp = Server.CreateObject("ADODB.Recordset")
tmp.ActiveConnection = Connect
tmp.Source = "SELECT info1 from q_info1 where id_info1="&info1
tmp.CursorType = 0
tmp.CursorLocation = 3
tmp.LockType = 3
tmp.Open()
while not tmp.EOF
	business = tmp("info1")
	tmp.MoveNext
wend
if business ="" then
	business ="All Businesses"
end if
tmp.Close

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
if site ="" then
	site ="All <% =BBPinfo3s %>"
end if

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
if act ="" then
	act ="All Activities"
end if

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
if tname ="" then
	tname ="All Topics"
end if

set questions = Server.CreateObject("ADODB.Recordset")
questions.ActiveConnection = Connect
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
	users.Source = users.Source & " order by a.user_firstname"
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
end if
'Response.Write users.Source
users.CursorType = 0
users.CursorLocation = 3
users.LockType = 3
users.Open()

%>
<BODY>
<%Response.Clear()
Response.AddHeader "Content-Disposition","inline; filename=matrix_users_" & day(now()) & "_" & month(now()) & "_" & year(now()) & ".xls"
Response.ContentType = "application/vnd.ms-excel"%>
<br><br>
<table border="0" cellspacing="1" cellpadding="0" align=center>
<tr><td width=15><b>Subject </td><td width=15> : </td><td width=250><%=questions("subject_name")%></td></tr>
<tr><td width=15><b>Topic </td><td width=15> : </td><td width=250><%=tname%></td></tr>
<tr><td width=15><b>Business </td><td width=15> : </td><td width=250><%=business%></td></tr>
<tr><td width=15><b><% =BBPinfo3 %> </td><td width=15> : </td><td width=250><%=site%></td></tr>
<tr><td width=15><b><% =BBPinfo3 %> </td><td width=15> : </td><td width=250><%=act%></td></tr>
</tr>
</table><br><br>

<table border="1" cellspacing="1" cellpadding="0" align=center>
<tr>
	<td colspan=2 rowspan=2 align=center width=200>Users</td>
	<td colspan=<%=questions.RecordCount + 1%> align=center >Question ID's<br><br></td>
</tr>
<tr>
	
	<%
	colspan = questions.RecordCount +1
	total_Q_cnt = questions.RecordCount
	dim quest_arr()
	redim quest_arr(total_Q_cnt) 
	t = 0
	while not questions.EOF
	question_text = questions("question_body")
	question_text = replace(question_text,chr(13)," ")
	question_text = replace(question_text,chr(10)," ")
	question_text = replace(question_text,chr(39),chr(96))
	question_text = replace(question_text,chr(34),chr(96))
	
	%>
		<td width=10 align=center>&nbsp;<%=questions("ID_question")%></td>
	<%
	quest_arr(t) = questions("ID_question")
	t = t + 1
	questions.MoveNext
	
	wend
	questions.Close%>
	<td>User result</td>
</tr>   

<%
Response.Flush
userid=0
if users.EOF then%>
	<tr>
	<td colspan=<%=colspan + 2%> align=center>No users found</td>
	</tr>	
<%else
sess = 0
while not users.EOF
%>
	<%
	if sess <>  users("id_session") then%>
	<tr>
	<%
		if userid <> users("id_user") then%>
			<td width=100 valign=top><%=users("user_firstname")%>&nbsp;<%=users("user_lastname")%></td>
			<td width=100 valign=top><%'=users("id_session")%><%=users("session_date")%></td>
		<%
		else%>
			<td width=100 valign=top>&nbsp;</td>
			<td width=100 valign=top><%'=users("id_session")%><%=users("session_date")%></td>
		<%end if
		set results = Server.CreateObject("ADODB.Recordset")
		results.ActiveConnection = Connect
		'results.Source = "select a.id_question, a.question_topic, b.id_result, c.id_session from q_question a, q_result b, q_session c, q_choice d where a.id_question = b.result_question and a.question_topic=" &topic &" and c.id_session = b.result_session and d.id_choice=b.result_answer and d.choice_cor = 1 "
		id_user = users("id_user")
		id_session = users("id_session")
		if cstr(topic)<> "0" then
			results.Source = "select a.id_question, a.question_topic, b.id_result, c.id_session, e.id_user, e.user_lastname,e.user_firstname, c.id_session,c.session_date, d.choice_cor from q_question a, q_result b, q_session c, q_choice d, q_user e where e.id_user = c.session_users and c.session_subject="&subj&" and a.id_question = b.result_question and a.question_topic="&topic&" and c.id_session = b.result_session and d.id_choice=b.result_answer and e.id_user = "&id_user&" and c.id_session = "&id_session&"  order by c.id_session, a.id_question"
		else
			results.Source = "select a.id_question, a.question_topic, b.id_result, c.id_session, e.id_user, e.user_lastname,e.user_firstname, c.id_session,c.session_date, d.choice_cor from q_question a, q_result b, q_session c, q_choice d, q_user e where e.id_user = c.session_users and c.session_subject="&subj&" and a.id_question = b.result_question and c.id_session = b.result_session and d.id_choice=b.result_answer and e.id_user = "&id_user&" and c.id_session = "&id_session&"  order by c.id_session, a.id_question"
		end if
		
		'Response.Write results.Source
		results.CursorType = 0
		results.CursorLocation = 3
		results.LockType = 3
		results.Open()
		i = 0 
		green= 0 
		red = 0
		temp = true
		while not results.EOF
			id_question = results("id_question")
			while temp
				if quest_arr(i)<> id_question then
					%>
					<td align=center>-</td>
					<%
				else
					temp = false
				end if
				i = i + 1
			wend
			if results("choice_cor") then
			%>
			<td align=center><font color="green">1<%'=results("id_question")%></font></td>
			<%
			green = green + 1
			else%>
			<td align=center><font color="red">X<%'=results("id_question")%></font></td>
			<%
			red = red + 1
			end if
			Response.Flush
			results.MoveNext
			temp = true
		wend
		results.close()
		while i < total_Q_cnt
			%>
			<td align=center>-</td>
			<%
			i = i + 1
		wend
		%>
		<td><font color="green"><%=green%></font>+<font color="red"><%=red%></font>=<%=green + red%></td>
	</tr> 
	<%
	userid = users("id_user")
	green= 0 
	red = 0
end if
sess =  users("id_session")
Response.Flush
users.MoveNext
wend
'users.close
%>
<tr>
<td colspan=2 align=right><font color="green">Correctly answered : </font></td>
<%
for i=0 to (ubound(quest_arr)-1)
	quest = cint(quest_arr(i))
	set total_answers = Server.CreateObject("ADODB.Recordset")
	total_answers.ActiveConnection = Connect
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
	'Response.Write total_answers.source
	total_answers.CursorType = 0
	total_answers.CursorLocation = 3
	total_answers.LockType = 3
	total_answers.Open()
	while not total_answers.EOF
	%>
	<td><font color="green"><%=total_answers("choice")%></font></td>
	<%
	total_answers.MoveNext
	wend
	total_answers.Close()
next
	%>
</tr>
<tr>
<td colspan=2 align=right><font color="red">Wrongly answered : </font></td>
<%
for i=0 to (ubound(quest_arr)-1)
	quest = cint(quest_arr(i))
	set total_answers = Server.CreateObject("ADODB.Recordset")
	total_answers.ActiveConnection = Connect
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
	total_answers.CursorType = 0
	total_answers.CursorLocation = 3
	total_answers.LockType = 3
	total_answers.Open()
	while not total_answers.EOF
	%>
	<td><font color="red"><%=total_answers("choice")%></font></td>
	<%
	total_answers.MoveNext
	wend
	total_answers.Close()
next
end if
%>
</tr>
</table>
<br>
<table>
         
</table>
</form>

</BODY>
</HTML>

<%
users.close

%>

