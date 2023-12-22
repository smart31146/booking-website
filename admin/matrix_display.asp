<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>Matrix</TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="javascript" type="text/javascript" src="overlib.js?v=bbp34">
</script>
<script>
//function printWindow() {
//bV = parseInt(navigator.appVersion);
//if (bV >= 4) window.print();
//}
function MM_showHideLayers() { //v3.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
    obj.visibility=v; }
}
function show(object) {
    if (document.layers && document.layers[object])
        document.layers[object].visibility = 'visible';
    else if (document.all)
        document.all[object].style.visibility = 'visible';
}
function hide(object) {
    if (document.layers && document.layers[object])
        document.layers[object].visibility = 'hidden';
    else if (document.all)
        document.all[object].style.visibility = 'hidden';
}
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}
</script>
</HEAD>
<%
server.ScriptTimeout=3600
subj = Request.QueryString("subj")
topic = Request.QueryString("topic")
info1 = Request.QueryString("info1")
info2 = Request.QueryString("info2")
info3 = Request.QueryString("info3")
' ADDED 3 JAN 2007 / Johan Bredenholt
fromdate = request.QueryString("fromdate")
todate = request.QueryString("todate")
sortby = request.QueryString("sortby")
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
tmp.Source = "SELECT s_topic from new_subjects where s_ID="&topic
tmp.CursorType = 0
tmp.CursorLocation = 3
tmp.LockType = 3
tmp.Open()
while not tmp.EOF
	tname = tmp("s_topic")
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
	questions.Source = "SELECT a.ID_question,a.question_body, b.s_ID, b.s_topic, b.s_qID, c.ID_subject, c.subject_name from q_question a, new_subjects b, subjects c where c.Id_subject = b.s_qID and b.s_ID = a.question_topic and b.s_ID="&topic & " and c.Id_subject="& subj & " order by a.ID_question"
else
	questions.Source = "SELECT a.ID_question,a.question_body, b.s_ID, b.s_topic, b.s_qID, c.ID_subject, c.subject_name from q_question a, new_subjects b, subjects c where c.Id_subject = b.s_qID and b.s_ID = a.question_topic and c.Id_subject="& subj & " order by a.ID_question"
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
		users.Source = users.Source & " and a.user_info1='"&info1&"'"
	end if
	if cstr(info2) <> "0" then
		users.Source = users.Source & " and a.user_info2='"&info2&"'"
	end if
	if cstr(info3) <> "0" then
		users.Source = users.Source & " and a.user_info3='"&info3&"'"
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
	users.Source = users.Source & " order by session_finish"
	else
	users.Source = users.Source & " order by a."&sortby
	end if
	' ------------------------- END
else
	users.Source = "select a.id_question, a.question_topic, b.id_result, c.id_session, e.id_user, e.user_lastname,e.user_firstname, c.id_session,c.session_date from q_question a, q_result b, q_session c, q_choice d, q_user e where 	e.id_user = c.session_users	and c.session_subject="&subj&" and c.session_done=1 and a.id_question = b.result_question and a.question_topic in (select s_ID from new_subjects where s_qID="&subj&") and c.id_session = b.result_session and d.id_choice=b.result_answer and d.choice_cor =1 "
	if cstr(info1) <> "0" then
		users.Source = users.Source & " and e.user_info1='"&info1&"'"
	end if
	if cstr(info2) <> "0" then
		users.Source = users.Source & " and e.user_info2='"&info2&"'"
	end if
	if cstr(info3) <> "0" then
		users.Source = users.Source & " and e.user_info3='"&info3&"'"
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
'response.end
users.CursorType = 0
users.CursorLocation = 3
users.LockType = 3
users.Open()

%>
<BODY>

<table>
<tr><td class="text" colspan=3 align=right><br><a href="matrix_export.asp?subj=<%=subj%>&fromdate=<%=fromdate%>&sortby=<%=sortby%>&todate=<%=todate%>&topic=<%=topic%>&info1=<%=info1%>&info2=<%=info2%>&info3=<%=info3%>"><img src="images/xls.gif" border=0></a></td></tr>
<tr><td class="text" width=15><b>Subject </td><td width=15 class="text"> : </td><td class="text" width=250><%=questions("subject_name")%></td></tr>
<tr><td class="text" width=15><b>Topic </td><td width=15 class="text"> : </td><td class="text" width=250><%=tname%></td></tr>
<tr><td class="text" width=15><b>Business </td><td width=15 class="text"> : </td><td class="text" width=250><%=business%></td></tr>
<tr><td class="text" width=15><b>Site </td><td width=15 class="text"> : </td><td class="text" width=250><%=site%></td></tr>
<tr><td class="text" width=15><b><% =BBPinfo3 %> </td><td width=15 class="text"> : </td><td class="text" width=250><%=act%></td></tr>

<% if cstr(fromdate) <> "" or cstr(todate) <>"" then%><tr><td class="text" width=150><b>Sessions <% if cstr(fromdate) <> "" and cstr(todate)= "" then
			response.write "from"
	elseif cstr(todate) <> "" and cstr(fromdate) ="" then
			response.write "to"
	elseif cstr(fromdate) <> "" and cstr(todate) <> "" then
			response.write "between"
	end if %> </td><td width=15 class="text"> : </td><td class="text"><% if cstr(fromdate) <> "" and cstr(todate)= "" then
			response.write "" & fromdate
	elseif cstr(todate) <> "" and cstr(fromdate) ="" then
			response.write  "" & todate
	elseif cstr(fromdate) <> "" and cstr(todate) <> "" then
			response.write  "" & fromdate & " and " & todate
	end if %></td></tr><% end if %>
<tr><td class="text" width=150><b>Sort by </td><td width=15 class="text"> : </td><td class="text"><%if sortby="user_firstname" then%>First name<%elseif sortby="user_lastname" then%>Last name<%else%>Date<%end if%></td></tr>

</table><br><br>
<div id="keyptlayer"> 
<font color="red">Please wait while data loads .......</font>
</div>
<table align=center>
<tr>
	<td colspan=2 class="subheads" rowspan=2 align=center width=200>Users</td>
	<td colspan=<%=questions.RecordCount + 1%> align=center class="subheads">Question ID's<br><br></td>
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
		<td class="text" width=10 align=center>&nbsp;<a href="javascript:void(0);" onmouseout="nd('');" onmouseover="overlib('<%=question_text%>',RIGHT);"><%=questions("ID_question")%></a></td>
	<%
	quest_arr(t) = questions("ID_question")
	t = t + 1
	questions.MoveNext
	
	wend
	questions.Close%>
	<td class="text">User result</td>
</tr>   

<%
Response.Flush
userid=0
if users.EOF then%>
	<tr>
	<td class="text" colspan=<%=colspan + 2%> align=center>No users found</td>
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
			<td class="text" width=100 valign=top><%=users("user_firstname")%>&nbsp;<%=users("user_lastname")%></td>
		<% else %>
			<td class="text" width=100 valign=top>&nbsp;</td>
		<% end if %>
		<%
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
					<td class="text" align=center>-</td>
					<%
				else
					temp = false
				end if
				i = i + 1
			wend
			if results("choice_cor") then
			%>
			<td class="text" align=center><font color="green"><img src="images/1_new.gif"><%'=results("id_question")%></font></td>
			<%
			green = green + 1
			else%>
			<td class="text" align=center><font color="red"><img src="images/0_new.gif"><%'=results("id_question")%></font></td>
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
			<td class="text" align=center>-</td>
			<%
			i = i + 1
		wend
		%>
		<td class="text"><font color="green"><%=green%></font>+<font color="red"><%=red%></font>=<%=green + red%></td>
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
<td colspan=2 class="text" align=right><font color="green">Correctly answered : </font></td>
<%
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
	total_answers.Open()
	while not total_answers.EOF
	%>
	<td class="text"><font color="green"><%=total_answers("choice")%></font></td>
	<%
	total_answers.MoveNext
	wend
	total_answers.Close()
next
	%>
</tr>
<tr>
<td colspan=2 class="text" align=right><font color="red">Wrongly answered : </font></td>
<%
for i=0 to (ubound(quest_arr)-1)
	quest = cint(quest_arr(i))
	'set total_answers = Server.CreateObject("ADODB.Recordset")
	'total_answers.ActiveConnection = Connect
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
	while not total_answers.EOF
	%>
	<td class="text"><font color="red"><%=total_answers("choice")%></font></td>
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
<tr>
	<td class="text">Note : Mouse over the question id displays the question text.</td>
</tr>
<!--<tr>
	<td class="text"><br><br><input type="button" onclick="printWindow();" name="print" value="Print Page" class="quiz_button"></td>
</tr>-->          
</table>
</form>
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<script>
MM_showHideLayers('keyptlayer','','hide');
</script>
</BODY>
</HTML>
<%
'call log_the_page ("Matrix List Topics: " & subj)
%>
<%
users.close

%>

