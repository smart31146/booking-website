<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
numbers=1
count = 1
SQL_having = ""
SQL_where = ""
results=request("results")
fromdate=request("fromdate")
fromdate=cdatesql(fromdate)
todate=request("todate")
if len(todate) < 12 and todate <> "" then
	todate=todate&" 23:59:59"
end if
todate=cdatesql(todate)
active = request("active")
IF request("passrate") <> "" then
passrate=  clng(request("passrate"))
else
passrate =80
END IF
if request("mths") <> "" then
	mths = cint(request("mths"))
else
	mths = ""
end if
noquiz = cstr(request("noquiz"))

if cStr(Request.Querystring("show_lines")) <> "" then show_lines = cInt(Request.Querystring("show_lines"))

If cStr(Request.Querystring("filter_username")) <> "" then
	SQL_having = " HAVING ((q_user.user_lastname) Like '%" + Replace(uCase(cStr(Request.Querystring("filter_username"))), "'", "''") + "%' OR  (q_user.user_firstname) Like '%" + Replace(uCase(cStr(Request.Querystring("filter_username"))), "'", "''") + "%') "
end if

subject_prm = 0

If cInt(Request.Querystring("subject")) <> 0 then
	subject_prm = cInt(Request.Querystring("subject"))
	SQL_where = " WHERE (q_session.Session_subject = " + (Request.Querystring("subject")) + ") "
	if cstr(request("active"))="1" then
		SQL_WHERE=SQL_where + "and q_user.user_active=1"
	else if cstr(request("active"))="0" then
		SQL_WHERE=SQL_where + "and q_user.user_active=0"
	end if
	end if

	if cstr(fromdate) <> "" and cstr(todate) = "" then
		SQL_WHERE = sql_where + "and q_session.session_finish >='"&fromdate&"'"
	else if cstr(todate) <> "" and cstr(fromdate) = "" then
		SQL_WHERE = sql_where + "and q_session.session_finish <='"&todate&"'"
	else if cstr(fromdate) <> "" and cstr(todate) <> "" then
		SQL_WHERE = sql_where + "and q_session.session_finish between '"&fromdate&"' and '"&todate&"'"
	end if
	end if
	end if
else
	if cstr(request("active"))="1" then
		sql_where="where  q_user.user_active=1"
	else if cstr(request("active"))="0" then
		sql_where="where  q_user.user_active=0"
	end if
	end if

	if cstr(fromdate) <> "" and cstr(todate)= "" then
		if request("active")<>"2" then
			SQL_WHERE = sql_where + "and session_finish >='"&fromdate&"'"
		else
			SQL_WHERE = "where q_session.session_finish >= '"&fromdate&"'"
		end if
	else if cstr(todate) <> "" and cstr(fromdate) ="" then
		if request("active")<>"2" then
			SQL_WHERE = sql_where + "and session_finish <= '"&todate&"'"
		else
			SQL_WHERE = "where  session_finish <='"&todate&"'"
		end if
	else if cstr(fromdate) <> "" and cstr(todate) <> "" then
		if request("active")<>"2" then
			SQL_WHERE = sql_where + "and q_session.session_finish <='"&todate&"' and session_finish >='"&fromdate&"'"
		else
			SQL_WHERE = "where q_session.session_finish between '"&fromdate&"' and '"&todate&"'"
		end if
	end if
	end if
	end if
end if

filter_info1_prm = 0
If cInt(Request.Querystring("filter_info1")) <> 0 then
	filter_info1_prm = cInt(Request.Querystring("filter_info1"))
	if SQL_having <> "" then
		SQL_having = SQL_having + " AND (q_user.user_info1)= " + (Request.Querystring("filter_info1")) + " "
	else
		SQL_having = " HAVING (q_user.user_info1)= " + (Request.Querystring("filter_info1")) + " "
	end if
end if

filter_info3_prm = 0
If cInt(Request.Querystring("filter_info3")) <> 0 then
	filter_info3_prm = cInt(Request.Querystring("filter_info3"))
	if SQL_having <> "" then
		SQL_having = SQL_having + " AND (q_user.user_info3)= " + (Request.Querystring("filter_info3")) + " "
	else
		SQL_having = " HAVING (q_user.user_info3)= " + (Request.Querystring("filter_info3")) + " "
	end if
end if

filter_info4_prm = 0
If cInt(Request.Querystring("filter_info4")) <> 0 then
	filter_info4_prm = cInt(Request.Querystring("filter_info4"))
	if SQL_having <> "" then
		SQL_having = SQL_having + " AND (q_user.user_info4)= " + (Request.Querystring("filter_info4")) + " "
	else
		SQL_having = " HAVING (q_user.user_info4)= " + (Request.Querystring("filter_info4")) + " "
	end if
end if

if request("mths")=1 then
	session("mths") = 1
else
	session("mths")=""
end if

set users = Server.CreateObject("ADODB.Recordset")
users.ActiveConnection = Connect
users.Source = "SELECT q_user.ID_user, q_user.user_lastname, q_user.user_firstname, q_info1.info1, q_info2.info2, q_info3.info3, q_info4.info4, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, q_user.user_info4, COUNT(q_session.ID_session) AS session_count FROM (q_info4 RIGHT JOIN (q_info3 RIGHT JOIN (q_info2 RIGHT JOIN (q_info1 RIGHT JOIN q_user ON q_info1.ID_info1 = q_user.user_info1) ON q_info2.ID_info2 = q_user.user_info2) ON q_info3.ID_info3 = q_user.user_info3) ON q_info4.ID_info4 = q_user.user_info4) LEFT JOIN q_session ON q_user.ID_user = q_session.Session_users " + SQL_where + " GROUP BY q_user.user_lastname, q_user.user_firstname, q_user.ID_user, q_info1.info1, q_info2.info2, q_info3.info3, q_info4.info4, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, q_user.user_info4 " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"
'SQL: "SELECT q_user.ID_user, q_user.user_lastname, q_user.user_firstname, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, COUNT(q_session.ID_session) AS session_count FROM q_info2 RIGHT OUTER JOIN q_info1 RIGHT OUTER JOIN q_session RIGHT OUTER JOIN q_user ON q_session.Session_users = q_user.ID_user ON q_info1.ID_info1 = q_user.user_info1 ON  q_info2.ID_info2 = q_user.user_info2 LEFT OUTER JOIN q_info3 ON q_user.user_info3 = q_info3.ID_info3 " + SQL_where + " GROUP BY q_user.user_lastname, q_user.user_firstname, q_user.ID_user, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3 " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"
'Access: "SELECT q_user.ID_user, q_user.user_lastname, q_user.user_firstname, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, COUNT(q_session.ID_session) AS session_count FROM (q_info3 RIGHT JOIN (q_info2 RIGHT JOIN (q_info1 RIGHT JOIN q_user ON q_info1.ID_info1 = q_user.user_info1) ON q_info2.ID_info2 = q_user.user_info2) ON q_info3.ID_info3 = q_user.user_info3) LEFT JOIN q_session ON q_user.ID_user = q_session.Session_users " + SQL_where + " GROUP BY q_user.user_lastname, q_user.user_firstname, q_user.ID_user, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3 " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"
'Response.Write users.Source
users.CursorType = 0
users.CursorLocation = 3
users.LockType = 3
users.Open()
users_numRows = 0

set filter_info1 = Server.CreateObject("ADODB.Recordset")
filter_info1.ActiveConnection = Connect
filter_info1.Source = "SELECT * FROM q_info1 order by info1"
filter_info1.CursorType = 0
filter_info1.CursorLocation = 3
filter_info1.LockType = 3
filter_info1.Open()
filter_info1_numRows = 0

set filter_info3 = Server.CreateObject("ADODB.Recordset")
filter_info3.ActiveConnection = Connect
filter_info3.Source = "SELECT * FROM q_info3 order by info3"
filter_info3.CursorType = 0
filter_info3.CursorLocation = 3
filter_info3.LockType = 3
filter_info3.Open()
filter_info3_numRows = 0

set filter_info4 = Server.CreateObject("ADODB.Recordset")
filter_info4.ActiveConnection = Connect
filter_info4.Source = "SELECT * FROM q_info4 order by info4"
filter_info4.CursorType = 0
filter_info4.CursorLocation = 3
filter_info4.LockType = 3
filter_info4.Open()
filter_info4_numRows = 0

set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT ID_subject, subject_name FROM subjects where subject_active_q <> 0"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0

set user_details = Server.CreateObject("ADODB.Recordset")
user_details.ActiveConnection = Connect
user_details.CursorType = 0
user_details.CursorLocation = 3
user_details.LockType = 3
user_details_numRows = 0

%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>Quiz users. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">

</HEAD>

<BODY BGCOLOR=#FFCC00 TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="0" topmargin="0" onload="check();">
<%
'Configure how you want to save the file
Response.Clear()
Response.AddHeader "Content-Disposition","inline; filename=USERS" & day(now()) & "_" & month(now()) & "_" & year(now()) & ".csv"
Response.ContentType = "application/vnd.ms-excel"
%>
"USERS"<%=vbcrlf%>
"Lastname","Firstname","Business","Active","Sessions","Avg Rate"<%=vbcrlf%>"---------------------------------------------------------------------------------------------------------------------"<%=vbcrlf%>
<%
           If Not users.EOF Or Not users.BOF Then %>
          <%
While (NOT users.EOF)
%>
          <%
if subject_prm <> 0 then
	subj_prm ="and (q_session.Session_subject ="&subject_prm&")"
else
	subj_prm =""
end if

if cstr(fromdate)="" and cstr(todate) <> "" then
	user_details.Source = "SELECT q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.Session_done, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and (session_finish <= '"&todate&"') "&subj_prm&" order by session_subject,session_date desc"
else if cstr(todate)="" and cstr(fromdate) <> "" then
	user_details.Source = "SELECT q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.Session_done, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and (session_finish >= '"&fromdate&"') "&subj_prm&" order by session_subject,session_date desc"
else if (cstr(todate)="" and cstr(fromdate)="") then
	user_details.Source = "SELECT q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.Session_done,q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) "&subj_prm&" order by session_subject,session_date desc"
else
	user_details.Source = "SELECT q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.Session_done,q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and ((session_finish >= '"&fromdate&"') and (session_finish <= '"&todate&"')) "&subj_prm&" order by session_subject,session_date desc"
end if
end if
end if
user_details.Open()
user_details_numRows = 0
user_session_rate = 0
user_session_count = 0
user_total_rate = 0

subid =0

While (NOT user_details.EOF)
if session("mths")="" then
	user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
	user_session_total = cInt(user_details.Fields.Item("session_total").Value)
	user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
	user_session_count = user_session_count + 1
else if cint(subid) <> cInt(user_details.Fields.Item("session_subject").Value) then
	subid = cInt(user_details.Fields.Item("session_subject").Value)
	user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
	user_session_total = cInt(user_details.Fields.Item("session_total").Value)
	user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
	user_session_count = user_session_count + 1
end if
end if
user_details.MoveNext()
Wend

user_details.Close()

if user_session_count > 0 then
	user_total_rate = (user_session_rate/user_session_count)
end if


if (cstr(noquiz)="1") then
	if cInt(users.Fields.Item("session_count").Value) = 0 then
%>
	 <%=(users.Fields.Item("user_lastname").Value)%>,<%=(users.Fields.Item("user_firstname").Value)%>,<%=(users.Fields.Item("info1").Value)%>,<%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "YES" else response.write "NO"%>,<%=(users.Fields.Item("session_count").Value)%>,<%="N/A"%>,
<%
count = count +1
end if
else


if cstr(results)="" or cstr(results)="2" then
%>
			<%
			if user_session_count = 0 then
				tmp =  "N/A"
			elseif (user_total_rate) >= passrate then
				tmp = FormatNumber(user_total_rate,2) & "%"
			else
				tmp = FormatNumber(user_total_rate,2) & "%"
			end if
			%>
			<%=(users.Fields.Item("user_lastname").Value)%>,<%=(users.Fields.Item("user_firstname").Value)%>,<%=(users.Fields.Item("info1").Value)%>,<%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "YES" else response.write "NO"%>,<%=(users.Fields.Item("session_count").Value)%>,<%=(tmp)%>,
<%
else if (cstr(results)="1") and (user_total_rate >= passrate) then
tmp= FormatNumber(user_total_rate,2) & "%"
count = count +1
%>
<%=(users.Fields.Item("user_lastname").Value)%>,<%=(users.Fields.Item("user_firstname").Value)%>,<%=(users.Fields.Item("info1").Value)%>,<%=(users.Fields.Item("session_count").Value)%>,<%=(tmp)%>,
<%
else if (cstr(results)="0") and (user_total_rate <= passrate) and (user_session_count<>0) then
tmp= FormatNumber(user_total_rate,2) & "%"
count = count +1
%>
<%=(users.Fields.Item("user_lastname").Value)%>,<%=(users.Fields.Item("user_firstname").Value)%>,<%=(users.Fields.Item("info1").Value)%>,<%=(users.Fields.Item("session_count").Value)%>,<%=(tmp)%>,
  <%
  end if
  end if
  end if
end if

  users.MoveNext()
  numbers=numbers+1
Wend
 users.Requery

%>
<%
'_______________________________________________________________________________

overall_session_rate = 0
overall_session_count = 0
overall_session_passed = 0
numbers=1
While (NOT users.EOF)

if subject_prm <> 0 then
	subj_prm ="and (q_session.Session_subject ="&subject_prm&")"
else
	subj_prm =""
end if

if cstr(fromdate)="" and cstr(todate) <> "" then
	user_details.Source = "SELECT q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and (session_finish <= '"&todate&"') "&subj_prm&" order by session_date desc,session_subject"
else if cstr(todate)="" and cstr(fromdate) <> "" then
	user_details.Source = "SELECT q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and (session_finish >= '"&fromdate&"') "&subj_prm&" order by session_date desc,session_subject"
else if (cstr(todate)="" and cstr(fromdate)="") then
	user_details.Source = "SELECT q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) "&subj_prm&" order by  session_date desc,session_subject"
else
	user_details.Source = "SELECT q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and ((session_finish >= '"&fromdate&"') and (session_finish <= '"&todate&"')) "&subj_prm&" order by session_date desc,session_subject"
end if
end if
end if
user_details.Open()
user_details_numRows = 0

user_session_rate = 0
user_session_count = 0
user_total_rate = 0
subid =0

While (NOT user_details.EOF)
if session("mths")="" then
	user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
	user_session_total = cInt(user_details.Fields.Item("session_total").Value)
	user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
	user_session_count = user_session_count + 1
else if cint(subid) <> cInt(user_details.Fields.Item("session_subject").Value) then
	subid = cInt(user_details.Fields.Item("session_subject").Value)
	user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
	user_session_total = cInt(user_details.Fields.Item("session_total").Value)
	user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
	user_session_count = user_session_count + 1
end if
end if
user_details.MoveNext()
Wend
user_details.close()
if user_session_count > 0 then
	user_total_rate = (user_session_rate/user_session_count)
	if cstr(results)="1" then
		if user_total_rate >= passrate then
			overall_session_rate = overall_session_rate + user_total_rate
			overall_session_count = overall_session_count + 1
			overall_session_passed = overall_session_passed + 1
		end if
	else if cstr(results)="0" then
		if user_total_rate <= passrate then
			overall_session_rate = overall_session_rate + user_total_rate
			overall_session_count = overall_session_count + 1
		end if
	else
		overall_session_rate = overall_session_rate + user_total_rate
		overall_session_count = overall_session_count + 1
		if user_total_rate >= passrate then overall_session_passed = overall_session_passed + 1
	end if
	end if
end if

  users.MoveNext()
  numbers=numbers+1
Wend

if overall_session_count > 0 then overll_pass_rate = overall_session_rate/overall_session_count
'________________________________________________________________________________
%>
Users in selection (filter) : <%
if (cstr(results)<> "2" and  cstr(results) <> "" ) or cstr(noquiz)="1" then
%>
<%Response.Write count-1%>
<%cnt = count -1%>
<%else%>
<%Response.Write numbers-1%>
<%cnt = numbers -1
end if
if cstr(noquiz)="1" then
overall_session_count=0
overall_session_passed =0
end if

%>
Users with at least 1 finished session:  <%=overall_session_count%>
<%
if overall_session_count = 0 then%>
Average score (%) of users with at least 1 finished session : <%response.write("Nothing to rate")
elseif overll_pass_rate >= passrate then%>
Average score (%) of users with at least 1 finished session : <%response.write(FormatNumber(overll_pass_rate,2))%>%
<%else %>
Average score (%) of users with at least 1 finished session : <%response.write(FormatNumber(overll_pass_rate,2))%>%
<%end if

if cint(cnt)= 0 then cnt =1
%>
Percentage of users who have completed a quiz : <%=FormatNumber((100 * (overall_session_count))/(cnt),2) %>%

<%=vbcrlf%>"---------------------------------------------------------------------------------------------------------------------"<%=vbcrlf%>"Generated on:","<%=Now()%>"<%=vbcrlf%><%=vbcrlf%>"Copyright <% Response.Write Year(now) %> (c) Law of the Jungle Pty Limited"

<% End If ' end Not users.EOF Or NOT users.BOF
call log_the_page ("Quiz List Users")
users.Close()
Set users = Nothing
filter_info1.Close()
filter_info3.Close()
filter_info4.Close()
subjects.Close()
%>
