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

set users = Server.CreateObject("ADODB.Recordset")
users.ActiveConnection = Connect
users.Source = "SELECT q_user.ID_user, q_user.user_lastname, q_user.user_firstname, q_info1.info1, q_info2.info2, q_info3.info3, q_info4.info4, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, q_user.user_info4, COUNT(q_session.ID_session) AS session_count FROM (q_info4 RIGHT JOIN (q_info3 RIGHT JOIN (q_info2 RIGHT JOIN (q_info1 RIGHT JOIN q_user ON q_info1.ID_info1 = q_user.user_info1) ON q_info2.ID_info2 = q_user.user_info2) ON q_info3.ID_info3 = q_user.user_info3) ON q_info4.ID_info4 = q_user.user_info4) LEFT JOIN q_session ON q_user.ID_user = q_session.Session_users " + SQL_where + " GROUP BY q_user.user_lastname, q_user.user_firstname, q_user.ID_user, q_info1.info1, q_info2.info2, q_info3.info3, q_info4.info4, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3 , q_user.user_info4 " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"
Response.Write users.Source
users.CursorType = 0
users.CursorLocation = 3
users.LockType = 3
users.Open()
users_numRows = 0

set filter_info1 = Server.CreateObject("ADODB.Recordset")
filter_info1.ActiveConnection = Connect
filter_info1.Source = "SELECT * FROM q_info1 where id_info1 = '" & filter_info1_prm & "' order by info1"
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
subjects.Source = "SELECT ID_subject, subject_name FROM subjects where subject_active_q <> 0 order by id_subject"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0

set sessions = Server.CreateObject("ADODB.Recordset")
sessions.ActiveConnection = Connect
sessions.CursorType = 0
sessions.CursorLocation = 3
sessions.LockType = 3

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
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz users. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">

</HEAD>

<BODY BGCOLOR=#FFCC00 TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="0" topmargin="0" onload="check();">
<%
Response.Clear()
Response.AddHeader "Content-Disposition","attachment; filename=USERS_Summary" & day(now()) & "_" & month(now()) & "_" & year(now()) & ".csv"
Response.ContentType="application/vnd.ms-excel"
count = subjects.RecordCount
i = 1
redim arr1(count)
%>
"USERS - Summary"<%=vbcrlf%>
"LAST NAME","FIRST NAME",<%while not subjects.eof
	arr1(i) = subjects.Fields.Item("id_subject").Value
	Response.write ucase(subjects.Fields.Item("subject_name").Value)
	Response.write ","
	i = i + 1
	subjects.movenext
wend
subjects.Close
%>
"-----------------------------------------------------------"<%=vbcrlf%>
<%
If Not users.EOF Or Not users.BOF Then
While (NOT users.EOF)

If subject_prm = 0 then
	sessions.Source = "SELECT q_session.ID_Session, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE Session_users = "&users.Fields.Item("id_user").Value&" AND (q_session.Session_done = 1) "&t&" "&t1&" ORDER BY subjects.id_subject, q_session.Session_date desc;"
else
	sessions.Source = "SELECT q_session.ID_Session, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE Session_users = "&users.Fields.Item("id_user").Value& " AND (q_session.Session_done = 1) and q_session.Session_subject="&subject_prm&" "&t&" "&t1&" ORDER BY subjects.id_subject, q_session.Session_date DESC ;"
end if

sessions.Open()
i = 1
while not sessions.EOF

		if cint(userid) <> (users.Fields.Item("id_user").Value)then
		Response.write vbcrlf
		userid = (users.Fields.Item("id_user").Value)
		subid=0

			if cint(subid) <> (sessions.Fields.Item("id_subject").Value)then
			subid = (sessions.Fields.Item("id_subject").Value)
				response.write users.Fields.Item("user_lastname").Value
				Response.write ","
				response.write users.Fields.Item("user_firstname").Value
				Response.write ","
				while arr1(i) <> sessions.Fields.Item("id_subject").Value
					Response.write ","
					i = i + 1
				wend
				if sessions.Fields.Item("id_subject").Value = arr1(i) then
					user_rate = FormatNumber((sessions.Fields.Item("Session_correct").Value)/(sessions.Fields.Item("Session_total").Value)*100,2)
					if cInt(user_rate) >= cInt(passrate) then
						Response.Write "PASS"
					else
						Response.Write "FAIL"
					end if
					Response.write ","
					i = i +1
				end if
			end if
		else
			if cint(subid) <> (sessions.Fields.Item("id_subject").Value)then
			subid = (sessions.Fields.Item("id_subject").Value)
				while arr1(i) <> sessions.Fields.Item("id_subject").Value
					Response.write ","
					i = i + 1
				wend
				if sessions.Fields.Item("id_subject").Value = arr1(i) then
					user_rate = FormatNumber((sessions.Fields.Item("Session_correct").Value)/(sessions.Fields.Item("Session_total").Value)*100,2)
					if cInt(user_rate) >= cInt(passrate) then
						Response.Write "PASS"
					else
						Response.Write "FAIL"
					end if
					Response.write ","
					i = i +1
				end if
			end if
		end if

sessions.MoveNext()
wend

sessions.Close()

users.MoveNext()
Wend
users.Requery
end if

call log_the_page ("Quiz List Users")
users.Close()
Set users = Nothing
filter_info1.Close()
filter_info3.Close()
filter_info4.Close()
%>
