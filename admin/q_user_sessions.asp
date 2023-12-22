<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
'show_lines = 50
'f cStr(Request.Querystring("show_lines")) <> "" then show_lines = cInt(Request.Querystring("show_lines"))
subject_prm = 0
If cInt(Request.Querystring("subject")) <> 0 then subject_prm = cInt(Request.Querystring("subject"))
filter_info1_prm = 0
If cInt(Request.Querystring("filter_info1")) <> 0 then filter_info1_prm = cInt(Request.Querystring("filter_info1"))
filter_info3_prm = 0
If cInt(Request.Querystring("filter_info3")) <> 0 then filter_info3_prm = cInt(Request.Querystring("filter_info3"))

Dim user__MMColParam
user__MMColParam = "1"
if (Request.QueryString("user") <> "") then user__MMColParam = Request.QueryString("user")

numbers=1

'PN050507 variables to store this users qinfo1 and qinfo3
current_users_qinfo1=0
current_users_qinfo3=0

' CHECK IF ONLINE OR OFFLINE
if request.querystring("status") = "" OR request.querystring("status") = "0" THEN
	status = 0
ELSEif request.querystring("status") = "1" THEN
	status = 1
ELSEif request.querystring("status") = "2" THEN
	status = 2
END IF
'END
' CHECK IF ONLINE OR OFFLINE
if request.querystring("status_q") = "" OR request.querystring("status_q") = "0" THEN
	status_q = 0
	status_sql = ""
ELSEif request.querystring("status_q") = "1" THEN
	status_q = 1
	status_sql = " and session_version = 0"
ELSEif request.querystring("status_q") = "2" THEN
	status_q = 2
	status_sql = " and session_version > 0"
END IF
'END

set user = Server.CreateObject("ADODB.Recordset")
user.ActiveConnection = Connect
user.Source = "SELECT * FROM q_user WHERE ID_user = " + Replace(user__MMColParam, "'", "''") + ""
user.CursorType = 0
user.CursorLocation = 3
user.LockType = 3
user.Open()
user_numRows = 0

'PN 050507 collect their division and activity
current_users_qinfo1=(user.Fields.Item("user_info1").Value)
current_users_qinfo3=(user.Fields.Item("user_info3").Value)


Dim sessions__MMColParam
sessions__MMColParam = "1"
if (Request.QueryString("user") <> "") then sessions__MMColParam = Request.QueryString("user")
if cint(request("subject")) <> 0 then
	t ="and (q_session.Session_subject ="&subject_prm&")"
	subject=cint(request("subject"))
else
	t = ""
	subject=0
end if
'Response.Write request("fromdate")
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
If session("mths") = "" or request("mths")= "" then
	sessions.Source = "SELECT q_session.session_version, q_session.ID_Session, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE Session_users = " + Replace(sessions__MMColParam, "'", "''") + " "&t&" "&t1&" "&status_sql&" ORDER BY subjects.id_subject, q_session.Session_date desc;"
else
	sessions.Source = "SELECT q_session.session_version, q_session.ID_Session, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE Session_users = " + Replace(sessions__MMColParam, "'", "''") + "AND (q_session.Session_done = 1) "&t&" "&t1&" "&status_sql&" ORDER BY subjects.id_subject, q_session.Session_date desc ;"
end if
sessions.CursorType = 0
sessions.CursorLocation = 3
sessions.LockType = 3
sessions.Open()
sessions_numRows = 0
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz user sessions. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
//-->
</script>
</HEAD>
<BODY>
<!--<%
	if Request.Cookies("passrate")<> "" then
		passrate= cint(Request.Cookies("passrate"))
	else
		passrate=50
	end if
%>-->
<table>
  <tr>
    <td class="heading"> Quiz user sessions</td>
  </tr>
  <tr>
    <td align="left" valign="bottom">


    <table>
        <tr>
          <td colspan="8" class="subheads"><%=(user.Fields.Item("user_firstname").Value)%>&nbsp;<%=(user.Fields.Item("user_lastname").Value)%>'s session(s):</td>
          <td class="subheads" align="right" valign="top"><a href="q_export_quiz_user_sessions.asp?user=<%=user__MMColParam%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=request("fromdate")%>&todate=<%=request("todate")%>&active=<%=request("active")%>&results=<%=request("results")%>&passrate=<%=passrate%>&mths=<%=request("mths")%>&noquiz=0"><img src="../admin/images/xls.gif" width="16" height="16" border="0"></a></td>
        </tr>
        <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
          <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
          <td colspan="8" class="text"><a href="../admin/q_list_of_users_results.asp?filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=request("fromdate")%>&todate=<%=request("todate")%>&active=<%=request("active")%>&results=<%=request("results")%>&passrate=<%=passrate%>&mths=<%=request("mths")%>&noquiz=0">...go
            up one level to list of users</a></td>
        </tr>
		
		<!-- OFFLINE/ONLINE FUNCTION-->
		<% IF pref_offline THEN %>
		<tr>
			<td >&nbsp;</td>
			<TD colspan="8" class="text"><br>
				<input type="radio" name="status" value="0" <% if clng(status_q)=0 THEN response.write "CHECKED"%> onClick="location.href='q_user_sessions.asp?user=<%=user__MMColParam%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=request("fromdate")%>&todate=<%=request("todate")%>&active=<%=request("active")%>&results=<%=request("results")%>&passrate=<%=passrate%>&mths=<%=request("mths")%>&noquiz=0&status=<% =status%>&status_q=0';"> All results
				&nbsp;&nbsp;
				<input type="radio" name="status" value="1" <% if clng(status_q)=1 THEN response.write "CHECKED"%> onClick="location.href='q_user_sessions.asp?user=<%=user__MMColParam%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=request("fromdate")%>&todate=<%=request("todate")%>&active=<%=request("active")%>&results=<%=request("results")%>&passrate=<%=passrate%>&mths=<%=request("mths")%>&noquiz=0&status=<% =status%>&status_q=1';"> Online
				&nbsp;&nbsp;
				<input type="radio" name="status" value="2" <% if clng(status_q)=2 THEN response.write "CHECKED"%> onClick="location.href='q_user_sessions.asp?user=<%=user__MMColParam%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=request("fromdate")%>&todate=<%=request("todate")%>&active=<%=request("active")%>&results=<%=request("results")%>&passrate=<%=passrate%>&mths=<%=request("mths")%>&noquiz=0&status=<% =status%>&status_q=2';"> Offline
			<br>
			</TD>
		</tr>
		<% END IF %>
		<!--END-->
        
		<% If Not sessions.EOF Or Not sessions.BOF Then %>
        <tr>
          <td class="text">&nbsp;</td>
          <td class="text">Subject</td>
          <td class="text">Date</td>
          <td class="text">Correct</td>
          <td class="text">Total Quest</td>
          <td class="text">Up to Page</td>
          <td class="text">Finished</td>
          <td class="text">Rate</td>
		  <td class="text">Pass Rate</td>
          <td class="text">Passed</td>

        </tr>
        <%
overall_rate = 0
sum_tests = 0
subid = 0

'PN 050507 variables to store user passes and fails
user_passes=0
user_fails=0
'PN050506 set up the variable that stores the passrate for the session
subject_pass_rate_percentage=default_passrate
While (NOT sessions.EOF)
if session("mths") <> "" then
	if cint(subid) <> (sessions.Fields.Item("id_subject").Value)then
	subid = (sessions.Fields.Item("id_subject").Value)


	'PN 050506 work out what the passrate will be for this users session
	subject_pass_rate_percentage = default_passrate
	if passrate_type = 1 then
		set pass_rate_query = Server.CreateObject("ADODB.Recordset")
		pass_rate_query.ActiveConnection = Connect
		pass_rate_query.Source = "SELECT * FROM q_certification where q_certification.q_session = "&sessions.Fields.Item("ID_Session").Value&""
		pass_rate_query.CursorType = 0
		pass_rate_query.CursorLocation = 3
		pass_rate_query.LockType = 3
		pass_rate_query.Open()
		if (not pass_rate_query.eof) then
			subject_pass_rate_percentage = pass_rate_query.Fields.Item("percentage_required").Value
		else
			set subject = Server.CreateObject("ADODB.Recordset")
			subject.ActiveConnection = Connect
			subject.Source = "SELECT * FROM subjects where subjects.id_subject = "&sessions.Fields.Item("id_subject").Value&""
			subject.CursorType = 0
			subject.CursorLocation = 3
			subject.LockType = 3
			subject.Open()
			subject_pass_rate_percentage = subject.fields.item("subject_passmark").value
			subject.Close()
		end if
		pass_rate_query.Close()
	elseif passrate_type=2 then
		set pass_rate_query = Server.CreateObject("ADODB.Recordset")
		pass_rate_query.ActiveConnection = Connect
		pass_rate_query.Source = "SELECT * FROM q_certification where q_certification.q_session = "&sessions.Fields.Item("ID_Session").Value&""
		pass_rate_query.CursorType = 0
		pass_rate_query.CursorLocation = 3
		pass_rate_query.LockType = 3
		pass_rate_query.Open()
		if (not pass_rate_query.eof) then
			subject_pass_rate_percentage = pass_rate_query.Fields.Item("percentage_required").Value
		else
			set pass_rate_query2 = Server.CreateObject("ADODB.Recordset")
			pass_rate_query2.ActiveConnection = Connect
			pass_rate_query2.Source = "SELECT pass_rate FROM pass_rates where subject="&sessions.Fields.Item("id_subject").Value&" and q_info1="&users.Fields.Item("user_info1").Value&" and q_info3="&users.Fields.Item("user_info3").Value&";"
			pass_rate_query2.CursorType = 0
			pass_rate_query2.CursorLocation = 3
			pass_rate_query2.LockType = 3
			pass_rate_query2.Open()
			if (not pass_rate_query2.eof) then
				subject_pass_rate_percentage = pass_rate_query2.Fields.Item("pass_rate").Value
			end if
			pass_rate_query2.Close()
		end if
		pass_rate_query.Close()
	elseif passrate_type=3 then
		set pass_rate_query = Server.CreateObject("ADODB.Recordset")
		pass_rate_query.ActiveConnection = Connect
		pass_rate_query.Source = "SELECT pass_rate FROM pass_rates where subject="&sessions.Fields.Item("id_subject").Value&" and q_info1="&users.Fields.Item("user_info1").Value&" and q_info3="&users.Fields.Item("user_info3").Value&";"
		pass_rate_query.CursorType = 0
		pass_rate_query.CursorLocation = 3
		pass_rate_query.LockType = 3
		pass_rate_query.Open()
		if (not pass_rate_query.eof) then
			subject_pass_rate_percentage = pass_rate_query.Fields.Item("pass_rate").Value
		end if
		pass_rate_query.Close()
	else
		set subject = Server.CreateObject("ADODB.Recordset")
		subject.ActiveConnection = Connect
		subject.Source = "SELECT * FROM subjects where subjects.id_subject = "&sessions.Fields.Item("id_subject").Value&""
		subject.CursorType = 0
		subject.CursorLocation = 3
		subject.LockType = 3
		subject.Open()
		subject_pass_rate_percentage = subject.fields.item("subject_passmark").value
		subject.Close()
	end if

	'________________________________________________________________________



	user_rate = FormatNumber((sessions.Fields.Item("Session_correct").Value)/(sessions.Fields.Item("Session_total").Value)*100,2)
	'PN050506 calulate if this session is a pass or a fail
	if abs(sessions.Fields.Item("Session_done").Value) = 1 then
		if(cInt(user_rate) >= cInt(subject_pass_rate_percentage)) then
			'it was a pass
			user_passes = user_passes + 1
		else
			'it was a fail
			user_fails = user_fails + 1
		end if
	end if
	if cInt(user_rate) >= cInt(subject_pass_rate_percentage) then user_pass = 1 else user_pass = 0
	overall_rate = overall_rate + user_rate
	%>
	        <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
	          <td class="text" width="10"><%=numbers%></td>
	          <td width="400" class="text"><a href="../admin/q_session_details.asp?user_session=<%=(sessions.Fields.Item("ID_Session").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&user=<%=(user.Fields.Item("ID_user").Value)%>&subject=<%=(sessions.Fields.Item("id_subject").Value)%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&todate=<%=request("todate")%>&fromdate=<%=request("fromdate")%>&active=<%=request("active")%>&results=<%=request("results")%>&passrate=<%=passrate%>&mths=<%=request("mths")%>&noquiz=0"><%=(sessions.Fields.Item("subject_name").Value)%></a></td>
	          <td class="text"><%=(sessions.Fields.Item("Session_date").Value)%></td>
	          <td width="20" class="text"><%=(sessions.Fields.Item("Session_correct").Value)%></td>
	          <td width="20" class="text"><%=(sessions.Fields.Item("Session_total").Value)%></td>
	          <td width="20" class="text"><%=(sessions.Fields.Item("Session_stop").Value)%></td>
	          <td class="text" width="20" align="center">

	            <%
				if abs(sessions.Fields.Item("Session_done").Value) = 1 then
					%><img src='images/1.gif'>
					</td>
					<td class="text" width="20" align="right">
						<%if user_pass = 1 then response.write ("<font color=green>" & user_rate & "%") else response.write ("<font color=red>" & user_rate & "%")%>
					</td>
					<td width="20" class="text"><%=subject_pass_rate_percentage%></td>
					<td class="text" width="20" align="center">
					    <%if user_pass = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
	          		</td> <%
					sum_tests = sum_tests+1
				else
					%><img src='images/0.gif'>
					</td>
					<td class="text" width="20" align="right">&nbsp;</td>
					<td width="20" class="text">&nbsp;</td>
					<td class="text" width="20" align="center">&nbsp;</td>
					<%
				end if
				%>
	        </tr>
	        <%
	  sessions.MoveNext()
	  numbers=numbers+1
else
	sessions.MoveNext()
end if
else
	'PN 050506 work out what the passrate will be for this users session
	subject_pass_rate_percentage = default_passrate
	if passrate_type = 1 then
		set pass_rate_query = Server.CreateObject("ADODB.Recordset")
		pass_rate_query.ActiveConnection = Connect
		pass_rate_query.Source = "SELECT * FROM q_certification where q_certification.q_session = "&sessions.Fields.Item("ID_Session").Value&""
		pass_rate_query.CursorType = 0
		pass_rate_query.CursorLocation = 3
		pass_rate_query.LockType = 3
		pass_rate_query.Open()
		if (not pass_rate_query.eof) then
			subject_pass_rate_percentage = pass_rate_query.Fields.Item("percentage_required").Value
		else
			set subject = Server.CreateObject("ADODB.Recordset")
			subject.ActiveConnection = Connect
			subject.Source = "SELECT * FROM subjects where subjects.id_subject = "&sessions.Fields.Item("id_subject").Value&""
			subject.CursorType = 0
			subject.CursorLocation = 3
			subject.LockType = 3
			subject.Open()
			subject_pass_rate_percentage = subject.fields.item("subject_passmark").value
			subject.Close()
		end if
		pass_rate_query.Close()
	elseif passrate_type=2 then
		set pass_rate_query = Server.CreateObject("ADODB.Recordset")
		pass_rate_query.ActiveConnection = Connect
		pass_rate_query.Source = "SELECT * FROM q_certification where q_certification.q_session = "&sessions.Fields.Item("ID_Session").Value&""
		pass_rate_query.CursorType = 0
		pass_rate_query.CursorLocation = 3
		pass_rate_query.LockType = 3
		pass_rate_query.Open()
		if (not pass_rate_query.eof) then
			subject_pass_rate_percentage = pass_rate_query.Fields.Item("percentage_required").Value
		else
			set pass_rate_query2 = Server.CreateObject("ADODB.Recordset")
			pass_rate_query2.ActiveConnection = Connect
			pass_rate_query2.Source = "SELECT pass_rate FROM pass_rates where subject="&sessions.Fields.Item("id_subject").Value&" and q_info1="&users.Fields.Item("user_info1").Value&" and q_info3="&users.Fields.Item("user_info3").Value&";"
			pass_rate_query2.CursorType = 0
			pass_rate_query2.CursorLocation = 3
			pass_rate_query2.LockType = 3
			pass_rate_query2.Open()
			if (not pass_rate_query2.eof) then
				subject_pass_rate_percentage = pass_rate_query2.Fields.Item("pass_rate").Value
			end if
			pass_rate_query2.Close()
		end if
		pass_rate_query.Close()
	elseif passrate_type=3 then
		set pass_rate_query = Server.CreateObject("ADODB.Recordset")
		pass_rate_query.ActiveConnection = Connect
		pass_rate_query.Source = "SELECT pass_rate FROM pass_rates where subject="&sessions.Fields.Item("id_subject").Value&" and q_info1="&users.Fields.Item("user_info1").Value&" and q_info3="&users.Fields.Item("user_info3").Value&";"
		pass_rate_query.CursorType = 0
		pass_rate_query.CursorLocation = 3
		pass_rate_query.LockType = 3
		pass_rate_query.Open()
		if (not pass_rate_query.eof) then
			subject_pass_rate_percentage = pass_rate_query.Fields.Item("pass_rate").Value
		end if
		pass_rate_query.Close()
	else
		set subject = Server.CreateObject("ADODB.Recordset")
		subject.ActiveConnection = Connect
		subject.Source = "SELECT * FROM subjects where subjects.id_subject = "&sessions.Fields.Item("id_subject").Value&""
		subject.CursorType = 0
		subject.CursorLocation = 3
		subject.LockType = 3
		subject.Open()
		subject_pass_rate_percentage = subject.fields.item("subject_passmark").value
		subject.Close()
	end if

	'________________________________________________________________________
	user_rate = ((sessions.Fields.Item("Session_correct").Value)/(sessions.Fields.Item("Session_total").Value)*100)
	'PN050506 calulate if this session is a pass or a fail
	if abs(sessions.Fields.Item("Session_done").Value) = 1 then
		if(cInt(user_rate) >= cInt(subject_pass_rate_percentage)) then
			'it was a pass
			user_passes = user_passes + 1
		else
			'it was a fail
			user_fails = user_fails + 1
		end if
	end if
	if cInt(user_rate) >= cInt(subject_pass_rate_percentage) then user_pass = 1 else user_pass = 0
	overall_rate = overall_rate + user_rate
	%>
	        <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
	          <td class="text" width="10"><%=numbers%></td>
	          <td width="400" class="text"><a href="../admin/q_session_details.asp?user_session=<%=(sessions.Fields.Item("ID_Session").Value)%>&user=<%=(user.Fields.Item("ID_user").Value)%>&filter_username=<%=request("filter_username")%>&subject=<%=(sessions.Fields.Item("id_subject").Value)%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=request("fromdate")%>&todate=<%=request("todate")%>&active=<%=request("active")%>&results=<%=request("results")%>&passrate=<%=passrate%>&mths=<%=request("mths")%>&noquiz=0""><%=(sessions.Fields.Item("subject_name").Value)%></a></td>
	          <td class="text"><%=(sessions.Fields.Item("Session_date").Value)%></td>
	          <td width="20" class="text"><%=(sessions.Fields.Item("Session_correct").Value)%></td>
	          <td width="20" class="text"><%=(sessions.Fields.Item("Session_total").Value)%></td>
	          <td width="20" class="text"><%=(sessions.Fields.Item("Session_stop").Value)%></td>
	          <td class="text" width="20" align="center">
	            <%
				if abs(sessions.Fields.Item("Session_done").Value) = 1 then
					%><img src='images/1.gif'>
					</td>
					<td class="text" width="20" align="right">
						<%if user_pass = 1 then response.write ("<font color=green>" & FormatNumber(user_rate,2) & "%") else response.write ("<font color=red>" & FormatNumber(user_rate,2) & "%")%>
					</td>
					<td width="20" class="text"><%=subject_pass_rate_percentage%></td>
					<td class="text" width="20" align="center">
					    <%if user_pass = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
	          		</td> <%
					sum_tests = sum_tests+1
				else
					%><img src='images/0.gif'>
					</td>
					<td class="text" width="20" align="right">&nbsp;</td>
					<td width="20" class="text">&nbsp;</td>
					<td class="text" width="20" align="center">&nbsp;</td>
					<%
				end if
				%>

	        </tr>
	        <%
	  sessions.MoveNext()
	  numbers=numbers+1
end if
Wend
if sum_tests = 0 then overall_pass = 0 else overall_pass = overall_rate/sum_tests
%>
        <tr>
          <td colspan="11">
            <hr>
          </td>
        </tr>
        <tr class="table_normal">
          <td class="text" width="10">&nbsp;</td>
          <td width="400" class="text">Completed Sessions</td>
          <td class="text">&nbsp;</td>
          <td colspan="3" class="text" align="left">Average %</td>
		  <td  class="text" align="left">Passes</td>
		  <td  class="text" align="left">Fails</td>
          <td class="text" colspan="3" align="left"></td>
        </tr>
        <tr class="table_normal">
          <td class="text" width="10">&nbsp;</td>
          <td width="400" class="text"><%=sum_tests%></td>
          <td class="text">&nbsp;</td>
          <td colspan="3" class="text" align="left">
            <%=FormatNumber(overall_pass,2)%>%
          </td>
		  <td  class="text" align="left"><font color=green><%=user_passes%></font></td>
		  <td  class="text" align="left"><font color=red><%=user_fails%></font></td>

          <td class="text" colspan="3" align="left">
            <!--<%
			if sum_tests = 0 then
				response.write("<font color=blue>N/A</font>")
			elseif (overall_pass) >= passrate then
				response.write("<font color=green>PASSED</font>")
			else
				response.write("<font color=red>FAILED</font>")
			end if
			%>-->
          </td>
        </tr>
        <tr>
          <td  colspan="9"></td>
        </tr>
        <% End If %>
		<% If sessions.EOF And sessions.BOF Then %>
        <tr>
          <td >&nbsp;</td>
          <td colspan="8" >Sorry,
            there are no user sessions in the quiz currently.</td>
        </tr>
		<% End If %>
      </table>
      <p>&nbsp;</p>
      </td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("Quiz User Sessions: " & user__MMColParam)
user.Close()
sessions.Close()
%>
